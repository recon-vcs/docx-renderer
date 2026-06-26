import { BreakType, DomType, OpenXmlElement, WmlBreak, WmlHyperlink } from '@docx/ooxml/wordprocessingml/document/model/dom';
import type { WmlParagraph } from '@docx/ooxml/wordprocessingml/document/model/paragraph';
import { parseParagraphProperty } from '@docx/ooxml/wordprocessingml/document/model/paragraph';
import type { WmlRun } from '@docx/ooxml/wordprocessingml/document/model/run';
import { parseBookmarkStart, parseBookmarkEnd } from '@docx/ooxml/wordprocessingml/document/model/bookmarks';
import { WmlCommentRangeStart, WmlCommentRangeEnd } from '@docx/ooxml/wordprocessingml/parts/comments/elements';
import xml from '@docx/xml/parsing/xml-parser';
import { xmlUtil, values } from '@docx/xml/parsing/parse-utils';
import { parseDefaultProperties } from './properties-parser';
import { parseSdt } from './content-parser';
import type { ParagraphParserContext } from './parse-context';

export function parseInserted(
	node: Element,
	ctx: ParagraphParserContext,
): OpenXmlElement {
	const wmlInserted: OpenXmlElement = {
		type: DomType.Inserted,
		children: [],
	};
	xmlUtil.foreach(node, (child) => {
		switch (child.localName) {
			case "r":
				wmlInserted.children.push(ctx.parseRun(child));
				break;
			default:
				if (ctx.options.debug) {
					console.warn(`DOCX:%c Unknown Inserted：${child.localName}`, 'color:#f75607');
				}
		}
	});
	return wmlInserted;
}

// TODO
export function parseDeleted(
	node: Element,
	ctx: ParagraphParserContext,
): OpenXmlElement {
	const wmlDeleted: OpenXmlElement = {
		type: DomType.Deleted,
		children: [],
	};
	xmlUtil.foreach(node, (child) => {
		switch (child.localName) {
			case "r":
				wmlDeleted.children.push(ctx.parseRun(child));
				break;
			default:
				if (ctx.options.debug) {
					console.warn(`DOCX:%c Unknown Inserted：${child.localName}`, 'color:#f75607');
				}
		}
	});
	return wmlDeleted;
}

export function parseParagraph(
	node: Element,
	ctx: ParagraphParserContext,
): WmlParagraph {
	const wmlParagraph: WmlParagraph = {
		type: DomType.Paragraph,
		children: [],
		props: {},
		cssStyle: {},
	};

	xmlUtil.foreach(node, (child) => {
		switch (child.localName) {
			// property
			case "pPr":
				parseParagraphProperties(child, wmlParagraph, ctx);
				break;

			case "r":
				wmlParagraph.children.push(ctx.parseRun(child));
				break;

			case "hyperlink":
				wmlParagraph.children.push(parseHyperlink(child, ctx));
				break;

			case "bookmarkStart":
				wmlParagraph.children.push(parseBookmarkStart(child, xml));
				break;

			case "bookmarkEnd":
				wmlParagraph.children.push(parseBookmarkEnd(child, xml));
				break;

			case "commentRangeStart":
				wmlParagraph.children.push(new WmlCommentRangeStart(xml.attr(child, "id")));
				break;

			case "commentRangeEnd":
				wmlParagraph.children.push(new WmlCommentRangeEnd(xml.attr(child, "id")));
				break;

			case "oMath":
			case "oMathPara":
				wmlParagraph.children.push(ctx.parseMathElement(child));
				break;

			// TODO Structured Document Tag
			case "sdt":
				wmlParagraph.children.push(...parseSdt(child, ctx));
				break;

			// TODO Inserted Math Control Character, Inserted Table Row, Inserted Numbering Properties
			case "ins":
				wmlParagraph.children.push(parseInserted(child, ctx));
				break;

			case "del":
				wmlParagraph.children.push(parseDeleted(child, ctx));
				break;

			default:
				if (ctx.options.debug) {
					console.warn(`DOCX:%c Unknown Paragraph Element：${child.localName}`, 'color:#f75607');
				}
		}
	});

	// when paragraph is empty, add a br tag to work with rich text editor and generate line height
	if (wmlParagraph.children.length === 0) {
		const wmlBreak: WmlBreak = { type: DomType.Break, "break": BreakType.TextWrapping };
		const wmlRun = { type: DomType.Run, children: [wmlBreak as OpenXmlElement] } as WmlRun;
		wmlParagraph.children = [wmlRun];
	}

	return wmlParagraph;
}

export function parseParagraphProperties(
	elem: Element,
	paragraph: WmlParagraph,
	ctx: ParagraphParserContext,
): void {
	parseDefaultProperties(elem, ctx.options, paragraph.cssStyle = {}, null, c => {
		if (parseParagraphProperty(c, paragraph.props, xml)) {
			return true;
		}

		switch (c.localName) {
			// Paragraph Conditional Formatting
			case "cnfStyle":
				paragraph.className = values.classNameOfCnfStyle(c);
				break;

			// Text Frame Properties
			case "framePr":
				parseFrame(c, paragraph);
				break;

			// Referenced Paragraph Style
			case "pStyle":
				paragraph.styleName = xml.attr(c, "val");
				break;

			default:
				return false;
		}

		return true;
	});
}

export function parseFrame(node: Element, paragraph: WmlParagraph): void {
	const dropCap = xml.attr(node, "dropCap");
	if (dropCap == "drop")
		paragraph.cssStyle["float"] = "left";
}

export function parseHyperlink(
	node: Element,
	ctx: ParagraphParserContext,
): WmlHyperlink {
	const wmlHyperlink: WmlHyperlink = <WmlHyperlink>{
		type: DomType.Hyperlink,
		children: [],
	};
	const anchor = xml.attr(node, "anchor");
	const relId = xml.attr(node, "id");

	if (anchor) {
		wmlHyperlink.href = "#" + anchor;
	}
	if (relId) {
		wmlHyperlink.id = relId;
	}

	xmlUtil.foreach(node, (child) => {
		switch (child.localName) {
			case "r":
				wmlHyperlink.children.push(ctx.parseRun(child));
				break;
			default:
				if (ctx.options.debug) {
					console.warn(`DOCX:%c Unknown Hyperlink Element：${child.localName}`, 'color:#f75607');
				}
		}
	});

	return wmlHyperlink;
}
