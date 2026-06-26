import { BreakType, DomType, OpenXmlElement, WmlBreak, WmlHyperlink } from '../document/dom';
import type { WmlParagraph } from '../document/paragraph';
import { parseParagraphProperty } from '../document/paragraph';
import type { WmlRun } from '../document/run';
import { parseBookmarkStart, parseBookmarkEnd } from '../document/bookmarks';
import { WmlCommentRangeStart, WmlCommentRangeEnd } from '../comments/elements';
import type { DocumentParserOptions } from '../document-parser';
import xml from './xml-parser';
import { xmlUtil, values } from './parse-utils';
import { parseDefaultProperties } from './properties-parser';
import { parseSdt } from './content-parser';

export interface ParagraphParserCallbacks {
	parseRun(node: Element): WmlRun;
	parseMathElement(node: Element): OpenXmlElement;
	parseBodyElements(node: Element): OpenXmlElement[];
}

// TODO Inserted Math Control Character, Inserted Table Row, Inserted Numbering Properties
export function parseInserted(
	node: Element,
	options: DocumentParserOptions,
	callbacks: ParagraphParserCallbacks,
): OpenXmlElement {
	let wmlInserted: OpenXmlElement = {
		type: DomType.Inserted,
		children: [],
	};
	xmlUtil.foreach(node, (child) => {
		switch (child.localName) {
			case "r":
				wmlInserted.children.push(callbacks.parseRun(child));
				break;
			default:
				if (options.debug) {
					console.warn(`DOCX:%c Unknown Inserted：${child.localName}`, 'color:#f75607');
				}
		}
	});
	return wmlInserted;
}

// TODO
export function parseDeleted(
	node: Element,
	options: DocumentParserOptions,
	callbacks: ParagraphParserCallbacks,
): OpenXmlElement {
	let wmlDeleted: OpenXmlElement = {
		type: DomType.Deleted,
		children: [],
	};
	xmlUtil.foreach(node, (child) => {
		switch (child.localName) {
			case "r":
				wmlDeleted.children.push(callbacks.parseRun(child));
				break;
			default:
				if (options.debug) {
					console.warn(`DOCX:%c Unknown Inserted：${child.localName}`, 'color:#f75607');
				}
		}
	});
	return wmlDeleted;
}

export function parseParagraph(
	node: Element,
	options: DocumentParserOptions,
	callbacks: ParagraphParserCallbacks,
): WmlParagraph {
	let wmlParagraph: WmlParagraph = {
		type: DomType.Paragraph,
		children: [],
		props: {},
		cssStyle: {},
	};

	xmlUtil.foreach(node, (child) => {
		switch (child.localName) {
			// property
			case "pPr":
				parseParagraphProperties(child, wmlParagraph, options);
				break;

			case "r":
				wmlParagraph.children.push(callbacks.parseRun(child));
				break;

			case "hyperlink":
				wmlParagraph.children.push(parseHyperlink(child, options, callbacks));
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
				wmlParagraph.children.push(callbacks.parseMathElement(child));
				break;

			// TODO Structured Document Tag
			case "sdt":
				wmlParagraph.children.push(...parseSdt(child, callbacks));
				break;

			// TODO Inserted Math Control Character, Inserted Table Row, Inserted Numbering Properties
			case "ins":
				wmlParagraph.children.push(parseInserted(child, options, callbacks));
				break;

			case "del":
				wmlParagraph.children.push(parseDeleted(child, options, callbacks));
				break;

			default:
				if (options.debug) {
					console.warn(`DOCX:%c Unknown Paragraph Element：${child.localName}`, 'color:#f75607');
				}
		}
	});

	// when paragraph is empty, add a br tag to work with rich text editor and generate line height
	if (wmlParagraph.children.length === 0) {
		let wmlBreak: WmlBreak = { type: DomType.Break, "break": BreakType.TextWrapping };
		let wmlRun = { type: DomType.Run, children: [wmlBreak as OpenXmlElement] } as WmlRun;
		wmlParagraph.children = [wmlRun];
	}

	return wmlParagraph;
}

export function parseParagraphProperties(
	elem: Element,
	paragraph: WmlParagraph,
	options: DocumentParserOptions,
): void {
	parseDefaultProperties(elem, options, paragraph.cssStyle = {}, null, c => {
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
	let dropCap = xml.attr(node, "dropCap");
	if (dropCap == "drop")
		paragraph.cssStyle["float"] = "left";
}

export function parseHyperlink(
	node: Element,
	options: DocumentParserOptions,
	callbacks: ParagraphParserCallbacks,
): WmlHyperlink {
	let wmlHyperlink: WmlHyperlink = <WmlHyperlink>{
		type: DomType.Hyperlink,
		children: [],
	};
	let anchor = xml.attr(node, "anchor");
	let relId = xml.attr(node, "id");

	if (anchor) {
		wmlHyperlink.href = "#" + anchor;
	}
	if (relId) {
		wmlHyperlink.id = relId;
	}

	xmlUtil.foreach(node, (child) => {
		switch (child.localName) {
			case "r":
				wmlHyperlink.children.push(callbacks.parseRun(child));
				break;
			default:
				if (options.debug) {
					console.warn(`DOCX:%c Unknown Hyperlink Element：${child.localName}`, 'color:#f75607');
				}
		}
	});

	return wmlHyperlink;
}
