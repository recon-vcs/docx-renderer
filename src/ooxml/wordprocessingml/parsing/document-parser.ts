import { DomType, IDomNumbering, OpenXmlElement } from '@docx/ooxml/wordprocessingml/document/model/dom';
import { DocumentElement } from '@docx/ooxml/wordprocessingml/document/model/document';
import { parseSectionProperties, SectionProperties } from '@docx/ooxml/wordprocessingml/document/model/section';
import xml from '@docx/xml/parsing/xml-parser';
import { IDomStyle, Ruleset } from '@docx/ooxml/wordprocessingml/document/model/style';
import { uuid } from '@docx/shared/utils';
import { xmlUtil } from '@docx/xml/parsing/parse-utils';
import {
	parseTable as parseTableFn,
} from './table-parser';
import {
	parseStylesFile as parseStylesFileFn,
	parseDefaultStyles as parseDefaultStylesFn,
	parseStyle as parseStyleFn,
	parseTableStyle as parseTableStyleFn,
} from './style-parser';
import {
	parseNumberingFile as parseNumberingFileFn,
} from './numbering-parser';
import {
	parseVmlPicture as parseVmlPictureFn,
	checkAlternateContent as checkAlternateContentFn,
	parseDrawing as parseDrawingFn,
} from '@docx/ooxml/drawingml/parsing/drawing-parser';
import {
	parseMathElement as parseMathElementFn,
} from '@docx/ooxml/omml/parsing/math-parser';
import {
	parseComments as parseCommentsFn,
	parseNotes as parseNotesFn,
	parseSdt as parseSdtFn,
} from './content-parser';
import {
	parseDefaultProperties as parseDefaultPropertiesFn,
} from './properties-parser';
import { parseParagraph as parseParagraphFn } from './paragraph-parser';
import { parseRun as parseRunFn } from './run-parser';
import type { ParseContext } from './parse-context';

export interface DocumentParserOptions {
	ignoreWidth: boolean;
	debug: boolean;
	ignoreTableWrap: boolean;
	ignoreImageWrap: boolean;
}

// Default parse options
export const defaultDocumentParserOptions: DocumentParserOptions = {
	ignoreWidth: false,
	debug: false,
	ignoreTableWrap: true,
	ignoreImageWrap: true,
}

// Module-level body element parser — needs ctx for recursive calls
function parseBodyElementsFn(element: Element, ctx: ParseContext): OpenXmlElement[] {
	const children: OpenXmlElement[] = [];

	xmlUtil.foreach(element, (child) => {
		switch (child.localName) {
			case "p":
				children.push(parseParagraphFn(child, ctx));
				break;

			case "tbl":
				children.push(parseTableFn(child, ctx));
				break;

			// TODO Structured Document
			case "sdt":
				children.push(...parseSdtFn(child, ctx));
				break;

			case "sectPr":
				// ignore — section property is parsed in parseDocumentFile
				break;

			default:
				if (ctx.options.debug) {
					console.warn(`DOCX:%c Unknown Body Element：${child.localName}`, 'color:red');
				}
		}
	});

	return children;
}

export class DocumentParser {
	options: DocumentParserOptions;
	private ctx: ParseContext;

	constructor(options?: Partial<DocumentParserOptions>) {
		this.options = {
			...defaultDocumentParserOptions,
			...options
		};

		// Build ParseContext using a lazy-bound object so all closures share
		// the same ctx reference by the time any method is invoked.
		const ctx = {} as ParseContext;
		const opts = this.options;
		Object.assign(ctx, {
			options: opts,
			parseBodyElements: (e: Element) => parseBodyElementsFn(e, ctx),
			parseParagraph: (e: Element) => parseParagraphFn(e, ctx),
			parseTable: (e: Element) => parseTableFn(e, ctx),
			parseRun: (e: Element) => parseRunFn(e, ctx),
			parseMathElement: (e: Element) => parseMathElementFn(e, ctx),
			parseDefaultProperties: (
				e: Element,
				s?: Record<string, string>,
				c?: Record<string, string>,
				h?: (p: Element) => boolean,
			) => parseDefaultPropertiesFn(e, opts, s, c, h),
			parseDrawing: (e: Element) => parseDrawingFn(e, ctx),
			parseVmlPicture: (e: Element) => parseVmlPictureFn(e, ctx),
			checkAlternateContent: (e: Element) => checkAlternateContentFn(e),
		} satisfies ParseContext);
		this.ctx = ctx;
	}

	parseDocumentFile(xmlDoc: Element): DocumentElement {
			const documentElement: DocumentElement = {
				uuid: 'root',
				sectProps: {} as SectionProperties,
				type: DomType.Document,
			};
		// background color
		const background = xml.element(xmlDoc, "background");
		documentElement.cssStyle = background ? this.parseBackground(background) : {};
		// parse child elements
		const body = xml.element(xmlDoc, "body");
		documentElement.children = this.parseBodyElements(body);
		// section properties
		const sectionProperties = xml.element(body, "sectPr");
		if (sectionProperties) {
			documentElement.sectProps = parseSectionProperties(sectionProperties, xml);
		}
		documentElement.sectProps.sectionId = uuid();

		return documentElement;
	}

	parseBackground(elem: Element): Record<string, string> {
		const result: Record<string, string> = {};
		const color = xmlUtil.colorAttr(elem, "color");

		if (color) {
			result["background-color"] = color;
		}

		return result;
	}

	parseBodyElements(element: Element): OpenXmlElement[] {
		return parseBodyElementsFn(element, this.ctx);
	}

	parseStylesFile(xstyles: Element): IDomStyle[] {
		return parseStylesFileFn(xstyles, this.ctx);
	}

	parseDefaultStyles(node: Element): IDomStyle {
		return parseDefaultStylesFn(node, this.ctx);
	}

	parseStyle(node: Element): IDomStyle {
		return parseStyleFn(node, this.ctx);
	}

	// TODO: multi-level nested table style rules are not yet applied
	parseTableStyle(node: Element): Ruleset[] {
		return parseTableStyleFn(node, this.ctx);
	}

	parseNumberingFile(xnums: Element): IDomNumbering[] {
		return parseNumberingFileFn(xnums, this.ctx);
	}

	parseNotes<T extends { id?: string; noteType?: string; children?: OpenXmlElement[] }>(
		xmlDoc: Element,
		elemName: string,
		elemClass: new () => T,
	): T[] {
		return parseNotesFn(xmlDoc, elemName, elemClass, this.ctx);
	}

	parseComments(xmlDoc: Element): ReturnType<typeof parseCommentsFn> {
		return parseCommentsFn(xmlDoc, this.ctx);
	}
}
