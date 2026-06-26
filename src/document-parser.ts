import { DomType, IDomNumbering, NumberingPicBullet, OpenXmlElement } from './document/dom';
import { DocumentElement } from './document/document';
import { WmlParagraph } from './document/paragraph';
import { parseSectionProperties, SectionProperties } from './document/section';
import xml from './parser/xml-parser';
import { WmlRun } from './document/run';
import { IDomStyle, Ruleset } from './document/style';
import { uuid } from "./utils";
import { xmlUtil } from './parser/parse-utils';
import {
	parseTable as parseTableFn,
	TableParserCallbacks,
} from './parser/table-parser';
import {
	parseStylesFile as parseStylesFileFn,
	parseDefaultStyles as parseDefaultStylesFn,
	parseStyle as parseStyleFn,
	parseTableStyle as parseTableStyleFn,
} from './parser/style-parser';
import {
	parseNumberingFile as parseNumberingFileFn,
	parseNumberingPicBullet as parseNumberingPicBulletFn,
	parseAbstractNumbering as parseAbstractNumberingFn,
	parseNumberingLevel as parseNumberingLevelFn,
} from './parser/numbering-parser';
import {
	parseVmlPicture as parseVmlPictureFn,
	checkAlternateContent as checkAlternateContentFn,
	parseDrawing as parseDrawingFn,
	DrawingParserCallbacks,
} from './parser/drawing-parser';
import {
	parseMathElement as parseMathElementFn,
	MathParserCallbacks,
} from './parser/math-parser';
import {
	parseComments as parseCommentsFn,
	parseNotes as parseNotesFn,
	parseSdt as parseSdtFn,
} from './parser/content-parser';
import { parseParagraph as parseParagraphFn } from './parser/paragraph-parser';
import { parseRun as parseRunFn } from './parser/run-parser';
import {
	parseDefaultProperties as parseDefaultPropertiesFn,
} from './parser/properties-parser';


export interface DocumentParserOptions {
	ignoreWidth: boolean;
	debug: boolean;
	ignoreTableWrap: boolean,
	ignoreImageWrap: boolean,
}

// 默认解析选项
export const defaultDocumentParserOptions: DocumentParserOptions = {
	ignoreWidth: false,
	debug: false,
	ignoreTableWrap: true,
	ignoreImageWrap: true,
}

export class DocumentParser {
	options: DocumentParserOptions;

	constructor(options?: Partial<DocumentParserOptions>) {
		this.options = {
			...defaultDocumentParserOptions,
			...options
		};
	}

	parseDocumentFile(xmlDoc: Element): DocumentElement {
		// document elements
		let documentElement: DocumentElement = {
			uuid: 'root',
			pages: [],
			sectProps: {} as SectionProperties,
			type: DomType.Document,
		};
		// 背景色
		let background = xml.element(xmlDoc, "background");
		documentElement.cssStyle = background ? this.parseBackground(background) : {};
		// 处理子元素
		let body = xml.element(xmlDoc, "body");
		documentElement.children = this.parseBodyElements(body);
		// 计算节属性
		let sectionProperties = xml.element(body, "sectPr");
		if (sectionProperties) {
			documentElement.sectProps = parseSectionProperties(sectionProperties, xml);
		}
		// 生成唯一uuid标识
		documentElement.sectProps.sectionId = uuid();

		return documentElement;
	}

	parseBackground(elem: Element): any {
		let result = {};
		let color = xmlUtil.colorAttr(elem, "color");

		if (color) {
			result["background-color"] = color;
		}

		return result;
	}

	parseBodyElements(element: Element): OpenXmlElement[] {
		let children = [];

		xmlUtil.foreach(element, (child) => {
			switch (child.localName) {
				case "p":
					children.push(this.parseParagraph(child));
					break;

				case "tbl":
					children.push(parseTableFn(child, this.options, this.tableCallbacks()));
					break;
				// TODO Structured Document
				case "sdt":
					children.push(...this.parseSdt(child));
					break;

				case "sectPr":
					// ignore,section property has parsed in parseDocumentFile
					break;

				default:
					if (this.options.debug) {
						console.warn(`DOCX:%c Unknown Body Element：${child.localName}`, 'color:red');
					}
			}
		});

		return children;
	}

	parseStylesFile(xstyles: Element): IDomStyle[] {
		return parseStylesFileFn(xstyles, this.options, { parseDefaultProperties: (e, s, c, h) => this.parseDefaultProperties(e, s, c, h) });
	}

	parseDefaultStyles(node: Element): IDomStyle {
		return parseDefaultStylesFn(node, this.options, { parseDefaultProperties: (e, s, c, h) => this.parseDefaultProperties(e, s, c, h) });
	}

	parseStyle(node: Element): IDomStyle {
		return parseStyleFn(node, this.options, { parseDefaultProperties: (e, s, c, h) => this.parseDefaultProperties(e, s, c, h) });
	}

	// TODO: multi-level nested table style rules are not yet applied
	parseTableStyle(node: Element): Ruleset[] {
		return parseTableStyleFn(node, this.options, { parseDefaultProperties: (e, s, c, h) => this.parseDefaultProperties(e, s, c, h) });
	}

	parseNumberingFile(xnums: Element): IDomNumbering[] {
		return parseNumberingFileFn(xnums, this.options, { parseDefaultProperties: (e, s, c, h) => this.parseDefaultProperties(e, s, c, h) });
	}

	parseNumberingPicBullet(elem: Element): NumberingPicBullet {
		return parseNumberingPicBulletFn(elem);
	}

	parseAbstractNumbering(node: Element, bullets: any[]): IDomNumbering[] {
		return parseAbstractNumberingFn(node, bullets, this.options, { parseDefaultProperties: (e, s, c, h) => this.parseDefaultProperties(e, s, c, h) });
	}

	parseNumberingLevel(id: string, node: Element, bullets: any[]): IDomNumbering {
		return parseNumberingLevelFn(id, node, bullets, this.options, { parseDefaultProperties: (e, s, c, h) => this.parseDefaultProperties(e, s, c, h) });
	}

	parseSdt(node: Element): OpenXmlElement[] {
		return parseSdtFn(node, { parseBodyElements: n => this.parseBodyElements(n) });
	}

	parseNotes(xmlDoc: Element, elemName: string, elemClass: any): any[] {
		return parseNotesFn(xmlDoc, elemName, elemClass, { parseBodyElements: n => this.parseBodyElements(n) });
	}

	parseComments(xmlDoc: Element): any[] {
		return parseCommentsFn(xmlDoc, { parseBodyElements: n => this.parseBodyElements(n) });
	}

	parseParagraph(node: Element): WmlParagraph {
		return parseParagraphFn(node, this.options, this.paragraphCallbacks());
	}

	parseRun(node: Element): WmlRun {
		return parseRunFn(node, this.options, this.runCallbacks());
	}

	private paragraphCallbacks() {
		return {
			parseRun: n => this.parseRun(n),
			parseMathElement: n => this.parseMathElement(n),
			parseBodyElements: n => this.parseBodyElements(n),
		};
	}

	private runCallbacks() {
		return {
			parseDrawing: n => parseDrawingFn(n, this.options, this.drawingCallbacks()),
			parseVmlPicture: n => parseVmlPictureFn(n, this.drawingCallbacks()),
			checkAlternateContent: n => checkAlternateContentFn(n),
		};
	}

	parseMathElement(elem: Element): OpenXmlElement {
		const callbacks: MathParserCallbacks = { parseRun: n => this.parseRun(n) };
		return parseMathElementFn(elem, this.options, callbacks);
	}

	parseDefaultProperties(elem: Element, style: Record<string, string> = null, childStyle: Record<string, string> = null, handler: (prop: Element) => boolean = null): Record<string, string> {
		return parseDefaultPropertiesFn(elem, this.options, style, childStyle, handler);
	}

	private drawingCallbacks(): DrawingParserCallbacks {
		return { parseBodyElements: n => this.parseBodyElements(n) };
	}

	private tableCallbacks(): TableParserCallbacks {
		return {
			parseParagraph: n => this.parseParagraph(n),
			parseTable: n => parseTableFn(n, this.options, this.tableCallbacks()),
			parseDefaultProperties: (e, s, c, h) => this.parseDefaultProperties(e, s, c, h),
		};
	}
}
