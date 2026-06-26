import type { OpenXmlElement, WmlTable } from '@docx/ooxml/wordprocessingml/document/model/dom';
import type { WmlParagraph } from '@docx/ooxml/wordprocessingml/document/model/paragraph';
import type { WmlRun } from '@docx/ooxml/wordprocessingml/document/model/run';
import type { DocumentParserOptions } from './document-parser';

export interface ParserOptionsContext {
	options: DocumentParserOptions;
}

export interface BodyParserContext extends ParserOptionsContext {
	parseBodyElements(elem: Element): OpenXmlElement[];
}

export interface ParagraphParserContext extends BodyParserContext {
	parseRun(elem: Element): WmlRun;
	parseMathElement(elem: Element): OpenXmlElement;
}

export interface RunParserContext extends ParserOptionsContext {
	parseDrawing(elem: Element): OpenXmlElement;
	parseVmlPicture(elem: Element): OpenXmlElement;
	checkAlternateContent(elem: Element): Element;
}

export interface MathParserContext extends ParserOptionsContext {
	parseRun(elem: Element): WmlRun;
}

export interface TableParserContext extends ParserOptionsContext {
	parseParagraph(elem: Element): WmlParagraph;
	parseTable(elem: Element): WmlTable;
	parseDefaultProperties(
		elem: Element,
		style?: Record<string, string>,
		childStyle?: Record<string, string>,
		handler?: (prop: Element) => boolean,
	): Record<string, string>;
}

export interface StyleParserContext extends ParserOptionsContext {
	parseDefaultProperties(
		elem: Element,
		style?: Record<string, string>,
		childStyle?: Record<string, string>,
		handler?: (prop: Element) => boolean,
	): Record<string, string>;
}

export interface NumberingParserContext extends StyleParserContext {}

export interface DrawingParserContext extends BodyParserContext {}

export interface ParseContext extends ParagraphParserContext, RunParserContext, MathParserContext, TableParserContext, StyleParserContext {
	parseDrawing(elem: Element): OpenXmlElement;
	parseVmlPicture(elem: Element): OpenXmlElement;
	checkAlternateContent(elem: Element): Element;
}
