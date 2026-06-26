import type { OpenXmlElement, WmlTable } from '@docx/ooxml/wordprocessingml/document/model/dom';
import type { WmlParagraph } from '@docx/ooxml/wordprocessingml/document/model/paragraph';
import type { WmlRun } from '@docx/ooxml/wordprocessingml/document/model/run';
import type { DocumentParserOptions } from './document-parser';

// Single context object threaded through all sub-parsers, replacing the
// multiple narrow callback-bag interfaces that previously existed.
export interface ParseContext {
	options: DocumentParserOptions;
	parseBodyElements(elem: Element): OpenXmlElement[];
	parseParagraph(elem: Element): WmlParagraph;
	parseTable(elem: Element): WmlTable;
	parseRun(elem: Element): WmlRun;
	parseMathElement(elem: Element): OpenXmlElement;
	parseDefaultProperties(
		elem: Element,
		style?: Record<string, string>,
		childStyle?: Record<string, string>,
		handler?: (prop: Element) => boolean,
	): Record<string, string>;
	parseDrawing(elem: Element): OpenXmlElement;
	parseVmlPicture(elem: Element): OpenXmlElement;
	checkAlternateContent(elem: Element): Element;
}
