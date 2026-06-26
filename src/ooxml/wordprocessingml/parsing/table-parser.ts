import { DomType, OpenXmlElement, WmlTable, WmlTableCell, WmlTableColumn, WmlTableRow } from '@docx/ooxml/wordprocessingml/document/model/dom';
import type { WmlParagraph } from '@docx/ooxml/wordprocessingml/document/model/paragraph';
import type { DocumentParserOptions } from '@docx/ooxml/wordprocessingml/parsing/document-parser';
import xml from '@docx/xml/parsing/xml-parser';
import { xmlUtil, values } from '@docx/xml/parsing/parse-utils';

// Callbacks required for cross-module calls from parseTableCell
export interface TableParserCallbacks {
	parseParagraph(node: Element): WmlParagraph;
	parseTable(node: Element): WmlTable;
	parseDefaultProperties(
		elem: Element,
		style?: Record<string, string>,
		childStyle?: Record<string, string>,
		handler?: (prop: Element) => boolean
	): Record<string, string>;
}

export function parseTable(
	node: Element,
	options: DocumentParserOptions,
	callbacks: TableParserCallbacks
): WmlTable {
	let result: WmlTable = { type: DomType.Table, children: [] };

	xmlUtil.foreach(node, c => {
		switch (c.localName) {
			case "tblPr":
				parseTableProperties(c, result, options, callbacks);
				break;

			case "tblGrid":
				result.columns = parseTableColumns(c, options);
				break;

			case "tr":
				result.children.push(parseTableRow(c, options, callbacks));
				break;

			default:
				if (options.debug) {
					console.warn(`DOCX:%c Unknown Table Element：${c.localName}`, 'color:#f75607');
				}
		}
	});

	return result;
}

export function parseTableColumns(node: Element, options: DocumentParserOptions): WmlTableColumn[] {
	let result: WmlTableColumn[] = [];

	xmlUtil.foreach(node, n => {
		switch (n.localName) {
			case "gridCol":
				result.push({ width: xml.lengthAttr(n, "w") });
				break;

			// TODO grid revision info
			case "tblGridChange":
				break;

			default:
				if (options.debug) {
					console.warn(`DOCX:%c Unknown Table Columns Element：${n.localName}`, 'color:#f75607');
				}
		}
	});

	return result;
}

export function parseTableProperties(
	elem: Element,
	table: WmlTable,
	options: DocumentParserOptions,
	callbacks: TableParserCallbacks
): void {
	table.cssStyle = {};
	table.cellStyle = {};

	callbacks.parseDefaultProperties(elem, table.cssStyle, table.cellStyle, c => {
		switch (c.localName) {
			case "tblStyle":
				table.styleName = xml.attr(c, "val");
				break;

			case "tblLook":
				table.className = values.classNameOftblLook(c);
				break;

			case "tblpPr":
				// Floating table position
				parseTablePosition(c, table, options);
				break;

			case "tblStyleColBandSize":
				table.colBandSize = xml.intAttr(c, "val");
				break;

			case "tblStyleRowBandSize":
				table.rowBandSize = xml.intAttr(c, "val");
				break;

			default:
				return false;
		}

		return true;
	});

	switch (table.cssStyle["text-align"]) {
		case "center":
			delete table.cssStyle["text-align"];
			table.cssStyle["margin-left"] = "auto";
			table.cssStyle["margin-right"] = "auto";
			break;

		case "right":
			delete table.cssStyle["text-align"];
			table.cssStyle["margin-left"] = "auto";
			break;

		default:
			if (options.debug) {
				console.warn(`DOCX:%c Unknown Table Align：${table.cssStyle["text-align"]}`, 'color:#f75607');
			}
	}
}

// Floating table — implements text wrap around table
export function parseTablePosition(
	node: Element,
	table: WmlTable,
	options: DocumentParserOptions
): void {
	// Floating tables displace subsequent elements; ignored by default
	if (options.ignoreTableWrap) {
		return;
	}

	let topFromText = xml.lengthAttr(node, "topFromText");
	let bottomFromText = xml.lengthAttr(node, "bottomFromText");
	let rightFromText = xml.lengthAttr(node, "rightFromText");
	let leftFromText = xml.lengthAttr(node, "leftFromText");

	table.cssStyle["float"] = 'left';
	table.cssStyle["margin-bottom"] = values.addSize(table.cssStyle["margin-bottom"], bottomFromText);
	table.cssStyle["margin-left"] = values.addSize(table.cssStyle["margin-left"], leftFromText);
	table.cssStyle["margin-right"] = values.addSize(table.cssStyle["margin-right"], rightFromText);
	table.cssStyle["margin-top"] = values.addSize(table.cssStyle["margin-top"], topFromText);
}

export function parseTableRow(
	node: Element,
	options: DocumentParserOptions,
	callbacks: TableParserCallbacks
): WmlTableRow {
	let result: WmlTableRow = { type: DomType.Row, children: [] };

	xmlUtil.foreach(node, c => {
		switch (c.localName) {
			case "tc":
				result.children.push(parseTableCell(c, options, callbacks));
				break;

			case "trPr":
				parseTableRowProperties(c, result, options, callbacks);
				break;

			default:
				if (options.debug) {
					console.warn(`DOCX:%c Unknown Table Row Element：${c.localName}`, 'color:#f75607');
				}
		}
	});

	return result;
}

export function parseTableRowProperties(
	elem: Element,
	row: WmlTableRow,
	options: DocumentParserOptions,
	callbacks: TableParserCallbacks
): void {
	row.cssStyle = callbacks.parseDefaultProperties(elem, {}, null, c => {
		switch (c.localName) {
			case "cnfStyle":
				row.className = values.classNameOfCnfStyle(c);
				break;

			// Repeat Table Row on Every New Page
			case "tblHeader":
				row.isHeader = xml.boolAttr(c, "val", true);
				break;

			default:
				return false;
		}

		return true;
	});
}

export function parseTableCell(
	node: Element,
	options: DocumentParserOptions,
	callbacks: TableParserCallbacks
): OpenXmlElement {
	let result: WmlTableCell = { type: DomType.Cell, children: [] };

	xmlUtil.foreach(node, c => {
		switch (c.localName) {
			case "tbl":
				result.children.push(callbacks.parseTable(c));
				break;

			case "p":
				result.children.push(callbacks.parseParagraph(c));
				break;

			case "tcPr":
				parseTableCellProperties(c, result, options, callbacks);
				break;

			default:
				if (options.debug) {
					console.warn(`DOCX:%c Unknown Table Cell Element：${c.localName}`, 'color:#f75607');
				}
		}
	});

	return result;
}

export function parseTableCellProperties(
	elem: Element,
	cell: WmlTableCell,
	options: DocumentParserOptions,
	callbacks: TableParserCallbacks
): void {
	cell.cssStyle = callbacks.parseDefaultProperties(elem, {}, null, c => {
		switch (c.localName) {
			case "gridSpan":
				cell.span = xml.intAttr(c, "val", null);
				break;

			case "vMerge":
				cell.verticalMerge = xml.attr(c, "val") ?? "continue";
				break;

			case "cnfStyle":
				cell.className = values.classNameOfCnfStyle(c);
				break;

			default:
				return false;
		}

		return true;
	});
}
