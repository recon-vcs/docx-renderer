import { DomType, OpenXmlElement, WmlTable, WmlTableCell, WmlTableColumn, WmlTableRow } from '@docx/ooxml/wordprocessingml/document/model/dom';
import xml from '@docx/xml/parsing/xml-parser';
import { xmlUtil, values } from '@docx/xml/parsing/parse-utils';
import type { TableParserContext } from './parse-context';

export function parseTable(
	node: Element,
	ctx: TableParserContext
): WmlTable {
	const result: WmlTable = { type: DomType.Table, children: [] };

	xmlUtil.foreach(node, c => {
		switch (c.localName) {
			case "tblPr":
				parseTableProperties(c, result, ctx);
				break;

			case "tblGrid":
				result.columns = parseTableColumns(c, ctx);
				break;

			case "tr":
				result.children.push(parseTableRow(c, ctx));
				break;

			default:
				if (ctx.options.debug) {
					console.warn(`DOCX:%c Unknown Table Element：${c.localName}`, 'color:#f75607');
				}
		}
	});

	return result;
}

export function parseTableColumns(node: Element, ctx: TableParserContext): WmlTableColumn[] {
	const result: WmlTableColumn[] = [];

	xmlUtil.foreach(node, n => {
		switch (n.localName) {
			case "gridCol":
				result.push({ width: xml.lengthAttr(n, "w") });
				break;

			// TODO grid revision info
			case "tblGridChange":
				break;

			default:
				if (ctx.options.debug) {
					console.warn(`DOCX:%c Unknown Table Columns Element：${n.localName}`, 'color:#f75607');
				}
		}
	});

	return result;
}

export function parseTableProperties(
	elem: Element,
	table: WmlTable,
	ctx: TableParserContext
): void {
	table.cssStyle = {};
	table.cellStyle = {};

	ctx.parseDefaultProperties(elem, table.cssStyle, table.cellStyle, c => {
		switch (c.localName) {
			case "tblStyle":
				table.styleName = xml.attr(c, "val");
				break;

			case "tblLook":
				table.className = values.classNameOftblLook(c);
				break;

			case "tblpPr":
				// Floating table position
				parseTablePosition(c, table, ctx);
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
			if (ctx.options.debug) {
				console.warn(`DOCX:%c Unknown Table Align：${table.cssStyle["text-align"]}`, 'color:#f75607');
			}
	}
}

// Floating table — implements text wrap around table
export function parseTablePosition(
	node: Element,
	table: WmlTable,
	ctx: TableParserContext
): void {
	// Floating tables displace subsequent elements; ignored by default
	if (ctx.options.ignoreTableWrap) {
		return;
	}

	const topFromText = xml.lengthAttr(node, "topFromText");
	const bottomFromText = xml.lengthAttr(node, "bottomFromText");
	const rightFromText = xml.lengthAttr(node, "rightFromText");
	const leftFromText = xml.lengthAttr(node, "leftFromText");

	table.cssStyle["float"] = 'left';
	table.cssStyle["margin-bottom"] = values.addSize(table.cssStyle["margin-bottom"], bottomFromText);
	table.cssStyle["margin-left"] = values.addSize(table.cssStyle["margin-left"], leftFromText);
	table.cssStyle["margin-right"] = values.addSize(table.cssStyle["margin-right"], rightFromText);
	table.cssStyle["margin-top"] = values.addSize(table.cssStyle["margin-top"], topFromText);
}

export function parseTableRow(
	node: Element,
	ctx: TableParserContext
): WmlTableRow {
	const result: WmlTableRow = { type: DomType.Row, children: [] };

	xmlUtil.foreach(node, c => {
		switch (c.localName) {
			case "tc":
				result.children.push(parseTableCell(c, ctx));
				break;

			case "trPr":
				parseTableRowProperties(c, result, ctx);
				break;

			default:
				if (ctx.options.debug) {
					console.warn(`DOCX:%c Unknown Table Row Element：${c.localName}`, 'color:#f75607');
				}
		}
	});

	return result;
}

export function parseTableRowProperties(
	elem: Element,
	row: WmlTableRow,
	ctx: TableParserContext
): void {
	row.cssStyle = ctx.parseDefaultProperties(elem, {}, null, c => {
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
	ctx: TableParserContext
): OpenXmlElement {
	const result: WmlTableCell = { type: DomType.Cell, children: [] };

	xmlUtil.foreach(node, c => {
		switch (c.localName) {
			case "tbl":
				result.children.push(ctx.parseTable(c));
				break;

			case "p":
				result.children.push(ctx.parseParagraph(c));
				break;

			case "tcPr":
				parseTableCellProperties(c, result, ctx);
				break;

			default:
				if (ctx.options.debug) {
					console.warn(`DOCX:%c Unknown Table Cell Element：${c.localName}`, 'color:#f75607');
				}
		}
	});

	return result;
}

export function parseTableCellProperties(
	elem: Element,
	cell: WmlTableCell,
	ctx: TableParserContext
): void {
	cell.cssStyle = ctx.parseDefaultProperties(elem, {}, null, c => {
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
