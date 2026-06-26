import { OpenXmlElement } from '@docx/ooxml/wordprocessingml/model/element';
import { WmlTable, WmlTableCell, WmlTableColumn } from '@docx/ooxml/wordprocessingml/model/table';
import { uuid } from '@docx/shared/utils';
import { Overflow, ChildrenType, createElement, appendChildren } from '@docx/rendering/dom/core/dom-utils';

export interface CellPos {
	col: number;
	row: number;
}

export type CellVerticalMergeType = Record<number, HTMLTableCellElement>;

// Mutable state shared across all table render calls (supports nested tables via stack)
export interface TableContext {
	tableVerticalMerges: Array<CellVerticalMergeType | null>;
	currentVerticalMerge: CellVerticalMergeType | null;
	tableCellPositions: Array<CellPos | null>;
	currentCellPosition: CellPos | null;
}

// Callbacks to the main renderer for cross-feature rendering
export interface TableRendererCallbacks {
	renderChildren(elem: OpenXmlElement, parent: HTMLElement): Promise<string>;
	appendChildren(parent: HTMLElement, children: Element): Promise<string>;
	renderClass(elem: OpenXmlElement, output: HTMLElement): void;
	renderStyleValues(style: Record<string, string>, output: HTMLElement): void;
}

// Render a <table> element
export async function renderTable(
	elem: WmlTable,
	parent: HTMLElement,
	ctx: TableContext,
	cbs: TableRendererCallbacks
): Promise<HTMLElement> {
	const oTable = createElement('table');
	// Assign a unique id to the table element
	oTable.dataset.uuid = uuid();
	// Push current cell position onto the stack (for nested tables)
	ctx.tableCellPositions.push(ctx.currentCellPosition);
	// Push current vertical-merge map onto the stack
	ctx.tableVerticalMerges.push(ctx.currentVerticalMerge);
	// Reset vertical merge map for this table
	ctx.currentVerticalMerge = {};
	// Reset cell position for this table
	ctx.currentCellPosition = { col: 0, row: 0 };
	// Apply CSS class
	cbs.renderClass(elem, oTable);
	// Apply inline styles
	cbs.renderStyleValues(elem.cssStyle, oTable);
	// Overflow detection: insert the table into the parent first
	let is_overflow: string;
	is_overflow = await cbs.appendChildren(parent, oTable);
	if (is_overflow === Overflow.TRUE) {
		oTable.dataset.overflow = Overflow.SELF;
		return oTable;
	}
	// Render <colgroup> if column widths are defined
	if (elem.columns) {
		renderTableColumns(elem.columns, oTable);
	}
	// Render rows and cells; record their overflow state
	oTable.dataset.overflow = await cbs.renderChildren(elem, oTable);
	// Restore previous vertical-merge map
	ctx.currentVerticalMerge = ctx.tableVerticalMerges.pop();
	// Restore previous cell position
	ctx.currentCellPosition = ctx.tableCellPositions.pop();
	return oTable;
}

// Render a <colgroup>/<col> block for fixed column widths
export function renderTableColumns(columns: WmlTableColumn[], parent: HTMLElement): HTMLElement {
	const oColGroup = createElement('colgroup');
	// Insert the colgroup (no overflow checking needed here)
	appendChildren(parent, oColGroup);
	for (const col of columns) {
		const oCol = createElement('col');
		if (col.width) {
			oCol.style.width = col.width;
		}
		appendChildren(oColGroup, oCol);
	}
	return oColGroup;
}

// Render a <tr> element
export async function renderTableRow(
	elem: OpenXmlElement,
	parent: HTMLElement,
	ctx: TableContext,
	cbs: TableRendererCallbacks
): Promise<HTMLElement> {
	const oTableRow = createElement('tr');
	// Reset column index at the start of each row
	ctx.currentCellPosition.col = 0;
	// Apply CSS class
	cbs.renderClass(elem, oTableRow);
	// Apply inline styles
	cbs.renderStyleValues(elem.cssStyle, oTableRow);
	// Overflow detection: insert the row into the parent first
	let is_overflow: string;
	is_overflow = await cbs.appendChildren(parent, oTableRow);
	if (is_overflow === Overflow.TRUE) {
		oTableRow.dataset.overflow = Overflow.SELF;
		return oTableRow;
	}
	// Render cells; record their overflow state
	oTableRow.dataset.overflow = await cbs.renderChildren(elem, oTableRow);
	// Advance row counter
	ctx.currentCellPosition.row++;
	return oTableRow;
}

// Render a <td> element
export async function renderTableCell(
	elem: WmlTableCell,
	parent: HTMLElement,
	ctx: TableContext,
	cbs: TableRendererCallbacks
): Promise<HTMLElement> {
	// Create a <td>; defaults colSpan=1, rowSpan=1
	const oTableCell = createElement('td');
	// Current column index for this cell
	const key = ctx.currentCellPosition.col;
	// Handle vertical cell merging
	if (elem.verticalMerge) {
		if (elem.verticalMerge === 'restart') {
			// Start/restart merged region
			ctx.currentVerticalMerge[key] = oTableCell;
			oTableCell.rowSpan = 1;
		} else if (ctx.currentVerticalMerge[key]) {
			// Continue merged region: extend the originating cell's rowSpan and hide this cell
			ctx.currentVerticalMerge[key].rowSpan += 1;
			oTableCell.style.display = 'none';
		}
	} else {
		ctx.currentVerticalMerge[key] = null;
	}
	// Apply CSS class
	cbs.renderClass(elem, oTableCell);
	// Apply inline styles
	cbs.renderStyleValues(elem.cssStyle, oTableCell);
	// Apply column span
	if (elem.span) {
		oTableCell.colSpan = elem.span;
	}
	// Advance column index by span width
	ctx.currentCellPosition.col += oTableCell.colSpan;
	// Overflow detection: insert the cell into the row first
	let is_overflow: string;
	is_overflow = await cbs.appendChildren(parent, oTableCell);
	if (is_overflow === Overflow.TRUE) {
		oTableCell.dataset.overflow = Overflow.SELF;
		return oTableCell;
	}
	// Render cell contents; record overflow state
	oTableCell.dataset.overflow = await cbs.renderChildren(elem, oTableCell);
	return oTableCell;
}
