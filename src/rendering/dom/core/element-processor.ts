import { DomType, OpenXmlElement } from '@docx/ooxml/wordprocessingml/document/model/dom';
import { WmlTable } from '@docx/ooxml/wordprocessingml/model/table';

export function copyStyleProperties(
	input: Record<string, string>,
	output: Record<string, string> | null,
	attrs: string[] | null = null,
): Record<string, string> {
	if (!input) {
		return output;
	}
	if (output == null) {
		output = {};
	}
	if (attrs == null) {
		attrs = Object.getOwnPropertyNames(input);
	}

	for (const key of attrs) {
		if (input.hasOwnProperty(key) && !output.hasOwnProperty(key))
			output[key] = input[key];
	}

	return output;
}

/** Copy table-level cell styles down to individual cells. */
export function processTable(table: WmlTable): void {
	for (const r of table.children) {
		for (const c of r.children) {
			c.cssStyle = copyStyleProperties(table.cellStyle, c.cssStyle, [
				'border-left',
				'border-right',
				'border-top',
				'border-bottom',
				'padding-left',
				'padding-right',
				'padding-top',
				'padding-bottom',
			]);
		}
	}
}

/**
 * Walk the element tree and propagate table cell styles.
 * Parent links must already be set (via linkParents) before rendering.
 */
export function processElement(element: OpenXmlElement): void {
	if (element.children) {
		for (const e of element.children) {
			if (e.type === DomType.Table) {
				processTable(e as WmlTable);
			}
			processElement(e);
		}
	}
}

/**
 * Walk the element tree and set parent references on every child.
 * Call this once after an element tree is constructed or re-assembled
 * (e.g. after parsing or after an overflow split).
 */
export function linkParents(element: OpenXmlElement): void {
	if (element.children) {
		for (const e of element.children) {
			e.parent = element;
			linkParents(e);
		}
	}
}
