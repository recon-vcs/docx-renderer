import { DomType, OpenXmlElement } from '../document/dom';
import { WmlTable } from '../model/table';

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

// 处理表格style样式
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

// 递归明确元素parent父级关系
export function processElement(element: OpenXmlElement): void {
	if (element.children) {
		for (const e of element.children) {
			e.parent = element;
			e.level = element?.level + 1;
			if (e.type == DomType.Table) {
				processTable(e as WmlTable);
				processElement(e);
			} else {
				processElement(e);
			}
		}
	}
}
