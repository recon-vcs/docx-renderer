import { DomType, OpenXmlElement, WmlTableCell, WmlTableRow } from '@docx/ooxml/wordprocessingml/document/model/dom';
import * as _ from 'lodash-es';

function isSplitParagraph(elem: OpenXmlElement): boolean {
	const { breakIndex, children } = elem;
	if (!breakIndex) return false;
	if (!children || children.length === 0) return false;
	const i = [...breakIndex][0];
	if (i === 0) return isSplitParagraph(children[i]);
	return i < children.length;
}

export function splitElementsByBreakIndex(current: OpenXmlElement, next: OpenXmlElement): void {
	for (let i = 0; i < next?.children.length; i++) {
		const child = next.children[i];
		const { type, breakIndex, children } = child;

		if (!breakIndex) continue;
		if (!children || children.length === 0) continue;

		const copy: OpenXmlElement = _.cloneDeepWith(child, (value, key) => {
			if (key === 'parent') return null;
		});

		const count = breakIndex.size > 0 ? [...breakIndex][0] : children.length;

		switch (type) {
			case DomType.Table: {
				const tableHeaders: WmlTableRow[] = children.filter((row: WmlTableRow) => row.isHeader);
				const unbrokenChildren = children.splice(0, count);
				children[0].children.forEach((cell: WmlTableCell) => {
					if (cell.verticalMerge === 'continue') {
						cell.verticalMerge = 'restart';
					}
				});
				if (tableHeaders.length > 0 && tableHeaders.length < children.length) {
					children.unshift(...tableHeaders);
				}
				copy.children = unbrokenChildren;
				current.children.push(copy);
				break;
			}

			case DomType.Row:
				if ((child as WmlTableRow)?.isHeader) continue;
				current.children.push(copy);
				break;

			case DomType.Cell:
				copy.children = children.splice(0, count);
				current.children[i] = copy;
				break;

			case DomType.Paragraph:
				copy.children = children.splice(0, count);
				current.children.push(copy);
				if (isSplitParagraph(child)) {
					child.cssStyle['text-indent'] = '0';
				}
				break;

			default:
				copy.children = children.splice(0, count);
				current.children.push(copy);
		}

		if (type !== DomType.Row && breakIndex.size > 0) {
			child.breakIndex = undefined;
		}

		if (children.length > 0) {
			splitElementsByBreakIndex(copy, child);
		}
	}
}
