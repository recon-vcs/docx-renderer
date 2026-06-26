import { XmlParser } from '@docx/xml/parsing/xml-parser';
import { Length } from "./common";
import { DocGridType, SectionProperties } from "./section";
import { ParagraphProperties } from "./paragraph";

export enum LineSpacingRule {
	// Minimum Line Height.
	AtLeast = "atLeast",
	// Automatically Determined Line Height. Default Value
	Auto = "auto",
	// Exact Line Height.
	Exact = "exact",
}

export interface SpacingBetweenLines {
	after?: Length;
	before?: Length;
	line?: number;
	lineRule?: LineSpacingRule;
}

export function parseSpacingBetweenLines(elem: Element, xml: XmlParser): SpacingBetweenLines {
	let spacing: SpacingBetweenLines = {
		lineRule: LineSpacingRule.Auto,
	};
	for (const attr of xml.attrs(elem)) {
		switch (attr.localName) {
			// Spacing after the last line in each paragraph
			case "after":
				spacing.after = xml.lengthAttr(elem, "after", undefined, '0pt');
				break;

			// TODO Automatically Determine Spacing after the last line in each paragraph
			case "afterAutospacing":
				break;

			// TODO Spacing Below Paragraph in Line Units
			case "afterLines":
				break;

			// Spacing before the first line in each paragraph
			case "before":
				spacing.before = xml.lengthAttr(elem, "before", undefined, '0pt');
				break;

			// TODO Automatically Determine Spacing before the first line in each paragraph
			case "beforeAutospacing":
				break;

			// TODO Spacing Above Paragraph in Line Units
			case "beforeLines":
				break;

			//  the amount of vertical spacing between lines of text within this paragraph.
			case "line":
				spacing.line = xml.intAttr(elem, "line", 0);
				break;

			// Type of Spacing Between Lines
			case "lineRule":
				spacing.lineRule = xml.attr(elem, "lineRule", LineSpacingRule.Auto) as LineSpacingRule;
				break;

			default:
				break;
		}
	}
	return spacing;
}

// TODO 处理AtLeast，行高不准确
export function parseLineSpacing(paragraphProperties: ParagraphProperties, sectionProperties?: SectionProperties): Record<string, any> {
	let { snapToGrid, spacing } = paragraphProperties;
	// original line spacing
	let lineSpacing: Record<string, any> = {};
	// original line number, kept outside the `if (spacing)` block so the
	// docGrid AtLeast correction below can still reference it.
	let originLine: number;

	if (spacing) {
		for (const key in spacing) {
			switch (key) {
				case 'line':
					originLine = spacing?.line;
					break;

				// Spacing after the last line in each paragraph
				case 'after':
					lineSpacing['margin-bottom'] = spacing[key];
					break;

				// Spacing before the first line in each paragraph
				case 'before':
					lineSpacing['margin-top'] = spacing[key];
					break;

				case 'afterLines':
					break;

				case 'beforeLines':
					break;

				case 'afterAutospacing':
					break;

				case 'beforeAutospacing':
					break;

				default:
					break;
			}
		}
		// Type of Spacing Between Lines
		switch (spacing?.lineRule) {
			// Automatically Determined Line Height.
			case "auto":
				lineSpacing['line-height'] = originLine / 240;
				break;

			// Minimum Line Height.
			case "atLeast":
				lineSpacing['line-height'] = `calc(100% + ${originLine / 20}pt)`;
				break;

			// Exact Line Height Override the DocGrid.
			case "exact":
				lineSpacing['line-height'] = `${originLine / 20}pt`;
				break;

			default:
				lineSpacing['line-height'] = originLine / 240;
				break;
		}

	}

	// if snapToGrid is false, return original line spacing
	if (snapToGrid === false) {
		return lineSpacing;
	}

	if (sectionProperties?.docGrid) {
		let { docGrid } = sectionProperties;
		switch (docGrid.type) {
			case DocGridType.Lines:
			case DocGridType.LinesAndChars:
			case DocGridType.SnapToChars:
				if (typeof lineSpacing['line-height'] === 'number') {
					lineSpacing['line-height'] = `${lineSpacing['line-height'] * docGrid.linePitch / 20}pt` as string;
				} else if (spacing?.lineRule === LineSpacingRule.AtLeast) {
					// AtLeast is a floor, not a fixed height: actual height is
					// whichever is larger between the explicit minimum and the
					// grid pitch. Previously skipped entirely because line-height
					// was already a string (calc(...)) by this point.
					lineSpacing['line-height'] = `max(${originLine / 20}pt, ${docGrid.linePitch / 20}pt)`;
				}
				break;
			case DocGridType.Default:
				return lineSpacing;

			default:
				console.warn(`DOCX:%c Unknown DocGrid Type：${docGrid.type}`, 'color:#f75607');
		}
	}
	return lineSpacing;
}
