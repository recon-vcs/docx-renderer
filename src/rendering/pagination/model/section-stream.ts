import { DomType, OpenXmlElement } from '@docx/ooxml/wordprocessingml/model/element';
import { SectionProperties, SectionType } from '@docx/ooxml/wordprocessingml/document/model/section';
import { LayoutRegion, RegionBreakBefore } from './layout-region';

function sectionTypeToBreak(type: SectionType | undefined): RegionBreakBefore {
	switch (type) {
		case SectionType.Continuous: return 'none';
		case SectionType.NextColumn: return 'column';
		case SectionType.EvenPage: return 'evenPage';
		case SectionType.OddPage: return 'oddPage';
		default: return 'page';
	}
}

function getParagraphSectionProps(el: OpenXmlElement): SectionProperties | undefined {
	if (el.type !== DomType.Paragraph) return undefined;
	return (el.props as { sectionProperties?: SectionProperties } | undefined)?.sectionProperties;
}

/**
 * Converts body children into a linear stream of layout regions.
 *
 * Word model: a sectPr embedded in a paragraph marks the END of that section.
 * The sectPr contains the properties of THAT section. The section's `type`
 * tells how the NEXT section begins. The body-level sectPr is the final section.
 */
export function buildSectionStream(
	bodyChildren: OpenXmlElement[],
	rootSectProps: SectionProperties,
): LayoutRegion[] {
	const regions: LayoutRegion[] = [];
	let currentChildren: OpenXmlElement[] = [];
	let pendingBreak: RegionBreakBefore = 'none';

	for (const el of bodyChildren) {
		currentChildren.push(el);
		const sectProps = getParagraphSectionProps(el);
		if (sectProps) {
			regions.push({
				section: sectProps,
				children: currentChildren,
				breakBefore: pendingBreak,
			});
			currentChildren = [];
			pendingBreak = sectionTypeToBreak(sectProps.type);
		}
	}

	if (currentChildren.length > 0 || regions.length === 0) {
		regions.push({
			section: rootSectProps,
			children: currentChildren,
			breakBefore: pendingBreak,
		});
	}

	return regions;
}
