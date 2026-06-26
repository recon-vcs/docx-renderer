import { BreakType, DomType, OpenXmlElement, WmlBreak } from '@docx/ooxml/wordprocessingml/model/element';
import { LayoutRegion, LayoutRegionHint, RegionBreakBefore } from '@docx/rendering/pagination/model/layout-region';

type ExplicitBreakKind = 'page' | 'column';

interface ElementBreakScan {
	explicitBreak?: ExplicitBreakKind;
	hints: LayoutRegionHint[];
}

function explicitBreakToRegionBreak(kind: ExplicitBreakKind): RegionBreakBefore {
	return kind === 'page' ? 'page' : 'column';
}

function scanElementBreaks(element: OpenXmlElement, path: number[] = []): ElementBreakScan {
	const hints: LayoutRegionHint[] = [];

	if (element.type === DomType.LastRenderedPageBreak) {
		hints.push({ kind: 'lastRenderedPageBreak', path });
	}

	if (element.type === DomType.Break) {
		const breakType = (element as WmlBreak).break;
		if (breakType === BreakType.Page) {
			return { explicitBreak: 'page', hints };
		}
		if (breakType === BreakType.Column) {
			return { explicitBreak: 'column', hints };
		}
	}

	for (let i = 0; i < (element.children?.length ?? 0); i++) {
		const childScan = scanElementBreaks(element.children[i], [...path, i]);
		hints.push(...childScan.hints);

		if (childScan.explicitBreak) {
			return { explicitBreak: childScan.explicitBreak, hints };
		}
	}

	return { hints };
}

function makeRegion(
	source: LayoutRegion,
	children: OpenXmlElement[],
	breakBefore: RegionBreakBefore,
	hints: LayoutRegionHint[],
): LayoutRegion {
	return {
		section: source.section,
		children,
		breakBefore,
		...(hints.length > 0 ? { hints } : {}),
	};
}

export function splitRegionsByExplicitBreaks(regions: LayoutRegion[]): LayoutRegion[] {
	const result: LayoutRegion[] = [];

	for (const region of regions) {
		let currentChildren: OpenXmlElement[] = [];
		let currentBreakBefore = region.breakBefore;
		let currentHints: LayoutRegionHint[] = [...(region.hints ?? [])];

		for (let i = 0; i < region.children.length; i++) {
			const child = region.children[i];
			const scan = scanElementBreaks(child, [i]);

			// A lastRenderedPageBreak inside an element means Word started a new page
			// at that element. Split BEFORE it so the element lands on the next page.
			if (
				scan.hints.some(h => h.kind === 'lastRenderedPageBreak') &&
				currentChildren.length > 0
			) {
				result.push(makeRegion(region, currentChildren, currentBreakBefore, currentHints));
				currentChildren = [];
				currentHints = [];
				currentBreakBefore = 'page';
			}

			currentChildren.push(child);
			currentHints.push(...scan.hints);

			if (!scan.explicitBreak) {
				continue;
			}

			result.push(makeRegion(region, currentChildren, currentBreakBefore, currentHints));
			currentChildren = [];
			currentHints = [];
			currentBreakBefore = explicitBreakToRegionBreak(scan.explicitBreak);
		}

		if (currentChildren.length > 0 || result.length === 0) {
			result.push(makeRegion(region, currentChildren, currentBreakBefore, currentHints));
		}
	}

	return result;
}
