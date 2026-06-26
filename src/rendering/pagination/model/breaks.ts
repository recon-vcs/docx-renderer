import type { RegionBreakBefore } from './layout-region';

export type PageParity = 'even' | 'odd';

export interface RegionBreakPlacement {
	breakBefore: RegionBreakBefore;
	startsNewPage: boolean;
	targetParity: PageParity | null;
}

export function normalizeRegionBreakBefore(
	breakBefore: RegionBreakBefore | null | undefined,
): RegionBreakBefore {
	return breakBefore ?? 'none';
}

export function interpretRegionBreakBefore(
	breakBefore: RegionBreakBefore | null | undefined,
): RegionBreakPlacement {
	const normalizedBreak = normalizeRegionBreakBefore(breakBefore);

	switch (normalizedBreak) {
		case 'page':
			return {
				breakBefore: normalizedBreak,
				startsNewPage: true,
				targetParity: null,
			};
		case 'evenPage':
			return {
				breakBefore: normalizedBreak,
				startsNewPage: true,
				targetParity: 'even',
			};
		case 'oddPage':
			return {
				breakBefore: normalizedBreak,
				startsNewPage: true,
				targetParity: 'odd',
			};
		case 'none':
		case 'column':
			return {
				breakBefore: normalizedBreak,
				startsNewPage: false,
				targetParity: null,
			};
	}
}

export function pageMatchesParity(pageNumber: number, parity: PageParity): boolean {
	return parity === 'even'
		? pageNumber % 2 === 0
		: pageNumber % 2 !== 0;
}

export function needsParityBlankPage(
	nextPageNumber: number,
	targetParity: PageParity | null,
): boolean {
	return targetParity !== null && !pageMatchesParity(nextPageNumber, targetParity);
}
