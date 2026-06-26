import type { LayoutRegion, PhysicalPage } from '@docx/rendering/pagination/model/layout-region';
import { interpretRegionBreakBefore, needsParityBlankPage } from '@docx/rendering/pagination/model/breaks';

export type BuildPhysicalPagesOptions = Record<string, never>;

function createPage(pageNumber: number): PhysicalPage {
	return {
		regions: [],
		pageNumber,
		sectionPageIndexes: new Map(),
	};
}

function nextPageNumber(pages: PhysicalPage[]): number {
	const lastPage = pages[pages.length - 1];
	return lastPage ? lastPage.pageNumber + 1 : 1;
}

function addRegionToPage(
	page: PhysicalPage,
	region: LayoutRegion,
	sectionPageCounts: Map<string, number>,
): void {
	page.regions.push(region);

	const sectionId = region.section?.sectionId;
	if (!sectionId || page.sectionPageIndexes.has(sectionId)) {
		return;
	}

	const sectionPageIndex = sectionPageCounts.get(sectionId) ?? 0;
	sectionPageCounts.set(sectionId, sectionPageIndex + 1);
	page.sectionPageIndexes.set(sectionId, sectionPageIndex);
}

export function buildPhysicalPages(
	regions: LayoutRegion[],
	_options: BuildPhysicalPagesOptions = {},
): PhysicalPage[] {
	const pages: PhysicalPage[] = [];
	const sectionPageCounts = new Map<string, number>();

	for (const region of regions) {
		const breakPlacement = interpretRegionBreakBefore(region.breakBefore);
		let currentPage = pages[pages.length - 1];

		if (!currentPage) {
			currentPage = createPage(1);
			pages.push(currentPage);
		} else if (breakPlacement.startsNewPage) {
			let regionPageNumber = nextPageNumber(pages);

			if (needsParityBlankPage(regionPageNumber, breakPlacement.targetParity)) {
				pages.push(createPage(regionPageNumber));
				regionPageNumber += 1;
			}

			currentPage = createPage(regionPageNumber);
			pages.push(currentPage);
		}

		addRegionToPage(currentPage, region, sectionPageCounts);
	}

	return pages;
}
