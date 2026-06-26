import { OpenXmlElement } from '@docx/ooxml/wordprocessingml/document/model/dom';
import { Page, PageProps } from './page';
import { SectionProperties } from '@docx/ooxml/wordprocessingml/document/model/section';
import { LayoutRegion } from '@docx/rendering/pagination/model/layout-region';
import { PageLayoutContext } from '@docx/rendering/pagination/model/page-numbering';
import { SplitTarget, splitElementsByBreakIndex } from './split-by-break';

export interface PageSplitResult {
	updatedCurrentPage: Page;
	nextPage: Page;
	nextPageIndex: number;
}

function activeSection(regions: LayoutRegion[] | undefined, fallback: SectionProperties): SectionProperties {
	return regions?.[regions.length - 1]?.section ?? fallback;
}

function pageNumberForNextPage(currentPage: Page, nextPageIndex: number): number {
	return currentPage.layoutContext?.physicalPageNumber
		? currentPage.layoutContext.physicalPageNumber + 1
		: nextPageIndex + 1;
}

function buildContext(
	currentPage: Page,
	nextPageIndex: number,
	nextSection: SectionProperties,
): PageLayoutContext | undefined {
	const sectionId = nextSection.sectionId;
	if (!sectionId) {
		return undefined;
	}

	const physicalPageNumber = pageNumberForNextPage(currentPage, nextPageIndex);
	const currentContext = currentPage.layoutContext;
	const sameSection = currentContext?.sectionId === sectionId;
	const sectionPageIndex = sameSection ? currentContext.sectionPageIndex + 1 : 0;

	return {
		physicalPageNumber,
		activeSection: nextSection,
		sectionId,
		sectionPageIndex,
		isFirstSectionPage: sectionPageIndex === 0,
		isEvenPage: physicalPageNumber % 2 === 0,
	};
}

function makeNextPage(
	currentPage: Page,
	pageIndex: number,
	nextChildren: OpenXmlElement[],
	nextRegions?: LayoutRegion[],
): Page {
	const nextPageIndex = pageIndex + 1;
	const nextSection = activeSection(nextRegions, currentPage.sectProps);
	const layoutContext = buildContext(currentPage, nextPageIndex, nextSection);

	return new Page({
		sectProps: nextSection,
		children: nextChildren,
		regions: nextRegions,
		layoutContext,
	} as PageProps);
}

/**
 * When a top-level element (level 2) overflows, split the current page into
 * current (already rendered) and next (remaining children).
 */
export function splitOnOverflow(
	currentPage: Page,
	pages: Page[],
	pageIndex: number,
	overflowIndex: number,
): PageSplitResult {
	const { sectProps, children: currentChildren } = currentPage;
	const nextPageChildren: OpenXmlElement[] = currentChildren.splice(overflowIndex);
	const nextPage = makeNextPage(currentPage, pageIndex, nextPageChildren);
	splitElementsByBreakIndex(currentPage, nextPage);
	currentPage.isSplit = true;
	pages[pageIndex] = currentPage;
	pages.splice(pageIndex + 1, 0, nextPage);
	return { updatedCurrentPage: currentPage, nextPage, nextPageIndex: pageIndex + 1 };
}

export function splitRegionOnOverflow(
	currentPage: Page,
	pages: Page[],
	pageIndex: number,
	regionIndex: number,
	overflowIndex: number,
): PageSplitResult {
	const regions = currentPage.regions ?? [];
	const region = regions[regionIndex];

	if (!region) {
		return splitOnOverflow(currentPage, pages, pageIndex, overflowIndex);
	}

	const currentRegion = {
		...region,
		children: region.children.slice(0, overflowIndex),
	};
	const nextRegion = {
		...region,
		breakBefore: 'page' as const,
		children: region.children.slice(overflowIndex),
	};

	const currentSplit: SplitTarget = { children: currentRegion.children, breakIndex: currentPage.breakIndex };
	const nextSplit: SplitTarget = { children: nextRegion.children };
	splitElementsByBreakIndex(currentSplit, nextSplit);
	currentRegion.children = currentSplit.children;
	nextRegion.children = nextSplit.children;

	const currentRegions = [
		...regions.slice(0, regionIndex),
		currentRegion,
	].filter(item => item.children.length > 0);
	const nextRegions = [
		...(nextRegion.children.length > 0 ? [nextRegion] : []),
		...regions.slice(regionIndex + 1),
	];

	currentPage.regions = currentRegions;
	currentPage.children = currentRegions.flatMap(item => item.children);
	currentPage.sectProps = activeSection(currentRegions, currentPage.sectProps);

	const nextPage = makeNextPage(
		currentPage,
		pageIndex,
		nextRegions.flatMap(item => item.children),
		nextRegions,
	);

	currentPage.isSplit = true;
	pages[pageIndex] = currentPage;
	pages.splice(pageIndex + 1, 0, nextPage);

	return { updatedCurrentPage: currentPage, nextPage, nextPageIndex: pageIndex + 1 };
}
