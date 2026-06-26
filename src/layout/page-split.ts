import { OpenXmlElement } from '../document/dom';
import { Page, PageProps } from '../document/page';
import { SectionProperties } from '../document/section';
import { LayoutRegion, PhysicalPage } from './layout-region';
import { PageLayoutContext } from './page-numbering';
import { splitElementsByBreakIndex } from '../render/split-by-break';

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

function buildPhysicalPage(
	regions: LayoutRegion[] | undefined,
	context: PageLayoutContext | undefined,
): PhysicalPage | undefined {
	if (!regions || !context) {
		return undefined;
	}

	return {
		regions,
		pageNumber: context.physicalPageNumber,
		sectionPageIndexes: new Map([[context.sectionId, context.sectionPageIndex]]),
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
		physicalPage: buildPhysicalPage(nextRegions, layoutContext),
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
	currentPage.checkingOverflow = false;
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

	const currentWrapper = new Page({
		sectProps: currentRegion.section,
		children: currentRegion.children,
	} as PageProps);
	const nextWrapper = new Page({
		sectProps: nextRegion.section,
		children: nextRegion.children,
	} as PageProps);

	splitElementsByBreakIndex(currentWrapper, nextWrapper);
	currentRegion.children = currentWrapper.children;
	nextRegion.children = nextWrapper.children;

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
	currentPage.checkingOverflow = false;
	currentPage.physicalPage = currentPage.layoutContext
		? buildPhysicalPage(currentRegions, currentPage.layoutContext)
		: currentPage.physicalPage;
	pages[pageIndex] = currentPage;
	pages.splice(pageIndex + 1, 0, nextPage);

	return { updatedCurrentPage: currentPage, nextPage, nextPageIndex: pageIndex + 1 };
}
