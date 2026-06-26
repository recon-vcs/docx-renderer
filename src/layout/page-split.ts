import { OpenXmlElement } from '../document/dom';
import { Page, PageProps } from '../document/page';
import { splitElementsByBreakIndex } from '../render/split-by-break';

export interface PageSplitResult {
	updatedCurrentPage: Page;
	nextPage: Page;
	nextPageIndex: number;
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
	const nextPage = new Page({ sectProps, children: nextPageChildren } as PageProps);
	splitElementsByBreakIndex(currentPage, nextPage);
	currentPage.isSplit = true;
	currentPage.checkingOverflow = false;
	pages[pageIndex] = currentPage;
	pages.splice(pageIndex + 1, 0, nextPage);
	return { updatedCurrentPage: currentPage, nextPage, nextPageIndex: pageIndex + 1 };
}
