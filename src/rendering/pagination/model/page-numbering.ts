import { SectionProperties } from '@docx/ooxml/wordprocessingml/document/model/section';
import { PhysicalPage } from './layout-region';

export interface PageLayoutContext {
	physicalPageNumber: number;
	activeSection: SectionProperties;
	sectionId: string;
	sectionPageIndex: number;
	isFirstSectionPage: boolean;
	isEvenPage: boolean;
}

function getActiveSection(page: PhysicalPage): SectionProperties {
	const activeRegion = page.regions[page.regions.length - 1];
	if (!activeRegion) {
		throw new Error('Cannot build a page layout context for a page with no regions.');
	}
	return activeRegion.section;
}

export function buildPageLayoutContexts(pages: PhysicalPage[]): PageLayoutContext[] {
	const fallbackSectionPageIndexes = new Map<string, number>();

	return pages.map((page) => {
		const activeSection = getActiveSection(page);
		const sectionId = activeSection.sectionId;
		const fallbackIndex = fallbackSectionPageIndexes.get(sectionId) ?? 0;
		const sectionPageIndex = page.sectionPageIndexes?.get(sectionId) ?? fallbackIndex;

		fallbackSectionPageIndexes.set(sectionId, fallbackIndex + 1);

		return {
			physicalPageNumber: page.pageNumber,
			activeSection,
			sectionId,
			sectionPageIndex,
			isFirstSectionPage: sectionPageIndex === 0,
			isEvenPage: page.pageNumber % 2 === 0,
		};
	});
}
