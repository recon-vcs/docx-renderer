import { describe, expect, it } from 'vitest';
import { SectionProperties } from '../../src/ooxml/wordprocessingml/document/model/section';
import { LayoutRegion, PhysicalPage } from '../../src/rendering/pagination/model/layout-region';
import { buildPageLayoutContexts } from '../../src/rendering/pagination/model/page-numbering';

function makeSection(sectionId: string): SectionProperties {
	return { sectionId } as SectionProperties;
}

function makeRegion(section: SectionProperties): LayoutRegion {
	return {
		section,
		children: [],
		breakBefore: 'none',
	};
}

function makePage(
	pageNumber: number,
	regions: LayoutRegion[],
	sectionPageIndexes = new Map<string, number>(),
): PhysicalPage {
	return {
		pageNumber,
		regions,
		sectionPageIndexes,
	};
}

describe('buildPageLayoutContexts', () => {
	it('uses the last region as the active section for a page', () => {
		const sectionA = makeSection('section-a');
		const sectionB = makeSection('section-b');

		const contexts = buildPageLayoutContexts([
			makePage(1, [makeRegion(sectionA), makeRegion(sectionB)]),
		]);

		expect(contexts).toHaveLength(1);
		expect(contexts[0].activeSection).toBe(sectionB);
		expect(contexts[0].sectionId).toBe('section-b');
	});

	it('uses sectionPageIndexes when available', () => {
		const section = makeSection('section-a');
		const contexts = buildPageLayoutContexts([
			makePage(7, [makeRegion(section)], new Map([['section-a', 3]])),
		]);

		expect(contexts[0].physicalPageNumber).toBe(7);
		expect(contexts[0].sectionPageIndex).toBe(3);
		expect(contexts[0].isFirstSectionPage).toBe(false);
	});

	it('computes section-local first page from fallback indexes', () => {
		const sectionA = makeSection('section-a');
		const sectionB = makeSection('section-b');

		const contexts = buildPageLayoutContexts([
			makePage(1, [makeRegion(sectionA)]),
			makePage(2, [makeRegion(sectionA)]),
			makePage(3, [makeRegion(sectionB)]),
		]);

		expect(contexts.map(context => context.sectionPageIndex)).toEqual([0, 1, 0]);
		expect(contexts.map(context => context.isFirstSectionPage)).toEqual([true, false, true]);
	});

	it('marks even physical pages', () => {
		const section = makeSection('section-a');

		const contexts = buildPageLayoutContexts([
			makePage(1, [makeRegion(section)]),
			makePage(2, [makeRegion(section)]),
		]);

		expect(contexts.map(context => context.isEvenPage)).toEqual([false, true]);
	});
});
