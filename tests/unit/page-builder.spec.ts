import { describe, expect, it } from 'vitest';
import { SectionProperties } from '../../src/ooxml/wordprocessingml/document/model/section';
import { buildPhysicalPages } from '../../src/rendering/pagination/core/page-builder';
import { LayoutRegion, RegionBreakBefore } from '../../src/rendering/pagination/model/layout-region';

function makeSection(sectionId?: string): SectionProperties {
	return { sectionId } as SectionProperties;
}

function makeRegion(
	sectionId: string | undefined,
	breakBefore: RegionBreakBefore,
): LayoutRegion {
	return {
		section: makeSection(sectionId),
		children: [],
		breakBefore,
	};
}

describe('buildPhysicalPages', () => {
	it('returns no pages for an empty region stream', () => {
		expect(buildPhysicalPages([])).toEqual([]);
	});

	it('starts the first region on physical page 1', () => {
		const region = makeRegion('s1', 'none');
		const pages = buildPhysicalPages([region]);

		expect(pages).toHaveLength(1);
		expect(pages[0].pageNumber).toBe(1);
		expect(pages[0].regions).toEqual([region]);
		expect(pages[0].sectionPageIndexes).toEqual(new Map([['s1', 0]]));
	});

	it('appends none and column breaks to the current page', () => {
		const first = makeRegion('s1', 'none');
		const continuous = makeRegion('s2', 'none');
		const column = makeRegion('s3', 'column');

		const pages = buildPhysicalPages([first, continuous, column]);

		expect(pages).toHaveLength(1);
		expect(pages[0].regions).toEqual([first, continuous, column]);
		expect(column.breakBefore).toBe('column');
		expect(pages[0].sectionPageIndexes).toEqual(new Map([
			['s1', 0],
			['s2', 0],
			['s3', 0],
		]));
	});

	it('starts a new physical page for page breaks', () => {
		const first = makeRegion('s1', 'none');
		const second = makeRegion('s2', 'page');

		const pages = buildPhysicalPages([first, second]);

		expect(pages).toHaveLength(2);
		expect(pages[0].pageNumber).toBe(1);
		expect(pages[0].regions).toEqual([first]);
		expect(pages[1].pageNumber).toBe(2);
		expect(pages[1].regions).toEqual([second]);
	});

	it('starts evenPage regions on an even physical page', () => {
		const first = makeRegion('s1', 'none');
		const second = makeRegion('s2', 'evenPage');

		const pages = buildPhysicalPages([first, second]);

		expect(pages).toHaveLength(2);
		expect(pages[1].pageNumber).toBe(2);
		expect(pages[1].regions).toEqual([second]);
	});

	it('inserts a blank page when evenPage would otherwise start on an odd page', () => {
		const first = makeRegion('s1', 'none');
		const second = makeRegion('s2', 'page');
		const third = makeRegion('s3', 'evenPage');

		const pages = buildPhysicalPages([first, second, third]);

		expect(pages.map(page => page.pageNumber)).toEqual([1, 2, 3, 4]);
		expect(pages[2].regions).toEqual([]);
		expect(pages[2].sectionPageIndexes).toEqual(new Map());
		expect(pages[3].regions).toEqual([third]);
	});

	it('inserts a blank page when oddPage would otherwise start on an even page', () => {
		const first = makeRegion('s1', 'none');
		const second = makeRegion('s2', 'oddPage');

		const pages = buildPhysicalPages([first, second]);

		expect(pages.map(page => page.pageNumber)).toEqual([1, 2, 3]);
		expect(pages[1].regions).toEqual([]);
		expect(pages[1].sectionPageIndexes).toEqual(new Map());
		expect(pages[2].regions).toEqual([second]);
	});

	it('does not insert a blank page when oddPage already targets an odd page', () => {
		const first = makeRegion('s1', 'none');
		const second = makeRegion('s2', 'page');
		const third = makeRegion('s3', 'oddPage');

		const pages = buildPhysicalPages([first, second, third]);

		expect(pages.map(page => page.pageNumber)).toEqual([1, 2, 3]);
		expect(pages[2].regions).toEqual([third]);
	});

	it('tracks section-local page indexes per physical page', () => {
		const pageOneFirst = makeRegion('s1', 'none');
		const pageOneSecond = makeRegion('s1', 'none');
		const pageTwo = makeRegion('s1', 'page');
		const otherSection = makeRegion('s2', 'none');

		const pages = buildPhysicalPages([
			pageOneFirst,
			pageOneSecond,
			pageTwo,
			otherSection,
		]);

		expect(pages).toHaveLength(2);
		expect(pages[0].sectionPageIndexes).toEqual(new Map([['s1', 0]]));
		expect(pages[1].sectionPageIndexes).toEqual(new Map([
			['s1', 1],
			['s2', 0],
		]));
	});

	it('skips section page indexes when a section id is not available', () => {
		const region = makeRegion(undefined, 'none');
		const pages = buildPhysicalPages([region]);

		expect(pages[0].sectionPageIndexes).toEqual(new Map());
	});
});
