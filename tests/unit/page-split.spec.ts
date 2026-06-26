import { describe, expect, it } from 'vitest';
import { DomType, OpenXmlElement } from '../../src/ooxml/wordprocessingml/document/model/dom';
import { Page } from '../../src/ooxml/wordprocessingml/document/model/page';
import { SectionProperties } from '../../src/ooxml/wordprocessingml/document/model/section';
import { LayoutRegion } from '../../src/rendering/pagination/model/layout-region';
import { splitRegionOnOverflow } from '../../src/rendering/pagination/model/page-split';

function section(sectionId: string): SectionProperties {
	return { sectionId } as SectionProperties;
}

function paragraph(name: string): OpenXmlElement {
	return {
		type: DomType.Paragraph,
		children: [],
		sourcePath: name,
	};
}

function region(sectionProps: SectionProperties, children: OpenXmlElement[]): LayoutRegion {
	return {
		section: sectionProps,
		children,
		breakBefore: 'none',
	};
}

describe('splitRegionOnOverflow', () => {
	it('splits a region at the overflowing block boundary', () => {
		const sectionProps = section('s1');
		const first = paragraph('first');
		const second = paragraph('second');
		const third = paragraph('third');
		const currentRegion = region(sectionProps, [first, second, third]);
		const page = new Page({
			sectProps: sectionProps,
			children: [first, second, third],
			regions: [currentRegion],
			layoutContext: {
				physicalPageNumber: 1,
				activeSection: sectionProps,
				sectionId: 's1',
				sectionPageIndex: 0,
				isFirstSectionPage: true,
				isEvenPage: false,
			},
		});
		const pages = [page];

		const result = splitRegionOnOverflow(page, pages, 0, 0, 1);

		expect(result.updatedCurrentPage.children).toEqual([first]);
		expect(result.nextPage.children).toEqual([second, third]);
		expect(result.nextPage.regions?.[0].breakBefore).toBe('page');
		expect(result.nextPage.layoutContext?.physicalPageNumber).toBe(2);
		expect(result.nextPage.layoutContext?.sectionPageIndex).toBe(1);
		expect(pages).toHaveLength(2);
	});

	it('moves later regions to the next page with the overflowing region remainder', () => {
		const firstSection = section('s1');
		const secondSection = section('s2');
		const first = paragraph('first');
		const second = paragraph('second');
		const third = paragraph('third');
		const fourth = paragraph('fourth');
		const firstRegion = region(firstSection, [first]);
		const secondRegion = region(secondSection, [second, third]);
		const thirdRegion = region(secondSection, [fourth]);
		const page = new Page({
			sectProps: secondSection,
			children: [first, second, third, fourth],
			regions: [firstRegion, secondRegion, thirdRegion],
			layoutContext: {
				physicalPageNumber: 1,
				activeSection: secondSection,
				sectionId: 's2',
				sectionPageIndex: 0,
				isFirstSectionPage: true,
				isEvenPage: false,
			},
		});
		const pages = [page];

		const result = splitRegionOnOverflow(page, pages, 0, 1, 1);

		expect(result.updatedCurrentPage.regions?.map(item => item.children)).toEqual([
			[first],
			[second],
		]);
		expect(result.nextPage.regions?.map(item => item.children)).toEqual([
			[third],
			[fourth],
		]);
		expect(result.nextPage.layoutContext?.sectionId).toBe('s2');
		expect(result.nextPage.layoutContext?.sectionPageIndex).toBe(1);
	});
});
