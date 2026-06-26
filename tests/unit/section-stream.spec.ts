import { describe, expect, it } from 'vitest';
import { buildSectionStream } from '../../src/rendering/pagination/model/section-stream';
import { DomType, OpenXmlElement } from '../../src/ooxml/wordprocessingml/model/element';
import { SectionProperties, SectionType } from '../../src/ooxml/wordprocessingml/document/model/section';

// ---------------------------------------------------------------------------
// Helper factories
// ---------------------------------------------------------------------------

function makeSectProps(type?: SectionType, id = 'sec-default'): SectionProperties {
	return { type, sectionId: id } as SectionProperties;
}

function makeParagraph(sectProps?: SectionProperties): OpenXmlElement {
	return {
		type: DomType.Paragraph,
		props: sectProps ? { sectionProperties: sectProps } : {},
	};
}

function makeTable(): OpenXmlElement {
	return { type: DomType.Table };
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

describe('buildSectionStream', () => {
	it('returns single region for document with no paragraph sectPr', () => {
		const rootSect = makeSectProps(undefined, 'root');
		const para1 = makeParagraph();
		const para2 = makeParagraph();

		const regions = buildSectionStream([para1, para2], rootSect);

		expect(regions).toHaveLength(1);
		expect(regions[0].section).toBe(rootSect);
		expect(regions[0].breakBefore).toBe('none');
		expect(regions[0].children).toEqual([para1, para2]);
	});

	it('single section boundary produces two regions when trailing content exists', () => {
		// In the implementation, the rootSectProps region is only emitted when
		// there are remaining children after all section boundaries OR there are
		// no regions at all. We include a trailing paragraph to get two regions.
		const sectA = makeSectProps(SectionType.NextPage, 'sectA');
		const rootSect = makeSectProps(undefined, 'root');
		const para1 = makeParagraph();
		const paraWithSect = makeParagraph(sectA);
		const para2 = makeParagraph();

		const regions = buildSectionStream([para1, paraWithSect, para2], rootSect);

		expect(regions).toHaveLength(2);

		// First region: terminated by the paragraph that holds sectA
		expect(regions[0].section).toBe(sectA);
		expect(regions[0].children).toEqual([para1, paraWithSect]);
		expect(regions[0].breakBefore).toBe('none');

		// Second region: rootSectProps, starts with a page break
		expect(regions[1].section).toBe(rootSect);
		expect(regions[1].children).toEqual([para2]);
		expect(regions[1].breakBefore).toBe('page');
	});

	it('Continuous section type gives breakBefore none on next region', () => {
		const sectA = makeSectProps(SectionType.Continuous, 'sectA');
		const rootSect = makeSectProps(undefined, 'root');

		const regions = buildSectionStream(
			[makeParagraph(sectA), makeParagraph()],
			rootSect,
		);

		expect(regions).toHaveLength(2);
		expect(regions[1].breakBefore).toBe('none');
	});

	it('NextColumn section type gives breakBefore column on next region', () => {
		const sectA = makeSectProps(SectionType.NextColumn, 'sectA');
		const rootSect = makeSectProps(undefined, 'root');

		const regions = buildSectionStream(
			[makeParagraph(sectA), makeParagraph()],
			rootSect,
		);

		expect(regions).toHaveLength(2);
		expect(regions[1].breakBefore).toBe('column');
	});

	it('EvenPage section type gives breakBefore evenPage', () => {
		const sectA = makeSectProps(SectionType.EvenPage, 'sectA');
		const rootSect = makeSectProps(undefined, 'root');

		const regions = buildSectionStream(
			[makeParagraph(sectA), makeParagraph()],
			rootSect,
		);

		expect(regions).toHaveLength(2);
		expect(regions[1].breakBefore).toBe('evenPage');
	});

	it('OddPage section type gives breakBefore oddPage', () => {
		const sectA = makeSectProps(SectionType.OddPage, 'sectA');
		const rootSect = makeSectProps(undefined, 'root');

		const regions = buildSectionStream(
			[makeParagraph(sectA), makeParagraph()],
			rootSect,
		);

		expect(regions).toHaveLength(2);
		expect(regions[1].breakBefore).toBe('oddPage');
	});

	it('undefined section type defaults to page break on next region', () => {
		// When SectionType is not set, sectionTypeToBreak returns 'page'.
		const sectA = makeSectProps(undefined, 'sectA');
		const rootSect = makeSectProps(undefined, 'root');

		const regions = buildSectionStream(
			[makeParagraph(sectA), makeParagraph()],
			rootSect,
		);

		expect(regions).toHaveLength(2);
		expect(regions[1].breakBefore).toBe('page');
	});

	it('NextPage section type defaults to page break on next region', () => {
		const sectA = makeSectProps(SectionType.NextPage, 'sectA');
		const rootSect = makeSectProps(undefined, 'root');

		const regions = buildSectionStream(
			[makeParagraph(sectA), makeParagraph()],
			rootSect,
		);

		expect(regions).toHaveLength(2);
		expect(regions[1].breakBefore).toBe('page');
	});

	it('multiple section boundaries produce correct region count', () => {
		// Three boundary paragraphs + one trailing element → 4 regions.
		const s1 = makeSectProps(SectionType.NextPage, 's1');
		const s2 = makeSectProps(SectionType.Continuous, 's2');
		const s3 = makeSectProps(SectionType.NextColumn, 's3');
		const rootSect = makeSectProps(undefined, 'root');

		const body: OpenXmlElement[] = [
			makeParagraph(),
			makeParagraph(s1),
			makeParagraph(),
			makeParagraph(s2),
			makeParagraph(),
			makeParagraph(s3),
			makeParagraph(), // trailing content triggers rootSectProps region
		];

		const regions = buildSectionStream(body, rootSect);

		expect(regions).toHaveLength(4);
		expect(regions[0].section).toBe(s1);
		expect(regions[1].section).toBe(s2);
		expect(regions[1].breakBefore).toBe('page'); // s1 is NextPage
		expect(regions[2].section).toBe(s3);
		expect(regions[2].breakBefore).toBe('none'); // s2 is Continuous
		expect(regions[3].section).toBe(rootSect);
		expect(regions[3].breakBefore).toBe('column'); // s3 is NextColumn
	});

	it('non-paragraph elements (tables) do not create section boundaries', () => {
		const sectA = makeSectProps(SectionType.NextPage, 'sectA');
		const rootSect = makeSectProps(undefined, 'root');
		const table1 = makeTable();
		const paraWithSect = makeParagraph(sectA);
		const table2 = makeTable();

		const regions = buildSectionStream([table1, paraWithSect, table2], rootSect);

		expect(regions).toHaveLength(2);
		// table1 and the boundary paragraph belong to region[0]
		expect(regions[0].children).toEqual([table1, paraWithSect]);
		// table2 is trailing content in the rootSectProps region
		expect(regions[1].section).toBe(rootSect);
		expect(regions[1].children).toEqual([table2]);
	});

	it('empty body produces single region with rootSectProps', () => {
		const rootSect = makeSectProps(undefined, 'root');

		const regions = buildSectionStream([], rootSect);

		expect(regions).toHaveLength(1);
		expect(regions[0].section).toBe(rootSect);
		expect(regions[0].children).toEqual([]);
		expect(regions[0].breakBefore).toBe('none');
	});

	it('first region always has breakBefore none', () => {
		const s1 = makeSectProps(SectionType.NextPage, 's1');
		const s2 = makeSectProps(SectionType.EvenPage, 's2');
		const rootSect = makeSectProps(undefined, 'root');

		const body: OpenXmlElement[] = [
			makeParagraph(s1),
			makeParagraph(s2),
			makeParagraph(),
		];

		const regions = buildSectionStream(body, rootSect);

		expect(regions[0].breakBefore).toBe('none');
	});

	it('paragraph with sectPr is included in the region it terminates', () => {
		const sectA = makeSectProps(SectionType.NextPage, 'sectA');
		const rootSect = makeSectProps(undefined, 'root');
		const para1 = makeParagraph();
		const paraWithSect = makeParagraph(sectA);
		const para2 = makeParagraph();

		const regions = buildSectionStream([para1, paraWithSect, para2], rootSect);

		// paraWithSect terminates region[0], so it must be in region[0].children
		expect(regions[0].children).toContain(paraWithSect);
		// and NOT in region[1].children
		expect(regions[1].children).not.toContain(paraWithSect);
	});
});
