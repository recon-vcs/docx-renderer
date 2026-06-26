import { describe, expect, it } from 'vitest';
import { DocumentElement } from '../../src/ooxml/wordprocessingml/document/model/document';
import { Columns, SectionProperties, SectionType } from '../../src/ooxml/wordprocessingml/document/model/section';
import { splitDocumentIntoPhysicalPages } from '../../src/rendering/pagination/core/modern-page-splitter';
import { BreakType, DomType, OpenXmlElement } from '../../src/ooxml/wordprocessingml/model/element';

function section(sectionId: string, type?: SectionType, columnCount?: number): SectionProperties {
	return {
		sectionId,
		type,
		...(columnCount ? { columns: { count: columnCount } as Columns } : {}),
	} as SectionProperties;
}

function paragraph(children: OpenXmlElement[] = [], sectProps?: SectionProperties): OpenXmlElement {
	return {
		type: DomType.Paragraph,
		children,
		props: sectProps ? { sectionProperties: sectProps } : {},
	};
}

function run(children: OpenXmlElement[]): OpenXmlElement {
	return { type: DomType.Run, children };
}

function text(): OpenXmlElement {
	return { type: DomType.Text };
}

function explicitPageBreak(): OpenXmlElement {
	return { type: DomType.Break, break: BreakType.Page } as OpenXmlElement;
}

function lastRenderedPageBreak(): OpenXmlElement {
	return { type: DomType.LastRenderedPageBreak };
}

function document(children: OpenXmlElement[], rootSection: SectionProperties): DocumentElement {
	return {
		type: DomType.Document,
		children,
		pages: [],
		sectProps: rootSection,
	};
}

describe('splitDocumentIntoPhysicalPages', () => {
	it('uses explicit page breaks to create physical pages', () => {
		const beforeBreak = paragraph([run([text(), explicitPageBreak()])]);
		const afterBreak = paragraph([run([text()])]);

		const split = splitDocumentIntoPhysicalPages(document(
			[beforeBreak, afterBreak],
			section('root'),
		));

		expect(split.regions.map(region => region.breakBefore)).toEqual(['none', 'page']);
		expect(split.pages.map(page => page.pageNumber)).toEqual([1, 2]);
		expect(split.pages[0].regions[0].children).toEqual([beforeBreak]);
		expect(split.pages[1].regions[0].children).toEqual([afterBreak]);
	});

	it('keeps lastRenderedPageBreak as a hint without creating physical pages', () => {
		const withHint = paragraph([run([text(), lastRenderedPageBreak()])]);
		const afterHint = paragraph([run([text()])]);

		const split = splitDocumentIntoPhysicalPages(document(
			[withHint, afterHint],
			section('root'),
		));

		expect(split.regions).toHaveLength(1);
		expect(split.regions[0].hints).toEqual([
			{ kind: 'lastRenderedPageBreak', path: [0, 0, 1] },
		]);
		expect(split.pages).toHaveLength(1);
		expect(split.pages[0].regions).toHaveLength(1);
		expect(split.pages[0].regions[0].children).toEqual([withHint, afterHint]);
	});

	it('keeps continuous sections as multiple regions on one physical page', () => {
		const firstSection = section('first', SectionType.Continuous, 1);
		const rootSection = section('root', undefined, 2);
		const sectionBoundary = paragraph([run([text()])], firstSection);
		const afterBoundary = paragraph([run([text()])]);

		const split = splitDocumentIntoPhysicalPages(document(
			[sectionBoundary, afterBoundary],
			rootSection,
		));

		expect(split.regions.map(region => region.breakBefore)).toEqual(['none', 'none']);
		expect(split.pages).toHaveLength(1);
		expect(split.pages[0].regions).toHaveLength(2);
		expect(split.pages[0].regions.map(region => region.section.columns?.count)).toEqual([1, 2]);
		expect(split.pages[0].regions.map(region => region.children)).toEqual([
			[sectionBoundary],
			[afterBoundary],
		]);
	});
});
