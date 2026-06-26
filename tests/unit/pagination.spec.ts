import { describe, expect, it } from 'vitest';
import { BreakType, DomType, OpenXmlElement } from '../../src/ooxml/wordprocessingml/model/element';
import { SectionProperties, SectionType } from '../../src/ooxml/wordprocessingml/document/model/section';
import { buildPaginationPlan } from '../../src/rendering/pagination/core/pagination';

function section(sectionId: string, type?: SectionType): SectionProperties {
	return { sectionId, type } as SectionProperties;
}

function paragraph(sectProps?: SectionProperties, children: OpenXmlElement[] = []): OpenXmlElement {
	return {
		type: DomType.Paragraph,
		props: sectProps ? { sectionProperties: sectProps } : {},
		children,
	};
}

function explicitPageBreak(): OpenXmlElement {
	return {
		type: DomType.Run,
		children: [{ type: DomType.Break, break: BreakType.Page } as OpenXmlElement],
	};
}

describe('buildPaginationPlan', () => {
	it('combines section stream, explicit breaks, and physical page construction', () => {
		const firstSection = section('s1', SectionType.Continuous);
		const rootSection = section('root');
		const afterBreak = paragraph(undefined, [explicitPageBreak()]);

		const plan = buildPaginationPlan([
			paragraph(firstSection),
			afterBreak,
			paragraph(),
		], rootSection);

		expect(plan.regions.map(region => region.breakBefore)).toEqual(['none', 'none', 'page']);
		expect(plan.pages.map(page => page.pageNumber)).toEqual([1, 2]);
		expect(plan.pages[0].regions).toHaveLength(2);
		expect(plan.pages[1].regions).toHaveLength(1);
	});

	it('uses section break parity when building physical pages', () => {
		const firstSection = section('s1', SectionType.OddPage);
		const rootSection = section('root');

		const plan = buildPaginationPlan([
			paragraph(firstSection),
			paragraph(),
		], rootSection);

		expect(plan.regions.map(region => region.breakBefore)).toEqual(['none', 'oddPage']);
		expect(plan.pages.map(page => page.pageNumber)).toEqual([1, 2, 3]);
		expect(plan.pages[1].regions).toEqual([]);
	});
});
