import { describe, expect, it } from 'vitest';
import { parseSectionProperties, SectionType } from '../../src/document/section';

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

function parseSectPr(inner: string): Element {
	const xml = `<w:sectPr xmlns:w="${W_NS}" xmlns:r="${R_NS}">${inner}</w:sectPr>`;
	return new DOMParser().parseFromString(xml, 'text/xml').documentElement;
}

describe('parseSectionProperties', () => {
	it('parses page size', () => {
		const el = parseSectPr('<w:pgSz w:w="12240" w:h="15840"/>');
		const props = parseSectionProperties(el);
		expect(props.pageSize).toBeDefined();
		expect(props.pageSize.width).toBeTruthy();
		expect(props.pageSize.height).toBeTruthy();
	});

	it('parses page margins', () => {
		const el = parseSectPr(
			'<w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="720" w:footer="720" w:gutter="0"/>',
		);
		const props = parseSectionProperties(el);
		expect(props.pageMargins).toBeDefined();
		expect(props.pageMargins.top).toBeTruthy();
		expect(props.pageMargins.left).toBeTruthy();
		expect(props.pageMargins.right).toBeTruthy();
		expect(props.pageMargins.bottom).toBeTruthy();
	});

	it('defaults section type to undefined when not set', () => {
		const el = parseSectPr('');
		const props = parseSectionProperties(el);
		expect(props.type).toBeUndefined();
	});

	it('parses NextPage section type', () => {
		const el = parseSectPr('<w:type w:val="nextPage"/>');
		const props = parseSectionProperties(el);
		expect(props.type).toBe(SectionType.NextPage);
	});

	it('parses Continuous section type', () => {
		const el = parseSectPr('<w:type w:val="continuous"/>');
		const props = parseSectionProperties(el);
		expect(props.type).toBe(SectionType.Continuous);
	});

	it('parses NextColumn section type', () => {
		const el = parseSectPr('<w:type w:val="nextColumn"/>');
		const props = parseSectionProperties(el);
		expect(props.type).toBe(SectionType.NextColumn);
	});

	it('parses column count', () => {
		const el = parseSectPr('<w:cols w:num="2" w:space="720"/>');
		const props = parseSectionProperties(el);
		expect(props.columns).toBeDefined();
		expect(props.columns.count).toBe(2);
	});

	it('parses titlePage flag', () => {
		const el = parseSectPr('<w:titlePg/>');
		const props = parseSectionProperties(el);
		expect(props.titlePage).toBe(true);
	});

	it('parses header references', () => {
		const el = parseSectPr('<w:headerReference w:type="default" r:id="rId1"/>');
		const props = parseSectionProperties(el);
		expect(props.headerRefs).toHaveLength(1);
		expect(props.headerRefs[0].type).toBe('default');
		expect(props.headerRefs[0].id).toBe('rId1');
	});

	it('parses footer references', () => {
		const el = parseSectPr('<w:footerReference w:type="even" r:id="rId2"/>');
		const props = parseSectionProperties(el);
		expect(props.footerRefs).toHaveLength(1);
		expect(props.footerRefs[0].type).toBe('even');
		expect(props.footerRefs[0].id).toBe('rId2');
	});

	it('computes contentSize width from page size minus margins', () => {
		const el = parseSectPr(
			'<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="720" w:footer="720" w:gutter="0"/>',
		);
		const props = parseSectionProperties(el);
		expect(props.contentSize).toBeDefined();
		expect(props.contentSize.width).toBeTruthy();
	});

	it('parses multiple header and footer references', () => {
		const el = parseSectPr(
			'<w:headerReference w:type="default" r:id="rId1"/>' +
			'<w:headerReference w:type="first" r:id="rId2"/>' +
			'<w:footerReference w:type="default" r:id="rId3"/>',
		);
		const props = parseSectionProperties(el);
		expect(props.headerRefs).toHaveLength(2);
		expect(props.footerRefs).toHaveLength(1);
	});
});
