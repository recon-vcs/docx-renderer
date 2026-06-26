import { describe, expect, it } from 'vitest';
import { BreakType, DomType, OpenXmlElement } from '../../src/ooxml/wordprocessingml/model/element';
import { SectionProperties } from '../../src/ooxml/wordprocessingml/document/model/section';
import { LayoutRegion } from '../../src/rendering/pagination/model/layout-region';
import { splitRegionsByExplicitBreaks } from '../../src/rendering/pagination/model/explicit-breaks';

function paragraph(children: OpenXmlElement[] = []): OpenXmlElement {
	return { type: DomType.Paragraph, children };
}

function run(children: OpenXmlElement[]): OpenXmlElement {
	return { type: DomType.Run, children };
}

function text(): OpenXmlElement {
	return { type: DomType.Text };
}

function breakElement(type: BreakType): OpenXmlElement {
	return { type: DomType.Break, break: type } as OpenXmlElement;
}

function lastRenderedBreak(): OpenXmlElement {
	return { type: DomType.LastRenderedPageBreak };
}

function region(children: OpenXmlElement[], breakBefore = 'none'): LayoutRegion {
	return {
		section: { sectionId: 's1' } as SectionProperties,
		children,
		breakBefore: breakBefore as LayoutRegion['breakBefore'],
	};
}

describe('splitRegionsByExplicitBreaks', () => {
	it('keeps a region unchanged when it has no explicit breaks', () => {
		const source = region([paragraph([run([text()])])]);

		expect(splitRegionsByExplicitBreaks([source])).toEqual([source]);
	});

	it('starts the following region on a new page after an explicit page break', () => {
		const before = paragraph([run([text(), breakElement(BreakType.Page)])]);
		const after = paragraph([run([text()])]);

		const result = splitRegionsByExplicitBreaks([region([before, after])]);

		expect(result).toHaveLength(2);
		expect(result[0].children).toEqual([before]);
		expect(result[0].breakBefore).toBe('none');
		expect(result[1].children).toEqual([after]);
		expect(result[1].breakBefore).toBe('page');
	});

	it('starts the following region on the next column after an explicit column break', () => {
		const before = paragraph([run([breakElement(BreakType.Column)])]);
		const after = paragraph([run([text()])]);

		const result = splitRegionsByExplicitBreaks([region([before, after])]);

		expect(result[1].breakBefore).toBe('column');
	});

	it('keeps lastRenderedPageBreak as a hint instead of splitting', () => {
		const withHint = paragraph([run([text(), lastRenderedBreak()])]);
		const after = paragraph([run([text()])]);

		const result = splitRegionsByExplicitBreaks([region([withHint, after])]);

		expect(result).toHaveLength(1);
		expect(result[0].children).toEqual([withHint, after]);
		expect(result[0].hints).toEqual([
			{ kind: 'lastRenderedPageBreak', path: [0, 0, 1] },
		]);
	});
});
