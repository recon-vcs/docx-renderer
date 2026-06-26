import { DocumentElement } from '@docx/ooxml/wordprocessingml/document/model/document';
import { SectionProperties } from '@docx/ooxml/wordprocessingml/document/model/section';
import { uuid } from '@docx/shared/utils';
import { splitRegionsByExplicitBreaks } from '@docx/rendering/pagination/model/explicit-breaks';
import { resolveHeaderFooterReferences } from '@docx/rendering/pagination/context/header-footer-context';
import { LayoutRegion, PhysicalPage } from '@docx/rendering/pagination/model/layout-region';
import { buildPhysicalPages } from '@docx/rendering/pagination/core/page-builder';
import { buildSectionStream } from '@docx/rendering/pagination/model/section-stream';

export interface ModernPageSplit {
	regions: LayoutRegion[];
	pages: PhysicalPage[];
}

export function splitDocumentIntoPhysicalPages(documentElement: DocumentElement): ModernPageSplit {
	const sectionRegions = normalizeRegionSections(
		buildSectionStream(documentElement.children ?? [], documentElement.sectProps),
	);
	const regions = splitRegionsByExplicitBreaks(sectionRegions);

	return {
		regions,
		pages: buildPhysicalPages(regions),
	};
}

function normalizeRegionSections(regions: LayoutRegion[]): LayoutRegion[] {
	let previousSection: SectionProperties | undefined;

	return regions.map((region) => {
		const section = normalizeSection(region.section, previousSection);
		previousSection = section;

		return {
			...region,
			section,
		};
	});
}

function normalizeSection(
	section: SectionProperties,
	previousSection: SectionProperties | undefined,
): SectionProperties {
	return {
		...section,
		sectionId: section.sectionId ?? uuid(),
		headerRefs: resolveHeaderFooterReferences(section.headerRefs, previousSection?.headerRefs),
		footerRefs: resolveHeaderFooterReferences(section.footerRefs, previousSection?.footerRefs),
	};
}
