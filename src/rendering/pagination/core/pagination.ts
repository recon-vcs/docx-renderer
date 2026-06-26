import { OpenXmlElement } from '@docx/ooxml/wordprocessingml/model/element';
import { SectionProperties } from '@docx/ooxml/wordprocessingml/document/model/section';
import { LayoutRegion, PhysicalPage } from '@docx/rendering/pagination/model/layout-region';
import { buildPhysicalPages } from '@docx/rendering/pagination/core/page-builder';
import { buildSectionStream } from '@docx/rendering/pagination/model/section-stream';
import { splitRegionsByExplicitBreaks } from '@docx/rendering/pagination/model/explicit-breaks';

export interface PaginationPlan {
	regions: LayoutRegion[];
	pages: PhysicalPage[];
}

export function buildPaginationPlan(
	bodyChildren: OpenXmlElement[],
	rootSectProps: SectionProperties,
): PaginationPlan {
	const sectionRegions = buildSectionStream(bodyChildren, rootSectProps);
	const regions = splitRegionsByExplicitBreaks(sectionRegions);
	const pages = buildPhysicalPages(regions);

	return { regions, pages };
}
