import { OpenXmlElement } from '@docx/ooxml/wordprocessingml/model/element';
import { SectionProperties } from '@docx/ooxml/wordprocessingml/document/model/section';

export type RegionBreakBefore = 'none' | 'page' | 'column' | 'evenPage' | 'oddPage';

export type LayoutRegionHintKind = 'lastRenderedPageBreak';

export interface LayoutRegionHint {
	kind: LayoutRegionHintKind;
	path: number[];
}

export interface LayoutRegion {
	section: SectionProperties;
	children: OpenXmlElement[];
	breakBefore: RegionBreakBefore;
	hints?: LayoutRegionHint[];
}

export interface PhysicalPage {
	regions: LayoutRegion[];
	pageNumber: number;
	sectionPageIndexes: Map<string, number>;
}
