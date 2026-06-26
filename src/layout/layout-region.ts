import { OpenXmlElement } from '../model/element';
import { SectionProperties } from '../document/section';

export type RegionBreakBefore = 'none' | 'page' | 'column' | 'evenPage' | 'oddPage';

export interface LayoutRegion {
	section: SectionProperties;
	children: OpenXmlElement[];
	breakBefore: RegionBreakBefore;
}

export interface PhysicalPage {
	regions: LayoutRegion[];
	pageNumber: number;
	sectionPageIndexes: Map<string, number>;
}
