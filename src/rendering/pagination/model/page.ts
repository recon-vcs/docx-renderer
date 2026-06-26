import { DomType, OpenXmlElement } from '@docx/ooxml/wordprocessingml/document/model/dom';
import { SectionProperties } from '@docx/ooxml/wordprocessingml/document/model/section';
import { uuid } from '@docx/shared/utils';
import type { LayoutRegion, PhysicalPage } from './layout-region';
import type { PageLayoutContext } from './page-numbering';

export interface TreeNode extends OpenXmlElement {
	prev?: TreeNode | null;
	next?: TreeNode | null;
}

export interface PageProps {
	sectProps?: SectionProperties;
	children: OpenXmlElement[];
	stack?: TreeNode[];
	isSplit?: boolean;
	isFirstPage?: boolean;
	isLastPage?: boolean;
	breakIndex?: Set<number>;
	regions?: LayoutRegion[];
	physicalPage?: PhysicalPage;
	layoutContext?: PageLayoutContext;
}

export class Page implements OpenXmlElement {
	type: DomType;
	pageId: string;
	sectProps?: SectionProperties;
	children: OpenXmlElement[];
	stack: TreeNode[];
	level?: number;
	isSplit: boolean;
	isFirstPage?: boolean;
	isLastPage?: boolean;
	breakIndex?: Set<number>;
	regions?: LayoutRegion[];
	physicalPage?: PhysicalPage;
	layoutContext?: PageLayoutContext;

	constructor({
		sectProps,
		children = [],
		stack = [],
		isSplit = false,
		isFirstPage = false,
		isLastPage = false,
		breakIndex = new Set(),
		regions,
		physicalPage,
		layoutContext,
	}: PageProps) {
		this.type = DomType.Page;
		this.level = 1;
		this.pageId = uuid();
		this.sectProps = sectProps;
		this.children = children;
		this.stack = stack;
		this.isSplit = isSplit;
		this.isFirstPage = isFirstPage;
		this.isLastPage = isLastPage;
		this.breakIndex = breakIndex;
		this.regions = regions;
		this.physicalPage = physicalPage;
		this.layoutContext = layoutContext;
	}
}
