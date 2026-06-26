import { WordDocument } from './word-document';
import { DomType, IDomNumbering, OpenXmlElement } from './document/dom';
import { Options } from './docx-preview';
import { DocumentElement } from './document/document';
import * as _ from 'lodash-es';
import { computePointToPixelRatio, updateTabStop } from './javascript';
import { FontTablePart } from './font-table/font-table';
import { FooterHeaderReference, SectionProperties } from './document/section';
import { Page, PageProps } from './document/page';
import { WmlFieldSimple } from './document/fields';
import { IDomStyle } from './document/style';
import { WmlBaseNote, WmlEndnote, WmlEndnotes, WmlFootnote, WmlFootnotes } from './notes/elements';
import { ThemePart } from './theme/theme-part';
import { Part } from './common/part';
import type { Stage } from 'konva/lib/Stage';
import type { Layer } from 'konva/lib/Layer';
import { Overflow, ChildrenType, createElement, createElementNS, removeAllElements, appendChildren as appendChildrenSync, removeElements, createStyleElement, appendComment, Node_DOM } from './render/dom-utils';
import { appendAndMeasure, inferOverflow } from './measure/overflow-measurer';
import { splitOnOverflow, splitRegionOnOverflow } from './layout/page-split';
import { buildPageLayoutContexts, type PageLayoutContext } from './layout/page-numbering';
import { TableContext } from './render/table-renderer';
import { renderNotes as renderNotesFn } from './render/notes-renderer';
import { MathRendererCallbacks, mathJustificationToTextAlign } from './render/math-renderer';
import { DrawingRenderContext, createKonva as createKonvaFn } from './render/drawing-renderer';
import { renderHeaderFooter as renderHeaderFooterFn } from './render/header-footer-renderer';
import { InlineRendererCallbacks, renderText as renderTextFn } from './render/inline-renderer';
import { FieldsRendererCallbacks, resolveFieldRuns as resolveFieldRunsFn, resolveSimpleField as resolveSimpleFieldFn } from './render/fields-renderer';
import { renderDefaultStyle as renderDefaultStyleFn, renderWrapper as renderWrapperFn } from './render/styles/default-styles';
import { processStyleName as processStyleNameFn, processStyles as processStylesFn, renderFontTable as renderFontTableFn, renderStyles as renderStylesFn, renderTheme as renderThemeFn } from './render/styles/document-styles';
import { levelTextToContent as levelTextToContentFn, numberingClass as numberingClassFn, numberingCounter as numberingCounterFn, numFormatToCssValue as numFormatToCssValueFn, processNumberings as processNumberingsFn, renderNumbering as renderNumberingFn, styleToString as styleToStringFn } from './render/styles/numbering-styles';
import { PageRendererCallbacks, createPage as createPageFn, createPageContent as createPageContentFn, renderHeaderFooterRef as renderHeaderFooterRefFn } from './render/page-renderer';
import { splitDocumentIntoPhysicalPages } from './layout/modern-page-splitter';
import { processElement, processTable } from './render/element-processor';
import { renderStyleValues, renderClass, renderCommonProperties } from './render/style-applier';
import { dispatchElement, ElementDispatchContext } from './render/element-dispatcher';

interface RenderSplitContext {
	regionIndex: number;
}

// HTML渲染器
export class HtmlRendererSync {
	className = 'docx';
	rootSelector: string;
	document: WordDocument;
	options: Options;
	styleMap: Record<string, IDomStyle> = {};
	bodyContainer: HTMLElement;
	wrapper: HTMLElement;
	// 当前操作的Part
	currentPart: Part = null;
	// 系统的PPI
	pointToPixelRatio: number;

	// 当前操作的Page
	currentPage: Page;
	// Table rendering context (cell position + vertical-merge stacks for nested tables)
	tableCtx: TableContext = {
		tableVerticalMerges: [],
		currentVerticalMerge: null,
		tableCellPositions: [],
		currentCellPosition: null,
	};

	footnoteMap: Record<string, WmlFootnote> = {};
	endnoteMap: Record<string, WmlEndnote> = {};
	currentFootnoteIds: string[];
	currentEndnoteIds: string[] = [];
	// 已使用的Header、Footer部分的数组。
	usedHeaderFooterParts: any[] = [];

	defaultTabSize: string;
	// 当前制表位
	currentTabs: any[] = [];

	// Konva框架--stage元素
	konva_stage: Stage;
	// Konva框架--layer元素
	konva_layer: Layer;

	/**
	 * Object对象 => HTML标签
	 *
	 * @param document word文档Object对象
	 * @param bodyContainer HTML生成容器
	 * @param styleContainer CSS样式生成容器
	 * @param options 渲染配置选项
	 */

	async render(document: WordDocument, bodyContainer: HTMLElement, styleContainer: HTMLElement = null, options: Options) {
		// word文档对象
		this.document = document;
		// 渲染选项
		this.options = options;
		// class类前缀
		this.className = options.className;
		// 根元素
		this.rootSelector = options.inWrapper ? `.${this.className}-wrapper` : ':root';
		// 文档CSS样式
		this.styleMap = null;
		// 主体容器
		this.bodyContainer = bodyContainer;
		// 样式容器，可传参指定，默认为主体容器
		styleContainer = styleContainer || bodyContainer;
		// 计算Point/Pixel换算比例
		this.pointToPixelRatio = computePointToPixelRatio();
		// CSS样式生成容器，清空所有CSS样式
		removeAllElements(styleContainer);
		// HTML生成容器，清空所有HTML元素
		removeAllElements(bodyContainer);

		// 添加注释
		appendComment(styleContainer, 'docxjs library predefined styles');
		// 添加默认CSS样式
		styleContainer.appendChild(this.renderDefaultStyle());

		// 主题CSS样式
		if (document.themePart) {
			appendComment(styleContainer, 'docxjs document theme values');
			this.renderTheme(document.themePart, styleContainer);
		}
		// 文档默认CSS样式，包含表格、列表、段落、字体，样式存在继承顺序
		if (document.stylesPart != null) {
			this.styleMap = this.processStyles(document.stylesPart.styles);

			appendComment(styleContainer, 'docxjs document styles');
			styleContainer.appendChild(this.renderStyles(document.stylesPart.styles));
		}
		// 多级列表样式
		if (document.numberingPart) {
			this.processNumberings(document.numberingPart.domNumberings);

			appendComment(styleContainer, "docxjs document numbering styles");
			styleContainer.appendChild(this.renderNumbering(document.numberingPart.domNumberings, styleContainer));
			//styleContainer.appendChild(this.renderNumbering2(document.numberingPart, styleContainer));
		}
		// 字体列表CSS样式
		if (!options.ignoreFonts && document.fontTablePart) {
			this.renderFontTable(document.fontTablePart, styleContainer);
		}
		// 生成脚注部分的Map
		if (document.footnotesPart) {
			this.footnoteMap = _.keyBy(document.footnotesPart.rootElement.children, 'id');
		}
		// 生成尾注部分的Map
		if (document.endnotesPart) {
			this.endnoteMap = _.keyBy(document.endnotesPart.rootElement.children, 'id');
		}
		// 文档设置
		if (document.settingsPart) {
			this.defaultTabSize = document.settingsPart.settings?.defaultTabStop;
		}
		this.assignSourcePaths(document.documentPart.body.children);
		// 根据option生成wrapper
		if (this.options.inWrapper) {
			this.wrapper = this.renderWrapper();
			bodyContainer.appendChild(this.wrapper);
		} else {
			this.wrapper = bodyContainer;
		}
		// 生成Canvas画布元素--Konva框架
		this.renderKonva();
		// 主文档--内容
		await this.renderPages(document.documentPart.body);
		// 渲染完成所有Page, 隐藏Stage
		this.konva_stage.visible(false);
		// 刷新制表符
		this.refreshTabStops();
	}

	assignSourcePaths(children: OpenXmlElement[]) {
		children.forEach((child, index) => {
			const path = `body/${index}`;
			child.sourcePath = path;
			this.assignNestedSourcePaths(child, path);
		});
	}

	assignNestedSourcePaths(element: OpenXmlElement, path: string) {
		if (element.type === DomType.Table) {
			element.children?.forEach((row, rowIndex) => {
				row.children?.forEach((cell, cellIndex) => {
					const cellPath = `${path}/cell/${rowIndex}/${cellIndex}`;
					cell.sourcePath = cellPath;
					cell.children?.forEach((child) => {
						child.sourcePath = cellPath;
						this.assignNestedSourcePaths(child, cellPath);
					});
				});
			});
			return;
		}

		element.children?.forEach((child) => {
			child.sourcePath = path;
			this.assignNestedSourcePaths(child, path);
		});
	}

	// Render built-in default styles.
	renderDefaultStyle() {
		return renderDefaultStyleFn(this.className);
	}

	// Render document theme CSS variables.
	renderTheme(themePart: ThemePart, styleContainer: HTMLElement) {
		renderThemeFn(themePart, styleContainer, this.documentStylesCallbacks());
	}

	// Build a CSS class name for a Word style.
	processStyleName(className: string): string {
		return processStyleNameFn(className, this.className);
	}

	// Merge inherited style rules. Base styles are ordered before dependent styles.
	processStyles(styles: IDomStyle[]) {
		return processStylesFn(styles, this.documentStylesCallbacks());
	}

	// Render document style rules.
	renderStyles(styles: IDomStyle[]): HTMLElement {
		return renderStylesFn(styles, this.documentStylesCallbacks());
	}

	processNumberings(numberings: IDomNumbering[]) {
		processNumberingsFn(numberings, this.numberingStylesCallbacks());
	}

	renderNumbering(numberings: IDomNumbering[], styleContainer: HTMLElement) {
		return renderNumberingFn(numberings, styleContainer, this.numberingStylesCallbacks());
	}

	numberingClass(id: string, lvl: number) {
		return numberingClassFn(this.className, id, lvl);
	}

	styleToString(selectors: string, declarations: Record<string, string>, cssText: string = null) {
		return styleToStringFn(selectors, declarations, cssText);
	}

	numberingCounter(id: string, lvl: number) {
		return numberingCounterFn(this.className, id, lvl);
	}

	levelTextToContent(text: string, suff: string, id: string, numformat: string) {
		return levelTextToContentFn(text, suff, id, numformat, (counterId, level) => this.numberingCounter(counterId, level));
	}

	numFormatToCssValue(format: string) {
		return numFormatToCssValueFn(format);
	}

	// Render embedded font-face rules.
	renderFontTable(fontsPart: FontTablePart, styleContainer: HTMLElement) {
		renderFontTableFn(fontsPart, styleContainer, this.documentStylesCallbacks());
	}

	// Render the optional wrapper container.
	renderWrapper() {
		return renderWrapperFn(this.className);
	}

	private documentStylesCallbacks() {
		return {
			className: this.className,
			options: this.options,
			styleToString: (selectors, declarations, cssText = null) => this.styleToString(selectors, declarations, cssText),
			processStyleName: (className) => this.processStyleName(className),
			createStyleElement: (cssText) => createStyleElement(cssText),
			appendComment: (styleContainer, text) => appendComment(styleContainer, text),
			loadFont: (id, key) => this.document.loadFont(id, key),
			refreshTabStops: () => this.refreshTabStops(),
		};
	}

	private numberingStylesCallbacks() {
		return {
			className: this.className,
			rootSelector: this.rootSelector,
			findStyle: (styleName) => this.findStyle(styleName),
			styleToString: (selectors, declarations, cssText = null) => this.styleToString(selectors, declarations, cssText),
			createStyleElement: (cssText) => createStyleElement(cssText),
			loadNumberingImage: (src) => this.document.loadNumberingImage(src),
			numberingClass: (id, level) => this.numberingClass(id, level),
			numberingCounter: (id, level) => this.numberingCounter(id, level),
			levelTextToContent: (text, suff, id, numformat) => this.levelTextToContent(text, suff, id, numformat),
			numFormatToCssValue: (format) => this.numFormatToCssValue(format),
		};
	}

	// Build physical pages from section regions and explicit break data.
	splitPageBySymbol(documentElement: DocumentElement): Page[] {
		const split = splitDocumentIntoPhysicalPages(documentElement);
		const physicalPagesWithRegions = split.pages.filter(physicalPage => physicalPage.regions.length > 0);
		const contexts = buildPageLayoutContexts(physicalPagesWithRegions);
		const contextByPage = new Map(physicalPagesWithRegions.map((physicalPage, index) => [
			physicalPage,
			contexts[index],
		]));

		return split.pages.map(physicalPage => {
			const activeRegion = physicalPage.regions[physicalPage.regions.length - 1];
			const children = physicalPage.regions.flatMap(region => region.children);
			const layoutContext = contextByPage.get(physicalPage);

			return new Page({
				isSplit: false,
				sectProps: layoutContext?.activeSection ?? activeRegion?.section ?? documentElement.sectProps,
				children,
				regions: physicalPage.regions,
				physicalPage,
				layoutContext,
			} as PageProps);
		});
	}

	// 生成所有的页面Page
	async renderPages(document: DocumentElement) {
		// 根据options.breakPages，选择是否分页
		let pages: Page[];
		if (this.options.breakPages) {
			// 根据分页符，初步拆分页面
			pages = this.splitPageBySymbol(document);
		} else {
			// 不分页则，只有一个page
			pages = [new Page({ isSplit: true, sectProps: document.sectProps, children: document.children, } as PageProps)];
		}
		// 初步分页结果,缓存至body中
		document.pages = pages;
		// 前一个节属性，判断分节符的第一个page
		let prevProps = null;
		// 浅拷贝初步分页结果，后续拆分操作将不断扩充数组，导致下面循环异常
		let origin_pages = [...pages];
		// 遍历生成每一个page
		for (let i = 0; i < origin_pages.length; i++) {
			this.currentFootnoteIds = [];
			const page: Page = origin_pages[i];
			const { sectProps } = page;
			// sectionProps属性不存在，则使用文档级别props;
			page.sectProps = sectProps ?? document.sectProps;
			// 是否本小节的第一个page
			page.isFirstPage = page.layoutContext?.isFirstSectionPage ?? prevProps != page.sectProps;
			// TODO 是否最后一个page,此时分页未完成，计算并不准确，影响到尾注的渲染
			page.isLastPage = i === origin_pages.length - 1;
			// 溢出检测默认不开启
			page.checkingOverflow = false;
			// 将上述数据存储在currentPage中
			this.currentPage = page;
			// 存储前一个节属性
			prevProps = page.sectProps;
			// 渲染单个page
			await this.renderPage();
		}
	}

	// Render a single page. If content overflows, recursively splits into the next page.
	async renderPage() {
		const { pageId, sectProps, children, isFirstPage, isLastPage, regions, layoutContext } = this.currentPage;
		processElement(this.currentPage);
		const pageElement = this.createPage(this.className, sectProps);

		renderStyleValues(this.document.documentPart.body.cssStyle, pageElement);

		const pages = this.document.documentPart.body.pages;
		const pageIndex = pages.findIndex((page) => page.pageId === pageId);

		let oHeader: HTMLElement = null;
		let oFooter: HTMLElement = null;
		if (this.options.renderHeaders) {
			oHeader = await this.renderHeaderFooterRef(sectProps.headerRefs, sectProps, pageIndex, isFirstPage, layoutContext, pageElement);
		}
		if (this.options.renderFooters) {
			oFooter = await this.renderHeaderFooterRef(sectProps.footerRefs, sectProps, pageIndex, isFirstPage, layoutContext, pageElement);
		}

		const getOffsetHeight = (el: HTMLElement) => (el?.offsetHeight ?? 0) * this.pointToPixelRatio;
		const { pageSize, pageMargins } = sectProps;
		const headerHeight = getOffsetHeight(oHeader);
		const footerHeight = getOffsetHeight(oFooter);
		const actualTop = _.max([parseFloat(pageMargins.top), headerHeight]);
		const actualBottom = _.max([parseFloat(pageMargins.bottom), footerHeight]);
		pageElement.style.paddingTop = `${actualTop}pt`;
		pageElement.style.paddingBottom = `${actualBottom}pt`;
		const contentHeight = parseFloat(pageSize.height) - actualTop - actualBottom;

		if (regions && regions.length > 1) {
			if (this.options.breakPages && !this.options.ignoreHeight) {
				pageElement.style.height = sectProps.pageSize.height;
			}

			let isOverflow = Overflow.FALSE;
			for (let regionIndex = 0; regionIndex < regions.length; regionIndex++) {
				const region = regions[regionIndex];
				const regionArticle = this.createPageContent(region.section);
				regionArticle.dataset.sectionId = region.section?.sectionId;
				regionArticle.dataset.breakBefore = region.breakBefore;
				pageElement.appendChild(regionArticle);
				this.currentPage.contentElement = pageElement;
				this.currentPage.checkingOverflow = this.options.breakPages;
				isOverflow = await this.renderElements(region.children, regionArticle, { regionIndex });
				if (isOverflow !== Overflow.FALSE && isOverflow !== Overflow.UNKNOWN && isOverflow !== Overflow.IGNORE) {
					break;
				}
			}
			if (isOverflow === Overflow.FALSE || isOverflow === Overflow.UNKNOWN || isOverflow === Overflow.IGNORE) {
				this.currentPage.isSplit = true;
				pages[pageIndex] = this.currentPage;
			}
			this.currentPage.checkingOverflow = false;
		} else {
			// Single-region page: standard overflow detection path.
			const contentElement = this.createPageContent(sectProps);
			if (this.options.breakPages) {
				contentElement.style.height = `${contentHeight}pt`;
			} else {
				contentElement.style.minHeight = `${contentHeight}pt`;
			}
			this.currentPage.contentElement = contentElement;
			pageElement.appendChild(contentElement);
			this.currentPage.checkingOverflow = true;
			const is_overflow = await this.renderElements(children, contentElement);
			if (is_overflow === Overflow.FALSE) {
				this.currentPage.isSplit = true;
				pages[pageIndex] = this.currentPage;
			}
			this.currentPage.checkingOverflow = false;
		}

		if (this.options.renderFootnotes) {
			await this.renderNotes(DomType.Footnotes, this.currentFootnoteIds, this.footnoteMap, pageElement);
		}
		if (this.options.renderEndnotes && isLastPage) {
			await this.renderNotes(DomType.Endnotes, this.currentEndnoteIds, this.endnoteMap, pageElement);
		}
	}

	// 创建Page
	createPage(className: string, props: SectionProperties) {
		return createPageFn(className, props, this.wrapper, {
			ignoreWidth: this.options.ignoreWidth,
			ignoreHeight: this.options.ignoreHeight,
		});
	}

	// TODO 分栏：一个页面可能存在多个章节section，每个section拥有不同的分栏
	// 多列分栏布局
	createPageContent(props: SectionProperties): HTMLElement {
		return createPageContentFn(props);
	}

	// TODO 分页不准确，页脚页码混乱，
	// TODO 支持奇数页偶数页不同页眉页脚
	// 渲染页眉/页脚的Ref
	async renderHeaderFooterRef(refs: FooterHeaderReference[], props: SectionProperties, pageIndex: number, isFirstPage: boolean, layoutContext: PageLayoutContext | undefined, parent: HTMLElement) {
		return renderHeaderFooterRefFn(refs, props, pageIndex, isFirstPage, layoutContext, parent, this.pageRendererCallbacks());
	}

	private pageRendererCallbacks(): PageRendererCallbacks {
		return {
			document: this.document,
			ignoreWidth: this.options.ignoreWidth,
			ignoreHeight: this.options.ignoreHeight,
			evenAndOddHeaders: this.document.settingsPart.settings.evenAndOddHeaders,
			usedHeaderFooterParts: this.usedHeaderFooterParts,
			setCurrentPart: part => { this.currentPart = part; },
			processElement: elem => processElement(elem),
			renderHeaderFooter: (elem, tagName, parent) => this.renderHeaderFooter(elem, tagName, parent),
		};
	}

	// TODO 字体太大，尾注位置不对
	// 渲染脚注/尾注
	async renderNotes(type: DomType = DomType.Footnotes, noteIds: string[], notesMap: Record<string, WmlBaseNote>, parent: HTMLElement) {
		return renderNotesFn(type, noteIds, notesMap, parent, {
			processElement: (e) => processElement(e),
			renderChildren: (e, p) => this.renderChildren(e, p),
		});
	}

	// 根据XML对象渲染多元素
	async renderElements(children: OpenXmlElement[], parent: HTMLElement | Element | Text, splitContext?: RenderSplitContext): Promise<Overflow> {
		const overflows: Overflow[] = [];
		const pages = this.document.documentPart.body.pages;
		const { pageId } = this.currentPage;
		const pageIndex = pages.findIndex(p => p.pageId === pageId);

		for (let i = 0; i < children.length; i++) {
			const elem = children[i];
			elem.index = i;
			if (!elem.breakIndex) elem.breakIndex = new Set();

			const rendered = await this.renderElement(elem, parent);
			let overflow: Overflow = rendered?.dataset?.overflow as Overflow ?? Overflow.UNKNOWN;
			let action = 'continue';

			switch (overflow) {
				case Overflow.SELF:
					elem.breakIndex.add(0);
					elem.parent.breakIndex.add(i);
					removeElements(rendered, parent);
					action = 'break';
					break;

				case Overflow.TRUE:
				case Overflow.FULL:
					elem.parent.breakIndex.add(i);
					if (elem.type !== DomType.Cell) removeElements(rendered, parent);
					action = 'break';
					break;

				case Overflow.PART:
					elem.parent.breakIndex.add(i);
					action = 'break';
					break;

				default:
					action = 'continue';
					if (overflow !== Overflow.FALSE && overflow !== Overflow.UNKNOWN && overflow !== Overflow.IGNORE && this.options.debug) {
						console.error('unhandled overflow', overflow, elem);
					}
			}

			if (elem.type === DomType.Cell) action = 'continue';
			overflows.push(overflow);

			if (action === 'break') {
				if (elem.level === 2) {
					if (splitContext) {
						splitRegionOnOverflow(this.currentPage, pages, pageIndex, splitContext.regionIndex, i);
					} else {
						splitOnOverflow(this.currentPage, pages, pageIndex, i);
					}
					processElement(this.currentPage);
					this.currentPage = pages[pageIndex + 1];
					await this.renderPage();
				}
				break;
			}
		}

		return inferOverflow(overflows);
	}

	// 根据XML对象渲染单个元素
	async renderElement(elem: OpenXmlElement, parent?: HTMLElement | Element | Text): Promise<Node_DOM> {
		const oNode = await dispatchElement(elem, parent, this.dispatchContext());
		if (oNode && oNode.nodeType === 1) {
			oNode.dataset.tag = elem.type;
			if (elem.sourcePath && this.isSourceAnchor(elem)) {
				oNode.dataset.vellumPath = elem.sourcePath;
			}
		}
		return oNode as Node_DOM;
	}

	private dispatchContext(): ElementDispatchContext {
		return {
			appendChildren: (p, c) => this.appendChildren(p, c),
			renderChildren: (e, p) => this.renderChildren(e, p),
			renderElements: (ch, p) => this.renderElements(ch, p),
			renderContainer: (e, t, pr) => this.renderContainer(e, t, pr),
			renderContainerNS: (e, n, t, pr) => this.renderContainerNS(e, n, t, pr),
			renderHeaderFooter: (e, t, p) => this.renderHeaderFooter(e, t, p),
			inlineCallbacks: () => this.inlineCallbacks(),
			mathCallbacks: () => this.mathCallbacks(),
			fieldsCallbacks: () => this.fieldsCallbacks(),
			drawingRenderContext: () => this.drawingRenderContext(),
			tableCtx: this.tableCtx,
			className: this.className,
			currentTabs: this.currentTabs,
			currentFootnoteIds: this.currentFootnoteIds,
			currentEndnoteIds: this.currentEndnoteIds,
			renderClass: (e, o) => renderClass(e, o, n => this.processStyleName(n)),
			renderStyleValues: (s, o) => renderStyleValues(s, o),
		};
	}

	isSourceAnchor(elem: OpenXmlElement): boolean {
		return elem.type === DomType.Paragraph || elem.type === DomType.Table || elem.type === DomType.Cell || elem.type === DomType.SectionBreak;
	}

	// 根据XML对象渲染子元素，并插入父级元素
	async renderChildren(elem: OpenXmlElement, parent: HTMLElement | Element | Text): Promise<Overflow> {
		return await this.renderElements(elem.children, parent);
	}

	// 插入子元素，针对后代元素进行溢出检测
	async appendChildren(parent: HTMLElement | Text, children: ChildrenType): Promise<Overflow> {
		return appendAndMeasure(parent, children, this.currentPage);
	}

	async renderContainer(elem: OpenXmlElement, tagName: keyof HTMLElementTagNameMap, props?: Record<string, any>) {
		const oContainer = createElement(tagName, props);

		oContainer.dataset.overflow = await this.renderChildren(elem, oContainer);
		return oContainer;
	}

	async renderContainerNS(elem: OpenXmlElement, ns: string, tagName: string, props?: Record<string, any>) {
		const parent = createElementNS(ns, tagName as any, props);
		await this.renderChildren(elem, parent);
		return parent;
	}

	mathJustificationToTextAlign(justification?: string) {
		return mathJustificationToTextAlign(justification);
	}

	private mathCallbacks(): MathRendererCallbacks {
		return {
			renderElement: (e, p) => this.renderElement(e, p as HTMLElement) as any,
			renderElements: (es, p) => this.renderElements(es, p as HTMLElement),
			renderChildren: (e, p) => this.renderChildren(e, p as HTMLElement),
			renderContainerNS: (e, n, t, pr) => this.renderContainerNS(e, n, t, pr),
			renderClass: (e, o) => renderClass(e, o, n => this.processStyleName(n)),
			renderStyleValues: (s, o) => renderStyleValues(s, o),
			className: this.className,
		};
	}

	resolveFieldRuns(runs: OpenXmlElement[]): OpenXmlElement[] {
		return resolveFieldRunsFn(runs, this.fieldsCallbacks());
	}

	resolveSimpleField(elem: WmlFieldSimple): OpenXmlElement[] {
		return resolveSimpleFieldFn(elem, this.fieldsCallbacks());
	}

	private fieldsCallbacks(): FieldsRendererCallbacks {
		return {
			processElement: (e) => processElement(e),
			renderChildren: (e, p) => this.renderChildren(e, p),
			appendChildren: (p, c) => this.appendChildren(p, c),
			appendChildrenWithoutOverflow: (p, c) => appendChildrenSync(p, c),
			renderText: (e, p) => renderTextFn(e, p, this.inlineCallbacks()),
			findComment: (id) => this.document.commentsPart?.commentMap[id],
			currentPageNumber: () => {
				const pages = this.document.documentPart.body.pages ?? [];
				return pages.findIndex(p => p.pageId === this.currentPage.pageId) + 1;
			},
			pageCount: () => this.document.documentPart.body.pages?.length ?? 0,
			renderChanges: () => this.options.renderChanges,
		};
	}

	private inlineCallbacks(): InlineRendererCallbacks {
		return {
			appendChildren: (p, c) => this.appendChildren(p, c),
			renderChildren: (e, p) => this.renderChildren(e, p),
			renderClass: (e, o) => renderClass(e, o, n => this.processStyleName(n)),
			renderCommonProperties: (s, p) => renderCommonProperties(s, p),
			renderStyleValues: (s, o) => renderStyleValues(s, o),
			resolveFieldRuns: (runs) => this.resolveFieldRuns(runs),
			findStyle: (styleName) => this.findStyle(styleName),
			numberingClass: (id, level) => this.numberingClass(id, level),
			currentPageIsSplit: () => this.currentPage.isSplit,
			currentSectionProperties: () => this.currentPage.sectProps,
			findExternalRelation: (id) => this.document.documentPart.rels.find(
				it => it.id == id && it.targetMode === 'External'
			),
		};
	}

	renderKonva() {
		const { stage, layer } = createKonvaFn(this.bodyContainer);
		this.konva_stage = stage;
		this.konva_layer = layer;
	}

	private drawingRenderContext(): DrawingRenderContext {
		return {
			document: this.document,
			currentPart: this.currentPart,
			options: this.options,
			konvaStage: this.konva_stage,
			konvaLayer: this.konva_layer,
			appendChildren: (p, c) => this.appendChildren(p, c),
			renderChildren: (e, p) => this.renderChildren(e, p),
			renderElement: (e, p) => this.renderElement(e, p),
			renderStyleValues: (s, o) => renderStyleValues(s, o),
		};
	}

	// 渲染页眉页脚
	async renderHeaderFooter(elem: OpenXmlElement, tagName: keyof HTMLElementTagNameMap, parent: HTMLElement) {
		return renderHeaderFooterFn(elem, tagName, parent, {
			renderChildren: (e, p) => this.renderChildren(e, p),
			renderStyleValues: (s, o) => renderStyleValues(s, o),
		});
	}

	// 查找内置默认style样式
	findStyle(styleName: string) {
		return styleName && this.styleMap?.[styleName];
	}

	tabStopClass() {
		return `${this.className}-tab-stop`;
	}

	// 刷新tab制表符
	refreshTabStops() {
		for (const tab of this.currentTabs) {
			updateTabStop(tab.span, tab.stops, this.defaultTabSize, this.pointToPixelRatio);
		}
	}
}
