import { WordDocument } from '@docx/word-document';
import { DomType, IDomNumbering, OpenXmlElement } from '@docx/ooxml/wordprocessingml/document/model/dom';
import type { Options } from '@docx/options';
import { DocumentElement } from '@docx/ooxml/wordprocessingml/document/model/document';
import * as _ from 'lodash-es';
import { computePointToPixelRatio, updateTabStop } from '@docx/javascript';
import type { TabStop } from '@docx/ooxml/wordprocessingml/document/model/paragraph';
import { FontTablePart } from '@docx/ooxml/wordprocessingml/parts/font-table/font-table';
import { FooterHeaderReference, SectionProperties } from '@docx/ooxml/wordprocessingml/document/model/section';
import { Page, PageProps } from '@docx/rendering/pagination/model/page';
import { WmlFieldSimple } from '@docx/ooxml/wordprocessingml/document/model/fields';
import { IDomStyle } from '@docx/ooxml/wordprocessingml/document/model/style';
import { WmlBaseNote, WmlEndnote, WmlEndnotes, WmlFootnote, WmlFootnotes } from '@docx/ooxml/wordprocessingml/parts/notes/elements';
import { ThemePart } from '@docx/ooxml/drawingml/theme/theme-part';
import { Part } from '@docx/opc/parts/part';
import type { Stage } from 'konva/lib/Stage';
import type { Layer } from 'konva/lib/Layer';
import { ChildrenType, createElement, createElementNS, removeAllElements, appendChildren as appendChildrenSync, removeElements, createStyleElement, appendComment, Node_DOM } from '@docx/rendering/dom/core/dom-utils';
import { measurePageOverflow, inferOverflow } from '@docx/rendering/measurement/overflow-measurer';
import { Overflow } from '@docx/rendering/measurement/overflow';
import { splitOnOverflow, splitRegionOnOverflow } from '@docx/rendering/pagination/model/page-split';
import { buildPageLayoutContexts, type PageLayoutContext } from '@docx/rendering/pagination/model/page-numbering';
import { TableContext } from '@docx/rendering/dom/elements/table-renderer';
import { renderNotes as renderNotesFn } from '@docx/rendering/dom/elements/notes-renderer';
import { MathRendererCallbacks, mathJustificationToTextAlign } from '@docx/rendering/dom/elements/math-renderer';
import { DrawingRenderContext, createKonva as createKonvaFn } from '@docx/rendering/dom/elements/drawing-renderer';
import { renderHeaderFooter as renderHeaderFooterFn } from '@docx/rendering/dom/elements/header-footer-renderer';
import { InlineRendererCallbacks, renderText as renderTextFn } from '@docx/rendering/dom/elements/inline-renderer';
import { FieldsRendererCallbacks, resolveFieldRuns as resolveFieldRunsFn, resolveSimpleField as resolveSimpleFieldFn } from '@docx/rendering/dom/elements/fields-renderer';
import { renderDefaultStyle as renderDefaultStyleFn, renderWrapper as renderWrapperFn } from '@docx/rendering/dom/styles/default-styles';
import { processStyleName as processStyleNameFn, processStyles as processStylesFn, renderFontTable as renderFontTableFn, renderStyles as renderStylesFn, renderTheme as renderThemeFn } from '@docx/rendering/dom/styles/document-styles';
import { levelTextToContent as levelTextToContentFn, numberingClass as numberingClassFn, numberingCounter as numberingCounterFn, numFormatToCssValue as numFormatToCssValueFn, processNumberings as processNumberingsFn, renderNumbering as renderNumberingFn, styleToString as styleToStringFn } from '@docx/rendering/dom/styles/numbering-styles';
import { PageRendererCallbacks, createPage as createPageFn, createPageContent as createPageContentFn, renderHeaderFooterRef as renderHeaderFooterRefFn } from '@docx/rendering/dom/elements/page-renderer';
import { splitDocumentIntoPhysicalPages } from '@docx/rendering/pagination/core/modern-page-splitter';
import { processElement, processTable } from '@docx/rendering/dom/core/element-processor';
import { renderStyleValues, renderClass, renderCommonProperties } from '@docx/rendering/dom/styles/style-applier';
import { dispatchElement, ElementDispatchContext } from '@docx/rendering/dom/core/element-dispatcher';

interface RenderSplitContext {
	regionIndex: number;
}

interface PendingTabStop {
	stops?: TabStop[];
	span: HTMLElement;
}

/** All mutable state that is created fresh for each render() call. */
interface RenderSession {
	/** Pages produced and mutated during the current render. */
	pages: Page[];
	/** The page currently being rendered. Advances as overflow splits pages. */
	currentPage: Page;
	/** The OPC Part that owns currently-rendered content (changes per header/footer/drawing). */
	currentPart: Part;
	/** Footnote ref ids collected for the current page. Reset per page. */
	currentFootnoteIds: string[];
	/** Endnote ref ids collected across all pages. */
	currentEndnoteIds: string[];
	/** Table cell position and vertical-merge stacks for nested tables. */
	tableCtx: TableContext;
	/** Header/footer Parts that have already been rendered (to avoid double-registering). */
	usedHeaderFooterParts: string[];
	/** Tab-stop spans queued for post-render measurement. */
	currentTabs: PendingTabStop[];
	/** Konva stage used for image clipping/transforms. Destroyed after render. */
	konvaStage: Stage;
	/** Konva layer paired with konvaStage. */
	konvaLayer: Layer;
	/** DOM element currently measured for page overflow. */
	overflowContentElement?: HTMLElement;
	/** Whether append operations should measure overflow. */
	checkingOverflow: boolean;
}

export class HtmlRendererSync {
	private className = 'docx';
	private rootSelector: string;
	private document: WordDocument;
	private options: Options;
	private styleMap: Record<string, IDomStyle> = {};
	private bodyContainer: HTMLElement;
	private wrapper: HTMLElement;
	private pointToPixelRatio: number;
	private footnoteMap: Record<string, WmlFootnote> = {};
	private endnoteMap: Record<string, WmlEndnote> = {};
	private defaultTabSize: string;

	/** Mutable per-render state. Null between renders. */
	private session: RenderSession | null = null;

	async render(document: WordDocument, bodyContainer: HTMLElement, styleContainer: HTMLElement = null, options: Options) {
		this.document = document;
		this.options = options;
		this.className = options.className;
		this.rootSelector = options.inWrapper ? `.${this.className}-wrapper` : ':root';
		this.styleMap = null;
		this.bodyContainer = bodyContainer;
		styleContainer = styleContainer || bodyContainer;
		this.pointToPixelRatio = computePointToPixelRatio();

		removeAllElements(styleContainer);
		removeAllElements(bodyContainer);

		appendComment(styleContainer, 'docxjs library predefined styles');
		styleContainer.appendChild(this.renderDefaultStyle());

		if (document.themePart) {
			appendComment(styleContainer, 'docxjs document theme values');
			this.renderTheme(document.themePart, styleContainer);
		}
		if (document.stylesPart != null) {
			this.styleMap = this.processStyles(document.stylesPart.styles);

			appendComment(styleContainer, 'docxjs document styles');
			styleContainer.appendChild(this.renderStyles(document.stylesPart.styles));
		}
		if (document.numberingPart) {
			this.processNumberings(document.numberingPart.domNumberings);

			appendComment(styleContainer, "docxjs document numbering styles");
			styleContainer.appendChild(this.renderNumbering(document.numberingPart.domNumberings, styleContainer));
		}
		if (!options.ignoreFonts && document.fontTablePart) {
			this.renderFontTable(document.fontTablePart, styleContainer);
		}
		if (document.footnotesPart) {
			this.footnoteMap = _.keyBy(document.footnotesPart.rootElement.children, 'id');
		}
		if (document.endnotesPart) {
			this.endnoteMap = _.keyBy(document.endnotesPart.rootElement.children, 'id');
		}
		if (document.settingsPart) {
			this.defaultTabSize = document.settingsPart.settings?.defaultTabStop;
		}
		this.assignSourcePaths(document.documentPart.body.children);

		if (this.options.inWrapper) {
			this.wrapper = this.renderWrapper();
			bodyContainer.appendChild(this.wrapper);
		} else {
			this.wrapper = bodyContainer;
		}

		const { stage, layer } = createKonvaFn(this.bodyContainer);
			this.session = {
				pages: [],
				currentPage: null,
			currentPart: null,
			currentFootnoteIds: [],
			currentEndnoteIds: [],
			tableCtx: {
				tableVerticalMerges: [],
				currentVerticalMerge: null,
				tableCellPositions: [],
				currentCellPosition: null,
			},
			usedHeaderFooterParts: [],
			currentTabs: [],
				konvaStage: stage,
				konvaLayer: layer,
				overflowContentElement: undefined,
				checkingOverflow: false,
			};

		await this.renderPages(document.documentPart.body);

		this.session.konvaStage.visible(false);
		this.refreshTabStops();
		this.session.konvaStage.destroy();
		this.session = null;
	}

	private assignSourcePaths(children: OpenXmlElement[]) {
		children.forEach((child, index) => {
			const path = `body/${index}`;
			child.sourcePath = path;
			this.assignNestedSourcePaths(child, path);
		});
	}

	private assignNestedSourcePaths(element: OpenXmlElement, path: string) {
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

	private renderDefaultStyle() {
		return renderDefaultStyleFn(this.className);
	}

	private renderTheme(themePart: ThemePart, styleContainer: HTMLElement) {
		renderThemeFn(themePart, styleContainer, this.documentStylesCallbacks());
	}

	private processStyleName(className: string): string {
		return processStyleNameFn(className, this.className);
	}

	private processStyles(styles: IDomStyle[]) {
		return processStylesFn(styles, this.documentStylesCallbacks());
	}

	private renderStyles(styles: IDomStyle[]): HTMLElement {
		return renderStylesFn(styles, this.documentStylesCallbacks());
	}

	private processNumberings(numberings: IDomNumbering[]) {
		processNumberingsFn(numberings, this.numberingStylesCallbacks());
	}

	private renderNumbering(numberings: IDomNumbering[], styleContainer: HTMLElement) {
		return renderNumberingFn(numberings, styleContainer, this.numberingStylesCallbacks());
	}

	private numberingClass(id: string, lvl: number) {
		return numberingClassFn(this.className, id, lvl);
	}

	private styleToString(selectors: string, declarations: Record<string, string>, cssText: string = null) {
		return styleToStringFn(selectors, declarations, cssText);
	}

	private numberingCounter(id: string, lvl: number) {
		return numberingCounterFn(this.className, id, lvl);
	}

	private levelTextToContent(text: string, suff: string, id: string, numformat: string) {
		return levelTextToContentFn(text, suff, id, numformat, (counterId, level) => this.numberingCounter(counterId, level));
	}

	private numFormatToCssValue(format: string) {
		return numFormatToCssValueFn(format);
	}

	private renderFontTable(fontsPart: FontTablePart, styleContainer: HTMLElement) {
		renderFontTableFn(fontsPart, styleContainer, this.documentStylesCallbacks());
	}

	private renderWrapper() {
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

	private splitPageBySymbol(documentElement: DocumentElement): Page[] {
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

	private async renderPages(document: DocumentElement) {
		let pages: Page[];
		if (this.options.breakPages) {
			pages = this.splitPageBySymbol(document);
		} else {
			pages = [new Page({ isSplit: true, sectProps: document.sectProps, children: document.children, } as PageProps)];
		}
			this.session.pages = pages;
			let prevProps = null;
		let origin_pages = [...pages];
		for (let i = 0; i < origin_pages.length; i++) {
			this.session.currentFootnoteIds = [];
			const page: Page = origin_pages[i];
			const { sectProps } = page;
			page.sectProps = sectProps ?? document.sectProps;
			page.isFirstPage = page.layoutContext?.isFirstSectionPage ?? prevProps != page.sectProps;
			page.isLastPage = i === origin_pages.length - 1;
				this.session.checkingOverflow = false;
				this.session.currentPage = page;
			prevProps = page.sectProps;
			await this.renderPage();
		}
	}

	private async renderPage() {
		const { pageId, sectProps, children, isFirstPage, isLastPage, regions, layoutContext } = this.session.currentPage;
		processElement(this.session.currentPage);
		const pageElement = this.createPage(this.className, sectProps);

		renderStyleValues(this.document.documentPart.body.cssStyle, pageElement);

			const pages = this.session.pages;
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
					this.session.overflowContentElement = pageElement;
					this.session.checkingOverflow = this.options.breakPages;
				isOverflow = await this.renderElements(region.children, regionArticle, { regionIndex });
				if (isOverflow !== Overflow.FALSE && isOverflow !== Overflow.UNKNOWN && isOverflow !== Overflow.IGNORE) {
					break;
				}
			}
			if (isOverflow === Overflow.FALSE || isOverflow === Overflow.UNKNOWN || isOverflow === Overflow.IGNORE) {
				this.session.currentPage.isSplit = true;
				pages[pageIndex] = this.session.currentPage;
			}
				this.session.checkingOverflow = false;
			} else {
			const contentElement = this.createPageContent(sectProps);
			if (this.options.breakPages) {
				contentElement.style.height = `${contentHeight}pt`;
			} else {
				contentElement.style.minHeight = `${contentHeight}pt`;
			}
				pageElement.appendChild(contentElement);
				this.session.overflowContentElement = contentElement;
				this.session.checkingOverflow = true;
			const is_overflow = await this.renderElements(children, contentElement);
			if (is_overflow === Overflow.FALSE) {
				this.session.currentPage.isSplit = true;
				pages[pageIndex] = this.session.currentPage;
			}
				this.session.checkingOverflow = false;
			}

		if (this.options.renderFootnotes) {
			await this.renderNotes(DomType.Footnotes, this.session.currentFootnoteIds, this.footnoteMap, pageElement);
		}
		if (this.options.renderEndnotes && isLastPage) {
			await this.renderNotes(DomType.Endnotes, this.session.currentEndnoteIds, this.endnoteMap, pageElement);
		}
	}

	private createPage(className: string, props: SectionProperties) {
		return createPageFn(className, props, this.wrapper, {
			ignoreWidth: this.options.ignoreWidth,
			ignoreHeight: this.options.ignoreHeight,
		});
	}

	private createPageContent(props: SectionProperties): HTMLElement {
		return createPageContentFn(props);
	}

	private async renderHeaderFooterRef(refs: FooterHeaderReference[], props: SectionProperties, pageIndex: number, isFirstPage: boolean, layoutContext: PageLayoutContext | undefined, parent: HTMLElement) {
		return renderHeaderFooterRefFn(refs, props, pageIndex, isFirstPage, layoutContext, parent, this.pageRendererCallbacks());
	}

	private pageRendererCallbacks(): PageRendererCallbacks {
		return {
			document: this.document,
			ignoreWidth: this.options.ignoreWidth,
			ignoreHeight: this.options.ignoreHeight,
			evenAndOddHeaders: this.document.settingsPart.settings.evenAndOddHeaders,
			usedHeaderFooterParts: this.session.usedHeaderFooterParts,
			setCurrentPart: part => { this.session.currentPart = part; },
			processElement: elem => processElement(elem),
			renderHeaderFooter: (elem, tagName, parent) => this.renderHeaderFooter(elem, tagName, parent),
		};
	}

	private async renderNotes(type: DomType = DomType.Footnotes, noteIds: string[], notesMap: Record<string, WmlBaseNote>, parent: HTMLElement) {
		return renderNotesFn(type, noteIds, notesMap, parent, {
			processElement: (e) => processElement(e),
			renderChildren: (e, p) => this.renderChildren(e, p),
		});
	}

	private async renderElements(children: OpenXmlElement[], parent: HTMLElement | Element | Text, splitContext?: RenderSplitContext): Promise<Overflow> {
		const overflows: Overflow[] = [];
			const pages = this.session.pages;
		const { pageId } = this.session.currentPage;
		const pageIndex = pages.findIndex(p => p.pageId === pageId);
		type RenderAction = 'continue' | 'break' | 'break-after-current';
		const BREAKING_OVERFLOWS = new Set([Overflow.SELF, Overflow.TRUE, Overflow.FULL, Overflow.PART]);

		for (let i = 0; i < children.length; i++) {
			const elem = children[i];
			elem.index = i;
			if (!elem.breakIndex) elem.breakIndex = new Set();

			const rendered = await this.renderElement(elem, parent);
			let overflow: Overflow = rendered?.dataset?.overflow as Overflow ?? Overflow.UNKNOWN;
			let action: RenderAction = 'continue';

			const isFirstPageElement = elem.level === 2 && i === 0;

			if (isFirstPageElement && BREAKING_OVERFLOWS.has(overflow)) {
				action = 'break-after-current';
			} else {
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
			}

			if (elem.type === DomType.Cell) action = 'continue';
			overflows.push(overflow);

			if (action === 'break' || action === 'break-after-current') {
				if (elem.level === 2) {
					const overflowIndex = action === 'break-after-current' ? i + 1 : i;
					if (overflowIndex < children.length) {
						if (splitContext) {
							splitRegionOnOverflow(this.session.currentPage, pages, pageIndex, splitContext.regionIndex, overflowIndex);
						} else {
							splitOnOverflow(this.session.currentPage, pages, pageIndex, overflowIndex);
						}
						processElement(this.session.currentPage);
						this.session.currentPage = pages[pageIndex + 1];
						await this.renderPage();
					} else {
						this.session.currentPage.isSplit = true;
						pages[pageIndex] = this.session.currentPage;
					}
				}
				break;
			}
		}

		return inferOverflow(overflows);
	}

	private async renderElement(elem: OpenXmlElement, parent?: HTMLElement | Element | Text): Promise<Node_DOM> {
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
			tableCtx: this.session.tableCtx,
			className: this.className,
			currentTabs: this.session.currentTabs,
			currentFootnoteIds: this.session.currentFootnoteIds,
			currentEndnoteIds: this.session.currentEndnoteIds,
			renderClass: (e, o) => renderClass(e, o, n => this.processStyleName(n)),
			renderStyleValues: (s, o) => renderStyleValues(s, o),
		};
	}

	private isSourceAnchor(elem: OpenXmlElement): boolean {
		return elem.type === DomType.Paragraph || elem.type === DomType.Table || elem.type === DomType.Cell || elem.type === DomType.SectionBreak;
	}

	private async renderChildren(elem: OpenXmlElement, parent: HTMLElement | Element | Text): Promise<Overflow> {
		return await this.renderElements(elem.children, parent);
	}

	private async appendChildren(parent: HTMLElement | Text, children: ChildrenType): Promise<Overflow> {
		appendChildrenSync(parent, children);
		return measurePageOverflow({
			isSplit: this.session.currentPage.isSplit,
			contentElement: this.session.overflowContentElement,
			checkingOverflow: this.session.checkingOverflow,
		});
	}

	private async renderContainer(elem: OpenXmlElement, tagName: keyof HTMLElementTagNameMap, props?: Record<string, any>) {
		const oContainer = createElement(tagName, props);

		oContainer.dataset.overflow = await this.renderChildren(elem, oContainer);
		return oContainer;
	}

	private async renderContainerNS(elem: OpenXmlElement, ns: string, tagName: string, props?: Record<string, any>) {
		const parent = createElementNS(ns, tagName as any, props);
		await this.renderChildren(elem, parent);
		return parent;
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

	private resolveFieldRuns(runs: OpenXmlElement[]): OpenXmlElement[] {
		return resolveFieldRunsFn(runs, this.fieldsCallbacks());
	}

	private resolveSimpleField(elem: WmlFieldSimple): OpenXmlElement[] {
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
					const pages = this.session.pages ?? [];
					return pages.findIndex(p => p.pageId === this.session.currentPage.pageId) + 1;
				},
				pageCount: () => this.session.pages?.length ?? 0,
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
			currentPageIsSplit: () => this.session.currentPage.isSplit,
			currentSectionProperties: () => this.session.currentPage.sectProps,
			findExternalRelation: (id) => this.document.documentPart.rels.find(
				it => it.id == id && it.targetMode === 'External'
			),
		};
	}

	private drawingRenderContext(): DrawingRenderContext {
		return {
			document: this.document,
			currentPart: this.session.currentPart,
			options: this.options,
			className: this.className,
			konvaStage: this.session.konvaStage,
			konvaLayer: this.session.konvaLayer,
			appendChildren: (p, c) => this.appendChildren(p, c),
			renderChildren: (e, p) => this.renderChildren(e, p),
			renderElement: (e, p) => this.renderElement(e, p),
			renderStyleValues: (s, o) => renderStyleValues(s, o),
		};
	}

	private async renderHeaderFooter(elem: OpenXmlElement, tagName: keyof HTMLElementTagNameMap, parent: HTMLElement) {
		return renderHeaderFooterFn(elem, tagName, parent, {
			renderChildren: (e, p) => this.renderChildren(e, p),
			renderStyleValues: (s, o) => renderStyleValues(s, o),
		});
	}

	private findStyle(styleName: string) {
		return styleName && this.styleMap?.[styleName];
	}

	private refreshTabStops() {
		for (const tab of this.session.currentTabs) {
			updateTabStop(tab.span, tab.stops, this.defaultTabSize, this.pointToPixelRatio);
		}
	}
}
