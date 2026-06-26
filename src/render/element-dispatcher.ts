import { BreakType, DomType, OpenXmlElement, WmlBreak, WmlCharacter, WmlDrawing, WmlHyperlink, WmlImage, WmlLastRenderedPageBreak, WmlNoteReference, WmlSectionBreak, WmlSymbol, WmlTableCell, WmlTableRow, WmlText, } from '../document/dom';
import { WmlTable } from '../model/table';
import { WmlParagraph } from '../document/paragraph';
import { WmlBookmarkStart } from '../document/bookmarks';
import { WmlFieldSimple } from '../document/fields';
import { WmlCommentRangeStart, WmlCommentReference } from '../comments/elements';
import { VmlElement } from '../vml/vml';
import { Overflow, ChildrenType, createElement, appendChildren, findParent, Node_DOM } from './dom-utils';
import { TableContext, renderTable as renderTableFn, renderTableRow as renderTableRowFn, renderTableCell as renderTableCellFn } from './table-renderer';
import { renderFootnoteReference as renderFootnoteReferenceFn, renderEndnoteReference as renderEndnoteReferenceFn } from './notes-renderer';
import { MathRendererCallbacks, renderMmlMathParagraph as renderMmlMathParagraphFn, renderMmlRadical as renderMmlRadicalFn, renderMmlDelimiter as renderMmlDelimiterFn, renderMmlNary as renderMmlNaryFn, renderMmlPreSubSuper as renderMmlPreSubSuperFn, renderMmlGroupChar as renderMmlGroupCharFn, renderMmlBar as renderMmlBarFn, renderMmlRun as renderMmlRunFn, renderMllList as renderMllListFn } from './math-renderer';
import { DrawingRenderContext, renderDrawing as renderDrawingFn, renderImage as renderImageFn, renderShape as renderShapeFn, renderVmlElement as renderVmlElementFn, renderVmlPicture as renderVmlPictureFn } from './drawing-renderer';
import { InlineRendererCallbacks, renderCharacter as renderCharacterFn, renderHyperlink as renderHyperlinkFn, renderParagraph as renderParagraphFn, renderRun as renderRunFn, renderText as renderTextFn } from './inline-renderer';
import { FieldsRendererCallbacks, renderBookmarkStart as renderBookmarkStartFn, renderCommentRangeEnd as renderCommentRangeEndFn, renderCommentRangeStart as renderCommentRangeStartFn, renderCommentReference as renderCommentReferenceFn, renderDeleted as renderDeletedFn, renderDeletedText as renderDeletedTextFn, renderInserted as renderInsertedFn, resolveSimpleField as resolveSimpleFieldFn } from './fields-renderer';

const ns = {
	html: 'http://www.w3.org/1999/xhtml',
	svg: 'http://www.w3.org/2000/svg',
	mathML: 'http://www.w3.org/1998/Math/MathML',
};

export interface ElementDispatchContext {
	appendChildren(parent: HTMLElement | Text, children: ChildrenType): Promise<Overflow>;
	renderChildren(elem: OpenXmlElement, parent: HTMLElement | Element | Text): Promise<Overflow>;
	renderElements(children: OpenXmlElement[], parent: HTMLElement | Element | Text): Promise<Overflow>;
	renderContainer(elem: OpenXmlElement, tagName: keyof HTMLElementTagNameMap, props?: Record<string, any>): Promise<HTMLElement>;
	renderContainerNS(elem: OpenXmlElement, ns: string, tagName: string, props?: Record<string, any>): Promise<Element>;
	renderHeaderFooter(elem: OpenXmlElement, tagName: keyof HTMLElementTagNameMap, parent: HTMLElement): Promise<HTMLElement>;
	inlineCallbacks(): InlineRendererCallbacks;
	mathCallbacks(): MathRendererCallbacks;
	fieldsCallbacks(): FieldsRendererCallbacks;
	drawingRenderContext(): DrawingRenderContext;
	tableCtx: TableContext;
	className: string;
	currentTabs: Array<{ stops: any; span: HTMLElement }>;
	currentFootnoteIds: string[];
	currentEndnoteIds: string[];
	renderClass(elem: OpenXmlElement, output: HTMLElement | Element): void;
	renderStyleValues(style: Record<string, string>, output: HTMLElement): void;
}

async function renderTab(elem: OpenXmlElement, parent: HTMLElement, ctx: ElementDispatchContext): Promise<HTMLElement> {
	const tabSpan = createElement('span');

	tabSpan.innerHTML = '&nbsp;';

	tabSpan.className = `${ctx.className}-tab-stop`;
	const stops = findParent<WmlParagraph>(elem, DomType.Paragraph).props?.tabs;
	ctx.currentTabs.push({ stops, span: tabSpan });

	// 作为子元素插入，执行溢出检测
	if (parent) {
		await ctx.appendChildren(parent, tabSpan);
	}

	return tabSpan;
}

async function renderSymbol(elem: WmlSymbol, parent: HTMLElement, ctx: ElementDispatchContext): Promise<HTMLElement> {
	const oSymbol = createElement('span');
	oSymbol.style.fontFamily = elem.font;
	oSymbol.innerHTML = `&#x${elem.char};`;
	// 溢出标识
	let is_overflow: Overflow;
	// oSymbol作为子元素插入，针对此元素进行溢出检测
	is_overflow = await ctx.appendChildren(parent, oSymbol);

	if (is_overflow === Overflow.TRUE) {
		oSymbol.dataset.overflow = Overflow.SELF;
	}

	oSymbol.dataset.overflow = is_overflow;

	return oSymbol;
}

// 渲染换行符号
async function renderBreak(elem: WmlBreak, parent: HTMLElement, ctx: ElementDispatchContext): Promise<HTMLElement> {
	let oBreak: HTMLElement;

	switch (elem.break) {
		// 分页符
		case BreakType.Page:
			oBreak = createElement('br');
			// 添加class
			oBreak.classList.add('break', 'page');
			break;

		// 	TODO 分栏符
		case BreakType.Column:
			oBreak = createElement('br');
			// 添加class
			oBreak.classList.add('break', 'column');
			break;

		// 强制换行
		case BreakType.TextWrapping:
		default:
			oBreak = createElement('br');
			// 添加class
			oBreak.classList.add('break', 'textWrap');
			break;
	}
	// oBreak作为子元素插入，针对此元素执行溢出检测
	let isOverflow = await ctx.appendChildren(parent, oBreak);

	if (isOverflow === Overflow.TRUE) {
		isOverflow = Overflow.SELF;
	}

	oBreak.dataset.overflow = isOverflow;

	return oBreak;
}

async function renderLastRenderedPageBreak(elem: WmlLastRenderedPageBreak, parent: HTMLElement, ctx: ElementDispatchContext): Promise<HTMLElement> {
	const oLastRenderedPageBreak = createElement('wbr');
	// 添加class
	oLastRenderedPageBreak.classList.add('lastRenderedPageBreak');
	// oLastRenderedPageBreak作为子元素插入，针对此元素执行溢出检测
	let isOverflow = await ctx.appendChildren(parent, oLastRenderedPageBreak);
	// if true,empty element should be Overflow.SELF
	if (isOverflow === Overflow.TRUE) {
		isOverflow = Overflow.SELF;
	}

	oLastRenderedPageBreak.dataset.overflow = isOverflow;

	return oLastRenderedPageBreak;
}

async function renderSectionBreak(elem: WmlSectionBreak, parent: HTMLElement, ctx: ElementDispatchContext): Promise<HTMLElement> {
	const oSectionBreak = createElement('s');
	// 添加class
	oSectionBreak.classList.add('break', 'section');
	// oSectionBreak作为子元素插入，针对此元素执行溢出检测
	let isOverflow = await ctx.appendChildren(parent, oSectionBreak);
	// if true,empty element should be Overflow.SELF
	if (isOverflow === Overflow.TRUE) {
		isOverflow = Overflow.SELF;
	}

	oSectionBreak.dataset.overflow = isOverflow;
	// break type
	oSectionBreak.dataset.type = elem.break;

	return oSectionBreak;
}

function tableCallbacks(ctx: ElementDispatchContext) {
	return {
		renderChildren: (e: OpenXmlElement, p: HTMLElement) => ctx.renderChildren(e, p),
		appendChildren: (p: HTMLElement, c: Element) => ctx.appendChildren(p, c),
		renderClass: (e: OpenXmlElement, o: HTMLElement) => ctx.renderClass(e, o),
		renderStyleValues: (s: Record<string, string>, o: HTMLElement) => ctx.renderStyleValues(s, o),
	};
}

// 根据XML对象渲染单个元素 (dispatch switch only; tag/path annotation done in renderElement)
export async function dispatchElement(
	elem: OpenXmlElement,
	parent: HTMLElement | Element | Text | undefined,
	ctx: ElementDispatchContext,
): Promise<Node_DOM | null> {
	// biome-ignore lint/suspicious/noImplicitAnyLet: return type varies by case
	let oNode: any;

	switch (elem.type) {
		case DomType.Paragraph:
			oNode = await renderParagraphFn(elem as WmlParagraph, parent as HTMLElement, ctx.inlineCallbacks());
			break;

		case DomType.Run:
			oNode = await renderRunFn(elem as any, parent as HTMLElement, ctx.inlineCallbacks());
			break;

		case DomType.SimpleField:
			// container has no visual representation of its own; its
			// resolved Run(s) are appended directly to parent.
			await ctx.renderElements(
				resolveSimpleFieldFn(elem as WmlFieldSimple, ctx.fieldsCallbacks()),
				parent as HTMLElement
			);
			oNode = null;
			break;

		case DomType.Text:
			oNode = await renderTextFn(elem as WmlText, parent as HTMLElement, ctx.inlineCallbacks());
			break;

		case DomType.Character:
			oNode = await renderCharacterFn(elem as WmlCharacter, parent as Text, ctx.inlineCallbacks());
			break;

		case DomType.Table:
			oNode = await renderTableFn(elem as WmlTable, parent as HTMLElement, ctx.tableCtx, tableCallbacks(ctx));
			break;

		case DomType.Row:
			oNode = await renderTableRowFn(elem as WmlTableRow, parent as HTMLElement, ctx.tableCtx, tableCallbacks(ctx));
			break;

		case DomType.Cell:
			oNode = await renderTableCellFn(elem as WmlTableCell, parent as HTMLElement, ctx.tableCtx, tableCallbacks(ctx));
			break;

		case DomType.Hyperlink:
			oNode = await renderHyperlinkFn(elem as WmlHyperlink, parent as HTMLElement, ctx.inlineCallbacks());
			break;

		case DomType.Drawing:
			oNode = await renderDrawingFn(elem as WmlDrawing, parent as HTMLElement, ctx.drawingRenderContext());
			break;

		case DomType.Image:
			oNode = await renderImageFn(elem as WmlImage, parent as HTMLElement, ctx.drawingRenderContext());
			break;

		case DomType.Shape:
			oNode = await renderShapeFn(elem, parent as HTMLElement, ctx.drawingRenderContext());
			break;

		case DomType.BookmarkStart:
			oNode = renderBookmarkStartFn(elem as WmlBookmarkStart, parent as HTMLElement, ctx.fieldsCallbacks());
			break;

		case DomType.BookmarkEnd:
			//ignore bookmark end
			oNode = null;
			break;

		case DomType.Tab:
			oNode = await renderTab(elem, parent as HTMLElement, ctx);
			break;

		case DomType.Symbol:
			oNode = await renderSymbol(elem as WmlSymbol, parent as HTMLElement, ctx);
			break;

		case DomType.Break:
			oNode = await renderBreak(elem as WmlBreak, parent as HTMLElement, ctx);
			break;

		case DomType.LastRenderedPageBreak:
			oNode = await renderLastRenderedPageBreak(elem as WmlLastRenderedPageBreak, parent as HTMLElement, ctx);
			break;

		case DomType.SectionBreak:
			oNode = await renderSectionBreak(elem as WmlSectionBreak, parent as HTMLElement, ctx);
			break;

		case DomType.Inserted:
			oNode = await renderInsertedFn(elem, parent as HTMLElement, ctx.fieldsCallbacks());
			break;

		case DomType.Deleted:
			oNode = await renderDeletedFn(elem, parent as HTMLElement, ctx.fieldsCallbacks());
			break;

		case DomType.DeletedText:
			oNode = await renderDeletedTextFn(elem as WmlText, parent as HTMLElement, ctx.fieldsCallbacks());
			break;

		case DomType.NoBreakHyphen:
			oNode = createElement('wbr');
			if (parent) {
				await ctx.appendChildren(parent as HTMLElement, oNode);
			}
			break;

		case DomType.CommentRangeStart:
			oNode = renderCommentRangeStartFn(elem as WmlCommentRangeStart);
			if (parent) appendChildren(parent, oNode);
			break;

		case DomType.CommentRangeEnd:
			oNode = renderCommentRangeEndFn(elem as WmlCommentRangeStart);
			if (parent) appendChildren(parent, oNode);
			break;

		case DomType.CommentReference:
			oNode = renderCommentReferenceFn(elem as WmlCommentReference, ctx.fieldsCallbacks());
			if (parent) appendChildren(parent, oNode);
			break;

		case DomType.Footer:
			oNode = await ctx.renderHeaderFooter(elem, 'footer', parent as HTMLElement);
			break;

		case DomType.Header:
			oNode = await ctx.renderHeaderFooter(elem, 'header', parent as HTMLElement);
			break;

		case DomType.Footnote:
		case DomType.Endnote:
			oNode = await ctx.renderContainer(elem, 'li');
			// 作为子元素插入,忽略溢出检测
			if (parent) {
				appendChildren(parent, oNode);
			}
			break;

		case DomType.FootnoteReference:
			oNode = renderFootnoteReferenceFn(elem as WmlNoteReference, ctx.currentFootnoteIds);
			if (parent) appendChildren(parent, oNode);
			break;

		case DomType.EndnoteReference:
			oNode = renderEndnoteReferenceFn(elem as WmlNoteReference, ctx.currentEndnoteIds);
			if (parent) appendChildren(parent, oNode);
			break;

		case DomType.VmlElement:
			oNode = await renderVmlElementFn(elem as VmlElement, parent as HTMLElement, ctx.drawingRenderContext());
			break;

		case DomType.VmlPicture:
			oNode = await renderVmlPictureFn(elem, ctx.drawingRenderContext());
			if (parent) appendChildren(parent, oNode);
			break;

		case DomType.MmlMath:
			oNode = await ctx.renderContainerNS(elem, ns.mathML, 'math', {
				xmlns: ns.mathML,
			});
			// TODO 作为子元素插入,针对此元素进行溢出检测
			if (parent) {
				oNode.dataset.overflow = await ctx.appendChildren(parent as HTMLElement, oNode);
			}
			break;

		case DomType.MmlMathParagraph:
			oNode = await renderMmlMathParagraphFn(elem, ctx.mathCallbacks());
			if (parent) appendChildren(parent, oNode);
			break;

		case DomType.MmlFraction:
			oNode = await ctx.renderContainerNS(elem, ns.mathML, 'mfrac');
			// 作为子元素插入,忽略溢出检测
			if (parent) {
				appendChildren(parent, oNode);
			}
			break;

		case DomType.MmlBase:
			oNode = await ctx.renderContainerNS(elem, ns.mathML, elem.parent.type == DomType.MmlMatrixRow ? "mtd" : "mrow");
			// 作为子元素插入,忽略溢出检测
			if (parent) {
				appendChildren(parent, oNode);
			}
			break;

		case DomType.MmlNumerator:
		case DomType.MmlDenominator:
		case DomType.MmlFunction:
		case DomType.MmlLimit:
		case DomType.MmlBox:
			oNode = await ctx.renderContainerNS(elem, ns.mathML, 'mrow');
			// 作为子元素插入,忽略溢出检测
			if (parent) {
				appendChildren(parent, oNode);
			}
			break;

		case DomType.MmlGroupChar:
			oNode = await renderMmlGroupCharFn(elem, ctx.mathCallbacks());
			if (parent) appendChildren(parent, oNode);
			break;

		case DomType.MmlLimitLower:
			oNode = await ctx.renderContainerNS(elem, ns.mathML, 'munder');
			// 作为子元素插入,忽略溢出检测
			if (parent) {
				appendChildren(parent, oNode);
			}
			break;

		case DomType.MmlMatrix:
			oNode = await ctx.renderContainerNS(elem, ns.mathML, 'mtable');
			// 作为子元素插入,忽略溢出检测
			if (parent) {
				appendChildren(parent, oNode);
			}
			break;

		case DomType.MmlMatrixRow:
			oNode = await ctx.renderContainerNS(elem, ns.mathML, 'mtr');
			// 作为子元素插入,忽略溢出检测
			if (parent) {
				appendChildren(parent, oNode);
			}
			break;

		case DomType.MmlRadical:
			oNode = await renderMmlRadicalFn(elem, ctx.mathCallbacks());
			if (parent) appendChildren(parent, oNode);
			break;

		case DomType.MmlSuperscript:
			oNode = await ctx.renderContainerNS(elem, ns.mathML, 'msup');
			// 作为子元素插入,忽略溢出検測
			if (parent) {
				appendChildren(parent, oNode);
			}
			break;

		case DomType.MmlSubscript:
			oNode = await ctx.renderContainerNS(elem, ns.mathML, 'msub');
			// 作为子元素插入,忽略溢出检测
			if (parent) {
				appendChildren(parent, oNode);
			}
			break;

		case DomType.MmlDegree:
		case DomType.MmlSuperArgument:
		case DomType.MmlSubArgument:
			oNode = await ctx.renderContainerNS(elem, ns.mathML, 'mrow');
			// 作为子元素插入,忽略溢出检测
			if (parent) {
				appendChildren(parent, oNode);
			}
			break;

		case DomType.MmlFunctionName:
			oNode = await ctx.renderContainerNS(elem, ns.mathML, 'mrow');
			// 作为子元素插入,忽略溢出检测
			if (parent) {
				appendChildren(parent, oNode);
			}
			break;

		case DomType.MmlDelimiter:
			oNode = await renderMmlDelimiterFn(elem, ctx.mathCallbacks());
			if (parent) appendChildren(parent, oNode);
			break;

		case DomType.MmlRun:
			oNode = await renderMmlRunFn(elem, ctx.mathCallbacks());
			if (parent) appendChildren(parent, oNode);
			break;

		case DomType.MmlNary:
			oNode = await renderMmlNaryFn(elem, ctx.mathCallbacks());
			if (parent) appendChildren(parent, oNode);
			break;

		case DomType.MmlPreSubSuper:
			oNode = await renderMmlPreSubSuperFn(elem, ctx.mathCallbacks());
			if (parent) {
				appendChildren(parent, oNode);
			}
			break;

		case DomType.MmlBar:
			oNode = await renderMmlBarFn(elem, ctx.mathCallbacks());
			if (parent) appendChildren(parent, oNode);
			break;

		case DomType.MmlEquationArray:
			oNode = await renderMllListFn(elem, ctx.mathCallbacks());
			if (parent) appendChildren(parent, oNode);
			break;
	}

	return oNode ?? null;
}
