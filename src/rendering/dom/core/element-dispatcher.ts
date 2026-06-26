import { BreakType, DomType, OpenXmlElement, WmlBreak, WmlCharacter, WmlHyperlink, WmlImage, WmlLastRenderedPageBreak, WmlNoteReference, WmlSectionBreak, WmlSymbol, WmlTableCell, WmlTableRow, WmlText, } from '@docx/ooxml/wordprocessingml/document/model/dom';
import { WmlTable } from '@docx/ooxml/wordprocessingml/model/table';
import { WmlParagraph } from '@docx/ooxml/wordprocessingml/document/model/paragraph';
import { WmlBookmarkStart } from '@docx/ooxml/wordprocessingml/document/model/bookmarks';
import { WmlFieldSimple } from '@docx/ooxml/wordprocessingml/document/model/fields';
import { WmlCommentRangeStart, WmlCommentReference } from '@docx/ooxml/wordprocessingml/parts/comments/elements';
import { VmlElement } from '@docx/ooxml/vml/vml';
import { ChildrenType, createElement, appendChildren, findParent, Node_DOM } from './dom-utils';
import { Overflow } from '@docx/rendering/measurement/overflow';
import { renderTable as renderTableFn, renderTableRow as renderTableRowFn, renderTableCell as renderTableCellFn } from '@docx/rendering/dom/elements/table-renderer';
import { renderFootnoteReference as renderFootnoteReferenceFn, renderEndnoteReference as renderEndnoteReferenceFn } from '@docx/rendering/dom/elements/notes-renderer';
import { renderMmlMathParagraph as renderMmlMathParagraphFn, renderMmlRadical as renderMmlRadicalFn, renderMmlDelimiter as renderMmlDelimiterFn, renderMmlNary as renderMmlNaryFn, renderMmlPreSubSuper as renderMmlPreSubSuperFn, renderMmlGroupChar as renderMmlGroupCharFn, renderMmlBar as renderMmlBarFn, renderMmlRun as renderMmlRunFn, renderMllList as renderMllListFn } from '@docx/rendering/dom/elements/math-renderer';
import { renderDrawing as renderDrawingFn, renderImage as renderImageFn, renderShape as renderShapeFn, renderVmlElement as renderVmlElementFn, renderVmlPicture as renderVmlPictureFn } from '@docx/rendering/dom/elements/drawing-renderer';
import { renderCharacter as renderCharacterFn, renderHyperlink as renderHyperlinkFn, renderParagraph as renderParagraphFn, renderRun as renderRunFn, renderText as renderTextFn } from '@docx/rendering/dom/elements/inline-renderer';
import { renderBookmarkStart as renderBookmarkStartFn, renderCommentRangeEnd as renderCommentRangeEndFn, renderCommentRangeStart as renderCommentRangeStartFn, renderCommentReference as renderCommentReferenceFn, renderDeleted as renderDeletedFn, renderDeletedText as renderDeletedTextFn, renderInserted as renderInsertedFn, resolveSimpleField as resolveSimpleFieldFn } from '@docx/rendering/dom/elements/fields-renderer';
import type { RenderContext } from '@docx/rendering/render-context';

const ns = {
	html: 'http://www.w3.org/1999/xhtml',
	svg: 'http://www.w3.org/2000/svg',
	mathML: 'http://www.w3.org/1998/Math/MathML',
};

function tableCallbacks(ctx: RenderContext) {
	return {
		renderChildren: (e: OpenXmlElement, p: HTMLElement) => ctx.renderChildren(e, p),
		appendChildren: (p: HTMLElement, c: Element) => ctx.appendChildren(p, c),
		renderClass: (e: OpenXmlElement, o: HTMLElement) => ctx.renderClass(e, o),
		renderStyleValues: (s: Record<string, string>, o: HTMLElement) => ctx.renderStyleValues(s, o),
	};
}

async function renderTab(elem: OpenXmlElement, parent: HTMLElement, ctx: RenderContext): Promise<HTMLElement> {
	const tabSpan = createElement('span');

	tabSpan.innerHTML = '&nbsp;';

	tabSpan.className = `${ctx.className}-tab-stop`;
	const stops = findParent<WmlParagraph>(elem, DomType.Paragraph).props?.tabs;
	ctx.currentTabs.push({ stops, span: tabSpan });

	if (parent) {
		await ctx.appendChildren(parent, tabSpan);
	}

	return tabSpan;
}

async function renderSymbol(elem: WmlSymbol, parent: HTMLElement, ctx: RenderContext): Promise<HTMLElement> {
	const oSymbol = createElement('span');
	oSymbol.style.fontFamily = elem.font;
	oSymbol.innerHTML = `&#x${elem.char};`;
	oSymbol.dataset.overflow = await ctx.appendChildren(parent, oSymbol);
	return oSymbol;
}

async function renderBreak(elem: WmlBreak, parent: HTMLElement, ctx: RenderContext): Promise<HTMLElement> {
	let oBreak: HTMLElement;

	switch (elem.break) {
		case BreakType.Page:
			oBreak = createElement('br');
			oBreak.classList.add('break', 'page');
			break;

		case BreakType.Column:
			oBreak = createElement('br');
			oBreak.classList.add('break', 'column');
			break;

		case BreakType.TextWrapping:
		default:
			oBreak = createElement('br');
			oBreak.classList.add('break', 'textWrap');
			break;
	}
	oBreak.dataset.overflow = await ctx.appendChildren(parent, oBreak);
	return oBreak;
}

async function renderLastRenderedPageBreak(elem: WmlLastRenderedPageBreak, parent: HTMLElement, ctx: RenderContext): Promise<HTMLElement> {
	const oLastRenderedPageBreak = createElement('wbr');
	oLastRenderedPageBreak.classList.add('lastRenderedPageBreak');
	oLastRenderedPageBreak.dataset.overflow = await ctx.appendChildren(parent, oLastRenderedPageBreak);
	return oLastRenderedPageBreak;
}

async function renderSectionBreak(elem: WmlSectionBreak, parent: HTMLElement, ctx: RenderContext): Promise<HTMLElement> {
	const oSectionBreak = createElement('s');
	oSectionBreak.classList.add('break', 'section');
	oSectionBreak.dataset.overflow = await ctx.appendChildren(parent, oSectionBreak);
	oSectionBreak.dataset.type = elem.break;
	return oSectionBreak;
}

// Dispatch a single element to its renderer (tag/path annotation done in renderElement)
export async function dispatchElement(
	elem: OpenXmlElement,
	parent: HTMLElement | Element | Text | undefined,
	ctx: RenderContext,
): Promise<Node_DOM | null> {
	// biome-ignore lint/suspicious/noImplicitAnyLet: return type varies by case
	let oNode: any;

	switch (elem.type) {
		case DomType.Paragraph:
			oNode = await renderParagraphFn(elem as WmlParagraph, parent as HTMLElement, ctx);
			break;

		case DomType.Run:
			oNode = await renderRunFn(elem as any, parent as HTMLElement, ctx);
			break;

		case DomType.SimpleField:
			await ctx.renderElements(
				resolveSimpleFieldFn(elem as WmlFieldSimple, ctx),
				parent as HTMLElement
			);
			oNode = null;
			break;

		case DomType.Text:
			oNode = await renderTextFn(elem as WmlText, parent as HTMLElement, ctx);
			break;

		case DomType.Character:
			oNode = await renderCharacterFn(elem as WmlCharacter, parent as Text, ctx);
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
			oNode = await renderHyperlinkFn(elem as WmlHyperlink, parent as HTMLElement, ctx);
			break;

		case DomType.Drawing:
			oNode = await renderDrawingFn(elem as any, parent as HTMLElement, ctx);
			break;

		case DomType.Image:
			oNode = await renderImageFn(elem as WmlImage, parent as HTMLElement, ctx);
			break;

		case DomType.Shape:
			oNode = await renderShapeFn(elem, parent as HTMLElement, ctx);
			break;

		case DomType.BookmarkStart:
			oNode = renderBookmarkStartFn(elem as WmlBookmarkStart, parent as HTMLElement, ctx);
			break;

		case DomType.BookmarkEnd:
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
			oNode = await renderInsertedFn(elem, parent as HTMLElement, ctx);
			break;

		case DomType.Deleted:
			oNode = await renderDeletedFn(elem, parent as HTMLElement, ctx);
			break;

		case DomType.DeletedText:
			oNode = await renderDeletedTextFn(elem as WmlText, parent as HTMLElement, ctx);
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
			oNode = renderCommentReferenceFn(elem as WmlCommentReference, ctx);
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
			oNode = await renderVmlElementFn(elem as VmlElement, parent as HTMLElement, ctx);
			break;

		case DomType.VmlPicture:
			oNode = await renderVmlPictureFn(elem, ctx);
			if (parent) appendChildren(parent, oNode);
			break;

		case DomType.MmlMath:
			oNode = await ctx.renderContainerNS(elem, ns.mathML, 'math', {
				xmlns: ns.mathML,
			});
			if (parent) {
				oNode.dataset.overflow = await ctx.appendChildren(parent as HTMLElement, oNode);
			}
			break;

		case DomType.MmlMathParagraph:
			oNode = await renderMmlMathParagraphFn(elem, ctx);
			if (parent) appendChildren(parent, oNode);
			break;

		case DomType.MmlFraction:
			oNode = await ctx.renderContainerNS(elem, ns.mathML, 'mfrac');
			if (parent) {
				appendChildren(parent, oNode);
			}
			break;

		case DomType.MmlBase:
			oNode = await ctx.renderContainerNS(elem, ns.mathML, elem.parent.type == DomType.MmlMatrixRow ? "mtd" : "mrow");
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
			if (parent) {
				appendChildren(parent, oNode);
			}
			break;

		case DomType.MmlGroupChar:
			oNode = await renderMmlGroupCharFn(elem, ctx);
			if (parent) appendChildren(parent, oNode);
			break;

		case DomType.MmlLimitLower:
			oNode = await ctx.renderContainerNS(elem, ns.mathML, 'munder');
			if (parent) {
				appendChildren(parent, oNode);
			}
			break;

		case DomType.MmlMatrix:
			oNode = await ctx.renderContainerNS(elem, ns.mathML, 'mtable');
			if (parent) {
				appendChildren(parent, oNode);
			}
			break;

		case DomType.MmlMatrixRow:
			oNode = await ctx.renderContainerNS(elem, ns.mathML, 'mtr');
			if (parent) {
				appendChildren(parent, oNode);
			}
			break;

		case DomType.MmlRadical:
			oNode = await renderMmlRadicalFn(elem, ctx);
			if (parent) appendChildren(parent, oNode);
			break;

		case DomType.MmlSuperscript:
			oNode = await ctx.renderContainerNS(elem, ns.mathML, 'msup');
			if (parent) {
				appendChildren(parent, oNode);
			}
			break;

		case DomType.MmlSubscript:
			oNode = await ctx.renderContainerNS(elem, ns.mathML, 'msub');
			if (parent) {
				appendChildren(parent, oNode);
			}
			break;

		case DomType.MmlDegree:
		case DomType.MmlSuperArgument:
		case DomType.MmlSubArgument:
			oNode = await ctx.renderContainerNS(elem, ns.mathML, 'mrow');
			if (parent) {
				appendChildren(parent, oNode);
			}
			break;

		case DomType.MmlFunctionName:
			oNode = await ctx.renderContainerNS(elem, ns.mathML, 'mrow');
			if (parent) {
				appendChildren(parent, oNode);
			}
			break;

		case DomType.MmlDelimiter:
			oNode = await renderMmlDelimiterFn(elem, ctx);
			if (parent) appendChildren(parent, oNode);
			break;

		case DomType.MmlRun:
			oNode = await renderMmlRunFn(elem, ctx);
			if (parent) appendChildren(parent, oNode);
			break;

		case DomType.MmlNary:
			oNode = await renderMmlNaryFn(elem, ctx);
			if (parent) appendChildren(parent, oNode);
			break;

		case DomType.MmlPreSubSuper:
			oNode = await renderMmlPreSubSuperFn(elem, ctx);
			if (parent) {
				appendChildren(parent, oNode);
			}
			break;

		case DomType.MmlBar:
			oNode = await renderMmlBarFn(elem, ctx);
			if (parent) appendChildren(parent, oNode);
			break;

		case DomType.MmlEquationArray:
			oNode = await renderMllListFn(elem, ctx);
			if (parent) appendChildren(parent, oNode);
			break;
	}

	return oNode ?? null;
}
