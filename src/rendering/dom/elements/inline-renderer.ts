import { WmlParagraph } from '@docx/ooxml/wordprocessingml/document/model/paragraph';
import { WmlRun } from '@docx/ooxml/wordprocessingml/document/model/run';
import { parseLineSpacing } from '@docx/ooxml/wordprocessingml/document/model/spacing-between-lines';
import { DomType, OpenXmlElement, WmlCharacter, WmlHyperlink, WmlText, WrapType } from '@docx/ooxml/wordprocessingml/document/model/dom';
import { appendChildren, createElement } from '@docx/rendering/dom/core/dom-utils';
import { applyGdiLineHeight } from '@docx/rendering/dom/styles/gdi-line-height';
import { Overflow } from '@docx/rendering/measurement/overflow';
import type { RenderContext } from '@docx/rendering/render-context';
import * as _ from 'lodash-es';

interface Node_DOM extends Node, Text {
	dataset: DOMStringMap;
}

export async function renderParagraph(
	elem: WmlParagraph,
	parent: HTMLElement,
	ctx: RenderContext
): Promise<HTMLParagraphElement> {
	const oParagraph = createElement('p');

	// Evaluate PAGE/NUMPAGES field codes into a local array. Never write this
	// back onto elem.children: header/footer paragraphs are shared objects
	// reused across every page, so mutating them would freeze the field value
	// resolved on the first render (e.g. PAGE always showing "1").
	const children = ctx.resolveFieldRuns(elem.children);
	oParagraph.dataset.uuid = elem.uuid;
	ctx.renderClass(elem, oParagraph);
	// snapToGrid is inheritable through the style chain, but elem.props only ever
	// holds what this paragraph's own pPr literally sets. Resolve the effective
	// value (own pPr > style chain > document default style > true) so docGrid
	// line-pitch snapping below only applies where Word would actually apply it.
	const snapToGrid = ctx.resolveSnapToGrid(elem.styleName, elem.props.snapToGrid);
	Object.assign(elem.cssStyle, parseLineSpacing({ ...elem.props, snapToGrid }, ctx.currentSectionProperties()));
	ctx.renderStyleValues(elem.cssStyle, oParagraph);
	ctx.renderCommonProperties(oParagraph.style, elem.props);

	const style = ctx.findStyle(elem.styleName);
	elem.props.tabs = _.unionBy(elem.props.tabs, style?.paragraphProps?.tabs, 'position');

	const numbering = elem.props.numbering ?? style?.paragraphProps?.numbering;

	if (numbering) {
		oParagraph.classList.add(
			ctx.numberingClass(numbering.id, numbering.level)
		);
	}

	// TODO Run children can contain multiple DrawingML objects; current positioning only handles one reliably.
	const shouldClear = children.some(run => {
		const hasTopAndBottomDrawing = run?.children?.some(
			child => child.type === DomType.Drawing && child.props.wrapType === WrapType.TopAndBottom
		);
		const hasClearBreak = run?.children?.some(
			child => child.type === DomType.Break && child?.props?.clear
		);
		return hasTopAndBottomDrawing || hasClearBreak;
	});

	if (shouldClear) {
		oParagraph.classList.add('clearfix');
	}

	oParagraph.style.position = 'relative';

	const isOverflow = await ctx.appendChildren(parent, oParagraph);
	applyGdiLineHeight(oParagraph);
	if (isOverflow === Overflow.SELF) {
		oParagraph.dataset.overflow = Overflow.SELF;
		return oParagraph;
	}

	oParagraph.dataset.overflow = await ctx.renderElements(children, oParagraph);
	alignDrawingOnlyParagraph(oParagraph, ctx);

	return oParagraph;
}

// Word sizes a line that holds ONLY an inline drawing to exactly the
// drawing's height, while a drawing on a line with text sits on the text
// baseline (leaving the font's descent below it). CSS baseline alignment
// always reserves the paragraph strut's descent, so a drawing-only
// paragraph would render ~1 descent taller than Word and push everything
// below it down. Bottom-aligning the drawings when the paragraph has no
// visible text reproduces Word's exact line height.
function alignDrawingOnlyParagraph(oParagraph: HTMLParagraphElement, ctx: RenderContext): void {
	const drawings = Array.from(oParagraph.querySelectorAll<HTMLElement>(`:scope span.${ctx.className}-drawing, :scope img`));
	if (drawings.length === 0) return;
	if ((oParagraph.textContent ?? '').trim() !== '') return;

	for (const drawing of drawings) {
		drawing.style.verticalAlign = 'text-bottom';
	}
}

export async function renderRun(
	elem: WmlRun,
	parent: HTMLElement,
	ctx: RenderContext
): Promise<HTMLSpanElement> {
	// TODO fieldRun ???
	if (elem.fieldRun) {
		return null;
	}

	const oSpan = createElement('span');
	ctx.renderClass(elem, oSpan);
	ctx.renderStyleValues(elem.cssStyle, oSpan);

	const isOverflow = await ctx.appendChildren(parent, oSpan);
	applyGdiLineHeight(oSpan);
	if (isOverflow === Overflow.SELF) {
		oSpan.dataset.overflow = Overflow.SELF;
		return oSpan;
	}

	if (elem.verticalAlign) {
		const oScript = createElement(elem.verticalAlign as any);
		appendChildren(oSpan, oScript);
		oSpan.dataset.overflow = await ctx.renderChildren(elem, oScript);
		return oSpan;
	}

	oSpan.dataset.overflow = await ctx.renderChildren(elem, oSpan);

	return oSpan;
}

export async function renderText(
	elem: WmlText,
	parent: HTMLElement,
	ctx: RenderContext
): Promise<Node_DOM> {
	const oText = document.createTextNode('') as Node_DOM;
	oText.dataset = { overflow: Overflow.UNCHECKED };
	appendChildren(parent, oText);

	if (ctx.currentPageIsSplit()) {
		oText.appendData(elem.text);
		return oText;
	}

	oText.dataset.overflow = await ctx.renderChildren(elem, oText);

	return oText;
}

export async function renderCharacter(
	elem: WmlCharacter,
	parent: Text,
	ctx: RenderContext
): Promise<Node_DOM> {
	const oCharacter = document.createTextNode(elem.char) as Node_DOM;
	oCharacter.dataset = { overflow: Overflow.UNCHECKED };
	oCharacter.dataset.overflow = await ctx.appendChildren(parent, oCharacter);

	return oCharacter;
}

export async function renderHyperlink(
	elem: WmlHyperlink,
	parent: HTMLElement,
	ctx: RenderContext
): Promise<HTMLAnchorElement> {
	const oAnchor = createElement('a');
	ctx.renderStyleValues(elem.cssStyle, oAnchor);

	const isOverflow = await ctx.appendChildren(parent, oAnchor);
	if (isOverflow === Overflow.SELF) {
		oAnchor.dataset.overflow = Overflow.SELF;
		return oAnchor;
	}

	if (elem.href) {
		oAnchor.href = elem.href;
	} else if (elem.id) {
		const rel = ctx.findExternalRelation(elem.id);
		oAnchor.href = rel?.target;
	}

	oAnchor.dataset.overflow = await ctx.renderChildren(elem, oAnchor);

	return oAnchor;
}
