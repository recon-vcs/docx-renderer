import { WmlParagraph } from '@docx/ooxml/wordprocessingml/document/model/paragraph';
import { WmlRun } from '@docx/ooxml/wordprocessingml/document/model/run';
import { parseLineSpacing } from '@docx/ooxml/wordprocessingml/document/model/spacing-between-lines';
import { DomType, OpenXmlElement, WmlCharacter, WmlHyperlink, WmlText, WrapType } from '@docx/ooxml/wordprocessingml/document/model/dom';
import { appendChildren, createElement } from '@docx/rendering/dom/core/dom-utils';
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

	// Evaluate PAGE/NUMPAGES field codes and replace stale cached values.
	elem.children = ctx.resolveFieldRuns(elem.children);
	oParagraph.dataset.uuid = elem.uuid;
	ctx.renderClass(elem, oParagraph);
	Object.assign(elem.cssStyle, parseLineSpacing(elem.props, ctx.currentSectionProperties()));
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
	const shouldClear = elem.children.some(run => {
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
	if (isOverflow === Overflow.SELF) {
		oParagraph.dataset.overflow = Overflow.SELF;
		return oParagraph;
	}

	oParagraph.dataset.overflow = await ctx.renderChildren(elem, oParagraph);

	return oParagraph;
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
