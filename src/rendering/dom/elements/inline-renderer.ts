import { WmlParagraph } from '@docx/ooxml/wordprocessingml/document/model/paragraph';
import { WmlRun } from '@docx/ooxml/wordprocessingml/document/model/run';
import { parseLineSpacing } from '@docx/ooxml/wordprocessingml/document/model/spacing-between-lines';
import { DomType, OpenXmlElement, WmlCharacter, WmlHyperlink, WmlText, WrapType } from '@docx/ooxml/wordprocessingml/document/model/dom';
import { IDomStyle } from '@docx/ooxml/wordprocessingml/document/model/style';
import { SectionProperties } from '@docx/ooxml/wordprocessingml/document/model/section';
import { CommonProperties } from '@docx/ooxml/wordprocessingml/document/model/common';
import { appendChildren, createElement } from '@docx/rendering/dom/core/dom-utils';
import { Overflow } from '@docx/rendering/measurement/overflow';
import * as _ from 'lodash-es';

interface Node_DOM extends Node, Text {
	dataset: DOMStringMap;
}

export interface InlineRendererCallbacks {
	appendChildren(parent: HTMLElement | Text, children: Node | Element): Promise<Overflow>;
	renderChildren(elem: OpenXmlElement, parent: HTMLElement | Text): Promise<Overflow>;
	renderClass(elem: OpenXmlElement, output: HTMLElement): void;
	renderCommonProperties(style: CSSStyleDeclaration, props: CommonProperties): void;
	renderStyleValues(style: Record<string, string>, output: HTMLElement): void;
	resolveFieldRuns(runs: OpenXmlElement[]): OpenXmlElement[];
	findStyle(styleName: string): IDomStyle;
	numberingClass(id: string, level: number): string;
	currentPageIsSplit(): boolean;
	currentSectionProperties(): SectionProperties;
	findExternalRelation(id: string): { target?: string } | undefined;
}

export async function renderParagraph(
	elem: WmlParagraph,
	parent: HTMLElement,
	callbacks: InlineRendererCallbacks
): Promise<HTMLParagraphElement> {
	const oParagraph = createElement('p');

	// Evaluate PAGE/NUMPAGES field codes and replace stale cached values.
	elem.children = callbacks.resolveFieldRuns(elem.children);
	oParagraph.dataset.uuid = elem.uuid;
	callbacks.renderClass(elem, oParagraph);
	Object.assign(elem.cssStyle, parseLineSpacing(elem.props, callbacks.currentSectionProperties()));
	callbacks.renderStyleValues(elem.cssStyle, oParagraph);
	callbacks.renderCommonProperties(oParagraph.style, elem.props);

	const style = callbacks.findStyle(elem.styleName);
	elem.props.tabs = _.unionBy(elem.props.tabs, style?.paragraphProps?.tabs, 'position');

	const numbering = elem.props.numbering ?? style?.paragraphProps?.numbering;

	if (numbering) {
		oParagraph.classList.add(
			callbacks.numberingClass(numbering.id, numbering.level)
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

	const isOverflow = await callbacks.appendChildren(parent, oParagraph);
	if (isOverflow === Overflow.TRUE) {
		oParagraph.dataset.overflow = Overflow.SELF;

		return oParagraph;
	}

	oParagraph.dataset.overflow = await callbacks.renderChildren(elem, oParagraph);

	return oParagraph;
}

export async function renderRun(
	elem: WmlRun,
	parent: HTMLElement,
	callbacks: InlineRendererCallbacks
): Promise<HTMLSpanElement> {
	// TODO fieldRun ???
	if (elem.fieldRun) {
		return null;
	}

	const oSpan = createElement('span');
	callbacks.renderClass(elem, oSpan);
	callbacks.renderStyleValues(elem.cssStyle, oSpan);

	const isOverflow = await callbacks.appendChildren(parent, oSpan);
	if (isOverflow === Overflow.TRUE) {
		oSpan.dataset.overflow = Overflow.SELF;

		return oSpan;
	}

	if (elem.verticalAlign) {
		const oScript = createElement(elem.verticalAlign as any);
		appendChildren(oSpan, oScript);
		oSpan.dataset.overflow = await callbacks.renderChildren(elem, oScript);

		return oSpan;
	}

	oSpan.dataset.overflow = await callbacks.renderChildren(elem, oSpan);

	return oSpan;
}

export async function renderText(
	elem: WmlText,
	parent: HTMLElement,
	callbacks: InlineRendererCallbacks
): Promise<Node_DOM> {
	const oText = document.createTextNode('') as Node_DOM;
	oText.dataset = { overflow: Overflow.UNKNOWN };
	appendChildren(parent, oText);

	if (callbacks.currentPageIsSplit()) {
		oText.appendData(elem.text);
		return oText;
	}

	oText.dataset.overflow = await callbacks.renderChildren(elem, oText);

	return oText;
}

export async function renderCharacter(
	elem: WmlCharacter,
	parent: Text,
	callbacks: InlineRendererCallbacks
): Promise<Node_DOM> {
	const oCharacter = document.createTextNode(elem.char) as Node_DOM;
	oCharacter.dataset = { overflow: Overflow.UNKNOWN };
	oCharacter.dataset.overflow = await callbacks.appendChildren(parent, oCharacter);

	return oCharacter;
}

export async function renderHyperlink(
	elem: WmlHyperlink,
	parent: HTMLElement,
	callbacks: InlineRendererCallbacks
): Promise<HTMLAnchorElement> {
	const oAnchor = createElement('a');
	callbacks.renderStyleValues(elem.cssStyle, oAnchor);

	const isOverflow = await callbacks.appendChildren(parent, oAnchor);
	if (isOverflow === Overflow.TRUE) {
		oAnchor.dataset.overflow = Overflow.SELF;

		return oAnchor;
	}

	if (elem.href) {
		oAnchor.href = elem.href;
	} else if (elem.id) {
		const rel = callbacks.findExternalRelation(elem.id);
		oAnchor.href = rel?.target;
	}

	oAnchor.dataset.overflow = await callbacks.renderChildren(elem, oAnchor);

	return oAnchor;
}
