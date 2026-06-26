import * as _ from 'lodash-es';
import { DomType, OpenXmlElement } from '@docx/ooxml/wordprocessingml/model/element';
import { createElement, createElementNS, appendChildren } from '@docx/rendering/dom/core/dom-utils';
import type { RenderContext } from '@docx/rendering/render-context';

const mathNs = 'http://www.w3.org/1998/Math/MathML';
const mathOperators = new Set(['=', '+', '-', '−', '*', '/', '×', '÷', '±', '<', '>', '≤', '≥']);

// Convert OOXML math justification to CSS text-align value
export function mathJustificationToTextAlign(justification?: string): string {
	switch (justification) {
		case 'left':
			return 'left';
		case 'right':
			return 'right';
		case 'center':
		case 'centerGroup':
		default:
			return 'center';
	}
}

// Render a math paragraph as an indivisible block.
export async function renderMmlMathParagraph(elem: OpenXmlElement, ctx: RenderContext): Promise<HTMLElement> {
	const oContainer = createElement('div');
	oContainer.classList.add(`${ctx.className}-math-paragraph`);
	oContainer.style.textAlign = mathJustificationToTextAlign(elem.props?.justification);
	oContainer.style.textIndent = '0px';
	oContainer.style.breakInside = 'avoid';
	oContainer.style.whiteSpace = 'normal';
	oContainer.dataset.overflow = await ctx.renderChildren(elem, oContainer) as string;
	return oContainer;
}

// Render <msqrt> or <mroot> (radical with optional degree)
export async function renderMmlRadical(elem: OpenXmlElement, ctx: RenderContext): Promise<MathMLElement> {
	const base = elem.children.find(el => el.type === DomType.MmlBase);
	let oParent: MathMLElement;
	if (elem.props?.hideDegree) {
		oParent = createElementNS(mathNs, 'msqrt', null);
		await ctx.renderElements([base], oParent);
		return oParent;
	}
	const degree = elem.children.find(el => el.type === DomType.MmlDegree);
	oParent = createElementNS(mathNs, 'mroot', null);
	await ctx.renderElements([base, degree], oParent);
	return oParent;
}

// Render <mrow> delimiter with optional open/close characters
export async function renderMmlDelimiter(elem: OpenXmlElement, ctx: RenderContext): Promise<MathMLElement> {
	const oMrow: MathMLElement = createElementNS(mathNs, 'mrow', null);
	// Opening character
	const oBegin: MathMLElement = createElementNS(mathNs, 'mo', null, [elem.props.beginChar ?? '(']);
	appendChildren(oMrow, oBegin);
	// Inner content
	await ctx.renderElements(elem.children, oMrow);
	// Closing character
	const oEnd: MathMLElement = createElementNS(mathNs, 'mo', null, [elem.props.endChar ?? ')']);
	appendChildren(oMrow, oEnd);
	return oMrow;
}

// Render an n-ary operator (sum, integral, etc.) with optional sub/superscript
export async function renderMmlNary(elem: OpenXmlElement, ctx: RenderContext): Promise<MathMLElement> {
	const grouped = _.keyBy(elem.children, 'type');

	const sup = grouped[DomType.MmlSuperArgument];
	const sub = grouped[DomType.MmlSubArgument];

	const supElem = sup ? await ctx.renderElement(sup) as unknown as MathMLElement : null;
	const subElem = sub ? await ctx.renderElement(sub) as unknown as MathMLElement : null;
	const charElem: MathMLElement = createElementNS(mathNs, 'mo', null, [elem.props?.char ?? '∫']);
	let operatorElem: MathMLElement = charElem;

	if (supElem && subElem) {
		operatorElem = createElementNS(mathNs, 'munderover', null, [charElem, subElem, supElem]);
	} else if (supElem) {
		operatorElem = createElementNS(mathNs, 'mover', null, [charElem, supElem]);
	} else if (subElem) {
		operatorElem = createElementNS(mathNs, 'munder', null, [charElem, subElem]);
	}

	const oMrow: MathMLElement = createElementNS(mathNs, 'mrow', null);
	appendChildren(oMrow, operatorElem);
	if (grouped[DomType.MmlBase]) {
		await ctx.renderElements(grouped[DomType.MmlBase].children, oMrow);
	}
	return oMrow;
}

// Render a pre-sub-superscript construct (<msubsup> with empty base)
export async function renderMmlPreSubSuper(elem: OpenXmlElement, ctx: RenderContext): Promise<MathMLElement> {
	const children = [];
	const grouped = _.keyBy(elem.children, 'type');

	const sup = grouped[DomType.MmlSuperArgument];
	const sub = grouped[DomType.MmlSubArgument];
	const supElem = sup ? await ctx.renderElement(sup) as unknown as MathMLElement : null;
	const subElem = sub ? await ctx.renderElement(sub) as unknown as MathMLElement : null;
	const stubElem: MathMLElement = createElementNS(mathNs, 'mo', null);

	children.push(createElementNS(mathNs, 'msubsup', null, [stubElem, subElem, supElem]));

	const oMrow: MathMLElement = createElementNS(mathNs, 'mrow', null);
	appendChildren(oMrow, children);
	await ctx.renderElements(grouped[DomType.MmlBase].children, oMrow);
	return oMrow;
}

// Render an <mover>/<munder> group character (accent above/below)
export async function renderMmlGroupChar(elem: OpenXmlElement, ctx: RenderContext): Promise<MathMLElement> {
	const tagName = elem.props.verticalJustification === 'bot' ? 'mover' : 'munder';
	const oGroupChar = await ctx.renderContainerNS(elem, mathNs, tagName);
	if (elem.props.char) {
		const oMo = createElementNS(mathNs, 'mo', null, [elem.props.char]);
		appendChildren(oGroupChar, oMo);
	}
	return oGroupChar as MathMLElement;
}

// Render an overline/underline bar decoration (<mrow> with text-decoration)
export async function renderMmlBar(elem: OpenXmlElement, ctx: RenderContext): Promise<MathMLElement> {
	const oMrow = await ctx.renderContainerNS(elem, mathNs, 'mrow') as MathMLElement;
	switch (elem.props.position) {
		case 'top':
			oMrow.style.textDecoration = 'overline';
			break;
		case 'bottom':
			oMrow.style.textDecoration = 'underline';
			break;
	}
	return oMrow;
}

function mathText(elem: OpenXmlElement): string {
	if (elem.type === DomType.Text || elem.type === DomType.DeletedText) {
		return (elem as OpenXmlElement & { text?: string }).text ?? '';
	}
	if (elem.type === DomType.Character) {
		return (elem as OpenXmlElement & { char?: string }).char ?? '';
	}
	return elem.children?.map(child => mathText(child)).join('') ?? '';
}

function appendMathToken(parent: MathMLElement, text: string, tagName: 'mi' | 'mn' | 'mo' | 'mtext', normal = false): void {
	const token = createElementNS(mathNs, tagName, null, [text]);
	if (normal) {
		token.setAttribute('mathvariant', 'normal');
	}
	appendChildren(parent, token);
}

function appendMathText(parent: MathMLElement, text: string, normalIdentifier = false): void {
	if (!text) {
		return;
	}

	if (normalIdentifier) {
		appendMathToken(parent, text, 'mi', true);
		return;
	}

	const parts = text.match(/\d+(?:\.\d+)?|[A-Za-zΑ-ω]|[^\s]/g) ?? [];
	for (const part of parts) {
		if (/^\d/.test(part)) {
			appendMathToken(parent, part, 'mn');
		} else if (mathOperators.has(part)) {
			appendMathToken(parent, part, 'mo');
		} else {
			appendMathToken(parent, part, 'mi');
		}
	}
}

// Render a math run as MathML token elements.
export async function renderMmlRun(elem: OpenXmlElement, ctx: RenderContext): Promise<MathMLElement> {
	const oMrow = createElementNS(mathNs, 'mrow') as MathMLElement;
	ctx.renderClass(elem, oMrow);
	ctx.renderStyleValues(elem.cssStyle, oMrow as HTMLElement);
	appendMathText(oMrow, mathText(elem), elem.parent?.type === DomType.MmlFunctionName);
	return oMrow;
}

// Render an equation array as an <mtable> where each child is a row
export async function renderMllList(elem: OpenXmlElement, ctx: RenderContext): Promise<HTMLElement> {
	const oMtable = createElementNS(mathNs, 'mtable') as HTMLElement;
	ctx.renderClass(elem, oMtable);
	ctx.renderStyleValues(elem.cssStyle, oMtable);
	for (const child of elem.children) {
		const oChild = await ctx.renderElement(child);
		const oMtd = createElementNS(mathNs, 'mtd', null, [oChild] as any);
		const oMtr = createElementNS(mathNs, 'mtr', null, [oMtd]);
		appendChildren(oMtable, oMtr);
	}
	return oMtable;
}
