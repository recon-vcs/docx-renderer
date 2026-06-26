import * as _ from 'lodash-es';
import { DomType, OpenXmlElement } from '../model/element';
import { createElement, createElementNS, appendChildren } from './dom-utils';

const mathNs = 'http://www.w3.org/1998/Math/MathML';

// Callbacks to the main renderer for math rendering
export interface MathRendererCallbacks {
	renderElement(elem: OpenXmlElement, parent?: HTMLElement | Element): Promise<any>;
	renderElements(elems: OpenXmlElement[], parent: Element): Promise<string>;
	renderChildren(elem: OpenXmlElement, parent: Element): Promise<string>;
	renderContainerNS(elem: OpenXmlElement, ns: string, tagName: string, props?: Record<string, any>): Promise<Element>;
	renderClass(elem: OpenXmlElement, output: Element): void;
	renderStyleValues(style: Record<string, string>, output: HTMLElement): void;
	className: string;
}

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
export async function renderMmlMathParagraph(elem: OpenXmlElement, cbs: MathRendererCallbacks): Promise<HTMLElement> {
	const oContainer = createElement('div');
	oContainer.classList.add(`${cbs.className}-math-paragraph`);
	oContainer.style.textAlign = mathJustificationToTextAlign(elem.props?.justification);
	oContainer.style.textIndent = '0px';
	oContainer.style.breakInside = 'avoid';
	oContainer.style.whiteSpace = 'normal';
	oContainer.dataset.overflow = await cbs.renderChildren(elem, oContainer) as string;
	return oContainer;
}

// Render <msqrt> or <mroot> (radical with optional degree)
export async function renderMmlRadical(elem: OpenXmlElement, cbs: MathRendererCallbacks): Promise<MathMLElement> {
	const base = elem.children.find(el => el.type === DomType.MmlBase);
	let oParent: MathMLElement;
	if (elem.props?.hideDegree) {
		oParent = createElementNS(mathNs, 'msqrt', null);
		await cbs.renderElements([base], oParent);
		return oParent;
	}
	const degree = elem.children.find(el => el.type === DomType.MmlDegree);
	oParent = createElementNS(mathNs, 'mroot', null);
	await cbs.renderElements([base, degree], oParent);
	return oParent;
}

// Render <mrow> delimiter with optional open/close characters
export async function renderMmlDelimiter(elem: OpenXmlElement, cbs: MathRendererCallbacks): Promise<MathMLElement> {
	const oMrow: MathMLElement = createElementNS(mathNs, 'mrow', null);
	// Opening character
	const oBegin: MathMLElement = createElementNS(mathNs, 'mo', null, [elem.props.beginChar ?? '(']);
	appendChildren(oMrow, oBegin);
	// Inner content
	await cbs.renderElements(elem.children, oMrow);
	// Closing character
	const oEnd: MathMLElement = createElementNS(mathNs, 'mo', null, [elem.props.endChar ?? ')']);
	appendChildren(oMrow, oEnd);
	return oMrow;
}

// Render an n-ary operator (sum, integral, etc.) with optional sub/superscript
export async function renderMmlNary(elem: OpenXmlElement, cbs: MathRendererCallbacks): Promise<MathMLElement> {
	const grouped = _.keyBy(elem.children, 'type');

	const sup = grouped[DomType.MmlSuperArgument];
	const sub = grouped[DomType.MmlSubArgument];

	const supElem = sup ? await cbs.renderElement(sup) as MathMLElement : null;
	const subElem = sub ? await cbs.renderElement(sub) as MathMLElement : null;
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
		await cbs.renderElements(grouped[DomType.MmlBase].children, oMrow);
	}
	return oMrow;
}

// Render a pre-sub-superscript construct (<msubsup> with empty base)
export async function renderMmlPreSubSuper(elem: OpenXmlElement, cbs: MathRendererCallbacks): Promise<MathMLElement> {
	const children = [];
	const grouped = _.keyBy(elem.children, 'type');

	const sup = grouped[DomType.MmlSuperArgument];
	const sub = grouped[DomType.MmlSubArgument];
	const supElem = sup ? await cbs.renderElement(sup) as MathMLElement : null;
	const subElem = sub ? await cbs.renderElement(sub) as MathMLElement : null;
	const stubElem: MathMLElement = createElementNS(mathNs, 'mo', null);

	children.push(createElementNS(mathNs, 'msubsup', null, [stubElem, subElem, supElem]));

	const oMrow: MathMLElement = createElementNS(mathNs, 'mrow', null);
	appendChildren(oMrow, children);
	await cbs.renderElements(grouped[DomType.MmlBase].children, oMrow);
	return oMrow;
}

// Render an <mover>/<munder> group character (accent above/below)
export async function renderMmlGroupChar(elem: OpenXmlElement, cbs: MathRendererCallbacks): Promise<MathMLElement> {
	const tagName = elem.props.verticalJustification === 'bot' ? 'mover' : 'munder';
	const oGroupChar = await cbs.renderContainerNS(elem, mathNs, tagName);
	if (elem.props.char) {
		const oMo = createElementNS(mathNs, 'mo', null, [elem.props.char]);
		appendChildren(oGroupChar, oMo);
	}
	return oGroupChar as MathMLElement;
}

// Render an overline/underline bar decoration (<mrow> with text-decoration)
export async function renderMmlBar(elem: OpenXmlElement, cbs: MathRendererCallbacks): Promise<MathMLElement> {
	const oMrow = await cbs.renderContainerNS(elem, mathNs, 'mrow') as MathMLElement;
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

// Render a math run (<ms> with class and style)
export async function renderMmlRun(elem: OpenXmlElement, cbs: MathRendererCallbacks): Promise<HTMLElement> {
	const oMs = createElementNS(mathNs, 'ms') as HTMLElement;
	cbs.renderClass(elem, oMs);
	cbs.renderStyleValues(elem.cssStyle, oMs);
	await cbs.renderChildren(elem, oMs);
	return oMs;
}

// Render an equation array as an <mtable> where each child is a row
export async function renderMllList(elem: OpenXmlElement, cbs: MathRendererCallbacks): Promise<HTMLElement> {
	const oMtable = createElementNS(mathNs, 'mtable') as HTMLElement;
	cbs.renderClass(elem, oMtable);
	cbs.renderStyleValues(elem.cssStyle, oMtable);
	for (const child of elem.children) {
		const oChild = await cbs.renderElement(child);
		const oMtd = createElementNS(mathNs, 'mtd', null, [oChild] as any);
		const oMtr = createElementNS(mathNs, 'mtr', null, [oMtd]);
		appendChildren(oMtable, oMtr);
	}
	return oMtable;
}
