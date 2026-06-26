import { OpenXmlElement } from '../model/element';
import { appendChildren, createElement } from './dom-utils';

export interface HeaderFooterRendererCallbacks {
	renderChildren(elem: OpenXmlElement, parent: HTMLElement): Promise<string>;
	renderStyleValues(style: Record<string, string>, output: HTMLElement): void;
}

export async function renderHeaderFooter(
	elem: OpenXmlElement,
	tagName: keyof HTMLElementTagNameMap,
	parent: HTMLElement,
	callbacks: HeaderFooterRendererCallbacks
): Promise<HTMLElement> {
	const oElement = createElement(tagName);
	appendChildren(parent, oElement);
	callbacks.renderStyleValues(elem.cssStyle, oElement);
	await callbacks.renderChildren(elem, oElement);

	return oElement;
}
