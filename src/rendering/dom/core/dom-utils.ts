import { OpenXmlElement, DomType } from '@docx/ooxml/wordprocessingml/model/element';

export type ChildrenType = Node[] | Node | Element[] | Element | string[] | string;

// Create an HTML element by tag name
export function createElement<T extends keyof HTMLElementTagNameMap>(
	tagName: T,
	props?: Partial<Record<keyof HTMLElementTagNameMap[T], any>>
): HTMLElementTagNameMap[T] {
	return createElementNS(null, tagName, props) as HTMLElementTagNameMap[T];
}

// Create an SVG element by tag name
export function createSvgElement<T extends keyof SVGElementTagNameMap>(
	tagName: T,
	props?: Partial<Record<keyof SVGElementTagNameMap[T], any>>
): SVGElementTagNameMap[T] {
	return createElementNS('http://www.w3.org/2000/svg', tagName, props) as SVGElementTagNameMap[T];
}

// Create a namespaced element (MathML, SVG, or HTML)
export function createElementNS<T extends keyof MathMLElementTagNameMap>(ns: string, tagName: T, props?: Partial<Record<any, any>>, children?: ChildrenType): MathMLElementTagNameMap[T];
export function createElementNS<T extends keyof SVGElementTagNameMap>(ns: string, tagName: T, props?: Partial<Record<any, any>>, children?: ChildrenType): SVGElementTagNameMap[T];
export function createElementNS<T extends keyof HTMLElementTagNameMap>(ns: string, tagName: T, props?: Partial<Record<any, any>>, children?: ChildrenType): HTMLElementTagNameMap[T];
export function createElementNS<T>(ns: string, tagName: T, props?: Partial<Record<any, any>>, children?: ChildrenType): Element | SVGElement | MathMLElement {
	let oParent: Element | SVGElement | MathMLElement;
	switch (ns) {
		case 'http://www.w3.org/1998/Math/MathML':
			oParent = document.createElementNS(ns, tagName as keyof MathMLElementTagNameMap);
			break;
		case 'http://www.w3.org/2000/svg':
			oParent = document.createElementNS(ns, tagName as keyof SVGElementTagNameMap);
			break;
		case 'http://www.w3.org/1999/xhtml':
			oParent = document.createElement(tagName as keyof HTMLElementTagNameMap);
			break;
		default:
			oParent = document.createElement(tagName as keyof HTMLElementTagNameMap);
	}
	if (props) {
		Object.assign(oParent, props);
	}
	if (children) {
		appendChildren(oParent, children);
	}
	return oParent;
}

// Clear all child elements from a container
export function removeAllElements(elem: HTMLElement): void {
	elem.innerHTML = '';
}

// Append children to a parent element (no overflow checking)
export function appendChildren(parent: Element | Text, children: ChildrenType): void {
	if (parent instanceof Element) {
		if (Array.isArray(children)) {
			parent.append(...(children as any[]));
		} else {
			if (typeof (children as any) === 'string') {
				parent.append(children as string);
			} else {
				parent.appendChild(children as Node);
			}
		}
	}
	if (parent instanceof Text) {
		if (Array.isArray(children)) {
			throw new Error('Text append children: children must be text node');
		} else {
			if (children instanceof Text) {
				parent.appendData(children.wholeText);
			}
		}
	}
}

// Remove one or more elements from the DOM
export function removeElements(target: Node[] | Node, parent: HTMLElement | Element | Text): void;
export function removeElements(target: Element[] | Element): void;
export function removeElements(target: ChildrenType, parent?: HTMLElement | Element | Text): void {
	if (parent === undefined) {
		if (Array.isArray(target)) {
			(target as Element[]).forEach(elem => {
				if (elem instanceof Element) {
					elem.remove();
				} else {
					throw new Error('removeElements: target must be Element!');
				}
			});
		} else {
			if (target instanceof Element) {
				target.remove();
			} else {
				throw new Error('removeElements: target must be Element!');
			}
		}
		return;
	}
	if (parent instanceof Text) {
		if (Array.isArray(target)) {
			throw new Error('Text remove target: target must be text node!');
		} else {
			if (target instanceof Text) {
				parent.deleteData(parent.length - (target as Text).length, (target as Text).length);
			}
		}
	}
	if (parent instanceof Element) {
		if (Array.isArray(target)) {
			(target as Element[]).forEach(elem => {
				if (elem instanceof Element) {
					elem.remove();
				} else {
					parent.removeChild(elem as any);
				}
			});
		} else {
			if (target instanceof Element) {
				target.remove();
			} else {
				parent.removeChild(target as any);
			}
		}
	}
}

// Create a <style> element with the given CSS text
export function createStyleElement(cssText: string): HTMLStyleElement {
	return createElement('style', { innerHTML: cssText });
}

// Append an HTML comment node to a container
export function appendComment(elem: HTMLElement, comment: string): void {
	elem.appendChild(document.createComment(comment));
}

// Walk up the parent chain to find the nearest ancestor of a given DomType
export function findParent<T extends OpenXmlElement>(elem: OpenXmlElement, type: DomType): T {
	let parent = elem.parent;
	while (parent != null && parent.type != type) {
		parent = parent.parent;
	}
	return parent as T;
}

// Node with dataset support (HTML element returned from renderElement)
export interface Node_DOM extends Node, Text {
	dataset: DOMStringMap;
}
