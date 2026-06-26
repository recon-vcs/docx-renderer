import { BaseHeaderFooterPart } from '../header-footer/parts';
import { Part } from '../common/part';
import { WordDocument } from '../word-document';
import { DomType, OpenXmlElement } from '../model/element';
import { FooterHeaderReference, SectionProperties } from '../document/section';
import { selectHeaderFooterReference } from '../layout/header-footer-context';
import { PageLayoutContext } from '../layout/page-numbering';
import { createElement } from './dom-utils';

export interface PageRendererCallbacks {
	document: WordDocument;
	ignoreWidth: boolean;
	ignoreHeight: boolean;
	evenAndOddHeaders: boolean;
	usedHeaderFooterParts: string[];
	setCurrentPart(part: Part | null): void;
	processElement(elem: OpenXmlElement): void;
	renderHeaderFooter(elem: OpenXmlElement, tagName: keyof HTMLElementTagNameMap, parent: HTMLElement): Promise<HTMLElement>;
}

export function createPage(
	className: string,
	props: SectionProperties,
	wrapper: HTMLElement,
	options: Pick<PageRendererCallbacks, 'ignoreWidth' | 'ignoreHeight'>
): HTMLElement {
	const oPage = createElement('section', { className });

	if (props) {
		oPage.dataset.sectionId = props.sectionId;
		if (props.pageMargins) {
			oPage.style.paddingLeft = props.pageMargins.left;
			oPage.style.paddingRight = props.pageMargins.right;
			oPage.style.paddingTop = props.pageMargins.top;
			oPage.style.paddingBottom = props.pageMargins.bottom;
		}
		if (props.pageSize) {
			if (!options.ignoreWidth) {
				oPage.style.width = props.pageSize.width;
			}
			if (!options.ignoreHeight) {
				oPage.style.minHeight = props.pageSize.height;
			}
		}
	}

	wrapper.appendChild(oPage);
	return oPage;
}

export function createPageContent(props: SectionProperties): HTMLElement {
	const oArticle = createElement('article');

	if (props.columns) {
		const { count, space, separator } = props.columns;
		if (count > 1) {
			oArticle.style.columnCount = `${count}`;
			oArticle.style.columnGap = space;
			oArticle.style.columnFill = 'auto';
		}
		if (separator) {
			oArticle.style.columnRule = '1px solid black';
		}
	}

	return oArticle;
}

export async function renderHeaderFooterRef(
	refs: FooterHeaderReference[],
	props: SectionProperties,
	pageIndex: number,
	isFirstPage: boolean,
	layoutContext: PageLayoutContext | undefined,
	parent: HTMLElement,
	callbacks: PageRendererCallbacks
): Promise<HTMLElement | null> {
	if (!refs) {
		return null;
	}

	const ref = selectHeaderFooterReference(refs, {
		titlePage: props.titlePage,
		isFirstSectionPage: layoutContext?.isFirstSectionPage ?? isFirstPage,
		evenAndOddHeaders: callbacks.evenAndOddHeaders,
		isEvenPage: layoutContext?.isEvenPage ?? (pageIndex + 1) % 2 === 0,
	});

	if (!ref) {
		console.error("Header/Footer reference is not found");
		return null;
	}

	const part = callbacks.document.findPartByRelId(ref.id, callbacks.document.documentPart) as BaseHeaderFooterPart;
	if (!part) {
		console.error(`Part corresponding to the reference with id:${ref.id} is not found`);
		return null;
	}

	callbacks.setCurrentPart(part);

	if (!callbacks.usedHeaderFooterParts.includes(part.path)) {
		callbacks.processElement(part.rootElement);
		callbacks.usedHeaderFooterParts.push(part.path);
	}

	let oElement: HTMLElement = null;
	switch (part.rootElement.type) {
		case DomType.Header:
			part.rootElement.cssStyle = {
				left: props.pageMargins?.left,
				'padding-top': props.pageMargins.header,
				width: props.contentSize?.width,
			};
			oElement = await callbacks.renderHeaderFooter(part.rootElement, 'header', parent);
			break;
		case DomType.Footer:
			part.rootElement.cssStyle = {
				left: props.pageMargins?.left,
				'padding-bottom': props.pageMargins.footer,
				width: props.contentSize?.width,
			};
			oElement = await callbacks.renderHeaderFooter(part.rootElement, 'footer', parent);
			break;
		default:
			console.warn('set header/footer style error', part.rootElement.type);
			break;
	}

	callbacks.setCurrentPart(null);
	return oElement;
}
