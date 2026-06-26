import Konva from 'konva';
import type { Layer } from 'konva/lib/Layer';
import type { Stage } from 'konva/lib/Stage';
import { DomType, OpenXmlElement } from '@docx/ooxml/wordprocessingml/model/element';
import { WmlDrawing, WmlImage } from '@docx/ooxml/drawingml/model/drawing';
import { Part } from '@docx/opc/parts/part';
import type { Options } from '@docx/options';
import { getPresetGeometryPaths } from '@docx/ooxml/drawingml/shapes/preset-geometry';
import { VmlElement } from '@docx/ooxml/vml/vml';
import { WordDocument } from '@docx/word-document';
import { appendChildren, createElement, createSvgElement } from '@docx/rendering/dom/core/dom-utils';
import { Overflow } from '@docx/rendering/measurement/overflow';

export interface DrawingRenderContext {
	document: WordDocument;
	currentPart: Part;
	options: Options;
	className: string;
	konvaStage: Stage;
	konvaLayer: Layer;
	appendChildren(parent: HTMLElement | Text, children: Element): Promise<Overflow>;
	renderChildren(elem: OpenXmlElement, parent: HTMLElement | Element | Text): Promise<Overflow>;
	renderElement(elem: OpenXmlElement, parent?: HTMLElement | Element | Text): Promise<Node>;
	renderStyleValues(style: Record<string, string>, output: HTMLElement): void;
}

async function waitForImageDecode(image: HTMLImageElement): Promise<void> {
	if (typeof image.decode !== 'function') {
		return;
	}

	try {
		await image.decode();
	} catch {
		// Broken or unsupported images should still render whatever the browser can show.
	}
}

export function createKonva(bodyContainer: HTMLElement): { stage: Stage; layer: Layer } {
	const oContainer = createElement('div');
	const containerId = `konva-container-${Math.random().toString(36).slice(2)}`;
	oContainer.id = containerId;
	appendChildren(bodyContainer, oContainer);

	const stage = new Konva.Stage({ container: containerId });
	const layer = new Konva.Layer({ listening: false });
	stage.add(layer);
	stage.visible(true);

	return { stage, layer };
}

export async function renderDrawing(
	elem: WmlDrawing,
	parent: HTMLElement,
	ctx: DrawingRenderContext
): Promise<HTMLElement> {
	const childHasAutoFit = elem.children?.some((c: OpenXmlElement) => c.props?.textbox?.autoFit === 'shape');

	const oDrawing = createElement('span');
	oDrawing.classList.add(`${ctx.className}-drawing`);
	oDrawing.style.textIndent = '0px';
	oDrawing.style.display = 'inline-block';
	oDrawing.dataset.wrap = elem?.props.wrapType;

	if (childHasAutoFit && elem.cssStyle['height']) {
		const drawingStyle = { ...elem.cssStyle };
		delete drawingStyle['height'];
		ctx.renderStyleValues(drawingStyle, oDrawing);
	} else {
		ctx.renderStyleValues(elem.cssStyle, oDrawing);
	}

	const isOverflow = await ctx.appendChildren(parent, oDrawing);
	if (isOverflow === Overflow.TRUE) {
		oDrawing.dataset.overflow = Overflow.SELF;
		return oDrawing;
	}

	oDrawing.dataset.overflow = await ctx.renderChildren(elem, oDrawing);
	return oDrawing;
}

export async function renderImage(
	elem: WmlImage,
	parent: HTMLElement,
	ctx: DrawingRenderContext
): Promise<HTMLImageElement> {
	const { is_clip, is_transform } = elem.props;
	const oImage = new Image();
	ctx.renderStyleValues(elem.cssStyle, oImage);

	const source = await ctx.document.loadDocumentImage(elem.src, ctx.currentPart);
	if (is_clip || is_transform) {
		try {
			oImage.src = await transformImage(elem, source, ctx);
		} catch (e) {
			console.error(`transform ${elem.src} image error:`, e);
		}
	} else {
		oImage.src = source;
	}

	await waitForImageDecode(oImage);
	oImage.dataset.overflow = await ctx.appendChildren(parent, oImage);
	return oImage;
}

export async function renderShape(
	elem: OpenXmlElement,
	parent: HTMLElement,
	ctx: DrawingRenderContext
): Promise<HTMLElement> {
	const textbox = elem.props?.textbox ?? {};
	const autoFit = textbox.autoFit === 'shape';

	// For autoFit textboxes, don't apply fixed height — let content drive size.
	const containerStyle = autoFit
		? Object.fromEntries(Object.entries(elem.cssStyle).filter(([k]) => k !== 'height'))
		: elem.cssStyle;

	const oContainer = createElement('span');
	oContainer.style.position = 'relative';
	oContainer.style.display = 'inline-block';
	ctx.renderStyleValues(containerStyle, oContainer);

	const oSvg = createSvgElement('svg');
	oSvg.setAttribute('viewBox', '0 0 21600 21600');
	oSvg.setAttribute('preserveAspectRatio', 'none');
	oSvg.style.position = 'absolute';
	oSvg.style.inset = '0';
	oSvg.style.width = '100%';
	oSvg.style.height = '100%';
	oSvg.style.overflow = 'visible';

	const fill: string = elem.props?.fill ?? 'none';
	const line: { width?: string; color?: string } = elem.props?.line ?? {};
	const strokeColor = line.color ?? '#000000';
	const strokeWidth = line.width ? (parseFloat(line.width) || 1) : 1;

	for (const d of getPresetGeometryPaths(elem.props?.preset)) {
		const oPath = createSvgElement('path');
		oPath.setAttribute('d', d);
		oPath.setAttribute('fill-rule', 'evenodd');
		oPath.setAttribute('fill', fill);
		oPath.setAttribute('stroke', strokeColor);
		oPath.setAttribute('stroke-width', `${strokeWidth}`);
		oPath.setAttribute('vector-effect', 'non-scaling-stroke');
		oSvg.appendChild(oPath);
	}

	oContainer.appendChild(oSvg);
	const isOverflow = await ctx.appendChildren(parent, oContainer);

	if (elem.children?.length) {
		const oText = createElement('div');
		oText.classList.add(`${ctx.className}-textbox`);
		oText.style.position = 'relative';
		oText.style.boxSizing = 'border-box';
		oText.style.overflow = 'hidden';
			oText.style.display = 'flex';
			oText.style.flexDirection = 'column';
			oText.style.alignItems = 'flex-start';
		oText.style.paddingLeft = textbox.paddingLeft ?? '';
		oText.style.paddingTop = textbox.paddingTop ?? '';
		oText.style.paddingRight = textbox.paddingRight ?? '';
		oText.style.paddingBottom = textbox.paddingBottom ?? '';
		oText.style.justifyContent = textbox.verticalAnchor === 'b' ? 'flex-end' : textbox.verticalAnchor === 'ctr' ? 'center' : 'flex-start';

		if (autoFit) {
			// Let the textbox size to content; use original (pre-rotation) width if available.
			oText.style.width = elem.props?.originalWidth ?? '100%';
			oText.style.height = 'auto';
		} else {
			// Fixed size: use original dimensions to avoid rotated bounding box distortion.
			oText.style.width = elem.props?.originalWidth ?? '100%';
			oText.style.height = elem.props?.originalHeight ?? '100%';
		}

		oContainer.appendChild(oText);
		await ctx.renderChildren(elem, oText);
	}

	oContainer.dataset.overflow = isOverflow;
	return oContainer;
}

export async function transformImage(
	elem: WmlImage,
	source: string,
	ctx: DrawingRenderContext
): Promise<string> {
	const { is_clip, clip, is_transform, transform } = elem.props;
	const img = new Image();
	img.src = source;
	await img.decode();

	const { naturalWidth, naturalHeight } = img;
	ctx.konvaStage.width(naturalWidth);
	ctx.konvaStage.height(naturalHeight);
	ctx.konvaLayer.removeChildren();

	const group = new Konva.Group();
	const image = new Konva.Image({
		image: img,
		x: naturalWidth / 2,
		y: naturalHeight / 2,
		width: naturalWidth,
		height: naturalHeight,
		offset: {
			x: naturalWidth / 2,
			y: naturalHeight / 2,
		},
	});

	if (is_clip) {
		const { left, right, top, bottom } = clip.path;
		const x = naturalWidth * left;
		const y = naturalHeight * top;
		const width = naturalWidth * (1 - left - right);
		const height = naturalHeight * (1 - top - bottom);
		image.crop({ x, y, width, height });
		image.size({ width, height });
	}

	if (is_transform) {
		for (const key in transform) {
			switch (key) {
				case 'scaleX':
					image.scaleX(transform[key]);
					break;
				case 'scaleY':
					image.scaleY(transform[key]);
					break;
				case 'rotate':
					image.rotation(transform[key]);
					break;
			}
		}
	}

	group.add(image);
	ctx.konvaLayer.add(group);

	if (ctx.options.useBase64URL) {
		return group.toDataURL();
	}

	const blob = (await group.toBlob()) as Blob;
	return URL.createObjectURL(blob);
}

export async function renderVmlElement(
	elem: VmlElement,
	parent: HTMLElement,
	ctx: DrawingRenderContext
): Promise<SVGElement> {
	const oSvg = createSvgElement('svg');
	oSvg.setAttribute('style', elem.cssStyleText);

	const oChildren = await renderVmlChildElement(elem, ctx);

	if (elem.imageHref?.id) {
		const source = await ctx.document?.loadDocumentImage(elem.imageHref.id, ctx.currentPart);
		oChildren.setAttribute('href', source);
	}

	appendChildren(oSvg, oChildren);

	requestAnimationFrame(() => {
		const bb = (oSvg.firstElementChild as any).getBBox();
		oSvg.setAttribute('width', `${Math.ceil(bb.x + bb.width)}`);
		oSvg.setAttribute('height', `${Math.ceil(bb.y + bb.height)}`);
	});

	if (parent) {
		oSvg.dataset.overflow = await ctx.appendChildren(parent, oSvg);
	}

	return oSvg;
}

export async function renderVmlPicture(
	elem: OpenXmlElement,
	ctx: DrawingRenderContext
): Promise<HTMLElement> {
	const oPictureContainer = createElement('span');
	await ctx.renderChildren(elem, oPictureContainer);
	return oPictureContainer;
}

export async function renderVmlChildElement(
	elem: VmlElement,
	ctx: DrawingRenderContext
): Promise<SVGElement> {
	const oVMLElement = createSvgElement(elem.tagName as any);
	Object.entries(elem.attrs).forEach(([k, v]) => oVMLElement.setAttribute(k, v));

	for (const child of elem.children) {
		if (child.type == DomType.VmlElement) {
			const oChild = await renderVmlChildElement(child as VmlElement, ctx);
			appendChildren(oVMLElement, oChild);
		} else {
			await ctx.renderElement(child as OpenXmlElement, oVMLElement);
		}
	}

	return oVMLElement;
}
