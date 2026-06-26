import Konva from 'konva';
import type { Layer } from 'konva/lib/Layer';
import type { Stage } from 'konva/lib/Stage';
import { DomType, OpenXmlElement } from '../model/element';
import { WmlDrawing, WmlImage } from '../model/drawing';
import { Part } from '../common/part';
import { Options } from '../docx-preview';
import { getPresetGeometryPaths } from '../shapes/preset-geometry';
import { VmlElement } from '../vml/vml';
import { WordDocument } from '../word-document';
import { Overflow, appendChildren, createElement, createSvgElement } from './dom-utils';

export interface DrawingRenderContext {
	document: WordDocument;
	currentPart: Part;
	options: Options;
	konvaStage: Stage;
	konvaLayer: Layer;
	appendChildren(parent: HTMLElement | Text, children: Element): Promise<Overflow>;
	renderChildren(elem: OpenXmlElement, parent: HTMLElement | Element | Text): Promise<Overflow>;
	renderElement(elem: OpenXmlElement, parent?: HTMLElement | Element | Text): Promise<Node>;
	renderStyleValues(style: Record<string, string>, output: HTMLElement): void;
}

export function createKonva(bodyContainer: HTMLElement): { stage: Stage; layer: Layer } {
	const oContainer = createElement('div');
	oContainer.id = 'konva-container';
	appendChildren(bodyContainer, oContainer);

	const stage = new Konva.Stage({ container: 'konva-container' });
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
	const oDrawing = createElement('span');
	oDrawing.style.textIndent = '0px';
	oDrawing.dataset.wrap = elem?.props.wrapType;
	ctx.renderStyleValues(elem.cssStyle, oDrawing);

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

	oImage.dataset.overflow = await ctx.appendChildren(parent, oImage);
	return oImage;
}

export async function renderShape(
	elem: OpenXmlElement,
	parent: HTMLElement,
	ctx: DrawingRenderContext
): Promise<HTMLElement> {
	const oContainer = createElement('span');
	oContainer.style.position = 'relative';
	oContainer.style.display = 'inline-block';
	ctx.renderStyleValues(elem.cssStyle, oContainer);

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
		oText.style.position = 'relative';
		oText.style.width = '100%';
		oText.style.height = '100%';
		oText.style.display = 'flex';
		oText.style.alignItems = 'center';
		oText.style.justifyContent = 'center';
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
