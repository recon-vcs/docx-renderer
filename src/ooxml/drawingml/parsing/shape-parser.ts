import { DomType, OpenXmlElement, WmlImage } from '@docx/ooxml/wordprocessingml/document/model/dom';
import type { DocumentParserOptions } from '@docx/ooxml/wordprocessingml/parsing/document-parser';
import xml from '@docx/xml/parsing/xml-parser';
import { LengthUsage, convertLength } from '@docx/ooxml/wordprocessingml/document/model/common';

// Callbacks for shape functions that need to parse embedded body content
export interface ShapeParserCallbacks {
	parseBodyElements(node: Element): OpenXmlElement[];
}

export function parseGraphic(
	elem: Element,
	options: DocumentParserOptions,
	callbacks: ShapeParserCallbacks
): OpenXmlElement {
	let graphicData = xml.element(elem, "graphicData");

	for (let n of xml.elements(graphicData)) {
		switch (n.localName) {
			// TODO DrawML其他元素
			// shape图形
			case "wsp":
				return parseShape(n, options, callbacks);

			// 图片
			case "pic":
				return parsePicture(n, options);

			default:
				if (options.debug) {
					console.warn(`DOCX:%c Unknown Graphic Element：${n.localName}`, 'color:#f75607');
				}
		}
	}

	return null;
}

// 解析图形shape
export function parseShape(
	node: Element,
	options: DocumentParserOptions,
	callbacks: ShapeParserCallbacks
): OpenXmlElement {
	let shape: OpenXmlElement = {
		type: DomType.Shape,
		cssStyle: {},
		children: [],
		props: {
			is_transform: false,
			transform: {},
		},
	};

	for (let n of xml.elements(node)) {
		switch (n.localName) {
			// 图形属性
			case "spPr":
				parseShapeProperties(n, shape, options);
				break;

			// 形状中的文本正文
			case "txbx":
			case "linkedTxbx": {
				let txbxContent = xml.element(n, "txbxContent");
				if (txbxContent) {
					shape.children.push(...callbacks.parseBodyElements(txbxContent));
				}
				break;
			}

			// 非几何信息，渲染不需要
			case "cNvPr":
			case "cNvSpPr":
			case "cNvCnPr":
			// 图形样式
			case "style":
			case "bodyPr":
				parseTextBodyProperties(n, shape);
				break;

			default:
				if (options.debug) {
					console.warn(`DOCX:%c Unknown Shape Element：${n.localName}`, 'color:#f75607');
				}
		}
	}
	return shape;
}

function parseTextBodyProperties(node: Element, target: OpenXmlElement): void {
	let autoFit = false;
	for (const child of xml.elements(node)) {
		if (child.localName === "spAutoFit") {
			autoFit = true;
		}
	}
	target.props.textbox = {
		paddingLeft: xml.lengthAttr(node, "lIns", LengthUsage.Emu),
		paddingTop: xml.lengthAttr(node, "tIns", LengthUsage.Emu),
		paddingRight: xml.lengthAttr(node, "rIns", LengthUsage.Emu),
		paddingBottom: xml.lengthAttr(node, "bIns", LengthUsage.Emu),
		verticalAnchor: xml.attr(node, "anchor"),
		autoFit,
	};
}

// 图形属性
export function parseShapeProperties(
	node: Element,
	target: OpenXmlElement,
	options: DocumentParserOptions
): void {
	for (let n of xml.elements(node)) {
		switch (n.localName) {
			case "xfrm":
				// 注意：存在多种变换组合的情况,需要统一合并处理
				// 水平翻转
				let flipH = xml.boolAttr(n, "flipH");
				if (flipH) {
					target.props.is_transform = true;
					target.props.transform.scaleX = -1;
				}
				// 垂直翻转
				let flipV = xml.boolAttr(n, "flipV");
				if (flipV) {
					target.props.is_transform = true;
					target.props.transform.scaleY = -1;
				}
				// 旋转角度
				let degree = xml.numberAttr(n, "rot", LengthUsage.degree, 0);
				if (degree) {
					target.props.is_transform = true;
					target.props.transform.rotate = degree;
				}
				// 子元素
				parseTransform2D(n, target, options);
				break;

			// 预制图形（矩形、椭圆、箭头等）
			case "prstGeom":
				target.props.preset = xml.attr(n, "prst");
				break;

			// 自定义路径图形，几何路径暂不支持，仅标记为自定义
			case "custGeom":
				target.props.preset = "custom";
				break;

			case "noFill":
				target.props.fill = "none";
				break;

			case "solidFill":
				target.props.fill = parseSolidFillColor(n);
				break;

			// 边框/线条
			case "ln":
				target.props.line = parseShapeLine(n);
				break;

			case "gradFill":
			case "blipFill":
			case "pattFill":
			case "grpFill":
			case "effectLst":
			case "effectDag":
			case "scene3d":
			case "sp3d":
			case "extLst":
				break;

			default:
				if (options.debug) {
					console.warn(`DOCX:%c Unknown Shape Property：${n.localName}`, 'color:#f75607');
				}
		}
	}
}

// 单色填充，仅支持直接RGB值；主题色（schemeClr）解析复杂，暂不支持
export function parseSolidFillColor(node: Element): string | null {
	let srgbClr = xml.element(node, "srgbClr");
	return srgbClr ? `#${xml.attr(srgbClr, "val")}` : null;
}

// 图形线条（边框）
export function parseShapeLine(node: Element): { width?: string; color?: string } {
	let result: { width?: string; color?: string } = {};
	let width = xml.intAttr(node, "w", 0);
	if (width) {
		result.width = String(convertLength(width, LengthUsage.Emu));
	}
	let fill = xml.element(node, "solidFill");
	if (fill) {
		result.color = parseSolidFillColor(fill);
	}
	return result;
}

// 解析图片
export function parsePicture(
	elem: Element,
	options: DocumentParserOptions
): WmlImage {
	let result: WmlImage = {
		type: DomType.Image,
		src: "",
		cssStyle: {},
		props: {
			is_clip: false,
			clip: {},
			is_transform: false,
			transform: {},
		}
	};
	for (let n of xml.elements(elem)) {
		switch (n.localName) {
			case "nvPicPr":
				break;
			case "blipFill":
				parseBlipFill(n, result, options);
				break;
			case "spPr":
				parseShapeProperties(n, result, options);
				break;
			default:
				if (options.debug) {
					console.warn(`DOCX:%c Unknown Picture Element：${n.localName}`, 'color:#f75607');
				}
		}
	}

	return result;
}

// 2D变換
export function parseTransform2D(
	node: Element,
	target: OpenXmlElement,
	options: DocumentParserOptions
): void {
	for (let n of xml.elements(node)) {
		switch (n.localName) {
			// 変換前の幅高さ（実際は参考値）
			case "ext":
				let { transform } = target.props;
				let origin_width = xml.intAttr(n, "cx", 0);
				let origin_height = xml.intAttr(n, "cy", 0);
				// 实际的宽高，单位emu
				let width: number;
				let height: number;
				// 根据旋转角度，重新计算宽高
				if (transform?.rotate) {
					// 换算为数字角度，单位：弧度，注意可能产生负值，-1
					let angel = Math.PI * transform.rotate / 180;
					width = Math.abs(origin_width * Math.cos(angel) + origin_height * Math.sin(angel));
					height = Math.abs(origin_width * Math.sin(angel) + origin_height * Math.cos(angel));
				} else {
					width = origin_width;
					height = origin_height;
				}
				target.props.width = convertLength(width, LengthUsage.Px, false);
				target.props.height = convertLength(height, LengthUsage.Px, false);
				target.props.originalWidth = convertLength(origin_width, LengthUsage.Emu, true);
				target.props.originalHeight = convertLength(origin_height, LengthUsage.Emu, true);
				target.cssStyle["width"] = convertLength(width, LengthUsage.Emu, true);
				target.cssStyle["height"] = convertLength(height, LengthUsage.Emu, true);
				break;

			// 変換後偏移量（実際は参考値）
			case "off":
				target.cssStyle["left"] = xml.lengthAttr(n, "x", LengthUsage.Emu);
				target.cssStyle["top"] = xml.lengthAttr(n, "y", LengthUsage.Emu);
				break;

			default:
				if (options.debug) {
					console.warn(`DOCX:%c Unknown Transform2D Element：${n.localName}`, 'color:#f75607');
				}
		}
	}
}

// 图像填充
export function parseBlipFill(
	node: Element,
	target: WmlImage,
	options: DocumentParserOptions
): void {
	for (let n of xml.elements(node)) {
		switch (n.localName) {
			// 填充效果
			case "blip":
				target.src = xml.attr(n, "embed");
				parseBlip(n, target, options);
				break;
			// 源矩形裁剪：距离源图片的4方位间距，单位百分比（%）
			case "srcRect":
				let left = xml.numberAttr(n, "l", LengthUsage.RelativeRect, 0);
				let right = xml.numberAttr(n, "r", LengthUsage.RelativeRect, 0);
				let top = xml.numberAttr(n, "t", LengthUsage.RelativeRect, 0);
				let bottom = xml.numberAttr(n, "b", LengthUsage.RelativeRect, 0);
				target.props.is_clip = [left, right, top, bottom].some((item) => item !== 0);
				target.props.clip.type = 'inset';
				target.props.clip.path = { top, right, bottom, left };
				break;
			case "stretch":
			case "tile":
				break;

			default:
				if (options.debug) {
					console.warn(`DOCX:%c Unknown Blip Fill Element：${n.localName}`, 'color:#f75607');
				}
		}
	}
}

// 图片填充效果
export function parseBlip(
	node: Element,
	target: OpenXmlElement,
	options: DocumentParserOptions
): void {
	for (let n of xml.elements(node)) {
		switch (n.localName) {
			case "alphaBiLevel":
			case "alphaCeiling":
			case "alphaFloor":
			case "alphaInv":
			case "alphaMod":
				break;
			// 透明度
			case "alphaModFix":
				let opacity = xml.lengthAttr(n, 'amt', LengthUsage.Opacity);
				target.cssStyle["opacity"] = opacity;
				break;

			default:
				if (options.debug) {
					console.warn(`DOCX:%c Unknown Blip Element：${n.localName}`, 'color:#f75607');
				}
				break;
		}
	}
}
