import { DomType, OpenXmlElement, WmlDrawing, WrapType } from '@docx/ooxml/wordprocessingml/document/model/dom';
import type { DocumentParserOptions } from '@docx/ooxml/wordprocessingml/parsing/document-parser';
import xml from '@docx/xml/parsing/xml-parser';
import { xmlUtil } from '@docx/xml/parsing/parse-utils';
import { convertLength, LengthUsage } from '@docx/ooxml/wordprocessingml/document/model/common';
import { parseVmlElement } from '@docx/ooxml/vml/vml';
import { parseGraphic } from '@docx/ooxml/drawingml/parsing/shape-parser';

const supportedNamespaceURIs = [
	"http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
];

// Callbacks required by drawing functions that delegate back to DocumentParser
export interface DrawingParserCallbacks {
	parseBodyElements(node: Element): OpenXmlElement[];
}

function normalizeWrapText(wrapText: string | undefined, posX?: { align?: string }): 'left' | 'right' {
	switch (wrapText) {
		case 'left':
		case 'right':
			return wrapText;
		case 'bothSides':
		case 'largest':
		default:
			return posX?.align === 'right' ? 'left' : 'right';
	}
}

export function parseVmlPicture(
	elem: Element,
	callbacks: DrawingParserCallbacks
): OpenXmlElement {
	const result = { type: DomType.VmlPicture, children: [] as OpenXmlElement[] };

	for (const el of xml.elements(elem)) {
		// callbacks satisfies VmlParserContext (structural match on parseBodyElements)
		const child = parseVmlElement(el, callbacks);
		child && result.children.push(child);
	}

	return result;
}

// 检测备选内容
export function checkAlternateContent(elem: Element): Element {
	if (elem.localName != 'AlternateContent') {
		return elem;
	}

	let choice = xml.element(elem, "Choice");
	// 备选项
	if (choice) {
		let requires = xml.attr(choice, "Requires");
		let namespaceURI = elem.lookupNamespaceURI(requires);

		if (supportedNamespaceURIs.includes(namespaceURI)) {
			return choice.firstElementChild;
		}
	}
	// 回退
	return xml.element(elem, "Fallback")?.firstElementChild;
}

export function parseDrawing(
	node: Element,
	options: DocumentParserOptions,
	callbacks: DrawingParserCallbacks
): OpenXmlElement {
	for (let n of xml.elements(node)) {
		switch (n.localName) {
			case "inline":
			case "anchor":
				return parseDrawingWrapper(n, options, callbacks);
			default:
				if (options.debug) {
					console.warn(`DOCX:%c Unknown Drawing Element：${n.localName}`, 'color:#f75607');
				}
		}
	}
}

// TODO 图片旋转、裁剪之后，文字环绕计算错误
// DrawingML对象有两种状态：内联（inline）-- 对象与文本对齐，浮动（anchor）--对象在文本中浮动，但可以相对于页面进行绝对定位
export function parseDrawingWrapper(
	node: Element,
	options: DocumentParserOptions,
	callbacks: DrawingParserCallbacks
): OpenXmlElement {
	// 是否布局在表格中
	let layoutInCell = xml.boolAttr(node, "layoutInCell");
	// 是否锁定
	let locked = xml.boolAttr(node, "locked");
	// 是否在文字后面显示
	let behindDoc = xml.boolAttr(node, "behindDoc");
	// 是否允许重叠
	let allowOverlap = xml.boolAttr(node, "allowOverlap");
	// 是否简单定位
	let simplePos = xml.boolAttr(node, "simplePos");
	// 层叠数值
	let relativeHeight = xml.intAttr(node, "relativeHeight", 1);
	// 计算DrawML对象相对于文字的上下左右间距；仅在浮动、文字环绕模式下有效；
	let distance = {
		left: xml.lengthAttr(node, "distL", LengthUsage.Emu),
		right: xml.lengthAttr(node, "distR", LengthUsage.Emu),
		top: xml.lengthAttr(node, "distT", LengthUsage.Emu),
		bottom: xml.lengthAttr(node, "distB", LengthUsage.Emu),
		distL: xml.intAttr(node, "distL", 0),
		distR: xml.intAttr(node, "distR", 0),
		distT: xml.intAttr(node, "distT", 0),
		distB: xml.intAttr(node, "distB", 0),
	}

	let result: WmlDrawing = {
		type: DomType.Drawing,
		children: [],
		cssStyle: {},
		props: {
			localName: node.localName,
			wrapType: null,
			layoutInCell,
			locked,
			behindDoc,
			allowOverlap,
			simplePos,
			relativeHeight,
			distance,
			extent: {},
		},
	};

	interface Position {
		relative: string;
		align: string;
		offset: string;
		origin: number;
	}

	// 横轴定位
	let posX: Position = { relative: "page", align: "left", offset: "0pt", origin: 0, };
	// 纵轴定位
	let posY: Position = { relative: "page", align: "top", offset: "0pt", origin: 0, };

	for (let n of xml.elements(node)) {
		switch (n.localName) {
			case "simplePos":
				// 简单定位
				if (simplePos) {
					posX.offset = xml.lengthAttr(n, "x", LengthUsage.Emu);
					posY.offset = xml.lengthAttr(n, "y", LengthUsage.Emu);
					posX.origin = xml.intAttr(n, "x", 0);
					posY.origin = xml.intAttr(n, "y", 0);
				}
				break;

			case "positionH":
				if (!simplePos) {
					let alignNode = xml.element(n, "align");
					let offsetNode = xml.element(n, "posOffset");

					posX.relative = xml.attr(n, "relativeFrom") ?? posX.relative;

					if (alignNode) {
						posX.align = alignNode.textContent;
					}

					if (offsetNode) {
						posX.offset = xmlUtil.sizeValue(offsetNode, LengthUsage.Emu);
						posX.origin = xmlUtil.parseTextContent(offsetNode);
					}
					// 设置横轴的属性
					result.props.posX = posX;
				}
				break;

			case "positionV":
				if (!simplePos) {
					let alignNode = xml.element(n, "align");
					let offsetNode = xml.element(n, "posOffset");

					posY.relative = xml.attr(n, "relativeFrom") ?? posY.relative;

					if (alignNode) {
						posY.align = alignNode.textContent;
					}

					if (offsetNode) {
						posY.offset = xmlUtil.sizeValue(offsetNode, LengthUsage.Emu);
						posY.origin = xmlUtil.parseTextContent(offsetNode);
					}
					// 设置纵轴的属性
					result.props.posY = posY;
				}
				break;

			// drawing外框尺寸
			case "extent":
				result.props.extent = {
					width: xml.lengthAttr(n, "cx", LengthUsage.Emu),
					height: xml.lengthAttr(n, "cy", LengthUsage.Emu),
					origin_width: xml.intAttr(n, "cx", 0),
					origin_height: xml.intAttr(n, "cy", 0),
				};
				break;

			// 特效占据空间
			case "effectExtent":
				result.props.effectExtent = {
					top: xml.lengthAttr(n, "t", LengthUsage.Emu),
					bottom: xml.lengthAttr(n, "b", LengthUsage.Emu),
					left: xml.lengthAttr(n, "l", LengthUsage.Emu),
					right: xml.lengthAttr(n, "r", LengthUsage.Emu),
					origin_top: xml.intAttr(n, "t", 0),
					origin_bottom: xml.intAttr(n, "b", 0),
					origin_left: xml.intAttr(n, "l", 0),
					origin_right: xml.intAttr(n, "r", 0),
				};
				break;

			// 图片
			case "graphic":
				let g = parseGraphic(n, options, callbacks);

				if (g) {
					result.children.push(g);
				}
				break;
			case "wrapTopAndBottom":
				result.props.wrapType = WrapType.TopAndBottom;
				break;

			case "wrapNone":
				result.props.wrapType = WrapType.None;
				break;

			case "wrapSquare":
				result.props.wrapType = WrapType.Square;
				// 文本环绕位置：bothSides、largest、left、right
				result.props.wrapText = xml.attr(n, "wrapText");
				break;

			case "wrapThrough":
			case "wrapTight":
				result.props.wrapType = WrapType.Tight;
				// 文本环绕位置：bothSides、largest、left、right
				result.props.wrapText = xml.attr(n, "wrapText");
				// 多边形数据
				let polygonNode = xml.element(n, "wrapPolygon");
				parsePolygon(polygonNode, result);
				break;
			default:
				if (options.debug) {
					console.warn(`DOCX:%c Unknown Drawing Property：${n.localName}`, 'color:#f75607');
				}
		}
	}
	// 重新计算DrawWrapper的空间
	let { extent, effectExtent } = result.props;
	let real_width = extent.origin_width + effectExtent.origin_left + effectExtent.origin_right;
	let real_height = extent.origin_height + effectExtent.origin_top + effectExtent.origin_bottom;
	result.cssStyle["width"] = convertLength(real_width, LengthUsage.Emu);
	result.cssStyle["height"] = convertLength(real_height, LengthUsage.Emu);
	// 内联（inline）--嵌入型环绕
	if (node.localName === "inline") {
		result.props.wrapType = WrapType.Inline;
	}
	// 浮动（anchor）--其他环绕
	if (node.localName === "anchor") {
		// 根据relativeHeight设置z-index
		result.cssStyle["position"] = "relative";
		// 根据behindDoc判断，衬于文字下方、浮于文字上方
		if (behindDoc) {
			result.cssStyle["z-index"] = -1;
		} else {
			result.cssStyle["z-index"] = relativeHeight;
		}
		// 图片文字环绕默认采用Inline
		if (options.ignoreImageWrap) {
			result.props.wrapType = WrapType.Inline;
		}
		// 文本环绕位置：bothSides、largest、left、right
		let { wrapType } = result.props;
		let wrapText = normalizeWrapText(result.props.wrapText, posX);

		switch (wrapType) {
			// 顶部底部文字环绕
			case WrapType.TopAndBottom:
				result.cssStyle['float'] = 'left';
				result.cssStyle['width'] = "100%";
				// 水平对齐方式，目前仅支持left、right、center
				result.cssStyle['text-align'] = posX.align;
				// 横轴位移补偿
				result.cssStyle["transform"] = `translate(${posX.offset},0)`;
				// 垂直方向，纵轴位移
				result.cssStyle["margin-top"] = `calc(${posY.offset} - ${distance.top})`;
				// 计算距离顶部的inset
				result.cssStyle["shape-outside"] = `inset(calc(${posY.offset} - ${distance.top}) 0 0 0)`;
				// TODO 图片位于文字中间，定位计算错误
				// DrawML对象与文字的上下间距
				result.cssStyle["box-sizing"] = "content-box";
				result.cssStyle["padding-top"] = distance.top;
				result.cssStyle["padding-bottom"] = distance.bottom;
				break;

			// 衬于文字下方、浮于文字上方
			case WrapType.None:
				result.cssStyle['position'] = 'absolute';
				// 水平对齐方式，目前仅支持left、right、center
				switch (posX.align) {
					case "left":
					case "right":
						result.cssStyle[posX.align] = posX.offset;
						break;
					case "center":
						result.cssStyle["left"] = "50%";
						result.cssStyle["transform"] = "translateX(-50%)";
				}
				// 垂直方向，纵轴位移
				result.cssStyle["top"] = posY.offset;

				break;

			// 矩形（四周型）环绕
			case WrapType.Square:
				result.cssStyle["float"] = wrapText === 'left' ? "right" : "left";
				// 垂直方向，纵轴位移
				result.cssStyle["margin-top"] = `calc(${posY.offset} - ${distance.top})`;
				// 计算距离顶部的inset
				result.cssStyle["shape-outside"] = `inset(calc(${posY.offset} - ${distance.top}) 0 0 0)`;
				// wrapText：文字所在的一侧
				switch (wrapText) {
					case "left":
						// 水平对齐方式，目前仅支持left、right、center
						switch (posX.align) {
							case "left":
								// 计算公式：段落width - posX.offset - Drawing对象width - Drawing对象padding-right
								result.cssStyle["margin-right"] = `calc(100% - ${extent.width} - ${posX.offset} - ${distance.right})`;
								break;
							case "right":
								result.cssStyle["margin-right"] = `calc(${posX.offset} - ${distance.right})`;
								break;
							case "center":
								result.cssStyle["margin-right"] = `calc( 50% - (${extent.width} - ${posX.offset}) / 2 - ${distance.right} )`;
						}
						break;
					case "right":
						// 水平对齐方式，目前仅支持left、right、center
						switch (posX.align) {
							case "left":
								result.cssStyle["margin-left"] = `calc(${posX.offset} - ${distance.left})`;
								break;
							case "right":
								// 计算公式：段落width - posX.offset - Drawing对象width - Drawing对象padding-right
								result.cssStyle["margin-left"] = `calc(100% - ${extent.width} - ${posX.offset} - ${distance.left})`;
								result.cssStyle["margin-right"] = `calc(${posX.offset} - ${distance.right})`;
								break;
							case "center":
								result.cssStyle["margin-left"] = `calc( 50% - (${extent.width} - ${posX.offset} ) / 2 - ${distance.left} )`;
						}

						break;
				}
				// DrawML对象与文字的上下间距
				result.cssStyle["box-sizing"] = "content-box";
				result.cssStyle["padding-top"] = distance.top;
				result.cssStyle["padding-bottom"] = distance.bottom;
				result.cssStyle["padding-left"] = distance.left;
				result.cssStyle["padding-right"] = distance.right;

				break;

			// 穿越型环绕
			case WrapType.Through:
			// 紧密型环绕
			case WrapType.Tight:
				result.cssStyle["float"] = wrapText === 'left' ? "right" : "left";
				// 根据多边形设置环绕
				let { polygonData } = result.props;
				result.cssStyle["shape-outside"] = `polygon(${polygonData})`;

				// TODO shape-margin目前4个方位只能设置统一的数值.暂时无法采用

				// 垂直方向，纵轴位移
				// TODO 存在上下padding时，定位错误
				result.cssStyle["margin-top"] = posY.offset;

				switch (wrapText) {
					case "left":
						// 水平对齐方式，目前仅支持left、right、center
						switch (posX.align) {
							case "left":
								// 计算公式：段落width - posX.offset - Drawing对象width
								result.cssStyle["margin-right"] = `calc(100% - ${extent.width} - ${posX.offset})`;
								break;
							case "right":
								result.cssStyle["margin-right"] = posX.offset;
								break;
							case "center":
								result.cssStyle["margin-right"] = `calc( 50% - (${extent.width} - ${posX.offset}) / 2 )`;
						}
						break;
					case "right":
						// 水平对齐方式，目前仅支持left、right、center
						switch (posX.align) {
							case "left":
								result.cssStyle["margin-left"] = posX.offset;
								break;
							case "right":
								// 计算公式：段落width - posX.offset - Drawing对象width
								result.cssStyle["margin-left"] = `calc(100% - ${extent.width} - ${posX.offset})`;
								break;
							case "center":
								result.cssStyle["margin-left"] = `calc( 50% - (${extent.width} - ${posX.offset} ) / 2 )`;
						}
						break;
				}
				break;
		}
	}

	return result;
}

/*
* 多边形端点数据
* Office Open XML将X和Y属性解释为固定坐标空间（21600x21600）中的坐标，每个坐标点在x轴和y轴上都有对应的值，范围从0到21599。
* 固定坐标空间 => 实际坐标空间：
* 实际坐标X = 固定坐标X(EMU) * 图形的Width / 21600
* 实际坐标Y = 固定坐标Y(EMU) * 图形的Height / 21600
*/
export function parsePolygon(node: Element, target: OpenXmlElement): void {
	let polygon = [];
	let { distance, extent, posX, posY } = target.props;
	let wrapText = normalizeWrapText(target.props.wrapText, posX);

	xmlUtil.foreach(node, (elem) => {
		// 原始值，单位：EMU
		let origin_x = xml.intAttr(elem, 'x', 0);
		let origin_y = xml.intAttr(elem, 'y', 0);
		// 实际坐标，单位EMU
		let real_x: number, real_y: number;
		// Point坐标，单位pt
		let point_x: string | number, point_y: string | number;
		// 修正坐标，补偿横向位移
		let revise_x: string | number, revise_y: string | number;
		/*
		* 根据wrapText，转换坐标
		* TODO 多边形：纵轴外边距暂时忽略，横轴补偿distance。当多边形超出DrawWrapper的范围时，补偿会被忽略，导致不准确
		*/
		switch (wrapText) {
			case "left":
				// 水平对齐方式，目前仅支持left、right、center
				switch (posX.align) {
					case "left":
						// 实际坐标
						real_x = origin_x * extent.origin_width / 21600 - distance.distL;
						real_y = origin_y * extent.origin_height / 21600 + posY.origin;
						// 修正坐标
						revise_x = convertLength(real_x, LengthUsage.Emu) ?? "0pt";
						revise_y = convertLength(real_y, LengthUsage.Emu) ?? "0pt";
						break;
					case "right":
						// 实际坐标
						real_x = origin_x * extent.origin_width / 21600 + posX.origin - distance.distL;
						real_y = origin_y * extent.origin_height / 21600 + posY.origin;
						// 修正坐标
						revise_x = convertLength(real_x, LengthUsage.Emu) ?? "0pt";
						revise_y = convertLength(real_y, LengthUsage.Emu) ?? "0pt";
						break;
					case "center":
						// 实际坐标
						real_x = origin_x * extent.origin_width / 21600 + posX.origin - distance.distL;
						real_y = origin_y * extent.origin_height / 21600 + posY.origin;
						// 修正坐标
						revise_x = convertLength(real_x, LengthUsage.Emu) ?? "0pt";
						revise_y = convertLength(real_y, LengthUsage.Emu) ?? "0pt";
				}
				break;
			case "right":
				// 水平对齐方式，目前仅支持left、right、center
				switch (posX.align) {
					case "left":
						// 实际坐标
						real_x = origin_x * extent.origin_width / 21600 + posX.origin + distance.distR;
						real_y = origin_y * extent.origin_height / 21600 + posY.origin;
						// 修正坐标
						revise_x = convertLength(real_x, LengthUsage.Emu) ?? "0pt";
						revise_y = convertLength(real_y, LengthUsage.Emu) ?? "0pt";
						break;
					case "right":
						// 实际坐标
						real_x = origin_x * extent.origin_width / 21600 + posX.origin + distance.distR;
						real_y = origin_y * extent.origin_height / 21600 + posY.origin;
						// Point坐标
						point_x = convertLength(real_x, LengthUsage.Emu) ?? "0pt";
						point_y = convertLength(real_y, LengthUsage.Emu) ?? "0pt";
						// 修正坐标，横轴补偿distance
						revise_x = `calc(100% + ${point_x} - ${extent.width})`;
						revise_y = point_y;

						break;
					case "center":
						// 实际坐标
						real_x = origin_x * extent.origin_width / 21600 + posX.origin + distance.distR;
						real_y = origin_y * extent.origin_height / 21600 + posY.origin;
						// Point坐标
						point_x = convertLength(real_x, LengthUsage.Emu) ?? "0pt";
						point_y = convertLength(real_y, LengthUsage.Emu) ?? "0pt";
						// 修正坐标，横轴补偿distance
						revise_x = `calc(50% + ${point_x})`;
						revise_y = point_y;
				}

				break;
		}

		let point = `${revise_x} ${revise_y}`;
		polygon.push(point);
	});
	target.props.polygonData = polygon.join(',');
}
