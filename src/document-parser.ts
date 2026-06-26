import { BreakType, DomType, IDomNumbering, NumberingPicBullet, OpenXmlElement, WmlBreak, WmlCharacter, WmlDrawing, WmlHyperlink, WmlImage, WmlLastRenderedPageBreak, WmlNoteReference, WmlSymbol, WmlTable, WmlTableCell, WmlTableColumn, WmlTableRow, WmlText, WrapType } from './document/dom';
import { DocumentElement } from './document/document';
import { parseParagraphProperties, parseParagraphProperty, WmlParagraph } from './document/paragraph';
import { parseSectionProperties, SectionProperties } from './document/section';
import xml from './parser/xml-parser';
import { parseRunProperties, WmlRun } from './document/run';
import { parseBookmarkEnd, parseBookmarkStart } from './document/bookmarks';
import { IDomStyle, Ruleset } from './document/style';
import { WmlFieldChar, WmlFieldSimple } from './document/fields';
import { convertLength, LengthUsage, LengthUsageType } from './document/common';
import { parseVmlElement } from './vml/vml';
import { uuid } from "./utils";
import { WmlComment, WmlCommentRangeEnd, WmlCommentRangeStart, WmlCommentReference } from './comments/elements';
import { parseLineSpacing } from "./document/spacing-between-lines";
import { autos, xmlUtil, values } from './parser/parse-utils';
import {
	parseTable as parseTableFn,
	parseTableColumns as parseTableColumnsFn,
	parseTableProperties as parseTablePropertiesFn,
	parseTablePosition as parseTablePositionFn,
	parseTableRow as parseTableRowFn,
	parseTableRowProperties as parseTableRowPropertiesFn,
	parseTableCell as parseTableCellFn,
	parseTableCellProperties as parseTableCellPropertiesFn,
} from './parser/table-parser';
import {
	parseStylesFile as parseStylesFileFn,
	parseDefaultStyles as parseDefaultStylesFn,
	parseStyle as parseStyleFn,
	parseTableStyle as parseTableStyleFn,
} from './parser/style-parser';
import {
	parseNumberingFile as parseNumberingFileFn,
	parseNumberingPicBullet as parseNumberingPicBulletFn,
	parseAbstractNumbering as parseAbstractNumberingFn,
	parseNumberingLevel as parseNumberingLevelFn,
} from './parser/numbering-parser';

const supportedNamespaceURIs = [
	"http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
];

const mmlTagMap = {
	"oMath": DomType.MmlMath,
	"oMathPara": DomType.MmlMathParagraph,
	"f": DomType.MmlFraction,
	"func": DomType.MmlFunction,
	"fName": DomType.MmlFunctionName,
	"num": DomType.MmlNumerator,
	"den": DomType.MmlDenominator,
	"rad": DomType.MmlRadical,
	"deg": DomType.MmlDegree,
	"e": DomType.MmlBase,
	"sSup": DomType.MmlSuperscript,
	"sSub": DomType.MmlSubscript,
	"sPre": DomType.MmlPreSubSuper,
	"sup": DomType.MmlSuperArgument,
	"sub": DomType.MmlSubArgument,
	"d": DomType.MmlDelimiter,
	"nary": DomType.MmlNary,
	"eqArr": DomType.MmlEquationArray,
	"lim": DomType.MmlLimit,
	"limLow": DomType.MmlLimitLower,
	"m": DomType.MmlMatrix,
	"mr": DomType.MmlMatrixRow,
	"box": DomType.MmlBox,
	"bar": DomType.MmlBar,
	"groupChr": DomType.MmlGroupChar
}

export interface DocumentParserOptions {
	ignoreWidth: boolean;
	debug: boolean;
	ignoreTableWrap: boolean,
	ignoreImageWrap: boolean,
}

// 默认解析选项
export const defaultDocumentParserOptions: DocumentParserOptions = {
	ignoreWidth: false,
	debug: false,
	ignoreTableWrap: true,
	ignoreImageWrap: true,
}

export class DocumentParser {
	options: DocumentParserOptions;

	constructor(options?: Partial<DocumentParserOptions>) {
		this.options = {
			...defaultDocumentParserOptions,
			...options
		};
	}

	parseDocumentFile(xmlDoc: Element): DocumentElement {
		// document elements
		let documentElement: DocumentElement = {
			uuid: 'root',
			pages: [],
			sectProps: {} as SectionProperties,
			type: DomType.Document,
		};
		// 背景色
		let background = xml.element(xmlDoc, "background");
		documentElement.cssStyle = background ? this.parseBackground(background) : {};
		// 处理子元素
		let body = xml.element(xmlDoc, "body");
		documentElement.children = this.parseBodyElements(body);
		// 计算节属性
		let sectionProperties = xml.element(body, "sectPr");
		if (sectionProperties) {
			documentElement.sectProps = parseSectionProperties(sectionProperties, xml);
		}
		// 生成唯一uuid标识
		documentElement.sectProps.sectionId = uuid();

		return documentElement;
	}

	parseBackground(elem: Element): any {
		let result = {};
		let color = xmlUtil.colorAttr(elem, "color");

		if (color) {
			result["background-color"] = color;
		}

		return result;
	}

	parseBodyElements(element: Element): OpenXmlElement[] {
		let children = [];

		xmlUtil.foreach(element, (child) => {
			switch (child.localName) {
				case "p":
					children.push(this.parseParagraph(child));
					break;

				case "tbl":
					children.push(this.parseTable(child));
					break;
				// TODO Structured Document
				case "sdt":
					children.push(...this.parseSdt(child));
					break;

				case "sectPr":
					// ignore,section property has parsed in parseDocumentFile
					break;

				default:
					if (this.options.debug) {
						console.warn(`DOCX:%c Unknown Body Element：${child.localName}`, 'color:red');
					}
			}
		});

		return children;
	}

	parseStylesFile(xstyles: Element): IDomStyle[] {
		return parseStylesFileFn(xstyles, this.options, { parseDefaultProperties: (e, s, c, h) => this.parseDefaultProperties(e, s, c, h) });
	}

	parseDefaultStyles(node: Element): IDomStyle {
		return parseDefaultStylesFn(node, this.options, { parseDefaultProperties: (e, s, c, h) => this.parseDefaultProperties(e, s, c, h) });
	}

	parseStyle(node: Element): IDomStyle {
		return parseStyleFn(node, this.options, { parseDefaultProperties: (e, s, c, h) => this.parseDefaultProperties(e, s, c, h) });
	}

	// TODO: multi-level nested table style rules are not yet applied
	parseTableStyle(node: Element): Ruleset[] {
		return parseTableStyleFn(node, this.options, { parseDefaultProperties: (e, s, c, h) => this.parseDefaultProperties(e, s, c, h) });
	}

	parseNumberingFile(xnums: Element): IDomNumbering[] {
		return parseNumberingFileFn(xnums, this.options, { parseDefaultProperties: (e, s, c, h) => this.parseDefaultProperties(e, s, c, h) });
	}

	parseNumberingPicBullet(elem: Element): NumberingPicBullet {
		return parseNumberingPicBulletFn(elem);
	}

	parseAbstractNumbering(node: Element, bullets: any[]): IDomNumbering[] {
		return parseAbstractNumberingFn(node, bullets, this.options, { parseDefaultProperties: (e, s, c, h) => this.parseDefaultProperties(e, s, c, h) });
	}

	parseNumberingLevel(id: string, node: Element, bullets: any[]): IDomNumbering {
		return parseNumberingLevelFn(id, node, bullets, this.options, { parseDefaultProperties: (e, s, c, h) => this.parseDefaultProperties(e, s, c, h) });
	}

	parseSdt(node: Element): OpenXmlElement[] {
		let result: OpenXmlElement[] = [];
		const sdtContent = xml.element(node, "sdtContent");
		if (sdtContent) {
			result = this.parseBodyElements(sdtContent);
		}
		return result;
	}

	parseNotes(xmlDoc: Element, elemName: string, elemClass: any): any[] {
		let result = [];

		for (let el of xml.elements(xmlDoc, elemName)) {
			const node = new elemClass();
			node.id = xml.attr(el, "id");
			node.noteType = xml.attr(el, "type");
			node.children = this.parseBodyElements(el);
			result.push(node);
		}

		return result;
	}

	parseComments(xmlDoc: Element): any[] {
		let result = [];

		for (let el of xml.elements(xmlDoc, "comment")) {
			const item = new WmlComment();
			item.id = xml.attr(el, "id");
			item.author = xml.attr(el, "author");
			item.initials = xml.attr(el, "initials");
			item.date = xml.attr(el, "date");
			item.children = this.parseBodyElements(el);
			result.push(item);
		}

		return result;
	}

	// TODO Inserted Math Control Character、Inserted Table Row、Inserted Numbering Properties
	parseInserted(node: Element): OpenXmlElement {
		let wmlInserted: OpenXmlElement = {
			type: DomType.Inserted,
			children: [],
		};
		xmlUtil.foreach(node, (child) => {
			switch (child.localName) {
				case "r":
					wmlInserted.children.push(this.parseRun(child));
					break;

				default:
					if (this.options.debug) {
						console.warn(`DOCX:%c Unknown Inserted：${child.localName}`, 'color:#f75607');
					}
			}
		});

		return wmlInserted;
	}

	// TODO
	parseDeleted(node: Element): OpenXmlElement {
		let wmlDeleted: OpenXmlElement = {
			type: DomType.Deleted,
			children: [],
		};
		xmlUtil.foreach(node, (child) => {
			switch (child.localName) {
				case "r":
					wmlDeleted.children.push(this.parseRun(child));
					break;

				default:
					if (this.options.debug) {
						console.warn(`DOCX:%c Unknown Inserted：${child.localName}`, 'color:#f75607');
					}
			}
		});

		return wmlDeleted;

	}

	parseParagraph(node: Element): WmlParagraph {
		let wmlParagraph: WmlParagraph = {
			type: DomType.Paragraph,
			children: [],
			props: {},
			cssStyle: {},
		};

		xmlUtil.foreach(node, (child) => {
			switch (child.localName) {
				// property
				case "pPr":
					this.parseParagraphProperties(child, wmlParagraph);
					break;

				case "r":
					wmlParagraph.children.push(this.parseRun(child));
					break;

				case "hyperlink":
					wmlParagraph.children.push(this.parseHyperlink(child));
					break;

				case "bookmarkStart":
					wmlParagraph.children.push(parseBookmarkStart(child, xml));
					break;

				case "bookmarkEnd":
					wmlParagraph.children.push(parseBookmarkEnd(child, xml));
					break;

				case "commentRangeStart":
					wmlParagraph.children.push(new WmlCommentRangeStart(xml.attr(child, "id")));
					break;

				case "commentRangeEnd":
					wmlParagraph.children.push(new WmlCommentRangeEnd(xml.attr(child, "id")));
					break;

				case "oMath":
				case "oMathPara":
					wmlParagraph.children.push(this.parseMathElement(child));
					break;

				// 	TODO Structured Document Tag
				case "sdt":
					wmlParagraph.children.push(...this.parseSdt(child));
					break;

				// TODO Inserted Math Control Character、Inserted Table Row、Inserted Numbering Properties
				case "ins":
					wmlParagraph.children.push(this.parseInserted(child));
					break;

				case "del":
					wmlParagraph.children.push(this.parseDeleted(child));
					break;

				default:
					if (this.options.debug) {
						console.warn(`DOCX:%c Unknown Paragraph Element：${child.localName}`, 'color:#f75607');
					}
			}
		})

		// when paragraph is empty, a br tag needs to be added to work with the rich text editor and generate line height
		// 当段落children为空，需要添加一个br标签，配合富文本编辑器，同时产生行高
		// TODO 实体符号来替换空行
		if (wmlParagraph.children.length === 0) {
			let wmlBreak: WmlBreak = { type: DomType.Break, "break": BreakType.TextWrapping };
			let wmlRun = { type: DomType.Run, children: [wmlBreak as OpenXmlElement] } as WmlRun;
			wmlParagraph.children = [wmlRun];
		}

		return wmlParagraph;
	}

	parseParagraphProperties(elem: Element, paragraph: WmlParagraph) {
		this.parseDefaultProperties(elem, paragraph.cssStyle = {}, null, c => {
			if (parseParagraphProperty(c, paragraph.props, xml)) {
				return true;
			}

			switch (c.localName) {
				// Paragraph Conditional Formatting
				case "cnfStyle":
					paragraph.className = values.classNameOfCnfStyle(c);
					break;

				// Text Frame Properties
				case "framePr":
					this.parseFrame(c, paragraph);
					break;

				// TODO pStyle should be a property of paragraph
				// Referenced Paragraph Style
				case "pStyle":
					paragraph.styleName = xml.attr(c, "val");
					break;

				default:
					// pass other properties to parseDefaultProperties function
					return false;
			}

			return true;
		});
	}

	parseFrame(node: Element, paragraph: WmlParagraph) {
		let dropCap = xml.attr(node, "dropCap");

		if (dropCap == "drop")
			paragraph.cssStyle["float"] = "left";
	}

	parseHyperlink(node: Element): WmlHyperlink {
		let wmlHyperlink: WmlHyperlink = <WmlHyperlink>{
			type: DomType.Hyperlink,
			children: [],
		};
		let anchor = xml.attr(node, "anchor");
		let relId = xml.attr(node, "id");

		if (anchor) {
			wmlHyperlink.href = "#" + anchor;
		}

		if (relId) {
			wmlHyperlink.id = relId;
		}

		xmlUtil.foreach(node, (child) => {
			switch (child.localName) {
				case "r":
					wmlHyperlink.children.push(this.parseRun(child));
					break;

				default:
					if (this.options.debug) {
						console.warn(`DOCX:%c Unknown Hyperlink Element：${child.localName}`, 'color:#f75607');
					}
			}
		});

		return wmlHyperlink;
	}

	parseRun(node: Element): WmlRun {
		let wmlRun: WmlRun = {
			type: DomType.Run,
			children: [],
		};

		xmlUtil.foreach(node, (child) => {
			// 检测备选内容
			child = this.checkAlternateContent(child);

			switch (child.localName) {
				// property
				case "rPr":
					this.parseRunProperties(child, wmlRun);
					break;

				case "t":
					wmlRun.children.push(this.parseText(child, DomType.Text));
					break;

				case "delText":
					wmlRun.children.push(this.parseText(child, DomType.DeletedText));
					break;

				case "commentReference":
					wmlRun.children.push(new WmlCommentReference(xml.attr(child, "id")));
					break;

				case "fldSimple":
					wmlRun.children.push(<WmlFieldSimple>{
						type: DomType.SimpleField,
						instruction: xml.attr(child, "instr"),
						lock: xml.boolAttr(child, "lock", false),
						dirty: xml.boolAttr(child, "dirty", false)
					});
					break;

				case "instrText":
					wmlRun.fieldRun = true;
					wmlRun.children.push(this.parseText(child, DomType.Instruction));
					break;

				case "fldChar":
					wmlRun.fieldRun = true;
					wmlRun.children.push(<WmlFieldChar>{
						type: DomType.ComplexField,
						charType: xml.attr(child, "fldCharType"),
						lock: xml.boolAttr(child, "lock", false),
						dirty: xml.boolAttr(child, "dirty", false)
					});
					break;

				case "noBreakHyphen":
					wmlRun.children.push({ type: DomType.NoBreakHyphen });
					break;

				case "br":
					wmlRun.children.push(<WmlBreak>{
						type: DomType.Break,
						break: xml.attr(child, "type") || "textWrapping",
						props: {
							clear: xml.attr(child, "clear")
						}
					});
					break;

				case "lastRenderedPageBreak":
					wmlRun.children.push(<WmlLastRenderedPageBreak>{
						type: DomType.LastRenderedPageBreak,
					});
					break;

				// SymbolChar：符号字符
				case "sym":
					wmlRun.children.push(<WmlSymbol>{
						type: DomType.Symbol,
						font: xml.attr(child, "font"),
						char: xml.attr(child, "char")
					});
					break;

				// TODO PositionalTab
				case "ptab":

					break;

				case "tab":
					wmlRun.children.push({ type: DomType.Tab });
					break;

				case "footnoteReference":
					wmlRun.children.push(<WmlNoteReference>{
						type: DomType.FootnoteReference,
						id: xml.attr(child, "id")
					});
					break;

				case "endnoteReference":
					wmlRun.children.push(<WmlNoteReference>{
						type: DomType.EndnoteReference,
						id: xml.attr(child, "id")
					});
					break;

				case "drawing":
					wmlRun.children.push(this.parseDrawing(child));
					break;

				case "pict":
					wmlRun.children.push(this.parseVmlPicture(child));
					break;

				default:
					if (this.options.debug) {
						console.warn(`DOCX:%c Unknown Run Element：${child.localName}`, 'color:#f75607');
					}
			}
		});

		return wmlRun;
	}

	parseText(elem: Element, type: DomType) {
		let wmlText = { type, text: '', } as WmlText;
		let textContent = elem.textContent;
		// 是否保留空格
		let is_preserve_space = xml.attr(elem, "xml:space") === "preserve";
		if (is_preserve_space) {
			// \u00A0 = 不间断空格，英文应该一个空格，中文两个空格。受到font-family影响。
			textContent = textContent.split(/\s/).join("\u00A0");
		}
		// whole text
		wmlText.text = textContent;
		// parse character
		if (textContent.length > 0) {
			wmlText.children = this.parseCharacter(textContent);
		}
		return wmlText;
	}

	parseCharacter(text: string): OpenXmlElement[] {
		let characters = [];
		// 检查字符串是否主要包含中文字符
		const isChinese = text.match(/[\u4e00-\u9fa5]+/g);
		// 主要是中文字符
		if (isChinese) {
			// 待完善正则表达式：/([\u4e00-\u9fff]|\w+)(\p{Punctuation}*)?|\s+/gu;丢失拉丁符号，右括号）
			// TODO 目前拆分方式，英文符号拆分为一个个字母，而非单词。
			characters = text.split('');
		} else {
			characters = text.match(/\S+|\s+/g);
		}
		return characters.map(character => {
			return { type: DomType.Character, char: character } as WmlCharacter
		});
	}

	parseRunProperties(elem: Element, run: WmlRun) {
		this.parseDefaultProperties(elem, run.cssStyle = {}, null, c => {
			switch (c.localName) {
				// Referenced Character Style
				case "rStyle":
					run.styleName = xml.attr(c, "val");
					break;

				// Subscript/Superscript Text
				case "vertAlign":
					run.verticalAlign = values.valueOfVertAlign(c, true);
					break;

				// Character Spacing Adjustment
				case "spacing":
					this.parseSpacing(c, run);
					break;

				default:
					// pass other properties to parseDefaultProperties function
					return false;
			}

			return true;
		});
	}

	parseMathElement(elem: Element): OpenXmlElement {
		const propsTag = `${elem.localName}Pr`;
		const mathElement: OpenXmlElement = {
			type: mmlTagMap[elem.localName],
			children: [],
		};

		xmlUtil.foreach(elem, (child) => {
			const childType = mmlTagMap[child.localName];

			if (childType) {
				mathElement.children.push(this.parseMathElement(child));
			} else if (child.localName == "r") {
				let wmlRun: WmlRun = this.parseRun(child);
				wmlRun.type = DomType.MmlRun;
				mathElement.children.push(wmlRun);
			} else if (child.localName == propsTag) {
				mathElement.props = this.parseMathProperties(child);
			}
		});

		return mathElement;
	}

	parseMathProperties(elem: Element): Record<string, any> {
		const result: Record<string, any> = {};

		for (const el of xml.elements(elem)) {
			switch (el.localName) {
				case "chr":
					result.char = xml.attr(el, "val");
					break;

				case "vertJc":
					result.verticalJustification = xml.attr(el, "val");
					break;

				case "jc":
					result.justification = xml.attr(el, "val");
					break;

				case "pos":
					result.position = xml.attr(el, "val");
					break;

				case "degHide":
					result.hideDegree = xml.boolAttr(el, "val");
					break;

				case "begChr":
					result.beginChar = xml.attr(el, "val");
					break;

				case "endChr":
					result.endChar = xml.attr(el, "val");
					break;

				default:
					if (this.options.debug) {
						console.warn(`DOCX:%c Unknown Math Property：${el.localName}`, 'color:#f75607');
					}
			}
		}

		return result;
	}

	parseVmlPicture(elem: Element): OpenXmlElement {
		const result = { type: DomType.VmlPicture, children: [] };

		for (const el of xml.elements(elem)) {
			const child = parseVmlElement(el, this);
			child && result.children.push(child);
		}

		return result;
	}

	// 检测备选内容
	checkAlternateContent(elem: Element): Element {
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

	parseDrawing(node: Element): OpenXmlElement {
		for (let n of xml.elements(node)) {
			switch (n.localName) {
				case "inline":
				case "anchor":
					return this.parseDrawingWrapper(n);
				default:
					if (this.options.debug) {
						console.warn(`DOCX:%c Unknown Drawing Element：${n.localName}`, 'color:#f75607');
					}
			}
		}
	}

	// TODO 图片旋转、裁剪之后，文字环绕计算错误
	// DrawingML对象有两种状态：内联（inline）-- 对象与文本对齐，浮动（anchor）--对象在文本中浮动，但可以相对于页面进行绝对定位
	parseDrawingWrapper(node: Element): OpenXmlElement {
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
					let g = this.parseGraphic(n);

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
					this.parsePolygon(polygonNode, result);
					break;
				default:
					if (this.options.debug) {
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
			if (this.options.ignoreImageWrap) {
				result.props.wrapType = WrapType.Inline;
			}
			// 文本环绕位置：bothSides、largest、left、right
			let { wrapText, wrapType } = result.props;

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
					// TODO 环绕位置bothSides、largest无法实现，目前仅支持left、right
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
						default:
							console.error(`text wrap picture on ${wrapText} is not supported！`)
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
					// TODO 环绕位置bothSides、largest无法实现，目前仅支持left、right
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
						default:
							console.error(`text wrap picture on ${wrapText} is not supported！`)
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
	parsePolygon(node: Element, target: OpenXmlElement) {
		let polygon = [];
		let { wrapText, distance, extent, posX, posY } = target.props;

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
				default:
					console.error(`text wrap picture on ${wrapText} is not supported！`)
					break;
			}

			let point = `${revise_x} ${revise_y}`;
			polygon.push(point);
		});
		target.props.polygonData = polygon.join(',');
	}

	parseGraphic(elem: Element): OpenXmlElement {
		let graphicData = xml.element(elem, "graphicData");

		for (let n of xml.elements(graphicData)) {
			switch (n.localName) {
				// TODO DrawML其他元素
				// shape图形
				case "wsp":
					return this.parseShape(n);

				// 图片
				case "pic":
					return this.parsePicture(n);

				default:
					if (this.options.debug) {
						console.warn(`DOCX:%c Unknown Graphic Element：${n.localName}`, 'color:#f75607');
					}
			}
		}

		return null;
	}

	// 解析图形shape
	parseShape(node: Element): OpenXmlElement {
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
					this.parseShapeProperties(n, shape);
					break;

				// 形状中的文本正文
				case "txbx":
				case "linkedTxbx": {
					let txbxContent = xml.element(n, "txbxContent");
					if (txbxContent) {
						shape.children.push(...this.parseBodyElements(txbxContent));
					}
					break;
				}

				// 非几何信息，渲染不需要
				case "cNvPr":
				case "cNvSpPr":
				case "cNvCnPr":
				// 图形样式
				case "style":
				// 指定形状中文本正文的正文属性。
				case "bodyPr":
					break;

				default:
					if (this.options.debug) {
						console.warn(`DOCX:%c Unknown Shape Element：${n.localName}`, 'color:#f75607');
					}
			}
		}
		return shape;
	}

	// 图形属性
	parseShapeProperties(node: Element, target: OpenXmlElement) {

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
					this.parseTransform2D(n, target);
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
					target.props.fill = this.parseSolidFillColor(n);
					break;

				// 边框/线条
				case "ln":
					target.props.line = this.parseShapeLine(n);
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
					if (this.options.debug) {
						console.warn(`DOCX:%c Unknown Shape Property：${n.localName}`, 'color:#f75607');
					}
			}
		}
	}

	// 单色填充，仅支持直接RGB值；主题色（schemeClr）解析复杂，暂不支持
	parseSolidFillColor(node: Element): string | null {
		let srgbClr = xml.element(node, "srgbClr");
		return srgbClr ? `#${xml.attr(srgbClr, "val")}` : null;
	}

	// 图形线条（边框）
	parseShapeLine(node: Element): { width?: string; color?: string } {
		let result: { width?: string; color?: string } = {};
		let width = xml.intAttr(node, "w", 0);
		if (width) {
			result.width = String(convertLength(width, LengthUsage.Emu));
		}
		let fill = xml.element(node, "solidFill");
		if (fill) {
			result.color = this.parseSolidFillColor(fill);
		}
		return result;
	}

	// 解析图片
	parsePicture(elem: Element): WmlImage {
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
					this.parseBlipFill(n, result);
					break;

				case "spPr":
					this.parseShapeProperties(n, result)
					break;
				default:
					if (this.options.debug) {
						console.warn(`DOCX:%c Unknown Picture Element：${n.localName}`, 'color:#f75607');
					}
			}
		}

		return result;
	}

	// 2D变换
	parseTransform2D(node: Element, target: OpenXmlElement) {
		for (let n of xml.elements(node)) {
			switch (n.localName) {
				// 变换之前的宽高，实际上无效
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
						// 无旋转
						width = origin_width;
						height = origin_height;
					}
					target.props.width = convertLength(width, LengthUsage.Px, false);
					target.props.height = convertLength(height, LengthUsage.Px, false);
					target.cssStyle["width"] = convertLength(width, LengthUsage.Emu, true);
					target.cssStyle["height"] = convertLength(height, LengthUsage.Emu, true);
					break;

				// 变换之后的偏移量，实际上无效
				case "off":
					target.cssStyle["left"] = xml.lengthAttr(n, "x", LengthUsage.Emu);
					target.cssStyle["top"] = xml.lengthAttr(n, "y", LengthUsage.Emu);
					break;

				default:
					if (this.options.debug) {
						console.warn(`DOCX:%c Unknown Transform2D Element：${n.localName}`, 'color:#f75607');
					}
			}
		}
	}

	// 图像填充
	parseBlipFill(node: Element, target: WmlImage) {
		// 图像填充
		for (let n of xml.elements(node)) {
			switch (n.localName) {
				// 填充效果
				case "blip":
					// embed属性：图片地址
					target.src = xml.attr(n, "embed");
					// 图片填充效果
					this.parseBlip(n, target);
					break;
				// 源矩形裁剪
				case "srcRect":
					// 距离源图片的4方位间距，单位百分比（%）
					let left = xml.numberAttr(n, "l", LengthUsage.RelativeRect, 0);
					let right = xml.numberAttr(n, "r", LengthUsage.RelativeRect, 0);
					let top = xml.numberAttr(n, "t", LengthUsage.RelativeRect, 0);
					let bottom = xml.numberAttr(n, "b", LengthUsage.RelativeRect, 0);
					// 裁剪路径
					target.props.is_clip = [left, right, top, bottom].some((item) => item !== 0);
					target.props.clip.type = 'inset';
					target.props.clip.path = { top, right, bottom, left };
					break;
				case "stretch":
					break;
				// 平铺
				case "tile":
					break;

				default:
					if (this.options.debug) {
						console.warn(`DOCX:%c Unknown Blip Fill Element：${n.localName}`, 'color:#f75607');
					}
			}
		}
	}

	// 图片填充效果
	parseBlip(node: Element, target: OpenXmlElement) {

		for (let n of xml.elements(node)) {
			switch (n.localName) {
				case "alphaBiLevel":

					break;
				case "alphaCeiling":

					break;
				case "alphaFloor":

					break;
				case "alphaInv":

					break;
				case "alphaMod":

					break;
				// 透明度
				case "alphaModFix":
					let opacity = xml.lengthAttr(n, 'amt', LengthUsage.Opacity);
					target.cssStyle["opacity"] = opacity;
					break;

				default:
					if (this.options.debug) {
						console.warn(`DOCX:%c Unknown Blip Element：${n.localName}`, 'color:#f75607');
					}
					break;
			}
		}

	}

	parseTable(node: Element): WmlTable {
		return parseTableFn(node, this.options, {
			parseParagraph: n => this.parseParagraph(n),
			parseTable: n => this.parseTable(n),
			parseDefaultProperties: (e, s, c, h) => this.parseDefaultProperties(e, s, c, h),
		});
	}

	parseTableColumns(node: Element): WmlTableColumn[] {
		return parseTableColumnsFn(node, this.options);
	}

	parseTableProperties(elem: Element, table: WmlTable): void {
		parseTablePropertiesFn(elem, table, this.options, {
			parseParagraph: n => this.parseParagraph(n),
			parseTable: n => this.parseTable(n),
			parseDefaultProperties: (e, s, c, h) => this.parseDefaultProperties(e, s, c, h),
		});
	}

	// Floating table — implements text wrap around table
	parseTablePosition(node: Element, table: WmlTable): void {
		parseTablePositionFn(node, table, this.options);
	}

	parseTableRow(node: Element): WmlTableRow {
		return parseTableRowFn(node, this.options, {
			parseParagraph: n => this.parseParagraph(n),
			parseTable: n => this.parseTable(n),
			parseDefaultProperties: (e, s, c, h) => this.parseDefaultProperties(e, s, c, h),
		});
	}

	parseTableRowProperties(elem: Element, row: WmlTableRow): void {
		parseTableRowPropertiesFn(elem, row, this.options, {
			parseParagraph: n => this.parseParagraph(n),
			parseTable: n => this.parseTable(n),
			parseDefaultProperties: (e, s, c, h) => this.parseDefaultProperties(e, s, c, h),
		});
	}

	parseTableCell(node: Element): OpenXmlElement {
		return parseTableCellFn(node, this.options, {
			parseParagraph: n => this.parseParagraph(n),
			parseTable: n => this.parseTable(n),
			parseDefaultProperties: (e, s, c, h) => this.parseDefaultProperties(e, s, c, h),
		});
	}

	parseTableCellProperties(elem: Element, cell: WmlTableCell): void {
		parseTableCellPropertiesFn(elem, cell, this.options, {
			parseParagraph: n => this.parseParagraph(n),
			parseTable: n => this.parseTable(n),
			parseDefaultProperties: (e, s, c, h) => this.parseDefaultProperties(e, s, c, h),
		});
	}

	// 公共属性，转化为确定的style样式，无需复杂计算
	parseDefaultProperties(elem: Element, style: Record<string, string> = null, childStyle: Record<string, string> = null, handler: (prop: Element) => boolean = null): Record<string, string> {
		style = style || {};

		xmlUtil.foreach(elem, c => {
			/**
			 * 根据提供的handler处理函数和条件执行逻辑。
			 * 如果handler处理函数存在并且调用处理函数返回真值，则终止当前逻辑。
			 *
			 * @param handler 可选的处理函数，接受一个参数 c，并返回一个布尔值。
			 * @param c 传递给处理函数的参数。
			 */
			if (handler?.(c)) {
				return;
			}

			switch (c.localName) {
				// Bold
				case "b":
					style["font-weight"] = xml.boolAttr(c, "val", true) ? "bold" : "normal";
					break;

				//TODO - maybe ignore
				case "bidi":

					break;

				// TODO Complex Script Bold
				case "bCs":
					break;

				// Text Border
				case "bdr":
					style["border"] = values.valueOfBorder(c);
					break;

				// Display All Characters As Capital Letters
				case "caps":
					style["text-transform"] = xml.boolAttr(c, "val", true) ? "uppercase" : "none";
					break;

				// Run Content Color
				case "color":
					style["color"] = xmlUtil.colorAttr(c, "val", null, autos.color);
					break;

				// TODO Use Complex Script Formatting on Run
				case "cs":
					break;

				// TODO Double Strikethrough
				case "dstrike":
					break;

				// TODO East Asian Typography Settings
				case "eastAsianLayout":
					break;

				// TODO Animated Text Effect
				case "effect":
					break;

				// TODO Emphasis Mark
				case "em":
					break;

				// TODO Embossing
				case "emboss":
					break;

				// TODO Manual Run Width
				case "fitText":
					break;

				// Text Highlighting
				case "highlight":
					style["background-color"] = xmlUtil.colorAttr(c, "val", null, autos.highlight);
					break;

				// Italics
				case "i":
					style["font-style"] = xml.boolAttr(c, "val", true) ? "italic" : "normal";
					break;

				// TODO Complex Script Italics
				case "iCs":
					break;

				// TODO Imprinting
				case "imprint":
					break;

				// TODO Font Kerning
				case "kern":
					// style['letter-spacing'] = xml.lengthAttr(c, 'val');
					break;

				// Languages for Run Content,check spelling and grammar
				case "lang":
					style["$lang"] = xml.attr(c, "val");
					break;

				// TODO Do Not Check Spelling or Grammar
				case "noProof":
					break;

				// TODO Display Character Outline
				case "outline":
					break;

				// Vertically Raised or Lowered Text
				case "position":
					style.verticalAlign = xml.lengthAttr(c, "val", LengthUsage.FontSize);
					break;

				// Run Fonts
				case "rFonts":
					this.parseFont(c, style);
					break;

				// TODO Revision Information for Run Properties
				case "rPrChange":
					break;

				// TODO Right To Left Text
				case "rtl":
					break;

				// TODO Shadow
				case "shadow":
					break;

				// Run Shading
				case "shd":
					style["background-color"] = xmlUtil.colorAttr(c, "fill", null, autos.shd);
					break;

				// Small Caps
				case "smallCaps":
					style["font-variant"] = xml.boolAttr(c, "val", true) ? "small-caps" : "none";
					break;

				// TODO Paragraph Mark Is Always Hidden
				case "specVanish":
					break;

				// Single Strikethrough
				case "strike":
					style["text-decoration"] = xml.boolAttr(c, "val", true) ? "line-through" : "none"
					break;

				// Non-Complex Script Font Size
				case "sz":
					// TODO 通过字符编码库或API来判断字符的编码范围，从而确定字符类型，字符类型决定字体大小
					style["font-size"] = style["min-height"] = xml.lengthAttr(c, "val", LengthUsage.FontSize);
					// style["font-size"] = xml.lengthAttr(c, "val", LengthUsage.FontSize);
					break;

				// Complex Script Font Size
				case "szCs":
					// TODO 通过字符编码库或API来判断字符的编码范围，从而确定字符类型，字符类型决定字体大小
					// style["font-size"] = style["min-height"] = xml.lengthAttr(c, "val", LengthUsage.FontSize);
					break;

				// Underline
				case "u":
					this.parseUnderline(c, style);
					break;

				// Hidden Text
				case "vanish":
					if (xml.boolAttr(c, "val", true))
						style["display"] = "none";
					break;

				// TODO	Subscript/Superscript Text
				case "vertAlign":
					// style.verticalAlign = values.valueOfVertAlign(c);
					break;

				// TODO Expanded/Compressed Text
				case "w":
					break;

				// TODO Web Hidden Text
				case "webHidden":
					break;

				case "jc":
					style["text-align"] = values.valueOfJc(c);
					break;

				case "textAlignment":
					style["vertical-align"] = values.valueOfTextAlignment(c);
					break;

				// 	TODO
				case "tcW":
					if (this.options.ignoreWidth) {
					}
					break;

				case "tblW":
					style["width"] = values.valueOfSize(c, "w");
					break;

				case "trHeight":
					this.parseTrHeight(c, style);
					break;

				case "ind":
				case "tblInd":
					this.parseIndentation(c, style);
					break;

				case "tblBorders":
					this.parseBorderProperties(c, childStyle || style);
					break;

				case "tblCellSpacing":
					style["border-spacing"] = values.valueOfMargin(c);
					style["border-collapse"] = "separate";
					break;

				case "pBdr":
					this.parseBorderProperties(c, style);
					break;

				case "tcBorders":
					this.parseBorderProperties(c, style);
					break;

				// TODO
				case "noWrap":
					//style["white-space"] = "nowrap";
					break;

				case "tblCellMar":
				case "tcMar":
					this.parseMarginProperties(c, childStyle || style);
					break;

				case "tblLayout":
					style["table-layout"] = values.valueOfTblLayout(c);
					break;

				case "vAlign":
					style["vertical-align"] = values.valueOfTextAlignment(c);
					break;

				case "wordWrap":
					if (xml.boolAttr(c, "val")) //TODO: test with examples
						style["overflow-wrap"] = "break-word";
					break;

				case "suppressAutoHyphens":
					style["hyphens"] = xml.boolAttr(c, "val", true) ? "none" : "auto";
					break;

				//ignore - tabs is parsed by other parser
				case "tabs":
				case "outlineLvl": //TODO
				case "contextualSpacing": //TODO
				case "tblStyleColBandSize": //TODO
				case "tblStyleRowBandSize": //TODO
				case "pageBreakBefore": //TODO - maybe ignore
				case "suppressLineNumbers": //TODO - maybe ignore
				case "keepLines": //TODO - maybe ignore
				case "keepNext": //TODO - maybe ignore
				case "widowControl": //TODO - maybe ignore

				default:
					if (this.options.debug) {
						console.warn(`DOCX:%c Unknown Property Element：${elem.localName}.${c.localName}`, 'color:green');
					}
					break;
			}
		});

		return style;
	}

	parseUnderline(node: Element, style: Record<string, string>) {
		let val = xml.attr(node, "val");

		if (val == null)
			return;

		switch (val) {
			case "dash":
			case "dashDotDotHeavy":
			case "dashDotHeavy":
			case "dashedHeavy":
			case "dashLong":
			case "dashLongHeavy":
			case "dotDash":
			case "dotDotDash":
				style["text-decoration"] = "underline dashed";
				break;

			case "dotted":
			case "dottedHeavy":
				style["text-decoration"] = "underline dotted";
				break;

			case "double":
				style["text-decoration"] = "underline double";
				break;

			case "single":
			case "thick":
				style["text-decoration"] = "underline";
				break;

			case "wave":
			case "wavyDouble":
			case "wavyHeavy":
				style["text-decoration"] = "underline wavy";
				break;

			case "words":
				style["text-decoration"] = "underline";
				break;

			case "none":
				style["text-decoration"] = "none";
				break;

			default:
				if (this.options.debug) {
					console.warn(`DOCX:%c Unknown Underline Property：${val}`, 'color:#f75607');
				}
		}

		let col = xmlUtil.colorAttr(node, "color");

		if (col) {
			style["text-decoration-color"] = col;
		}

	}

	// 转换Run字体，包含四种，ascii，eastAsia，ComplexScript，高 ANSI Font
	// TODO 通过字符编码库或API来判断字符的编码范围，从而确定字符类型，字符类型决定字体大小
	parseFont(node: Element, style: Record<string, string>) {
		// 字体
		let fonts = new Set();
		// ascii字体
		let ascii = xml.attr(node, "ascii");
		let ascii_theme = values.themeValue(node, "asciiTheme");
		fonts.add(ascii).add(ascii_theme);
		// eastAsia
		let east_Asia = xml.attr(node, "eastAsia");
		let east_Asia_theme = values.themeValue(node, "eastAsiaTheme");
		fonts.add(east_Asia).add(east_Asia_theme);
		// ComplexScript
		let complex_script = xml.attr(node, "cs");
		let complex_script_theme = values.themeValue(node, "cstheme");
		fonts.add(complex_script).add(complex_script_theme);
		// 高 ANSI Font
		let high_ansi = xml.attr(node, "hAnsi");
		let high_ansi_theme = values.themeValue(node, "hAnsiTheme");
		fonts.add(high_ansi).add(high_ansi_theme);
		// 去除重复字体，去除null
		let unique_fonts = [...fonts].filter(x => x);
		if (unique_fonts.length > 0) {
			// 合并成一个字体配置
			style["font-family"] = unique_fonts.join(', ');
		}

		// 字体提示：hint，拥有三种值：ComplexScript（cs）、Default（default）、EastAsia（eastAsia）
		style["_hint"] = xml.attr(node, "hint");
	}

	parseIndentation(node: Element, style: Record<string, string>) {
		let indentation: Record<string, any> = {};
		// 不同的单位将会产生不同的属性
		for (const attr of xml.attrs(node)) {
			switch (attr.localName) {
				case "end":
					indentation.end = xml.lengthAttr(node, "end");
					break;

				case "endChars":
					indentation.endCharacters = xml.lengthAttr(node, "endChars");
					break;

				case "firstLine":
					indentation.firstLine = xml.lengthAttr(node, "firstLine");
					break;

				case "firstLineChars":
					indentation.firstLineChars = xml.lengthAttr(node, "firstLineChars");
					break;

				case "hanging":
					indentation.hanging = xml.lengthAttr(node, "hanging");
					break;

				case "hangingChars":
					indentation.hangingChars = xml.lengthAttr(node, "hangingChars");
					break;

				case "left":
					indentation.left = xml.lengthAttr(node, "left");
					break;

				case "leftChars":
					indentation.leftChars = xml.lengthAttr(node, "leftChars");
					break;

				case "right":
					indentation.right = xml.lengthAttr(node, "right");
					break;

				case "rightChars":
					indentation.rightChars = xml.lengthAttr(node, "rightChars");
					break;

				case "start":
					indentation.start = xml.lengthAttr(node, "start");
					break;

				case "startChars":
					indentation.startChars = xml.lengthAttr(node, "startChars");
					break;

				default:
					if (this.options.debug) {
						console.warn(`DOCX:%c Unknown Indentation Property：${attr.localName}`, 'color:#f75607');
					}
			}
		}
		// TODO 处理文本缩进
		if (indentation.firstLine) style["text-indent"] = indentation.firstLine;
		if (indentation.hanging) style["text-indent"] = `-${indentation.hanging}`;
		// 段落缩进，通过padding实现
		if (indentation.left || indentation.start) style["padding-left"] = indentation.left || indentation.start;
		if (indentation.right || indentation.end) style["padding-right"] = indentation.right || indentation.end;
	}

	// the additional amount of character pitch to the contents of a run
	parseSpacing(node: Element, run: WmlRun) {
		for (const attr of xml.attrs(node)) {
			switch (attr.localName) {
				// Character Spacing Adjustment
				case "val":
					run.cssStyle["margin-bottom"] = xml.lengthAttr(node, "val");
					break;

				default:
					if (this.options.debug) {
						console.warn(`DOCX:%c Unknown Spacing Property：${attr.localName}`, 'color:#f75607');
					}
			}
		}
	}

	parseMarginProperties(node: Element, output: Record<string, string>) {
		xmlUtil.foreach(node, c => {
			switch (c.localName) {
				case "left":
					output["padding-left"] = values.valueOfMargin(c);
					break;

				case "right":
					output["padding-right"] = values.valueOfMargin(c);
					break;

				case "top":
					output["padding-top"] = values.valueOfMargin(c);
					break;

				case "bottom":
					output["padding-bottom"] = values.valueOfMargin(c);
					break;

				default:
					if (this.options.debug) {
						console.warn(`DOCX:%c Unknown Margin Property：${c.localName}`, 'color:#f75607');
					}
			}
		});
	}

	parseTrHeight(node: Element, output: Record<string, string>) {
		switch (xml.attr(node, "hRule")) {
			case "exact":
				output["height"] = xml.lengthAttr(node, "val");
				break;

			case "atLeast":
			default:
				output["height"] = xml.lengthAttr(node, "val");
				// min-height doesn't work for tr
				//output["min-height"] = xml.sizeAttr(node, "val");
				break;
		}
	}

	parseBorderProperties(node: Element, output: Record<string, string>) {
		xmlUtil.foreach(node, c => {
			switch (c.localName) {
				case "start":
				case "left":
					output["border-left"] = values.valueOfBorder(c);
					break;

				case "end":
				case "right":
					output["border-right"] = values.valueOfBorder(c);
					break;

				case "top":
					output["border-top"] = values.valueOfBorder(c);
					break;

				case "bottom":
					output["border-bottom"] = values.valueOfBorder(c);
					break;

				default:
					if (this.options.debug) {
						console.warn(`DOCX:%c Unknown Border Property：${c.localName}`, 'color:#f75607');
					}
			}
		});
	}
}

