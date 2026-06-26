import { BreakType, DomType, IDomNumbering, NumberingPicBullet, OpenXmlElement, WmlBreak, WmlCharacter, WmlHyperlink, WmlImage, WmlLastRenderedPageBreak, WmlNoteReference, WmlSymbol, WmlTable, WmlTableCell, WmlTableColumn, WmlTableRow, WmlText } from './document/dom';
import { DocumentElement } from './document/document';
import { parseParagraphProperties, parseParagraphProperty, WmlParagraph } from './document/paragraph';
import { parseSectionProperties, SectionProperties } from './document/section';
import xml from './parser/xml-parser';
import { parseRunProperties, WmlRun } from './document/run';
import { parseBookmarkEnd, parseBookmarkStart } from './document/bookmarks';
import { IDomStyle, Ruleset } from './document/style';
import { WmlFieldChar, WmlFieldSimple } from './document/fields';
import { LengthUsage, LengthUsageType } from './document/common';
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
import {
	parseVmlPicture as parseVmlPictureFn,
	checkAlternateContent as checkAlternateContentFn,
	parseDrawing as parseDrawingFn,
	parseDrawingWrapper as parseDrawingWrapperFn,
	parsePolygon as parsePolygonFn,
	DrawingParserCallbacks,
} from './parser/drawing-parser';
import {
	parseGraphic as parseGraphicFn,
	parseShape as parseShapeFn,
	parseShapeProperties as parseShapePropertiesFn,
	parseSolidFillColor as parseSolidFillColorFn,
	parseShapeLine as parseShapeLineFn,
	parsePicture as parsePictureFn,
	parseTransform2D as parseTransform2DFn,
	parseBlipFill as parseBlipFillFn,
	parseBlip as parseBlipFn,
} from './parser/shape-parser';
import {
	parseMathElement as parseMathElementFn,
	parseMathProperties as parseMathPropertiesFn,
	MathParserCallbacks,
} from './parser/math-parser';


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
		const callbacks: MathParserCallbacks = { parseRun: n => this.parseRun(n) };
		return parseMathElementFn(elem, this.options, callbacks);
	}

	parseMathProperties(elem: Element): Record<string, any> {
		return parseMathPropertiesFn(elem, this.options);
	}

	parseVmlPicture(elem: Element): OpenXmlElement {
		const callbacks: DrawingParserCallbacks = { parseBodyElements: n => this.parseBodyElements(n) };
		return parseVmlPictureFn(elem, callbacks);
	}

	// 检测备选内容
	checkAlternateContent(elem: Element): Element {
		return checkAlternateContentFn(elem);
	}

	parseDrawing(node: Element): OpenXmlElement {
		const callbacks: DrawingParserCallbacks = { parseBodyElements: n => this.parseBodyElements(n) };
		return parseDrawingFn(node, this.options, callbacks);
	}

	parseDrawingWrapper(node: Element): OpenXmlElement {
		const callbacks: DrawingParserCallbacks = { parseBodyElements: n => this.parseBodyElements(n) };
		return parseDrawingWrapperFn(node, this.options, callbacks);
	}

	parsePolygon(node: Element, target: OpenXmlElement): void {
		parsePolygonFn(node, target);
	}

	parseGraphic(elem: Element): OpenXmlElement {
		const callbacks: DrawingParserCallbacks = { parseBodyElements: n => this.parseBodyElements(n) };
		return parseGraphicFn(elem, this.options, callbacks);
	}

	parseShape(node: Element): OpenXmlElement {
		const callbacks: DrawingParserCallbacks = { parseBodyElements: n => this.parseBodyElements(n) };
		return parseShapeFn(node, this.options, callbacks);
	}

	parseShapeProperties(node: Element, target: OpenXmlElement): void {
		parseShapePropertiesFn(node, target, this.options);
	}

	parseSolidFillColor(node: Element): string | null {
		return parseSolidFillColorFn(node);
	}

	parseShapeLine(node: Element): { width?: string; color?: string } {
		return parseShapeLineFn(node);
	}

	parsePicture(elem: Element): WmlImage {
		return parsePictureFn(elem, this.options);
	}

	parseTransform2D(node: Element, target: OpenXmlElement): void {
		parseTransform2DFn(node, target, this.options);
	}

	parseBlipFill(node: Element, target: WmlImage): void {
		parseBlipFillFn(node, target, this.options);
	}

	parseBlip(node: Element, target: OpenXmlElement): void {
		parseBlipFn(node, target, this.options);
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

