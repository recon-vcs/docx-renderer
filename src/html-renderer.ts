import { WordDocument } from './word-document';
import {
	DomType,
	WmlImage,
	IDomNumbering,
	OpenXmlElement,
	WmlBreak,
	WmlDrawing,
	WmlHyperlink,
	WmlNoteReference,
	WmlSymbol,
	WmlTable,
	WmlTableCell,
	WmlTableColumn,
	WmlTableRow,
	WmlText,
	WrapType,
} from './document/dom';
import { CommonProperties } from './document/common';
import { Options } from './docx-preview';
import { DocumentElement } from './document/document';
import { WmlParagraph } from './document/paragraph';
import * as _ from 'lodash-es';
import { asArray, escapeClassName, uuid } from './utils';
import { computePointToPixelRatio, updateTabStop } from './javascript';
import { FontTablePart } from './font-table/font-table';
import { FooterHeaderReference, SectionProperties, SectionType } from './document/section';
import { Page, PageProps } from './document/page';
import { RunProperties, WmlRun } from './document/run';
import { WmlBookmarkStart } from './document/bookmarks';
import { IDomStyle } from './document/style';
import { WmlBaseNote, WmlFootnote } from './notes/elements';
import { ThemePart } from './theme/theme-part';
import { BaseHeaderFooterPart } from './header-footer/parts';
import { Part } from './common/part';
import { VmlElement } from './vml/vml';
import Konva from 'konva';
import type { Stage } from 'konva/lib/Stage';
import type { Layer } from 'konva/lib/Layer';
import type { Group } from 'konva/lib/Group';

const ns = {
	html: 'http://www.w3.org/1999/xhtml',
	svg: 'http://www.w3.org/2000/svg',
	mathML: 'http://www.w3.org/1998/Math/MathML',
};

interface CellPos {
	col: number;
	row: number;
}

type CellVerticalMergeType = Record<number, HTMLTableCellElement>;

interface Node_DOM extends Node, Comment, CharacterData {
	dataset: DOMStringMap;
}

enum Overflow {
	TRUE = 'true',
	FALSE = 'false',
	UNKNOWN = 'undetected',
}

// HTML渲染器

export class HtmlRenderer {

	className: string = "docx";
	rootSelector: string;
	document: WordDocument;
	options: Options;
	styleMap: Record<string, IDomStyle> = {};
	currentPart: Part = null;
	wrapper: HTMLElement;

	// 当前操作的Page
	currentPage: Page;

	tableVerticalMerges: CellVerticalMergeType[] = [];
	currentVerticalMerge: CellVerticalMergeType = null;
	tableCellPositions: CellPos[] = [];
	currentCellPosition: CellPos = null;

	footnoteMap: Record<string, WmlFootnote> = {};
	endnoteMap: Record<string, WmlFootnote> = {};
	currentFootnoteIds: string[];
	currentEndnoteIds: string[] = [];
	usedHederFooterParts: any[] = [];

	defaultTabSize: string;
	// 当前制表位
	currentTabs: any[] = [];
	tabsTimeout: any = 0;

	// Konva框架--stage元素
	konva_stage: Stage;
	// Konva框架--layer元素
	konva_layer: Layer;

	/**
	 * Object对象 => HTML标签
	 *
	 * @param document word文档Object对象
	 * @param bodyContainer HTML生成容器
	 * @param styleContainer CSS样式生成容器
	 * @param options 渲染配置选项
	 */

	render(document: WordDocument, bodyContainer: HTMLElement, styleContainer: HTMLElement = null, options: Options) {
		this.document = document;
		this.options = options;
		// class类前缀
		this.className = options.className;
		// 根元素
		this.rootSelector = options.inWrapper ? `.${this.className}-wrapper` : ':root';
		// 文档CSS样式
		this.styleMap = null;
		// 主体容器
		this.wrapper = bodyContainer;
		// styleContainer== null，styleContainer = bodyContainer
		styleContainer = styleContainer || bodyContainer;

		// CSS样式生成容器，清空所有CSS样式
		removeAllElements(styleContainer);
		// HTML生成容器，清空所有HTML元素
		removeAllElements(bodyContainer);

		// 添加注释
		appendComment(styleContainer, "docxjs library predefined styles");
		// 添加默认CSS样式
		styleContainer.appendChild(this.renderDefaultStyle());

		// 主题CSS样式
		if (document.themePart) {
			appendComment(styleContainer, "docxjs document theme values");
			this.renderTheme(document.themePart, styleContainer);
		}
		// 文档默认CSS样式，包含表格、列表、段落、字体，样式存在继承顺序
		if (document.stylesPart != null) {
			this.styleMap = this.processStyles(document.stylesPart.styles);

			appendComment(styleContainer, "docxjs document styles");
			styleContainer.appendChild(this.renderStyles(document.stylesPart.styles));
		}
		// 多级列表样式
		if (document.numberingPart) {
			this.processNumberings(document.numberingPart.domNumberings);

			appendComment(styleContainer, "docxjs document numbering styles");
			styleContainer.appendChild(this.renderNumbering(document.numberingPart.domNumberings, styleContainer));
			//styleContainer.appendChild(this.renderNumbering2(document.numberingPart, styleContainer));
		}
		// 字体列表CSS样式
		if (!options.ignoreFonts && document.fontTablePart) {
			this.renderFontTable(document.fontTablePart, styleContainer);
		}
		// 生成脚注部分的Map
		if (document.footnotesPart) {
			this.footnoteMap = _.keyBy(document.footnotesPart.rootElement.children, 'id');
		}
		// 生成尾注部分的Map
		if (document.endnotesPart) {
			this.endnoteMap = _.keyBy(document.endnotesPart.rootElement.children, 'id');
		}
		// 文档设置
		if (document.settingsPart) {
			this.defaultTabSize = document.settingsPart.settings?.defaultTabStop;
		}
		// 主文档--内容
		let pageElements = this.renderPages(document.documentPart.body);
		if (this.options.inWrapper) {
			bodyContainer.appendChild(this.renderWrapper(pageElements));
		} else {
			appendChildren(bodyContainer, pageElements);
		}

		// 刷新制表符
		this.refreshTabStops();
	}

	// 渲染默认样式
	renderDefaultStyle() {
		let c = this.className;
		let styleText = `
			.${c}-wrapper { background: gray; padding: 30px; padding-bottom: 0px; display: flex; flex-flow: column; align-items: center; } 
			.${c}-wrapper>section.${c} { background: white; box-shadow: 0 0 10px rgba(0, 0, 0, 0.5); margin-bottom: 30px; }
			.${c} { color: black; hyphens: auto; text-underline-position: from-font; }
			section.${c} { box-sizing: border-box; display: flex; flex-flow: column nowrap; position: relative; overflow: hidden; }
            section.${c}>header { position: absolute; top: 0; z-index: 1; display: flex; align-items: flex-end; }
			section.${c}>article { z-index: 1; }
			section.${c}>footer { position: absolute; bottom: 0; z-index: 1; }
			.${c} table { border-collapse: collapse; }
			.${c} table td, .${c} table th { vertical-align: top; }
			.${c} p { margin: 0pt; min-height: 1em; }
			.${c} span { white-space: pre-wrap; overflow-wrap: break-word; }
			.${c} math { vertical-align: middle; }
			.${c} .${c}-math-paragraph math { display: inline-block; }
			.${c} a { color: inherit; text-decoration: inherit; }
			.${c} img, ${c} svg { vertical-align: baseline; }
			.${c} .clearfix::after { content: ""; display: block; line-height: 0; clear: both; }
		`;

		return createStyleElement(styleText);
	}

	// 文档CSS主题样式
	renderTheme(themePart: ThemePart, styleContainer: HTMLElement) {
		const variables = {};
		const fontScheme = themePart.theme?.fontScheme;

		if (fontScheme) {
			if (fontScheme.majorFont) {
				variables['--docx-majorHAnsi-font'] = fontScheme.majorFont.latinTypeface;
			}

			if (fontScheme.minorFont) {
				variables['--docx-minorHAnsi-font'] = fontScheme.minorFont.latinTypeface;
			}
		}

		const colorScheme = themePart.theme?.colorScheme;

		if (colorScheme) {
			for (let [k, v] of Object.entries(colorScheme.colors)) {
				variables[`--docx-${k}-color`] = `#${v}`;
			}
		}

		const cssText = this.styleToString(`.${this.className}`, variables);
		styleContainer.appendChild(createStyleElement(cssText));
	}

	// 计算className，小写，默认前缀："docx_"
	processStyleName(className: string): string {
		return className ? `${this.className}_${escapeClassName(className)}` : this.className;
	}

	// 处理样式继承
	processStyles(styles: IDomStyle[]) {
		//
		const stylesMap = _.keyBy(styles.filter(x => x.id != null), 'id');
		// 遍历base_on关系,合并样式
		for (const style of styles.filter(x => x.basedOn)) {
			let baseStyle = stylesMap[style.basedOn];

			if (baseStyle) {
				// 深度合并
				style.paragraphProps = _.merge(style.paragraphProps, baseStyle.paragraphProps);
				style.runProps = _.merge(style.runProps, baseStyle.runProps);

				for (const baseValues of baseStyle.rulesets) {
					const styleValues = style.rulesets.find(x => x.target == baseValues.target);

					if (styleValues) {
						this.copyStyleProperties(baseValues.declarations, styleValues.declarations);
					} else {
						style.rulesets.push({ ...baseValues, declarations: { ...baseValues.declarations } });
					}
				}
			} else if (this.options.debug) {
				console.warn(`Can't find base style ${style.basedOn}`);
			}
		}

		for (let style of styles) {
			style.cssName = this.processStyleName(style.id);
		}

		return stylesMap;
	}

	renderStyles(styles: IDomStyle[]): HTMLElement {
		let styleText = "";
		const stylesMap = this.styleMap;
		const defaultStyles = _.keyBy(styles.filter(s => s.isDefault), 'target');

		for (const style of styles) {
			let subStyles = style.rulesets;

			if (style.linked) {
				let linkedStyle = style.linked && stylesMap[style.linked];

				if (linkedStyle) {
					subStyles = subStyles.concat(linkedStyle.rulesets);
				} else if (this.options.debug) {
					console.warn(`Can't find linked style ${style.linked}`);
				}
			}

			for (const subStyle of subStyles) {
				//TODO temporary disable modificators until test it well
				let selector = `${style.type ?? ''}.${style.cssName}`; //${subStyle.mod ?? ''}

				if (style.type != subStyle.target) {
					selector += ` ${subStyle.target}`;
				}

				if (defaultStyles[style.type] == style) {
					selector = `.${this.className} ${style.type}, ` + selector;
				}

				styleText += this.styleToString(selector, subStyle.declarations);
			}
		}

		return createStyleElement(styleText);
	}

	processNumberings(numberings: IDomNumbering[]) {
		for (let num of numberings.filter(n => n.pStyleName)) {
			const style = this.findStyle(num.pStyleName);

			if (style?.paragraphProps?.numbering) {
				style.paragraphProps.numbering.level = num.level;
			}
		}
	}

	renderNumbering(numberings: IDomNumbering[], styleContainer: HTMLElement) {
		let styleText = "";
		let resetCounters = [];

		for (let num of numberings) {
			let selector = `p.${this.numberingClass(num.id, num.level)}`;
			let listStyleType = "none";

			if (num.bullet) {
				let valiable = `--${this.className}-${num.bullet.src}`.toLowerCase();

				styleText += this.styleToString(`${selector}:before`, {
					"content": "' '",
					"display": "inline-block",
					"background": `var(${valiable})`
				}, num.bullet.style);

				this.document.loadNumberingImage(num.bullet.src).then(data => {
					let text = `${this.rootSelector} { ${valiable}: url(${data}) }`;
					styleContainer.appendChild(createStyleElement(text));
				});
			} else if (num.levelText) {
				let counter = this.numberingCounter(num.id, num.level);
				const counterReset = counter + " " + (num.start - 1);
				if (num.level > 0) {
					styleText += this.styleToString(`p.${this.numberingClass(num.id, num.level - 1)}`, {
						"counter-reset": counterReset
					});
				}
				// reset all level counters with start value
				resetCounters.push(counterReset);

				styleText += this.styleToString(`${selector}:before`, {
					"content": this.levelTextToContent(num.levelText, num.suff, num.id, this.numFormatToCssValue(num.format)),
					"counter-increment": counter,
					...num.rStyle,
				});
			} else {
				listStyleType = this.numFormatToCssValue(num.format);
			}

			styleText += this.styleToString(selector, {
				"display": "list-item",
				"list-style-position": "inside",
				"list-style-type": listStyleType,
				...num.pStyle
			});
		}

		if (resetCounters.length > 0) {
			styleText += this.styleToString(this.rootSelector, {
				"counter-reset": resetCounters.join(" ")
			});
		}

		return createStyleElement(styleText);
	}

	numberingClass(id: string, lvl: number) {
		return `${this.className}-num-${id}-${lvl}`;
	}

	styleToString(selectors: string, values: Record<string, string>, cssText: string = null) {
		let result = `${selectors} {\r\n`;

		for (const key in values) {
			if (key.startsWith('$')) {
				continue;
			}

			result += `  ${key}: ${values[key]};\r\n`;
		}

		if (cssText) {
			result += cssText;
		}

		return result + "}\r\n";
	}

	numberingCounter(id: string, lvl: number) {
		return `${this.className}-num-${id}-${lvl}`;
	}

	levelTextToContent(text: string, suff: string, id: string, numformat: string) {
		const suffMap = {
			"tab": "\\9",
			"space": "\\a0",
		};

		let result = text.replace(/%\d*/g, s => {
			let lvl = parseInt(s.substring(1), 10) - 1;
			return `"counter(${this.numberingCounter(id, lvl)}, ${numformat})"`;
		});

		return `"${result}${suffMap[suff] ?? ""}"`;
	}

	numFormatToCssValue(format: string) {
		let mapping = {
			none: "none",
			bullet: "disc",
			decimal: "decimal",
			lowerLetter: "lower-alpha",
			upperLetter: "upper-alpha",
			lowerRoman: "lower-roman",
			upperRoman: "upper-roman",
			decimalZero: "decimal-leading-zero", // 01,02,03,...
			// ordinal: "", // 1st, 2nd, 3rd,...
			// ordinalText: "", //First, Second, Third, ...
			// cardinalText: "", //One,Two Three,...
			// numberInDash: "", //-1-,-2-,-3-, ...
			// hex: "upper-hexadecimal",
			aiueo: "katakana",
			aiueoFullWidth: "katakana",
			chineseCounting: "simp-chinese-informal",
			chineseCountingThousand: "simp-chinese-informal",
			chineseLegalSimplified: "simp-chinese-formal", // 中文大写
			chosung: "hangul-consonant",
			ideographDigital: "cjk-ideographic",
			ideographTraditional: "cjk-heavenly-stem", // 十天干
			ideographLegalTraditional: "trad-chinese-formal",
			ideographZodiac: "cjk-earthly-branch", // 十二地支
			iroha: "katakana-iroha",
			irohaFullWidth: "katakana-iroha",
			japaneseCounting: "japanese-informal",
			japaneseDigitalTenThousand: "cjk-decimal",
			japaneseLegal: "japanese-formal",
			thaiNumbers: "thai",
			koreanCounting: "korean-hangul-formal",
			koreanDigital: "korean-hangul-formal",
			koreanDigital2: "korean-hanja-informal",
			hebrew1: "hebrew",
			hebrew2: "hebrew",
			hindiNumbers: "devanagari",
			ganada: "hangul",
			taiwaneseCounting: "cjk-ideographic",
			taiwaneseCountingThousand: "cjk-ideographic",
			taiwaneseDigital: "cjk-decimal",
		};

		return mapping[format] ?? format;
	}

	// renderNumbering2(numberingPart: NumberingPartProperties, container: HTMLElement): HTMLElement {
	// 	let css = "";
	// 	const numberingMap = keyBy(numberingPart.abstractNumberings, x => x.id);
	// 	const bulletMap = keyBy(numberingPart.bulletPictures, x => x.id);
	// 	const topCounters = [];
	//
	// 	for (let num of numberingPart.numberings) {
	// 		const absNum = numberingMap[num.abstractId];
	//
	// 		for (let lvl of absNum.levels) {
	// 			const className = this.numberingClass(num.id, lvl.level);
	// 			let listStyleType = "none";
	//
	// 			if (lvl.text && lvl.format == 'decimal') {
	// 				const counter = this.numberingCounter(num.id, lvl.level);
	//
	// 				if (lvl.level > 0) {
	// 					css += this.styleToString(`p.${this.numberingClass(num.id, lvl.level - 1)}`, {
	// 						"counter-reset": counter
	// 					});
	// 				} else {
	// 					topCounters.push(counter);
	// 				}
	//
	// 				css += this.styleToString(`p.${className}:before`, {
	// 					"content": this.levelTextToContent(lvl.text, num.id),
	// 					"counter-increment": counter
	// 				});
	// 			} else if (lvl.bulletPictureId) {
	// 				let pict = bulletMap[lvl.bulletPictureId];
	// 				let variable = `--${this.className}-${pict.referenceId}`.toLowerCase();
	//
	// 				css += this.styleToString(`p.${className}:before`, {
	// 					"content": "' '",
	// 					"display": "inline-block",
	// 					"background": `var(${variable})`
	// 				}, pict.style);
	//
	// 				this.document.loadNumberingImage(pict.referenceId).then(data => {
	// 					var text = `.${this.className}-wrapper { ${variable}: url(${data}) }`;
	// 					container.appendChild(createStyleElement(text));
	// 				});
	// 			} else {
	// 				listStyleType = this.numFormatToCssValue(lvl.format);
	// 			}
	//
	// 			css += this.styleToString(`p.${className}`, {
	// 				"display": "list-item",
	// 				"list-style-position": "inside",
	// 				"list-style-type": listStyleType,
	// 				//TODO
	// 				//...num.style
	// 			});
	// 		}
	// 	}
	//
	// 	if (topCounters.length > 0) {
	// 		css += this.styleToString(`.${this.className}-wrapper`, {
	// 			"counter-reset": topCounters.join(" ")
	// 		});
	// 	}
	//
	// 	return createStyleElement(css);
	// }

	// 字体列表CSS样式
	renderFontTable(fontsPart: FontTablePart, styleContainer: HTMLElement) {
		for (let f of fontsPart.fonts) {
			for (let ref of f.embedFontRefs) {
				this.document.loadFont(ref.id, ref.key).then(fontData => {
					const cssValues = {
						'font-family': f.name,
						'src': `url(${fontData})`
					};

					if (ref.type == "bold" || ref.type == "boldItalic") {
						cssValues['font-weight'] = 'bold';
					}

					if (ref.type == "italic" || ref.type == "boldItalic") {
						cssValues['font-style'] = 'italic';
					}

					appendComment(styleContainer, `docxjs ${f.name} font`);
					const cssText = this.styleToString("@font-face", cssValues);
					styleContainer.appendChild(createStyleElement(cssText));
					this.refreshTabStops();
				});
			}
		}
	}

	// 生成父级容器
	renderWrapper(children: HTMLElement[]) {
		return createElement("div", { className: `${this.className}-wrapper` }, children);
	}

	// 复制CSS样式
	copyStyleProperties(input: Record<string, string>, output: Record<string, string>, attrs: string[] = null): Record<string, string> {
		if (!input) {
			return output;
		}
		if (output == null) {
			output = {};
		}
		if (attrs == null) {
			attrs = Object.getOwnPropertyNames(input);
		}

		for (let key of attrs) {
			if (input.hasOwnProperty(key) && !output.hasOwnProperty(key))
				output[key] = input[key];
		}

		return output;
	}

	// 递归明确元素parent父级关系
	processElement(element: OpenXmlElement) {
		if (element.children) {
			for (let e of element.children) {
				e.parent = element;
				// 判断类型
				if (e.type == DomType.Table) {
					// 渲染表格
					this.processTable(e);
				} else {
					// 递归渲染
					this.processElement(e);
				}
			}
		}
	}

	// 表格style样式
	processTable(table: WmlTable) {
		for (let r of table.children) {
			for (let c of r.children) {
				c.cssStyle = this.copyStyleProperties(table.cellStyle, c.cssStyle, [
					"border-left", "border-right", "border-top", "border-bottom",
					"padding-left", "padding-right", "padding-top", "padding-bottom"
				]);

				this.processElement(c);
			}
		}
	}

	/*
	 * section与page概念区别
	 * 章节(section)是根据内容的逻辑结构和组织来划分的，不同章节设置独立的格式。
	 * 页面是文档实际呈现的物理单位，而章节则是逻辑上的分割点。
	 */

	// 根据分页符拆分页面
	splitPage(elements: OpenXmlElement[]): Page[] {
		// 当前操作page，elements数组包含子元素
		let current_page: Page = new Page({} as PageProps);
		// 切分出的所有pages
		const pages: Page[] = [current_page];

		for (const elem of elements) {
			// 标记顶层元素的层级level
			elem.level = 1;
			// 添加elem进入当前操作page
			current_page.children.push(elem);

			/* 段落基本结构：paragraph => run => text... */
			if (elem.type == DomType.Paragraph) {
				const p = elem as WmlParagraph;
				// 节属性，代表分节符
				const sectProps: SectionProperties = p.props.sectionProperties;
				// 节属性生成唯一uuid，每一个节中page均是同一个uuid，代表属于同一个节
				if (sectProps) {
					sectProps.sectionId = uuid();
				}
				// 查找内置默认段落样式
				const default_paragraph_style = this.findStyle(p.styleName);

				// 检测段落内置样式是否存在段前分页符
				if (default_paragraph_style?.paragraphProps?.pageBreakBefore) {
					// 标记当前page已拆分
					current_page.isSplit = true;
					// 保存当前page的sectionProps
					current_page.sectProps = sectProps;
					// 重置新的page
					current_page = new Page({} as PageProps);
					// 添加新page
					pages.push(current_page);
				}

				// 段落部分Break索引
				let pBreakIndex = -1;
				// Run部分Break索引
				let rBreakIndex = -1;

				// 查询段落中Break索引
				if (p.children) {
					// 计算段落Break索引
					pBreakIndex = p.children.findIndex(r => {
						// 计算Run Break索引
						rBreakIndex = r.children?.findIndex((t: OpenXmlElement) => {
							// 如果不是分页符、换行符、分栏符
							if (t.type !== DomType.Break && t.type !== DomType.LastRenderedPageBreak) {
								return false;
							}
							// 默认忽略lastRenderedPageBreak，
							if (t.type === DomType.LastRenderedPageBreak) {
								// 判断前一个p段落，
								// 如果含有分页符、分节符，那它们一定位于上一个page，数组为空；
								// 如果前一个段落是普通段落，数组长度大于0，则代表文字过多超过一页，需要自动分页
								return (current_page.children.length > 2 || !this.options.ignoreLastRenderedPageBreak);
							}
							// 分页符
							if ((t as WmlBreak).break === "page") {
								return true;
							}
						});
						rBreakIndex = rBreakIndex ?? -1;
						return rBreakIndex != -1;
					});
				}
				// 段落Break索引
				if (pBreakIndex != -1) {
					// 一般情况下，标记当前page：已拆分
					current_page.isSplit = true;
					// 检测分页符之前的所有元素是否存在表格
					const exist_table: boolean = current_page.children.some(
						elem => elem.type === DomType.Table
					);
					// 存在表格
					if (exist_table) {
						// 表格可能需要计算之后拆分，标记当前page：未拆分
						current_page.isSplit = false;
					}
					// 检测分页符之前的所有元素是否存在目录
					let exist_TOC: boolean = current_page.children.some((paragraph) => {
						return paragraph.children.some((elem) => {
							if (elem.type === DomType.Hyperlink) {
								return (elem as WmlHyperlink)?.href?.includes('Toc')
							}
							return false;
						});
					});
					// 	存在目录
					if (exist_TOC) {
						// 目录可能需要计算之后拆分，标记当前page：未拆分
						current_page.isSplit = false;
					}
				}
				/*
				 *
				 * 分页有两种情况：
				 * 1、段落中存在节属性sectProps，且类型不是continuous/nextColumn
				 * 2、段落存在Break索引
				 *
				 */
				if (pBreakIndex != -1 || (sectProps && sectProps.type != SectionType.Continuous && sectProps.type != SectionType.NextColumn)) {
					// 保存当前page的pageProps
					current_page.sectProps = sectProps;
					// 重置新的page
					current_page = new Page({} as PageProps);
					// 添加新page
					pages.push(current_page);
				}

				// 根据段落Break索引，拆分Run部分
				if (pBreakIndex != -1) {
					// 即将拆分的Run部分
					let breakRun = p.children[pBreakIndex];
					// 是否需要拆分Run
					let is_split = rBreakIndex < breakRun.children.length - 1;

					if (pBreakIndex < p.children.length - 1 || is_split) {
						// 原始的Run
						let origin_run = p.children;
						// 切出Break索引后面的Run，创建新段落
						const new_paragraph: WmlParagraph = {
							...p,
							children: origin_run.slice(pBreakIndex),
						};
						// 保存Break索引前面的Run
						p.children = origin_run.slice(0, pBreakIndex);
						// 添加新段落
						current_page.children.push(new_paragraph);

						if (is_split) {
							// Run下面原始的元素
							const origin_elements = breakRun.children;
							// 切出Run Break索引前面的元素，创建新Run
							const newRun = {
								...breakRun,
								children: origin_elements.slice(0, rBreakIndex),
							};
							// 将新Run放入上一个page的段落
							p.children.push(newRun);
							// 切出Run Break索引后面的元素
							breakRun.children = origin_elements.slice(rBreakIndex);
						}
					}
				}
			}

			// elem元素是表格，需要渲染过程中拆分page
			if (elem.type === DomType.Table) {
				// 标记当前page：未拆分
				current_page.isSplit = false;
			}
		}
		// 一个节可能分好几个页，但是节属性sectionProps存在当前节中最后一段对应的 paragraph 元素的子元素。即：[null,null,null,setPr];
		let currentSectProps = null;
		// 倒序给每一页填充sectionProps，方便后期页面渲染
		for (let i = pages.length - 1; i >= 0; i--) {
			if (pages[i].sectProps == null) {
				pages[i].sectProps = currentSectProps;
			} else {
				currentSectProps = pages[i].sectProps;
			}
		}
		return pages;
	}

	// 生成所有的页面Page
	renderPages(document: DocumentElement): HTMLElement[] {
		const result = [];
		// 生成页面parent父级关系
		this.processElement(document);
		// 根据options.breakPages，选择是否分页
		let pages: Page[];
		if (this.options.breakPages) {
			// 拆分页面
			pages = this.splitPage(document.children);
		} else {
			// 不分页则，只有一个page
			pages = [new Page({ sectProps: document.sectProps, children: document.children, } as PageProps)];
		}
		// 缓存分页的结果
		document.pages = pages;
		// 前一个节属性，判断分节符的第一个page
		let prevProps = null;
		// 遍历生成每一个page
		for (let i = 0, l = pages.length; i < l; i++) {
			this.currentFootnoteIds = [];
			const page: Page = pages[i];
			const { sectProps } = page;
			// sectionProps属性不存在，则使用文档级别props;
			let sectionProps = sectProps ?? document.sectProps;
			// 页码
			let pageIndex = result.length;
			// 是否本小节的第一个page
			let isFirstPage = prevProps != sectionProps;
			// TODO 是否最后一个page,此时分页未完成，计算并不准确，影响到尾注的渲染
			let isLastPage = i === l - 1;
			// 渲染单个page，有可能多个page
			let pageElements: HTMLElement[] = this.renderPage(page, sectionProps, document.cssStyle, pageIndex, isFirstPage, isLastPage);

			result.push(...pageElements);
			// 存储前一个节属性
			prevProps = sectProps;
		}

		return result;
	}

	// 生成单个page
	renderPage(page: Page, props: SectionProperties, sectionStyle: Record<string, string>, pageIndex: number, isFirstPage: boolean, isLastPage: boolean): HTMLElement[] {
		// 根据sectProps，创建page
		const pageElement = this.createPage(this.className, props);
		// 给page添加style样式
		this.renderStyleValues(sectionStyle, pageElement);
		// 渲染page页眉
		if (this.options.renderHeaders) {
			this.renderHeaderFooterRef(props.headerRefs, props, pageIndex, isFirstPage, pageElement);
		}
		// page主体内容
		let contentElement = createElement("article");
		// 根据options.breakPages，设置article的高度
		if (this.options.breakPages) {
			// 不分页则，拥有最小高度
			contentElement.style.minHeight = props.contentSize.height;
		}

		// 生成article内容
		this.renderElements(page.children, contentElement);
		// 放入page
		pageElement.appendChild(contentElement);

		// 渲染page脚注
		if (this.options.renderFootnotes) {
			this.renderNotes(this.currentFootnoteIds, this.footnoteMap, pageElement);
		}
		// 渲染page尾注，判断最后一页
		if (this.options.renderEndnotes && isLastPage) {
			this.renderNotes(this.currentEndnoteIds, this.endnoteMap, pageElement);
		}
		// 渲染page页脚
		if (this.options.renderFooters) {
			this.renderHeaderFooterRef(props.footerRefs, props, pageIndex, isFirstPage, pageElement);
		}
		return [pageElement]
	}

	// 创建Page
	createPage(className: string, props: SectionProperties) {
		let oPage = createElement("section", { className });

		if (props) {
			if (props.pageMargins) {
				oPage.style.paddingLeft = props.pageMargins.left;
				oPage.style.paddingRight = props.pageMargins.right;
				oPage.style.paddingTop = props.pageMargins.top;
				oPage.style.paddingBottom = props.pageMargins.bottom;
			}

			if (props.pageSize) {
				if (!this.options.ignoreWidth) {
					oPage.style.width = props.pageSize.width;
				}
				if (!this.options.ignoreHeight) {
					oPage.style.minHeight = props.pageSize.height;
				}
			}

			if (props.columns && props.columns.count) {
				oPage.style.columnCount = `${props.columns.count}`;
				oPage.style.columnGap = props.columns.space;

				if (props.columns.separator) {
					oPage.style.columnRule = "1px solid black";
				}
			}
		}

		return oPage;
	}

	// TODO 分页不准确，页脚页码混乱
	// 渲染页眉/页脚的Ref
	renderHeaderFooterRef(refs: FooterHeaderReference[], props: SectionProperties, page: number, isFirstPage: boolean, parent: HTMLElement) {
		if (!refs) return;
		// 查找奇数偶数的ref指向
		let ref = (props.titlePage && isFirstPage ? refs.find(x => x.type == "first") : null)
			?? (page % 2 == 1 ? refs.find(x => x.type == "even") : null)
			?? refs.find(x => x.type == "default");

		// 查找ref对应的part部分
		let part = ref && this.document.findPartByRelId(ref.id, this.document.documentPart) as BaseHeaderFooterPart;

		if (part) {
			this.currentPart = part;
			if (!this.usedHederFooterParts.includes(part.path)) {
				this.processElement(part.rootElement);
				this.usedHederFooterParts.push(part.path);
			}
			// 根据页眉页脚，设置CSS
			switch (part.rootElement.type) {
				case DomType.Header:
					part.rootElement.cssStyle = {
						left: props.pageMargins?.left,
						width: props.contentSize?.width,
						height: props.pageMargins?.top,
					}
					break;
				case DomType.Footer:
					part.rootElement.cssStyle = {
						left: props.pageMargins?.left,
						width: props.contentSize?.width,
						height: props.pageMargins?.bottom,
					}
					break;
				default:
					console.warn('set header/footer style error', part.rootElement.type);
					break;
			}

			this.renderElements([part.rootElement], parent);
			this.currentPart = null;
		}
	}

	// 渲染脚注/尾注
	renderNotes(noteIds: string[], notesMap: Record<string, WmlBaseNote>, parent: HTMLElement) {
		let notes = noteIds.map(id => notesMap[id]).filter(x => x);

		if (notes.length > 0) {
			let children = this.renderElements(notes);
			let result = createElement("ol", null, children);
			parent.appendChild(result);
		}
	}

	// 渲染多元素，
	renderElements(elems: OpenXmlElement[], parent?: HTMLElement): Node[] {
		if (elems == null) {
			return null;
		}

		let result: Node[] = [];

		for (let i = 0; i < elems.length; i++) {
			let element = this.renderElement(elems[i]);
			if (Array.isArray(element)) {
				result.push(...element);
			} else if (element) {
				result.push(element);
			}
		}

		if (parent) {
			appendChildren(parent, result);
		}

		return result;
	}

	// 渲染单个元素
	renderElement(elem: OpenXmlElement): Node | Node[] {
		switch (elem.type) {
			case DomType.Paragraph:
				return this.renderParagraph(elem as WmlParagraph);

			case DomType.BookmarkStart:
				return this.renderBookmarkStart(elem as WmlBookmarkStart);

			case DomType.BookmarkEnd:
				return null; //ignore bookmark end

			case DomType.Run:
				return this.renderRun(elem as WmlRun);

			case DomType.Table:
				return this.renderTable(elem);

			case DomType.Row:
				return this.renderTableRow(elem);

			case DomType.Cell:
				return this.renderTableCell(elem);

			case DomType.Hyperlink:
				return this.renderHyperlink(elem);

			case DomType.Drawing:
				return this.renderDrawing(elem);

			case DomType.Image:
				return this.renderImage(elem as WmlImage);

			case DomType.Text:
				return this.renderText(elem as WmlText);

			case DomType.DeletedText:
				return this.renderDeletedText(elem as WmlText);

			case DomType.Tab:
				return this.renderTab(elem);

			case DomType.Symbol:
				return this.renderSymbol(elem as WmlSymbol);

			case DomType.Break:
				return this.renderBreak(elem as WmlBreak);

			case DomType.Footer:
				return this.renderHeaderFooter(elem, "footer");

			case DomType.Header:
				return this.renderHeaderFooter(elem, "header");

			case DomType.Footnote:
			case DomType.Endnote:
				return this.renderContainer(elem, "li");

			case DomType.FootnoteReference:
				return this.renderFootnoteReference(elem as WmlNoteReference);

			case DomType.EndnoteReference:
				return this.renderEndnoteReference(elem as WmlNoteReference);

			case DomType.NoBreakHyphen:
				return createElement("wbr");

			case DomType.VmlPicture:
				return this.renderVmlPicture(elem);

			case DomType.VmlElement:
				return this.renderVmlElement(elem as VmlElement);

			case DomType.MmlMath:
				return this.renderContainerNS(elem, ns.mathML, "math", { xmlns: ns.mathML });

			case DomType.MmlMathParagraph:
				return this.renderMmlMathParagraph(elem);

			case DomType.MmlFraction:
				return this.renderContainerNS(elem, ns.mathML, "mfrac");

			case DomType.MmlBase:
				return this.renderContainerNS(elem, ns.mathML, elem.parent.type == DomType.MmlMatrixRow ? "mtd" : "mrow");

			case DomType.MmlNumerator:
			case DomType.MmlDenominator:
			case DomType.MmlFunction:
			case DomType.MmlLimit:
			case DomType.MmlBox:
				return this.renderContainerNS(elem, ns.mathML, "mrow");

			case DomType.MmlGroupChar:
				return this.renderMmlGroupChar(elem);

			case DomType.MmlLimitLower:
				return this.renderContainerNS(elem, ns.mathML, "munder");

			case DomType.MmlMatrix:
				return this.renderContainerNS(elem, ns.mathML, "mtable");

			case DomType.MmlMatrixRow:
				return this.renderContainerNS(elem, ns.mathML, "mtr");

			case DomType.MmlRadical:
				return this.renderMmlRadical(elem);


			case DomType.MmlSuperscript:
				return this.renderContainerNS(elem, ns.mathML, "msup");

			case DomType.MmlSubscript:
				return this.renderContainerNS(elem, ns.mathML, "msub");

			case DomType.MmlDegree:
			case DomType.MmlSuperArgument:
			case DomType.MmlSubArgument:
				return this.renderContainerNS(elem, ns.mathML, "mn");

			case DomType.MmlFunctionName:
				return this.renderContainerNS(elem, ns.mathML, "ms");

			case DomType.MmlDelimiter:
				return this.renderMmlDelimiter(elem);

			case DomType.MmlRun:
				return this.renderMmlRun(elem);

			case DomType.MmlNary:
				return this.renderMmlNary(elem);

			case DomType.MmlPreSubSuper:
				return this.renderMmlPreSubSuper(elem);

			case DomType.MmlBar:
				return this.renderMmlBar(elem);

			case DomType.MmlEquationArray:
				return this.renderMllList(elem);

			case DomType.Inserted:
				return this.renderInserted(elem);

			case DomType.Deleted:
				return this.renderDeleted(elem);

		}

		return null;
	}

	// 判断是否存在分页元素
	isPageBreakElement(elem: OpenXmlElement): boolean {
		// 分页符、换行符、分栏符
		if (elem.type !== DomType.Break && elem.type !== DomType.LastRenderedPageBreak) {
			return false;
		}
		// 默认以lastRenderedPageBreak作为分页依据
		if (elem.type === DomType.LastRenderedPageBreak) {
			return !this.options.ignoreLastRenderedPageBreak;
		}
		// 分页符
		if ((elem as WmlBreak).break === "page") {
			return true;
		}
	}

	renderChildren(elem: OpenXmlElement, parent?: HTMLElement): Node[] {
		return this.renderElements(elem.children, parent);
	}

	renderContainer(elem: OpenXmlElement, tagName: keyof HTMLElementTagNameMap, props?: Record<string, any>) {
		return createElement(tagName, props, this.renderChildren(elem));
	}

	renderContainerNS(elem: OpenXmlElement, ns: string, tagName: string, props?: Record<string, any>) {
		return createElementNS(ns, tagName, props, this.renderChildren(elem));
	}

	renderMmlMathParagraph(elem: OpenXmlElement) {
		const result = createElement("span");

		result.classList.add(`${this.className}-math-paragraph`);
		result.style.display = "block";
		result.style.textAlign = this.mathJustificationToTextAlign(elem.props?.justification);
		result.style.textIndent = "0px";

		this.renderChildren(elem, result);

		return result;
	}

	mathJustificationToTextAlign(justification?: string) {
		switch (justification) {
			case "left":
				return "left";
			case "right":
				return "right";
			case "center":
			case "centerGroup":
			default:
				return "center";
		}
	}

	renderParagraph(elem: WmlParagraph) {
		let result = createElement("p");

		const style = this.findStyle(elem.styleName);
		elem.props.tabs ??= style?.paragraphProps?.tabs;  //TODO

		this.renderClass(elem, result);
		this.renderChildren(elem, result);
		this.renderStyleValues(elem.cssStyle, result);
		this.renderCommonProperties(result.style, elem.props);

		const numbering = elem.props.numbering ?? style?.paragraphProps?.numbering;

		if (numbering) {
			result.classList.add(this.numberingClass(numbering.id, numbering.level));
		}

		return result;
	}

	renderRun(elem: WmlRun) {
		if (elem.fieldRun)
			return null;

		const result = createElement("span");

		this.renderClass(elem, result);
		this.renderStyleValues(elem.cssStyle, result);

		if (elem.verticalAlign) {
			const wrapper = createElement(elem.verticalAlign as any);
			this.renderChildren(elem, wrapper);
			result.appendChild(wrapper);
		} else {
			this.renderChildren(elem, result);
		}

		return result;
	}

	renderText(elem: WmlText) {
		return document.createTextNode(elem.text);
	}

	renderHyperlink(elem: WmlHyperlink) {
		let result = createElement("a");

		this.renderChildren(elem, result);
		this.renderStyleValues(elem.cssStyle, result);

		if (elem.href) {
			result.href = elem.href;
		} else if (elem.id) {
			const rel = this.document.documentPart.rels
				.find(it => it.id == elem.id && it.targetMode === "External");
			result.href = rel?.target;
		}

		return result;
	}

	renderDrawing(elem: OpenXmlElement) {
		let result = createElement("div");

		result.style.display = "inline-block";
		result.style.position = "relative";
		result.style.textIndent = "0px";

		this.renderChildren(elem, result);
		this.renderStyleValues(elem.cssStyle, result);

		return result;
	}

	// 渲染图片，默认转换blob--异步
	renderImage(elem: WmlImage) {
		let result = createElement("img");

		this.renderStyleValues(elem.cssStyle, result);

		if (this.document) {
			this.document
				.loadDocumentImage(elem.src, this.currentPart)
				.then(src => {
					result.src = src;
				});
		}

		return result;
	}

	renderDeletedText(elem: WmlText) {
		return this.options.renderEndnotes ? document.createTextNode(elem.text) : null;
	}

	renderBreak(elem: WmlBreak) {
		if (elem.break == "textWrapping") {
			return createElement("br");
		}

		return null;
	}

	renderInserted(elem: OpenXmlElement): Node | Node[] {
		if (this.options.renderChanges) {
			return this.renderContainer(elem, "ins");
		}

		return this.renderChildren(elem);
	}

	renderDeleted(elem: OpenXmlElement): Node {
		if (this.options.renderChanges) {
			return this.renderContainer(elem, "del");
		}

		return null;
	}

	renderSymbol(elem: WmlSymbol) {
		let span = createElement("span");
		span.style.fontFamily = elem.font;
		span.innerHTML = `&#x${elem.char};`
		return span;
	}

	// 渲染页眉页脚
	renderHeaderFooter(elem: OpenXmlElement, tagName: keyof HTMLElementTagNameMap,) {

		let result: HTMLElement = createElement(tagName);
		// 渲染子元素
		this.renderChildren(elem, result);
		// 渲染style样式
		this.renderStyleValues(elem.cssStyle, result);

		return result;
	}

	renderFootnoteReference(elem: WmlNoteReference) {
		let result = createElement("sup");
		this.currentFootnoteIds.push(elem.id);
		result.textContent = `${this.currentFootnoteIds.length}`;
		return result;
	}

	renderEndnoteReference(elem: WmlNoteReference) {
		let result = createElement("sup");
		this.currentEndnoteIds.push(elem.id);
		result.textContent = `${this.currentEndnoteIds.length}`;
		return result;
	}

	// 渲染制表符
	renderTab(elem: OpenXmlElement) {
		let tabSpan = createElement("span");

		tabSpan.innerHTML = "&emsp;";//"&nbsp;";

		if (this.options.experimental) {
			tabSpan.className = this.tabStopClass();
			let stops = findParent<WmlParagraph>(elem, DomType.Paragraph).props?.tabs;
			this.currentTabs.push({ stops, span: tabSpan });
		}

		return tabSpan;
	}

	renderBookmarkStart(elem: WmlBookmarkStart): HTMLElement {
		let result = createElement("span");
		result.id = elem.name;
		return result;
	}

	renderTable(elem: WmlTable) {
		let oTable = createElement("table");

		this.tableCellPositions.push(this.currentCellPosition);
		this.tableVerticalMerges.push(this.currentVerticalMerge);
		this.currentVerticalMerge = {};
		this.currentCellPosition = { col: 0, row: 0 };
		// 渲染表格column列
		if (elem.columns) {
			oTable.appendChild(this.renderTableColumns(elem.columns));
		}

		this.renderClass(elem, oTable);
		this.renderChildren(elem, oTable);
		this.renderStyleValues(elem.cssStyle, oTable);

		this.currentVerticalMerge = this.tableVerticalMerges.pop();
		this.currentCellPosition = this.tableCellPositions.pop();
		return oTable;
	}

	renderTableColumns(columns: WmlTableColumn[]) {
		let result = createElement("colgroup");

		for (let col of columns) {
			let colElem = createElement("col");

			if (col.width)
				colElem.style.width = col.width;

			result.appendChild(colElem);
		}

		return result;
	}

	renderTableRow(elem: OpenXmlElement) {
		let result = createElement("tr");

		this.currentCellPosition.col = 0;

		this.renderClass(elem, result);
		this.renderChildren(elem, result);
		this.renderStyleValues(elem.cssStyle, result);

		this.currentCellPosition.row++;

		return result;
	}

	renderTableCell(elem: WmlTableCell) {
		let result = createElement("td");

		const key = this.currentCellPosition.col;

		if (elem.verticalMerge) {
			if (elem.verticalMerge == "restart") {
				this.currentVerticalMerge[key] = result;
				result.rowSpan = 1;
			} else if (this.currentVerticalMerge[key]) {
				this.currentVerticalMerge[key].rowSpan += 1;
				result.style.display = "none";
			}
		} else {
			this.currentVerticalMerge[key] = null;
		}

		this.renderClass(elem, result);
		this.renderChildren(elem, result);
		this.renderStyleValues(elem.cssStyle, result);

		if (elem.span)
			result.colSpan = elem.span;

		this.currentCellPosition.col += result.colSpan;

		return result;
	}

	renderVmlPicture(elem: OpenXmlElement) {
		let result = createElement("div");
		this.renderChildren(elem, result);
		return result;
	}

	renderVmlElement(elem: VmlElement): SVGElement {
		let container = createSvgElement("svg");

		container.setAttribute("style", elem.cssStyleText);

		const result = this.renderVmlChildElement(elem);

		if (elem.imageHref?.id) {
			this.document?.loadDocumentImage(elem.imageHref.id, this.currentPart)
				.then(x => result.setAttribute("href", x));
		}

		container.appendChild(result);

		requestAnimationFrame(() => {
			const bb = (container.firstElementChild as any).getBBox();

			container.setAttribute("width", `${Math.ceil(bb.x + bb.width)}`);
			container.setAttribute("height", `${Math.ceil(bb.y + bb.height)}`);
		});

		return container;
	}

	renderVmlChildElement(elem: VmlElement) {
		const result = createSvgElement(elem.tagName as any);
		Object.entries(elem.attrs).forEach(([k, v]) => result.setAttribute(k, v));

		for (let child of elem.children) {
			if (child.type == DomType.VmlElement) {
				result.appendChild(this.renderVmlChildElement(child as VmlElement));
			} else {
				result.append(...asArray(this.renderElement(child as any)));
			}
		}

		return result;
	}

	renderMmlRadical(elem: OpenXmlElement): HTMLElement {
		const base = elem.children.find(el => el.type == DomType.MmlBase);

		if (elem.props?.hideDegree) {
			return createElementNS(ns.mathML, "msqrt", null, this.renderElements([base]));
		}

		const degree = elem.children.find(el => el.type == DomType.MmlDegree);
		return createElementNS(ns.mathML, "mroot", null, this.renderElements([base, degree]));
	}

	renderMmlDelimiter(elem: OpenXmlElement): HTMLElement {
		const children = [];

		children.push(createElementNS(ns.mathML, "mo", null, [elem.props.beginChar ?? '(']));
		children.push(...this.renderElements(elem.children));
		children.push(createElementNS(ns.mathML, "mo", null, [elem.props.endChar ?? ')']));

		return createElementNS(ns.mathML, "mrow", null, children);
	}

	renderMmlNary(elem: OpenXmlElement): HTMLElement {
		const children = [];
		const grouped = _.keyBy(elem.children, 'type');

		const sup = grouped[DomType.MmlSuperArgument];
		const sub = grouped[DomType.MmlSubArgument];
		const supElem = sup ? createElementNS(ns.mathML, "mo", null, asArray(this.renderElement(sup))) : null;
		const subElem = sub ? createElementNS(ns.mathML, "mo", null, asArray(this.renderElement(sub))) : null;


		const charElem = createElementNS(ns.mathML, "mo", null, [elem.props?.char ?? '\u222B']);

		if (supElem || subElem) {
			children.push(createElementNS(ns.mathML, "munderover", null, [charElem, subElem, supElem]));
		} else if (supElem) {
			children.push(createElementNS(ns.mathML, "mover", null, [charElem, supElem]));
		} else if (subElem) {
			children.push(createElementNS(ns.mathML, "munder", null, [charElem, subElem]));
		} else {
			children.push(charElem);
		}

		children.push(...this.renderElements(grouped[DomType.MmlBase].children));

		return createElementNS(ns.mathML, "mrow", null, children);
	}

	renderMmlPreSubSuper(elem: OpenXmlElement) {
		const children = [];
		const grouped = _.keyBy(elem.children, 'type');

		const sup = grouped[DomType.MmlSuperArgument];
		const sub = grouped[DomType.MmlSubArgument];
		const supElem = sup ? createElementNS(ns.mathML, "mo", null, asArray(this.renderElement(sup))) : null;
		const subElem = sub ? createElementNS(ns.mathML, "mo", null, asArray(this.renderElement(sub))) : null;
		const stubElem = createElementNS(ns.mathML, "mo", null);

		children.push(createElementNS(ns.mathML, "msubsup", null, [stubElem, subElem, supElem]));
		children.push(...this.renderElements(grouped[DomType.MmlBase].children));

		return createElementNS(ns.mathML, "mrow", null, children);
	}

	renderMmlGroupChar(elem: OpenXmlElement) {
		const tagName = elem.props.verticalJustification === "bot" ? "mover" : "munder";
		const result = this.renderContainerNS(elem, ns.mathML, tagName);

		if (elem.props.char) {
			result.appendChild(createElementNS(ns.mathML, "mo", null, [elem.props.char]));
		}

		return result;
	}

	renderMmlBar(elem: OpenXmlElement) {
		const result = this.renderContainerNS(elem, ns.mathML, "mrow");

		switch (elem.props.position) {
			case "top":
				result.style.textDecoration = "overline";
				break
			case "bottom":
				result.style.textDecoration = "underline";
				break
		}

		return result;
	}

	renderMmlRun(elem: OpenXmlElement) {
		const result = createElementNS(ns.mathML, "ms");

		this.renderClass(elem, result);
		this.renderStyleValues(elem.cssStyle, result);
		this.renderChildren(elem, result);

		return result;
	}

	renderMllList(elem: OpenXmlElement) {
		const result = createElementNS(ns.mathML, "mtable");
		// 添加class类
		this.renderClass(elem, result);
		// 渲染style样式
		this.renderStyleValues(elem.cssStyle, result);

		const children = this.renderChildren(elem);

		for (let child of children) {
			result.appendChild(createElementNS(ns.mathML, "mtr", null, [
				createElementNS(ns.mathML, "mtd", null, [child])
			]));
		}

		return result;
	}

	// 设置元素style样式
	renderStyleValues(style: Record<string, string>, output: HTMLElement) {
		for (let k in style) {
			if (k.startsWith("$")) {
				output.setAttribute(k.slice(1), style[k]);
			} else {
				output.style[k] = style[k];
			}
		}
	}

	renderRunProperties(style: any, props: RunProperties) {
		this.renderCommonProperties(style, props);
	}

	renderCommonProperties(style: any, props: CommonProperties) {
		if (props == null)
			return;

		if (props.color) {
			style["color"] = props.color;
		}

		if (props.fontSize) {
			style["font-size"] = props.fontSize;
		}
	}

	// 添加class类名
	renderClass(input: OpenXmlElement, output: HTMLElement) {
		if (input.className) {
			output.className = input.className;
		}

		if (input.styleName) {
			output.classList.add(this.processStyleName(input.styleName));
		}
	}

	// 查找内置默认style样式
	findStyle(styleName: string) {
		return styleName && this.styleMap?.[styleName];
	}

	tabStopClass() {
		return `${this.className}-tab-stop`;
	}

	// 刷新tab制表符
	refreshTabStops() {
		if (!this.options.experimental) {
			return;
		}

		clearTimeout(this.tabsTimeout);

		this.tabsTimeout = setTimeout(() => {
			const pixelToPoint = computePointToPixelRatio();

			for (let tab of this.currentTabs) {
				updateTabStop(tab.span, tab.stops, this.defaultTabSize, pixelToPoint);
			}
		}, 500);
	}

}

type ChildType = Node | string;

function createElement<T extends keyof HTMLElementTagNameMap>(tagName: T, props?: Partial<Record<keyof HTMLElementTagNameMap[T], any>>, children?: ChildType[]): HTMLElementTagNameMap[T] {
	return createElementNS(undefined, tagName, props, children);
}

function createSvgElement<T extends keyof SVGElementTagNameMap>(tagName: T, props?: Partial<Record<keyof SVGElementTagNameMap[T], any>>, children?: ChildType[]): SVGElementTagNameMap[T] {
	return createElementNS(ns.svg, tagName, props, children);
}

function createElementNS(ns: string, tagName: string, props?: Partial<Record<any, any>>, children?: ChildType[]): any {
	let result = ns ? document.createElementNS(ns, tagName) : document.createElement(tagName);
	Object.assign(result, props);
	children && appendChildren(result, children);
	return result;
}

function removeAllElements(elem: HTMLElement) {
	elem.innerHTML = '';
}

// 插入子元素
function appendChildren(parent: HTMLElement | Element, children: (Node | string)[]) {
	children.forEach(child => {
		parent.appendChild(_.isString(child) ? document.createTextNode(child) : child)
	});
}

// 创建style标签
function createStyleElement(cssText: string) {
	return createElement("style", { innerHTML: cssText });
}

// 插入注释
function appendComment(elem: HTMLElement, comment: string) {
	elem.appendChild(document.createComment(comment));
}

function findParent<T extends OpenXmlElement>(elem: OpenXmlElement, type: DomType): T {
	let parent = elem.parent;

	while (parent != null && parent.type != type)
		parent = parent.parent;

	return <T>parent;
}
