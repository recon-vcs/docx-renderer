import type { DocumentParserOptions } from '../document-parser';
import xml from './xml-parser';
import { xmlUtil, values, autos } from './parse-utils';
import { LengthUsage } from '../document/common';
import type { WmlRun } from '../document/run';

export function parseDefaultProperties(
	elem: Element,
	options: DocumentParserOptions,
	style: Record<string, string> = null,
	childStyle: Record<string, string> = null,
	handler: (prop: Element) => boolean = null,
): Record<string, string> {
	style = style || {};

	xmlUtil.foreach(elem, c => {
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
				break;

			// Languages for Run Content, check spelling and grammar
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
				parseFont(c, style);
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
				style["text-decoration"] = xml.boolAttr(c, "val", true) ? "line-through" : "none";
				break;

			// Non-Complex Script Font Size
			case "sz":
				style["font-size"] = style["min-height"] = xml.lengthAttr(c, "val", LengthUsage.FontSize);
				break;

			// Complex Script Font Size
			case "szCs":
				break;

			// Underline
			case "u":
				parseUnderline(c, style, options);
				break;

			// Hidden Text
			case "vanish":
				if (xml.boolAttr(c, "val", true))
					style["display"] = "none";
				break;

			// TODO Subscript/Superscript Text
			case "vertAlign":
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

			case "tcW":
				if (options.ignoreWidth) {
				}
				break;

			case "tblW":
				style["width"] = values.valueOfSize(c, "w");
				break;

			case "trHeight":
				parseTrHeight(c, style);
				break;

			case "ind":
			case "tblInd":
				parseIndentation(c, style, options);
				break;

			case "tblBorders":
				parseBorderProperties(c, childStyle || style, options);
				break;

			case "tblCellSpacing":
				style["border-spacing"] = values.valueOfMargin(c);
				style["border-collapse"] = "separate";
				break;

			case "pBdr":
				parseBorderProperties(c, style, options);
				break;

			case "tcBorders":
				parseBorderProperties(c, style, options);
				break;

			// TODO
			case "noWrap":
				break;

			case "tblCellMar":
			case "tcMar":
				parseMarginProperties(c, childStyle || style, options);
				break;

			case "tblLayout":
				style["table-layout"] = values.valueOfTblLayout(c);
				break;

			case "vAlign":
				style["vertical-align"] = values.valueOfTextAlignment(c);
				break;

			case "wordWrap":
				if (xml.boolAttr(c, "val"))
					style["overflow-wrap"] = "break-word";
				break;

			case "suppressAutoHyphens":
				style["hyphens"] = xml.boolAttr(c, "val", true) ? "none" : "auto";
				break;

			// ignored — tabs parsed separately
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
				if (options.debug) {
					console.warn(`DOCX:%c Unknown Property Element：${elem.localName}.${c.localName}`, 'color:green');
				}
				break;
		}
	});

	return style;
}

export function parseUnderline(node: Element, style: Record<string, string>, options: DocumentParserOptions): void {
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
			if (options.debug) {
				console.warn(`DOCX:%c Unknown Underline Property：${val}`, 'color:#f75607');
			}
	}

	let col = xmlUtil.colorAttr(node, "color");
	if (col) {
		style["text-decoration-color"] = col;
	}
}

// Convert run fonts (ascii, eastAsia, ComplexScript, high ANSI)
export function parseFont(node: Element, style: Record<string, string>): void {
	let fonts = new Set();
	let ascii = xml.attr(node, "ascii");
	let ascii_theme = values.themeValue(node, "asciiTheme");
	fonts.add(ascii).add(ascii_theme);
	let east_Asia = xml.attr(node, "eastAsia");
	let east_Asia_theme = values.themeValue(node, "eastAsiaTheme");
	fonts.add(east_Asia).add(east_Asia_theme);
	let complex_script = xml.attr(node, "cs");
	let complex_script_theme = values.themeValue(node, "cstheme");
	fonts.add(complex_script).add(complex_script_theme);
	let high_ansi = xml.attr(node, "hAnsi");
	let high_ansi_theme = values.themeValue(node, "hAnsiTheme");
	fonts.add(high_ansi).add(high_ansi_theme);
	let unique_fonts = [...fonts].filter(x => x);
	if (unique_fonts.length > 0) {
		style["font-family"] = unique_fonts.join(', ');
	}
	style["_hint"] = xml.attr(node, "hint");
}

export function parseIndentation(node: Element, style: Record<string, string>, options: DocumentParserOptions): void {
	let indentation: Record<string, any> = {};
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
				if (options.debug) {
					console.warn(`DOCX:%c Unknown Indentation Property：${attr.localName}`, 'color:#f75607');
				}
		}
	}
	if (indentation.firstLine) style["text-indent"] = indentation.firstLine;
	if (indentation.hanging) style["text-indent"] = `-${indentation.hanging}`;
	if (indentation.left || indentation.start) style["padding-left"] = indentation.left || indentation.start;
	if (indentation.right || indentation.end) style["padding-right"] = indentation.right || indentation.end;
}

// Additional character pitch spacing applied to run content
export function parseSpacing(node: Element, run: WmlRun, options: DocumentParserOptions): void {
	for (const attr of xml.attrs(node)) {
		switch (attr.localName) {
			case "val":
				run.cssStyle["margin-bottom"] = xml.lengthAttr(node, "val");
				break;
			default:
				if (options.debug) {
					console.warn(`DOCX:%c Unknown Spacing Property：${attr.localName}`, 'color:#f75607');
				}
		}
	}
}

export function parseMarginProperties(node: Element, output: Record<string, string>, options: DocumentParserOptions): void {
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
				if (options.debug) {
					console.warn(`DOCX:%c Unknown Margin Property：${c.localName}`, 'color:#f75607');
				}
		}
	});
}

export function parseTrHeight(node: Element, output: Record<string, string>): void {
	switch (xml.attr(node, "hRule")) {
		case "exact":
			output["height"] = xml.lengthAttr(node, "val");
			break;
		case "atLeast":
		default:
			output["height"] = xml.lengthAttr(node, "val");
			// min-height doesn't work for tr
			break;
	}
}

export function parseBorderProperties(node: Element, output: Record<string, string>, options: DocumentParserOptions): void {
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
				if (options.debug) {
					console.warn(`DOCX:%c Unknown Border Property：${c.localName}`, 'color:#f75607');
				}
		}
	});
}
