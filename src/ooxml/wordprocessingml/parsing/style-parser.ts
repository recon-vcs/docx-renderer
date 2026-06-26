import { IDomStyle, Ruleset } from '@docx/ooxml/wordprocessingml/document/model/style';
import { parseParagraphProperties } from '@docx/ooxml/wordprocessingml/document/model/paragraph';
import { parseRunProperties } from '@docx/ooxml/wordprocessingml/document/model/run';
import { parseLineSpacing } from '@docx/ooxml/wordprocessingml/document/model/spacing-between-lines';
import type { DocumentParserOptions } from '@docx/ooxml/wordprocessingml/parsing/document-parser';
import xml from '@docx/xml/parsing/xml-parser';
import { xmlUtil } from '@docx/xml/parsing/parse-utils';

// Callback for parseDefaultProperties which remains in DocumentParser
export interface StyleParserCallbacks {
	parseDefaultProperties(
		elem: Element,
		style?: Record<string, string>,
		childStyle?: Record<string, string>,
		handler?: (prop: Element) => boolean
	): Record<string, string>;
}

export function parseStylesFile(
	xstyles: Element,
	options: DocumentParserOptions,
	callbacks: StyleParserCallbacks
): IDomStyle[] {
	let result: IDomStyle[] = [];

	xmlUtil.foreach(xstyles, n => {
		switch (n.localName) {
			case "style":
				result.push(parseStyle(n, options, callbacks));
				break;

			case "docDefaults":
				result.push(parseDefaultStyles(n, options, callbacks));
				break;

			default:
				if (options.debug) {
					console.warn(`DOCX:%c Unknown Style File：${n.localName}`, 'color:#f75607');
				}
		}
	});

	return result;
}

export function parseDefaultStyles(
	node: Element,
	options: DocumentParserOptions,
	callbacks: StyleParserCallbacks
): IDomStyle {
	let result = <IDomStyle>{
		basedOn: null,
		id: null,
		name: null,
		rulesets: [],
		type: null
	};

	xmlUtil.foreach(node, c => {
		switch (c.localName) {
			case "rPrDefault": {
				let rPr = xml.element(c, "rPr");
				if (rPr) {
					result.rulesets.push({
						target: "span",
						declarations: callbacks.parseDefaultProperties(rPr, {})
					});
				}
				break;
			}

			case "pPrDefault": {
				let pPr = xml.element(c, "pPr");
				if (pPr) {
					let paragraphProperties = parseParagraphProperties(pPr, xml);
					let ruleset = {
						target: "p",
						declarations: callbacks.parseDefaultProperties(pPr, {})
					};
					Object.assign(ruleset.declarations, parseLineSpacing(paragraphProperties));
					result.rulesets.push(ruleset);
				}
				break;
			}

			default:
				if (options.debug) {
					console.warn(`DOCX:%c Unknown Default Style：${c.localName}`, 'color:#f75607');
				}
		}
	});

	return result;
}

export function parseStyle(
	node: Element,
	options: DocumentParserOptions,
	callbacks: StyleParserCallbacks
): IDomStyle {
	let result: IDomStyle = <IDomStyle>{
		basedOn: null,
		id: null,
		name: null,
		rulesets: [],
		type: null,
	};

	for (const attr of xml.attrs(node)) {
		switch (attr.localName) {
			// User-Defined Style
			case "customStyle":
				result.customStyle = xml.boolAttr(node, "customStyle", false);
				break;

			// Default Style
			case "default":
				result.isDefault = xml.boolAttr(node, "default", false);
				break;

			// Style ID
			case "styleId":
				result.id = xml.attr(node, "styleId");
				break;

			// Style Type
			case "type":
				result.type = xml.attr(node, "type");
				const typeToLabelMap: Record<string, string> = {
					"paragraph": "p",
					"table": "table",
					"character": "span",
					"numbering": "p",
				};
				if (typeToLabelMap.hasOwnProperty(result.type)) {
					result.label = typeToLabelMap[result.type];
				} else {
					if (options && options.debug) {
						console.warn(`DOCX:%c Unknown Style Type：${result.type}`, 'color:#f75607');
					}
				}
				break;

			default:
				if (options.debug) {
					console.warn(`DOCX:%c Unknown Style Property：${attr.localName}`, 'color:#f75607');
				}
		}
	}

	xmlUtil.foreach(node, n => {
		switch (n.localName) {
			// Alternate Style Names
			case "aliases":
				result.aliases = xml.attr(n, "val").split(",");
				break;

			// Automatically Merge User Formatting Into Style Definition
			case "autoRedefine":
				result.autoRedefine = true;
				break;

			// Parent Style ID
			case "basedOn":
				result.basedOn = xml.attr(n, "val");
				break;

			// Hide Style From User Interface
			case "hidden":
				result.hidden = true;
				break;

			// Linked Style Reference
			case "link":
				result.linked = xml.attr(n, "val");
				break;

			// Style Cannot Be Applied
			case "locked":
				result.locked = true;
				break;

			// Primary Style Name
			case "name":
				result.name = xml.attr(n, "val");
				break;

			// Style For Next Paragraph
			case "next":
				result.next = xml.attr(n, "val");
				break;

			// E-Mail Message Text Style
			case "personal":
				result.personal = xml.boolAttr(n, "val");
				break;

			// E-Mail Message Composition Style
			case "personalCompose":
				result.personalCompose = xml.boolAttr(n, "val");
				break;

			// E-Mail Message Reply Style
			case "personalReply":
				result.personalReply = xml.boolAttr(n, "val");
				break;

			// Style Paragraph Properties
			case "pPr": {
				result.paragraphProps = parseParagraphProperties(n, xml);
				let ruleset = {
					target: "p",
					declarations: callbacks.parseDefaultProperties(n, {})
				};
				Object.assign(ruleset.declarations, parseLineSpacing(result.paragraphProps));
				result.rulesets.push(ruleset);
				break;
			}

			// Specifies Primary Style
			case "qFormat":
				result.primaryStyle = true;
				break;

			// Run Properties
			case "rPr":
				result.rulesets.push({
					target: "span",
					declarations: callbacks.parseDefaultProperties(n, {})
				});
				result.runProps = parseRunProperties(n, xml);
				break;

			// Revision Identifier for Style Definition
			case "rsid":
				result.rsid = xml.hexAttr(n, "val");
				break;

			// Hide Style From Main User Interface
			case "semiHidden":
				result.semiHidden = true;
				break;

			// Style Table Properties
			case "tblPr":
				result.rulesets.push({
					target: "td",
					declarations: callbacks.parseDefaultProperties(n, {})
				});
				break;

			// Style Table Row Properties
			case "trPr":
				// TODO: maybe move to processor
				result.rulesets.push({
					target: "tr",
					declarations: callbacks.parseDefaultProperties(n, {})
				});
				break;

			// Style Table Cell Properties
			case "tcPr":
				result.rulesets.push({
					target: "td",
					declarations: callbacks.parseDefaultProperties(n, {})
				});
				break;

			// Style Conditional Table Formatting Properties
			case "tblStylePr":
				for (let s of parseTableStyle(n, options, callbacks)) {
					result.rulesets.push(s);
				}
				break;

			// Optional User Interface Sorting Order
			case "uiPriority":
				result.uiPriority = xml.intAttr(n, "val", Infinity);
				break;

			// Remove Semi-Hidden Property When Style Is Used
			case "unhideWhenUsed":
				result.unhideWhenUsed = true;
				break;

			default:
				if (options.debug) {
					console.warn(`DOCX:%c Unknown Style element：${n.localName}`, 'color:blue');
				}
		}
	});

	return result;
}

// TODO: multi-level nested table style rules are not yet applied
export function parseTableStyle(
	node: Element,
	options: DocumentParserOptions,
	callbacks: StyleParserCallbacks
): Ruleset[] {
	let result: Ruleset[] = [];

	let type = xml.attr(node, "type");
	let selector = "";
	let modifier = "";

	switch (type) {
		case "firstRow":
			modifier = ".first-row";
			selector = "tr.first-row td";
			break;
		case "lastRow":
			modifier = ".last-row";
			selector = "tr.last-row td";
			break;
		case "firstCol":
			modifier = ".first-col";
			selector = "td.first-col";
			break;
		case "lastCol":
			modifier = ".last-col";
			selector = "td.last-col";
			break;
		case "band1Vert":
			modifier = ":not(.no-vband)";
			selector = "td.odd-col";
			break;
		case "band2Vert":
			modifier = ":not(.no-vband)";
			selector = "td.even-col";
			break;
		case "band1Horz":
			modifier = ":not(.no-hband)";
			selector = "tr.odd-row";
			break;
		case "band2Horz":
			modifier = ":not(.no-hband)";
			selector = "tr.even-row";
			break;
		default:
			return [];
	}

	xmlUtil.foreach(node, n => {
		switch (n.localName) {
			case "pPr": {
				let paragraphProperties = parseParagraphProperties(n, xml);
				let ruleset = {
					target: `${selector} p`,
					modifier: modifier,
					declarations: callbacks.parseDefaultProperties(n, {})
				};
				Object.assign(ruleset.declarations, parseLineSpacing(paragraphProperties));
				result.push(ruleset);
				break;
			}

			case "rPr":
				result.push({
					target: `${selector} span`,
					modifier: modifier,
					declarations: callbacks.parseDefaultProperties(n, {})
				});
				break;

			case "tblPr":
			case "tcPr":
				result.push({
					target: selector, // TODO: maybe move to processor
					modifier: modifier,
					declarations: callbacks.parseDefaultProperties(n, {})
				});
				break;

			default:
				if (options.debug) {
					console.warn(`DOCX:%c Unknown Table Style：${n.localName}`, 'color:#f75607');
				}
		}
	});

	return result;
}
