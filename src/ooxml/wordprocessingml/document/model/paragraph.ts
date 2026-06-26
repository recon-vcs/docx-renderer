import { OpenXmlElement } from "./dom";
import { CommonProperties, Length, ns, parseCommonProperty } from "./common";
import { Borders } from "./border";
import { parseSectionProperties, SectionProperties } from "./section";
import { SpacingBetweenLines, parseSpacingBetweenLines } from "./spacing-between-lines";
import { XmlParser } from '@docx/xml/parsing/xml-parser';
import { parseRunProperties, RunProperties } from "./run";

// TODO WmlParagraph段落不应该继承段落属性ParagraphProperties，参照微软SDK文档
export interface WmlParagraph extends OpenXmlElement {
	props: ParagraphProperties;
}

export interface ParagraphProperties extends CommonProperties {
	keepLines?: boolean;
	keepNext?: boolean;
	numbering?: ParagraphNumbering;
	outlineLevel?: number;
	pageBreakBefore?: boolean;
	paragraphBorders?: Borders;
	runProperties?: RunProperties;
	sectionProperties?: SectionProperties;
	snapToGrid?: boolean;
	spacing?: SpacingBetweenLines;
	tabs?: TabStop[];

	textAlignment?: "auto" | "baseline" | "bottom" | "center" | "top" | string;
}

export interface TabStop {
	style: "bar" | "center" | "clear" | "decimal" | "end" | "num" | "start" | "left" | "right";
	leader: "none" | "dot" | "heavy" | "hyphen" | "middleDot" | "underscore";
	position: number;
}

export interface ParagraphNumbering {
	id: string;
	level: number;
}

export function parseParagraphProperties(elem: Element, xml: XmlParser): ParagraphProperties {
	let properties = <ParagraphProperties>{};

	for (let el of xml.elements(elem)) {
		parseParagraphProperty(el, properties, xml);
	}

	return properties;
}

export function parseParagraphProperty(elem: Element, props: ParagraphProperties, xml: XmlParser) {
	// namespace check
	if (elem.namespaceURI != ns.wordml) {
		return false;
	}

	if (parseCommonProperty(elem, props, xml))
		return true;

	switch (elem.localName) {
		// TODO Automatically Adjust Right Indent When Using Document Grid
		case "adjustRightInd":

			break;

		// TODO Automatically Adjust Spacing of Latin and East Asian Text
		case "autoSpaceDE":

			break;

		// TODO Automatically Adjust Spacing of East Asian Text and Numbers
		case "autoSpaceDN":

			break;

		// TODO Ignore Spacing Above and Below When Using Identical Styles
		case "contextualSpacing":

			break;

		// TODO Associated HTML div ID
		case "divId":

			break;

		// 	Keep All Lines On One Page
		case "keepLines":
			props.keepLines = xml.boolAttr(elem, "val", true);
			break;

		// Keep Paragraph With Next Paragraph
		case "keepNext":
			props.keepNext = xml.boolAttr(elem, "val", true);
			break;

		// Numbering Definition Instance Reference
		case "numPr":
			props.numbering = parseNumbering(elem, xml);
			break;

		// 	Associated Outline Level
		case "outlineLvl":
			props.outlineLevel = xml.intAttr(elem, "val");
			break;

		// 	Start Paragraph on Next Page
		case "pageBreakBefore":
			props.pageBreakBefore = xml.boolAttr(elem, "val", true);
			break;

		// TODO Run Properties for the Paragraph Mark
		case "rPr":
			props.runProperties = parseRunProperties(elem, xml);
			break;

		// Section Properties
		case "sectPr":
			props.sectionProperties = parseSectionProperties(elem, xml);
			break;

		// TODO Use Document Grid Settings For Inter-Character Spacing
		case "snapToGrid":
			props.snapToGrid = xml.boolAttr(elem, "val", true);
			break;

		// Spacing Between Lines and Above/Below Paragraph
		case "spacing":
			props.spacing = parseSpacingBetweenLines(elem, xml);
			return false; //TODO
			break;

		// Set of Custom Tab Stops
		case "tabs":
			props.tabs = parseTabs(elem, xml);
			break;

		// Vertical Character Alignment on Line
		case "textAlignment":
			props.textAlignment = xml.attr(elem, "val");
			return false; //TODO
			break;

		default:
			return false;
	}

	return true;
}

export function parseTabs(elem: Element, xml: XmlParser): TabStop[] {
	return xml.elements(elem, "tab")
		.map(e => <TabStop>{
			position: xml.numberAttr(e, "pos"),
			leader: xml.attr(e, "leader"),
			style: xml.attr(e, "val")
		});
}

export function parseNumbering(elem: Element, xml: XmlParser): ParagraphNumbering {
	let result = <ParagraphNumbering>{};

	for (let e of xml.elements(elem)) {
		switch (e.localName) {
			case "numId":
				result.id = xml.attr(e, "val");
				break;

			case "ilvl":
				result.level = xml.intAttr(e, "val");
				break;
		}
	}

	return result;
}
