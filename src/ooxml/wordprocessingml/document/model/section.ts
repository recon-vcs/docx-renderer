import globalXmlParser, { XmlParser } from '@docx/xml/parsing/xml-parser';
import { Borders, parseBorders } from './border';
import { Length, convertLength } from './common';

export interface Column {
	space: Length;
	width: Length;
}

export interface Columns {
	columns: Column[];
	count: number;
	equalWidth: boolean;
	separator: boolean;
	space: Length;
}

export interface ContentSize {
	height: Length,
	width: Length,
}

export interface PageSize extends ContentSize {
	orientation: "landscape" | string
}

export interface PageNumber {
	chapSep: "colon" | "emDash" | "endash" | "hyphen" | "period" | string;
	chapStyle: string;
	format: "none" | "cardinalText" | "decimal" | "decimalEnclosedCircle" | "decimalEnclosedFullstop"
		| "decimalEnclosedParen" | "decimalZero" | "lowerLetter" | "lowerRoman"
		| "ordinalText" | "upperLetter" | "upperRoman" | string;
	start: number;
}

export interface PageMargins {
	bottom: Length;
	footer: Length;
	gutter: Length;
	header: Length;
	left: Length;
	right: Length;
	top: Length;
}

export enum SectionType {
	Continuous = "continuous",
	// default
	NextPage = "nextPage",
	NextColumn = "nextColumn",
	EvenPage = "evenPage",
	OddPage = "oddPage",
}

export interface FooterHeaderReference {
	id: string;
	type: string | "first" | "even" | "default";
}

export interface DocGrid {
	characterSpace?: number;
	linePitch?: number;
	type?: DocGridType;
}

export enum DocGridType {
	// No Document Grid.Default Value
	Default = 'default',
	// Line Grid Only.
	Lines = 'lines',
	// Line and Character Grid.
	LinesAndChars = 'linesAndChars',
	// Character Grid Only.
	SnapToChars = 'snapToChars',
}

export interface SectionProperties {
	columns: Columns;
	contentSize: ContentSize;
	docGrid: DocGrid;
	footerRefs: FooterHeaderReference[],
	headerRefs: FooterHeaderReference[],
	pageBorders: Borders;
	pageMargins: PageMargins;
	pageNumber: PageNumber;
	pageSize: PageSize;
	sectionId: string;
	titlePage: boolean;
	type: SectionType;
}

// 原始尺寸数据，单位：dxa
interface OriginSize {
	pageMargins: {
		bottom: number;
		footer: number;
		gutter: number;
		header: number;
		left: number;
		right: number;
		top: number;
	},
	pageSize: {
		height: number,
		width: number,
	}
}

export function parseSectionProperties(elem: Element, xml: XmlParser = globalXmlParser): SectionProperties {
	let section = <SectionProperties>{
		contentSize: {},
	};
	// 原始尺寸，单位：dxa
	let origin = <OriginSize>{};

	for (let e of xml.elements(elem)) {
		switch (e.localName) {
			// TODO Right to Left Section Layout
			case "bidi":
				break;

			// Column Definitions
			case "cols":
				section.columns = parseColumns(e, xml);
				break;

			// Document Grid
			case "docGrid":
				section.docGrid = parseDocGrid(e, xml);
				break;

			// TODO Section-Wide Endnote Properties
			case "endnotePr":

				break;

			// Footer Reference
			case "footerReference":
				(section.footerRefs ?? (section.footerRefs = [])).push(parseFooterHeaderReference(e, xml));
				break;

			// TODO Section-Wide Footnote Properties
			case "footnotePr":

				break;

			// TODO Only Allow Editing of Form Fields
			case "formProt":

				break;

			// Header Reference
			case "headerReference":
				(section.headerRefs ?? (section.headerRefs = [])).push(parseFooterHeaderReference(e, xml));
				break;

			// TODO Line Numbering Settings
			case "lnNumType":

				break;

			// TODO Suppress Endnotes In Document
			case "noEndnote":

				break;

			// TODO Paper Source Information
			case "paperSrc":

				break;

			// Page Borders
			case "pgBorders":
				section.pageBorders = parseBorders(e, xml);
				break;

			// Page Margins
			case "pgMar":
				section.pageMargins = {
					left: xml.lengthAttr(e, "left"),
					right: xml.lengthAttr(e, "right"),
					top: xml.lengthAttr(e, "top"),
					bottom: xml.lengthAttr(e, "bottom"),
					header: xml.lengthAttr(e, "header"),
					footer: xml.lengthAttr(e, "footer"),
					gutter: xml.lengthAttr(e, "gutter"),
				};
				// 记录原始尺寸
				origin.pageMargins = {
					left: xml.intAttr(e, "left"),
					right: xml.intAttr(e, "right"),
					top: xml.intAttr(e, "top"),
					bottom: xml.intAttr(e, "bottom"),
					header: xml.intAttr(e, "header"),
					footer: xml.intAttr(e, "footer"),
					gutter: xml.intAttr(e, "gutter"),
				}
				break;

			// Page Numbering Settings
			case "pgNumType":
				section.pageNumber = parsePageNumber(e, xml);
				break;

			// Page Size
			case "pgSz":
				section.pageSize = {
					width: xml.lengthAttr(e, "w"),
					height: xml.lengthAttr(e, "h"),
					orientation: xml.attr(e, "orient")
				}
				// 记录原始尺寸
				origin.pageSize = {
					width: xml.intAttr(e, "w"),
					height: xml.intAttr(e, "h"),
				}
				break;

			// TODO Reference to Printer Settings Data
			case "printerSettings":

				break;

			// TODO Gutter on Right Side of Page
			case "rtlGutter":

				break;

			// TODO Revision Information for Section Properties
			case "sectPrChange":

				break;

			// TODO Text Flow Direction
			case "textDirection":

				break;

			// Different First Page Headers and Footers
			case "titlePg":
				section.titlePage = xml.boolAttr(e, "val", true);
				break;

			// Section Type
			case "type":
				section.type = xml.attr(e, "val") as SectionType;
				break;

			// TODO Vertical Text Alignment on Page
			case "vAlign":

				break;

			default:
				break;
		}
	}
	// Compute content width from page size minus horizontal margins when both are present.
	if (origin.pageSize && origin.pageMargins) {
		const { width } = origin.pageSize;
		const { left, right } = origin.pageMargins;
		section.contentSize.width = convertLength(width - left - right) as string;
	}

	return section;
}

function parseColumns(elem: Element, xml: XmlParser): Columns {
	return {
		count: xml.intAttr(elem, "num"),
		space: xml.lengthAttr(elem, "space"),
		separator: xml.boolAttr(elem, "sep"),
		equalWidth: xml.boolAttr(elem, "equalWidth", true),
		columns: xml.elements(elem, "col")
			.map(e => <Column>{
				width: xml.lengthAttr(e, "w"),
				space: xml.lengthAttr(e, "space")
			})
	};
}

function parseFooterHeaderReference(elem: Element, xml: XmlParser): FooterHeaderReference {
	return {
		id: xml.attr(elem, "id"),
		type: xml.attr(elem, "type"),
	}
}

// TODO only support linePitch property temporarily
function parseDocGrid(elem: Element, xml: XmlParser): DocGrid {
	let grid: DocGrid = {
		type: DocGridType.Default,
	};
	for (let attr of xml.attrs(elem)) {
		switch (attr.localName) {
			case "charSpace":
				grid.characterSpace = xml.intAttr(elem, "charSpace");
				break;

			case "linePitch":
				grid.linePitch = xml.intAttr(elem, "linePitch");
				break;

			case "type":
				grid.type = xml.attr(elem, "type", DocGridType.Default) as DocGridType;
				break;

			default:
				break;
		}
	}
	return grid;
}

function parsePageNumber(elem: Element, xml: XmlParser): PageNumber {
	return {
		chapSep: xml.attr(elem, "chapSep"),
		chapStyle: xml.attr(elem, "chapStyle"),
		format: xml.attr(elem, "fmt"),
		start: xml.intAttr(elem, "start")
	};
}


