import { Length } from '@docx/ooxml/wordprocessingml/document/model/common';
import { XmlParser } from '@docx/xml/parsing/xml-parser';

export interface WmlSettings {
	autoHyphenation: boolean;
	defaultTabStop: Length;
	endnoteProps: NoteProperties;
	evenAndOddHeaders: boolean;
	footnoteProps: NoteProperties;
	themeFontLang: ThemeFontLanguages;
}

export interface ThemeFontLanguages {
	val?: string;
	eastAsia?: string;
	bidi?: string;
}

export interface NoteProperties {
	nummeringFormat: string;
	defaultNoteIds: string[];
}

export function parseSettings(elem: Element, xml: XmlParser) {
	var result = {} as WmlSettings;
	// TODO support more settings
	for (let el of xml.elements(elem)) {
		switch (el.localName) {
			// Automatically Hyphenate Document Contents When Displayed
			case "autoHyphenation":
				result.autoHyphenation = xml.boolAttr(el, "val");
				break;
			// Distance Between Automatic Tab Stops
			case "defaultTabStop":
				result.defaultTabStop = xml.lengthAttr(el, "val");
				break;
			// Document-Wide Endnote Properties
			case "endnotePr":
				result.endnoteProps = parseNoteProperties(el, xml);
				break;
			// 	Different Even/Odd Page Headers and Footers
			case "evenAndOddHeaders":
				result.evenAndOddHeaders = xml.boolAttr(el, "val", true);
				break;
			// Document-Wide Footnote Properties
			case "footnotePr":
				result.footnoteProps = parseNoteProperties(el, xml);
				break;
			// Theme Font Languages - select which <a:font script="..."> entry
			// substitutes for an empty theme typeface
			case "themeFontLang":
				result.themeFontLang = {
					val: xml.attr(el, "val"),
					eastAsia: xml.attr(el, "eastAsia"),
					bidi: xml.attr(el, "bidi"),
				};
				break;

		}
	}

	return result;
}

export function parseNoteProperties(elem: Element, xml: XmlParser) {
	var result = {
		defaultNoteIds: []
	} as NoteProperties;

	for (let el of xml.elements(elem)) {
		switch (el.localName) {
			case "numFmt":
				result.nummeringFormat = xml.attr(el, "val");
				break;

			case "footnote":
			case "endnote":
				result.defaultNoteIds.push(xml.attr(el, "id"));
				break;
		}
	}

	return result;
}
