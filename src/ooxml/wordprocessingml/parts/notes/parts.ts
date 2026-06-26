import { OpenXmlPackage } from '@docx/opc/package/open-xml-package';
import { Part } from '@docx/opc/parts/part';
import { DocumentParser } from '@docx/ooxml/wordprocessingml/parsing/document-parser';
import { WmlNotes, WmlFootnotes, WmlFootnote, WmlEndnotes, WmlEndnote } from "./elements";

export class BaseNotePart<T extends WmlNotes> extends Part {
	rootElement: T;

	protected _documentParser: DocumentParser;

	constructor(pkg: OpenXmlPackage, path: string, parser: DocumentParser) {
		super(pkg, path);
		this._documentParser = parser;
	}

}

export class FootnotesPart extends BaseNotePart<WmlFootnotes> {
	constructor(pkg: OpenXmlPackage, path: string, parser: DocumentParser) {
		super(pkg, path, parser);
	}

	parseXml(root: Element) {
		this.rootElement = new WmlFootnotes();
		this.rootElement.children = this._documentParser.parseNotes(root, "footnote", WmlFootnote);
	}
}

export class EndnotesPart extends BaseNotePart<WmlEndnotes> {
	constructor(pkg: OpenXmlPackage, path: string, parser: DocumentParser) {
		super(pkg, path, parser);
	}

	parseXml(root: Element) {
		this.rootElement = new WmlEndnotes();
		this.rootElement.children = this._documentParser.parseNotes(root, "endnote", WmlEndnote);
	}
}
