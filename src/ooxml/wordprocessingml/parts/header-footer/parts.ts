import { OpenXmlPackage } from '@docx/opc/package/open-xml-package';
import { Part } from '@docx/opc/parts/part';
import { DocumentParser } from '@docx/ooxml/wordprocessingml/parsing/document-parser';
import { OpenXmlElement } from '@docx/ooxml/wordprocessingml/document/model/dom';
import { WmlHeader, WmlFooter } from "./elements";

export abstract class BaseHeaderFooterPart<T extends OpenXmlElement = OpenXmlElement> extends Part {
	rootElement: T;

	private _documentParser: DocumentParser;

	constructor(pkg: OpenXmlPackage, path: string, parser: DocumentParser) {
		super(pkg, path);
		this._documentParser = parser;
	}

	parseXml(root: Element) {
		this.rootElement = this.createRootElement();
		this.rootElement.level = 1;
		this.rootElement.children = this._documentParser.parseBodyElements(root);
	}

	protected abstract createRootElement(): T;
}

export class HeaderPart extends BaseHeaderFooterPart<WmlHeader> {
	protected createRootElement(): WmlHeader {
		return new WmlHeader();
	}
}

export class FooterPart extends BaseHeaderFooterPart<WmlFooter> {
	protected createRootElement(): WmlFooter {
		return new WmlFooter();
	}
}
