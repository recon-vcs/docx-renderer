import { OpenXmlPackage } from '@docx/opc/package/open-xml-package';
import { Part } from '@docx/opc/parts/part';
import { DocumentParser } from '@docx/ooxml/wordprocessingml/parsing/document-parser';
import { DocumentElement } from '@docx/ooxml/wordprocessingml/document/model/document';

export class DocumentPart extends Part {
    private _documentParser: DocumentParser;

    constructor(pkg: OpenXmlPackage, path: string, parser: DocumentParser) {
        super(pkg, path);
        this._documentParser = parser;
    }
    
    body: DocumentElement

    parseXml(root: Element) {
        this.body = this._documentParser.parseDocumentFile(root);
    }
}