import { OpenXmlPackage } from '@docx/opc/package/open-xml-package';
import { Part } from '@docx/opc/parts/part';
import { DocumentParser } from '@docx/ooxml/wordprocessingml/parsing/document-parser';
import { IDomStyle } from '@docx/ooxml/wordprocessingml/document/model/style';

export class StylesPart extends Part {
    styles: IDomStyle[];

    private _documentParser: DocumentParser;

    constructor(pkg: OpenXmlPackage, path: string, parser: DocumentParser) {
        super(pkg, path);
        this._documentParser = parser;
    }

    parseXml(root: Element) {
        this.styles = this._documentParser.parseStylesFile(root);     
    }
}