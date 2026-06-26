import { OpenXmlPackage } from '@docx/opc/package/open-xml-package';
import { Part } from '@docx/opc/parts/part';
import { DocumentParser } from '@docx/ooxml/wordprocessingml/parsing/document-parser';
import { IDomNumbering } from '@docx/ooxml/wordprocessingml/document/model/dom';
import { AbstractNumbering, Numbering, NumberingBulletPicture, NumberingPartProperties, parseNumberingPart } from "./numbering";

export class NumberingPart extends Part implements NumberingPartProperties {
    private _documentParser: DocumentParser;

    constructor(pkg: OpenXmlPackage, path: string, parser: DocumentParser) {
        super(pkg, path);
        this._documentParser = parser;
    }

    numberings: Numbering[];
    abstractNumberings: AbstractNumbering[];
    bulletPictures: NumberingBulletPicture[];
    
    domNumberings: IDomNumbering[];

    parseXml(root: Element) {
        Object.assign(this, parseNumberingPart(root, this._package.xmlParser));
        this.domNumberings = this._documentParser.parseNumberingFile(root);  
    }
}