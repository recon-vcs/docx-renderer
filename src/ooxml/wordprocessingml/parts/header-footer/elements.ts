import { OpenXmlElementBase, DomType } from '@docx/ooxml/wordprocessingml/document/model/dom';

export class WmlHeader extends OpenXmlElementBase {
    type: DomType = DomType.Header;
}

export class WmlFooter extends OpenXmlElementBase {
    type: DomType = DomType.Footer;
}