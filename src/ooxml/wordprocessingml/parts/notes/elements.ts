import { OpenXmlElementBase, DomType, OpenXmlElement } from '@docx/ooxml/wordprocessingml/document/model/dom';

export class WmlNotes<T extends OpenXmlElement = OpenXmlElement> extends OpenXmlElementBase {
	type: DomType;
	children?: T[] = [];
}

export abstract class WmlBaseNote implements OpenXmlElementBase {
	type: DomType;
	id: string;
	noteType: string;
}

export class WmlFootnotes extends WmlNotes<WmlFootnote> {
	type = DomType.Footnotes;
}

export class WmlFootnote extends WmlBaseNote {
	type = DomType.Footnote
}

export class WmlEndnotes extends WmlNotes<WmlEndnote> {
	type = DomType.Endnotes;
}

export class WmlEndnote extends WmlBaseNote {
	type = DomType.Endnote
}
