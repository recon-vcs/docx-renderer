import { OpenXmlElement } from '@docx/ooxml/wordprocessingml/model/element';

export interface WmlDrawing extends OpenXmlElement {}

export interface WmlImage extends OpenXmlElement {
	src: string;
}
