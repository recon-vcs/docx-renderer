import { OpenXmlElement } from './element';

export interface WmlDrawing extends OpenXmlElement {}

export interface WmlImage extends OpenXmlElement {
	src: string;
}
