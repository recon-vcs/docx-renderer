import { OpenXmlElement } from "./dom";
import { SectionProperties } from "./section";

export interface DocumentElement extends OpenXmlElement {
	sectProps: SectionProperties;
}
