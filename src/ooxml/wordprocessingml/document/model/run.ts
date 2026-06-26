import { XmlParser } from '@docx/xml/parsing/xml-parser';
import { CommonProperties, parseCommonProperty } from "./common";
import { OpenXmlElement } from "./dom";

export interface WmlRun extends OpenXmlElement, RunProperties {
    verticalAlign?: string;
	fieldRun?: boolean;
}

export interface RunProperties extends CommonProperties {

}

export function parseRunProperties(elem: Element, xml: XmlParser): RunProperties {
    let result = <RunProperties>{};

    for(let el of xml.elements(elem)) {
        parseRunProperty(el, result, xml);
    }

    return result;
}

export function parseRunProperty(elem: Element, props: RunProperties, xml: XmlParser) {
    return parseCommonProperty(elem, props, xml);
}
