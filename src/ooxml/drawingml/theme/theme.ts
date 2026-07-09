import { XmlParser } from '@docx/xml/parsing/xml-parser';

export class DmlTheme {
    colorScheme: DmlColorScheme;
    fontScheme: DmlFontScheme;
}

export interface DmlColorScheme {
    name: string;
    colors: Record<string, string>;
}

export interface DmlFontScheme {
    name: string;
    majorFont: DmlFormInfo,
    minorFont: DmlFormInfo
}

export interface DmlFormInfo {
    latinTypeface: string;
    eaTypeface: string;
    csTypeface: string;
    // <a:font script="Jpan" typeface="..."/> entries: per-script faces that
    // substitute for an empty latin/ea/cs typeface, selected through the
    // document's themeFontLang.
    scriptTypefaces: Record<string, string>;
}

export function parseTheme(elem: Element, xml: XmlParser) {
    var result = new DmlTheme();
    var themeElements = xml.element(elem, "themeElements");

    for (let el of xml.elements(themeElements)) {
        switch(el.localName) {
            case "clrScheme": result.colorScheme = parseColorScheme(el, xml); break;
            case "fontScheme": result.fontScheme = parseFontScheme(el, xml); break;
        }
    }

    return result;
}

export function parseColorScheme(elem: Element, xml: XmlParser) {
    var result: DmlColorScheme = { 
        name: xml.attr(elem, "name"),
        colors: {}
    };

    for (let el of xml.elements(elem)) {
        var srgbClr = xml.element(el, "srgbClr");
        var sysClr = xml.element(el, "sysClr");

        if (srgbClr) {
            result.colors[el.localName] = xml.attr(srgbClr, "val");
        }
        else if (sysClr) {
            result.colors[el.localName] = xml.attr(sysClr, "lastClr");
        }
    }

    return result;
}

export function parseFontScheme(elem: Element, xml: XmlParser) {
    var result: DmlFontScheme = { 
        name: xml.attr(elem, "name"),
    } as DmlFontScheme;

    for (let el of xml.elements(elem)) {
        switch (el.localName) {
            case "majorFont": result.majorFont = parseFontInfo(el, xml); break;
            case "minorFont": result.minorFont = parseFontInfo(el, xml); break;
        }
    }

    return result;
}

export function parseFontInfo(elem: Element, xml: XmlParser): DmlFormInfo {
    const scriptTypefaces: Record<string, string> = {};

    for (const el of xml.elements(elem, "font")) {
        const script = xml.attr(el, "script");
        const typeface = xml.attr(el, "typeface");
        if (script && typeface) {
            scriptTypefaces[script] = typeface;
        }
    }

    return {
        latinTypeface: xml.elementAttr(elem, "latin", "typeface"),
        eaTypeface: xml.elementAttr(elem, "ea", "typeface"),
        csTypeface: xml.elementAttr(elem, "cs", "typeface"),
        scriptTypefaces,
    };
}