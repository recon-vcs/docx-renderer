import { OpenXmlPackage } from '@docx/opc/package/open-xml-package';
import { Part } from '@docx/opc/parts/part';
import { DmlTheme, parseTheme } from '@docx/ooxml/drawingml/theme/theme';

export class ThemePart extends Part {
    theme: DmlTheme;

    constructor(pkg: OpenXmlPackage, path: string) {
        super(pkg, path);
    }

    parseXml(root: Element) {
        this.theme = parseTheme(root, this._package.xmlParser);
    }
}