import { OpenXmlPackage } from '@docx/opc/package/open-xml-package';
import { Part } from '@docx/opc/parts/part';
import { WmlSettings, parseSettings } from "./settings";

export class SettingsPart extends Part {
	settings: WmlSettings;

	constructor(pkg: OpenXmlPackage, path: string) {
		super(pkg, path);
	}

	parseXml(root: Element) {
		this.settings = parseSettings(root, this._package.xmlParser);
	}
}