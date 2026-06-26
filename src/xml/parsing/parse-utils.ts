import { LengthUsage, LengthUsageType, convertLength } from '@docx/ooxml/wordprocessingml/document/model/common';
import xml from '@docx/xml/parsing/xml-parser';

// Auto colors for docx properties
export const autos = {
	shd: "inherit",
	color: "black",
	borderColor: "black",
	highlight: "transparent"
};

const knownColors = [
	'black', 'blue', 'cyan', 'darkBlue', 'darkCyan', 'darkGray', 'darkGreen',
	'darkMagenta', 'darkRed', 'darkYellow', 'green', 'lightGray', 'magenta',
	'none', 'red', 'white', 'yellow'
];

// DOM iteration helpers
export class xmlUtil {
	static foreach(node: Element, callback: (n: Element, i: number) => void) {
		for (let i = 0; i < node.children.length; i++) {
			let n = node.children[i];
			if (n.nodeType == Node.ELEMENT_NODE) {
				callback(n, i);
			}
		}
	}

	static colorAttr(node: Element, attrName: string, defValue: string = null, autoColor: string = 'black') {
		let v = xml.attr(node, attrName);

		if (v) {
			if (v == "auto") {
				return autoColor;
			} else if (knownColors.includes(v)) {
				return v;
			}
			return `#${v}`;
		}

		let themeColor = xml.attr(node, "themeColor");
		return themeColor ? `var(--docx-${themeColor}-color)` : defValue;
	}

	static sizeValue(node: Element, type: LengthUsageType = LengthUsage.Dxa): string {
		return convertLength(node.textContent, type) as string;
	}

	static parseTextContent(node: Element, defaultValue: number = 0): number {
		let textContent: string = node.textContent;
		return textContent ? parseInt(textContent) : defaultValue;
	}
}

// CSS value converters for docx properties
// TODO: some methods overlap with XmlParser methods — consolidate in a later phase
export class values {
	static themeValue(c: Element, attr: string) {
		let val = xml.attr(c, attr);
		return val ? `var(--docx-${val}-font)` : null;
	}

	static valueOfSize(c: Element, attr: string) {
		let type = LengthUsage.Dxa;

		switch (xml.attr(c, "type")) {
			case "dxa":
				break;
			case "pct":
				type = LengthUsage.TablePercent;
				break;
			case "auto":
				return "auto";
		}

		return xml.lengthAttr(c, attr, type);
	}

	static valueOfMargin(c: Element) {
		return xml.lengthAttr(c, "w");
	}

	static valueOfBorder(c: Element) {
		let type = xml.attr(c, "val");

		if (type == "nil") {
			return "none";
		}

		let color = xmlUtil.colorAttr(c, "color");
		let size = xml.lengthAttr(c, "sz", LengthUsage.Border);

		return `${size} solid ${color == "auto" ? autos.borderColor : color}`;
	}

	static valueOfTblLayout(c: Element) {
		let type = xml.attr(c, "type");
		return type == "fixed" ? "fixed" : "auto";
	}

	static classNameOfCnfStyle(c: Element) {
		const val = xml.attr(c, "val");
		const classes = [
			'first-row', 'last-row', 'first-col', 'last-col',
			'odd-col', 'even-col', 'odd-row', 'even-row',
			'ne-cell', 'nw-cell', 'se-cell', 'sw-cell'
		];
		return classes.filter((_, i) => val[i] == '1').join(' ');
	}

	static valueOfJc(c: Element) {
		let type = xml.attr(c, "val");

		switch (type) {
			case "start":
			case "left":
				return "left";
			case "center":
				return "center";
			case "end":
			case "right":
				return "right";
			case "both":
				return "justify";
		}

		return type;
	}

	static valueOfVertAlign(c: Element, asTagName: boolean = false) {
		let type = xml.attr(c, "val");

		switch (type) {
			case "subscript":
				return "sub";
			case "superscript":
				return asTagName ? "sup" : "super";
		}

		return asTagName ? null : type;
	}

	static valueOfTextAlignment(c: Element) {
		let type = xml.attr(c, "val");

		switch (type) {
			case "auto":
			case "baseline":
				return "baseline";
			case "top":
				return "top";
			case "center":
				return "middle";
			case "bottom":
				return "bottom";
		}

		return type;
	}

	static addSize(a: string, b: string): string {
		if (a == null) return b;
		if (b == null) return a;
		return `calc(${a} + ${b})`; // TODO
	}

	static classNameOftblLook(c: Element) {
		const val = xml.hexAttr(c, "val", 0);
		let className = "";

		if (xml.boolAttr(c, "firstRow") || (val & 0x0020)) className += " first-row";
		if (xml.boolAttr(c, "lastRow") || (val & 0x0040)) className += " last-row";
		if (xml.boolAttr(c, "firstColumn") || (val & 0x0080)) className += " first-col";
		if (xml.boolAttr(c, "lastColumn") || (val & 0x0100)) className += " last-col";
		if (xml.boolAttr(c, "noHBand") || (val & 0x0200)) className += " no-hband";
		if (xml.boolAttr(c, "noVBand") || (val & 0x0400)) className += " no-vband";

		return className.trim();
	}
}
