import { DomType, OpenXmlElement } from '../document/dom';
import type { WmlRun } from '../document/run';
import xml from './xml-parser';
import { xmlUtil } from './parse-utils';
import type { DocumentParserOptions } from '../document-parser';

const mmlTagMap: Record<string, DomType> = {
	"oMath": DomType.MmlMath,
	"oMathPara": DomType.MmlMathParagraph,
	"f": DomType.MmlFraction,
	"func": DomType.MmlFunction,
	"fName": DomType.MmlFunctionName,
	"num": DomType.MmlNumerator,
	"den": DomType.MmlDenominator,
	"rad": DomType.MmlRadical,
	"deg": DomType.MmlDegree,
	"e": DomType.MmlBase,
	"sSup": DomType.MmlSuperscript,
	"sSub": DomType.MmlSubscript,
	"sPre": DomType.MmlPreSubSuper,
	"sup": DomType.MmlSuperArgument,
	"sub": DomType.MmlSubArgument,
	"d": DomType.MmlDelimiter,
	"nary": DomType.MmlNary,
	"eqArr": DomType.MmlEquationArray,
	"lim": DomType.MmlLimit,
	"limLow": DomType.MmlLimitLower,
	"m": DomType.MmlMatrix,
	"mr": DomType.MmlMatrixRow,
	"box": DomType.MmlBox,
	"bar": DomType.MmlBar,
	"groupChr": DomType.MmlGroupChar,
};

export interface MathParserCallbacks {
	parseRun(node: Element): WmlRun;
}

export function parseMathElement(
	elem: Element,
	options: DocumentParserOptions,
	callbacks: MathParserCallbacks
): OpenXmlElement {
	const propsTag = `${elem.localName}Pr`;
	const mathElement: OpenXmlElement = {
		type: mmlTagMap[elem.localName],
		children: [],
	};

	xmlUtil.foreach(elem, (child) => {
		const childType = mmlTagMap[child.localName];

		if (childType) {
			mathElement.children.push(parseMathElement(child, options, callbacks));
		} else if (child.localName == "r") {
			let wmlRun: WmlRun = callbacks.parseRun(child);
			wmlRun.type = DomType.MmlRun;
			mathElement.children.push(wmlRun);
		} else if (child.localName == propsTag) {
			mathElement.props = parseMathProperties(child, options);
		}
	});

	return mathElement;
}

export function parseMathProperties(
	elem: Element,
	options: DocumentParserOptions
): Record<string, any> {
	const result: Record<string, any> = {};

	for (const el of xml.elements(elem)) {
		switch (el.localName) {
			case "chr":
				result.char = xml.attr(el, "val");
				break;

			case "vertJc":
				result.verticalJustification = xml.attr(el, "val");
				break;

			case "jc":
				result.justification = xml.attr(el, "val");
				break;

			case "pos":
				result.position = xml.attr(el, "val");
				break;

			case "degHide":
				result.hideDegree = xml.boolAttr(el, "val");
				break;

			case "begChr":
				result.beginChar = xml.attr(el, "val");
				break;

			case "endChr":
				result.endChar = xml.attr(el, "val");
				break;

			default:
				if (options.debug) {
					console.warn(`DOCX:%c Unknown Math Property：${el.localName}`, 'color:#f75607');
				}
		}
	}

	return result;
}
