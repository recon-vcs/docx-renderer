import { DomType, OpenXmlElement, WmlBreak, WmlCharacter, WmlLastRenderedPageBreak, WmlNoteReference, WmlSymbol, WmlText } from '@docx/ooxml/wordprocessingml/document/model/dom';
import type { WmlRun } from '@docx/ooxml/wordprocessingml/document/model/run';
import { WmlFieldChar, WmlFieldSimple } from '@docx/ooxml/wordprocessingml/document/model/fields';
import { WmlCommentReference } from '@docx/ooxml/wordprocessingml/parts/comments/elements';
import xml from '@docx/xml/parsing/xml-parser';
import { xmlUtil, values } from '@docx/xml/parsing/parse-utils';
import { parseDefaultProperties, parseSpacing } from './properties-parser';
import type { RunParserContext } from './parse-context';

export function parseRun(
	node: Element,
	ctx: RunParserContext,
): WmlRun {
	const wmlRun: WmlRun = {
		type: DomType.Run,
		children: [],
	};

	xmlUtil.foreach(node, (child) => {
		// check for alternate content
		child = ctx.checkAlternateContent(child);

		switch (child.localName) {
			// property
			case "rPr":
				parseRunProperties(child, wmlRun, ctx);
				break;

			case "t":
				wmlRun.children.push(parseText(child, DomType.Text));
				break;

			case "delText":
				wmlRun.children.push(parseText(child, DomType.DeletedText));
				break;

			case "commentReference":
				wmlRun.children.push(new WmlCommentReference(xml.attr(child, "id")));
				break;

			case "fldSimple":
				wmlRun.children.push(<WmlFieldSimple>{
					type: DomType.SimpleField,
					instruction: xml.attr(child, "instr"),
					lock: xml.boolAttr(child, "lock", false),
					dirty: xml.boolAttr(child, "dirty", false)
				});
				break;

			case "instrText":
				wmlRun.fieldRun = true;
				wmlRun.children.push(parseText(child, DomType.Instruction));
				break;

			case "fldChar":
				wmlRun.fieldRun = true;
				wmlRun.children.push(<WmlFieldChar>{
					type: DomType.ComplexField,
					charType: xml.attr(child, "fldCharType"),
					lock: xml.boolAttr(child, "lock", false),
					dirty: xml.boolAttr(child, "dirty", false)
				});
				break;

			case "noBreakHyphen":
				wmlRun.children.push({ type: DomType.NoBreakHyphen });
				break;

			case "br":
				wmlRun.children.push(<WmlBreak>{
					type: DomType.Break,
					break: xml.attr(child, "type") || "textWrapping",
					props: {
						clear: xml.attr(child, "clear")
					}
				});
				break;

			case "lastRenderedPageBreak":
				wmlRun.children.push(<WmlLastRenderedPageBreak>{
					type: DomType.LastRenderedPageBreak,
				});
				break;

			// SymbolChar
			case "sym":
				wmlRun.children.push(<WmlSymbol>{
					type: DomType.Symbol,
					font: xml.attr(child, "font"),
					char: xml.attr(child, "char")
				});
				break;

			// TODO PositionalTab
			case "ptab":
				break;

			case "tab":
				wmlRun.children.push({ type: DomType.Tab });
				break;

			case "footnoteReference":
				wmlRun.children.push(<WmlNoteReference>{
					type: DomType.FootnoteReference,
					id: xml.attr(child, "id")
				});
				break;

			case "endnoteReference":
				wmlRun.children.push(<WmlNoteReference>{
					type: DomType.EndnoteReference,
					id: xml.attr(child, "id")
				});
				break;

			case "drawing":
				wmlRun.children.push(ctx.parseDrawing(child));
				break;

			case "pict":
				wmlRun.children.push(ctx.parseVmlPicture(child));
				break;

			default:
				if (ctx.options.debug) {
					console.warn(`DOCX:%c Unknown Run Element：${child.localName}`, 'color:#f75607');
				}
		}
	});

	return wmlRun;
}

export function parseText(elem: Element, type: DomType): WmlText {
	const wmlText = { type, text: '' } as WmlText;
	let textContent = elem.textContent;
	// preserve whitespace
	const is_preserve_space = xml.attr(elem, "xml:space") === "preserve";
	if (is_preserve_space) {
		//   = non-breaking space
		textContent = textContent.split(/\s/).join(" ");
	}
	wmlText.text = textContent;
	if (textContent.length > 0) {
		wmlText.children = parseCharacter(textContent);
	}
	return wmlText;
}

export function parseCharacter(text: string): OpenXmlElement[] {
	let characters: string[];
	// check whether string is primarily Chinese characters
	const isChinese = text.match(/[一-龥]+/g);
	if (isChinese) {
		characters = text.split('');
	} else {
		characters = text.match(/\S+|\s+/g);
	}
	return characters.map(character => (
		{ type: DomType.Character, char: character } as WmlCharacter
	));
}

export function parseRunProperties(
	elem: Element,
	run: WmlRun,
	ctx: RunParserContext,
): void {
	parseDefaultProperties(elem, ctx.options, run.cssStyle = {}, null, c => {
		switch (c.localName) {
			// Referenced Character Style
			case "rStyle":
				run.styleName = xml.attr(c, "val");
				break;

			// Subscript/Superscript Text
			case "vertAlign":
				run.verticalAlign = values.valueOfVertAlign(c, true);
				break;

			// Character Spacing Adjustment
			case "spacing":
				parseSpacing(c, run, ctx.options);
				break;

			default:
				return false;
		}

		return true;
	});
}
