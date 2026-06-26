import { IDomNumbering, NumberingPicBullet } from '@docx/ooxml/wordprocessingml/model/numbering-types';
import xml from '@docx/xml/parsing/xml-parser';
import { xmlUtil } from '@docx/xml/parsing/parse-utils';
import type { ParseContext } from './parse-context';

export function parseNumberingFile(
	xnums: Element,
	ctx: ParseContext
): IDomNumbering[] {
	const result: IDomNumbering[] = [];
	const mapping: Record<string, string> = {};
	const bullets: NumberingPicBullet[] = [];

	xmlUtil.foreach(xnums, n => {
		switch (n.localName) {
			case "abstractNum":
				parseAbstractNumbering(n, bullets, ctx)
					.forEach(x => result.push(x));
				break;

			case "numPicBullet":
				bullets.push(parseNumberingPicBullet(n));
				break;

			case "num": {
				const numId = xml.attr(n, "numId");
				const abstractNumId = xml.elementAttr(n, "abstractNumId", "val");
				mapping[abstractNumId] = numId;
				break;
			}

			default:
				if (ctx.options.debug) {
					console.warn(`DOCX:%c Unknown Numbering File：${n.localName}`, 'color:#f75607');
				}
		}
	});

	result.forEach(x => x.id = mapping[x.id]);

	return result;
}

export function parseNumberingPicBullet(elem: Element): NumberingPicBullet {
	const pict = xml.element(elem, "pict");
	const shape = pict && xml.element(pict, "shape");
	const imagedata = shape && xml.element(shape, "imagedata");

	return imagedata ? {
		id: xml.intAttr(elem, "numPicBulletId"),
		src: xml.attr(imagedata, "id"),
		style: xml.attr(shape, "style")
	} : null;
}

export function parseAbstractNumbering(
	node: Element,
	bullets: NumberingPicBullet[],
	ctx: ParseContext
): IDomNumbering[] {
	const result: IDomNumbering[] = [];
	const id = xml.attr(node, "abstractNumId");

	xmlUtil.foreach(node, n => {
		switch (n.localName) {
			case "lvl":
				result.push(parseNumberingLevel(id, n, bullets, ctx));
				break;

			default:
				if (ctx.options.debug) {
					console.warn(`DOCX:%c Unknown Abstract Numbering：${n.localName}`, 'color:#f75607');
				}
		}
	});

	return result;
}

export function parseNumberingLevel(
	id: string,
	node: Element,
	bullets: NumberingPicBullet[],
	ctx: ParseContext
): IDomNumbering {
	const result: IDomNumbering = {
		id: id,
		level: xml.intAttr(node, "ilvl"),
		start: 1,
		pStyleName: undefined,
		pStyle: {},
		rStyle: {},
		suff: "tab"
	};

	xmlUtil.foreach(node, n => {
		switch (n.localName) {
			case "start":
				result.start = xml.intAttr(n, "val");
				break;

			case "pPr":
				ctx.parseDefaultProperties(n, result.pStyle);
				break;

			case "rPr":
				ctx.parseDefaultProperties(n, result.rStyle);
				break;

			case "lvlPicBulletId": {
				const picId = xml.intAttr(n, "val");
				result.bullet = bullets.find(x => x?.id == picId);
				break;
			}

			case "lvlText":
				result.levelText = xml.attr(n, "val");
				break;

			case "pStyle":
				result.pStyleName = xml.attr(n, "val");
				break;

			case "numFmt":
				result.format = xml.attr(n, "val");
				break;

			case "suff":
				result.suff = xml.attr(n, "val");
				break;

			default:
				if (ctx.options.debug) {
					console.warn(`DOCX:%c Unknown Numbering Level：${n.localName}`, 'color:#f75607');
				}
		}
	});

	return result;
}
