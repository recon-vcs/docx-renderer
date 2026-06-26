import { WmlComment } from '@docx/ooxml/wordprocessingml/parts/comments/elements';
import { OpenXmlElement } from '@docx/ooxml/wordprocessingml/document/model/dom';
import xml from '@docx/xml/parsing/xml-parser';
import type { BodyParserContext } from './parse-context';

type NoteConstructor<T> = new () => T;

export function parseSdt(
	node: Element,
	ctx: BodyParserContext
): OpenXmlElement[] {
	const sdtContent = xml.element(node, "sdtContent");
	return sdtContent ? ctx.parseBodyElements(sdtContent) : [];
}

export function parseNotes<T extends { id?: string; noteType?: string; children?: OpenXmlElement[] }>(
	xmlDoc: Element,
	elemName: string,
	elemClass: NoteConstructor<T>,
	ctx: BodyParserContext
): T[] {
	const result: T[] = [];

	for (const el of xml.elements(xmlDoc, elemName)) {
		const node = new elemClass();
		node.id = xml.attr(el, "id");
		node.noteType = xml.attr(el, "type");
		node.children = ctx.parseBodyElements(el);
		result.push(node);
	}

	return result;
}

export function parseComments(
	xmlDoc: Element,
	ctx: BodyParserContext
): WmlComment[] {
	const result: WmlComment[] = [];

	for (const el of xml.elements(xmlDoc, "comment")) {
		const item = new WmlComment();
		item.id = xml.attr(el, "id");
		item.author = xml.attr(el, "author");
		item.initials = xml.attr(el, "initials");
		item.date = xml.attr(el, "date");
		item.children = ctx.parseBodyElements(el);
		result.push(item);
	}

	return result;
}
