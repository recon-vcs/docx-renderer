import { WmlComment } from '../comments/elements';
import { OpenXmlElement } from '../document/dom';
import xml from './xml-parser';

export interface ContentParserCallbacks {
	parseBodyElements(node: Element): OpenXmlElement[];
}

type NoteConstructor<T> = new () => T;

export function parseSdt(
	node: Element,
	callbacks: ContentParserCallbacks
): OpenXmlElement[] {
	const sdtContent = xml.element(node, "sdtContent");
	return sdtContent ? callbacks.parseBodyElements(sdtContent) : [];
}

export function parseNotes<T extends { id?: string; noteType?: string; children?: OpenXmlElement[] }>(
	xmlDoc: Element,
	elemName: string,
	elemClass: NoteConstructor<T>,
	callbacks: ContentParserCallbacks
): T[] {
	const result: T[] = [];

	for (const el of xml.elements(xmlDoc, elemName)) {
		const node = new elemClass();
		node.id = xml.attr(el, "id");
		node.noteType = xml.attr(el, "type");
		node.children = callbacks.parseBodyElements(el);
		result.push(node);
	}

	return result;
}

export function parseComments(
	xmlDoc: Element,
	callbacks: ContentParserCallbacks
): WmlComment[] {
	const result: WmlComment[] = [];

	for (const el of xml.elements(xmlDoc, "comment")) {
		const item = new WmlComment();
		item.id = xml.attr(el, "id");
		item.author = xml.attr(el, "author");
		item.initials = xml.attr(el, "initials");
		item.date = xml.attr(el, "date");
		item.children = callbacks.parseBodyElements(el);
		result.push(item);
	}

	return result;
}
