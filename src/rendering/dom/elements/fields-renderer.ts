import { WmlBookmarkStart } from '@docx/ooxml/wordprocessingml/document/model/bookmarks';
import { DomType, OpenXmlElement, WmlCharacter, WmlText } from '@docx/ooxml/wordprocessingml/document/model/dom';
import { WmlFieldChar, WmlFieldSimple, WmlInstructionText } from '@docx/ooxml/wordprocessingml/document/model/fields';
import { WmlRun } from '@docx/ooxml/wordprocessingml/document/model/run';
import { WmlCommentRangeStart, WmlCommentReference } from '@docx/ooxml/wordprocessingml/parts/comments/elements';
import { ChildrenType, createElement } from '@docx/rendering/dom/core/dom-utils';
import { Overflow } from '@docx/rendering/measurement/overflow';

export interface FieldsRendererCallbacks {
	processElement(elem: OpenXmlElement): void;
	renderChildren(elem: OpenXmlElement, parent: HTMLElement): Promise<Overflow>;
	appendChildren(parent: HTMLElement | Text, children: ChildrenType): Promise<Overflow>;
	appendChildrenWithoutOverflow(parent: Element | Text, children: ChildrenType): void;
	renderText(elem: WmlText, parent: HTMLElement): Promise<Node>;
	findComment(id: string): { id: string; author: string; date: string } | undefined;
	currentPageNumber(): number;
	pageCount(): number;
	renderChanges(): boolean;
}

export function resolveFieldRuns(
	runs: OpenXmlElement[],
	callbacks: FieldsRendererCallbacks
): OpenXmlElement[] {
	const pageNumber = callbacks.currentPageNumber();
	const totalPages = callbacks.pageCount();
	const result: OpenXmlElement[] = [];
	let i = 0;

	while (i < runs.length) {
		const run = runs[i] as WmlRun;
		const beginChar = run.type === DomType.Run
			&& run.children?.find(c => c.type === DomType.ComplexField) as WmlFieldChar;

		if (!beginChar || beginChar.charType !== 'begin') {
			result.push(run);
			i++;
			continue;
		}

		let j = i + 1;
		let instruction = '';
		while (j < runs.length) {
			const childRun = runs[j] as WmlRun;
			const fieldChar = childRun.children?.find(c => c.type === DomType.ComplexField) as WmlFieldChar;
			const instr = childRun.children?.find(c => c.type === DomType.Instruction) as WmlInstructionText;
			if (instr) {
				instruction += instr.text;
			}
			j++;
			if (fieldChar?.charType === 'separate') {
				break;
			}
		}

		const resultStart = j;
		while (j < runs.length) {
			const fieldChar = (runs[j] as WmlRun).children?.find(c => c.type === DomType.ComplexField) as WmlFieldChar;
			if (fieldChar?.charType === 'end') {
				break;
			}
			j++;
		}
		const endIndex = j;
		const fieldName = instruction.trim().split(/\s+/)[0]?.toUpperCase();

		if (fieldName === 'PAGE' || fieldName === 'NUMPAGES') {
			const value = String(fieldName === 'PAGE' ? pageNumber : totalPages);
			const template = (runs[resultStart] as WmlRun) ?? run;
			const replacement = createFieldReplacement(value, template, run.parent);
			callbacks.processElement(replacement);
			result.push(replacement);
		} else {
			result.push(...runs.slice(i, endIndex + 1));
		}
		i = endIndex + 1;
	}

	return result;
}

export function resolveSimpleField(
	elem: WmlFieldSimple,
	callbacks: FieldsRendererCallbacks
): OpenXmlElement[] {
	const pageNumber = callbacks.currentPageNumber();
	const totalPages = callbacks.pageCount();
	const fieldName = elem.instruction?.trim().split(/\s+/)[0]?.toUpperCase();

	if (fieldName !== 'PAGE' && fieldName !== 'NUMPAGES') {
		return elem.children ?? [];
	}

	const value = String(fieldName === 'PAGE' ? pageNumber : totalPages);
	const template = elem.children?.[0] as WmlRun;
	const replacement = createFieldReplacement(value, template, elem.parent);
	callbacks.processElement(replacement);
	return [replacement];
}

export function renderBookmarkStart(
	elem: WmlBookmarkStart,
	parent: HTMLElement,
	callbacks: FieldsRendererCallbacks
): HTMLElement {
	const oSpan = createElement('span');
	oSpan.id = elem.name;
	callbacks.appendChildrenWithoutOverflow(parent, oSpan);
	oSpan.dataset.overflow = Overflow.IGNORE;

	return oSpan;
}

export async function renderInserted(
	elem: OpenXmlElement,
	parent: HTMLElement,
	callbacks: FieldsRendererCallbacks
): Promise<HTMLModElement | HTMLSpanElement> {
	const tagName: keyof HTMLElementTagNameMap = callbacks.renderChanges() ? 'ins' : 'span';
	const oInserted = createElement(tagName) as HTMLModElement | HTMLSpanElement;
	let isOverflow = await callbacks.appendChildren(parent, oInserted);

	if (isOverflow === Overflow.TRUE) {
		oInserted.dataset.overflow = Overflow.SELF;

		return oInserted;
	}

	oInserted.dataset.overflow = await callbacks.renderChildren(elem, oInserted);

	return oInserted;
}

export async function renderDeleted(
	elem: OpenXmlElement,
	parent: HTMLElement,
	callbacks: FieldsRendererCallbacks
): Promise<HTMLModElement> {
	const oDeleted = createElement('del');

	if (callbacks.renderChanges() === false) {
		oDeleted.style.display = 'none';
	}

	let isOverflow = await callbacks.appendChildren(parent, oDeleted);

	if (isOverflow === Overflow.TRUE) {
		oDeleted.dataset.overflow = Overflow.SELF;

		return oDeleted;
	}

	oDeleted.dataset.overflow = await callbacks.renderChildren(elem, oDeleted);

	return oDeleted;
}

export async function renderDeletedText(
	elem: WmlText,
	parent: HTMLElement,
	callbacks: FieldsRendererCallbacks
): Promise<Node> {
	return callbacks.renderText(elem, parent);
}

export function renderCommentRangeStart(commentStart: WmlCommentRangeStart): Comment {
	return document.createComment(`start of comment #${commentStart.id}`);
}

export function renderCommentRangeEnd(commentEnd: WmlCommentRangeStart): Comment {
	return document.createComment(`end of comment #${commentEnd.id}`);
}

export function renderCommentReference(
	commentRef: WmlCommentReference,
	callbacks: FieldsRendererCallbacks
): Comment {
	const comment = callbacks.findComment(commentRef.id);

	if (!comment) return null;

	return document.createComment(
		`comment #${comment.id} by ${comment.author} on ${comment.date}`
	);
}

function createFieldReplacement(
	value: string,
	template: WmlRun,
	parent: OpenXmlElement
): WmlRun {
	return {
		type: DomType.Run,
		cssStyle: { ...template?.cssStyle },
		verticalAlign: template?.verticalAlign,
		parent,
		fieldRun: false,
		children: [{
			type: DomType.Text,
			text: value,
			children: [{ type: DomType.Character, char: value } as WmlCharacter],
		} as WmlText],
	} as WmlRun;
}
