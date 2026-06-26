import { DomType, OpenXmlElement, WmlNoteReference } from '@docx/ooxml/wordprocessingml/document/model/dom';
import { WmlBaseNote, WmlFootnotes, WmlEndnotes } from '@docx/ooxml/wordprocessingml/parts/notes/elements';
import { createElement } from '@docx/rendering/dom/core/dom-utils';
import type { RenderContext } from '@docx/rendering/render-context';

// Render a footnotes or endnotes list into the page element
export async function renderNotes(
	type: DomType,
	noteIds: string[],
	notesMap: Record<string, WmlBaseNote>,
	parent: HTMLElement,
	ctx: RenderContext
): Promise<void> {
	// Gather only the notes that appear on this page
	const children: WmlBaseNote[] = noteIds.map(id => notesMap[id]).filter(x => x);
	if (children.length > 0) {
		const oList = createElement('ol', null);
		// Build a synthetic container element
		const notes = type === DomType.Footnotes ? new WmlFootnotes() : new WmlEndnotes();
		notes.children = children;
		// Establish parent links and apply table style propagation
		ctx.linkParents(notes);
		ctx.processElement(notes);
		// Render each note as an <li>
		await ctx.renderChildren(notes, oList);
		parent.appendChild(oList);
	}
}

// Render a footnote reference marker (superscript counter)
export function renderFootnoteReference(
	elem: WmlNoteReference,
	currentFootnoteIds: string[]
): HTMLElement {
	const oSup = createElement('sup');
	currentFootnoteIds.push(elem.id);
	oSup.textContent = `${currentFootnoteIds.length}`;
	return oSup;
}

// Render an endnote reference marker (superscript counter)
export function renderEndnoteReference(
	elem: WmlNoteReference,
	currentEndnoteIds: string[]
): HTMLElement {
	const oSup = createElement('sup');
	currentEndnoteIds.push(elem.id);
	oSup.textContent = `${currentEndnoteIds.length}`;
	return oSup;
}
