import { DomType, OpenXmlElement } from '../model/element';
import { WmlNoteReference } from '../model/element';
import { WmlBaseNote, WmlFootnotes, WmlEndnotes } from '../notes/elements';
import { createElement } from './dom-utils';

// Callbacks to the main renderer for notes rendering
export interface NotesRendererCallbacks {
	processElement(elem: OpenXmlElement): void;
	renderChildren(elem: OpenXmlElement, parent: HTMLElement): Promise<string>;
}

// Render a footnotes or endnotes list into the page element
export async function renderNotes(
	type: DomType,
	noteIds: string[],
	notesMap: Record<string, WmlBaseNote>,
	parent: HTMLElement,
	cbs: NotesRendererCallbacks
): Promise<void> {
	// Gather only the notes that appear on this page
	const children: WmlBaseNote[] = noteIds.map(id => notesMap[id]).filter(x => x);
	if (children.length > 0) {
		const oList = createElement('ol', null);
		// Build a synthetic container element
		const notes = type === DomType.Footnotes ? new WmlFootnotes() : new WmlEndnotes();
		notes.children = children;
		// Re-establish parent relationships for the notes tree
		cbs.processElement(notes);
		// Render each note as an <li>
		await cbs.renderChildren(notes, oList);
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
