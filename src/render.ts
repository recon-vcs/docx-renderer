import { DocumentParser } from '@docx/ooxml/wordprocessingml/parsing/document-parser';
import { HtmlRendererSync } from '@docx/html-renderer-sync';
import { resolveOptions, type DocumentSource, type Options } from '@docx/options';
import { createRenderResult, RenderResult } from '@docx/render-result';
import { WordDocument } from '@docx/word-document';

/** Parses a docx file into a WordDocument model without rendering it. */
export function parseAsync(data: DocumentSource, userOptions: Partial<Options> | null = null): Promise<WordDocument> {
	const ops = resolveOptions(userOptions);
	return WordDocument.load(data, new DocumentParser(ops), ops);
}

/** Renders an already-parsed WordDocument into the given containers. */
export async function renderDocument(
	document: WordDocument,
	bodyContainer: HTMLElement,
	styleContainer?: HTMLElement | null,
	userOptions?: Partial<Options> | null,
): Promise<RenderResult> {
	const ops = resolveOptions(userOptions);
	const renderer = new HtmlRendererSync();
	await renderer.render(document, bodyContainer, styleContainer ?? undefined, ops);
	return createRenderResult(document, bodyContainer, ops.className);
}

/** Parses and renders with the pagination-aware renderer. */
export async function renderSync(
	data: DocumentSource,
	bodyContainer: HTMLElement,
	styleContainer: HTMLElement | null = null,
	userOptions: Partial<Options> | null = null,
): Promise<RenderResult> {
	const doc = await parseAsync(data, userOptions);
	return renderDocument(doc, bodyContainer, styleContainer, userOptions);
}
