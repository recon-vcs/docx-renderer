// The paginator measures each paragraph's rendered height (via scrollHeight/
// clientHeight) synchronously while appending content, to decide where a page
// overflows. Fonts referenced only by name in CSS (system fonts like Meiryo/
// Yu Gothic/MS Gothic, not backed by an embedded @font-face) are loaded
// lazily by the browser on first use, and that load is asynchronous even when
// the font is installed locally. If pagination measures paragraphs before a
// document's fonts have finished loading, it uses fallback-font metrics -
// systematically smaller than the real ones for CJK "Gothic"/"Mincho" fonts -
// and packs too much content onto early pages. Once the fonts finish loading
// a moment later, the already-decided page split no longer matches the
// (now correctly sized) rendered content.
//
// This waits for every font-family referenced in the generated stylesheet to
// finish loading before pagination begins, so every overflow measurement
// uses final, stable metrics.
export async function waitForDeclaredFonts(styleContainer: HTMLElement): Promise<void> {
	if (typeof document === 'undefined' || !('fonts' in document)) {
		return;
	}

	const cssText = styleContainer.textContent ?? '';
	const fontFamilies = new Set<string>();

	for (const match of cssText.matchAll(/font-family:\s*([^;}]+);/g)) {
		for (const rawName of match[1].split(',')) {
			const name = rawName.trim().replace(/^["']|["']$/g, '');
			if (name && name !== 'inherit' && name !== 'initial') {
				fontFamilies.add(name);
			}
		}
	}

	await Promise.all(
		Array.from(fontFamilies).map(name => document.fonts.load(`16px "${name}"`).catch(() => [] as FontFace[]))
	);

	// Also wait on the document-wide font readiness signal: it covers fonts
	// triggered by @font-face rules (embedded fonts) that don't need an
	// explicit load() call to start fetching.
	await document.fonts.ready;
}
