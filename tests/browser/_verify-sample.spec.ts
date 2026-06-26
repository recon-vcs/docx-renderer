import { expect, test } from '@playwright/test';

interface DocxGlobal {
	renderSync(data: Blob, bodyContainer: HTMLElement, styleContainer: HTMLElement, options?: Record<string, unknown>): Promise<unknown>;
}

type HarnessWindow = Window & typeof globalThis & { docx: DocxGlobal };

test('verify zz-sample-analyze rendering after fixes', async ({ page }) => {
	const pageErrors: Error[] = [];
	const consoleErrors: string[] = [];
	page.on('pageerror', (error) => pageErrors.push(error));
	page.on('console', (msg) => {
		if (msg.type() === 'error') consoleErrors.push(msg.text());
	});

	await page.goto('/tests/browser/harness.html');
	const result = await page.evaluate(async () => {
		const response = await fetch('/tests/fixtures/zz-sample-analyze.docx');
		const blob = await response.blob();
		const body = document.querySelector<HTMLElement>('#document-container')!;
		const styleContainer = document.querySelector<HTMLElement>('#style-container')!;
		const docx = (window as unknown as HarnessWindow).docx;
		await docx.renderSync(blob, body, styleContainer, { breakPages: true });
		const sections = Array.from(document.querySelectorAll('section.docx-wrapper, section[class*="docx"]'));
		const mathParagraphs = Array.from(document.querySelectorAll<HTMLElement>('.docx-math-paragraph'));
		const mathElements = Array.from(document.querySelectorAll<HTMLElement>('math'));
		const images = Array.from(document.querySelectorAll<HTMLImageElement>('img'));
		const textboxes = Array.from(document.querySelectorAll<HTMLElement>('.docx-textbox'));

		return {
			pageCount: sections.length,
			hasMathML: !!document.querySelector('math'),
			invalidMathWrappers: document.querySelectorAll('math mo > mrow, math mo > mn, math mo > ms').length,
			footerTexts: Array.from(document.querySelectorAll('footer')).map(f => f.textContent?.trim()),
			columnCounts: Array.from(document.querySelectorAll('article')).map(a => getComputedStyle(a).columnCount),
			mathParagraphs: mathParagraphs.map(el => {
				const style = getComputedStyle(el);
				const rect = el.getBoundingClientRect();
				return {
					tagName: el.tagName,
					breakInside: style.breakInside || style.pageBreakInside,
					pageBreakInside: style.pageBreakInside,
					columnSpan: style.columnSpan,
					width: rect.width,
					height: rect.height,
				};
			}),
			mathWhiteSpace: mathElements.map(el => getComputedStyle(el).whiteSpace),
			images: images.map(el => ({
				complete: el.complete,
				naturalWidth: el.naturalWidth,
				naturalHeight: el.naturalHeight,
			})),
			textboxes: textboxes.map(el => {
				const style = getComputedStyle(el);
				return {
					boxSizing: style.boxSizing,
					overflow: style.overflow,
					paddingLeft: style.paddingLeft,
					paddingTop: style.paddingTop,
					paddingRight: style.paddingRight,
					paddingBottom: style.paddingBottom,
					alignItems: style.alignItems,
				};
			}),
		};
	});

	console.log('RESULT:', JSON.stringify(result, null, 2));
	await page.screenshot({ path: 'test-results/zz-sample-full.png', fullPage: true });

	expect(pageErrors).toEqual([]);
	expect(consoleErrors).toEqual([]);
	expect(result.pageCount).toBeGreaterThanOrEqual(1);
	expect(result.hasMathML).toBe(true);
	expect(result.invalidMathWrappers).toBe(0);
	expect(result.mathParagraphs.length).toBeGreaterThan(0);
	expect(result.mathParagraphs.every(item => item.tagName === 'DIV')).toBe(true);
	expect(result.mathParagraphs.every(item => item.breakInside === 'avoid' || item.pageBreakInside === 'avoid')).toBe(true);
	expect(result.mathParagraphs.every(item => item.columnSpan === 'none')).toBe(true);
	expect(result.mathWhiteSpace.every(value => value === 'nowrap')).toBe(true);
	expect(result.images.every(image => image.complete && image.naturalWidth > 0 && image.naturalHeight > 0)).toBe(true);
	expect(result.textboxes.every(textbox => textbox.boxSizing === 'border-box' && textbox.overflow === 'hidden' && textbox.alignItems === 'flex-start')).toBe(true);
});
