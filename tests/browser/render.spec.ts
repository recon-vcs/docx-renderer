import { expect, test, type Page } from '@playwright/test';

/**
 * Shape of the UMD global exposed by dist/docx-vellum.umd.js inside the
 * harness page. Only the members used by these tests are declared.
 */
interface DocxGlobal {
	renderSync(data: Blob, bodyContainer: HTMLElement): Promise<unknown>;
}

type HarnessWindow = Window & typeof globalThis & { docx: DocxGlobal };

/** Representative fixtures for the synchronous renderer smoke tests. */
const SYNC_SMOKE_CASES = [
	'text',
	'table',
	'numbering',
	'header-footer',
	'footnote',
	'page-layout',
	'image-clip',
	'text-wrap-square',
	'comment',
	'break-page',
	'underlines',
	'text-break',
	'revision',
	'line-spacing',
	'equation',
] as const;

async function renderInPage(page: Page, fixture: string): Promise<string> {
	await page.goto('/tests/browser/harness.html');
	return page.evaluate(
		async ({ fixture }) => {
			const response = await fetch(`/tests/fixtures/${fixture}.docx`);
			if (!response.ok) {
				throw new Error(`failed to fetch fixture ${fixture}: ${response.status}`);
			}
			const blob = await response.blob();
			const body = document.querySelector<HTMLElement>('#document-container');
			if (!body) {
				throw new Error('harness is missing #document-container');
			}
			const docx = (window as HarnessWindow).docx;
			await docx.renderSync(blob, body);
			return body.innerHTML;
		},
		{ fixture },
	);
}

test.describe('renderSync smoke', () => {
	for (const name of SYNC_SMOKE_CASES) {
		test(`renders ${name} without errors`, async ({ page }) => {
			const pageErrors: Error[] = [];
			page.on('pageerror', (error) => pageErrors.push(error));

			const html = await renderInPage(page, name);

			expect(pageErrors).toEqual([]);
			expect(html.length).toBeGreaterThan(0);
			expect(html).toContain('docx-wrapper');
		});
	}
});

const SECTION_SMOKE_CASES = [
	'section-break',
	'columns',
	'column-break',
	'footer-with-section',
	'header-section',
	'break-page-section-break',
] as const;

test.describe('section and break smoke', () => {
	for (const name of SECTION_SMOKE_CASES) {
		test(`renders ${name} without errors`, async ({ page }) => {
			const pageErrors: Error[] = [];
			page.on('pageerror', (error) => pageErrors.push(error));

			const html = await renderInPage(page, name);

			expect(pageErrors).toEqual([]);
			expect(html.length).toBeGreaterThan(0);
			expect(html).toContain('docx-wrapper');
		});
	}

	test('section-break renders at least one page', async ({ page }) => {
		const pageErrors: Error[] = [];
		page.on('pageerror', (error) => pageErrors.push(error));

		await renderInPage(page, 'section-break');

		expect(pageErrors).toEqual([]);
		const count = await page.locator('section.docx').count();
		expect(count).toBeGreaterThanOrEqual(1);
	});

	test('columns renders multi-column content', async ({ page }) => {
		const pageErrors: Error[] = [];
		page.on('pageerror', (error) => pageErrors.push(error));

		await renderInPage(page, 'columns');
		const articleColumnCountsByPage = await page.evaluate(() => {
			const pages = Array.from(document.querySelectorAll<HTMLElement>('section.docx'));

			return pages.map(renderedPage => {
				const articles = Array.from(renderedPage.querySelectorAll<HTMLElement>(':scope > article'));

				return articles.map(article => getComputedStyle(article).columnCount);
			});
		});
		const allColumnCounts = articleColumnCountsByPage.flat();

		expect(pageErrors).toEqual([]);
		expect(articleColumnCountsByPage.some(columnCounts => columnCounts.length > 1)).toBe(true);
		expect(allColumnCounts).toContain('2');
		expect(allColumnCounts).toContain('3');
	});

	test('header image pagination does not repeat the same overflowing block', async ({ page }) => {
		const pageErrors: Error[] = [];
		page.on('pageerror', (error) => pageErrors.push(error));

		await renderInPage(page, 'break-page-header-image');
		const metrics = await page.evaluate(() => {
			const pages = Array.from(document.querySelectorAll<HTMLElement>('section.docx'));
			const images = Array.from(document.querySelectorAll<HTMLImageElement>('img'));

			return {
				pageCount: pages.length,
				imageCount: images.length,
				imagesDecoded: images.every(image => image.complete && image.naturalWidth > 0 && image.naturalHeight > 0),
				pageTexts: pages.map(renderedPage => renderedPage.querySelector('article')?.textContent?.trim() ?? ''),
			};
		});

		expect(pageErrors).toEqual([]);
		expect(metrics.pageCount).toBeGreaterThan(1);
		expect(metrics.pageCount).toBeLessThan(10);
		expect(metrics.imageCount).toBeGreaterThan(0);
		expect(metrics.imagesDecoded).toBe(true);
		expect(metrics.pageTexts.some(text => text.includes('UN'))).toBe(true);
	});

});
