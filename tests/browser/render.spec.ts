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
		const count = await page.locator('.section.docx').count();
		expect(count).toBeGreaterThanOrEqual(1);
	});

	test('columns renders multi-column content', async ({ page }) => {
		const pageErrors: Error[] = [];
		page.on('pageerror', (error) => pageErrors.push(error));

		const html = await renderInPage(page, 'columns');

		expect(pageErrors).toEqual([]);
		expect(html).toMatch(/column-count|columns:/);
	});
});
