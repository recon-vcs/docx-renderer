import { render as renderOurs } from '../src/docx-preview';
import { renderAsync as renderSync } from 'docx-preview-sync';
import { renderAsync as renderOriginal } from 'docx-preview';

const fixtureUrls = import.meta.glob<string>('../tests/fixtures/a.docx', { query: '?url', import: 'default', eager: true });
const FIXTURE_URL = Object.values(fixtureUrls)[0];

const select = document.querySelector<HTMLSelectElement>('#renderer-select')!;
const status = document.querySelector<HTMLElement>('#status')!;
const styleContainer = document.querySelector<HTMLElement>('#style-container')!;
const documentContainer = document.querySelector<HTMLElement>('#document-container')!;

async function renderWith(renderer: string, blob: Blob) {
	status.textContent = 'rendering...';
	styleContainer.innerHTML = '';
	documentContainer.innerHTML = '';
	try {
		const options = { breakPages: true };
		if (renderer === 'docx-renderer') {
			await renderOurs(blob, documentContainer, styleContainer, options);
		} else if (renderer === 'docx-preview-sync') {
			await renderSync(blob, documentContainer, styleContainer, options);
		} else {
			await renderOriginal(blob, documentContainer, styleContainer, options);
		}
		status.textContent = `done (${document.querySelectorAll('section').length} pages)`;
	} catch (err) {
		status.textContent = `error: ${(err as Error).message}`;
		console.error(err);
	}
}

select.addEventListener('change', async () => {
	const res = await fetch(FIXTURE_URL);
	await renderWith(select.value, await res.blob());
});

select.dispatchEvent(new Event('change'));
