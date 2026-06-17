import { renderSync } from '../src/docx-preview';

const fixtureUrls = import.meta.glob<string>('/tests/fixtures/*.docx', { query: '?url', import: 'default', eager: true });

const select = document.querySelector<HTMLSelectElement>('#fixture-select')!;
const fileInput = document.querySelector<HTMLInputElement>('#file-input')!;
const status = document.querySelector<HTMLElement>('#status')!;
const styleContainer = document.querySelector<HTMLElement>('#style-container')!;
const documentContainer = document.querySelector<HTMLElement>('#document-container')!;

for (const path of Object.keys(fixtureUrls).sort()) {
	const name = path.split('/').pop()!;
	const option = document.createElement('option');
	option.value = fixtureUrls[path];
	option.textContent = name;
	select.appendChild(option);
}

async function render(source: Blob | ArrayBuffer) {
	status.textContent = 'rendering...';
	styleContainer.innerHTML = '';
	documentContainer.innerHTML = '';
	try {
		await renderSync(source, documentContainer, styleContainer, { inWrapper: true });
		status.textContent = 'done';
	} catch (err) {
		status.textContent = `error: ${(err as Error).message}`;
		console.error(err);
	}
}

select.addEventListener('change', async () => {
	if (!select.value) return;
	const res = await fetch(select.value);
	await render(await res.arrayBuffer());
});

fileInput.addEventListener('change', async () => {
	const file = fileInput.files?.[0];
	if (!file) return;
	await render(file);
});

if (select.options.length > 0) {
	select.dispatchEvent(new Event('change'));
}
