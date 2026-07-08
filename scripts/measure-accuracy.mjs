// Compares the rendered output of tests/fixtures/a.docx against the ground-truth
// page images in real/a-page-*.png (rendered from a.pdf, the same document exported
// to PDF by Microsoft Word — see README "Measured accuracy" section) using per-page
// pixel diffing. Regenerate the rendered screenshot via the playground before running
// this script.
import { PNG } from 'pngjs';
import pixelmatch from 'pixelmatch';
import { readFileSync, writeFileSync, mkdirSync, readdirSync } from 'node:fs';
import { fileURLToPath } from 'node:url';

const ROOT = fileURLToPath(new URL('..', import.meta.url));
const RENDERED = process.argv[2];
const REAL_DIR = `${ROOT}/real`;
const OUT_DIR = `${ROOT}/accuracy`;

if (!RENDERED) {
	console.error('usage: node scripts/measure-accuracy.mjs <path-to-rendered-full-page-screenshot.png>');
	process.exit(1);
}

// Bounding boxes of the rendered <section class="docx"> pages, read via
// `document.querySelectorAll('section')[i].getBoundingClientRect()` in the playground
// at viewport 1240x1754.
const sections = [
	{ top: 92, left: 223.140625, width: 793.71875, height: 1122.53125 },
	{ top: 1244.53125, left: 223.140625, width: 793.71875, height: 1122.53125 },
	{ top: 2397.0625, left: 223.140625, width: 793.71875, height: 1190.5 },
	{ top: 3617.5625, left: 212, width: 816, height: 1056 },
	{ top: 4703.5625, left: 223.140625, width: 793.71875, height: 1122.53125 },
];

function loadPng(path) {
	return PNG.sync.read(readFileSync(path));
}

function crop(png, rect) {
	const x = Math.round(rect.left);
	const y = Math.round(rect.top);
	const w = Math.round(rect.width);
	const h = Math.round(rect.height);
	const out = new PNG({ width: w, height: h });
	PNG.bitblt(png, out, x, y, w, h, 0, 0);
	return out;
}

function resizeNearest(src, targetW, targetH) {
	const out = new PNG({ width: targetW, height: targetH });
	for (let y = 0; y < targetH; y++) {
		const sy = Math.min(src.height - 1, Math.floor((y * src.height) / targetH));
		for (let x = 0; x < targetW; x++) {
			const sx = Math.min(src.width - 1, Math.floor((x * src.width) / targetW));
			const si = (sy * src.width + sx) * 4;
			const di = (y * targetW + x) * 4;
			out.data[di] = src.data[si];
			out.data[di + 1] = src.data[si + 1];
			out.data[di + 2] = src.data[si + 2];
			out.data[di + 3] = src.data[si + 3];
		}
	}
	return out;
}

mkdirSync(OUT_DIR, { recursive: true });

const pageFiles = readdirSync(REAL_DIR)
	.filter((name) => /^a-page-\d+\.png$/.test(name))
	.sort((a, b) => Number(a.match(/\d+/)[0]) - Number(b.match(/\d+/)[0]));

const rendered = loadPng(RENDERED);
const results = [];

for (let i = 0; i < pageFiles.length; i++) {
	const refPage = loadPng(`${REAL_DIR}/${pageFiles[i]}`);
	const renderedPage = crop(rendered, sections[i]);
	const renderedResized = resizeNearest(renderedPage, refPage.width, refPage.height);

	const diff = new PNG({ width: refPage.width, height: refPage.height });
	const diffPixels = pixelmatch(refPage.data, renderedResized.data, diff.data, refPage.width, refPage.height, {
		threshold: 0.15,
	});
	const totalPixels = refPage.width * refPage.height;
	const score = 1 - diffPixels / totalPixels;

	writeFileSync(`${OUT_DIR}/page-${i + 1}-reference.png`, PNG.sync.write(refPage));
	writeFileSync(`${OUT_DIR}/page-${i + 1}-rendered.png`, PNG.sync.write(renderedResized));
	writeFileSync(`${OUT_DIR}/page-${i + 1}-diff.png`, PNG.sync.write(diff));

	results.push({ page: i + 1, score, diffPixels, totalPixels });
}

const overall = results.reduce((sum, r) => sum + r.score, 0) / results.length;

console.log(JSON.stringify({ results, overall }, null, 2));
