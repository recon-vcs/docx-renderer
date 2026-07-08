import { PNG } from 'pngjs';
import pixelmatch from 'pixelmatch';
import { readFileSync, writeFileSync, mkdirSync, readdirSync } from 'node:fs';

export function loadPng(path) {
	return PNG.sync.read(readFileSync(path));
}

export function crop(png, rect) {
	const x = Math.round(rect.left);
	const y = Math.round(rect.top);
	const w = Math.round(rect.width);
	const h = Math.round(rect.height);
	const out = new PNG({ width: w, height: h });
	PNG.bitblt(png, out, x, y, w, h, 0, 0);
	return out;
}

export function resizeNearest(src, targetW, targetH) {
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

export function listReferencePages(realDir) {
	return readdirSync(realDir)
		.filter((name) => /^a-page-\d+\.png$/.test(name))
		.sort((a, b) => Number(a.match(/\d+/)[0]) - Number(b.match(/\d+/)[0]));
}

/**
 * Scores a renderer's full-page screenshot against the reference pages.
 * Missing pages (renderer produced fewer sections than the reference has pages)
 * score 0 against the fixed reference page count instead of being averaged away.
 */
export function scoreRenderer({ renderedScreenshotPath, sections, realDir, outDir }) {
	mkdirSync(outDir, { recursive: true });
	const pageFiles = listReferencePages(realDir);
	const rendered = loadPng(renderedScreenshotPath);
	const results = [];

	for (let i = 0; i < pageFiles.length; i++) {
		const refPage = loadPng(`${realDir}/${pageFiles[i]}`);

		if (i >= sections.length) {
			results.push({ page: i + 1, score: 0, missing: true });
			continue;
		}

		const renderedPage = crop(rendered, sections[i]);
		const renderedResized = resizeNearest(renderedPage, refPage.width, refPage.height);

		const diff = new PNG({ width: refPage.width, height: refPage.height });
		const diffPixels = pixelmatch(refPage.data, renderedResized.data, diff.data, refPage.width, refPage.height, {
			threshold: 0.15,
		});
		const totalPixels = refPage.width * refPage.height;
		const score = 1 - diffPixels / totalPixels;

		writeFileSync(`${outDir}/page-${i + 1}-reference.png`, PNG.sync.write(refPage));
		writeFileSync(`${outDir}/page-${i + 1}-rendered.png`, PNG.sync.write(renderedResized));
		writeFileSync(`${outDir}/page-${i + 1}-diff.png`, PNG.sync.write(diff));

		results.push({ page: i + 1, score, diffPixels, totalPixels });
	}

	const overall = results.reduce((sum, r) => sum + r.score, 0) / results.length;
	return { pagesRendered: sections.length, results, overall };
}
