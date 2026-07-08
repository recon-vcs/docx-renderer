// Compares a rendered screenshot of tests/fixtures/a.docx against the ground-truth
// page images in real/a-page-*.png (rasterized from a.pdf, a Word/Microsoft 365 PDF
// export of the same file — see README "Measured accuracy" section).
//
// Manual/one-off usage. `pnpm measure:all` runs the full multi-library pipeline in one
// shot (drives agent-browser itself) and is what CI uses.
import { writeFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { scoreRenderer } from './lib/pixel-diff.mjs';

const ROOT = fileURLToPath(new URL('..', import.meta.url));
const RENDERED = process.argv[2];
const LABEL = process.argv[3] ?? 'renderer';
const SECTIONS_JSON = process.argv[4];

if (!RENDERED || !SECTIONS_JSON) {
	console.error(
		'usage: node scripts/measure-accuracy.mjs <rendered-screenshot.png> <label> \'<section-rects-json>\'',
	);
	process.exit(1);
}

const result = scoreRenderer({
	renderedScreenshotPath: RENDERED,
	sections: JSON.parse(SECTIONS_JSON),
	realDir: `${ROOT}/real`,
	outDir: `${ROOT}/accuracy/${LABEL}`,
});

console.log(JSON.stringify({ label: LABEL, ...result }, null, 2));
