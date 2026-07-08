// Regenerates accuracy/comparison-{light,dark}.svg from the latest ACCURACY.md
// summary numbers, without re-running the full render-and-screenshot pipeline.
// Use `pnpm measure:all` to regenerate real scores end-to-end.
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { renderComparisonChart } from './lib/comparison-chart.mjs';

const ROOT = fileURLToPath(new URL('..', import.meta.url));
const report = readFileSync(`${ROOT}/ACCURACY.md`, 'utf8');

const summaryRows = [...report.matchAll(/^\| (.+?) \| \d+ \/ \d+ \| \*\*([\d.]+)%\*\* \|$/gm)];
if (summaryRows.length === 0) {
	console.error('No summary rows found in ACCURACY.md — run `pnpm measure:all` first.');
	process.exit(1);
}

const data = summaryRows.map(([, title, scorePct]) => ({
	title,
	overall: Number(scorePct) / 100,
	highlight: title.includes('this project'),
}));

renderComparisonChart(data, `${ROOT}/accuracy`);
console.log('wrote accuracy/comparison-light.svg and accuracy/comparison-dark.svg');
