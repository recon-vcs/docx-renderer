import { writeFileSync } from 'node:fs';

// Categorical slots 1/2/3 (blue/aqua/yellow) in fixed order — validated with
// scripts/validate_palette.js (see dataviz skill). Axis is truncated at 40%
// (all three scores sit above it) with a visible break mark near the origin
// so the truncation reads as intentional, not a zeroed baseline.
const THEMES = {
	light: {
		surface: 'transparent',
		textPrimary: '#0b0b0b',
		textMuted: '#898781',
		gridline: '#e1e0d9',
		baseline: '#c3c2b7',
		series: ['#2a78d6', '#1baf7a', '#eda100'],
	},
	dark: {
		surface: 'transparent',
		textPrimary: '#ffffff',
		textMuted: '#c3c2b7',
		gridline: '#2c2c2a',
		baseline: '#383835',
		series: ['#3987e5', '#199e70', '#c98500'],
	},
};

const WIDTH = 680;
const LABEL_X = 8;
const BAR_X0 = 260;
const BAR_MAX_W = 360; // AXIS_MAX px width
const BAR_H = 26;
const ROW_H = 54;
const ROW_Y0 = 22;
const RADIUS = 4;
const FONT = "system-ui, -apple-system, 'Segoe UI', sans-serif";
const AXIS_MIN = 0.4;
const AXIS_MAX = 1.0;
const AXIS_TICKS = [0.4, 0.6, 0.8, 1.0];

function roundedBarPath(x, y, w, h, r) {
	if (w < r) r = w;
	return `M${x},${y} L${x + w - r},${y} A${r},${r} 0 0 1 ${x + w},${y + r} L${x + w},${y + h - r} A${r},${r} 0 0 1 ${x + w - r},${y + h} L${x},${y + h} Z`;
}

function valueToX(value) {
	return BAR_X0 + ((value - AXIS_MIN) / (AXIS_MAX - AXIS_MIN)) * BAR_MAX_W;
}

/**
 * @param {{title: string, overall: number}[]} data
 * @param {string} outDir
 */
export function renderComparisonChart(data, outDir) {
	const height = ROW_Y0 + data.length * ROW_H + 32;
	const axisY = ROW_Y0 + data.length * ROW_H - (ROW_H - BAR_H) + 18;

	for (const theme of Object.keys(THEMES)) {
		const t = THEMES[theme];

		const gridlines = AXIS_TICKS.map((v) => {
			const x = valueToX(v);
			return `<line x1="${x}" y1="${ROW_Y0 - 8}" x2="${x}" y2="${axisY}" stroke="${t.gridline}" stroke-width="1"/>`;
		}).join('\n\t');

		const ticks = AXIS_TICKS.map((v) => {
			const x = valueToX(v);
			return `<text x="${x}" y="${axisY + 16}" font-size="11" fill="${t.textMuted}" text-anchor="middle" font-family="${FONT}">${Math.round(v * 100)}%</text>`;
		}).join('\n\t');

		// Axis-break glyph: two short parallel diagonals over the baseline at the
		// truncated origin, the standard "this axis does not start at zero" mark.
		const breakX = valueToX(AXIS_MIN);
		const axisBreak = `
	<line x1="${breakX - 4}" y1="${axisY + 5}" x2="${breakX + 2}" y2="${axisY - 5}" stroke="${t.surface === 'transparent' ? t.baseline : t.surface}" stroke-width="6"/>
	<line x1="${breakX - 5}" y1="${axisY + 3}" x2="${breakX + 1}" y2="${axisY - 7}" stroke="${t.textMuted}" stroke-width="1.5"/>
	<line x1="${breakX - 1}" y1="${axisY + 3}" x2="${breakX + 5}" y2="${axisY - 7}" stroke="${t.textMuted}" stroke-width="1.5"/>`;

		const rows = data
			.map((d, i) => {
				const y = ROW_Y0 + i * ROW_H;
				const x1 = valueToX(Math.max(AXIS_MIN, Math.min(AXIS_MAX, d.overall)));
				const w = Math.max(0, x1 - BAR_X0);
				const fill = t.series[i % t.series.length];
				const path = roundedBarPath(BAR_X0, y, w, BAR_H, RADIUS);
				const pctLabel = `${(d.overall * 100).toFixed(1)}%`;
				return `
	<text x="${LABEL_X}" y="${y + BAR_H / 2 + 4}" font-size="13" font-weight="500" fill="${t.textPrimary}" font-family="${FONT}">${d.title}</text>
	<path d="${path}" fill="${fill}"/>
	<text x="${BAR_X0 + w + 10}" y="${y + BAR_H / 2 + 4}" font-size="13" font-weight="600" fill="${t.textPrimary}" font-family="${FONT}">${pctLabel}</text>`;
			})
			.join('\n');

		const ariaLabel = `Pixel-accuracy comparison against Word's PDF export of a.docx (axis truncated at 40%): ${data.map((d) => `${d.title} ${(d.overall * 100).toFixed(1)}%`).join(', ')}`;

		const svg = `<svg viewBox="0 0 ${WIDTH} ${height}" xmlns="http://www.w3.org/2000/svg" role="img" aria-label="${ariaLabel}">
	<rect x="0" y="0" width="${WIDTH}" height="${height}" fill="${t.surface}"/>
	${gridlines}
	<line x1="${BAR_X0}" y1="${ROW_Y0 - 8}" x2="${BAR_X0}" y2="${axisY}" stroke="${t.baseline}" stroke-width="1.5"/>
	${rows}
	${ticks}
	${axisBreak}
</svg>
`;

		writeFileSync(`${outDir}/comparison-${theme}.svg`, svg);
	}
}
