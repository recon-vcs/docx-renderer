// Word (via GDI/DirectWrite) computes a run's line contribution by scaling
// the font's ascent and descent to the pixel size and rounding EACH one up
// to a whole pixel before summing. Browsers scale the same metrics but round
// the sum, so a font like Meiryo at 14px measures 21px in Chromium while
// Word uses ceil(14.84) + ceil(6.16) = 22px - a systematic 1px-per-line
// vertical drift against Word output.
//
// The font's true ascent/descent ratios are recovered at runtime by probing
// the font at a large size on a canvas (integer rounding noise becomes
// negligible), so this needs no per-font metrics tables and works for any
// installed or embedded font. The GDI-rounded height is then applied as an
// explicit line-height wherever the CSS value would otherwise be `normal`;
// explicit line heights (exact/atLeast/multiple spacing) always win and are
// left untouched.

const PROBE_PX = 1024;
// Guards ceil() against float noise when a scaled metric is a whole number.
const CEIL_EPSILON = 1e-4;

interface FontRatios {
	ascent: number;
	descent: number;
}

const ratioCache = new Map<string, FontRatios | null>();
let probeContext: CanvasRenderingContext2D | null = null;

function measureFontRatios(probeFont: string): FontRatios | null {
	const cached = ratioCache.get(probeFont);
	if (cached !== undefined) return cached;

	if (!probeContext) {
		probeContext = document.createElement('canvas').getContext('2d');
	}

	let ratios: FontRatios | null = null;
	if (probeContext) {
		probeContext.font = probeFont;
		const metrics = probeContext.measureText('Hg');
		if (typeof metrics.fontBoundingBoxAscent === 'number' && typeof metrics.fontBoundingBoxDescent === 'number') {
			ratios = {
				ascent: metrics.fontBoundingBoxAscent / PROBE_PX,
				descent: metrics.fontBoundingBoxDescent / PROBE_PX,
			};
		}
	}

	ratioCache.set(probeFont, ratios);
	return ratios;
}

function gdiLineHeightPx(ratios: FontRatios, fontSizePx: number): number {
	const ascent = Math.ceil(ratios.ascent * fontSizePx - CEIL_EPSILON);
	const descent = Math.ceil(ratios.descent * fontSizePx - CEIL_EPSILON);
	return ascent + descent;
}

/**
 * Replaces an element's `line-height: normal` with the GDI-rounded line
 * height of its computed font, so line stacking matches Word. The element
 * must already be attached to the document.
 */
export function applyGdiLineHeight(el: HTMLElement): void {
	const cs = getComputedStyle(el);
	if (cs.lineHeight !== 'normal') return;

	const fontSizePx = parseFloat(cs.fontSize);
	if (!Number.isFinite(fontSizePx) || fontSizePx <= 0) return;

	// The full computed family stack resolves on the canvas exactly like it
	// does on the element, and weight/style can select faces with different
	// vertical metrics.
	const probeFont = `${cs.fontStyle} ${cs.fontWeight} ${PROBE_PX}px ${cs.fontFamily}`;
	const ratios = measureFontRatios(probeFont);
	if (!ratios) return;

	el.style.lineHeight = `${gdiLineHeightPx(ratios, fontSizePx)}px`;
}
