// Minimal preset-geometry support for DrawingML shapes (<a:prstGeom prst="...">).
// Coordinates use OOXML's normalized 21600x21600 shape space. These paths are
// visual approximations (not the exact ECMA-376 adjustable formulas) — good
// enough to render a recognizable shape instead of nothing. Unknown presets
// fall back to a plain rectangle.

const RECT = 'M0,0 L21600,0 L21600,21600 L0,21600 Z';

const PRESET_PATHS: Record<string, string> = {
	rect: RECT,
	roundRect: 'M3600,0 L18000,0 Q21600,0 21600,3600 L21600,18000 Q21600,21600 18000,21600 L3600,21600 Q0,21600 0,18000 L0,3600 Q0,0 3600,0 Z',
	ellipse: 'M0,10800 A10800,10800 0 1,0 21600,10800 A10800,10800 0 1,0 0,10800 Z',
	triangle: 'M10800,0 L21600,21600 L0,21600 Z',
	rtTriangle: 'M0,0 L21600,21600 L0,21600 Z',
	diamond: 'M10800,0 L21600,10800 L10800,21600 L0,10800 Z',
	parallelogram: 'M6500,0 L21600,0 L15100,21600 L0,21600 Z',
	trapezoid: 'M0,21600 L5400,0 L16200,0 L21600,21600 Z',
	pentagon: 'M10800,0 L21600,8260 L17550,21600 L4050,21600 L0,8260 Z',
	hexagon: 'M5400,0 L16200,0 L21600,10800 L16200,21600 L5400,21600 L0,10800 Z',
	octagon: 'M6320,0 L15280,0 L21600,6320 L21600,15280 L15280,21600 L6320,21600 L0,15280 L0,6320 Z',
	rightArrow: 'M0,5400 L13500,5400 L13500,0 L21600,10800 L13500,21600 L13500,16200 L0,16200 Z',
	leftArrow: 'M21600,5400 L8100,5400 L8100,0 L0,10800 L8100,21600 L8100,16200 L21600,16200 Z',
	upArrow: 'M5400,21600 L5400,8100 L0,8100 L10800,0 L21600,8100 L16200,8100 L16200,21600 Z',
	downArrow: 'M5400,0 L16200,0 L16200,13500 L21600,13500 L10800,21600 L0,13500 L5400,13500 Z',
	line: 'M0,0 L21600,21600',
	straightConnector1: 'M0,0 L21600,21600',
};

// Shapes that need more than one filled outline (e.g. a ring plus a bar).
// Rendered with fill-rule="evenodd" so nested/overlapping subpaths cut holes.
const PRESET_COMPOUND_PATHS: Record<string, string[]> = {
	// "No entry" / prohibited symbol: a ring (outer circle minus inner circle)
	// plus a diagonal bar across it.
	noSmoking: [
		// Two half-arc circles (instead of one near-360° arc) avoid precision
		// glitches some renderers have with start/end points that nearly coincide.
		'M1800,10800 A9000,9000 0 1,0 19800,10800 A9000,9000 0 1,0 1800,10800 Z'
		+ ' M5400,10800 A5400,5400 0 1,0 16200,10800 A5400,5400 0 1,0 5400,10800 Z',
		'M3500,3500 L17900,17900 L19800,16000 L5400,1600 Z',
	],
};

/** Returns one or more SVG path `d` strings (21600x21600 viewBox) for a preset geometry name. */
export function getPresetGeometryPaths(preset: string | undefined): string[] {
	if (preset && PRESET_COMPOUND_PATHS[preset]) {
		return PRESET_COMPOUND_PATHS[preset];
	}
	if (preset && PRESET_PATHS[preset]) {
		return [PRESET_PATHS[preset]];
	}
	return [RECT];
}
