import { createElement, createStyleElement } from '@docx/rendering/dom/core/dom-utils';

export function renderDefaultStyle(className: string): HTMLElement {
	const c = className;
	const styleText = `
			.${c} { font-family: system-ui, -apple-system, "Segoe UI", Roboto, "Helvetica Neue", "Noto Sans", "Liberation Sans", Arial, sans-serif }
			.${c}-wrapper { background: gray; padding: 30px; padding-bottom: 0px; display: flex; flex-flow: column; align-items: center; line-height:normal; font-weight:normal; } 
			.${c}-wrapper>section.${c} { background: white; box-shadow: 0 0 10px rgba(0, 0, 0, 0.5); margin-bottom: 30px; }
			.${c} { color: black; hyphens: auto; text-underline-position: from-font; }
			section.${c} { box-sizing: border-box; display: flex; flex-flow: column nowrap; position: relative; overflow: hidden; }
            section.${c}>header { position: absolute; top: 0; z-index: 1; display: flex; flex-direction: column; justify-content: flex-end; }
			section.${c}>article { z-index: 1; }
			section.${c}>footer { position: absolute; bottom: 0; z-index: 1; }
			.${c} table { border-collapse: collapse; break-inside: avoid; }
			.${c} table td, .${c} table th { vertical-align: top; }
			.${c} p { margin: 0pt; min-height: 1em; }
			.${c} span { white-space: pre-wrap; overflow-wrap: break-word; }
			.${c} math { vertical-align: middle; }
			.${c} .${c}-math-paragraph { break-inside: avoid; page-break-inside: avoid; }
			.${c} .${c}-math-paragraph math { max-width: 100%; overflow: visible; }
			.${c} .${c}-math-paragraph math * { overflow-wrap: normal; }
			.${c} a { color: inherit; text-decoration: inherit; }
			.${c} img, .${c} svg { vertical-align: baseline; break-inside: avoid; page-break-inside: avoid; }
			.${c} .${c}-drawing { display: inline-block; break-inside: avoid; page-break-inside: avoid; }
			.${c} svg { fill: transparent; break-inside: avoid; page-break-inside: avoid; }
			.${c} .clearfix::after { content: ""; display: block; line-height: 0; clear: both; }
			.${c} br.break.column { break-after: column; }
			.${c} s.break.section { display: block; }
			.${c} s.break.section[data-type="nextColumn"] { break-before: column; }
		`;

	return createStyleElement(styleText);
}

export function renderWrapper(className: string): HTMLElement {
	return createElement('div', { className: `${className}-wrapper` });
}
