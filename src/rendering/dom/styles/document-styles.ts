import * as _ from 'lodash-es';
import type { Options } from '@docx/options';
import { FontTablePart } from '@docx/ooxml/wordprocessingml/parts/font-table/font-table';
import { ThemePart } from '@docx/ooxml/drawingml/theme/theme-part';
import { escapeClassName } from '@docx/shared/utils';
import { IDomStyle, Ruleset } from '@docx/ooxml/wordprocessingml/document/model/style';

export interface DocumentStylesCallbacks {
	className: string;
	options: Options;
	styleToString(selectors: string, declarations: Record<string, string>, cssText?: string): string;
	processStyleName(className: string): string;
	createStyleElement(cssText: string): HTMLElement;
	appendComment(styleContainer: HTMLElement, text: string): void;
	loadFont(id: string, key: string): Promise<string>;
	refreshTabStops(): void;
}

export function processStyleName(className: string, baseClassName: string): string {
	return className ? `${baseClassName}_${escapeClassName(className)}` : baseClassName;
}

export function renderTheme(
	themePart: ThemePart,
	styleContainer: HTMLElement,
	callbacks: DocumentStylesCallbacks
): void {
	const variables: Record<string, string> = {};
	const fontScheme = themePart.theme?.fontScheme;

	if (fontScheme) {
		if (fontScheme.majorFont) {
			variables['--docx-majorHAnsi-font'] = fontScheme.majorFont.latinTypeface;
		}

		if (fontScheme.minorFont) {
			variables['--docx-minorHAnsi-font'] = fontScheme.minorFont.latinTypeface;
		}
	}

	const colorScheme = themePart.theme?.colorScheme;

	if (colorScheme) {
		for (const [k, v] of Object.entries(colorScheme.colors)) {
			variables[`--docx-${k}-color`] = `#${v}`;
		}
	}

	const cssText = callbacks.styleToString(`.${callbacks.className}`, variables);
	styleContainer.appendChild(callbacks.createStyleElement(cssText));
}

export function processStyles(
	styles: IDomStyle[],
	callbacks: Pick<DocumentStylesCallbacks, 'processStyleName' | 'options'>
): Record<string, IDomStyle> {
	const stylesMap = _.keyBy(styles, 'id');

	for (const childStyle of styles) {
		childStyle.cssName = callbacks.processStyleName(childStyle.id);

		if (childStyle.basedOn === null) {
			continue;
		}

		const parentStyle = stylesMap[childStyle.basedOn];

		if (parentStyle) {
			if (parentStyle?.paragraphProps) {
				childStyle.paragraphProps = _.merge({}, parentStyle?.paragraphProps, childStyle.paragraphProps);
			}
			if (parentStyle?.runProps) {
				childStyle.runProps = _.merge({}, parentStyle?.runProps, childStyle.runProps);
			}

			for (const parentRuleset of parentStyle.rulesets) {
				const childRuleset: Ruleset = childStyle.rulesets.find(r => r.target == parentRuleset.target);

				if (childRuleset) {
					childRuleset.declarations = _.merge({}, parentRuleset.declarations, childRuleset.declarations);
				} else {
					childStyle.rulesets.push({ ...parentRuleset });
				}
			}
		} else if (callbacks.options.debug) {
			console.warn(`Can't find base style ${childStyle.basedOn}`);
		}
	}

	return stylesMap;
}

// The only run-formatting properties that affect a block's own "strut" (the
// invisible reference line box every block-level element inserts, sized from
// its own font, that a line must be at least as tall as). Only these two are
// mirrored onto paragraph elements below - anything else from a run ruleset
// (color, background, decoration...) would visibly paint a full-width box
// behind/around the text if applied to the paragraph itself, which real strut
// boxes never do.
const STRUT_PROPERTIES = ['font-family', 'font-size'] as const;

function pickStrutDeclarations(declarations: Record<string, string>): Record<string, string> | null {
	const picked: Record<string, string> = {};
	for (const key of STRUT_PROPERTIES) {
		if (declarations[key] !== undefined) {
			picked[key] = declarations[key];
		}
	}
	return Object.keys(picked).length > 0 ? picked : null;
}

function buildSelector(style: IDomStyle, target: string, className: string): string {
	let selector = `${style.label ?? ''}.${style.cssName}`;
	if (style.label !== target) {
		selector += ` ${target}`;
	}
	if (style.isDefault) {
		selector = `.${className} ${style.label}, ` + selector;
	}
	return selector;
}

export function renderStyles(
	styles: IDomStyle[],
	callbacks: Pick<DocumentStylesCallbacks, 'className' | 'styleToString' | 'createStyleElement'>
): HTMLElement {
	let styleText = "";

	for (const style of styles) {
		// TODO Handle linked styles; linked styles can reference each other.

		// Document defaults (docDefaults' rPrDefault, style.id === null) and
		// paragraph/numbering styles' own rPr define the *default run*
		// formatting for a paragraph - i.e. what an empty/under-filled line
		// falls back to. Word derives such a line's height from that default
		// run's font, so the paragraph element's own strut must use the same
		// font metrics, not the browser's UA default. Character styles are
		// excluded: they only ever apply to explicitly rStyle'd runs, never
		// define a paragraph's default formatting.
		const isParagraphScoped = style.id === null || style.label === 'p';

		for (const ruleset of style.rulesets) {
			//TODO temporary disable modifier until test it well
			const selector = buildSelector(style, ruleset.target, callbacks.className);

			styleText += callbacks.styleToString(selector, ruleset.declarations);

			if (isParagraphScoped && ruleset.target === 'span') {
				const strutDeclarations = pickStrutDeclarations(ruleset.declarations);
				if (strutDeclarations) {
					const strutSelector = buildSelector(style, 'p', callbacks.className);
					styleText += callbacks.styleToString(strutSelector, strutDeclarations);
				}
			}
		}

		// Ignore Spacing Above and Below When Using Identical Styles
		// (w:contextualSpacing): Word drops the spacing between two paragraphs
		// of the same style when they're contiguous (typically list items).
		// This is a purely structural, class-based relationship, so it maps
		// directly onto adjacent-sibling CSS - no per-paragraph traversal needed.
		if (style.label === 'p' && style.paragraphProps?.contextualSpacing) {
			const selfSelector = `p.${style.cssName}`;
			styleText += callbacks.styleToString(`${selfSelector} + ${selfSelector}`, { 'padding-top': '0pt' });
			styleText += callbacks.styleToString(`${selfSelector}:has(+ ${selfSelector})`, { 'margin-bottom': '0pt' });
		}
	}

	return callbacks.createStyleElement(styleText);
}

export function renderFontTable(
	fontsPart: FontTablePart,
	styleContainer: HTMLElement,
	callbacks: Pick<DocumentStylesCallbacks, 'styleToString' | 'createStyleElement' | 'appendComment' | 'loadFont' | 'refreshTabStops'>
): void {
	for (const f of fontsPart.fonts) {
		for (const ref of f.embedFontRefs) {
			callbacks.loadFont(ref.id, ref.key).then(fontData => {
				const cssValues: Record<string, string> = {
					'font-family': f.name,
					src: `url(${fontData})`,
				};

				if (ref.type == 'bold' || ref.type == 'boldItalic') {
					cssValues['font-weight'] = 'bold';
				}

				if (ref.type == 'italic' || ref.type == 'boldItalic') {
					cssValues['font-style'] = 'italic';
				}

				callbacks.appendComment(styleContainer, `docxjs ${f.name} font`);
				const cssText = callbacks.styleToString('@font-face', cssValues);
				styleContainer.appendChild(callbacks.createStyleElement(cssText));
				callbacks.refreshTabStops();
			});
		}
	}
}
