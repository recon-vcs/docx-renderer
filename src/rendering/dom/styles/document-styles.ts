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

export function renderStyles(
	styles: IDomStyle[],
	callbacks: Pick<DocumentStylesCallbacks, 'className' | 'styleToString' | 'createStyleElement'>
): HTMLElement {
	let styleText = "";

	for (const style of styles) {
		// TODO Handle linked styles; linked styles can reference each other.

		for (const ruleset of style.rulesets) {
			//TODO temporary disable modifier until test it well
			let selector = `${style.label ?? ''}.${style.cssName}`; //${subStyle.mod ?? ''}
			if (style.label !== ruleset.target) {
				selector += ` ${ruleset.target}`;
			}
			if (style.isDefault) {
				selector = `.${callbacks.className} ${style.label}, ` + selector;
			}

			styleText += callbacks.styleToString(selector, ruleset.declarations);
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
