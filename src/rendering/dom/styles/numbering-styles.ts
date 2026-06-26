import { IDomNumbering } from '@docx/ooxml/wordprocessingml/document/model/dom';
import { IDomStyle } from '@docx/ooxml/wordprocessingml/document/model/style';

export interface NumberingStylesCallbacks {
	className: string;
	rootSelector: string;
	findStyle(styleName: string): IDomStyle;
	styleToString(selectors: string, declarations: Record<string, string>, cssText?: string): string;
	createStyleElement(cssText: string): HTMLElement;
	loadNumberingImage(src: string): Promise<string>;
	numberingClass(id: string, level: number): string;
	numberingCounter(id: string, level: number): string;
	levelTextToContent(text: string, suff: string, id: string, numformat: string): string;
	numFormatToCssValue(format: string): string;
}

export function processNumberings(
	numberings: IDomNumbering[],
	callbacks: Pick<NumberingStylesCallbacks, 'findStyle'>
): void {
	for (const num of numberings.filter(n => n.pStyleName)) {
		const style = callbacks.findStyle(num.pStyleName);

		if (style?.paragraphProps?.numbering) {
			style.paragraphProps.numbering.level = num.level;
		}
	}
}

export function renderNumbering(
	numberings: IDomNumbering[],
	styleContainer: HTMLElement,
	callbacks: NumberingStylesCallbacks
): HTMLElement {
	let styleText = '';
	const resetCounters = [];

	for (const num of numberings) {
		const selector = `p.${callbacks.numberingClass(num.id, num.level)}`;
		let listStyleType = 'none';

		if (num.bullet) {
			const valiable = `--${callbacks.className}-${num.bullet.src}`.toLowerCase();

			styleText += callbacks.styleToString(`${selector}:before`, {
				"content": "' '",
				"display": "inline-block",
				"background": `var(${valiable})`
			}, num.bullet.style);

			callbacks.loadNumberingImage(num.bullet.src).then(data => {
				const text = `${callbacks.rootSelector} { ${valiable}: url(${data}) }`;
				styleContainer.appendChild(callbacks.createStyleElement(text));
			});
		} else if (num.levelText) {
			const counter = callbacks.numberingCounter(num.id, num.level);
			const counterReset = counter + ' ' + (num.start - 1);
			if (num.level > 0) {
				styleText += callbacks.styleToString(`p.${callbacks.numberingClass(num.id, num.level - 1)}`, {
					"counter-reset": counterReset
				});
			}
			resetCounters.push(counterReset);

			styleText += callbacks.styleToString(`${selector}:before`, {
				"content": callbacks.levelTextToContent(num.levelText, num.suff, num.id, callbacks.numFormatToCssValue(num.format)),
				"counter-increment": counter,
				...num.rStyle,
			});
		} else {
			listStyleType = callbacks.numFormatToCssValue(num.format);
		}

		styleText += callbacks.styleToString(selector, {
			display: 'list-item',
			'list-style-position': 'inside',
			'list-style-type': listStyleType,
			...num.pStyle,
		});
	}

	if (resetCounters.length > 0) {
		styleText += callbacks.styleToString(callbacks.rootSelector, {
			'counter-reset': resetCounters.join(' '),
		});
	}

	return callbacks.createStyleElement(styleText);
}

export function numberingClass(className: string, id: string, lvl: number): string {
	return `${className}-num-${id}-${lvl}`;
}

export function styleToString(selectors: string, declarations: Record<string, string>, cssText: string = null): string {
	let result = `${selectors} {\r\n`;

	for (const key in declarations) {
		if (key.startsWith('$')) {
			continue;
		}

		result += `  ${key}: ${declarations[key]};\r\n`;
	}

	if (cssText) {
		result += cssText;
	}

	return result + '}\r\n';
}

export function numberingCounter(className: string, id: string, lvl: number): string {
	return `${className}-num-${id}-${lvl}`;
}

export function levelTextToContent(
	text: string,
	suff: string,
	id: string,
	numformat: string,
	resolveCounter: (id: string, level: number) => string
): string {
	const suffMap = {
		tab: '\\9',
		space: '\\a0',
	};

	const result = text.replace(/%\d*/g, s => {
		const lvl = parseInt(s.substring(1), 10) - 1;
		return `"counter(${resolveCounter(id, lvl)}, ${numformat})"`;
	});

	return `"${result}${suffMap[suff] ?? ''}"`;
}

export function numFormatToCssValue(format: string): string {
	const mapping = {
		none: 'none',
		bullet: 'disc',
		decimal: 'decimal',
		lowerLetter: 'lower-alpha',
		upperLetter: 'upper-alpha',
		lowerRoman: 'lower-roman',
		upperRoman: 'upper-roman',
		decimalZero: 'decimal-leading-zero',
		aiueo: 'katakana',
		aiueoFullWidth: 'katakana',
		chineseCounting: 'simp-chinese-informal',
		chineseCountingThousand: 'simp-chinese-informal',
		chineseLegalSimplified: 'simp-chinese-formal',
		chosung: 'hangul-consonant',
		ideographDigital: 'cjk-ideographic',
		ideographTraditional: 'cjk-heavenly-stem',
		ideographLegalTraditional: 'trad-chinese-formal',
		ideographZodiac: 'cjk-earthly-branch',
		iroha: 'katakana-iroha',
		irohaFullWidth: 'katakana-iroha',
		japaneseCounting: 'japanese-informal',
		japaneseDigitalTenThousand: 'cjk-decimal',
		japaneseLegal: 'japanese-formal',
		thaiNumbers: 'thai',
		koreanCounting: 'korean-hangul-formal',
		koreanDigital: 'korean-hangul-formal',
		koreanDigital2: 'korean-hanja-informal',
		hebrew1: 'hebrew',
		hebrew2: 'hebrew',
		hindiNumbers: 'devanagari',
		ganada: 'hangul',
		taiwaneseCounting: 'cjk-ideographic',
		taiwaneseCountingThousand: 'cjk-ideographic',
		taiwaneseDigital: 'cjk-decimal',
	};

	return mapping[format] ?? format;
}
