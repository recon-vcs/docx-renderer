import { OpenXmlElement } from '../document/dom';
import { CommonProperties } from '../document/common';
import { RunProperties } from '../document/run';

// 设置元素style样式
export function renderStyleValues(style: Record<string, string>, output: HTMLElement): void {
	for (const k in style) {
		if (k.startsWith('$')) {
			output.setAttribute(k.slice(1), style[k]);
		} else {
			output.style[k] = style[k];
		}
	}
}

export function renderCommonProperties(style: any, props: CommonProperties | null): void {
	if (props == null) return;

	if (props.color) {
		style['color'] = props.color;
	}

	if (props.fontSize) {
		style['font-size'] = props.fontSize;
	}
}

export function renderRunProperties(style: any, props: RunProperties | null): void {
	renderCommonProperties(style, props);
}

// 添加class类名
export function renderClass(
	input: OpenXmlElement,
	output: HTMLElement | Element,
	processStyleName: (name: string) => string,
): void {
	if (input.className) {
		output.className = input.className;
	}

	if (input.styleName) {
		output.classList.add(processStyleName(input.styleName));
	}
}
