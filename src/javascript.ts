import { Length } from '@docx/ooxml/wordprocessingml/document/model/common';
import { TabStop } from '@docx/ooxml/wordprocessingml/document/model/paragraph';

const defaultTab: TabStop = { position: 0, leader: "none", style: "left" };
const maxTabs = 50;

// calculate pixel to point ratio.
// ratio = point / pixel;
export function computePointToPixelRatio(container: HTMLElement = document.body) {
	// temp element
	const temp = document.createElement("div");
	// set width 100pt
	temp.style.width = '100pt';
	// append to body
	container.appendChild(temp);
	//
	const ratio = 100 / temp.offsetWidth;
	container.removeChild(temp);

	return ratio
}

export function updateTabStop(element: HTMLElement, tabs: TabStop[], defaultTabSize: Length, pixelToPoint: number = 72 / 96) {
	// element's parent paragraph
	const oParagraph: HTMLParagraphElement = element.closest("p");
	// element rect
	const elementRect: DOMRect = element.getBoundingClientRect();
	// paragraph rect
	const paragraphRect: DOMRect = oParagraph.getBoundingClientRect();
	// paragraph style
	const paragraphStyle: CSSStyleDeclaration = getComputedStyle(oParagraph);
	// tabStops with order from small to large
	const tabStops: TabStop[] = tabs?.length > 0 ? tabs.sort((a, b) => a.position - b.position) : [defaultTab];
	// last tab
	const lastTab = tabStops[tabStops.length - 1];
	// paragraph width in point
	const paragraphWidth = paragraphRect.width * pixelToPoint;
	// default tab stop size in point
	const size = parseFloat(defaultTabSize);
	// last tab stop position
	let position = lastTab.position + size;
	// TODO
	if (position < paragraphWidth) {
		for (; position < paragraphWidth && tabStops.length < maxTabs; position += size) {
			tabStops.push({ ...defaultTab, position: position });
		}
	}
	// paragraph left offset in point
	const marginLeft = parseFloat(paragraphStyle.marginLeft);
	const paragraphOffset = paragraphRect.left + marginLeft;
	// element actual left offset in point
	const left = (elementRect.left - paragraphOffset) * pixelToPoint;
	// find first tab stop
	const tab = tabStops.find(tab => tab.style != "clear" && tab.position > left);
	// no tab stop
	if (tab == null) {
		return;
	}

	let width: number = 1;

	if (tab.style == "right" || tab.style == "center") {
		// tab stop elements
		const tabStopElements = Array.from(oParagraph.querySelectorAll(`.${element.className}`));
		// next tab stop index
		const nextIndex = tabStopElements.indexOf(element) + 1;
		// create range
		const range = document.createRange();
		// set the start position of the Range
		range.setStartBefore(element);

		if (nextIndex < tabStopElements.length) {
			// set the end position of the Range，before the next tab stop element
			// 在next tabStop元素之前,设置 Range 的终点位置
			range.setEndBefore(tabStopElements[nextIndex]);
		} else {
			// set the end position of the Range，after the paragraph.
			// 在段落元素之后,设置 Range 的终点位置
			range.setEndAfter(oParagraph);
		}

		const mul = tab.style === "center" ? 0.5 : 1;
		// range rect
		const rangeRect = range.getBoundingClientRect();
		// offset
		const offset = rangeRect.left + mul * rangeRect.width - (paragraphRect.left - marginLeft);

		width = tab.position - offset * pixelToPoint;
	} else {
		width = tab.position - left;
	}

	element.innerHTML = "&nbsp;";
	element.style.textDecoration = "inherit";
	element.style.wordSpacing = `${width.toFixed(0)}pt`;

	switch (tab.leader) {
		case "dot":
		case "middleDot":
			element.style.textDecorationLine = "underline";
			element.style.textDecorationStyle = "dotted";
			break;

		case "hyphen":
			element.style.textDecorationLine = "underline";
			element.style.textDecorationStyle = "dashed";
			break;

		case "heavy":
		case "underscore":
			element.style.textDecorationLine = "underline";
			element.style.textDecorationStyle = "solid";
			break;

		case "none":
		default:
			element.style.textDecorationLine = "none";
			break;
	}
}
