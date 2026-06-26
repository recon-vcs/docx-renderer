import { appendChildren as appendChildrenUtil, Overflow, ChildrenType } from '@docx/rendering/dom/core/dom-utils';

export interface OverflowMetrics {
	clientHeight: number;
	scrollHeight: number;
}

export function isOverflowing(metrics: OverflowMetrics): boolean {
	return metrics.clientHeight < metrics.scrollHeight;
}

export function measureElementOverflow(el: HTMLElement): boolean {
	const currentOverflow = getComputedStyle(el).overflow;
	if (!currentOverflow || currentOverflow === 'visible') {
		el.style.overflow = 'hidden';
	}

	const overflow = isOverflowing({
		clientHeight: el.clientHeight,
		scrollHeight: el.scrollHeight,
	});

	el.style.overflow = currentOverflow;
	return overflow;
}

export interface PageOverflowState {
	isSplit: boolean;
	contentElement?: HTMLElement;
	checkingOverflow?: boolean;
}

/** Appends children to parent, then measures overflow against the page content element. */
export function appendAndMeasure(
	parent: HTMLElement | Text,
	children: ChildrenType,
	pageState: PageOverflowState,
): Overflow {
	appendChildrenUtil(parent, children);
	if (pageState.isSplit) return Overflow.UNKNOWN;
	if (pageState.checkingOverflow && pageState.contentElement) {
		return measureElementOverflow(pageState.contentElement) ? Overflow.TRUE : Overflow.FALSE;
	}
	return Overflow.UNKNOWN;
}

/** Infer parent overflow status from an array of child overflow results. */
export function inferOverflow(overflows: Overflow[]): Overflow {
	if (overflows.length === 0) return Overflow.FALSE;
	const overflowStatus: Overflow[] = [Overflow.FULL, Overflow.SELF, Overflow.TRUE, Overflow.PART];
	if (overflows.every(o => o === Overflow.PART)) return Overflow.PART;
	if (overflows.every(o => overflowStatus.includes(o))) return Overflow.FULL;
	if (overflows.every(o => o === Overflow.UNKNOWN)) return Overflow.UNKNOWN;
	if (overflows.every(o => [Overflow.FALSE, Overflow.UNKNOWN, Overflow.IGNORE].includes(o))) return Overflow.FALSE;
	if (overflows.some(o => overflowStatus.includes(o))) return Overflow.PART;
	return Overflow.UNKNOWN;
}
