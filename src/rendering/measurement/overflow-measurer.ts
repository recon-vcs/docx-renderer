import { Overflow } from './overflow';

export interface OverflowMetrics {
	clientHeight: number;
	scrollHeight: number;
}

export function isOverflowing(metrics: OverflowMetrics): boolean {
	return metrics.clientHeight < metrics.scrollHeight;
}

export function measureElementOverflow(el: HTMLElement): boolean {
	const currentOverflow = getComputedStyle(el).overflow;
	const skippedElements = Array.from(el.querySelectorAll<HTMLElement>('[data-overflow="skip"]'));
	const skippedDisplayValues = skippedElements.map(item => item.style.display);

	for (const item of skippedElements) {
		item.style.display = 'none';
	}

	if (!currentOverflow || currentOverflow === 'visible') {
		el.style.overflow = 'hidden';
	}

	const overflow = isOverflowing({
		clientHeight: el.clientHeight,
		scrollHeight: el.scrollHeight,
	});

	el.style.overflow = currentOverflow;
	for (let i = 0; i < skippedElements.length; i++) {
		skippedElements[i].style.display = skippedDisplayValues[i];
	}
	return overflow;
}

export interface PageOverflowState {
	isSplit: boolean;
	contentElement?: HTMLElement;
	checkingOverflow?: boolean;
}

export function measurePageOverflow(pageState: PageOverflowState): Overflow {
	if (pageState.isSplit) return Overflow.UNCHECKED;
	if (pageState.checkingOverflow && pageState.contentElement) {
		return measureElementOverflow(pageState.contentElement) ? Overflow.SELF : Overflow.NONE;
	}
	return Overflow.UNCHECKED;
}

/** Infer parent overflow status from an array of child overflow results. */
export function inferOverflow(overflows: Overflow[]): Overflow {
	if (overflows.length === 0) return Overflow.NONE;
	const overflowStatus: Overflow[] = [Overflow.FULL, Overflow.SELF, Overflow.PARTIAL];
	if (overflows.every(o => o === Overflow.PARTIAL)) return Overflow.PARTIAL;
	if (overflows.every(o => overflowStatus.includes(o))) return Overflow.FULL;
	if (overflows.every(o => o === Overflow.UNCHECKED)) return Overflow.UNCHECKED;
	if (overflows.every(o => [Overflow.NONE, Overflow.UNCHECKED, Overflow.SKIP].includes(o))) return Overflow.NONE;
	if (overflows.some(o => overflowStatus.includes(o))) return Overflow.PARTIAL;
	return Overflow.UNCHECKED;
}
