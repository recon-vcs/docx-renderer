import { describe, expect, it } from 'vitest';
import { isOverflowing, measureElementOverflow } from '../../src/rendering/measurement/overflow-measurer';

describe('isOverflowing', () => {
	it('returns false when scroll height fits client height', () => {
		expect(isOverflowing({ clientHeight: 100, scrollHeight: 100 })).toBe(false);
	});

	it('returns true when scroll height exceeds client height', () => {
		expect(isOverflowing({ clientHeight: 100, scrollHeight: 101 })).toBe(true);
	});
});

describe('measureElementOverflow', () => {
	it('excludes skipped overflow elements while preserving their display style', () => {
		const container = document.createElement('div');
		const skipped = document.createElement('span');
		skipped.dataset.overflow = 'skip';
		skipped.style.display = 'block';
		container.appendChild(skipped);
		document.body.appendChild(container);

		Object.defineProperty(container, 'clientHeight', { configurable: true, value: 100 });
		Object.defineProperty(container, 'scrollHeight', {
			configurable: true,
			get: () => skipped.style.display === 'none' ? 100 : 200,
		});

		expect(measureElementOverflow(container)).toBe(false);
		expect(skipped.style.display).toBe('block');

		container.remove();
	});
});
