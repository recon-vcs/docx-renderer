import { describe, expect, it } from 'vitest';
import { isOverflowing } from '../../src/rendering/measurement/overflow-measurer';

describe('isOverflowing', () => {
	it('returns false when scroll height fits client height', () => {
		expect(isOverflowing({ clientHeight: 100, scrollHeight: 100 })).toBe(false);
	});

	it('returns true when scroll height exceeds client height', () => {
		expect(isOverflowing({ clientHeight: 100, scrollHeight: 101 })).toBe(true);
	});
});
