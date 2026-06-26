import { describe, expect, it } from 'vitest';
import {
	interpretRegionBreakBefore,
	needsParityBlankPage,
	normalizeRegionBreakBefore,
	pageMatchesParity,
} from '../../src/rendering/pagination/model/breaks';

describe('layout break interpretation', () => {
	it('normalizes missing break values to none', () => {
		expect(normalizeRegionBreakBefore(undefined)).toBe('none');
		expect(normalizeRegionBreakBefore(null)).toBe('none');
		expect(normalizeRegionBreakBefore('page')).toBe('page');
	});

	it('keeps none and column breaks on the current physical page', () => {
		expect(interpretRegionBreakBefore('none')).toEqual({
			breakBefore: 'none',
			startsNewPage: false,
			targetParity: null,
		});

		expect(interpretRegionBreakBefore('column')).toEqual({
			breakBefore: 'column',
			startsNewPage: false,
			targetParity: null,
		});
	});

	it('interprets page breaks as new physical pages without parity', () => {
		expect(interpretRegionBreakBefore('page')).toEqual({
			breakBefore: 'page',
			startsNewPage: true,
			targetParity: null,
		});
	});

	it('interprets odd and even page breaks with target parity', () => {
		expect(interpretRegionBreakBefore('evenPage')).toEqual({
			breakBefore: 'evenPage',
			startsNewPage: true,
			targetParity: 'even',
		});

		expect(interpretRegionBreakBefore('oddPage')).toEqual({
			breakBefore: 'oddPage',
			startsNewPage: true,
			targetParity: 'odd',
		});
	});

	it('checks physical page parity', () => {
		expect(pageMatchesParity(1, 'odd')).toBe(true);
		expect(pageMatchesParity(2, 'odd')).toBe(false);
		expect(pageMatchesParity(2, 'even')).toBe(true);
		expect(pageMatchesParity(3, 'even')).toBe(false);
	});

	it('detects when a blank page is needed before a parity break', () => {
		expect(needsParityBlankPage(2, 'even')).toBe(false);
		expect(needsParityBlankPage(3, 'even')).toBe(true);
		expect(needsParityBlankPage(3, 'odd')).toBe(false);
		expect(needsParityBlankPage(4, 'odd')).toBe(true);
		expect(needsParityBlankPage(4, null)).toBe(false);
	});
});
