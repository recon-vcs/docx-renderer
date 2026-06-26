import { describe, expect, it } from 'vitest';
import { Length } from '../../src/shared/measurement/length';

describe('Length', () => {
	it('parses a pt value', () => {
		const length = Length.parse('12.5pt');
		expect(length.value).toBe(12.5);
		expect(length.type).toBe('pt');
	});

	it('parses a px value', () => {
		const length = Length.parse('3px');
		expect(length.value).toBe(3);
		expect(length.type).toBe('px');
	});

	it('creates from a string', () => {
		const length = Length.from('10pt');
		expect(length).toBeInstanceOf(Length);
		expect(length?.value).toBe(10);
	});

	it('passes through an existing Length', () => {
		const original = new Length(4, 'px');
		expect(Length.from(original)).toBe(original);
	});

	it('returns null for unsupported input', () => {
		expect(Length.from(42)).toBeNull();
	});

	it('adds lengths of the same unit', () => {
		const sum = new Length(1, 'pt').add(new Length(2, 'pt'));
		expect(sum.value).toBe(3);
		expect(sum.type).toBe('pt');
	});

	it('subtracts lengths of the same unit', () => {
		const diff = new Length(5, 'px').minus(new Length(2, 'px'));
		expect(diff.value).toBe(3);
	});

	it('throws when mixing units', () => {
		expect(() => new Length(1, 'pt').add(new Length(1, 'px'))).toThrow();
	});

	it('multiplies and divides by scalars', () => {
		expect(new Length(2, 'pt').multiply(3).value).toBe(6);
		expect(new Length(6, 'pt').divide(3).value).toBe(2);
	});

	it('formats with two decimals and unit', () => {
		expect(new Length(1.005, 'pt').toString()).toBe('1.00pt');
		expect(new Length(7).toString()).toBe('7.00');
	});
});
