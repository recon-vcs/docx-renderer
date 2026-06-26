import { describe, expect, it } from 'vitest';
import { asArray, escapeClassName, resolvePath, splitPath, uuid } from '../../src/shared/utils';

describe('splitPath', () => {
	it('splits a nested path into folder and file name', () => {
		expect(splitPath('word/document.xml')).toEqual(['word/', 'document.xml']);
	});

	it('handles a bare file name', () => {
		expect(splitPath('document.xml')).toEqual(['', 'document.xml']);
	});
});

describe('resolvePath', () => {
	it('resolves a relative path against a base folder', () => {
		expect(resolvePath('media/image1.png', 'word/')).toBe('word/media/image1.png');
	});

	it('resolves parent navigation', () => {
		expect(resolvePath('../media/image1.png', 'word/theme/')).toBe('word/media/image1.png');
	});
});

describe('escapeClassName', () => {
	it('replaces spaces and dots with dashes and lowercases', () => {
		expect(escapeClassName('Heading 1.Title')).toBe('heading-1-title');
	});

	it('replaces ampersands with "and"', () => {
		expect(escapeClassName('Black & White')).toBe('black-and-white');
	});

	it('throws on undefined input', () => {
		expect(() => escapeClassName(undefined)).toThrow();
	});
});

describe('asArray', () => {
	it('wraps a scalar into an array', () => {
		expect(asArray(1)).toEqual([1]);
	});

	it('returns arrays unchanged', () => {
		const input = [1, 2];
		expect(asArray(input)).toBe(input);
	});
});

describe('uuid', () => {
	it('produces values in UUID v4 format', () => {
		expect(uuid()).toMatch(/^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[0-9a-f]{4}-[0-9a-f]{12}$/i);
	});

	it('produces unique values', () => {
		expect(uuid()).not.toBe(uuid());
	});
});
