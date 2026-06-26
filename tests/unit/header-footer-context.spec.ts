import { describe, expect, it } from 'vitest';
import {
	resolveHeaderFooterReferences,
	selectHeaderFooterReference,
} from '../../src/rendering/pagination/context/header-footer-context';

interface TestRef {
	id: string;
	type: string;
}

function ref(type: string, id = type): TestRef {
	return { type, id };
}

describe('resolveHeaderFooterReferences', () => {
	it('keeps current section references before inherited references', () => {
		const current = [ref('default', 'current-default')];
		const previous = [ref('first', 'previous-first'), ref('even', 'previous-even')];

		const resolved = resolveHeaderFooterReferences(current, previous);

		expect(resolved).toEqual([
			ref('default', 'current-default'),
			ref('first', 'previous-first'),
			ref('even', 'previous-even'),
		]);
	});

	it('does not inherit a previous reference when current section has the same type', () => {
		const current = [ref('default', 'current-default')];
		const previous = [ref('default', 'previous-default'), ref('even', 'previous-even')];

		const resolved = resolveHeaderFooterReferences(current, previous);

		expect(resolved).toEqual([
			ref('default', 'current-default'),
			ref('even', 'previous-even'),
		]);
	});
});

describe('selectHeaderFooterReference', () => {
	it('selects the first-page reference for a title page on the first section page', () => {
		const refs = [ref('default'), ref('first')];

		const selected = selectHeaderFooterReference(refs, {
			titlePage: true,
			isFirstSectionPage: true,
			isEvenPage: false,
		});

		expect(selected).toEqual(ref('first'));
	});

	it('selects the even reference when even and odd references are enabled', () => {
		const refs = [ref('default'), ref('even')];

		const selected = selectHeaderFooterReference(refs, {
			isFirstSectionPage: false,
			evenAndOddHeaders: true,
			isEvenPage: true,
		});

		expect(selected).toEqual(ref('even'));
	});

	it('selects the default reference for odd pages when even and odd references are enabled', () => {
		const refs = [ref('default'), ref('even')];

		const selected = selectHeaderFooterReference(refs, {
			isFirstSectionPage: false,
			evenAndOddHeaders: true,
			isEvenPage: false,
		});

		expect(selected).toEqual(ref('default'));
	});

	it('selects the default reference when even and odd references are disabled', () => {
		const refs = [ref('default'), ref('even')];

		const selected = selectHeaderFooterReference(refs, {
			isFirstSectionPage: false,
			isEvenPage: true,
		});

		expect(selected).toEqual(ref('default'));
	});
});
