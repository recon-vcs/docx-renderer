export interface HeaderFooterReferenceLike {
	type: string;
}

export interface HeaderFooterSelectionContext {
	titlePage?: boolean;
	isFirstSectionPage: boolean;
	evenAndOddHeaders?: boolean;
	isEvenPage: boolean;
}

export function resolveHeaderFooterReferences<TRef extends HeaderFooterReferenceLike>(
	currentRefs: readonly TRef[] | undefined,
	previousRefs: readonly TRef[] | undefined,
): TRef[] {
	const resolvedRefs = [...(currentRefs ?? [])];
	const resolvedTypes = new Set(resolvedRefs.map(ref => ref.type));

	for (const ref of previousRefs ?? []) {
		if (!resolvedTypes.has(ref.type)) {
			resolvedRefs.push(ref);
			resolvedTypes.add(ref.type);
		}
	}

	return resolvedRefs;
}

export function selectHeaderFooterReference<TRef extends HeaderFooterReferenceLike>(
	refs: readonly TRef[] | undefined,
	context: HeaderFooterSelectionContext,
): TRef | undefined {
	if (!refs) return undefined;

	if (context.titlePage && context.isFirstSectionPage) {
		return refs.find(ref => ref.type === 'first');
	}

	if (context.evenAndOddHeaders) {
		if (context.isEvenPage) {
			return refs.find(ref => ref.type === 'even');
		}

		return refs.find(ref => ref.type === 'default' || ref.type === 'odd');
	}

	return refs.find(ref => ref.type === 'default');
}
