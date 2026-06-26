import { describe, expect, it } from 'vitest';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';

const repoRoot = process.cwd();

describe('public API docs', () => {
	it('documents the current render API without stale renderAsync references', () => {
		const readme = readFileSync(join(repoRoot, 'README.md'), 'utf8');

		expect(readme).toContain('RenderResult');
		expect(readme).toContain('dispose()');
		expect(readme).not.toContain('renderAsync');
		expect(readme).not.toContain('docs/renderAsync.html');
	});
});
