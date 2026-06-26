import { describe, expect, it } from 'vitest';
import { readdirSync, readFileSync, statSync } from 'node:fs';
import { join, relative } from 'node:path';

const repoRoot = process.cwd();
const maxProductionFileLines = 800;

function lineCount(path: string): number {
	const text = readFileSync(path, 'utf8');
	return text.length === 0 ? 0 : text.split(/\r?\n/).length;
}

function listTypeScriptFiles(dir: string): string[] {
	return readdirSync(dir)
		.flatMap(entry => {
			const path = join(dir, entry);
			const stat = statSync(path);

			if (stat.isDirectory()) {
				return listTypeScriptFiles(path);
			}

			return entry.endsWith('.ts') ? [path] : [];
		});
}

describe('production source line counts', () => {
	it(`keeps src TypeScript files at or below ${maxProductionFileLines} lines`, () => {
		const oversized = listTypeScriptFiles(join(repoRoot, 'src'))
			.map(path => ({
				path: relative(repoRoot, path),
				lines: lineCount(path),
			}))
			.filter(file => file.lines > maxProductionFileLines);

		expect(oversized).toEqual([]);
	});
});
