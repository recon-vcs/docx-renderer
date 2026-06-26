import { describe, expect, it } from 'vitest';
import { readdirSync, readFileSync, statSync } from 'node:fs';
import { join, relative } from 'node:path';

const repoRoot = process.cwd();

function listFiles(dir: string, predicate: (path: string) => boolean): string[] {
	return readdirSync(dir)
		.flatMap(entry => {
			const path = join(dir, entry);
			const stat = statSync(path);
			if (stat.isDirectory()) return listFiles(path, predicate);
			return predicate(path) ? [path] : [];
		});
}

function productionTsFiles(dir: string): string[] {
	return listFiles(join(repoRoot, dir), path => path.endsWith('.ts'));
}

function filesMatching(files: string[], pattern: RegExp): string[] {
	return files
		.filter(path => pattern.test(readFileSync(path, 'utf8')))
		.map(path => relative(repoRoot, path));
}

describe('architecture boundaries', () => {
	it('keeps OOXML model and parser layers independent from rendering', () => {
		const offenders = filesMatching(
			productionTsFiles('src/ooxml'),
			/from\s+['"]@docx\/rendering\//,
		);

		expect(offenders).toEqual([]);
	});

	it('keeps pagination model independent from DOM rendering helpers', () => {
		const offenders = filesMatching(
			productionTsFiles('src/rendering/pagination'),
			/from\s+['"]@docx\/rendering\/dom\//,
		);

		expect(offenders).toEqual([]);
	});

	it('keeps measurement independent from DOM mutation helpers', () => {
		const offenders = filesMatching(
			productionTsFiles('src/rendering/measurement'),
			/from\s+['"]@docx\/rendering\/dom\//,
		);

		expect(offenders).toEqual([]);
	});

	it('keeps sub-parsers on narrow parser contexts', () => {
		const parserFiles = productionTsFiles('src/ooxml/wordprocessingml/parsing')
			.filter(path => !path.endsWith('document-parser.ts') && !path.endsWith('parse-context.ts'));
		const offenders = filesMatching(
			parserFiles,
			/import\s+type\s+\{\s*ParseContext\s*\}/,
		);

		expect(offenders).toEqual([]);
	});
});
