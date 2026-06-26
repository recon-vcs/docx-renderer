import { readFile } from 'node:fs/promises';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { DocumentParser } from '../../src/ooxml/wordprocessingml/parsing/document-parser';
import { WordDocument } from '../../src/word-document';

const PARSE_OPTIONS = {
	trimXmlDeclaration: true,
	useBase64URL: true,
	debug: false,
};

async function loadFixture(name: string): Promise<WordDocument> {
	// import.meta.url is an http URL under the jsdom environment; resolve from cwd instead
	const data = await readFile(join(process.cwd(), 'tests', 'fixtures', `${name}.docx`));
	return WordDocument.load(data, new DocumentParser({ debug: false }), PARSE_OPTIONS);
}

describe('extended properties (docProps/app.xml)', () => {
	it('loads the extended properties part', async () => {
		const doc = await loadFixture('extended-props');

		expect(doc.extendedPropsPart).toBeTruthy();

		const props = doc.extendedPropsPart.props;
		expect(props.application).toBe('Microsoft Office Word');
		expect(props.appVersion).toBe('16.0000');
		expect(props.template).toBe('Normal.dotm');
		expect(props.pages).toBe(1);
		expect(props.words).toBe(68);
		expect(props.characters).toBe(393);
		expect(props.lines).toBe(3);
		expect(props.paragraphs).toBe(1);
		expect(props.company).toBe('');
	});

	it('loads the main document part', async () => {
		const doc = await loadFixture('extended-props');

		expect(doc.documentPart).toBeTruthy();
		expect(doc.documentPart.body).toBeTruthy();
	});
});
