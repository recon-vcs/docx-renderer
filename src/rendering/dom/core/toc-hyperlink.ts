import { DomType, OpenXmlElement } from '@docx/ooxml/wordprocessingml/document/model/dom';
import { WmlFieldChar, WmlInstructionText } from '@docx/ooxml/wordprocessingml/document/model/fields';
import { IDomStyle } from '@docx/ooxml/wordprocessingml/document/model/style';

// Word does not paint the "Hyperlink" character style on runs inside a TOC
// field's result: entries display with the TOC paragraph styles only (black,
// no underline), even though the XML carries rStyle=Hyperlink on every entry
// run. A TOC field spans multiple paragraphs (fldChar begin ... end), so this
// walks the body once in document order, tracks open complex fields with
// their instructions, and strips the Hyperlink character style reference from
// runs that sit inside a TOC instruction's result.
export function stripTocHyperlinkStyles(body: OpenXmlElement, styles: IDomStyle[] | undefined): void {
	const hyperlinkStyleIds = new Set(
		(styles ?? [])
			.filter(style => style.name === 'Hyperlink')
			.map(style => style.id),
	);
	if (hyperlinkStyleIds.size === 0) return;

	// Instructions of currently open complex fields, innermost last. Text
	// accumulates until the field's "separate" char, which is early enough:
	// everything before "separate" is instruction, everything after is result.
	const openInstructions: { text: string }[] = [];

	const insideTocField = () => openInstructions.some(f => f.text.trimStart().toUpperCase().startsWith('TOC'));

	const walk = (elem: OpenXmlElement): void => {
		for (const child of elem.children ?? []) {
			switch (child.type) {
				case DomType.ComplexField: {
					const charType = (child as WmlFieldChar).charType;
					if (charType === 'begin') {
						openInstructions.push({ text: '' });
					} else if (charType === 'end') {
						openInstructions.pop();
					}
					break;
				}

				case DomType.Instruction: {
					const current = openInstructions[openInstructions.length - 1];
					if (current) {
						current.text += (child as WmlInstructionText).text ?? '';
					}
					break;
				}

				case DomType.Run:
					if (child.styleName && hyperlinkStyleIds.has(child.styleName) && insideTocField()) {
						delete child.styleName;
					}
					walk(child);
					break;

				default:
					walk(child);
			}
		}
	};

	walk(body);
}
