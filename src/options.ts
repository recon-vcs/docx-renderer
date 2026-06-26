/** Accepted input formats for a docx document. */
export type DocumentSource = Blob | ArrayBuffer | Uint8Array;

export interface Options {
	breakPages: boolean;
	className: string;

	ignoreFonts: boolean;
	ignoreHeight: boolean;
	ignoreImageWrap: boolean;
	ignoreLastRenderedPageBreak: boolean;
	ignoreTableWrap: boolean;
	ignoreWidth: boolean;

	inWrapper: boolean;

	renderChanges: boolean;
	renderEndnotes: boolean;
	renderFooters: boolean;
	renderFootnotes: boolean;
	renderHeaders: boolean;

	trimXmlDeclaration: boolean;
	useBase64URL: boolean;

	debug: boolean;
}

export const defaultOptions: Options = {
	breakPages: true,
	className: "docx",

	ignoreFonts: false,
	ignoreHeight: false,
	ignoreImageWrap: false,
	ignoreLastRenderedPageBreak: true,
	ignoreTableWrap: true,
	ignoreWidth: false,

	inWrapper: true,

	renderChanges: false,
	renderEndnotes: true,
	renderFooters: true,
	renderFootnotes: true,
	renderHeaders: true,

	trimXmlDeclaration: true,
	useBase64URL: false,

	debug: false,
};

export function resolveOptions(userOptions: Partial<Options> | null = null): Options {
	return { ...defaultOptions, ...userOptions };
}
