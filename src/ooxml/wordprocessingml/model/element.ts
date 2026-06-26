import { SectionType } from '@docx/ooxml/wordprocessingml/document/model/section';

export enum DomType {
	Document = "document",
	Paragraph = "paragraph",
	Run = "run",
	Break = "break",
	LastRenderedPageBreak = "lastRenderedPageBreak",
	SectionBreak = "sectionBreak",
	NoBreakHyphen = "noBreakHyphen",
	Table = "table",
	Row = "row",
	Cell = "cell",
	Hyperlink = "hyperlink",
	Drawing = "drawing",
	Image = "image",
	Text = "text",
	Character = "character",
	Tab = "tab",
	Symbol = "symbol",
	BookmarkStart = "bookmarkStart",
	BookmarkEnd = "bookmarkEnd",
	Footer = "footer",
	Header = "header",
	FootnoteReference = "footnoteReference",
	EndnoteReference = "endnoteReference",
	Footnotes = "footnotes",
	Footnote = "footnote",
	Endnotes = "endnotes",
	Endnote = "endnote",
	SimpleField = "simpleField",
	ComplexField = "complexField",
	Instruction = "instruction",
	VmlPicture = "vmlPicture",
	Shape = "shape",
	MmlMath = "mmlMath",
	MmlMathParagraph = "mmlMathParagraph",
	MmlFraction = "mmlFraction",
	MmlFunction = "mmlFunction",
	MmlFunctionName = "mmlFunctionName",
	MmlNumerator = "mmlNumerator",
	MmlDenominator = "mmlDenominator",
	MmlRadical = "mmlRadical",
	MmlBase = "mmlBase",
	MmlDegree = "mmlDegree",
	MmlSuperscript = "mmlSuperscript",
	MmlSubscript = "mmlSubscript",
	MmlPreSubSuper = "mmlPreSubSuper",
	MmlSubArgument = "mmlSubArgument",
	MmlSuperArgument = "mmlSuperArgument",
	MmlNary = "mmlNary",
	MmlDelimiter = "mmlDelimiter",
	MmlRun = "mmlRun",
	MmlEquationArray = "mmlEquationArray",
	MmlLimit = "mmlLimit",
	MmlLimitLower = "mmlLimitLower",
	MmlMatrix = "mmlMatrix",
	MmlMatrixRow = "mmlMatrixRow",
	MmlBox = "mmlBox",
	MmlBar = "mmlBar",
	MmlGroupChar = "mmlGroupChar",
	VmlElement = "vmlElement",
	Inserted = "inserted",
	Deleted = "deleted",
	DeletedText = "deletedText",
	Comment = "comment",
	CommentReference = "commentReference",
	CommentRangeStart = "commentRangeStart",
	CommentRangeEnd = "commentRangeEnd",
}

export enum MathDomType {
	Base = "mmlBase",
	Bar = "mmlBar",
	Box = "mmlBox",
	Delimiter = "mmlDelimiter",
	Degree = "mmlDegree",
	Denominator = "mmlDenominator",
	Function = "mmlFunction",
	FunctionName = "mmlFunctionName",
	Fraction = "mmlFraction",
	GroupChar = "mmlGroupChar",
	Limit = "mmlLimit",
	LimitLower = "mmlLimitLower",
	Matrix = "mmlMatrix",
	MatrixRow = "mmlMatrixRow",
	Math = "mmlMath",
	MathParagraph = "mmlMathParagraph",
	Nary = "mmlNary",
	Numerator = "mmlNumerator",
	PreSubSuper = "mmlPreSubSuper",
	Radical = "mmlRadical",
	SubArgument = "mmlSubArgument",
	Subscript = "mmlSubscript",
	Superscript = "mmlSuperscript",
}

export enum BreakType {
	Column = "column",
	Page = "page",
	TextWrapping = "textWrapping",
}

export enum WrapType {
	Inline = "Inline",
	None = "None",
	TopAndBottom = "TopAndBottom",
	Tight = "Tight",
	Through = "Through",
	Square = "Square",
	Polygon = "Polygon",
}

export interface OpenXmlElement {
	uuid?: string;
	type: DomType;
	children?: OpenXmlElement[];
	cssStyle?: Record<string, any>;
	props?: Record<string, any>;
	breakIndex?: Set<number>;
	styleName?: string;
	className?: string;
	parent?: OpenXmlElement;
}

export abstract class OpenXmlElementBase implements OpenXmlElement {
	type: DomType;
	children?: OpenXmlElement[] = [];
	cssStyle?: Record<string, any> = {};
	props?: Record<string, any>;
	breakIndex?: Set<number>;
	styleName?: string;
	className?: string;
	parent?: OpenXmlElement;
}

export interface WmlHyperlink extends OpenXmlElement {
	id?: string;
	href?: string;
}

export interface WmlNoteReference extends OpenXmlElement {
	id: string;
}

export interface WmlBreak extends OpenXmlElement {
	break: BreakType;
}

export interface WmlSectionBreak extends OpenXmlElement {
	break: SectionType;
}

export interface WmlLastRenderedPageBreak extends OpenXmlElement {}

export interface WmlText extends OpenXmlElement {
	text: string;
}

export interface WmlCharacter extends OpenXmlElement {
	char: string;
}

export interface WmlSymbol extends OpenXmlElement {
	font: string;
	char: string;
}
