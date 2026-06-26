import { Part } from '@docx/opc/parts/part';
import { OpenXmlPackage } from '@docx/opc/package/open-xml-package';
import { DocumentParser } from '@docx/ooxml/wordprocessingml/parsing/document-parser';
import { WmlComment } from "./elements";
import * as _ from 'lodash-es';

export class CommentsPart extends Part {
	protected _documentParser: DocumentParser;

	comments: WmlComment[]
	commentMap: Record<string, WmlComment>;

	constructor(pkg: OpenXmlPackage, path: string, parser: DocumentParser) {
		super(pkg, path);
		this._documentParser = parser;
	}

	parseXml(root: Element) {
		this.comments = this._documentParser.parseComments(root);
		this.commentMap = _.keyBy(this.comments, 'id');
	}
}
