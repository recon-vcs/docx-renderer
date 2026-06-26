import { OpenXmlElement } from './element';

export interface WmlTableColumn {
	width?: string;
}

export interface WmlTable extends OpenXmlElement {
	columns?: WmlTableColumn[];
	cellStyle?: Record<string, string>;
	colBandSize?: number;
	rowBandSize?: number;
}

export interface WmlTableRow extends OpenXmlElement {
	isHeader?: boolean;
}

export interface WmlTableCell extends OpenXmlElement {
	verticalMerge?: 'restart' | 'continue' | string;
	span?: number;
}
