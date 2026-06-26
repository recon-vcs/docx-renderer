import { ParagraphProperties } from "./paragraph";
import { RunProperties } from "./run";

export interface IDomStyle {
	aliases?: string[];
	autoRedefine?: boolean;
	basedOn?: string;
	customStyle?: boolean;
	cssName?: string;
	hidden?: boolean;
	id: string;
	isDefault?: boolean;
	linked?: string;
	locked?: boolean;
	name?: string;
	next?: string;
	paragraphProps?: ParagraphProperties;
	personal?: boolean;
	personalCompose?: boolean;
	personalReply?: boolean;
	primaryStyle?: boolean;
	rsid?: number;
	rulesets: Ruleset[];
	runProps?: RunProperties;
	semiHidden?: boolean;
	type: string;
	label?: string;
	uiPriority?: number;
	unhideWhenUsed?: boolean;
}

export interface Ruleset {
	target: string;
	modifier?: string;
	declarations: Record<string, string>;
}
