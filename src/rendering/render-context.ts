/**
 * Unified rendering context that consolidates all callback interfaces used by
 * the sub-renderers (page, inline, math, fields, drawing, element-dispatch).
 *
 * HtmlRendererSync creates one RenderContext per render call and passes it to
 * all sub-renderer functions, replacing the previous pattern of six separate
 * callback-factory methods.
 */

import type { Part } from '@docx/opc/parts/part';
import type { Stage } from 'konva/lib/Stage';
import type { Layer } from 'konva/lib/Layer';
import { OpenXmlElement } from '@docx/ooxml/wordprocessingml/model/element';
import { WmlText } from '@docx/ooxml/wordprocessingml/document/model/dom';
import type { TabStop } from '@docx/ooxml/wordprocessingml/document/model/paragraph';
import type { CommonProperties } from '@docx/ooxml/wordprocessingml/document/model/common';
import type { IDomStyle } from '@docx/ooxml/wordprocessingml/document/model/style';
import type { SectionProperties } from '@docx/ooxml/wordprocessingml/document/model/section';
import type { Options } from '@docx/options';
import type { WordDocument } from '@docx/word-document';
import { Overflow } from '@docx/rendering/measurement/overflow';
import type { ChildrenType } from '@docx/rendering/dom/core/dom-utils';
import type { TableContext } from '@docx/rendering/dom/elements/table-renderer';
import type { Node_DOM } from '@docx/rendering/dom/core/dom-utils';

export interface RenderContext {
	// ── identity ──────────────────────────────────────────────────────────────
	className: string;
	document: WordDocument;
	options: Options;

	// ── session state (read-only from sub-renderers) ─────────────────────────
	currentPart: Part | null;
	konvaStage: Stage;
	konvaLayer: Layer;
	tableCtx: TableContext;
	currentTabs: Array<{ stops?: TabStop[]; span: HTMLElement }>;
	currentFootnoteIds: string[];
	currentEndnoteIds: string[];
	usedHeaderFooterParts: string[];
	evenAndOddHeaders: boolean;
	ignoreWidth: boolean;
	ignoreHeight: boolean;

	// ── element tree helpers ──────────────────────────────────────────────────
	/** Set parent links on the element tree (replaces render-time processElement parent-setting). */
	linkParents(elem: OpenXmlElement): void;
	/** Apply table cell style propagation (and recursively process children). */
	processElement(elem: OpenXmlElement): void;

	// ── core render operations ────────────────────────────────────────────────
	appendChildren(parent: HTMLElement | Text, children: ChildrenType): Promise<Overflow>;
	appendChildrenWithoutOverflow(parent: Element | Text, children: ChildrenType): void;
	runWithoutOverflowChecking<T>(callback: () => Promise<T>): Promise<T>;
	renderChildren(elem: OpenXmlElement, parent: HTMLElement | Element | Text): Promise<Overflow>;
	renderElements(children: OpenXmlElement[], parent: HTMLElement | Element | Text): Promise<Overflow>;
	renderElement(elem: OpenXmlElement, parent?: HTMLElement | Element | Text): Promise<Node_DOM>;
	renderContainer(elem: OpenXmlElement, tagName: keyof HTMLElementTagNameMap, props?: Record<string, any>): Promise<HTMLElement>;
	renderContainerNS(elem: OpenXmlElement, ns: string, tagName: string, props?: Record<string, any>): Promise<Element>;
	renderHeaderFooter(elem: OpenXmlElement, tagName: keyof HTMLElementTagNameMap, parent: HTMLElement): Promise<HTMLElement>;

	// ── style helpers ─────────────────────────────────────────────────────────
	renderClass(elem: OpenXmlElement, output: HTMLElement | Element): void;
	renderStyleValues(style: Record<string, string>, output: HTMLElement): void;
	renderCommonProperties(style: CSSStyleDeclaration, props: CommonProperties): void;

	// ── lookup helpers ────────────────────────────────────────────────────────
	findStyle(styleName: string): IDomStyle;
	numberingClass(id: string, level: number): string;
	findExternalRelation(id: string): { target?: string } | undefined;
	findComment(id: string): { id: string; author: string; date: string } | undefined;

	// ── page state accessors ──────────────────────────────────────────────────
	currentPageIsSplit(): boolean;
	currentSectionProperties(): SectionProperties;
	currentPageNumber(): number;
	pageCount(): number;

	// ── field helpers ─────────────────────────────────────────────────────────
	resolveFieldRuns(runs: OpenXmlElement[]): OpenXmlElement[];
	renderText(elem: WmlText, parent: HTMLElement): Promise<Node>;
	renderChanges(): boolean;

	// ── session mutations ─────────────────────────────────────────────────────
	setCurrentPart(part: Part | null): void;
}
