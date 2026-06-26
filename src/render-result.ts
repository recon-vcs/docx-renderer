import { WordDocument } from '@docx/word-document';

export interface PageHandle {
	index: number;
	element: HTMLElement;
	blockPaths: string[];
}

export interface SourceMap {
	elementsFor(blockPath: string): HTMLElement[];
	pathFor(element: HTMLElement): string | null;
}

export interface AttachOptions {
	className?: string;
	placement?: 'right' | 'left';
	offsetX?: number;
	offsetY?: number;
}

export interface OverlayHandle {
	element: HTMLElement;
	anchor: HTMLElement;
	update(): void;
	dispose(): void;
}

export interface OverlayLayer {
	attach(anchor: HTMLElement, content: HTMLElement, opts?: AttachOptions): OverlayHandle;
	clear(): void;
	dispose(): void;
}

export interface RenderResult {
	document: WordDocument;
	pages: PageHandle[];
	sourceMap: SourceMap;
	overlay: OverlayLayer;
}

class DomSourceMap implements SourceMap {
	private elementsByPath = new Map<string, HTMLElement[]>();

	constructor(private readonly root: HTMLElement) {
		const elements = Array.from(root.querySelectorAll<HTMLElement>('[data-vellum-path]'));
		for (const element of elements) {
			const path = element.dataset.vellumPath;
			if (!path) continue;
			const existing = this.elementsByPath.get(path);
			if (existing) {
				existing.push(element);
			} else {
				this.elementsByPath.set(path, [element]);
			}
		}
	}

	elementsFor(blockPath: string): HTMLElement[] {
		const exact = this.elementsByPath.get(blockPath);
		if (exact?.length) return [...exact];

		const bodyPath = blockPath.match(/^body\/\d+/)?.[0];
		if (!bodyPath || bodyPath === blockPath) return [];

		return [...(this.elementsByPath.get(bodyPath) ?? [])];
	}

	pathFor(element: HTMLElement): string | null {
		const target = element.closest<HTMLElement>('[data-vellum-path]');
		if (!target || !this.root.contains(target)) return null;
		return target.dataset.vellumPath ?? null;
	}
}

class DomOverlayLayer implements OverlayLayer {
	private readonly handles = new Set<DomOverlayHandle>();
	private readonly resizeObserver: ResizeObserver | null;
	private readonly mutationObserver: MutationObserver;
	private frameId: number | null = null;

	constructor(private readonly root: HTMLElement, private readonly pages: PageHandle[]) {
		this.resizeObserver = typeof ResizeObserver === 'undefined'
			? null
			: new ResizeObserver(() => this.scheduleUpdate());
		this.mutationObserver = new MutationObserver(() => this.scheduleUpdate());

		for (const page of pages) {
			page.element.style.position ||= 'relative';
			this.ensurePageOverlay(page.element);
			this.resizeObserver?.observe(page.element);
		}

		this.mutationObserver.observe(root, {
			childList: true,
			subtree: true,
			attributes: true,
			attributeFilter: ['style', 'class'],
		});
		window.addEventListener('resize', this.onWindowResize);
	}

	attach(anchor: HTMLElement, content: HTMLElement, opts: AttachOptions = {}): OverlayHandle {
		const handle = new DomOverlayHandle(this, anchor, content, opts);
		this.handles.add(handle);
		handle.update();
		return handle;
	}

	clear(): void {
		for (const handle of [...this.handles]) {
			handle.dispose();
		}
		this.handles.clear();
	}

	remove(handle: DomOverlayHandle): void {
		this.handles.delete(handle);
	}

	overlayFor(anchor: HTMLElement): HTMLElement | null {
		const page = this.pageFor(anchor);
		return page ? this.ensurePageOverlay(page) : null;
	}

	positionFor(anchor: HTMLElement, placement: 'right' | 'left', offsetX: number, offsetY: number): { left: number; top: number } | null {
		const page = this.pageFor(anchor);
		if (!page) return null;

		const pageRect = page.getBoundingClientRect();
		const anchorRect = anchor.getBoundingClientRect();
		const scale = page.offsetWidth > 0 ? pageRect.width / page.offsetWidth : 1;
		const normalizedScale = scale || 1;
		const anchorLeft = placement === 'left' ? anchorRect.left - pageRect.left : anchorRect.right - pageRect.left;

		return {
			left: anchorLeft / normalizedScale + offsetX,
			top: (anchorRect.top - pageRect.top) / normalizedScale + offsetY,
		};
	}

	dispose(): void {
		this.clear();
		this.resizeObserver?.disconnect();
		this.mutationObserver.disconnect();
		window.removeEventListener('resize', this.onWindowResize);
		if (this.frameId !== null) {
			cancelAnimationFrame(this.frameId);
			this.frameId = null;
		}
	}

	private readonly onWindowResize = () => this.scheduleUpdate();

	private scheduleUpdate(): void {
		if (this.frameId !== null) return;
		this.frameId = requestAnimationFrame(() => {
			this.frameId = null;
			for (const handle of this.handles) {
				handle.update();
			}
		});
	}

	private pageFor(anchor: HTMLElement): HTMLElement | null {
		return this.pages.find((page) => page.element.contains(anchor))?.element ?? null;
	}

	private ensurePageOverlay(page: HTMLElement): HTMLElement {
		let overlay = page.querySelector<HTMLElement>(':scope > .vellum-overlay-layer');
		if (overlay) return overlay;

		overlay = document.createElement('div');
		overlay.className = 'vellum-overlay-layer';
		overlay.style.position = 'absolute';
		overlay.style.inset = '0';
		overlay.style.pointerEvents = 'none';
		overlay.style.zIndex = '5';
		page.appendChild(overlay);
		return overlay;
	}
}

class DomOverlayHandle implements OverlayHandle {
	readonly element: HTMLElement;

	constructor(
		private readonly layer: DomOverlayLayer,
		readonly anchor: HTMLElement,
		content: HTMLElement,
		private readonly opts: AttachOptions,
	) {
		this.element = document.createElement('div');
		this.element.className = opts.className ?? '';
		this.element.style.position = 'absolute';
		this.element.style.pointerEvents = 'auto';
		this.element.appendChild(content);
		this.layer.overlayFor(anchor)?.appendChild(this.element);
	}

	update(): void {
		if (!this.anchor.isConnected) {
			this.dispose();
			return;
		}

		const overlay = this.layer.overlayFor(this.anchor);
		if (!overlay) {
			this.element.remove();
			return;
		}

		if (this.element.parentElement !== overlay) {
			overlay.appendChild(this.element);
		}

		const position = this.layer.positionFor(
			this.anchor,
			this.opts.placement ?? 'right',
			this.opts.offsetX ?? 8,
			this.opts.offsetY ?? 0,
		);
		if (!position) return;

		this.element.style.left = `${position.left}px`;
		this.element.style.top = `${position.top}px`;
	}

	dispose(): void {
		this.element.remove();
		this.layer.remove(this);
	}
}

export function createRenderResult(document: WordDocument, bodyContainer: HTMLElement, className: string): RenderResult {
	const pageElements = Array.from(bodyContainer.querySelectorAll<HTMLElement>(`section.${className}`));
	const pages = pageElements.map<PageHandle>((element, index) => ({
		index,
		element,
		blockPaths: Array.from(
			new Set(
				Array.from(element.querySelectorAll<HTMLElement>('[data-vellum-path]'))
					.map((node) => node.dataset.vellumPath)
					.filter((path): path is string => Boolean(path)),
			),
		),
	}));

	return {
		document,
		pages,
		sourceMap: new DomSourceMap(bodyContainer),
		overlay: new DomOverlayLayer(bodyContainer, pages),
	};
}
