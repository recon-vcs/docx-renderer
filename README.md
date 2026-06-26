# docx-renderer

**The highest-fidelity DOCX-to-HTML renderer for the browser.**

`docx-renderer` converts Microsoft Word documents into paginated HTML entirely in the browser. It parses Office Open XML directly, reconstructs page geometry, and renders into a DOM container — no server, no conversion, no compromise on accuracy.

**Goal: 99% visual agreement with Microsoft Word.**

The renderer already produces high-quality output across a wide range of real-world documents. Development is focused on systematically closing the remaining differences in pagination, typography, spacing, tables, drawings, headers, footers, and other Word layout behaviors.

`docx-renderer` is the document preview engine used by [Recon](https://github.com/recon-vcs) and is published as a standalone npm library.

**[Live Playground →](https://recon-vcs.github.io/docx-renderer/)**

---

## Why not just convert DOCX to HTML?

Rendering a Word document correctly is not the same as extracting its text.

A `.docx` file is a layout system: page geometry, style inheritance, section properties, numbering engines, floating objects, table measurements, font metrics, line breaking rules, and dozens of interacting compatibility settings. Even a small error in any one of these can shift content between lines or pages and cause the entire document to diverge from Word.

`docx-renderer` is designed around **visual fidelity**, not approximate conversion. Every layout decision traces back to the Office Open XML specification and observed Microsoft Word behavior.

## Features

- **Paginated HTML output** — pages match Word's dimensions, margins, and layout
- **No server round-trip** — parsing and rendering run entirely in the browser
- **Headers, footers, footnotes, endnotes** — including first-page and odd/even variants
- **Tables** — including nested tables, merged cells, column spans, and fixed headers
- **Images and drawings** — inline and anchored, with rotation and clipping
- **Text wrapping** — inline, square, tight, through, behind-text, and in-front-of-text
- **Numbering and lists** — inherited through style chains
- **Section breaks and columns** — continuous, page, and odd/even section transitions
- **Style inheritance** — full paragraph and character style chains
- **Document grid** — `lines`, `linesAndChars`, `snapToChars`, and exact/at-least line rule overrides
- **Source-to-DOM mapping** — bidirectional link between document structure and rendered elements
- **Overlay system** — positioned annotations, comments, and editor interfaces
- **Revision marks** — tracked changes rendering
- **ESM, CommonJS, and UMD builds**
- **Deterministic output** — suitable for golden-file regression testing

## Project status

> **0.x — API may change**

The renderer is usable and produces high-quality output on real-world documents. The public API shape may still change while the architecture and compatibility surface are being refined.

Pin an exact package version when depending on the current API shape.

## Installation

```bash
npm install docx-renderer
```

Runtime dependencies (`jszip`, `konva`, `lodash-es`) are installed automatically.

## Quick start

```ts
import { renderSync } from "docx-renderer";

const response = await fetch("/sample.docx");
const blob = await response.blob();

const result = await renderSync(
  blob,
  document.getElementById("document-container"),
  document.getElementById("style-container"),
  { breakPages: true },
);

// Release observers and overlay resources when done.
result.dispose();
```

## API

### `parseAsync`

```ts
parseAsync(
  data: Blob | ArrayBuffer | Uint8Array,
  options?: Partial<Options>,
): Promise<WordDocument>
```

Parses a DOCX file into a `WordDocument` model without rendering. Use this when you need to inspect, cache, or transform the document model before rendering.

### `renderSync`

```ts
renderSync(
  data: Blob | ArrayBuffer | Uint8Array,
  bodyContainer: HTMLElement,
  styleContainer?: HTMLElement,
  options?: Partial<Options>,
): Promise<RenderResult>
```

Parses and renders a DOCX file into the supplied DOM container in one step.

### `renderDocument`

```ts
renderDocument(
  document: WordDocument,
  bodyContainer: HTMLElement,
  styleContainer?: HTMLElement,
  options?: Partial<Options>,
): Promise<RenderResult>
```

Renders an already-parsed `WordDocument`. Use this when you have called `parseAsync` separately.

### `defaultOptions` / `resolveOptions`

```ts
import { defaultOptions, resolveOptions } from "docx-renderer";

const options = resolveOptions({ breakPages: true });
```

## Options

| Option | Type | Default | Description |
|---|---|---|---|
| `breakPages` | `boolean` | `true` | Render each Word page as a separate `<section>` element |
| `className` | `string` | `"docx"` | CSS class prefix applied to rendered elements |
| `ignoreFonts` | `boolean` | `false` | Skip font embedding from the document |
| `ignoreHeight` | `boolean` | `false` | Ignore page height constraints |
| `ignoreImageWrap` | `boolean` | `false` | Ignore image text-wrapping settings |
| `ignoreLastRenderedPageBreak` | `boolean` | `true` | Ignore `<w:lastRenderedPageBreak>` hints from Word |
| `ignoreTableWrap` | `boolean` | `true` | Ignore table text-wrapping |
| `ignoreWidth` | `boolean` | `false` | Ignore page width constraints |
| `inWrapper` | `boolean` | `true` | Wrap output in a container element |
| `renderChanges` | `boolean` | `false` | Render tracked changes (revision marks) |
| `renderEndnotes` | `boolean` | `true` | Render endnotes |
| `renderFooters` | `boolean` | `true` | Render footers |
| `renderFootnotes` | `boolean` | `true` | Render footnotes |
| `renderHeaders` | `boolean` | `true` | Render headers |
| `trimXmlDeclaration` | `boolean` | `true` | Strip XML declarations from inline SVG |
| `useBase64URL` | `boolean` | `false` | Encode images as base64 data URLs instead of blob URLs |
| `debug` | `boolean` | `false` | Emit debug attributes on rendered elements |

## RenderResult

```ts
interface RenderResult {
  document:  WordDocument;   // Parsed document model
  pages:     PageHandle[];   // Per-page metadata and DOM references
  sourceMap: SourceMap;      // Bidirectional DOCX path ↔ DOM element mapping
  overlay:   OverlayLayer;   // Positioned overlay system for annotations
  dispose(): void;           // Disconnect observers and release resources
}
```

### Pages

`pages` is an array of `PageHandle` objects:

```ts
interface PageHandle {
  index:      number;        // Zero-based page index
  element:    HTMLElement;   // The rendered <section> element
  blockPaths: string[];      // Document block paths rendered on this page
}
```

Use `pages` to implement page navigation, virtualized viewers, or document inspection tools.

### Source map

`sourceMap` connects the parsed DOCX structure to the rendered DOM.

```ts
interface SourceMap {
  elementsFor(blockPath: string): HTMLElement[];
  pathFor(element: HTMLElement): string | null;
}
```

Use the source map to locate DOM elements from document paths, attach review comments, synchronize with an external editor, or implement selection and diff tools.

### Overlay layer

`overlay` provides a positioned layer for user interface elements anchored to document content.

```ts
interface OverlayLayer {
  attach(anchor: HTMLElement, content: HTMLElement, opts?: AttachOptions): OverlayHandle;
  clear(): void;
  dispose(): void;
}
```

The overlay layer repositions attached elements automatically when the document is resized or scrolled. It is used by Recon for review comments and annotation markers.

## Separate parse and render

```ts
import { parseAsync, renderDocument } from "docx-renderer";

const documentModel = await parseAsync(blob);

// Inspect, transform, or cache documentModel here.

const result = await renderDocument(
  documentModel,
  document.getElementById("document-container"),
  undefined,
  { breakPages: true },
);
```

## UMD / browser global

Load dependencies before the UMD bundle. The library is exposed as `window.docx`.

```html
<script src="https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/lodash@4.17.21/lodash.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/konva@9.3.6/konva.min.js"></script>
<script src="./docx-renderer.umd.js"></script>

<script>
  const result = await docx.renderSync(blob, container, undefined, { breakPages: true });
</script>
```

## Rendering fidelity

The long-term target is complete visual agreement with Microsoft Word.

Accurate rendering requires understanding how Word interprets Office Open XML, not just how the spec is written. Some layout decisions depend on Word's internal rendering engine, compatibility mode settings, or undocumented behavior. These are treated as measurable engineering problems rather than unknowns.

Areas of active development:

- Page dimensions and margins
- Section breaks and section-scoped layout
- Paragraph spacing and indentation
- Line height, line breaking, and line grid snapping
- Font resolution and fallback behavior
- Style inheritance through paragraph and character chains
- Numbering and list layout
- Table width measurement and cell sizing
- Borders, shading, and table cell backgrounds
- Headers and footers, including first-page and odd/even variants
- Footnotes and endnotes
- Images, drawings, and anchored objects
- Text wrapping modes
- Page breaks and pagination
- Word compatibility settings

Rendering changes should be validated against real DOCX fixtures and visual regression tests. The test suite in `tests/` includes over 100 fixture documents covering the above areas.

## Development

```bash
pnpm install
pnpm build          # ESM + CJS + UMD + TypeScript declarations
pnpm test           # Unit tests
pnpm test:browser   # Playwright rendering tests (requires Chromium)
pnpm dev            # Vite playground with HMR
```

Install Chromium for browser tests:

```bash
npx playwright install chromium
```

## Testing

```
tests/
├── browser/     Playwright rendering tests against real DOCX fixtures
├── fixtures/    100+ DOCX input documents covering layout edge cases
├── golden/      Expected HTML output for regression testing
└── unit/        Isolated parser and renderer unit tests
```

- Unit tests cover pagination logic, style resolution, section parsing, and layout primitives.
- Browser tests render each fixture in Chromium and compare against golden HTML.
- Rendering fixes should include a representative DOCX fixture and a regression test.

## Origin and attribution

`docx-renderer` is a fork of [docx-preview-sync](https://github.com/millet0328/docx-preview-sync) by **millet0328**, which is itself derived from [docx-preview / docxjs](https://github.com/VolodymyrBaydalka/docxjs) by **Volodymyr Baydalka**.

Both upstream projects are licensed under the Apache License 2.0. This project retains that license and the required attribution.

See [`NOTICE`](./NOTICE) for details.

## Contributing

Contributions that improve rendering correctness, OOXML compatibility, or test coverage are welcome.

For rendering changes, include:

1. A DOCX file that reproduces the issue
2. A description of the expected Microsoft Word output
3. A regression test against the new fixture
4. The smallest practical implementation change

Correctness and OOXML compatibility take priority over document-specific CSS patches.

## License

[Apache License 2.0](./LICENSE)

This project includes code derived from `docx-preview-sync` and `docx-preview`. See [`NOTICE`](./NOTICE) for attribution.
