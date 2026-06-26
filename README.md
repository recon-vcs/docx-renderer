# docx-vellum

Synchronous docx → HTML rendering library for the browser.

docx-vellum parses Office Open XML (`.docx`) files with no server round-trip and
renders them as paginated HTML directly into a DOM container. It is the document
preview engine of the [Recon](https://github.com/recon-vcs) project and is being
evolved into a standalone, publishable npm library.

> **Status: 0.x — the public API is being redesigned.** The current
> `renderSync` / `renderAsync` entry points are inherited from upstream and will
> change in an upcoming release (structured `RenderResult` with page handles,
> source maps and an overlay layer). Pin an exact version if you depend on the
> current shape.

## Origin and credits

docx-vellum is a fork of
[docx-preview-sync](https://github.com/millet0328/docx-preview-sync) by
**millet0328**, which is itself derived from
[docx-preview / docxjs](https://github.com/VolodymyrBaydalka/docxjs) by
**Volodymyr Baydalka**. Both projects are licensed under the Apache License 2.0,
and this fork keeps that license. See [NOTICE](./NOTICE) for attribution details.

## Installation

```bash
npm install docx-vellum
```

Runtime dependencies (`jszip`, `konva`, `lodash-es`) are declared as regular
dependencies and are installed automatically.

## Usage

### ESM / bundler

```ts
import { renderSync } from "docx-vellum";

const response = await fetch("/sample.docx");
const blob = await response.blob();

const body = document.getElementById("document-container");
const style = document.getElementById("style-container");

// Parses the document and renders paginated HTML into `body`.
const result = await renderSync(blob, body, style, { breakPages: true });
// Call result.overlay.dispose() when done to release observers.
```

### Browser global (UMD)

Load `dist/docx-vellum.umd.js` after `jszip`, `konva` and `lodash`; the library
is then available as the global `docx`:

```html
<script src="https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/lodash@4.17.21/lodash.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/konva@9.3.6/konva.min.js"></script>
<script src="docx-vellum.umd.js"></script>
<script>
  docx.renderSync(blob, document.getElementById("container"));
</script>
```

### API

| Function | Description |
| --- | --- |
| `parseAsync(data, options?)` | Parses a docx file (`Blob`, `ArrayBuffer` or `Uint8Array`) into a `WordDocument` model without rendering. |
| `renderSync(data, body, style?, options?)` | Parses and renders with the synchronous, pagination-aware renderer. Returns a `RenderResult`. |
| `renderDocument(doc, body, style?, options?)` | Renders an already-parsed `WordDocument`. Returns a `RenderResult`. |
| `defaultOptions` | The default `Options` values. |

See the `Options` interface in `src/docx-preview.ts` for all rendering flags
(page breaking, header/footer/footnote rendering, experimental tab stops, ...).

## Demo

`docs/index.html` (renderSync) and `docs/renderAsync.html` (renderAsync) are
static demo pages. Build first, then serve the repository root with any static
file server and open the pages:

```bash
npm run build
npx serve .   # or: python3 -m http.server
# open http://localhost:3000/docs/
```

## Development

```bash
npm install
npm run build          # ESM + CJS + UMD bundles and .d.ts into dist/
npm test               # vitest unit tests (tests/unit)
npx playwright install chromium   # once, before the first browser test run
npm run test:browser   # Playwright rendering tests (tests/browser)
```

Test fixtures live in `tests/fixtures/*.docx`; golden HTML outputs for the
rendering regression tests live in `tests/golden/*.html`.

## License

[Apache-2.0](./LICENSE). Includes code from docx-preview-sync and
docx-preview/docxjs — see [NOTICE](./NOTICE).
