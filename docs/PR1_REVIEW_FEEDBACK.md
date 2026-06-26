# PR #1 Review Feedback

Target: https://github.com/recon-vcs/docx-vellum/pull/1

Reviewed state:

- Branch: `big-refactor`
- Head: `31b27dde73a254caf09d3a4a8d315c601737368b`
- Scope: `main...HEAD`, 130 files, +6965 / -6035
- `pnpm test`: passed, 16 files / 97 tests
- `pnpm run build`: passed
- `pnpm run test:browser`: failed, 24 passed / 1 failed

## 1. Browser-level quality gate is red

The PR currently fails `pnpm run test:browser`. The failing test is the large sample verification in `tests/browser/_verify-sample.spec.ts`, and the failure is inside the textbox rendering expectations: the test expects every textbox to have `alignItems === 'flex-start'`, but the rendered sample returns `normal`.

This is not a small isolated unit-test mismatch. The PR adds a complicated visual/rendering refactor and also adds sample rendering verification, but that verification does not pass. Before discussing architecture polish, the browser rendering gate should be green or the expected rendering contract should be deliberately changed.

Also, the current sample verification is still mostly DOM/CSS probes. `CLAUDE.md` and `real/zz.png` imply a visual comparison workflow, but that is not represented as a real test. If the visual golden is intentional, make it a proper test fixture and assertion path. If it is only local workflow scratch, it should not be in the PR.

Relevant files:

- `tests/browser/_verify-sample.spec.ts`
- `CLAUDE.md`
- `real/zz.png`

## 2. The refactor split files, but rendering and layout are still coupled

`src/html-renderer-sync.ts` is under the new line-count limit, but it still owns too much of the system: page creation, header/footer measurement, overflow decisions, region splitting, source path mutation, note rendering, field resolution, table state, tab measurement, and renderer callback wiring.

The most concerning part is that DOM rendering is still deciding pagination while mutating both the document model and page list. `renderElements` renders a node, reads `dataset.overflow`, mutates `elem.index` / `elem.breakIndex`, calls `splitOnOverflow` or `splitRegionOnOverflow`, replaces `session.currentPage`, and recursively renders the next page. That keeps layout policy, DOM measurement, and rendering in one control loop.

This feels more like a mechanical extraction from the old class than a systematic architecture. A cleaner direction would make layout/split decisions explicit data, then let DOM rendering consume the layout result. If incremental rendering with measurement is unavoidable, the mutable session should still be an explicit layout/render session object with narrower responsibilities, not a central renderer object that every feature module reaches back into through callbacks.

Relevant files:

- `src/html-renderer-sync.ts`
- `src/rendering/pagination/model/page-split.ts`
- `src/rendering/measurement/overflow-measurer.ts`

## 3. Parser boundaries are still broad

The parser decomposition is an improvement over one giant file, but the new boundary is still a wide `ParseContext` passed into almost every sub-parser. That context exposes paragraph, run, table, math, drawing, VML, default properties, alternate content, and body parsing all at once.

That makes dependencies easy to grab and hard to reason about. Adding a new parser feature can become "thread one more method through the global context" instead of defining a cohesive parser for one OOXML domain. It also weakens unit tests because a supposedly focused parser can depend on much more of the parsing pipeline than its API suggests.

I would narrow this before the structure settles: split context by domain, pass only the dependencies a parser actually needs, and keep recursive parsing entry points explicit. The goal should be that parser modules reveal their true dependency surface from their function signatures.

Relevant files:

- `src/ooxml/wordprocessingml/parsing/document-parser.ts`
- `src/ooxml/wordprocessingml/parsing/parse-context.ts`

## 4. Public API state is inconsistent

The code now exports `parseAsync`, `renderDocument`, `renderSync`, and `RenderResult` types from `src/docx-preview.ts`. README also says `renderSync` and `renderDocument` return `RenderResult`.

But the README status block still says `renderSync` / `renderAsync` are inherited upstream entry points and that structured `RenderResult` is an upcoming change. Later, the demo section still references `docs/renderAsync.html`, which is not present in `docs/`, and `renderAsync` is not exported from `src/docx-preview.ts`.

This makes the public contract unclear: is `RenderResult` the current API, or a future API? Is `renderAsync` intentionally removed, or accidentally missing? For a PR this large, the public surface needs one clear story across exports, README, demos, and tests.

Relevant files:

- `README.md`
- `src/docx-preview.ts`
- `src/render.ts`

## 5. RenderResult introduces lifecycle responsibility without enforcing it

`RenderResult` creates an overlay layer with `ResizeObserver`, `MutationObserver`, and a `window.resize` listener. Cleanup exists through `result.overlay.dispose()`, but this is only documented in one ESM snippet. The UMD example and browser tests discard the return value, so they do not exercise or enforce the lifecycle contract.

This is an API design issue, not just a docs issue. Repeated rendering into the same container can leave callers responsible for knowing about hidden observers. Either the renderer should own cleanup for resources it creates, or the returned result should be the explicit owner of all disposables and examples/tests should consistently dispose it.

Relevant files:

- `src/render-result.ts`
- `README.md`
- `tests/browser/render.spec.ts`
- `tests/browser/_verify-sample.spec.ts`

## 6. The current quality gates do not prove the new architecture

The new line-count test only checks that production `.ts` files stay under 800 lines. That is a weak proxy for maintainability. It passes even when the renderer still has central mutable session state and broad callback surfaces.

The unit tests are useful for characterization, and they now pass, but many of them lock in implementation details of the new pagination helpers rather than enforcing architectural boundaries. If the purpose of this PR is a systematic refactor, add gates that protect the intended boundaries: renderer modules should not own parser concerns, layout modules should not depend on DOM rendering, parser modules should not receive a global parsing surface by default, and browser rendering samples should be green.

Relevant files:

- `tests/unit/line-count.spec.ts`
- `src/html-renderer-sync.ts`
- `src/ooxml/wordprocessingml/parsing/parse-context.ts`

## 7. Strict TypeScript is still mostly aspirational

`tsconfig.json` enables `strict`, but disables `noImplicitAny`, `strictNullChecks`, and `strictPropertyInitialization`. That may be necessary for inherited code, but this PR creates many new module boundaries where those checks would be most valuable.

At minimum, new modules should not normalize the old looseness. If full strictness is too much for this PR, consider a smaller scoped rule: new parser/rendering/pagination modules should avoid `any`, explicit nullable state should be modeled, and compatibility with old non-strict code should be isolated at the adapter boundary.

Relevant file:

- `tsconfig.json`

## 8. Layer dependencies still run in the wrong direction

The new directory layout suggests separate OOXML model, pagination, measurement, and DOM rendering layers, but some dependencies still cross those boundaries in the wrong direction.

`src/ooxml/wordprocessingml/document/model/page.ts` imports pagination types and stores `contentElement?: HTMLElement`, so the document model knows about rendering-time DOM state. `src/rendering/pagination/model/page-split.ts` imports `splitElementsByBreakIndex` from `src/rendering/dom/core/split-by-break.ts`, so pagination model logic depends on a DOM rendering helper. `src/rendering/measurement/overflow-measurer.ts` also imports append helpers and the `Overflow` enum from DOM utilities, which keeps measurement coupled to DOM mutation.

This is exactly the kind of coupling the refactor appears to be trying to remove. The directory names are cleaner, but the dependency graph still says "model/layout/rendering are mutually aware." A more durable version would keep parsed OOXML model free of DOM and render-time page state, keep pagination split logic independent of DOM helper modules, and expose measurement as a narrow service returning metrics or decisions.

Relevant files:

- `src/ooxml/wordprocessingml/document/model/page.ts`
- `src/rendering/pagination/model/page-split.ts`
- `src/rendering/dom/core/split-by-break.ts`
- `src/rendering/measurement/overflow-measurer.ts`
