# PR #1 Review Feedback

Target: https://github.com/recon-vcs/docx-vellum/pull/1

This file records the broad review findings and the follow-up fixes applied for
them. It intentionally avoids small style comments.

## Resolution State

- `pnpm test`: passed, 18 files / 102 tests
- `pnpm run build`: passed
- `pnpm run test:browser`: passed, 25 tests

## Fixed Items

1. Browser rendering gate was red.
   - Fixed textbox alignment so the sample verification passes.
   - Browser tests now dispose the returned `RenderResult`.

2. Local workflow artifacts were mixed into the PR.
   - Removed `CLAUDE.md`.
   - Removed `real/zz.png`.

3. Public API docs were inconsistent.
   - README now describes `RenderResult` as the current return type.
   - Removed stale `renderAsync` and `docs/renderAsync.html` references.
   - Added `tests/unit/public-api-docs.spec.ts` to keep this from regressing.

4. `RenderResult` lifecycle was not enforced.
   - Added `RenderResult.dispose()`.
   - New renders dispose any previous render result attached to the same body container.
   - Overlay cleanup now removes overlay DOM nodes as well as observers/listeners.

5. OOXML model and rendering state were coupled.
   - Moved `Page` from `src/ooxml/.../document/model/page.ts` to `src/rendering/pagination/model/page.ts`.
   - Renderer pages now live in `RenderSession` instead of mutating `document.documentPart.body.pages`.
   - Removed render-time DOM fields from the parsed document model.

6. Pagination depended on DOM rendering helpers.
   - Moved `splitElementsByBreakIndex` from DOM core to pagination model.
   - Moved `Overflow` out of DOM utilities into measurement.
   - Measurement no longer imports DOM mutation helpers.

7. Parser boundaries were too broad.
   - Split parser context types by domain: body, paragraph, run, math, table, style, numbering, and drawing.
   - Sub-parsers now depend on narrow context interfaces instead of the full `ParseContext`.

8. Architecture gates were too weak.
   - Added `tests/unit/architecture-boundaries.spec.ts`.
   - The new test blocks OOXML -> rendering imports, pagination -> DOM rendering imports, measurement -> DOM helper imports, and broad `ParseContext` imports in sub-parsers.

## Remaining Risk

The renderer still has a central orchestration class and inherited non-strict OOXML model shapes. Those were improved where they directly caused the review findings, but a full renderer pipeline rewrite would be a larger follow-up than this fix pass.
