# docx-vellum 移行計画: legacy廃止 + experimental既定化 + 段組み/数式/図形修正

## 背景

レンダラー2系統並存:

- `HtmlRenderer` (`src/html-renderer.ts`, 1711行) — `renderAsync()`使用。コード内コメントで明示"legacy"。
- `HtmlRendererSync` (`src/html-renderer-sync.ts`, 3166行) — `renderSync()`使用。ページネーション対応本線。

legacy機能劣化確認済み(数式justification一時欠落等)。本線側にも独立バグ3件発見。再現fixture`recon-cli/docx-xml-sample-analyze.docx`をplayground+`renderSync`で表示→ユーザー添付スクショと同じ崩れ再現済み:

1. **2段組み(`w:cols w:num="2"`)で表/数式/図形が極小に圧縮**
   `createPageContent()`(`html-renderer-sync.ts:1243-1258`) `column-count`をcontent要素に直適用。`column-fill`未指定+高さ`auto`→Chrome balance(均等分配)レイアウト発生。Word段組みは「上から詰めて溢れたら次列」逐次充填モデル、balanceとは根本異質。表/数式/floatする図形が予期しない位置/サイズに圧縮される主因。

2. **数式(MathML)レンダリングロジック自体は実装済み**(`renderMmlNary`/`renderMmlFraction`/`renderMmlRadical`等、2805-2920行付近)。要因1解消で正常組版見込み。

3. **図形("禁止"マーク等)完全非表示。バグ2系統重複:**
   - VML側: `parseVmlElement`の`case "shape"`、`<g>`生成のみ。`path`/`adj`/`formulas`(VML図形アウトライン定義)変換一切無し(`src/vml/vml.ts:35-37`)。`convertPath()`関数(同120-127行)定義済みだが**呼び出し元無し**(未使用コード)。
   - DrawingML側: `parseShape()`(`document-parser.ts:1619-1645`)、`cNvPr`/`spPr`等同caseにfall-through、最初の子要素で`parseShapeProperties()`呼び即`return`。`parseShapeProperties()`(1648-1696行)、`prstGeom`(図形種類)含む全case`default`にfall-through→**常に`null`返却**。レンダラー側(`html-renderer-sync.ts`)に`DomType.Shape`のcase**存在せず**。新旧両図形経路、パーサー・レンダラー双方機能不全確定。
   - `checkAlternateContent()`(`document-parser.ts:1116-1133`)、`supportedNamespaceURIs`空配列のため`mc:Choice`(モダンDrawingML)常時無視、`mc:Fallback`(VML)採用固定。

4. **`options.experimental`フラグ**(既定`false`)抑制対象:
   - タブストップ/リーダー位置再計算(`renderTab`/`refreshTabStops`、TOCドットリーダー・右タブズレ原因)
   - コメント注釈レンダリング(`renderCommentRangeStart/End/Reference`、2693/2702/2711行)
   常時有効化、フラグ自体削除(死分岐残さず)。

## 方針

**docx-vellum単体のみ修正**。`recon/vendor/docx-vellum`反映はユーザー側で別途実施、`recon`コード側触らない。

5フェーズ進行。フェーズ毎に独立ビルド/テストgreen維持。

---

### Phase 1 — legacy `HtmlRenderer`完全削除

- `src/html-renderer.ts` 削除
- `src/docx-preview.ts`: `HtmlRenderer`import削除、`renderDocument()`の`sync`パラメータ削除(`HtmlRendererSync`常時使用)、`renderAsync()`エクスポート削除
- `tests/browser/render.spec.ts`: `'renderAsync golden output'`スイート(12ケース)削除。`renderSync`側カバレッジ無いケースは移植
- `rollup.config.mjs`/型エクスポート: `renderAsync`参照削除、再ビルド
- `playground/main.ts`: `renderSync`使用へ変更済み(対応完了)

### Phase 2 — `experimental`フラグ削除(常時有効化)

- `src/docx-preview.ts`の`Options`型・`defaultOptions`から`experimental`削除
- `html-renderer-sync.ts`内`if (this.options.experimental)`/`if (!this.options.experimental)`分岐(2531, 2693, 2702, 2711, 2992-2998行)全撤去、常時実行コードへ
- 影響: コメント注釈常時出力(意図通り)。タブストップ再計算は同期1回実行のためパフォーマンス影響無し見込み

### Phase 3 — 段組み(multi-column)レイアウト修正

- `createPage`/`createPageContent`(`html-renderer-sync.ts:1211-1258`)修正:
  - content要素に`column-fill: auto`追加
  - 高さ`minHeight`→確定値(`contentSize.height`)に変更、balanceモデル回避、Word的「上から詰め次列」挙動へ
  - 既存ページ分割(overflow検出)との整合確認、列が複数ページ跨ぐケース(`break-page-section-break.docx`等)回帰確認
- 再現用`tests/fixtures/zz-sample-analyze.docx`を正式fixture名(例:`shapes-and-math-mixed.docx`)へリネーム、golden test追加

### Phase 4 — 図形(Shape)レンダリング実装

優先: モダンDrawingML(`wsp`/`prstGeom`)経路を本線化。VML任意ベジエformula言語(`path`/`adj`/`formulas`)対象外、VML側既存`rect`/`oval`/`line`/`imagedata`のみ維持。

- `document-parser.ts`:
  - `parseShape()`: fall-through不具合修正、`cNvPr`/`spPr`/`txbx`等個別処理、全要素走査後に結果返却
  - `parseShapeProperties()`: `prstGeom`(`prst`属性=図形名)/`custGeom`/`solidFill`/`noFill`/`ln`(枠線)実読取。`prstGeom`は属性保持のみ、SVGパス変換は新規マッピングテーブルで実施(rect/roundRect/ellipse/triangle/diamond/各種arrow/`noSmokingSymbol`等頻出20〜30種から開始、未知presetは矩形近似)
  - `checkAlternateContent()`: `supportedNamespaceURIs`に`wps`名前空間URI追加、Choice(モダン)優先採用
- `html-renderer-sync.ts`: `DomType.Shape`レンダリングcase新設(`renderShape()`)。`prstGeom`名→SVG`<path>`変換テーブル参照、`spPr`の塗り/線/変形(`xfrm`回転・反転は既存`parseTransform2D`結果再利用)適用。テキスト含む図形(`txbx`)は既存`foreignObject`パターン(VML textbox実装参考)で対応
- VML側`convertPath()`(現未使用): 明示`path`属性持つ単純図形(矢印・吹き出し等頻出typeのみ)限定で接続検討。formula言語フル実装はしない

### Phase 5 — 検証

- `npm run build`/`tsc --noEmit`コンパイルエラー無し確認
- 既存golden test(`tests/browser/render.spec.ts`, `tests/unit/*`)実行、Phase1削除分(`renderAsync`系)以外全green
- playground目視確認(agent-browserでスクショ):
  - `zz-sample-analyze.docx`(2段組み+数式+図形): 表圧縮無し、数式が本文サイズで組版、"禁止"マークが矩形+斜線アイコンで表示
  - `table-of-contents.docx`, `columns.docx`, `column-break.docx`: 既存段組み/目次系fixture回帰無し
  - `comment.docx`: コメント注釈常時表示確認
