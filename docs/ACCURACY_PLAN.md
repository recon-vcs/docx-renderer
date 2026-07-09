# Accuracy 改善計画 — p1〜p4 100% 目標

現行の文字ズレ対策は [TEXT_ALIGNMENT_PLAN.md](./TEXT_ALIGNMENT_PLAN.md) を優先する。以下は過去ラウンドの調査履歴。

開始時 (`ACCURACY.md`, pixelmatch threshold 0.15):
p1 99.4 / p2 92.5 / p3 99.8 / p4 98.9 / p5 97.0

Phase 1-3 実施後: p1 99.4 / p2 92.6 / p3 99.8 / p4 98.9 / p5 97.0（回帰なし、5ページ維持）

100% 到達可能性: 閾値 0.15 でフォント AA 差は吸収される。位置完全一致なら diff≈0（p3 99.8 が証左）。ただし調査の結果、残差の主因は docGrid ではなく **本環境に Meiryo/Yu Gothic/Yu Mincho が未インストールで、フォールバックフォントの自然行高が Word 実測より系統的に小さいこと**と判明（詳細は Phase 2 参照）。これはコードの一般ロジックでは解消できない環境制約であり、100% 到達の主な障害として残る。図形(VML/DrawingML)関連の残差はユーザー指示により別PRへ先送り。

## 原因（diff 画像 + コード診断済）

### 1. PAGE field 凍結 — 全ページ「1」表示 [p3,p4,p5]

ref: p3=「3」p4=「4」。rendered: 全部「1」。
根本原因: `src/rendering/dom/elements/inline-renderer.ts:22`
`elem.children = ctx.resolveFieldRuns(elem.children)` が段落 children を破壊的置換。
header/footer part の rootElement は全ページ共有 → 初回レンダ時のページ番号「1」がリテラル text run として tree に焼き付き、以降のページで再評価されない。
a.docx に pgNumType なし → 連番が正解。

修正: field 解決結果を tree に書き戻さない。レンダ毎に一時配列で解決（または header/footer part のみ clone）。

### 2. docGrid 行送り未適用 — 縦ドリフト主因 [p2,p4、p1/p3 微小]

ref 行ピッチ 38px（= linePitch 360twips = 18pt × viewport scale 1.5615）、rendered 29px（自然行高）。
p4 で累積 122px、p2 で画像以降 +12px 等のドリフト。
根本原因: `src/ooxml/wordprocessingml/document/model/spacing-between-lines.ts` `parseLineSpacing` は `w:spacing` を持つ段落しか line-height を設定しない。a.docx の docDefaults は `w:spacing w:before="160" w:after="80"` のみ（line なし）→ 通常段落は grid snap されず自然行高。

修正: `docGrid.type = lines/linesAndChars` かつ段落 `snapToGrid ≠ false` なら、spacing.line 不在でも line-height を grid に snap。
- 自然行高 ≤ pitch → line-height = pitch
- 超過 → ceil(自然行高 / pitch) × pitch
- docDefaults の before/after margin と grid の相互作用は Word 実測（ref のピッチ 38px は margin 込みで 1 grid 行 = margin が grid に吸収されている）→ snap 対象段落は before/after も grid 整合させる必要あり。実装時に ref と突き合わせて確定。

最大の系統的修正。p5 への回帰リスクあり → 測定必須。

### 3. p4 2段組: 第2カラムへ分配されない [p4]

ref: 左カラム 余白狭い〜リンク、右カラム 背景オレンジ/太文字/フォント違い/2段組み。
rendered: ほぼ全部左カラム縦一列。column break は docx に存在しない → 自然フロー。
根本原因: region article（`page-renderer.ts` `createPageContent`）は `columnCount` + `columnFill:auto` を設定するが高さ未指定 → CSS multicol が無限高の単一カラムとして流す。

修正: region article に contentHeight（renderPage で算出済）を設定し、Word 同様「左カラムをページ底まで埋めてから右へ」を CSS に再現させる。

### 4. p2 アンカー/回り込み [p2 残差]

原因2修正後に再測定してから着手（位置が全部ズレるため）。現時点の観測:
- 冒頭 code 画像が +12px 下方
- 四角形埋め込み（square wrap）横の行（2〜6）の行送りが grid 非適用
- anchored table（緑）横のテキスト「あ」「d f g h」の x/y が ref と別物 → wrap flow 差
- 写真前面（in-front-of-text）アンカー周辺のテキスト位置・改行差（「→写真前面」→「面」のみ視認）
- リスト番号後 gap（number→text の tab stop）が狭い

### 5. 小物 [p1,p3]

- p1: TOC leader 行・見出しの数 px 縦オフセット → 大半は原因2で解消見込み。残れば TOC スタイル spacing を個別調整
- p3: footnote 行位置・セクション区切り位置 → 原因2 + 原因1 で解消見込み

## 工程

各 Phase 後 `pnpm measure:all` で全5ページ回帰確認。

1. **Phase 1**: PAGE field 修正（原因1）— 完了。`inline-renderer.ts` の `elem.children = resolveFieldRuns(...)` 破壊的代入を廃止、ローカル変数 + `ctx.renderElements` に変更。p3/p4/p5 のページ番号が正しく描画されることを目視確認済み（スコア表示は小数点丸めで同値だが diff 画像のページ番号ブロブは解消）。回帰なし。

2. **Phase 2**: docGrid snap — 完了、ただし当初仮説は誤りと判明。`a.docx` の styles.xml を直接確認したところ、既定段落スタイル "Normal"(styleId `a`) は `<w:snapToGrid w:val="0"/>` を明示しており、本文の大半は本来 grid snap の対象外（`ad` = TOC Heading のみ true に戻す）。一方 `elem.props.snapToGrid` は段落自身の pPr にしか値が入らず、スタイル継承が一切行われていなかった（`inline-renderer.ts` は `elem.props` を直接 `parseLineSpacing` に渡すのみ）。当初計画通り「spacing.line 不在なら常に pitch へ snap」を実装したところ、ほぼ全段落に line-height が加算されて **p2→p3 の改行位置がずれてページ数が 5→6 に増加する回帰**が発生（`測定必須`のとおり検出・停止）。
   - 根本修正: `RenderContext.resolveSnapToGrid(styleName, ownSnapToGrid)` を新設（`html-renderer-sync.ts`）。直接指定 > スタイルの `basedOn` チェーン > ドキュメント既定段落スタイル > true（OOXML既定）の順で解決し、`inline-renderer.ts` はこれを `parseLineSpacing` に渡すよう変更。
   - grid snap 自体（`spacing-between-lines.ts`）は元の設計通り実装して残したが、正しく解決された `snapToGrid` によりほぼ全段落でスキップされるため、実質的な影響は僅小（p2 92.5%→92.6%）。
   - 実測した ref のピッチ約38px（p4 Letter セクションの "H"/"Bbb" 制御段落群、Latin cap-top基準で複数段落間隔から算出）は docGrid 由来ではなく、Meiryo フォント（本環境に未インストール、IPA Gothic 等へフォールバック）の実際の自然行高が、フォールバックフォントの行高より系統的に大きいことに起因すると判明。これはフォントメトリクスの環境差であり、コード側の一般ロジックでは解消不可（残差として報告）。
   - 副次的に発見した独立バグを合わせて修正: 段落の `margin-top`/`margin-bottom`（`w:spacing before/after`）はブラウザの既定 margin collapse により隣接段落間で `max()` に縮退していたが、Word は before/after を常に加算する。`default-styles.ts` の `.${c} p` に `padding-top: 0.05px` を追加し、視覚上のオフセットなしに collapse を止めた（一般的な CSS 手法、全段落に一律適用）。

3. **Phase 3**: 2段組カラム高さ — 部分的に完了、当初仮説の一部も誤りと判明。region article には実際には既に `contentHeight` が設定されていた（`html-renderer-sync.ts` の単一セクション分岐 `renderPage`）。ブラウザ DOM 調査で確認済み。真因は `columnFill: auto`（最初の列をページ全高まで埋めてから次列へ）がそのまま機能していたことで、Word の実際の分割点より段落が「短すぎる」ため列1に入りすぎていた（Phase 2 で判明したフォントメトリクス差が根本原因）。`columnFill: balance` も試したが reference は明らかに不均等分割（列1が列2よりずっと長い）であり balance は悪化したため `auto` に確定（コメントに実測根拠を明記）。列高さ自体の割当ロジックに変更は不要だった。

4. **Phase 4 / 5**: p2/p1/p3 の残差を diff 画像で再確認。コード画像・テキストボックス・箇条書き番号・TOCリーダー線・footnote 位置はいずれも「二重像」パターン（赤+黄が数px〜十数pxずれて重なる）で、すべて Phase 2 で特定したフォントメトリクス環境差の帰結。p2 のアンカーテーブル（緑色セル表）・p5 の表は見た目上ボーダー/シェーディング/セル配置は reference と一致しており、diff が真っ赤に見えるのは高密度な境界線パターンに数px の累積ズレが乗っているため（表そのものの実装バグではない）。四角形埋め込み（図形）と写真前面まわりの wrap 差、および p5 の禁止マーク図形はユーザー指示によりスコープ外（別PR）。

## ラウンド2 (2026-07-09、Windowsフォント fontconfig 導入後)

fontconfig で /mnt/c/Windows/Fonts を参照（メイリオ実体で測定可能に）→ baseline 98.4%。
そこから縦ドリフトの真因を分解して修正。結果: **overall 98.4→98.9%（p1 99.4 / p2 98.7 / p3 99.8 / p4 99.4 / p5 97.1）**、5ページ維持。

確定した真因と修正（すべて一般ロジック、magic 定数なし）:

1. **段落 strut が UA 既定 16px**: p 要素に run 既定フォント（docDefaults/スタイル rPr）が適用されず、行箱が 16px メイリオ strut(24px) で決まっていた。スタイルの span 向け宣言から font-family/font-size を p にもミラー（document-styles.ts）。
2. **margin collapse で after 消失**: before→`padding-top` / after→`margin-bottom` に変更（spacing-between-lines.ts）。padding は隣接 margin と衝突せず加算、末尾 margin はページ底で自然に破棄され Word の「ページ末 after-spacing 破棄」と一致。
3. **GDI 行高丸め**: Word は ascent/descent を px に**別々に ceil**して合算（メイリオ 10.5pt: ceil(14.84)+ceil(6.16)=22px）、ブラウザは合計丸めで 21px → 1px/行の系統ドリフト。canvas で実フォントの ascent/descent 比を実測し `line-height` を GDI 式で設定（gdi-line-height.ts、フォントデータテーブル不要）。`line-height: normal` の箇所のみ適用、exact/atLeast は不変。
4. **ページ先頭の before-spacing**: Word は自然改ページで到達した先頭段落の space-before を破棄（section 先頭は保持）。`article[data-page-start] > p:first-child { padding-top: 0 }`。
5. **測定時の幻影 margin**: overflow 測定が article を scroll container 化し最終子の margin-bottom が scrollHeight に算入 → Word なら収まる行が 1px 溢れ 6ページ化。`article > :last-child { margin-bottom: 0 }`（Word のページ末 after 破棄と同義、:last-child が append 順に自動追従）。
6. **wrapSquare bothSides の左右誤り**: wrapText 未指定/bothSides を無条件 float:left（テキスト右側）にしていたが Word は広い側へ流す。render 時に containing block の実測ギャップ比較で float 側を決定（drawing-renderer.ts）。p2「ｄｆｇｈ」の2行折返し + 6px 溢れが解消。
7. **画像のみの行の descent**: Word は inline 画像だけの行を画像高さぴったりにする（font descent を足さない）。テキストなし段落の drawing を `vertical-align: text-bottom` に（inline-renderer.ts）。p2 の一律 +8px ズレ解消。
8. contextualSpacing（同一スタイル連続時の before/after 抑制）を実装: `p.X + p.X { padding-top: 0 }` + `:has(+ p.X) { margin-bottom: 0 }`（Title / List Paragraph が使用）。
9. フォント読込完了待ち（font-loading.ts）: ページ分割測定前に document.fonts を待機。

前ラウンドの「フォントメトリクス環境差でロジック解消不可」という結論は**誤り**だった（1〜3 が真因）。

残差（今後）:
- p2 1.3%: 四角形埋め込み float 周辺の行送り・リスト番号 tab gap・テキストボックス枠線
- p1 0.6%: TOC リーダー線・見出し微小オフセット
- p4 0.6%: 2段組の分割点微差
- p5 2.9%: 禁止マーク図形（スコープ外・別PR）・数式位置
- 図形本体（VML/DrawingML geometry）は別PR

## ラウンド3 (2026-07-09、「文字ずれ」対応)

ユーザー指摘「図・写真は合うが文字がずれる」の真因は位置ではなく**フォント解決**だった:

1. **theme CSS 変数の未定義**: renderTheme は `--docx-{major,minor}HAnsi-font` しか設定せず、スタイル CSS が参照する `--docx-*EastAsia-font` / `--docx-*Bidi-font` が未定義 → `var()` 未定義で font-family 宣言全体が invalid → 見出し・表題が UA フォールバック（system-ui、太く見える）で描画されていた。さらに theme の `<a:ea typeface=""/>`（空）は「`<a:font script="Jpan">` を themeFontLang で引け」の意味なのに未実装。修正: settings.xml の `themeFontLang` をパース（settings.ts）、theme の script 別 typeface を収集（theme.ts）、6変数すべてを言語→script 解決付きで定義（document-styles.ts）。表題/見出しが 游ゴシック Light で正しく描画。
2. **TOC 内 Hyperlink スタイル**: TOC entry run は rStyle=Hyperlink（色+下線）を持つが、Word は TOC field 結果内でそれを表示しない（本物のリンクは表示する）。fldChar begin/end を文書順に追跡して TOC instruction 内の run から Hyperlink 参照を除去（toc-hyperlink.ts、複数段落跨ぎ対応）。TOC が黒・下線なしに。

テキストボックスは position/文字とも ±1px（ラウンド2 の修正で解消済みだったことを実測確認）。
p1 99.4（内訳 99.39→99.41）。残る TOC 差はリーダー線の描画方式（border-dotted vs '.' グリフ）と 内容→見出し1 の gap 約10px。
