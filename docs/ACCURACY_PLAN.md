# Accuracy 改善計画 — p1〜p4 100% 目標

現状 (`ACCURACY.md`, pixelmatch threshold 0.15):
p1 99.4 / p2 92.5 / p3 99.8 / p4 98.9 / p5 97.0

100% 到達可能性: 閾値 0.15 でフォント AA 差は吸収される。位置完全一致なら diff≈0（p3 99.8 が証左）。残差は全て layout 差。

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

1. **Phase 1**: PAGE field 修正（原因1）— 独立・低リスク・p3/p4/p5 即効
2. **Phase 2**: docGrid snap 一般化（原因2）— 最大インパクト。p2/p4 大幅改善見込み、p5 回帰監視
3. **Phase 3**: 2段組カラム高さ（原因3）
4. **Phase 4**: p2 wrap 個別修正（原因4）— Phase 2 後の新 diff 基準で
5. **Phase 5**: 残差潰し（原因5）— p1〜p4 100% まで反復
