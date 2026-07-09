# 文字ズレ改善計画

目的: `tests/fixtures/a.docx` の p1〜p4 を Word PDF に寄せる。図形本体より文字位置を先に直す。スコアだけ見ない。p1 の表題、TOC、p4 の2段組み文字を目視基準にする。

## 現状

- `ACCURACY.md`: p1 99.4 / p2 98.7 / p3 99.8 / p4 99.4 / p5 97.1。
- p1 diff は表題と TOC が赤黄の二重像。数字は高いが、人間にはズレが目立つ。
- `a.docx` は `w:characterSpacingControl="compressPunctuation"`、`w:useFELayout`、`w:balanceSingleByteDoubleByteWidth`、`w:docGrid type="lines" linePitch="360"`。
- Normal は `snapToGrid=0`。全段落を grid に押し込むのは誤り。TOC Heading だけ grid 対象。
- 表題 style は `w:spacing w:val="-10"`。現実装は run spacing を `margin-bottom` にしている。これは誤り。文字間隔として扱う。
- TOC leader は現実装が dotted underline。Word は tab leader の dot glyph を tab stop まで敷く。underline 代用では dot の高さ・間隔・終端が合わない。

## 修正順

1. **文字位置計測を固定** — 実施済み
   - `pnpm measure:all` で現状画像を再生成。
   - `agent-browser` で p1 を開き、表題、TOC 見出し、leader、ページ番号の DOM rect と computed style を取る。
   - reference PNG 側は `pngjs` で同じ文字塊の bbox を取る。以後は「どの文字塊が何 px ずれたか」で判断する。

2. **run `w:spacing` を直す** — 実施済み
   - `src/ooxml/wordprocessingml/parsing/properties-parser.ts` の direct run `parseSpacing` を `letter-spacing` へ変更。
   - style/default `rPr` の `w:spacing` も `letter-spacing` へ変換。表題 style の `w:spacing="-10"` はここを通る。
   - `margin-bottom` への変換を消した。run spacing は縦余白ではない。

3. **TOC tab leader を Word 型にする** — 実施済み
   - `<w:tab>` は underline ではなく leader glyph span にする。
   - tab stop までの残り幅を測り、現在 font/size/letter-spacing で glyph を繰り返す。
   - font load 後の `refreshTabStops()` 再計算に統合。

4. **CJK 互換設定を実装範囲に入れる** — 未実施
   - settings parser に `characterSpacingControl`、`useFELayout`、`balanceSingleByteDoubleByteWidth` を保持させる。
   - まず `compressPunctuation` だけ実測で効く箇所を確認する。全テキストへ固定 letter-spacing を足さない。
   - `leftChars` / `hangingChars` / `firstLineChars` は、char grid がある時だけ文字単位インデントとして評価する。`left` と二重適用しない。

5. **docGrid は対象を限定** — 継続
   - Normal の `snapToGrid=0` を尊重する。
   - TOC Heading と明示 `snapToGrid` 段落だけ line pitch を確認する。
   - p4 の2段組みは、文字幅と行高が直った後に column split を再測定する。先に column fill をいじらない。

## 検証

- 必須: `pnpm test`、`pnpm measure:all`。
- 目標: p1 の表題/TOC 二重像を消す。p1〜p4 はページ数維持で 99.8% 以上を先に狙う。100% は diff を見て残差を個別に潰す。
- 回帰条件: 5ページが6ページになる、p2 の画像/表位置が崩れる、p4 の2段組み分配が悪化する。この場合は止めて原因を測る。

## 今回の結果

- `pnpm test`: 103 tests passed。
- `pnpm run build`: 成功。
- `pnpm measure:all`: 5183 使用中で失敗。代替として `agent-browser` + `scripts/measure-accuracy.mjs` で docx-renderer のみ採点。
- 結果: p1 99.53% / p2 98.70% / p3 99.79% / p4 99.39% / p5 97.10%、5ページ維持。
- `ACCURACY.md` の丸め表示は p1 99.5%。`accuracy/docx-renderer` の画像は最新採点で更新。

根本原因候補: run spacing を縦余白として扱う誤変換、TOC leader の描画モデル違い、CJK 互換設定未反映。

## 測定環境修正

- 2026-07-09: 通常の agent-browser は `/mnt/c/Windows/Fonts` を fontconfig 検索対象にしておらず、`Meiryo` / `Yu Gothic` / `游ゴシック` が fallback と同じ metrics になっていた。
- `scripts/measure-all.mjs` が project-local `fonts.conf` を生成し、`/mnt/c/Windows/Fonts` がある時だけ追加するよう修正。agent-browser は専用 profile で起動し直す。
- `pnpm measure:all`: docx-renderer 98.901%、p1 99.531 / p2 98.702 / p3 99.787 / p4 99.387 / p5 97.101。既存 accuracy 画像と一致。

結論: 行高を増やす修正は p2 を壊すので却下。まず測定ブラウザのフォント解決を固定した。残る p1 の上寄りは Windows フォント使用後の本物の layout 差として扱う。
