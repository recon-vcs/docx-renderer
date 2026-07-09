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

1. **文字位置計測を固定**
   - `pnpm measure:all` で現状画像を再生成。
   - `agent-browser` で p1 を開き、表題、TOC 見出し、leader、ページ番号の DOM rect と computed style を取る。
   - reference PNG 側は `pngjs` で同じ文字塊の bbox を取る。以後は「どの文字塊が何 px ずれたか」で判断する。

2. **run `w:spacing` を直す**
   - `src/ooxml/wordprocessingml/parsing/properties-parser.ts` の `parseSpacing` を `letter-spacing` へ変更。
   - `margin-bottom` への変換を消す。run spacing は縦余白ではない。
   - paragraph style の run ruleset から生成された span/p strut にも、必要な範囲で同じ文字幅が効くことを確認。
   - 期待効果: 表題、見出し、p4 の「大きい文字」など、文字幅由来の x/y 連鎖ズレが減る。

3. **TOC tab leader を Word 型にする**
   - `<w:tab>` は underline ではなく leader span にする。
   - tab stop までの残り幅を測り、現在 font/size/letter-spacing で dot glyph を繰り返す。
   - 最後のページ番号の右端を `w:tab w:pos="8494"` に合わせる。
   - font load 後に再計算する。既存の `refreshTabStops()` に統合する。
   - 期待効果: p1 TOC の長い横線差分を消す。

4. **CJK 互換設定を実装範囲に入れる**
   - settings parser に `characterSpacingControl`、`useFELayout`、`balanceSingleByteDoubleByteWidth` を保持させる。
   - まず `compressPunctuation` だけ実測で効く箇所を確認する。全テキストへ固定 letter-spacing を足さない。
   - `leftChars` / `hangingChars` / `firstLineChars` は、char grid がある時だけ文字単位インデントとして評価する。`left` と二重適用しない。

5. **docGrid は対象を限定**
   - Normal の `snapToGrid=0` を尊重する。
   - TOC Heading と明示 `snapToGrid` 段落だけ line pitch を確認する。
   - p4 の2段組みは、文字幅と行高が直った後に column split を再測定する。先に column fill をいじらない。

## 検証

- 必須: `pnpm test`、`pnpm measure:all`。
- 目標: p1 の表題/TOC 二重像を消す。p1〜p4 はページ数維持で 99.8% 以上を先に狙う。100% は diff を見て残差を個別に潰す。
- 回帰条件: 5ページが6ページになる、p2 の画像/表位置が崩れる、p4 の2段組み分配が悪化する。この場合は止めて原因を測る。

根本原因候補: run spacing を縦余白として扱う誤変換、TOC leader の描画モデル違い、CJK 互換設定未反映。
