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

## ラウンド5（2026-07-09、実バグ3件特定・修正だが視覚的ズレは未解決）

### 見つけて直したバグ

DOM実機測定（`getBoundingClientRect`/`getComputedStyle`/canvas font metrics probe）で機序を特定。3件とも live DOM mutation で効果を先に確認してからコード反映。

1. **mixed-font 段落の line-height 握り潰し**（`gdi-line-height.ts` / `inline-renderer.ts`）
   `renderParagraph` が子 run 描画前に段落自身の line-height を確定させ、子 span 側の `applyGdiLineHeight` は「継承した非 normal 値」を見て「もう明示済み」と誤判定しスキップしていた。段落の strut フォントと実際の run フォントが異なる場合（例: 見出し1本文＝段落 strut は游ゴシック Light だが実テキスト run はメイリオ）、run 側の高さ補正が効かず段落全体が低くなる。
   修正: run 呼び出し側に `force` オプション追加、継承由来の非 normal は無視して自前フォントで再計算。run は OOXML 上そもそも自分の line-height を持たない（段落プロパティのみ）ので、常に上書きしてよい。

2. **docGrid line-height スナップが段落固有 props だけ見ていた**（`inline-renderer.ts`）
   `renderParagraph` の `parseLineSpacing` 呼び出しが `elem.props`（そのパラグラフ instance が直接持つ pPr）だけを渡していて、style 側の `spacing`（before/after/line）をマージしていなかった。ほとんどの段落は spacing を style 側から継承するため、この呼び出しは「明示 spacing 無し」と誤判定し、`snapToGrid=true` の段落（TOC Heading 等）を doc grid 1 行分（linePitch そのまま）に強制スナップ、CSS クラス側の正しい line-height（style 由来）を inline style で上書きしていた。
   修正: `elem.props.spacing ?? style?.paragraphProps?.spacing` でマージしてから渡す。

3. **TOC tab leader の vertical-align 欠落**（`javascript.ts` `applyTabLeader`）
   `display:inline-block` + `overflow:hidden` だけで `vertical-align` を明示していなかった。CSS2.1 10.8.1 により、overflow が visible でない inline-block の baseline 基準は「テキストの baseline」でなく「box の下端」になる。結果、leader の点々が本来の位置より浮いて見え、かつ line box がその分膨張していた。
   修正: `vertical-align: bottom` を明示。

### 検証結果と重要な気づき

- `pnpm test` 103 pass、回帰なし。
- `pnpm measure:all`: p1 は 99.531%（バグ2件修正後）→ tab leader 修正適用後 99.486%（数値上は悪化）。
- **だが `diff.png` を目視すると、3件とも「個々には正しいバグ修正」なのに、ページ全体のズレは消えていない。ズレの発生箇所が入れ替わっただけ。** 修正前は 内容/TOC見出し1/↑目次/見出し1本文/本文/改行のみ が一致、見出し2・3(TOC)/見出し2・3(本文) がズレていた。修正後は 内容/TOC見出し1/見出し1〜3本文 が新たにズレ、↑目次 は逆に一致するようになった。
  → 一部の段落は「別の未解決バグと偶然相殺して、たまたま合っていた」だけだった。片方を直すと相殺が崩れて別の場所にズレが出る。トータルのズレ量はスコア的にもほぼ変わっていない（フラット〜微減）。
- tab leader 修正はスコアを下げるが、点々が baseline に乗る＝視覚的には明確に正しいので維持を選択（対症療法ではなく実際のCSSバグ）。数値より見た目を優先する判断はユーザー確認済み。

### 未解決: 表題→内容の gap 不足（-20px 実測）

- `表題`(style a3, 游ゴシック Light) → `内容`(TOC Heading) の段落間 gap が、reference 比で約20px 不足。
- フォント一致は確認済み（段落 strut・実 run とも游ゴシック Light、mixed-font ではない）。GDI line-height 計算（`gdi-line-height.ts`）も他の行と同じロジックで一貫しており、単体のバグ箇所が見当たらない。
- `表題` の `margin-bottom` を実験的に +20px すると、この1箇所だけ reference に一致し p1 スコアも 99.533% まで上がる。ただし「なぜ20pxか」の根拠がなく、コードへの反映（`Title` style への決め打ち）は保留。

### 結論

残るズレの本丸は、Word内部の line-height 算出（GDI/DirectWrite ベースと推定、未文書化）と、こちらの `line-height: normal`（ブラウザの font-metrics 依存）の間の乖離そのもの。個別段落へのパッチはモグラ叩きになりやすい（今回がまさにそれ）。
次に進むなら、段落単体（表題・見出し1/2/3・TOC・本文）ごとに「reference が要求する実際の line-height」を先に一覧化し、フォント／サイズとの相関パターンを見てから直すべき。パターンが見えないまま個別修正を重ねるのは前回・今回と同じ失敗を繰り返すだけ。
