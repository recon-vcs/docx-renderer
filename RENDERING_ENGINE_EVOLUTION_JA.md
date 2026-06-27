# docx-renderer は従来の docxjs / docx-preview-sync と何が違うのか

## 概要

`docx-renderer` は、docxjs の「Office Open XML を読み、ブラウザの HTML/CSS として表示する」という基礎と、millet0 の `docx-preview-sync` が導入した「実際の DOM の高さを測りながらページを分割する」という発想を引き継いでいる。

一方で、現在の実装は従来コードへの個別パッチの集積ではない。高精度化のために、レンダリング処理を次の独立した問題へ分解し直している。

1. DOCX パッケージと OOXML の解釈
2. Word の論理構造からレイアウト領域への変換
3. セクション、改ページ、段組みから物理ページを計画する処理
4. HTML/CSS への変換
5. ブラウザが確定した寸法を使うオーバーフロー測定
6. 測定結果を文書構造へ反映する再分割

この分離により、Word の意味を失わずにブラウザのレイアウトエンジンを利用できるようになった。精度向上の本質は、CSS の調整量ではなく、**Word の論理モデルとブラウザの物理レイアウトの間に、明示的な変換層を置いたこと**にある。

## 系譜

### docxjs

docxjs は、DOCX をブラウザだけで解析し、段落、文字装飾、表、画像、番号、ヘッダー、フッターなどを HTML/CSS へ変換する基盤を作った。サーバー側で PDF や画像へ変換せず、DOM として扱える点は現在も重要な設計資産である。

ただし、基本的な思想は「文書要素を対応する DOM に変換する」ことにあり、Word と同じページ境界を再現するためのレイアウトモデルは限定的だった。DOCX 内に保存された `lastRenderedPageBreak` のようなヒントへ依存する場面もあり、編集後の文書、異なるフォント環境、複雑な表や図形ではページ位置がずれやすい。

### millet0 の docx-preview-sync

`docx-preview-sync` は、DOM を描画しながら `clientHeight` と `scrollHeight` を比較し、はみ出した要素を次ページへ送る仕組みを追加した。`breakIndex` をツリー上に記録し、段落や表の内部まで再帰的に分割することで、docxjs よりも実際のブラウザ表示に即したページネーションを実現した。

これは重要な前進だった一方、ページ、OOXML 要素、DOM 測定状態が同じオブジェクトグラフに混在していた。改ページのたびに深いコピーとツリーの書き換えが発生し、セクション、段組み、ヘッダー・フッター、配置画像を組み合わせるほど状態管理が難しくなっていた。

### 現在の docx-renderer

現在の実装は、従来の DOM 実測方式を捨てず、その前後に明示的なモデルと責務境界を追加した。これにより「仕様から決定できる配置」と「ブラウザで測らなければ決定できない配置」を分けて扱っている。

## 従来との違い

| 観点 | docxjs | docx-preview-sync | docx-renderer |
|---|---|---|---|
| 主目的 | DOCX の HTML 表示 | DOM 実測によるページ分割 | Word に近い物理ページの再構成 |
| ページ境界 | 保存済みヒントと明示改ページが中心 | 描画中のオーバーフローを再帰分割 | セクション・改ページの事前計画と DOM 実測の併用 |
| セクション | ページに付随する属性として扱いがち | `Page` がセクションと描画状態を保持 | `LayoutRegion` と `PhysicalPage` を分離 |
| 1ページ内の複数セクション | 表現が難しい | 表現が難しい | 1物理ページに複数 region を保持可能 |
| ヘッダー・フッター | ページ配列の位置に依存 | first/even/default の状態が描画処理へ混在 | セクション内ページ番号と継承を専用コンテキストで解決 |
| 配置画像 | CSS の位置指定を中心に処理 | 通常フローと同じ測定に巻き込まれる場合がある | OOXML の wrap 意味に基づいて測定対象外を判定 |
| 実装構造 | 大きな parser / renderer | 大きな同期 renderer と可変ツリー | OOXML、pagination、measurement、DOM rendering を分離 |
| 検証 | 基本機能のサンプル中心 | 個別不具合へのテスト | 小さい純粋関数の単体テストと実 DOCX のブラウザテストを併用 |

## 高精度化を可能にした設計判断

### 1. 「セクション」と「ページ」を別の概念にした

Word のセクションは、ページそのものではない。セクションは、余白、用紙サイズ、段組み、文書グリッド、ヘッダー・フッター参照などが有効になる論理範囲である。

たとえば `continuous` セクション区切りでは、同じ物理ページの途中から段組みや書式が変わる。従来の「1 Page = 1 sectPr」というモデルでは、この文書を正しく表現できない。

現在は次の中間モデルを使う。

```text
Document body
    ↓ section-stream
LayoutRegion[]
    ↓ explicit-breaks
改ページ・改段位置を付与した LayoutRegion[]
    ↓ page-builder
PhysicalPage[]
    ↓ DOM rendering + measurement
最終的な paginated HTML
```

- `LayoutRegion` は、同じセクション設定が適用される連続した内容を表す。
- `PhysicalPage` は、実際に表示する1ページを表し、複数の `LayoutRegion` を持てる。
- `PageLayoutContext` は、物理ページ番号とセクション内ページ番号を別々に保持する。

これにより、continuous / next page / next column / odd page / even page を同じ仕組みで解釈できる。奇数・偶数ページ区切りで必要になる空白ページも、DOM 描画の副作用ではなくページ計画として生成できる。

関連実装:

- `src/rendering/pagination/model/section-stream.ts`
- `src/rendering/pagination/model/layout-region.ts`
- `src/rendering/pagination/core/page-builder.ts`
- `src/rendering/pagination/model/page-numbering.ts`

### 2. 改ページを「事前に分かるもの」と「測定が必要なもの」に分けた

改ページには性質の違う2種類がある。

- `w:br`、セクション区切り、奇数・偶数ページ指定など、OOXML だけで決定できるもの
- フォント、行折り返し、表、画像サイズなど、ブラウザがレイアウトして初めて決定できるもの

前者は `section-stream`、`explicit-breaks`、`page-builder` で先に物理ページへ変換する。後者は DOM 挿入後に測定し、オーバーフロー位置で現在の region を分割して次ページへ送る。

このハイブリッド方式には次の利点がある。

- 明示改ページを偶然の高さ計算へ依存させない。
- すべてを独自組版エンジンで再実装せず、ブラウザの実フォントメトリクスを使える。
- 実測による分割がセクション境界を壊さない。
- ページ分割の判断を、小さい純粋関数として単体テストできる。

`lastRenderedPageBreak` も絶対的な正解とはせず、レイアウト上の hint として保持する。Word が最後に保存した環境と現在のブラウザ環境が異なる可能性を前提にした判断である。

### 3. オーバーフローを状態機械として扱った

「はみ出した / はみ出していない」の二値だけでは、入れ子の段落、run、表、セルを安全に分割できない。現在は `NONE`、`SELF`、`PARTIAL`、`FULL`、`UNCHECKED`、`SKIP` を区別し、子要素の結果から親要素の状態を合成する。

重要なのは状態数そのものではなく、測定結果と処理方針を分離したことである。

- 要素自身がページを越えたのか
- 子の一部だけが越えたのか
- 全体が次ページへ移るべきか
- そもそも通常フロー外なので測定すべきでないか

これらを区別することで、段落内の途中分割、表の行・セル内部の分割、先頭要素がページより大きい場合の無限再分割防止を同じ処理系で扱える。

### 4. 配置画像を通常フローから明確に除外した

Word の前面・背面配置や `wrapNone` の図形は、本文の高さを押し広げる要素ではない。しかし、DOM 上の `scrollHeight` だけを見ると、絶対配置された画像が本文のオーバーフロー原因に見える場合がある。

現在は、DOM の `position: absolute` という結果だけでなく、OOXML の `WrapType.None` という意味も使って out-of-flow 要素を判定する。対象には DOM 挿入前から `data-overflow="skip"` を付け、ページ測定時だけ一時的に非表示にして高さから除外し、測定後は必ず元の表示状態へ戻す。

この工夫により、次のような誤動作を防いでいる。

- 配置画像だけを理由に不要なページが増える。
- 同じ画像と本文ブロックが次ページへ繰り返し送られる。
- ヘッダー画像が本文のページネーションへ干渉する。

### 5. ヘッダー・フッターをセクション局所のページ文脈で解決した

Word のヘッダー・フッターは、文書全体の単純な「1ページ目 / 偶数ページ / 通常ページ」では決まらない。

- `titlePage` の first は、そのセクションの先頭ページに適用される。
- even/default の選択には物理ページの偶奇が必要になる。
- 参照が省略されたセクションでは、前セクションから種類ごとに継承する。

現在は `PageLayoutContext` で、物理ページ番号、セクション ID、セクション内ページ番号、セクション先頭か、偶数ページかを保持する。参照の継承と first/even/default の選択も専用関数へ分離した。

その結果、ページ番号とセクション番号を混同せず、セクションをまたぐヘッダー・フッターを安定して選択できる。

### 6. Word 固有の行組みを CSS の単純変換で済ませない

ページ差の多くは、1行あたり数 px の誤差が積み重なって起きる。そのため、段落間隔と行間を単に `line-height` へ写すだけでは足りない。

現在は次を区別している。

- `auto`: 240分率の Word 行間
- `exact`: 指定値を固定行高として使用
- `atLeast`: 指定値を最小値として使用
- `snapToGrid`: 文書グリッドへ吸着するか
- `docGrid`: `lines`、`linesAndChars`、`snapToChars` などの section 設定

特に `atLeast` と document grid が同時に指定された場合は、固定値として上書きせず、明示最小値と grid pitch の大きい方を採用する。これは Word の「最低値」という意味を CSS 上でも維持するための処理である。

### 7. 図形、画像、テキストボックスを意味単位で解析した

DrawingML を単なる `<img>` へ落とすのではなく、位置、extent、回転、切り抜き、wrap、shape、text body を別々に解析するようにした。

主な改善点は次のとおり。

- 画像の 30 / 90 / 180 / 270 / 360 度回転
- 回転後の bounding box を考慮した寸法
- `srcRect` による画像クリッピングと回転の組み合わせ
- inline / square / tight / through / top-and-bottom / behind / in-front の wrap
- テキストボックスの inset と縦方向 anchor
- `spAutoFit` / `normAutofit` / `noAutofit` の区別
- shape geometry と元の extent の分離

ここでも、OOXML の意味を parser で保持し、renderer はそのモデルを DOM へ変換する。解析中に CSS 文字列だけへ潰さないため、回転、切り抜き、自動サイズ調整を組み合わせても判断材料が残る。

### 8. 数式はブラウザの native MathML を壊さない形で出力した

OMML 数式を MathML へ変換する際、見た目だけを合わせる wrapper を追加すると Chrome の native MathML layout を壊すことがある。実際に `math { display: inline-block }` は内部の `mrow` を不正に積み上げる原因になった。

現在は、数式 run を内容に応じて `<mi>`、`<mn>`、`<mo>` へ分解し、MathML 本来のレイアウトへ委ねる。数式段落では不要な `column-span` も使わず、段組み内で数式が別領域へ飛ぶことを避けている。

これは「CSS で強制的に似せる」より「ブラウザが理解できる正しい意味構造を渡す」方を選んだ例である。

### 9. 巨大な parser / renderer を領域別に分割した

従来は、数千行の `document-parser.ts` と `html-renderer-sync.ts` が多くの仕様と可変状態を抱えていた。現在は主に次の境界へ分割している。

```text
src/opc/                       DOCX パッケージ、Part、Relationship
src/ooxml/                     WordprocessingML / DrawingML / OMML のモデルと解析
src/rendering/pagination/      section、break、page の計画と再分割
src/rendering/measurement/     DOM のオーバーフロー測定
src/rendering/dom/             styles と要素別 DOM renderer
src/render-result.ts           page handle、source map、overlay、lifecycle
```

単にファイルを分割しただけではない。`Page` を `OpenXmlElement` から外し、レンダリング時の状態を OOXML モデルへ混入させないようにした。parser も万能コンテキストではなく、paragraph / run / table / style など領域ごとの狭い context を受け取る。

DOM renderer 間の依存は `RenderContext` に集約し、個別 callback interface の乱立を抑えた。`sourcePath` はモデルへ書き込まず `WeakMap` で管理するため、解析済みモデルをレンダリング都合のメタデータで汚しにくい。

### 10. 精度改善を再現可能なテストへ変えた

高精度レンダリングでは、ある文書を直した結果、別の文書のページ数や配置が崩れることが多い。そのため、実装と同じくらい fixture の設計が重要になる。

現在のリポジトリには、次のような境界条件を分離した DOCX fixture がある。

- section break、continuous、next column、odd/even page
- first/even/default header・footer と継承
- 明示改ページ、`lastRenderedPageBreak`、空白ページ
- 単一セル、複数セル、結合セル、複数ページ表、固定ヘッダー行
- document grid、line rule、文字間隔、tab stop
- 画像回転、切り抜き、各種 text wrap
- 数式、脚注、文末脚注、変更履歴、目次

検証は二層に分かれる。

- pagination、番号、継承、overflow 合成などを DOM なしで検証する単体テスト
- 実 DOCX を Chromium で描画し、ページ数、DOM、画像の decode、配置状態を確認するブラウザテスト

不具合を「その文書だけの CSS 補正」で直すのではなく、最小 fixture と純粋関数のテストへ落とすことで、改善を他の文書にも適用できる設計知識へ変えている。

## 実装上の工夫

### ブラウザを組版エンジンとして利用する

Word の全組版規則を JavaScript だけで再実装するのは現実的ではない。特にフォント fallback、glyph 幅、禁則、MathML、表の intrinsic sizing はブラウザがすでに高度な実装を持っている。

そのため、docx-renderer は次の分担を採用している。

- OOXML の意味、セクション境界、改ページ規則は自前で解釈する。
- 文字幅、折り返し、表の実寸、フォントメトリクスはブラウザに計算させる。
- 計算結果を測定し、Word のページ制約に合わせて再分割する。

この分担が、精度と実装可能性のバランスを取っている。

### ページ全体を毎回作り直さず、分割位置を伝播する

オーバーフローが起きた要素から親へ分割位置を伝播し、現在ページと次ページへ必要な subtree を分ける。表セルや段落内のような深い位置でも、最上位ブロックを丸ごと失わずに次ページへ継続できる。

従来の全面的な `cloneDeep` 依存は縮小し、ページ分割は `SplitTarget` の小さな契約へ寄せた。これにより、`Page` を一時的な OOXML 要素として偽装せず、region と page の両方へ同じ分割処理を適用できる。

### レンダリング対象外ではなく「測定対象外」を表現する

配置画像やヘッダーは表示する必要があるが、本文の高さには含めてはいけない。要素を削除するのではなく `SKIP` として測定だけから除外することで、表示とページネーションの責務を分離している。

### API に描画結果のライフサイクルを持たせる

`render()` は DOM を生成するだけでなく、ページ一覧、source map、overlay、`dispose()` を含む `RenderResult` を返す。

- `pages`: 物理ページと DOM の対応
- `sourceMap`: DOCX 内の source path と DOM の双方向対応
- `overlay`: コメントや差分 UI を本文 DOM から分離して配置
- `dispose()`: observer や overlay を確実に解放

これは直接の組版精度だけでなく、Recon のようなレビュー UI で「どの原文要素がどのページのどこに描かれたか」を安定して扱うための設計である。

## 何が精度向上に最も効いたのか

改善効果が大きかった順に整理すると、次のようになる。

1. セクションと物理ページを分離した中間モデル
2. OOXML で決定できる改ページと DOM 実測が必要な改ページの分離
3. out-of-flow drawing を本文の overflow 測定から除外する処理
4. セクション局所のヘッダー・フッターとページ番号コンテキスト
5. document grid、行間、段組みなど小さな誤差が累積する箇所の意味的な解釈
6. DrawingML、OMML、表を大きな renderer から独立させたこと
7. 実 DOCX fixture と小さい pagination test の組み合わせ

個々の CSS 修正も必要だが、それだけでは高精度にはならない。ページ差は複数機能の相互作用で発生するため、**誤差がどの段階で生まれたかを特定できる構造**が最も重要だった。

## 現在地

docx-renderer は Microsoft Word 自体を再実装したものではなく、完全互換を達成したと主張するものでもない。フォントの有無、ブラウザ差、Word の非公開互換動作、未実装の OOXML 要素によって差は残る。

ただし、従来の「DOCX をそれらしく HTML にする」方式からは明確に進んでいる。現在は、Word の論理構造を保った中間モデル、ブラウザによる実測、測定結果に基づく再ページネーション、fixture による検証を一つのパイプラインとして持つ。

そのため、新しい不一致が見つかったときも、個別文書向けの例外を増やすのではなく、次のどこに原因があるかを分けて改善できる。

- OOXML の解釈不足
- layout region への変換不足
- page planning の規則不足
- DOM/CSS 変換の不足
- overflow 測定の誤り
- ブラウザ固有差

この原因分解可能性こそが、現在の高精度と、今後さらに精度を上げられる拡張性の両方を支えている。
