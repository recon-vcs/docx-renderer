# docx-vellum 描画精度 再設計計画書

対象: `tests/fixtures/zz-sample-analyze.docx` 解析結果ベース。
目的: 対症療法でなく、拡張性高く描画精度保証できる設計への再構築。

## 0. 調査サンプル構造(zz-sample-analyze.docx)

- `word/document.xml` 70KB。3つの `w:sectPr` 存在(複数セクション文書)。
- セクション1: headerReference/footerReference 各3種(even/default/first)持つ。用紙11906x16838(A4)、単カラム。
- セクション2: header/footerReference 無し(直前セクション継承前提)。用紙12240x15840(Letter)、2カラム。
- セクション3: header/footerReference 無し。用紙11906x16838、2カラム。
- `word/header1-3.xml` `word/footer1-3.xml` 各2.7-3.3KB。
- 数式: `m:oMathPara`/`m:oMath` 2ブロック(べき乗・上付き下付き・分数・Σ記号付き総和)。
- 図形: `mc:AlternateContent` 内に `wps:wsp`(Choice, DrawingML)+`w:pict`(Fallback, VML)。テキストボックス(`wps:txbx`/`w:txbxContent`)複数。
- 最終図形の正体: `a:prstGeom prst="noSmoking"`(Word「禁止」マーク=リング+斜めバー)。`adj="1936"`(調整値)付き。VML側 `v:shapetype` も対応する手書き数式定義(`v:formulas`)持つ。
- 画像3枚(`word/media/image1-3.png`)、`w:drawing`としてインライン挿入。`w:lastRenderedPageBreak` 複数箇所に存在(Word保存時点の改ページ記録)。

## 1. アーキテクチャ現状

パイプライン: `document-parser.ts`(XML→IR) → `html-renderer-sync.ts`(IR→DOM、ページネーション込み)。
DrawingML/VMLは別実装(`document-parser.ts`内のDrawingML解析、`src/vml/vml.ts`がVML解析)。図形描画は`renderShape`(SVG生成)とKonva(画像変換のみ、回転/クリップ)。

## 2. 問題ごと根本原因

### 2-1. 数式(OMML)崩れ・無駄な改行

- パース: `document-parser.ts` mmlTagMap(L27-52)で oMath系13要素対応。`parseMathProperties`(L1060-1101)は属性6種(chr/vertJc/jc/pos/degHide/begChr/endChr)のみ対応。色・フォント・サイズ等のOOXML数式属性は無視。
- レンダリング: `html-renderer-sync.ts` L1838-2005, 2053-2967でMathML(`<math>`/`<mfrac>`/`<msup>`/`<msub>`等)へ直接変換。KaTeX/MathJax等の組版エンジン未使用、ブラウザ標準MathML実装に委譲。
- 根本原因: ブラウザ間MathML実装差異 + 属性未対応により Word本来のフォントサイズ・配置比率が再現されない。「無駄な改行」はoMathParaのブロック要素マージン処理、もしくはMathML `<mrow>` の折り返し制御(white-space/display値)未設定が濃厚。要再現テストでの特定。

### 2-2. 図形が弱い・複雑図形が単純楕円化

- `src/shapes/preset-geometry.ts`(52行): 対応プリセット17種類(rect/roundRect/ellipse/三角系/矢印系/line)。各パスはECMA-376公式の調整可能数式(`a:avLst`の`gd`)を計算せず、固定比率の近似パスをハードコード。`noSmoking`のみ複合パス(2サブパス)対応の特例。
- 未知プリセットはrectへ無条件フォールバック(コメントL1-5に明記、設計として承知の上の妥当ライン)。
- `src/vml/vml.ts`(127行): v:rect/v:oval/v:shape/v:textboxの素朴な変換のみ。custGeomパスデータの幾何パース無し、グループ図形(wpg:wgp)非対応、回転/変換属性パース欠落、グラデーション無視。
- スタイル解決: `document-parser.ts` L1695-1749は`solidFill`直接指定のみ対応。`wps:style`内の`a:fillRef`/`a:lnRef`/`a:effectRef`(テーマカラースキーム参照、本サンプルのnoSmoking図形が実際に使っている方式)は一切解決されない。`html-renderer-sync.ts` L2448-2452のコメントでも明記の既知制約。
- `adj`調整値(本サンプルでは`adj="1936"`、斜めバーの位置決定パラメータ)は完全無視、固定ハードコード比率使用。
- 根本原因: (1)プリセット形状が公式調整可能数式を計算していない、(2)テーマスタイル参照(fillRef/lnRef)が未実装、(3)DrawingMLとVMLが別実装で対応範囲・精度がそもそも噛み合っていない。これら3点が重なり「Wordの禁止マーク」が「単純な塗り楕円」へ大幅劣化。

### 2-3. テキストボックスサイズ異常

- EMU→px/pt変換自体(`convertLength`, `LengthUsage.Emu = 1/12700`)は正しい。
- `wps:txbx`は親Shapeのextent(width/height)を継承し`width:100%, height:100%`の相対配置(`html-renderer-sync.ts` `renderShape` L2476-2486)。親Shapeの`a:xfrm`/`a:ext`が正しく読めていれば連動するはずだが、回転を伴うShape(`parseTransform2D`)では再計算済みの幅高さとテキストボックスの相対配置が不整合を起こしやすい。
- 根本原因: テキストボックスが「自身の絶対サイズ」を持たず常に親依存。回転・複合変形時のズレが温床。

### 2-4. セクション区切り未認識・改ページ崩れ

- `src/document/section.ts`: `w:sectPr`の`pgSz`/`pgMar`/`cols`/`headerReference`/`footerReference`パースは実装済み(L109-266)。`contentSize.height`は計算されておらず「ヘッダーフッターDOM確定後に再計算」とコメントのみ(L261-263)。
- `html-renderer-sync.ts` L1036-1047: ヘッダー/フッター未指定セクションへの継承ロジックは実装済み(`_.unionBy`で直前セクションのrefsとマージ)。継承自体は対症療法ではなく構造的に対応されている。
- カラム数切替(`w:cols num=2`)はL1247で対応コードあるが、本サンプルのような「セクション単位で単カラム→2カラムへ切替わる」ケースの実機検証記録が見当たらない(golden testでの複数セクション+カラム切替の網羅性不明)。
- 改ページ: `splitPageBySymbol`(L670-)は`w:lastRenderedPageBreak`(Word保存時点の静的記録)を基準に粗い分割。テーブル/TOC含むページのみ`isSplit=false`にして実測オーバーフロー再検出(`Overflow` enum、L1180-1191)の対象にする設計。通常の段落+画像のみのページは静的マーカー位置をそのまま信用。
- 根本原因: 画像が新規挿入/差し替えされた場合、Word保存時点の`lastRenderedPageBreak`位置と実際の高さが食い違う。動的overflow検出機構自体は存在するが、適用対象がテーブル/TOCに限定されており、画像主体ページでは機能しない設計上の穴。

### 2-5. ヘッダー/フッター

- `renderHeaderFooterRef`(L1271-)で first/even/default解決、高さ実測(`getOffsetHeight`)してpadding-top/bottomへ反映(L1148-1174)するロジックは実装済みで設計として妥当。
- `src/header-footer/parts.ts`(36行)はパート読み込みのみの薄い層、ロジックはhtml-renderer-sync側に集約。
- 機能自体は実装済み。優先度は「動かない」ではなく「3セクション×3種別の組合せを回帰テストで保証していない」点。

## 3. 設計方針

対症療法(個別パッチの積み上げ)でなく、以下3軸で構造的に再設計する。

- **責務分離**: 「XMLパース」「OOXML中間表現(IR)構築」「レイアウト計算」「DOM/SVG描画」を明確に分ける。現状DrawingMLとVMLが別実装で同じ概念(図形・パス・塗り)を別々に持っているのが拡張性を阻害する根。
- **公式仕様準拠の段階的強化**: 近似パスのハードコードから、ECMA-376準拠の計算へ移行(優先: 図形ジオメトリ、数式属性)。
- **計測ファースト**: レイアウト崩れは「静的記録を信用する」から「実測して検証・補正する」へ。lastRenderedPageBreakは初期ヒントとして使い、最終判定は常に実測オーバーフローに委ねる。

## 4. 再設計案

### 4-1. 図形ジオメトリ統一エンジン(優先度最高、視覚インパクト最大)

- 新設: `ShapeGeometryIR { kind, basePaths: PathCommand[], adjustables: Record<string, GuideFormula> }` という中間表現を定義。
- ECMA-376 Part 1 §20.1.9.x の `prstGeom` 調整可能数式(`gd`/`avLst`)を解釈する計算エンジンを実装(主要十数種から段階導入、noSmoking含む)。
- `a:fillRef`/`a:lnRef`/`a:effectRef` → `theme.ts`の配色・線スキームへの解決関数を追加。`solidFill`直接指定と統合した「FillResolver」に一本化。
- DrawingML(`wps:wsp`)とVML(`v:shape`)を同じIRへ変換するアダプタを書き、レンダラ(SVG生成)を1本に統合。VML専用のcustGeom/グループ/回転対応もこのIR層で吸収する。
- 効果: noSmoking等の複合図形が正しい形状・正しいテーマ色で再現される。新規プリセット追加が「パス文字列のハードコード追加」でなく「公式数式の登録」になり拡張性が上がる。

### 4-2. レイアウト計測の二段化(ページネーション/改ページ信頼性)

- 全要素(段落・画像・図形・テーブル)を`isSplit=false`相当の動的オーバーフロー検出対象に統一する。`lastRenderedPageBreak`は「初期分割ヒント」のレイヤーに格下げし、レンダリング後に実測offsetHeightで再検証・補正するフックを追加。
- 画像はCSS固定サイズ(EMU変換済み)のため計測自体は安定済み。問題は「検出ロジックの適用範囲」であり画像変換コード自体の修正は不要。
- 効果: 画像追加・差し替え時にWord保存時点の記録とズレてもページ境界が崩れない。

### 4-3. 数式属性カバレッジ拡大+改行原因の特定修正

- `parseMathProperties`の対応属性をOOXML仕様の主要プロパティ(色/フォント/サイズ/太字italic/上下添字スケール比)まで拡張。
- MathML出力側のCSS強化(`display: inline-flex`相当の折り返し制御、`mfrac`/`munderover`の行間調整)。「無駄な改行」は再現テストケースを切り出し、`oMathPara`のブロックmargin処理かMathML要素のwhite-space設定かを確定後に修正。

### 4-4. テキストボックス独立サイズ対応

- `wps:txbx`に親Shape依存のwidth:100%だけでなく、`a:xfrm`があれば自身のEMUサイズを優先する分岐を追加。回転変形時は`parseTransform2D`計算後の実寸を反映。

### 4-5. セクション/ヘッダーフッター回帰保証

- 機能自体は実装済みのため新規実装は最小限。`tests/golden`に「3セクション・カラム数混在・header even/default/first混在」のケースを追加し、退行検出を仕組み化する。

## 5. 実装フェーズ

1. **Phase 0 計測基盤**: golden image diffテストにサンプル(zz-sample-analyze.docx)由来のケースを追加(図形/数式/セクション境界/画像混在ページ)。先にテストを赤くしてから着手。
2. **Phase 1 図形ジオメトリ統一エンジン**(4-1): 最大の見た目インパクト。DrawingML/VML共通IR化、FillResolver実装、主要プリセットの公式数式対応。
3. **Phase 2 レイアウト二段化**(4-2): 動的オーバーフロー検出を全要素に拡張、lastRenderedPageBreakをヒント層へ格下げ。
4. **Phase 3 数式属性拡張+改行修正**(4-3): 再現ケース確定後にCSS/属性対応。
5. **Phase 4 テキストボックス/セクション仕上げ**(4-4, 4-5): 回帰テスト整備、残課題クローズ。

各フェーズ完了条件: 該当golden testが通り、`zz-sample-analyze.docx`のレンダリング結果がWord出力(比較画像)と視覚的に一致すること。
