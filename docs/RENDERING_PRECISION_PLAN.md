# docx-renderer 描画精度 再設計計画書

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

### 2-1. 数式(OMML)崩れ・無駄な改行 ── セクション/カラムレイアウト起因が濃厚(フォント・色は無関係)

`document.xml`実測: 問題の数式2ブロック(`m:oMathPara` L1928, L1983)はいずれも**セクション3内**(L1377 sectPrより後、L2353 sectPrより前)に配置されている。セクション3は`<w:cols w:num="2"/>`(2カラムレイアウト)。つまり数式はCSS `column-count: 2`で組まれた狭いカラム幅の中に押し込まれている。フォント/色/属性の問題ではなく、**カラム幅に対して数式ブロックが大きすぎることによる強制折り返し・分断**が第一の疑い。

- カラム実装: `html-renderer-sync.ts` `createPageContent`(L1244-1266)で`columnCount`/`columnGap`/`columnFill:auto`をCSSに設定。コメントL1242に開発者自身の TODO「分栏：一个页面可能存在多个章节section，每个section拥有不同的分栏」(1ページ内に複数セクションが存在し各セクションが異なるカラム数を持つケース未対応)と明記。本サンプルはセクション単位でカラム数が変わる(セクション1: 単カラム / セクション2,3: 2カラム)構成であり、この既知の弱点に直撃する可能性が高い。
- CSS Multicolの既定動作は「ブロック要素はカラム境界で分割されうる」。MathML `<math>`/`<mfrac>`等のコンテナに`break-inside: avoid`相当の指定が無ければ、数式がカラム境界やカラム幅で意図せず分割・圧縮される。これが「無駄な改行」の直接原因として最有力。
- 数式パース・属性カバレッジの不足(`parseMathProperties` L1060-1101が6属性のみ対応、色/フォント/サイズ未対応)も実在する制約だが、見た目崩れの主因ではなく、二次的な精度劣化要因として扱う。

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
- SectionBreak自体の改ページ処理は実装あり(`splitElementsBySymbol`内 L868-894): `NextPage`/`EvenPage`/`OddPage`/defaultで`startNewPage()`を呼ぶ。だが**`Continuous`と`NextColumn`のcaseは空(no-op、L873-880)で未実装**。本サンプルは全sectPrに`w:type`指定が無くデフォルト`NextPage`扱いのため直接は発火しないが、type指定付きの文書(continuous断面やカラム内改段)では区切りが無視されるバグが確定的に存在する。
- カラム数切替(`w:cols num=2`)はL1247で対応コードあるが、`createPageContent`のコメントL1242で「1ページに複数セクションが混在し、各セクションが異なるカラム数を持つケースは未対応」と開発者自身が明記。本サンプルはまさにセクション単位でカラム数が変わる構成(2-1参照)であり、この欠落の影響を受ける本命候補。
- 改ページ: `splitPageBySymbol`(L670-)は`w:lastRenderedPageBreak`(Word保存時点の静的記録)を基準に粗い分割。テーブル/TOC含むページのみ`isSplit=false`にして実測オーバーフロー再検出(`Overflow` enum、L1180-1191)の対象にする設計。通常の段落+画像のみのページは静的マーカー位置をそのまま信用。
- 根本原因: (a)画像が新規挿入/差し替えされた場合、Word保存時点の`lastRenderedPageBreak`位置と実際の高さが食い違う。動的overflow検出機構自体は存在するが、適用対象がテーブル/TOCに限定されており画像主体ページでは機能しない。(b)Continuous/NextColumn未実装、(c)1ページ内複数セクション・複数カラム未対応、の3点が複合してセクション境界の改ページ・改カラムが意図通りに発生しない事例を生む。

### 2-5. ヘッダー/フッター ── 開発者自身がバグを認めている既知制約

- `renderHeaderFooterRef`(L1271-)で first/even/default解決、高さ実測(`getOffsetHeight`)してpadding-top/bottomへ反映(L1148-1174)するロジック自体は存在する。
- だが`createPageContent`直前のコメント(L1268-1269)に**開発者自身のTODOが明記**: 「分页不准确，页脚页码混乱」(ページ分割が不正確、フッターのページ番号が混乱する)、「支持奇数页偶数页不同页眉页脚」(奇数/偶数ページで異なるヘッダーフッターは未対応、の意の未完了TODO)。
- 本サンプルはheaderReference/footerReferenceに`even`/`default`/`first`の3種類全てを使用しており、このTODOが直撃する構成。「実装済みだから動くはず」ではなく、**実装はあるが既知バグとして公言されている状態**と認識すべき。
- `src/header-footer/parts.ts`(36行)はパート読み込みのみの薄い層、ロジックはhtml-renderer-sync側に集約。

### 2-6. 行間設定の微妙なズレ ── docGrid補正がlineRule次第で適用漏れ

`src/document/spacing-between-lines.ts`に2つの確定済み欠陥:

- L71 開発者コメント「TODO 处理AtLeast，行高不准确」(AtLeastルールの行高処理が不正確、と明記)。`atLeast`時は`calc(100% + Xpt)`という近似式(L121-123)で、Wordの実際のベースライン計算と一致する保証がない。
- L142-158: 本サンプルの全sectPrには`<w:docGrid w:type="lines" w:linePitch="360"/>`が指定されている(実測確認済み)。このdocGrid補正は`lineSpacing['line-height']`が**数値型の場合のみ**(L148 `typeof === 'number'`)適用される。だが`lineRule="exact"`または`"atLeast"`の段落では、L122/L127で`line-height`が**既に文字列**(`calc(...)`や`"Xpt"`)にセットされているため、L148の条件に掛からず**docGrid補正がスキップされる**。lineRule次第で行間計算経路が分岐し、グリッド指定文書で見た目が揃わない不整合が生じる。

## 3. 設計方針

対症療法(個別パッチの積み上げ)でなく、以下3軸で構造的に再設計する。

- **責務分離**: 「XMLパース」「OOXML中間表現(IR)構築」「レイアウト計算」「DOM/SVG描画」を明確に分ける。現状DrawingMLとVMLが別実装で同じ概念(図形・パス・塗り)を別々に持っているのが拡張性を阻害する根。
- **公式仕様準拠の段階的強化**: 近似パスのハードコードから、ECMA-376準拠の計算へ移行(優先: 図形ジオメトリ、数式属性)。
- **計測ファースト**: レイアウト崩れは「静的記録を信用する」から「実測して検証・補正する」へ。lastRenderedPageBreakは初期ヒントとして使い、最終判定は常に実測オーバーフローに委ねる。

## 4. 再設計案

### 4-0. セクション/カラム/ページネーション統合エンジン(優先度最高、複数不具合の共通根)

2-1・2-4・2-5・2-6で確認した不具合は「1セクション=1ページ」という暗黙の前提と、ページ分割ロジックの未実装ケース(Continuous/NextColumn)・既知バグ(ヘッダーフッター混乱、docGrid補正漏れ)に共通の根を持つ。個別パッチでなく以下を一体で再設計する。

- `createPageContent`のTODO(L1242)を解消: セクション境界をページ内のサブ領域として扱えるレイアウトモデルへ変更し、1ページ内でカラム数が変わるセクション構成(Continuous断面)を表現できるようにする。
- `splitElementsBySymbol`のSectionBreak switch(L868-894)で空実装のままの`Continuous`/`NextColumn`ケースを実装する。
- カラムレイアウト内のブロック要素(数式・図形・表)に`break-inside: avoid`相当の制御を導入し、カラム幅を超える・カラム境界で分割されるべきでない要素を保護する。数式の「無駄な改行」はここで解消を狙う。
- ヘッダー/フッターのfirst/even/default解決とページ番号付与(L1268-1269のTODO)をセクション統合エンジンの一部として再実装し、テスト可能な単位に切り出す。
- `parseLineSpacing`(spacing-between-lines.ts)のdocGrid補正をlineRule(`auto`/`exact`/`atLeast`)全パターンに対して一貫した経路で適用するよう修正し、AtLeastの近似式(L121-123)をWordのベースライン計算に近づける。

### 4-1. 図形ジオメトリ統一エンジン(視覚インパクト最大)

- 新設: `ShapeGeometryIR { kind, basePaths: PathCommand[], adjustables: Record<string, GuideFormula> }` という中間表現を定義。
- ECMA-376 Part 1 §20.1.9.x の `prstGeom` 調整可能数式(`gd`/`avLst`)を解釈する計算エンジンを実装(主要十数種から段階導入、noSmoking含む)。
- `a:fillRef`/`a:lnRef`/`a:effectRef` → `theme.ts`の配色・線スキームへの解決関数を追加。`solidFill`直接指定と統合した「FillResolver」に一本化。
- DrawingML(`wps:wsp`)とVML(`v:shape`)を同じIRへ変換するアダプタを書き、レンダラ(SVG生成)を1本に統合。VML専用のcustGeom/グループ/回転対応もこのIR層で吸収する。
- 効果: noSmoking等の複合図形が正しい形状・正しいテーマ色で再現される。新規プリセット追加が「パス文字列のハードコード追加」でなく「公式数式の登録」になり拡張性が上がる。

### 4-2. 画像・lastRenderedPageBreakのオーバーフロー検出拡張

- 全要素(段落・画像・図形・テーブル)を`isSplit=false`相当の動的オーバーフロー検出対象に統一する。`lastRenderedPageBreak`は「初期分割ヒント」のレイヤーに格下げし、レンダリング後に実測offsetHeightで再検証・補正するフックを追加。
- 画像はCSS固定サイズ(EMU変換済み)のため計測自体は安定済み。問題は「検出ロジックの適用範囲」であり画像変換コード自体の修正は不要。
- 効果: 画像追加・差し替え時にWord保存時点の記録とズレてもページ境界が崩れない。4-0のセクション統合エンジンと組み合わせて初めて、セクション境界+画像オーバーフローの複合崩れが解消する。

### 4-3. 数式属性カバレッジ拡大(二次的精度向上)

「無駄な改行」自体の主因はカラムレイアウト未対応(2-1, 4-0で対応)。本項はその後に残る精度差分の話。

- `parseMathProperties`の対応属性をOOXML仕様の主要プロパティ(色/フォント/サイズ/太字italic/上下添字スケール比)まで拡張。
- MathML出力側のCSS強化(`mfrac`/`munderover`の行間調整)。4-0でカラム起因の折り返しを解消した後、残る微細なズレをここで詰める。

### 4-4. テキストボックス独立サイズ対応

- `wps:txbx`に親Shape依存のwidth:100%だけでなく、`a:xfrm`があれば自身のEMUサイズを優先する分岐を追加。回転変形時は`parseTransform2D`計算後の実寸を反映。

### 4-5. 回帰テスト整備

- `tests/golden`に「3セクション・カラム数混在・header even/default/first混在・docGrid lines指定」のケース(本サンプル相当)を追加し、4-0で修正した内容の退行を検出できるようにする。

## 5. 実装フェーズ

1. **Phase 0 計測基盤**: golden image diffテストにサンプル(zz-sample-analyze.docx)由来のケースを追加(セクション境界/カラム混在/数式/図形/ヘッダーフッター混在)。先にテストを赤くしてから着手。
2. **Phase 1 セクション/カラム/ページネーション統合エンジン**(4-0): 数式の「無駄な改行」・セクション区切りの改ページ漏れ・ヘッダーフッター混乱・行間ズレ、4つの不具合の共通根を一体で解消する最優先フェーズ。
3. **Phase 2 図形ジオメトリ統一エンジン**(4-1): 視覚インパクト最大の図形崩れに着手。DrawingML/VML共通IR化、FillResolver実装、主要プリセットの公式数式対応。
4. **Phase 3 オーバーフロー検出拡張**(4-2): 画像主体ページの動的検出をPhase 1の上に積む。
5. **Phase 4 数式属性拡張+テキストボックス仕上げ**(4-3, 4-4): 残る精度差分を詰める。
6. **Phase 5 回帰テスト整備**(4-5): 各フェーズの修正を恒久的に固定する。

各フェーズ完了条件: 該当golden testが通り、`zz-sample-analyze.docx`のレンダリング結果がWord出力(比較画像)と視覚的に一致すること。
