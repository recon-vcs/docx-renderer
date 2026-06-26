# big-refactor PR レビュー

> 観点: 対症療法的でないか、体系的か、洗練されているか、拡張性が高いか

---

## 総評

構造の方向性は正しい（OOXML model / parsing / rendering / pagination の層分離）が、
実装には複数の根本的な設計問題が残っている。
「体系的に見える構造」の裏に、責務の混乱・二重実装・汚染されたモデルが潜んでいる。

---

## 問題 1: Page クラスがレイヤー境界を破壊している

**ファイル**: `src/rendering/pagination/model/page.ts`

`Page` が `OpenXmlElement` を implements している。
これはレンダリング時の概念がOOXMLモデルレイヤーに混入することを意味する。

```ts
export class Page implements OpenXmlElement {
    type: DomType;   // ← OOXML の DomType を持つ
    children: OpenXmlElement[];
    regions?: LayoutRegion[];  // ← LayoutRegion も持つ
    physicalPage?: PhysicalPage;  // ← 計画フェーズの概念も持つ
}
```

`Page` は `children`（OOXML要素の配列）と `regions`（レイアウト計画の配列）を
**同時に持つ**。これらは同じ内容を別の形式で保持している二重表現であり、
どちらが真実かが常に曖昧。`splitRegionOnOverflow` 内では
`currentPage.children = currentRegions.flatMap(item => item.children)` という
同期コードが必要になっており、これが対症療法の典型。

**根本問題**: OOXML model・render-time page・物理ページ計画の3概念が1クラスに同居している。

---

## 問題 2: ページネーションの二重実装

**ファイル**: `src/rendering/pagination/core/pagination.ts` と `src/rendering/pagination/core/modern-page-splitter.ts`

両者は同じことをしている:

```ts
// pagination.ts
export function buildPaginationPlan(bodyChildren, rootSectProps): PaginationPlan {
    const sectionRegions = buildSectionStream(bodyChildren, rootSectProps);
    const regions = splitRegionsByExplicitBreaks(sectionRegions);
    const pages = buildPhysicalPages(regions);
    return { regions, pages };
}

// modern-page-splitter.ts
export function splitDocumentIntoPhysicalPages(documentElement): ModernPageSplit {
    const sectionRegions = normalizeRegionSections(
        buildSectionStream(documentElement.children, documentElement.sectProps)
    );
    const regions = splitRegionsByExplicitBreaks(sectionRegions);
    return { regions, pages: buildPhysicalPages(regions) };
}
```

`modern-page-splitter.ts` に `normalizeRegionSections`（ヘッダー/フッター継承）が
追加されているだけで、本質的に同一の処理フローが2つ存在している。
ファイル名に "modern" が付いている時点で「古い実装を置き換えたが消せていない」状態の証拠。
実際に使われているのは `modern-page-splitter.ts` 側のみで、`pagination.ts` は
テストからのみ参照されている。

---

## 問題 3: HtmlRendererSync が God Object のまま

**ファイル**: `src/html-renderer-sync.ts` (676行)

表面上は各種 `*Fn` 関数にデリゲートしているが、
`HtmlRendererSync` クラス自体が全ての状態（session、options、document、styleMap、etc.）を
所有し続けている。`documentStylesCallbacks()` や `numberingStylesCallbacks()` が
毎回クロージャを生成して `this` を包み込む構造は、
「モジュールに分割した」という外見だけで実質的に分離されていない。

```ts
// 676行のクラスが、30個以上のモジュールに this を渡し続ける
private documentStylesCallbacks() {
    return {
        styleToString: (selectors, declarations, cssText = null) => this.styleToString(...),
        processStyleName: (className) => this.processStyleName(className),
        loadFont: (id, key) => this.document.loadFont(id, key),
        // ...
    };
}
```

また `renderPages` 内での `origin_pages` という snake_case 変数、
`this.session.pages` へのミューテーション、ループ中での `page.isFirstPage` 設定など、
手続き的なミューテーションが散在している。

---

## 問題 4: splitRegionOnOverflow が抽象化を壊している

**ファイル**: `src/rendering/pagination/model/page-split.ts`

```ts
export function splitRegionOnOverflow(...) {
    // LayoutRegion を扱う関数なのに Page を一時的に作る
    const currentWrapper = new Page({ sectProps: currentRegion.section, children: ... });
    const nextWrapper = new Page({ sectProps: nextRegion.section, children: ... });
    splitElementsByBreakIndex(currentWrapper, nextWrapper);
    currentRegion.children = currentWrapper.children;  // 取り出す
    nextRegion.children = nextWrapper.children;         // 取り出す
```

`LayoutRegion` を分割するために `Page` オブジェクトを作り、
`splitElementsByBreakIndex` を呼んで、また `children` を取り出す。
`splitElementsByBreakIndex` が `Page` に依存しているため、
`LayoutRegion` で直接呼べない。これは `Page` の設計ミスが引き起こした対症療法。

---

## 問題 5: Overflow 状態機械が過度に複雑

**ファイル**: `src/rendering/measurement/overflow.ts`, `overflow-measurer.ts`

```ts
export enum Overflow {
    TRUE = 'true',
    FALSE = 'false',
    SELF = 'self',
    FULL = 'full',
    PART = 'part',
    UNKNOWN = 'undetected',
    IGNORE = 'ignore',
}
```

7つの状態。`inferOverflow` は5段階の条件チェック。
呼び出し側では `isOverflow !== Overflow.FALSE && isOverflow !== Overflow.UNKNOWN && isOverflow !== Overflow.IGNORE`
という否定の連鎖が必要になっており、読み手が状態の意味を把握するのが困難。
各状態が何を意味し、いつセットされ、どこで消費されるかのドキュメントがない。

---

## 問題 6: 中国語コメントが残留している

**ファイル**: `src/rendering/dom/core/element-dispatcher.ts`, `element-processor.ts`

```ts
// 作为子元素插入，执行溢出检测
// 渲染换行符号
// 分页符
// 处理表格style样式
// 递归明确元素parent父级关系
```

コードの出自が別プロジェクト（中国語圏のfork）であることを示している。
コメント言語の混在（中国語・英語・なし）は、コードの知的所有権と
保守責任の所在が曖昧であることを示唆する。

---

## 問題 7: パース済みモデルへのレンダリング時ミューテーション

**ファイル**: `src/html-renderer-sync.ts`

```ts
private assignSourcePaths(children: OpenXmlElement[]) {
    children.forEach((child, index) => {
        child.sourcePath = path;  // ← パース済みモデルを書き換える
```

また `element-processor.ts` の `processElement` では:

```ts
e.parent = element;  // ← OOXML要素に parent を後付けする
e.level = element?.level + 1;
```

パース済みの不変モデルが、レンダリングフェーズで変更されている。
これにより同一の `WordDocument` を複数回レンダリングすると
`parent` / `sourcePath` / `level` が上書きされ、
再利用時の安全性が保証されない。

---

## 問題 8: コールバックインターフェースの乱立

**ファイル群**: `element-dispatcher.ts`, `page-renderer.ts`, `inline-renderer.ts`, `drawing-renderer.ts` 等

`ElementDispatchContext`, `PageRendererCallbacks`, `InlineRendererCallbacks`,
`MathRendererCallbacks`, `FieldsRendererCallbacks`, `DrawingRenderContext` が全て存在し、
全て `HtmlRendererSync` の `this` で満たされる。
インターフェース自体は良いが、全部が同じGodオブジェクトに繋がっているため、
「モジュール化されている」という外見だけで、実質的な依存の独立性はゼロ。
個別モジュールを独立してテストできない構造になっている。

---

## まとめ: 対症療法か体系的か

| 箇所 | 判定 |
|---|---|
| OOXML model / parsing / rendering の層分離 | ✅ 体系的 |
| `parsing/parse-context.ts` のnarrow context分割 | ✅ 体系的 |
| `pagination/` 以下のファイル構造 | ✅ 体系的 |
| `Page implements OpenXmlElement` | ❌ 根本設計ミス |
| `pagination.ts` と `modern-page-splitter.ts` の共存 | ❌ 二重実装 |
| `splitRegionOnOverflow` の一時 Page 生成 | ❌ 対症療法 |
| `HtmlRendererSync` の callback delegation | ❌ 偽の分離 |
| パース済みモデルへのレンダリング時ミューテーション | ❌ 設計混在 |

層の名前は正しく付けられているが、`Page` クラスが複数の層を跨いでいる点が
最も根本的な問題。ここを直すと残りの問題の多くが連鎖的に解決できる見込みがある。
