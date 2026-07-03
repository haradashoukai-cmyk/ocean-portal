# 回航ごとのルートレイヤー表示機能（実装プランメモ）

> 2026-07-02 時点の実装プラン。同日にこのプランどおり `saily.html` へ実装済み。

## Context（背景）

投稿は地図上にマーカーとして表示されるが、どの回航（航海）の記録なのか、どの順番で航海したのかが地図から読み取れない。回航ごとにレイヤーを作り、その回航期間中の投稿を時系列順にポリライン（線）でつなぎ、順番の番号を付けて航跡が一目で分かるようにする。

**確定事項:**
1. 投稿と回航の紐付け = **投稿日時で自動判定**（タイムスタンプが回航の startDate〜endDate 内なら紐付け。バックエンド変更なし）
2. 切替UI = **地図右上のレイヤーパネル**（回航ごとのチェックボックスで表示/非表示）
3. 線上の各投稿地点に **順番の番号**（1,2,3…）を表示

## 対象ファイル

**`saily.html` のみ**（単一HTMLアプリ、ビルドなし、Leaflet 1.9.4 CDN）。

主要アンカー: `VOYAGES`（L1036）、CSS（~L85 `.voyage-dot` 付近）、グローバル変数（L1392）、`initMap()`（L2069、pane作成 L2075-2078、notice control L2079-2086）、`renderPosts()`（L2136）、`renderVoyageMarkers()`（L2243 の直後に新関数群）、`focusPost()`（L2295）、`formatDate()`（L2327 — 日時パースの既存パターン）。

## 実装ステップ

### 1. データ: VOYAGES に色を追加（L1036）
既存 `V2026-01` に `color: '#e11d48'` を追加。新JSブロックにフォールバックパレットを定義:
```js
const ROUTE_COLORS = ['#e11d48', '#f97316', '#7c3aed', '#0d9488', '#ca8a04'];
```
（既存マーカーが紺/ティール/アンバーなので、海図上で映える暖色・紫系）

### 2. CSS 追加（`.voyage-dot` の後）
- `.route-seq` — 20px の番号バッジ（丸、白縁、回航色の背景、クリック可）
- `.voyage-layer-control` — 右上パネル（白半透明、角丸、コンパクト）。`.vlc-head`（タイトル行、クリックで開閉）、`.collapsed .vlc-body { display:none }`、`.vlc-row`（チェックボックス行）、`.vlc-swatch`（色見本）、`.vlc-count`（件数）
- 既存の `#mapView.popup-active .timeline-toggle` ルール（L93付近）に `.voyage-layer-control` を追記し、ポップアップ表示中はパネルをフェード

### 3. グローバル変数（L1392 の直後）
```js
let voyageRouteLayers = {}, voyageRouteVisibility = {};
```
`voyageRouteVisibility` はチェック状態（回航IDキー）。`renderPosts()` の再構築をまたいで保持。

### 4. 新JS関数群（`renderVoyageMarkers()` の直後、L2250 以降）

**4a. 日時ヘルパー + 紐付け判定**（投稿は `YYYY-MM-DD HH:MM:SS`、回航は `YYYY/MM/DD` — フォーマット差に注意。既存 `formatDate()` と同じ `replace(' ', 'T')` でローカル時刻パース）:
```js
function postTime(p) { const t = new Date(String(p.timestamp || '').replace(' ', 'T')).getTime(); return Number.isFinite(t) ? t : NaN; }
function voyageDateMs(s, endOfDay) { const [y, m, d] = String(s).split('/').map(Number); return new Date(y, m - 1, d, endOfDay ? 23 : 0, endOfDay ? 59 : 0, endOfDay ? 59 : 0).getTime(); }
function voyageForPost(p) { /* 座標・日時が有効で、期間内（endDate は 23:59:59 まで含む）の最初の回航を find で返す。なければ null */ }
```

**4b. `renderVoyageRoutes()`** — 全レイヤー再構築:
- `voyageRouteLayers` の全グループを map から remove して `{}` にリセット
- `posts` を `voyageForPost` で回航ごとにグループ化 → 各グループを `postTime` 昇順ソート（バックエンドは新しい順なので必須）
- 回航ごとに `L.layerGroup()` を作成: 投稿2件以上なら `L.polyline(..., { pane: 'routePane', color, weight: 3, opacity: 0.8, dashArray: '6 6' })`、各投稿地点に番号バッジ `L.divIcon`（`.route-seq`、`iconAnchor` を写真マーカー(42px)に被らないよう左上にオフセット）
- バッジの click は **`focusPost(p.id)` を呼ぶ**（既存のクラスタ対応ズーム+ポップアップを再利用。クリック時に `markers[id]` を解決するので再構築後も安全）
- 件数を保持し、`voyageRouteVisibility[v.id] ??= (v.visible !== false)` で初期化、ON なら `group.addTo(map)`
- 最後に `updateVoyageLayerControl()` を呼ぶ

**4c. `updateVoyageLayerControl()`** — パネル DOM（`#voyageLayerControl`）を innerHTML で再構築。ヘッダ「回航ルート」（クリックで開閉）+ 回航ごとの行（チェックボックス、色見本、回航名、`N件`。0件は disabled）。inline onchange で `toggleVoyageRoute()`（既存コードのスタイルに合わせる）

**4d. `toggleVoyageRoute(voyageId, checked)`** — `voyageRouteVisibility` を更新し、layerGroup を addTo/removeLayer

### 5. `initMap()` へのフック
- `voyagePane` 作成（L2078）の後に: `routePane` を z-index **660** で作成（portPane 650 と voyagePane 670 の間。投稿写真マーカー(markerPane 600)の上、ポップアップ(1200)の下）
- `notice.addTo(map)`（L2086）の後に: 既存 notice と同じパターンで `L.control({ position: 'topright' })` を作成。`L.DomUtil.create('div', 'voyage-layer-control')`、`L.DomEvent.disableClickPropagation` / `disableScrollPropagation`。**モバイル（max-width: 768px）では初期状態 collapsed**

### 6. `renderPosts()` へのフック（L2136）
`updateMenuCounts();`（L2143）の直後、**投稿0件の early return より前**に `renderVoyageRoutes();` を挿入。posts が空になっても古いレイヤーが確実にクリアされ、作成/編集/削除の全経路を1箇所でカバー。`fitBounds`（L2151）は変更不要（線の頂点＝投稿座標で既に bounds に含まれる）。

## エッジケース
- 期間内投稿0件: レイヤーなし、パネル行は disabled「0件」
- 1件のみ: バッジ「1」のみ、線なし
- 座標・日時が不正な投稿: `voyageForPost` で除外
- どの回航にも属さない投稿: 従来どおり（変更なし）
- 期間が重複する回航: `find` で最初の1つのみに紐付け

## 検証方法

GAS の実データに依存するため、Playwright + Chromium でスタブ検証:

1. `python3 -m http.server 8899`（リポジトリルート、バックグラウンド）
2. Playwright スクリプト: `page.route('**?action=get*')` で投稿5件のフィクスチャを返す — 回航期間内4件（**配列は時系列バラバラ**にしてソートを証明、うち1件は `2026-05-24 23:00:00` で endDate 終日包含を証明）+ 期間外 `2026-06-15` 1件
3. 確認項目:
   - (a) 右上に「回航ルート」パネル、「新居浜→統栄2026年5月 4件」チェック済み
   - (b) `routePane` にポリライン1本 + `.route-seq` バッジが時系列順に 1〜4
   - (c) チェックOFF→線とバッジが消える、ON→復活
   - (d) バッジクリック→該当投稿のポップアップが開く
   - (e) 期間外の6月投稿にはバッジなし（通常マーカーのみ）
   - (f) ビューポート375×700でパネルが collapsed 起動、タップで展開
4. コンソールで `posts = posts.slice(1); renderPosts()` → 3件で再構築され線が重複しない（編集/削除経路の確認）
5. スクリーンショットを取得して確認

---

# 追加機能: 区間色分け（往路/復路）と寄港中バッジ（2026-07-03）

## 概要

- **区間（legs）**: 回航に日付の区切りを設定すると、線とバッジの色が区間ごとに変わる（例: 往路=赤、復路=紫）。番号は通し番号のまま。区間未設定の回航は従来どおり1色
- **寄港中バッジ**: 登録済みの港から**2km以内**の投稿は、番号バッジが「⚓番号」のピル型に自動で変わる。港リスト（スプレッドシート）に港を追加すればその港も判定対象になる
- パネルには区間の凡例と「⚓ = 港に滞在中」の注記を表示

## 区間の設定方法（後日、区切り日が決まったとき）

`saily.html` の `VOYAGES` の該当回航に `legs` を追加するだけ:

```js
{
  id: 'V2026-01',
  name: '新居浜→統栄2026年5月',
  ...
  legs: [
    { name: '往路', color: '#e11d48' },                        // 開始日省略 = 回航の開始日から
    { name: '復路', startDate: '2026/05/19', color: '#7c3aed' } // この日の0時以降は復路
  ]
}
```

- 区間はいくつでも追加できる（例: 往路・上陸・復路の3つ）
- `color` を省略すると自動で色が割り当てられる
- 投稿は「開始日が自分の日時以前である最後の区間」に入る

## 実装メモ

- `voyageLegs(v)` / `postLegIndex(legs, p)` — 区間の正規化と投稿の区間判定
- `isNearPort(p)` — 等距円筒近似で港との距離を計算（投稿数×港数の単純ループ、負荷は無視できる）
- ポリラインは区間の連続区切りごとに分割描画し、直前区間の最終地点を先頭に加えて線を途切れさせない。境目の線は新しい区間の色
- `loadPorts()` 完了後に `renderVoyageRoutes()` を呼び、実データの港で⚓判定を更新
