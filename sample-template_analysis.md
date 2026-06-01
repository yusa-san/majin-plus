# template_analysis.md

> このファイルは、テンプレートPPTXを定型解析して書き出した「設計情報」です。
> まじん式+.md と元データに**この情報を併せて**与え、テンプレの構成・配色・レイアウトを活かした slideData を生成してください。
> ※テンプレ例スライドの本文・図表の中身は含みません（構成・型・配置のみ）。

## 1. メタ情報
- スライドサイズ: 13.33 × 7.5 in
- 形式: 16:9
- レイアウト数: 15 / 例スライド数: 29

## 2. カラーパレット（テンプレ色 → 9論理名 → 役割）

| theme | HEX | 最近傍の論理名 | 役割 |
| --- | --- | --- | --- |
| dk1 | #000000 | black | 文字/濃色 |
| lt1 | #FFFFFF | white | 背景/淡色 |
| dk2 | #42A5F5 | skyblue | 文字/濃色 |
| lt2 | #FFFFFF | white | 背景/淡色 |
| accent1 | #42A5F5 | skyblue | メインカラー候補 |
| accent2 | #4CAF50 | green | メインカラー候補 |
| accent3 | #FFC107 | yellow | アクセント候補 |
| accent4 | #1565C0 | blue | アクセント候補 |
| accent5 | #FF7043 | orange | アクセント候補 |
| accent6 | #F44336 | red | アクセント候補 |
| hlink | #2E7D32 | deepgreen | リンク |
| folHlink | #000000 | black | リンク |

> slideData の `accentColor` は上表の**論理名**で指定。メインカラーは accent1またはaccent2（または役割「メインカラー候補」）を基調に統一してください。

## 3. フォント
- 見出し(major): メイリオ
- 本文(minor): ＭＳ 明朝
- テンプレに現れる全フォント: Century Gothic / メイリオ / Georgia / ＭＳ 明朝 / BIZ UDPゴシック / Meiryo レギュラー / Century Gothic レギュラー / 游明朝 / Meiryo / Windows Office Compatible MS Mincho / Georgia レギュラー

## 4. レイアウトカタログ（まじん式タイプへの対応）

| id | 分類 | 推奨まじん式タイプ | 例スライドでの使用 | タイトル枠 | 本文枠 | 図枠 | 列 |
| --- | --- | --- | --- | --- | --- | --- | --- |
| L1 | title | title | — | あり | — | — | 1 |
| L2 | content | content | — | あり | 1枠 | 1枠 | 1 |
| L3 | content | content | — | あり | 1枠 | 1枠 | 1 |
| L4 | content | content | — | — | 1枠 | 1枠 | 1 |
| L5 | other | content | slide 29 | — | 1枠 | — | 1 |
| L6 | other | content | — | — | 1枠 | — | 1 |
| L7 | other | content | — | — | 1枠 | — | 1 |
| L8 | title | title | slide 1 | あり | 1枠 | — | 1 |
| L9 | title | title | — | あり | 1枠 | 1枠 | 1 |
| L10 | title | title | slide 2,7,18,22,27 | あり | 1枠 | — | 1 |
| L11 | twoCol | content（twoColumn） | — | あり | 2枠 | — | 2 |
| L12 | content | content | — | あり | 1枠 | — | 1 |
| L13 | content | content | slide 3,4,5,6,17,19,20,21,23,24,25,26,28 | あり | 1枠 | — | 1 |
| L14 | content | content | slide 9,10,11,12,13,14,15,16 | あり | 1枠 | — | 1 |
| L15 | content | content | slide 8 | あり | 1枠 | — | 1 |

### 重要原型の詳細（幾何座標・テキスト許容量）

#### L8（title / title） — 表紙2
- タイトル枠: left 0.94, top 1.23, w 11.44, h 1.18 (in)
- 本文枠1: left 0.94, top 4.75, w 6.99, h 0.36 (in)
- テキスト許容量(目安): タイトル≈116字 / 本文≈55字×1行
- 装飾図形数: 2（背景あり）

#### L11（twoCol / content（twoColumn）） — 1_タイトルとテキスト_ヘッドラインあり
- タイトル枠: left 0.47, top 0.36, w 12.39, h 0.47 (in)
- 本文枠1: left 0.47, top 1.56, w 12.39, h 3.59 (in)
- 本文枠2: left 0.47, top 1.16, w 6.59, h 0.4 (in)
- テキスト許容量(目安): タイトル≈63字 / 本文≈99字×10行
- 装飾図形数: 2（背景あり）

#### L13（content / content） — タイトルのみ
- タイトル枠: left 0.47, top 0.36, w 12.39, h 0.47 (in)
- 本文枠1: left 0.47, top 1.06, w 12.39, h 5.54 (in)
- テキスト許容量(目安): タイトル≈63字 / 本文≈99字×15行
- 装飾図形数: 1（背景あり）

## 5. 例スライド構成サマリ（中身なし・構成のみ）

| # | 参照レイアウト | 分類 | 推奨まじん式タイプ | 要素(本文枠/図形/画像/線/表) |
| --- | --- | --- | --- | --- |
| 1 | L8 | title | title | 0/1/1/1/0 |
| 2 | L10 | title | title | 0/0/0/0/0 |
| 3 | L13 | content | content | 0/3/2/0/0 |
| 4 | L13 | cards | cards / headerCards | 0/6/2/1/0 |
| 5 | L13 | cards | cards / headerCards | 0/12/1/0/0 |
| 6 | L13 | table | table | 0/2/4/0/1 |
| 7 | L10 | title | title | 0/0/0/0/0 |
| 8 | L15 | cards | cards / headerCards | 0/10/8/0/0 |
| 9 | L14 | cards | cards / headerCards | 0/6/2/0/0 |
| 10 | L14 | cards | cards / headerCards | 0/6/2/0/0 |
| 11 | L14 | cards | cards / headerCards | 0/6/2/0/0 |
| 12 | L14 | cards | cards / headerCards | 0/6/2/0/0 |
| 13 | L14 | cards | cards / headerCards | 0/7/2/0/0 |
| 14 | L14 | cards | cards / headerCards | 0/6/2/0/0 |
| 15 | L14 | cards | cards / headerCards | 0/6/2/0/0 |
| 16 | L14 | cards | cards / headerCards | 0/6/2/0/0 |
| 17 | L13 | content | content | 0/2/2/0/0 |
| 18 | L10 | title | title | 0/0/0/0/0 |
| 19 | L13 | cards | cards / headerCards | 0/15/6/0/0 |
| 20 | L13 | cards | cards / headerCards | 0/13/5/0/0 |
| 21 | L13 | cards | cards / headerCards | 0/10/5/0/0 |
| 22 | L10 | title | title | 0/0/0/0/0 |
| 23 | L13 | cards | cards / headerCards | 0/4/3/0/0 |
| 24 | L13 | cards | cards / headerCards | 0/5/2/0/0 |
| 25 | L13 | content | content | 0/3/2/0/0 |
| 26 | L13 | cards | cards / headerCards | 0/10/7/0/0 |
| 27 | L10 | title | title | 0/0/0/0/0 |
| 28 | L13 | cards | cards / headerCards | 0/10/5/0/0 |
| 29 | L5 | other | content | 0/0/1/0/0 |

## 6. slideData 生成時の指示（このテンプレに合わせる）

- **テンプレが持つ型を優先**: 上のカタログ「推奨まじん式タイプ」と例スライドの並び・粒度に倣って slideData の type を選ぶ。
- **配色**: `accentColor`/`color` は §2 の論理名で指定。メインカラー1色を全 `section`/`title`/`closing`/章末 `iconBanner` に統一（まじん式+ §6.2）。
- **フォント**: 文字数は §4 の「テキスト許容量(目安)」を超えない。タイトル/本文の字数上限を厳守。
- **レイアウト指定（任意）**: 特定のテンプレ原型に強く寄せたいスライドは、共通プロパティ `templateLayout` にカタログの id（例 "L1"）を指定してよい。**座標は書かない**（描画ツールがテンプレ実枠へ配置する）。
- **図の差し込み**: 図/写真/図表を入れたいスライドは共通プロパティ `figure` に「何を入れるか」を記述（ツールがダミー枠を描画）。図枠ありレイアウト: L2, L3, L4, L9 を `templateLayout` に指定するとテンプレ図位置に入る。 **画像URLは生成・推定しない**。
- **例スライドの本文・図表は引用しない**（構成だけを参考にする）。
