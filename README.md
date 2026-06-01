# Majin Plus

[日本語](#japanese) | [English](#english)

---

<a name="japanese"></a>
## 日本語

### 概要

**Majin Plus** は、Google Apps Script（GAS）アプリケーションとして公開されている [まじん式 v3](https://note.com/majin_108/n/n11fc2f2190e9) をベースにしたローカルHTMLアプリケーションです。このプロジェクトは、まじん式の生産性向上メソッドをスタンドアロンのローカルアプリとして再実装し、特にNotionユーザー向けの機能を強化しています。

### 特徴

- **ローカルアプリケーション**: 外部サービスを必要とせず、ブラウザ上で完結
- **強化されたスライドタイプ**: Notionユーザー向けに特別に設計された追加のスライド形式
- **プライバシー重視**: すべてのデータはローカルマシン上に保存
- **使いやすさ**: インストール不要のシンプルなHTMLベースのインターフェース

### 基本的な使い方

1. `index.html` をブラウザで開く
2. （必要であれば）テンプレートとなる `.pptx` ファイルを追加する
3. 色のマッピングを行う
4. `majin-plus_v6.md` を Notion AI スキルに登録したうえで、Notion ページをもとに「アウトプットを出して」と伝える
5. スライドの構成を確認されると同時に `template_analyzer.md` があるか聞かれるので、必要であればコピーして貼り付ける。スライドの構成も必要であれば修正する
6. プライマリーカラーとフォントは好みのものを選ぶ
7. 「スライドを生成」ボタンを押す

### 謝辞

このプロジェクトは、**まじん式**とその開発者である [@majin_108](https://note.com/majin_108) さんのおかげで作成されました。
[note.comのまじん式 v3](https://note.com/majin_108/n/n11fc2f2190e9)

### ライセンス

このプロジェクトはMITライセンスの下でライセンスされています。詳細は[LICENSE](LICENSE)ファイルをご覧ください。

Contributions are welcome! Please feel free to submit issues or pull requests.

### Contact

For questions or feedback, please open an issue on this repository.

---

<a name="english"></a>
## English

### Overview

**Majin Plus** is a local HTML application based on the [Majin Method v3](https://note.com/majin_108/n/n11fc2f2190e9) originally published as a Google Apps Script (GAS) application. This project reimplements the Majin-style productivity method as a standalone local app with enhanced features, particularly tailored for Notion users.

### Features

- **Local Application**: Runs entirely in your browser without requiring external services
- **Enhanced Slide Types**: Additional slide formats designed specifically for Notion users
- **Privacy-Focused**: All data stays on your local machine
- **Easy to Use**: Simple HTML-based interface requiring no installation

### Basic Usage

1. Open `index.html` in your browser
2. (If needed) Add a `.pptx` file to use as a template
3. Configure the color mapping
4. Register `majin-plus_v6.md` as a Notion AI skill, then open your Notion page and say "Generate output"
5. You will be asked to confirm the slide structure and whether `template_analyzer.md` is available — paste it in if needed, and adjust the slide structure as necessary
6. Choose your preferred primary color and font
7. Click the "Generate Slides" button

### Acknowledgments

This project is built on the **Majin Method** (まじん式) and its creator [@majin_108](https://note.com/majin_108).
[Majin Method v3 on note.com](https://note.com/majin_108/n/n11fc2f2190e9)

### License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
