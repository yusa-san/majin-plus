# Majin Plus

[日本語](#japanese) | [English](#english)

---

<a name="japanese"></a>
## 日本語

### 概要

**Majin Plus** は、Google Apps Script（GAS）アプリケーションとして公開されている [まじん式 v3](https://note.com/majin_108/n/n11fc2f2190e9) をベースにしたローカルHTMLアプリケーションです。このプロジェクトは、まじん式の生産性向上メソッドをスタンドアロンのローカルアプリとして再実装し、特にNotionユーザー向けの機能を強化しています。
<img width="1903" height="1027" alt="image" src="https://github.com/user-attachments/assets/6795f9bd-906b-4e04-9b90-f0f45a1cad9d" />

### 特徴

- **ローカルアプリケーション**: 外部サービスを必要とせず、ブラウザ上で完結
- **強化されたスライドタイプ**: Notionユーザー向けに特別に設計された追加のスライド形式
- **プライバシー重視**: すべてのデータはローカルマシン上に保存
- **使いやすさ**: インストール不要のシンプルなHTMLベースのインターフェース

### 基本的な使い方

1. `index.html` をブラウザで開く（インストール不要）
2. （テンプレートを使う場合）アプリ右側「デザイン設定」パネルの「ファイルを選択(.pptx)」からテンプレートPPTXを読み込む
3. 色のマッピングを行う：「アクセントカラーの対応付け」で、テンプレートのテーマカラー（色スウォッチ）を slideData の論理名（`blue` / `red` など）に割り当てる。テンプレートなしの場合は標準パレットが表示される
4. `majin-plus_v6.md` を Notion AI のカスタムスキルとして登録したうえで、スライド化したい Notion ページを開き「アウトプットを出して」と伝える
5. AI にスライド構成を確認される。同時に `template_analysis.md` があるか聞かれるので、テンプレートを使う場合はアプリの「📄 template_analysis.md を生成」ボタンを押してコピーし、チャットに貼り付ける（テンプレートなしの場合はそのまま「標準で」と答える）。構成も必要に応じて修正する
6. AI が出力した slideData（JSON）をコピーし、アプリ左側「スライドデータを入力」欄に貼り付ける
7. プライマリーカラーとフォントを好みのものに設定する
8. 「プレゼンテーションを生成」ボタンを押す

### 謝辞

このプロジェクトは、**まじん式**とその開発者である [@majin_108](https://note.com/majin_108) さんのおかげで作成されました。
[note.comのまじん式 v3](https://note.com/majin_108/n/n11fc2f2190e9)

### ライセンス

このプロジェクトはMITライセンスの下でライセンスされています。詳細は[LICENSE](LICENSE)ファイルをご覧ください。

### お問い合わせ

コントリビューション歓迎。Issue や Pull Request はお気軽にどうぞ！
質問やフィードバックは、このリポジトリの Issue からお願いします。

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

1. Open `index.html` in your browser (no installation required)
2. (If using a template) In the "Design Settings" panel on the right, click "Select File (.pptx)" to load your template PPTX
3. Configure color mapping: in the "Accent Color Mapping" section, assign each template theme color (color swatches) to a logical name used in slideData (`blue`, `red`, etc.). If no template is loaded, the default palette is shown instead
4. Register `majin-plus_v6.md` as a custom Notion AI skill, then open the Notion page you want to turn into slides and say "Generate output"
5. The AI will ask you to confirm the slide structure. It will also ask whether you have a `template_analysis.md` — if you loaded a template in step 2, click the "📄 Generate template_analysis.md" button in the app, copy the output, and paste it into the chat. If not using a template, simply reply "Use standard layout." Adjust the slide structure as needed
6. Copy the slideData (JSON) output from the AI and paste it into the "Enter Slide Data" field on the left side of the app
7. Set your preferred primary color and font
8. Click the "Generate Presentation" button

### Acknowledgments

This project is built on the **Majin Method** (まじん式) and its creator [@majin_108](https://note.com/majin_108).
[Majin Method v3 on note.com](https://note.com/majin_108/n/n11fc2f2190e9)

### License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

### Contact

Contributions are welcome! Please feel free to submit issues or pull requests.
For questions or feedback, please open an issue on this repository.
