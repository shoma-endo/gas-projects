# Google Apps Script Projects

このリポジトリは、複数のGoogle Apps Script (GAS) プロジェクトを管理するモノレポです。
各ディレクトリは独立したGASプロジェクトとして構成されており、`clasp` を使用して管理されています。

## プロジェクト一覧

| プロジェクト名 | 概要 |
| :--- | :--- |
| **automatic-submission-sheet-creation** | **入稿シート自動作成システム**<br>OpenAI API (GPT-4o) を利用して、自然言語の与件テキストから広告媒体ごとの入稿情報を構造化して抽出し、スプレッドシートに自動入力します。 |
| **elme-access-control-system** | **入退室管理システム**<br>Webアプリケーションとしてデプロイし、入室・退室のログをスプレッドシートに記録します。QRコード等からのアクセスを想定した簡易認証トークン機能付き。 |
| **instagram-base-integration** | **Instagram基盤連携ツール**<br>Instagram Graph APIを利用して、自社アカウントのインサイト（フォロワー属性、リーチ、インプレッションなど）を定期的に取得し、スプレッドシートに蓄積します。 |
| **instagram-competitor-analysis** | **Instagram競合分析ツール**<br>指定した競合アカウントの投稿データやパフォーマンス指標（いいね、コメント数など）を取得し、自社アカウントと比較分析するためのデータを収集します。 |
| **lark-ocr** | **Lark OCR連携検証**<br>Lark (Lark Suite) のAPIとGemini APIを組み合わせて、アップロードされたファイルのOCR処理や解析を行うための検証用プロジェクトです。 |
| **sim-creation** | **広告シミュレーション作成ツール**<br>動画・静止画広告の予算、CPM、CTRなどのパラメータから、KPI（表示回数、クリック数、視聴単価など）を試算し、シミュレーションシートを出力するWebツールです。 |
| **youtube-data-acquisition** | **YouTubeデータ収集ツール**<br>YouTube Data API v3を利用して、指定したチャンネルの登録者数、再生回数、動画数などの統計情報を定期的に取得し、スプレッドシートに記録します。 |
| **LarkCalendar-Base-Task** | Larkカレンダー＆タスクをLark内Baseに自動同期（差分更新） |
| **OfficialLine-LarkBase** | LINEBotを使い、Webhook受信・Lark内Baseからのメッセージ送信・LINE Login移行を統合処理 |
| **proLine-LarkBase** | 外部WebhookをスプレッドシートでLark内Baseへ中継（フィールド自動生成・重複防止付き） |
| **UTAGE-LarkBase** | UTAGEの複数シートをLark内Baseへ一括アップロード（日付・数値自動変換） |

## 開発環境のセットアップ

このリポジトリは `clasp` (Command Line Apps Script Projects) を使用しています。

### 1. 依存関係のインストール

```bash
npm install -g @google/clasp
```

### 2. Googleアカウントでのログイン

```bash
clasp login
```

### 3. 各プロジェクトの連携

各ディレクトリに移動し、既存のGASプロジェクトと連携していない場合は `.clasp.json` を設定するか、新規作成してください。

```bash
cd <project-directory>
# 既存のスクリプトIDと紐付ける場合
clasp setting scriptId <YOUR_SCRIPT_ID>
# または
clasp clone <YOUR_SCRIPT_ID>
```

## 共通の運用ルール

*   **機密情報の管理**: APIキー、トークン、Webhook URL、特定の環境に依存するID（スプレッドシートIDなど）はコードにハードコーディングせず、**スクリプトプロパティ (Script Properties)** を使用してください。
*   **プッシュとプル**: コードの変更は `clasp push` でGASエディタに反映し、ブラウザ上での変更は `clasp pull` でローカルに取り込んでください。
了解。以下は **README.md にそのまま貼り付け可能な日本語版 Markdown です**（最小限のブランチ命名ガイドライン）。

## 📌 ブランチ命名ルール（最低限ガイドライン）

このリポジトリは複数の Google Apps Script プロジェクトを含むモノリポ構成です。
どのプロジェクトに対する変更か判別しやすくするため、ブランチ作成時に以下の命名ガイドラインを推奨します。

### ✏️ 命名フォーマット（ガイドライン）

```
<project-name>/<type>/<brief-description>
```

#### 各要素の意味

* **project-name**
  対象プロジェクトのフォルダ名
  例: `instagram-base-integration`、`sim-creation`、`youtube-data-acquisition` など

* **type**（任意）
  変更種別を簡潔な英語で表す

  * `feat` : 新機能
  * `fix` : バグ修正
  * `refactor` : リファクタリング
  * `docs` : ドキュメント
  * `chore` : 保守／設定変更

* **brief-description**
  短い説明（ハイフン区切り）

#### 例

```
instagram-base-integration/feat/add-metrics-sync
sim-creation/fix/budget-calculation-error
youtube-data-acquisition/docs/update-readme
```

## ⚠️ 運用における注意
* ブランチ名は **対象プロジェクトが何か一目で分かること** を優先してください。
* `type` や `brief-description` は理解を助けるため推奨しますが、必須ではありません。

これは厳密なルールではなく、**運用上の目安**です。必要に応じて柔軟に運用してください。
