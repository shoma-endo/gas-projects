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
