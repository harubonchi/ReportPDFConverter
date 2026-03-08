# Report PDF Converter (Rust版)

Windows専用のZIP->Word->PDF結合ツールです。

## 機能

- ZIPアップロード
- ZIP内の `.doc` / `.docx` を Microsoft Word COM でPDF化
- 生成PDFを結合して1つのPDFを出力
- 進捗確認API
- プリンタ一覧取得と印刷API
- Gmail API（Python版の認証情報を流用可）またはSMTP経由のメール送信API（任意）

## 前提

- Windows
- Microsoft Word インストール済み
- Rust toolchain
- Visual Studio Build Tools (MSVC C++ ビルドツール, `link.exe`)

## 開発実行

```powershell
cargo run --release
```

デフォルトURL: `http://127.0.0.1:80`

環境変数で変更可能:

- `RPC_PORT` (既定: `80`)
- `RPC_BASE_DIR` (既定: Debug時はカレントディレクトリ、Release時はexe配置ディレクトリ)

## Gmail API送信（推奨）

Python版で作成済みの `gmail_api_credentials/credentials.json` と `token.json` をそのまま使えます。

- `GMAIL_CREDENTIALS_JSON` (任意)
- `GMAIL_TOKEN_JSON` (任意)
- `EMAIL_SENDER` (任意, 既定 `roboken.report.tool@gmail.com`)
- `EMAIL_DISPLAY_NAME` (任意, 既定 `ロボ研報告書作成ツール`)

`GMAIL_CREDENTIALS_JSON` / `GMAIL_TOKEN_JSON` 未指定時は以下順で探索します。

1. `Rust版/gmail_api_credentials`
2. 兄弟フォルダ `Python版/gmail_api_credentials`

## SMTP送信（任意・フォールバック）

メール送信API (`POST /api/email/:job_id`) を使う場合は以下を設定してください。

- `SMTP_HOST`
- `SMTP_PORT` (任意, 既定 `587`)
- `SMTP_USER`
- `SMTP_PASS`
- `SMTP_FROM` (任意)
- `SMTP_FROM_NAME` (任意)

`.env.sample` を `.env` に変更して値を設定できます。

## ビルド

```powershell
cargo build --release
```

生成物:

- `target/release/report-pdf-converter.exe`

## インストーラ作成（Inno Setup）

1. Inno Setup 6 をインストール
2. 以下を実行

```powershell
.\scripts\build_installer.ps1
```

出力先:

- `dist/installer/ReportPDFConverter-Setup.exe`

## API概要

- `POST /api/process` (multipart, field: `zip_file`)
- `GET /api/status/:job_id`
- `GET /download/:job_id`
- `GET /api/printers`
- `POST /api/print/:job_id`
- `POST /api/email/:job_id`

## 注意

- Word COM変換は環境依存のため、Wordのダイアログやアドイン状態で失敗することがあります。
- 印刷は既定PDFハンドラに依存します。
- ZIPエントリ名の特殊エンコードによっては解凍時に文字化けする場合があります。
