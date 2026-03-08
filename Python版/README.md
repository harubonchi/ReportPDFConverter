# ReportPDFConverter

ZIP内のWordファイルをPDFに変換・結合し、必要に応じてGmail APIで送信するツールです。

## Directory Layout

```text
D:\Python\ReportPDFConverter\Python版
├─ python/
│  ├─ app.py
│  ├─ tray_launcher.py
│  ├─ email_service.py
│  ├─ pdf_merge.py
│  ├─ word_to_pdf_converter.py
│  └─ word_file_prefixer.py
├─ gmail_api_credentials/
│  ├─ credentials.json
│  └─ token.json
├─ lab_members/
│  ├─ order.json
│  └─ faculty_contacts.json
├─ templates/
├─ static/
├─ data/
├─ pyproject.toml
└─ report_pdf_converter.spec
```

## Prerequisites

- Windows
- Python 3.12+
- Microsoft Word (Word→PDF変換で使用)

## Setup

1. 依存関係をインストール

```powershell
uv sync
```

2. Gmail APIクライアント情報を配置

- `gmail_api_credentials/credentials.json` を配置
- `gmail_api_credentials/token.json` は初回認証後に自動作成

3. メンバー関連JSONを配置

- `lab_members/order.json`
- `lab_members/faculty_contacts.json`

## Run

1. Flaskアプリを直接起動

```powershell
python python/app.py
```

- 起動URL: `http://127.0.0.1:8000`

2. トレイランチャー起動（推奨）

```powershell
python python/tray_launcher.py
```

- 既定で `80` 番ポートを試行し、使用中の場合は空きポートへ切り替えます。

## Gmail Authentication Files

- `credentials.json`: Google Cloudで作成したOAuthクライアント情報
- `token.json`: 初回認証後に生成されるユーザーアクセストークン情報

環境変数で上書き可能です。

- `GMAIL_CREDENTIALS_JSON`
- `GMAIL_TOKEN_JSON`
- `EMAIL_SENDER`
- `EMAIL_DISPLAY_NAME`

## Build (PyInstaller)

```powershell
pyinstaller report_pdf_converter.spec
```

## Notes

- 機密情報保護のため、`gmail_api_credentials/credentials.json` と `gmail_api_credentials/token.json` は `.gitignore` に含めています。
