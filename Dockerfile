FROM python:3.10-slim

ENV DEBIAN_FRONTEND=noninteractive \
    PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1

# Java/Aspose を利用した変換処理は廃止し、Windows 環境での運用を前提としています。
# この Dockerfile はバックエンド API の検証用途のみを想定しています。
RUN set -eux; \
    apt-get update; \
    apt-get install -y --no-install-recommends \
        ca-certificates \
        fontconfig \
    ; \
    rm -rf /var/lib/apt/lists/*

WORKDIR /app

# ---- プロジェクト同梱フォントを配置（TTF/OTF/TTC なんでもOK）----
# 例: プロジェクト直下に fonts/Arial.ttf, fonts/MSGOTHIC.TTC, fonts/NotoSansCJK.ttc など
COPY fonts/ /usr/local/share/fonts/

# 任意: エイリアスと優先順位（必要なら調整）
RUN set -eux; \
    mkdir -p /etc/fonts; \
    printf '%s\n' \
'<?xml version="1.0"?>' \
'<!DOCTYPE fontconfig SYSTEM "fonts.dtd">' \
'<fontconfig>' \
'  <!-- generic family の優先順位（手元フォント → システム） -->' \
'  <alias><family>serif</family><prefer>' \
'    <family>Times New Roman</family>' \
'    <family>MS Mincho</family>' \
'  </prefer></alias>' \
'  <alias><family>sans-serif</family><prefer>' \
'    <family>MS Gothic</family>' \
'  </prefer></alias>' \
'' \
'  <!-- スキャン順でローカルを優先 -->' \
'  <dir>/usr/local/share/fonts</dir>' \
'  <dir>/usr/share/fonts</dir>' \
'</fontconfig>' > /etc/fonts/local.conf; \
    # フォントキャッシュ更新（配置後に実行するのが重要）
    fc-cache -f -v

# 依存パッケージ
COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt

# アプリ本体
COPY . .

EXPOSE 8000
CMD ["python", "app.py"]
