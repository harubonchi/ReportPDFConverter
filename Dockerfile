FROM python:3.10-slim

WORKDIR /app

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1

# LibreOffice と日本語フォールバック(Noto)に加えて fontconfig を入れる
RUN apt-get update \
    && apt-get install -y --no-install-recommends \
        libreoffice \
        openjdk-17-jdk \
        fonts-noto-cjk \
        fontconfig \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

ENV JAVA_HOME=/usr/lib/jvm/java-17-openjdk-amd64 \
    PATH="${JAVA_HOME}/bin:${PATH}"

# ---- ローカル購入フォントを最優先で組み込み（MS Gothic / MS Mincho / Times New Roman）----
# プロジェクト直下に置いた TTF をコンテナのシステムフォントに配置
# ※ファイル名は質問どおり：MSGOTHIC.TTF / MSMINCHO.TTF / times-new-roman.ttf
COPY MSGOTHIC.TTF MSMINCHO.TTF times-new-roman.ttf /usr/local/share/fonts/truetype/local/

# fontconfig でエイリアスと優先順位を設定：
# 1) DOCX 内の表記ゆれ（全角名 / P系 / PostScript 名）をローカル実体に収束
# 2) generic family(serif/sans-serif) でもローカル3種を最優先
RUN set -eux; \
    mkdir -p /etc/fonts; \
    printf '%s\n' \
'<?xml version="1.0"?>' \
'<!DOCTYPE fontconfig SYSTEM "fonts.dtd">' \
'<fontconfig>' \
'  <!-- ==== Family aliases to converge name variants to local TTFs ==== -->' \
'  <!-- MS Gothic group -->' \
'  <alias><family>ＭＳ ゴシック</family><prefer><family>MS Gothic</family></prefer></alias>' \
'  <alias><family>MS PGothic</family><prefer><family>MS Gothic</family></prefer></alias>' \
'  <alias><family>MS Gothic</family><prefer><family>MS Gothic</family></prefer></alias>' \
'' \
'  <!-- MS Mincho group -->' \
'  <alias><family>ＭＳ 明朝</family><prefer><family>MS Mincho</family></prefer></alias>' \
'  <alias><family>MS PMincho</family><prefer><family>MS Mincho</family></prefer></alias>' \
'  <alias><family>MS Mincho</family><prefer><family>MS Mincho</family></prefer></alias>' \
'' \
'  <!-- Times New Roman variants (some DOCX embed PostScript names) -->' \
'  <alias><family>TimesNewRomanPSMT</family><prefer><family>Times New Roman</family></prefer></alias>' \
'  <alias><family>Times New Roman</family><prefer><family>Times New Roman</family></prefer></alias>' \
'' \
'  <!-- ==== Generic family preferences (prefer local TTFs, then Noto fallbacks) ==== -->' \
'  <alias><family>serif</family>' \
'    <prefer>' \
'      <family>Times New Roman</family>' \
'      <family>MS Mincho</family>' \
'      <family>Noto Serif CJK JP</family>' \
'    </prefer>' \
'  </alias>' \
'  <alias><family>sans-serif</family>' \
'    <prefer>' \
'      <family>MS Gothic</family>' \
'      <family>Noto Sans CJK JP</family>' \
'    </prefer>' \
'  </alias>' \
'  <!-- monospace は今回は明示せず（必要なら追加可） -->' \
'' \
'  <!-- ==== Prefer locally-installed fonts over system ones when families tie ==== -->' \
'  <!-- Boost local dir priority by ordering: prepend local dirs in scan order -->' \
'  <dir>/usr/local/share/fonts/truetype/local</dir>' \
'</fontconfig>' > /etc/fonts/local.conf; \
    # フォントキャッシュ更新
    fc-cache -f -v
# ---- ここまで ----

COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

EXPOSE 8000

CMD ["python", "app.py"]
