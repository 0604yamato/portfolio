# Python 3.11イメージを使用
FROM python:3.11-slim

# 作業ディレクトリを設定
WORKDIR /app

# 依存パッケージをコピーしてインストール
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# アプリケーションファイルをコピー
COPY main.py .
COPY article_generation_prompt.txt .

# ポート8080を公開
EXPOSE 8080

# アプリケーションを起動
CMD exec gunicorn --bind :$PORT --workers 1 --threads 8 --timeout 0 main:app
