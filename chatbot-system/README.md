# AIチャットボットシステム

> Dialogflow CX + Vertex AI / Dify + LINE Bot によるマルチプラットフォーム対応

## 概要

長期インターン先のクライアント企業向けに開発した、AIチャットボットシステムです。
Webサイトへの埋め込みとLINE Bot連携の2つのチャネルに対応し、生成AI（Gemini / LLM）を活用して顧客からの問い合わせに24時間自動応答します。

---

## 1. Webサイト埋め込み型チャットボット

### 課題

- 営業時間外の問い合わせに対応できない
- 定型的な質問への回答に人的リソースがかかる
- FAQページだけでは顧客が必要な情報を見つけられない

### 解決策

Google Cloud の Dialogflow CX と Vertex AI（Gemini）を活用し、Webサイトに埋め込み可能なAIチャットボットを開発。ナレッジベースを参照して正確な回答を生成します。

### システム構成

```
┌─────────────────┐     ┌─────────────────┐     ┌─────────────────┐
│   Webサイト      │────▶│  Dialogflow CX  │────▶│   Vertex AI     │
│ (df-messenger)  │     │    (Agent)      │     │  (Gemini 2.5)   │
└─────────────────┘     └────────┬────────┘     └─────────────────┘
                                 │
                                 ▼
                        ┌─────────────────┐
                        │  Data Store     │
                        │ (ナレッジベース)  │
                        └─────────────────┘
```

### 主な機能

| 機能 | 説明 |
|------|------|
| Playbookによる回答生成 | 生成AIが自然な対話形式で回答 |
| ナレッジベース参照 | 企業情報・サービス内容を正確に回答 |
| df-messenger | Googleが提供するWebウィジェットで簡単埋め込み |
| トークン管理 | 会話履歴の長さに応じたトークン制限設定 |

### 技術スタック

- Google Dialogflow CX
- Vertex AI (Gemini 2.5 Flash)
- df-messenger（Webウィジェット）
- Data Store（ナレッジベース）

### 設定のポイント

| 項目 | 推奨設定 |
|------|---------|
| Model | gemini-2.5-flash |
| Input token limit | Up to 32k |
| Output token limit | Up to 2048 |
| Temperature | 0.3〜0.5（安定性重視） |

---

## 2. LINE Bot連携システム

### 課題

- LINE公式アカウントへの問い合わせに手動で対応していた
- 営業時間外は返信が遅れる
- 複数の問い合わせチャネルを一元管理したい

### 解決策

Dify（LLMアプリケーション開発プラットフォーム）とLINE Messaging APIを連携させ、LINEからの問い合わせにAIが自動応答するシステムを開発。

### システム構成

```
┌─────────────────┐     ┌─────────────────┐     ┌─────────────────┐
│    LINE App     │────▶│   Cloud Run     │────▶│    Dify API     │
│   (ユーザー)     │     │   (Express)     │     │   (LLM Agent)   │
└─────────────────┘     └────────┬────────┘     └─────────────────┘
                                 │
                                 ▼
                        ┌─────────────────┐
                        │ Secret Manager  │
                        │  (APIキー管理)   │
                        └─────────────────┘
```

### 主な機能

| 機能 | 説明 |
|------|------|
| LINE Webhook | LINEからのメッセージをリアルタイム受信 |
| Dify連携 | LLMによる自然な回答生成 |
| ストリーミング対応 | SSE形式のレスポンスをパース |
| 署名検証 | LINE署名によるセキュリティ確保 |
| Cloud Runデプロイ | サーバーレスでスケーラブルに運用 |

### 技術スタック

- Node.js (Express)
- LINE Messaging API (@line/bot-sdk)
- Dify API
- Google Cloud Run
- Google Cloud Secret Manager
- Docker

### APIエンドポイント

| エンドポイント | 用途 |
|---------------|------|
| `POST /api/chat` | Webフロントエンドからのチャットリクエスト |
| `POST /webhook` | LINE Webhookエンドポイント |

### セキュリティ対策

1. **LINE署名検証**: HMAC-SHA256による署名検証で不正リクエストを防止
2. **Secret Manager**: APIキー・トークンを環境変数ではなくSecret Managerで管理
3. **CORSポリシー**: 許可されたオリジンからのみアクセス可能

---

## デプロイ構成

### Google Cloud Run

```bash
# イメージビルド
gcloud builds submit --tag gcr.io/PROJECT_ID/chatbot-app

# デプロイ
gcloud run deploy chatbot \
  --image gcr.io/PROJECT_ID/chatbot-app \
  --platform managed \
  --region asia-northeast1 \
  --allow-unauthenticated \
  --set-secrets DIFY_API_KEY=DIFY_API_KEY:latest \
  --memory 512Mi \
  --min-instances 0 \
  --max-instances 10
```

### コスト最適化

- `--min-instances 0`: アイドル時はインスタンス数を0に
- `--max-instances 10`: 最大インスタンス数を制限
- `--memory 512Mi`: 必要最小限のメモリ割り当て

---

## 使用技術まとめ

| カテゴリ | 技術 |
|---------|------|
| AI/LLM | Vertex AI (Gemini), Dify |
| 対話基盤 | Google Dialogflow CX (Playbook) |
| バックエンド | Node.js, Express |
| クラウド | Google Cloud Run, Secret Manager |
| メッセージング | LINE Messaging API |
| コンテナ | Docker |
| フロントエンド | df-messenger |

---

## 工夫した点

1. **ハルシネーション対策**
   - ナレッジベースに情報がない場合は「その情報はございません」と正直に回答
   - プロンプトで回答範囲を明確に制限

2. **エラーハンドリング**
   - API通信エラー時のリトライ処理
   - ユーザーへのエラーメッセージ送信

3. **ストリーミングレスポンス対応**
   - Server-Sent Events (SSE) 形式のレスポンスをパース
   - 複数のイベントタイプ（agent_message, message）に対応

4. **マルチチャネル設計**
   - Web・LINE両方で同じLLMバックエンドを利用
   - チャネルごとに最適化されたUI/UX

5. **運用効率化**
   - Cloud Runによるサーバーレス運用
   - Secret Managerによる機密情報の安全な管理

---

## 学んだこと

- Google Cloud の AI/ML サービス（Dialogflow CX, Vertex AI）の活用
- LLMアプリケーション開発プラットフォーム（Dify）の利用
- LINE Bot開発とWebhook連携
- Cloud Runを使ったサーバーレスデプロイ
- ストリーミングAPI（SSE）のハンドリング
- 生成AIのプロンプト設計（ハルシネーション対策）
