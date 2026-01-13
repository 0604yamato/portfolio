# サンプルクリニックチャットボット - Cloud Runデプロイガイド

このガイドでは、チャットボットをGoogle Cloud Runにデプロイする手順を説明します。

## 前提条件

- Google Cloud Platform（GCP）アカウント
- Google Cloud CLI（gcloud）がインストール済み
- Dify APIキー
- （オプション）LINE Bot設定（LINE連携を使用する場合）

## 1. GCPプロジェクトの準備

```bash
# Google Cloud CLIにログイン
gcloud auth login

# プロジェクトを作成（既存のプロジェクトを使用する場合はスキップ）
gcloud projects create your-gcp-project --name="サンプルクリニックチャットボット"

# プロジェクトを設定
gcloud config set project your-gcp-project

# 必要なAPIを有効化
gcloud services enable cloudbuild.googleapis.com
gcloud services enable run.googleapis.com
gcloud services enable secretmanager.googleapis.com
```

## 2. 環境変数の設定（Secret Managerを使用）

本番環境では、Secret Managerを使って機密情報を安全に管理します。

```bash
# Dify APIキーを保存
echo -n "your-dify-api-key" | gcloud secrets create DIFY_API_KEY --data-file=-

# （オプション）LINE設定を保存
echo -n "your-line-channel-secret" | gcloud secrets create LINE_CHANNEL_SECRET --data-file=-
echo -n "your-line-channel-access-token" | gcloud secrets create LINE_CHANNEL_ACCESS_TOKEN --data-file=-

# Dify API URLを保存（デフォルト: https://api.dify.ai/v1/chat-messages）
echo -n "https://api.dify.ai/v1/chat-messages" | gcloud secrets create DIFY_URL --data-file=-
```

## 3. Dockerイメージのビルドとデプロイ

```bash
# dify-proxyディレクトリに移動
cd C:\Users\yamat\OneDrive\UC\dify-proxy

# Cloud Buildでイメージをビルド
gcloud builds submit --tag gcr.io/your-gcp-project/chatbot-app

# Cloud Runにデプロイ
gcloud run deploy your-gcp-project \
  --image gcr.io/your-gcp-project/chatbot-app \
  --platform managed \
  --region asia-northeast1 \
  --allow-unauthenticated \
  --set-secrets DIFY_API_KEY=DIFY_API_KEY:latest,DIFY_URL=DIFY_URL:latest,LINE_CHANNEL_SECRET=LINE_CHANNEL_SECRET:latest,LINE_CHANNEL_ACCESS_TOKEN=LINE_CHANNEL_ACCESS_TOKEN:latest \
  --memory 512Mi \
  --cpu 1 \
  --max-instances 10 \
  --min-instances 0
```

**注意**: LINE連携を使用しない場合は、`--set-secrets`からLINE関連の設定を削除してください：

```bash
gcloud run deploy your-gcp-project \
  --image gcr.io/your-gcp-project/chatbot-app \
  --platform managed \
  --region asia-northeast1 \
  --allow-unauthenticated \
  --set-secrets DIFY_API_KEY=DIFY_API_KEY:latest,DIFY_URL=DIFY_URL:latest \
  --memory 512Mi \
  --cpu 1 \
  --max-instances 10 \
  --min-instances 0
```

## 4. デプロイの確認

デプロイが完了すると、URLが表示されます：

```
Service [your-gcp-project] revision [your-gcp-project-00001-xyz] has been deployed and is serving 100 percent of traffic.
Service URL: https://your-gcp-project-xxxxx-an.a.run.app
```

ブラウザでこのURLにアクセスすると、チャットボットUIが表示されます。

## 5. カスタムドメインの設定（オプション）

独自ドメインを使用する場合：

```bash
# ドメインマッピングを作成
gcloud run domain-mappings create --service your-gcp-project --domain chat.example.com --region asia-northeast1
```

その後、DNSレコードを設定してドメインをCloud Runサービスに向けます。

## 6. LINE Webhook URLの設定（LINE連携を使用する場合）

LINE Developers Consoleで、Webhook URLを以下のように設定します：

```
https://your-gcp-project-xxxxx-an.a.run.app/webhook
```

## 7. 更新とデプロイ

コードを更新した後、再度デプロイするには：

```bash
# イメージを再ビルド
gcloud builds submit --tag gcr.io/your-gcp-project/chatbot-app

# 再デプロイ（前回と同じコマンド）
gcloud run deploy your-gcp-project \
  --image gcr.io/your-gcp-project/chatbot-app \
  --platform managed \
  --region asia-northeast1 \
  --allow-unauthenticated \
  --set-secrets DIFY_API_KEY=DIFY_API_KEY:latest,DIFY_URL=DIFY_URL:latest \
  --memory 512Mi \
  --cpu 1
```

## 8. ログの確認

```bash
# リアルタイムログを表示
gcloud run services logs read your-gcp-project --region asia-northeast1 --follow

# 最新のログを表示
gcloud run services logs read your-gcp-project --region asia-northeast1 --limit 50
```

## トラブルシューティング

### エラー: "Secret not found"

Secret Managerでシークレットが正しく作成されているか確認してください：

```bash
gcloud secrets list
```

### エラー: "Permission denied"

Cloud Runサービスアカウントに必要な権限を付与してください：

```bash
# サービスアカウントにSecret Accessorロールを付与
gcloud projects add-iam-policy-binding your-gcp-project \
  --member serviceAccount:PROJECT_NUMBER-compute@developer.gserviceaccount.com \
  --role roles/secretmanager.secretAccessor
```

### 環境変数の確認

```bash
# デプロイされたサービスの設定を確認
gcloud run services describe your-gcp-project --region asia-northeast1
```

## コスト最適化

- `--min-instances 0`: アイドル時にインスタンス数を0にしてコストを削減
- `--max-instances 10`: 最大インスタンス数を制限
- `--memory 512Mi`: 必要最小限のメモリを設定

## セキュリティ

- 本番環境では`--allow-unauthenticated`を削除し、認証を有効にすることを推奨
- APIキーやトークンは必ずSecret Managerで管理
- `.env`ファイルをGitにコミットしない

## サポート

問題が発生した場合は、ログを確認するか、Google Cloud Supportにお問い合わせください。
