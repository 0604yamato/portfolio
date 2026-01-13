# 業務自動化スイート

Google Apps Scriptを使用した業務効率化ツール群です。

## プロジェクト構成

```
business-automation/
├── alert-system/           # 顧客アラート通知システム
│   ├── customer_notification.gs    # Chatwork通知
│   └── customer_master_setup.gs    # 顧客マスター管理
├── influencer-campaign/    # インフルエンサー施策管理
│   ├── project_management.js       # 案件管理
│   └── application_management.js   # 応募管理
└── quote-system/           # 見積書管理システム
    ├── config.gs           # 設定ファイル
    ├── sync_data.gs        # データ同期
    └── history.gs          # 履歴管理
```

## 機能概要

### 1. 顧客アラート通知システム
- 「次回やること」の日付を超過した顧客を自動検出
- Chatwork APIで担当者に自動通知
- 毎日定時実行（トリガー設定可能）

### 2. インフルエンサー施策管理
- 飲食店向けインフルエンサーマーケティング案件管理
- Web API（doGet/doPost）でフロントエンドと連携
- 案件登録 → 審査 → 公開 → 応募受付のワークフロー

### 3. 見積書管理システム
- 管理者・担当者間のデータ同期
- 見積書履歴の保存・復元機能
- カスタムメニューからの操作

## 使用技術

- **Google Apps Script** - サーバーレス実行環境
- **Google Sheets API** - データストレージ
- **Chatwork API** - チャット通知

## セットアップ

1. 新しいGoogleスプレッドシートを作成
2. 拡張機能 → Apps Script を開く
3. 各.gsファイルのコードを貼り付け
4. スクリプト内の定数（スプレッドシートID等）を設定
5. 初期設定関数を実行

## 環境変数

`.env.example` を参照してください。
Google Apps Scriptでは、スクリプト内のCONFIGオブジェクトを直接編集します。
