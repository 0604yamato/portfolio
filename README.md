# Portfolio - Yamato

工学院大学 情報学部 情報通信工学科 3年（2027年卒予定）

クラウドや生成AIを活用した**業務自動化・DX推進**に強い関心を持ち、
長期インターンでは実務レベルのシステム開発に取り組んでいます。

---

## Skills

### Languages
![Python](https://img.shields.io/badge/Python-3776AB?style=flat-square&logo=python&logoColor=white)
![JavaScript](https://img.shields.io/badge/JavaScript-F7DF1E?style=flat-square&logo=javascript&logoColor=black)
![Google Apps Script](https://img.shields.io/badge/Google%20Apps%20Script-4285F4?style=flat-square&logo=google&logoColor=white)

### Cloud / AI
![Google Cloud](https://img.shields.io/badge/Google%20Cloud-4285F4?style=flat-square&logo=googlecloud&logoColor=white)
![OpenAI](https://img.shields.io/badge/OpenAI-412991?style=flat-square&logo=openai&logoColor=white)
![Vertex AI](https://img.shields.io/badge/Vertex%20AI-4285F4?style=flat-square&logo=googlecloud&logoColor=white)

### Tools
![Git](https://img.shields.io/badge/Git-F05032?style=flat-square&logo=git&logoColor=white)
![Docker](https://img.shields.io/badge/Docker-2496ED?style=flat-square&logo=docker&logoColor=white)
![VS Code](https://img.shields.io/badge/VS%20Code-007ACC?style=flat-square&logo=visualstudiocode&logoColor=white)

---

## Project: SEO記事自動生成システム

> OpenAI GPT-4o + Google Cloud Run による業務自動化

### 概要

SEO記事の作成プロセスを自動化するシステムを開発しました。
キーワード分析から構成案作成、初稿生成までを一気通貫で自動化し、
クライアントのメディア運営を効率化しています。

### システム構成

```
┌─────────────────┐     ┌─────────────────┐     ┌─────────────────┐
│  Google Sheets  │────▶│   Cloud Run     │────▶│  Google Docs    │
│  (見出し構造)    │     │  (Python API)   │     │  (生成記事)      │
└─────────────────┘     └────────┬────────┘     └─────────────────┘
                                 │
                                 ▼
                        ┌─────────────────┐
                        │   OpenAI API    │
                        │   (GPT-4o)      │
                        └─────────────────┘
```

### 技術スタック

| カテゴリ | 技術 |
|---------|------|
| 言語 | Python 3.11 |
| AI/LLM | OpenAI API (GPT-4o) |
| クラウド | Google Cloud Run |
| API連携 | Google Sheets API, Google Docs API, Google Drive API, Search Console API |
| コンテナ | Docker |
| その他 | Google Apps Script |

### 主な機能

1. **SEO記事自動生成**: スプレッドシートの見出し構造から約6,000文字の記事を自動生成
2. **Search Consoleキーワード分析**: 検索キーワードデータを自動取得・分析
3. **ステータス管理**: 処理状況をスプレッドシートで自動更新
4. **Cloud Run対応**: HTTPエンドポイントとして公開、スケーラブルに運用

### 成果

| 指標 | Before | After |
|------|--------|-------|
| 1記事あたりの作成時間 | 3時間 | **30分**（80%削減） |
| 月間記事生成数 | - | **50記事以上** |

### 工夫した点

- **プロンプトエンジニアリング**: SEOに最適化された記事を安定生成するためのプロンプト設計
- **エラーハンドリング**: API制限への対応（リトライロジック）、処理中断時の再開機能
- **運用効率化**: Google Apps Scriptによるワンクリック実行、スプレッドシートでの進捗管理

---

## Project: 業務自動化システム

> Google Apps Script + Chatwork API による営業管理・業務効率化

### 概要

クライアント企業の営業管理業務を自動化するシステム群を開発しました。
顧客対応の漏れ防止・データ管理の効率化を実現しています。

### システム一覧

| システム名 | 機能 |
|-----------|------|
| 顧客対応アラート通知 | 対応期限超過の顧客をChatworkに自動通知 |
| 要対応リスト自動生成 | 対応が必要な顧客を専用シートに一覧表示 |
| 見積書管理・データ同期 | 複数スプレッドシート間のデータ同期と履歴管理 |

### 技術スタック

| カテゴリ | 技術 |
|---------|------|
| 言語 | Google Apps Script (JavaScript) |
| API連携 | Chatwork API, Google Sheets API |
| 自動化 | 時間ベーストリガー、イベントトリガー |

### 成果

- 営業担当者の確認作業を自動化
- 対応漏れの削減
- データ管理の一元化

[詳細はこちら](./business-automation/README.md)

---

## Internship Experience

| 期間 | 企業 | 業務内容 |
|------|------|----------|
| 2024年〜現在 | IT企業（長期インターン） | 生成AIを活用したDX推進、業務自動化システム開発 |

---

## Education

**工学院大学 情報学部 情報通信工学科**
2023年4月入学（2027年3月卒業予定）

---

## Contact

- **Email**: abudahe634@gmail.com
- **GitHub**: [@0604yamato](https://github.com/0604yamato)

---

*このポートフォリオは随時更新しています*
