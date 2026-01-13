"""
クライアント企業メディア記事自動生成 Cloud Run版

Cloud Runで動作するAPIサーバー
GASから呼び出されて、記事を生成する
"""

from flask import Flask, request, jsonify
import os
import logging
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from openai import OpenAI
import anthropic
import json
import random
import requests
from io import BytesIO
import re
import base64
from google.cloud import aiplatform
from PIL import Image
import google.generativeai as genai
from concurrent.futures import ThreadPoolExecutor, as_completed
import threading
import time
from bs4 import BeautifulSoup
from janome.tokenizer import Tokenizer
from collections import Counter
from google.cloud import tasks_v2
from google.protobuf import timestamp_pb2
import datetime

# ロギング設定
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = Flask(__name__)

# Google API のスコープ
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/documents',
    'https://www.googleapis.com/auth/drive',  # Drive全体へのアクセス（画像フォルダー読み取りに必要）
    'https://www.googleapis.com/auth/webmasters.readonly',  # Search Console（読み取り専用）
    'https://www.googleapis.com/auth/cloud-platform'  # Vertex AI（画像生成）に必要
]

def send_slack_notification(message, webhook_url=None):
    """Slackに通知を送信"""
    webhook_url = webhook_url or os.environ.get('SLACK_WEBHOOK_URL')
    if not webhook_url:
        logger.warning("SLACK_WEBHOOK_URL が設定されていません")
        return False

    try:
        response = requests.post(
            webhook_url,
            json={'text': message},
            headers={'Content-Type': 'application/json'}
        )
        if response.status_code == 200:
            logger.info("Slack通知を送信しました")
            return True
        else:
            logger.error(f"Slack通知エラー: {response.status_code} - {response.text}")
            return False
    except Exception as e:
        logger.error(f"Slack通知エラー: {e}")
        return False

class ArticleAutomation:
    def __init__(self, spreadsheet_id, openai_api_key, image_folder_id=None, project_id=None, image_generation_method='existing_folder',
                 master_spreadsheet_id=None, keyword_column='G', article_url_column='N', anthropic_api_key=None):
        self.spreadsheet_id = spreadsheet_id
        self.image_folder_id = image_folder_id
        self.project_id = project_id or os.environ.get('GCP_PROJECT_ID', 'YOUR_GCP_PROJECT_ID')
        self.image_generation_method = image_generation_method  # 'existing_folder', 'vertex_ai', or 'both'
        # マスターシート関連
        self.master_spreadsheet_id = master_spreadsheet_id  # 初稿URLを書き込むマスターシート
        self.keyword_column = keyword_column  # キーワード列（デフォルト: G）
        self.article_url_column = article_url_column  # 初稿URL列（デフォルト: N）
        # OpenAI クライアントを初期化
        self.openai_client = OpenAI(
            api_key=openai_api_key
        )
        # Claude クライアントを初期化（APIキーがあれば）
        self.anthropic_api_key = anthropic_api_key or os.environ.get('ANTHROPIC_API_KEY')
        self.claude_client = None
        if self.anthropic_api_key:
            self.claude_client = anthropic.Anthropic(api_key=self.anthropic_api_key)
        self.sheets_service = None
        self.docs_service = None
        self.drive_service = None
        self.image_cache = None  # サブフォルダーと画像のキャッシュ
        self.credentials = None  # Google認証情報を保存

    def authenticate_google(self):
        """サービスアカウントで認証"""
        # 環境変数からサービスアカウントキーを取得
        service_account_key = os.environ.get('GOOGLE_SERVICE_ACCOUNT_KEY')

        if not service_account_key:
            raise ValueError("GOOGLE_SERVICE_ACCOUNT_KEY environment variable is not set")

        logger.info(f"[DEBUG] 環境変数の最初の20文字: {service_account_key[:20]}")
        logger.info(f"[DEBUG] 環境変数の長さ: {len(service_account_key)}")

        try:
            # JSONとしてパース
            service_account_info = json.loads(service_account_key)
            logger.info(f"[DEBUG] JSON解析成功。client_email: {service_account_info.get('client_email', 'N/A')}")
        except json.JSONDecodeError as e:
            logger.error(f"JSONパースエラー: {e}")
            logger.error(f"環境変数の値（最初の100文字）: {service_account_key[:100]}")
            raise ValueError(f"Invalid JSON in GOOGLE_SERVICE_ACCOUNT_KEY: {e}")

        credentials = service_account.Credentials.from_service_account_info(
            service_account_info,
            scopes=SCOPES
        )

        # 認証情報を保存
        self.credentials = credentials

        self.sheets_service = build('sheets', 'v4', credentials=credentials)
        self.docs_service = build('docs', 'v1', credentials=credentials)
        self.drive_service = build('drive', 'v3', credentials=credentials)

        # Vertex AIの初期化（認証情報を使用）
        aiplatform.init(
            project=self.project_id,
            location='us-central1',
            credentials=credentials
        )

    def get_all_sheets(self):
        """すべてのシート名を取得"""
        try:
            logger.info(f"[DEBUG] スプレッドシートID: {self.spreadsheet_id}")
            logger.info(f"[DEBUG] Sheets Service: {self.sheets_service}")
            sheet_metadata = self.sheets_service.spreadsheets().get(
                spreadsheetId=self.spreadsheet_id
            ).execute()
            logger.info(f"[DEBUG] メタデータ取得成功")
            sheets = sheet_metadata.get('sheets', [])
            logger.info(f"[DEBUG] シート数: {len(sheets)}")
            sheet_titles = [sheet['properties']['title'] for sheet in sheets]
            logger.info(f"[DEBUG] シート名: {sheet_titles}")
            return sheet_titles
        except HttpError as err:
            logger.error(f"エラー: シート一覧の取得に失敗しました - {err}")
            logger.error(f"[DEBUG] エラー詳細: {err.content}")
            return []
        except Exception as e:
            logger.error(f"予期しないエラー: {e}")
            return []

    def get_headings_from_sheet(self, sheet_name, force=False):
        """シートから見出しデータを抽出（列ズレ対応版）

        Args:
            sheet_name: シート名
            force: Trueの場合、処理済みでも再生成する
        """
        try:
            range_name = f"'{sheet_name}'!A1:H100"  # H列まで広めに取得
            result = self.sheets_service.spreadsheets().values().get(
                spreadsheetId=self.spreadsheet_id,
                range=range_name
            ).execute()

            values = result.get('values', [])

            if not values:
                logger.info(f"[DEBUG] シート '{sheet_name}': データが空")
                return None

            if len(values) < 7:
                logger.info(f"[DEBUG] シート '{sheet_name}': 行数不足（{len(values)}行、最低7行必要）")
                return None

            # キーワード抽出（3行目付近を探索）
            keyword = ""
            if len(values) > 2:
                for cell in values[2]:
                    cell_text = cell.strip() if cell else ""
                    # ラベル（メインKW、キーワード等）をスキップ
                    if cell_text and not cell_text.startswith("メインKW") and not cell_text.startswith("キーワード"):
                        keyword = cell_text
                        break
            
            # ステータス確認（F2セル相当を探す）
            # 2行目の「処理済み」を含むセルを探す
            status = ""
            if len(values) > 1:
                for cell in values[1]:
                    if cell == "処理済み":
                        status = "処理済み"
                        break
            
            logger.info(f"[DEBUG] シート '{sheet_name}': ステータス='{status}', force={force}")
            if status == "処理済み" and not force:
                logger.info(f"[DEBUG] シート '{sheet_name}': 処理済みのためスキップ")
                return None
            elif status == "処理済み" and force:
                logger.info(f"[DEBUG] シート '{sheet_name}': 処理済みですが、force=Trueのため再生成します")

            # 見出し抽出（全行スキャン、行の中にH1-H4があるか探す）
            h1_title = ""
            headings = []

            # 2行目のタイトル案（H1）をまずチェック
            if len(values) > 1:
                row = values[1]
                # "タイトル案" の右側、または H1タグがあるか
                for i, cell in enumerate(row):
                    if not cell: continue
                    # H1タグがある場合、その右側をタイトルとする
                    if cell.strip() == "H1" and i + 1 < len(row):
                        h1_title = row[i+1].strip()
                        break
                    # 単に長い文字列があればタイトル候補とする（後でH1が見つからなければこれを使う）
                    if len(cell) > 10 and "タイトル" not in cell and not h1_title:
                         h1_title = cell.strip()

            # 7行目以降をスキャンして見出しを探す
            for i in range(6, len(values)):
                row = values[i]
                if not row: continue

                hierarchy = ""
                heading_text = ""

                # 行内で H1, H2... を探す
                for j, cell in enumerate(row):
                    if not cell: continue
                    val = cell.strip()
                    if val in ["H1", "H2", "H3", "H4"]:
                        hierarchy = val
                        # マーカーの右側にある最初の空でないセルを見出しテキストとする
                        for k in range(j + 1, len(row)):
                            if row[k] and row[k].strip():
                                heading_text = row[k].strip()
                                break
                        break
                
                if hierarchy and heading_text:
                    if hierarchy == "H1":
                        h1_title = heading_text
                    else:
                        headings.append({
                            'level': hierarchy,
                            'text': heading_text
                        })

            # H2の数をカウント
            h2_count = len([h for h in headings if h['level'] == 'H2'])

            logger.info(f"[DEBUG] シート '{sheet_name}': H1='{h1_title}', 見出し総数={len(headings)}, H2数={h2_count}")

            if h1_title and headings:
                return {
                    'keyword': keyword,
                    'h1_title': h1_title,
                    'headings': headings,
                    'sheet_name': sheet_name
                }
            
            # データが見つからなかった場合のデバッグログ
            if not headings:
                logger.info(f"[DEBUG] シート '{sheet_name}': 見出しが見つかりませんでした。行データをダンプします:")
                for i in range(min(15, len(values))):
                     logger.info(f"Row {i}: {values[i]}")

            return None

        except HttpError as err:
            logger.error(f"エラー: スプレッドシートの取得に失敗 - {err}")
            return None

    def _group_headings_by_h2(self, headings):
        """見出しをH2ごとにグループ化"""
        h2_groups = []
        current_group = None

        for heading in headings:
            if heading['level'] == 'H2':
                if current_group:
                    h2_groups.append(current_group)
                current_group = {
                    'h2': heading['text'],
                    'sub_headings': []
                }
            elif heading['level'] in ['H3', 'H4'] and current_group:
                current_group['sub_headings'].append(heading)

        if current_group:
            h2_groups.append(current_group)

        return h2_groups

    def _summarize_section(self, section_content, h2_text):
        """セクションの要約を作成（次のセクション生成時に使用）"""
        try:
            response = self.openai_client.chat.completions.create(
                model="gpt-4o-mini",
                messages=[
                    {"role": "system", "content": "あなたは記事要約の専門家です。簡潔に要約してください。"},
                    {"role": "user", "content": f"""以下のセクションを100字以内で要約してください。

【セクション】
{section_content}

【要約の指示】
- H2「{h2_text}」で書いた内容の要点のみ
- 次のセクションで重複を避けるための参考情報として使う
- 100字以内で簡潔に

要約のみを出力してください（前置きや説明は不要）。"""}
                ],
                max_completion_tokens=200
            )

            summary = response.choices[0].message.content.strip()
            logger.info(f"要約作成完了: {summary[:50]}...")
            return summary

        except Exception as e:
            logger.error(f"エラー: 要約の作成に失敗 - {e}")
            return ""

    def _generate_h2_section(self, keyword, h2_text, sub_headings, target_chars, section_index, total_sections, previous_summary=""):
        """1つのH2セクションを生成"""
        try:
            # サブ見出しをマークダウン形式に変換
            sub_headings_text = ""
            for sub in sub_headings:
                if sub['level'] == 'H3':
                    sub_headings_text += f"\n### {sub['text']}"
                elif sub['level'] == 'H4':
                    sub_headings_text += f"\n#### {sub['text']}"

            # 番号付きリストの判定
            numbered_list_keywords = ['ランキング', '選', 'TOP', 'おすすめ']
            use_numbered_list = any(kw in h2_text for kw in numbered_list_keywords)
            numbered_list_instruction = ""
            if use_numbered_list:
                numbered_list_instruction = "\n- 「ランキング」「〇〇選」が含まれているため、必ず「1. 2. 3.」の番号付きリストを使用してください"

            # 前セクションの要約指示
            previous_context = ""
            if previous_summary:
                previous_context = f"""

【前のセクションで書いた内容】
{previous_summary}

【注意事項】
- 上記と重複する内容は書かないでください
- 上記の流れを受けて、このセクションの話題に自然に進んでください
- 同じ具体例・数字・表現を繰り返さないでください"""

            logger.info(f"[{section_index}/{total_sections}] H2セクション生成中: {h2_text}（目標{target_chars}字）")

            response = self.openai_client.chat.completions.create(
                model="gpt-4o-mini",
                messages=[
                    {"role": "system", "content": f"""あなたはウェブマガジンの専門ライターです。

【このセクションの必須要件】
1) 文字数: {target_chars}字を必ず達成（{target_chars - 50}字未満は不合格）
2) 見出し: H2「{h2_text}」とそのサブ見出しのみを使用（追加・変更禁止）
3) 「働」の漢字のみ平仮名（はたらく等）。履歴書・自己PR・面接・職場・時給等は漢字のまま。
4) 太字マークダウン（**）: 使用しない。強調は「」（鉤括弧）のみ{numbered_list_instruction}

【書き方の指示】
- 段落形式の文章を中心に構成すること
- 具体例・体験談・数値データを3～5個含めること
- 読者が具体的にイメージできる詳細な説明をすること
- 抽象的な説明ではなく、具体的なシーン・状況・数字を含める
- 文字数が不足する場合は、さらに具体例を追加する

【箇条書きの使用】
以下のような内容がある場合は、積極的に箇条書きを使用してください:
  * メリット・デメリットの比較 → 箇条書きで列挙
  * 手順やステップの説明 → 箇条書きで列挙
  * 選択肢やランキングの列挙 → 箇条書きで列挙
  * 注意点やポイントのまとめ → 箇条書きで列挙
  * チェックリスト → 箇条書きで列挙

【表の使用】
以下のような内容がある場合は、積極的に表を使用してください:
  * 時給・料金の比較 → 比較表を作成
  * プランや特徴の比較 → 比較表を作成
  * メリット・デメリット → 表で整理
  * 分類・カテゴリ整理 → 分類表を作成
  * 手順・チェックリスト → 表で整理

※ 記事全体で箇条書き3～4箇所、表2～3つが必要です。このセクションでも可能な限り使用してください（ただし使いすぎないこと）"""},
                    {"role": "user", "content": f"""キーワード「{keyword}」に関するセクションを書いてください。
{previous_context}

【見出し構造】
## {h2_text}{sub_headings_text}

【執筆方法】
1. H2冒頭に100～150字の導入文を書く
2. 各H3/H4セクションを詳しく書く（具体例を含む）
3. 文字数目標: {target_chars}字（この文字数を必ず達成してください）
4. 段落形式の文章をメインに、補助的に箇条書き・表を使う

【重要】文字数を確実に達成するために:
- 各セクションに具体例を3～5個含める
- 「例えば、〜」「実際に〜」「具体的には〜」などで詳しく説明
- 読者が具体的にイメージできる詳細な描写を心がける
- {target_chars}字に満たない場合は不合格です

セクション本文のみを出力してください（H1見出しは不要、H2見出しから開始）。
※文字数カウント（例：「文字数：○○字」）は絶対に出力しないこと。"""}
                ],
                max_completion_tokens=2000
            )

            section_content = response.choices[0].message.content
            section_char_count = len(section_content)
            logger.info(f"[{section_index}/{total_sections}] セクション生成完了: {section_char_count}字")

            return section_content

        except Exception as e:
            logger.error(f"エラー: H2セクション「{h2_text}」の生成に失敗 - {e}")
            return f"## {h2_text}\n\nERROR: {str(e)}"

    def generate_design(self, keyword, h1_title, headings):
        """Step0: 全体設計を生成（本文は書かない）"""
        try:
            logger.info(f"[Step0] 全体設計を生成中...")

            # 見出し構造をMarkdown形式で文字列化
            headings_md = self._format_headings_md(headings)

            response = self.openai_client.chat.completions.create(
                model="gpt-4o-mini",
                messages=[
                    {"role": "system", "content": """あなたは「読者目線」を最重視するSEO編集者。本文はまだ書かない。
以下のH2/H3見出しは確定。変更・追加・削除は禁止。

★最重要★ 読者は「専門家」ではなく「初めて調べる一般人」。難しい言葉は必ず説明が必要。

作るもの:

## 1. 読者の悩み・疑問（3〜5個）
このキーワードで検索する人が「本当に知りたいこと」を具体的に書く
例：「いくらもらえるの？」「手続きは難しい？」「損しない方法は？」

## 2. 専門用語リスト（初出時の説明文付き）
記事に出てくる専門用語を洗い出し、（）内に入れる説明文を用意
例：第3号被保険者（会社員の配偶者として扶養される人）

## 3. 各H2の設計（H2ごとに以下を書く）
- このH2で解決する読者の疑問：「○○って何？」「○○はどうすれば？」
- 冒頭の要点（1〜2文）：読者が最初に知りたい情報
- 共感ポイント：「〜と不安な方も多いですよね」等
- 末尾の呼びかけ：「まずは○○を確認してみましょう」等

## 4. 記事全体のポイント3点
読者がこの記事から持ち帰る「行動できる」ポイント

## 5. NG事項
- 説明なしの専門用語
- 長すぎる文（40字超）
- 呼びかけなしで終わるH2

出力：Markdownのみ"""},
                    {"role": "user", "content": f"""# テーマ
{keyword}

# H1見出し
{h1_title}

# 見出し
{headings_md}

# 参考メモ
- 読者：バイト・副業・生活情報を探す20～40代（専門知識なし）
- 読者の気持ち：「難しそう」「よくわからない」「損したくない」
- 目標：読者が「なるほど、これならできそう」と思える記事

# 表記ルール
- 「働」のみ平仮名（はたらく等）。共働き→共ばたらき
- 履歴書・自己PR・面接・職場・時給等は漢字のまま
- 太字マークダウン（**）は使用禁止。強調は「」（鉤括弧）のみ。"""}
                ],
                max_completion_tokens=3000
            )

            design_md = response.choices[0].message.content
            usage = response.usage
            logger.info(f"[Step0] 設計完了（tokens: {usage.total_tokens}）")

            return design_md, usage

        except Exception as e:
            logger.error(f"エラー: 設計の生成に失敗 - {e}")
            return f"ERROR: {str(e)}", None

    def generate_draft(self, keyword, h1_title, headings, design_md):
        """Step1: 初稿生成（設計に従って本文を書く）"""
        try:
            logger.info(f"[Step1] 初稿を生成中...")

            # 見出し構造をMarkdown形式で文字列化
            headings_md = self._format_headings_md(headings)

            # H2の数をカウントして各セクションの目標文字数を計算
            h2_count = len([h for h in headings if h['level'] == 'H2'])
            target_total_body = 4500
            target_per_section = target_total_body // max(1, h2_count)
            min_target = target_per_section - 100
            max_target = target_per_section + 100
            target_str = f"{min_target}～{max_target}"

            response = self.openai_client.chat.completions.create(
                model="gpt-5.2",
                messages=[
                    {"role": "system", "content": f"""あなたはクライアント企業メディアの記事ライターです。

# ■■■ 絶対厳守ルール（違反は不合格） ■■■

## 【ルール1】見出しは全て使用（省略厳禁）
提供された見出しを「全て」「そのまま」使用すること。
- H2が{h2_count}個あれば、記事にもH2を{h2_count}個全て書く
- 見出しを1つでも省略したら不合格
- 見出しを1文字でも変更したら不合格
- 見出しの順序を変更したら不合格
- 見出しを追加したら不合格

## 【ルール2】文字数は5,000〜6,000字（厳守）★最重要★
- 記事全体: 5,000〜6,000文字（絶対厳守）
- 理想は5,500字（中央値を狙う）
- 各H2セクション: {target_str}文字
- 導入部: 400〜500文字
- ★5,000字未満は不合格（具体例や説明が不足）★
- ★6,000字超過も不合格（冗長すぎる）★
- 両方の境界を意識して5,500字前後を目指す

## 【ルール3】禁止事項
- 「働」の漢字は全て平仮名に変換（例外なし）
  - 働く → はたらく
  - 働き方 → はたらき方
  - 共働き → 共ばたらき ★これも必ず平仮名★
  - 働いて → はたらいて
  ※他の漢字・カタカナは絶対にそのまま使用：
  - ✅ 履歴書（りれきしょ ❌）
  - ✅ 自己PR（じこぴーあーる ❌）
  - ✅ 面接（めんせつ ❌）
  - ✅ 職場（しょくば ❌）
  - ✅ 時給（じきゅう ❌）
  - ✅ 電気代、エアコン等もそのまま
- 太字（**）→ 使用禁止
- 乱暴な口調・煽り・過度なスラング
- 根拠のない断定（「必ず」「絶対」）や誇大表現の連発

# 文体・口調（クライアント企業メディア）

## 基本トーン
- です・ます調。丁寧だが堅すぎない（読みやすいWeb記事の口調）
- 読者に寄り添う言い回しを多用（例：「〜不安に思う方も多いのではないでしょうか？」）
- 押しつけや断定を避ける：不確かな場合は「〜とされています」「〜の傾向があります」「目安です」
- 読者の行動を促す締め：「まずは〜をチェックしてみてください」「ぜひ参考にしてみてください」

## 語彙・言い回し
- わかりやすさ重視：「ここでは〜をまとめて解説します」「ポイントは〜」「〜の目安」「〜しやすい」
- 安心・注意喚起：「事前に確認」「〜すると安心」「必ずチェック」「ミスマッチを防ぐ」
- ブランド文脈では「はたらく／はたらき方」「しごと」など、ひらがな表記を自然に混ぜる

## 文章リズム
- 1文は40〜60字以内（長い文は2文に分ける）
- 要点→理由→具体例→補足の順で書く
- 見出し直下に要点を簡潔に述べる
- 「！」を適度に使用する（記事全体で5〜10個程度）
  - 例：「ぜひ参考にしてみてください！」「おすすめです！」
  - 使いすぎに注意。自然な文脈でのみ使用

## 【ルール5】読みやすさ（最重要）★これで記事の質が決まる★

### ■ 文章の基本ルール
- 1文は40〜60字以内（それ以上は2文に分ける）
- 1段落は3〜4行以内（長くなったら改行）
- Web記事は「短く区切る」が基本

### ■ ビフォーアフター例（この違いを意識して書く）

【悪い例❌】
「第3号被保険者に該当するかどうかで手続きや将来の受給額が大きく変わりますので、まずは自分がどの区分に該当するのかを確認しておくことが重要です。」
→ 1文が長すぎる、専門用語の説明なし、堅い

【良い例✅】
「第3号被保険者（会社員の配偶者として扶養される人）かどうかで、手続きが変わります。受給額にも影響するので、まずは自分の区分を確認してみましょう。」
→ 2文に分割、専門用語を説明、呼びかけで締める

【悪い例❌】
「年金の受給開始年齢は原則65歳からとなっています。」
→ 堅い、説明だけで終わっている

【良い例✅】
「年金は原則65歳から受け取れます。「まだ先の話」と思うかもしれませんが、「いつから」「いくらもらえるか」を知っておくと安心ですよ。」
→ 親しみやすい、読者に寄り添っている

### ■ 各セクションの構成テンプレート（必ず守る）

【H2セクションの構成】
1. 冒頭1〜2文：要点を簡潔に述べる
2. 本文：理由→具体例→補足の順で展開
3. 末尾：自然な形で締める（毎回同じ呼びかけパターンは禁止）

【H3セクションの構成】
1. 冒頭：そのH3で伝えたいポイント
2. 本文：具体的な説明・例
3. 必要に応じて呼びかけ

### ■ 専門用語は必ず説明（初出時）
- 「第3号被保険者（会社員の配偶者として扶養される人）」
- 「老齢基礎年金（国民年金から支給される年金）」
- 「繰下げ受給（受給開始を遅らせて月額を増やす制度）」
- 読者は専門家ではない。説明なしの専門用語は不合格。

### ■ 具体的な数字を入れる
- 「年金額は人による」→「満額で年間約81万円（月約6.8万円）」
- 「多くの人が」→「約○○%の人が」「○○万人以上が」
- 数字があると信頼性・説得力が上がる

### ■ 共感表現（自然な文脈でのみ使用）
- 読者に寄り添う表現は「自然な文脈」でのみ使用
- 同じ共感フレーズを記事内で2回以上使わない
- テンプレート的な表現の連発は禁止
- 例：「〜と不安に思う方もいるかもしれません」「実は〜」などは自然な箇所でのみ

### ■ H2セクションの締め方
- 毎回同じパターンの呼びかけは禁止
- 自然な文脈で締める（無理に呼びかけなくてOK）
- 「〜してみましょう」「〜してみてください」の連発禁止
- H2ごとに異なる締め方をする（バリエーションを持たせる）

### ■ 興味を引く表現（適度に使う）
- 「実は〜」「意外にも〜」
- 「知っておきたいのが〜」
- 「ポイントは〜」「コツは〜」
- 「よくある間違いが〜」
- 「見落としがちなのが〜」

## 【ルール4】読者目線の言葉選び（重要）

### ■ 読者は「初心者」「未経験者」が多い
- 読者は「これからバイトを始める人」「初めて挑戦する人」
- 企業側・採用側の視点ではなく、読者（応募者）の視点で書く

### ■ 言葉選びの例
- ❌「新規採用のバイト」→ ✅「未経験の人」「初心者」「バイト初日は〜」「最初は誰でも〜」
- ❌「おすすめ時間」→ ✅「おすすめの時間」（自然な日本語に）
- ❌「勤務者」「従業員」→ ✅「はたらく人」「スタッフ」
- 読み手が違和感を感じない自然な表現を選ぶ

### ■ 当たり前すぎる文章を避ける
- 誰でも知っている情報だけの文章は価値がない
- ❌「バイトは仕事です」「お金をもらえます」「働くと疲れます」
- ✅ 具体的な数字、体験談、コツ、注意点など「知って得する情報」を入れる
- 「読者がこれを読んで何を得られるか？」を常に意識

## 【ルール5】見出しと内容の整合性（必須）

### ■ 見出しで約束した内容を必ず書く
- 見出しに「1日の流れ」があれば → 時系列で流れを説明する
- 見出しに「〇〇選」があれば → その数だけ紹介する
- 見出しに「比較」があれば → 比較表や比較説明を入れる
- 見出しと本文がズレていると読者は混乱する

### ■ 見出しの内容を網羅する
- 見出し「仕事内容・向いてる人・1日の流れ」なら、3つ全て説明する
- 1つでも抜けていたら不合格

## 【ルール6】簡潔→補足の構造（重要）

### ■ 冒頭で要点を述べる
- 最初の1〜2文で「このセクションで伝えたいこと」を簡潔に述べる
- その後に補足・具体例・理由を追加
- ❌ 冒頭から長々と背景説明をしない

### ■ 悪い例と良い例

【悪い例❌】冒頭が長い
「バイトを始めるにあたって、時給がどのくらいなのか気になる方も多いのではないでしょうか。時給は地域や職種によって異なりますが、一般的には...」

【良い例✅】簡潔に要点から
「時給の相場は900〜1,200円が目安です。地域や職種で差がありますが、都市部は高め、地方はやや低めの傾向があります。」

### ■ 改行のタイミング
- 話題が変わるタイミングで改行
- 長い説明が続いたら改行で区切る
- 「読みやすさ」を最優先

## 【ルール7】参照元・出典の記載

### ■ 統計データ・公的情報は出典を明記
- 「厚生労働省によると〜」「○○調査（2024年）では〜」
- 「○○省のデータでは〜」「○○協会の発表によれば〜」

### ■ 出典が必要な情報
- 統計データ（○○%、○○万人など）
- 法律・制度に関する情報
- 公的機関の発表
- 信頼性を高めるための根拠

### ■ 出典不要な情報
- 一般的な知識・常識
- 自明な事実
- 体験談・感想

## 【ルール8】表と箇条書き（必須）
- 表を記事全体で2〜3個必ず含める（1個以下は不合格）
- 箇条書きを記事全体で2〜3個必ず含める（まとめセクションには入れない）

【表の書き方（Markdown形式）】★必ずMarkdown形式で書くこと★
表は必ず以下のMarkdown形式で書くこと:

| 項目 | 内容A | 内容B |
|------|-------|-------|
| 特徴1 | 説明 | 説明 |
| 特徴2 | 説明 | 説明 |

★ 各行は | で始まり | で終わること
★ 2行目は区切り行（|------|------|）を必ず入れること
表の種類例: 比較表、時給一覧、メリット・デメリット表、手順表など

【箇条書きの書き方】
- ポイント1の説明文
- ポイント2の説明文
- ポイント3の説明文

箇条書きの種類例: ポイント、注意点、手順、チェックリストなど

## 【ルール9】AI感の排除（最重要）

### ■ リード文の書き方（厳守）
- ❌「この記事では」「本記事では」で始めない
- ❌「〜を解説します」「〜を紹介します」で始めない
- ✅ 読者の悩みや状況から書き始める
- ✅ 例：「共ばたらきで家事の分担に悩んでいませんか？」「家事分担がうまくいかない...」
- リード文は2〜3文で簡潔に。記事の価値を端的に伝える

### ■ 同語尾・同型文の禁止
- 同じ語尾（〜です/〜ます/〜しょう）が3回以上連続禁止
- 同じ文型（AはBです。CはDです。）が3回以上連続禁止
- 短文（20字以下）と中文（40〜60字）を混ぜて「ゆらぎ」を作る

### ■ 語尾のバリエーション（必須）
以下の語尾をバランスよく使い分ける：
- 〜です / 〜ます（基本・使いすぎ注意）
- 〜でしょう / 〜かもしれません（推測）
- 〜できます / 〜なります（可能・変化）
- 〜ください / 〜ましょう（呼びかけ）※記事内で各2回まで
- 〜ものです / 〜ことがあります（一般論）

### ■ 抽象語の禁止（具体に置換）
以下の抽象語は使用禁止。数字・判断基準・手順・具体例に置き換える：
- ❌「重要です」→ ✅「〜すると○○円節約できます」
- ❌「〜と言えます」→ ✅「〜という傾向があります」
- ❌「総じて」「〜が求められます」→ ✅ 具体的な条件や数字で説明

### ■ AIっぽい定型句の禁止
以下の表現は使用禁止（1回も使わない）：
- ❌「本記事では」「この記事では」「ここでは」
- ❌「結論から言うと」「結論としては」
- ❌「つまり」「すなわち」の多用
- ❌「〜を解説します」「〜を紹介します」「〜を説明します」（完全禁止）
- ❌「〜についてみていきましょう」「〜を見ていきます」
- ❌「いかがでしたか？」「参考になれば幸いです」
- ❌「〜してみてはいかがでしょうか」
- ❌「〜が挙げられます」「〜が考えられます」の連発

### ■ 前置き・水増しの禁止
- 冒頭から本題に入る（長い前置き禁止）
- 「まず〜」「はじめに〜」で始めすぎない
- 網羅で水増ししない（読者の判断/行動に効く情報のみ）
- 当たり前すぎる説明は省略

### ■ 根拠のない断定の禁止
- 与えた素材以外の情報は断定禁止
- 素材外は「〜の傾向があります」「〜とされています」で推測扱い
- 統計データは必ず出典を明記

# 書式
H1: #、H2: ##、H3: ###、H4: ####
箇条書き: -
表: Markdown形式（| 項目 | 内容 | で書く）"""},
                    {"role": "user", "content": f"""# ■ 記事執筆依頼 ■

キーワード: {keyword}

## H1見出し
# {h1_title}

## 見出し構造（以下{h2_count}個のH2を全て使用すること）
{headings_md}

## 参考: 全体設計
{design_md}

---

# ■■■ 重要な指示 ■■■

1. 上記の見出し構造に含まれる見出しを「全て」使ってください
   - H2が{h2_count}個あります。{h2_count}個全て書いてください
   - 見出しは1文字も変えないでください
   - 見出しを省略しないでください

2. ★文字数は5,000〜6,000文字で書いてください★（最重要）
   - 理想は5,500字前後（中央値を狙う）
   - 各H2セクション: {target_str}文字
   - ★5,000字未満は不合格（内容不足）★
   - ★6,000字超過も不合格（冗長すぎ）★

3. 表と箇条書き（必須）
   - 表を2〜3個必ず入れてください（Markdown形式）
   - ★必ずMarkdown形式の表を使うこと★
   - 箇条書きを2〜3個必ず入れてください（まとめセクション以外に）

   表の例（この形式を厳守）:
   | 項目 | 内容A | 内容B |
   |------|-------|-------|
   | 特徴 | 説明 | 説明 |
   ↑ 各行は | で始まり | で終わること！

   箇条書きの例:
   - ポイント1
   - ポイント2
   - ポイント3

4. 禁止事項
   - 「働」の漢字は全て平仮名に（共働き→共ばたらき も含む）
   - ※履歴書、自己PR、面接、職場、時給、電気代、エアコン等はそのまま使用（ひらがなNG）
   - 太字（**） → 使用禁止

5. 読みやすさ（最重要）★これで記事の質が決まる★

   【文章ルール】
   - 1文は40〜60字以内（長い文は2文に分ける）
   - 1段落は3〜4行以内
   - 話題が変わるタイミングで改行

   【簡潔→補足の構造】★重要★
   - 冒頭1〜2文で要点を述べる
   - その後に補足・具体例を追加
   - ❌ 冒頭から長々と説明しない

   【専門用語】
   - 初出時に必ず（）で説明を入れる
   - 説明なしの専門用語は不合格

   【具体的な数字・出典】
   - 「約○○円」「○○%」「○○万人」など数字で説得力を出す
   - 統計データは出典を明記：「厚生労働省によると〜」「○○調査（2024年）では〜」

   【H2セクションの構成】★必ず守る★
   - 冒頭：要点を1〜2文で簡潔に
   - 本文：理由→具体例→補足
   - 末尾：自然な形で締める（毎回同じ呼びかけパターンは禁止）

   【親しみやすさ】
   - 読者に寄り添う表現は自然な文脈でのみ使用
   - 同じ共感フレーズを記事内で2回以上使わない
   - 堅い説明文ではなく、読者に語りかける文体で

6. 読者目線の言葉選び（重要）
   - 読者は「初心者」「未経験者」が多い
   - ❌「新規採用のバイト」→ ✅「未経験の人」「初心者」「バイト初日は〜」
   - ❌「おすすめ時間」→ ✅「おすすめの時間」（自然な日本語に）
   - 企業側ではなく、読者（応募者）の視点で書く

7. 当たり前すぎる文章を避ける
   - ❌「バイトは仕事です」「お金をもらえます」
   - ✅ 具体的な数字、体験談、コツ、注意点を入れる
   - 「読者がこれを読んで何を得られるか？」を意識

8. 見出しと内容の整合性（必須）
   - 見出しで約束した内容を必ず本文で書く
   - 「1日の流れ」があれば → 時系列で説明
   - 「仕事内容・向いてる人・1日の流れ」なら → 3つ全て説明（抜け禁止）

記事本文のみを出力してください。

【絶対禁止】以下は出力に含めないこと：
- 文字数カウント（例：「文字数：5,000字」）
- チェック結果やコメント
- 説明文や前置き
- 「以上です」等の締めの言葉"""}
                ],
                max_completion_tokens=8000
            )

            draft_md = response.choices[0].message.content
            usage = response.usage
            char_count = len(draft_md)
            logger.info(f"[Step1] 初稿完了（文字数: {char_count}字、tokens: {usage.total_tokens}）")

            return draft_md, usage

        except Exception as e:
            logger.error(f"エラー: 初稿の生成に失敗 - {e}")
            return f"ERROR: {str(e)}", None

    def generate_draft_with_claude(self, keyword, h1_title, headings):
        """Claude APIを使用して初稿を生成（性能テスト用・最小フロー）"""
        try:
            if not self.claude_client:
                return "ERROR: Claude APIキーが設定されていません", None

            logger.info(f"[Claude] 初稿を生成中...")

            # 見出し構造をMarkdown形式で文字列化
            headings_md = self._format_headings_md(headings)

            # H2の数をカウントして各セクションの目標文字数を計算
            h2_count = len([h for h in headings if h['level'] == 'H2'])
            target_total_body = 4500
            target_per_section = target_total_body // max(1, h2_count)
            min_target = target_per_section - 100
            max_target = target_per_section + 100
            target_str = f"{min_target}～{max_target}"

            system_prompt = f"""あなたはクライアント企業メディアの記事ライターです。

# 絶対厳守ルール

## 【ルール1】見出しは全て使用（省略厳禁）
提供された見出しを「全て」「そのまま」使用すること。
- H2が{h2_count}個あれば、記事にもH2を{h2_count}個全て書く
- 見出しを1つでも省略したら不合格
- 見出しを1文字でも変更したら不合格

## 【ルール2】文字数は5,000〜6,000字（厳守）
- 記事全体: 5,000〜6,000文字
- 理想は5,500字
- 各H2セクション: {target_str}文字
- 導入部: 400〜500文字

## 【ルール3】禁止事項
- 「働」の漢字は全て平仮名に変換（働く→はたらく、共働き→共ばたらき）
- 太字（**）→ 使用禁止

## 【ルール4】文体・口調
- です・ます調。丁寧だが堅すぎない
- 読者に寄り添う言い回しを使用
- 1文は40〜60字以内
- 専門用語は初出時に（）で説明

## 【ルール5】必須要素
- 表を2〜3個必ず含める（Markdown形式）
- 箇条書きを2〜3個必ず含める
- 各H2セクションは自然な形で締める（毎回同じパターン禁止）

## 書式
H1: #、H2: ##、H3: ###、H4: ####
箇条書き: -
表: Markdown形式（| 項目 | 内容 |）"""

            user_prompt = f"""# 記事執筆依頼

キーワード: {keyword}

## H1見出し
# {h1_title}

## 見出し構造（以下{h2_count}個のH2を全て使用すること）
{headings_md}

---

# 重要な指示

1. 上記の見出し構造に含まれる見出しを「全て」使ってください
   - H2が{h2_count}個あります。{h2_count}個全て書いてください
   - 見出しは1文字も変えないでください

2. 文字数は5,000〜6,000文字で書いてください
   - 理想は5,500字前後
   - 各H2セクション: {target_str}文字

3. 表と箇条書き（必須）
   - 表を2〜3個必ず入れてください（Markdown形式）
   - 箇条書きを2〜3個必ず入れてください

記事本文のみを出力してください。説明や前置きは不要です。"""

            response = self.claude_client.messages.create(
                model="claude-sonnet-4-20250514",
                max_tokens=8000,
                messages=[
                    {"role": "user", "content": user_prompt}
                ],
                system=system_prompt
            )

            draft_md = response.content[0].text
            char_count = len(draft_md)

            # 使用量情報を作成
            usage_info = {
                'input_tokens': response.usage.input_tokens,
                'output_tokens': response.usage.output_tokens,
                'total_tokens': response.usage.input_tokens + response.usage.output_tokens
            }

            logger.info(f"[Claude] 初稿完了（文字数: {char_count}字、tokens: {usage_info['total_tokens']}）")

            return draft_md, usage_info

        except Exception as e:
            logger.error(f"エラー: Claude初稿の生成に失敗 - {e}")
            return f"ERROR: {str(e)}", None

    def generate_article(self, keyword, h1_title, headings):
        """記事生成のメインフロー（Step0〜Step3）"""
        try:
            usage_log = {}

            # Step0: 設計
            design_md, usage0 = self.generate_design(keyword, h1_title, headings)
            if "ERROR:" in design_md:
                return design_md
            usage_log['step0'] = usage0

            # Step1: 初稿
            draft_md, usage1 = self.generate_draft(keyword, h1_title, headings, design_md)
            if "ERROR:" in draft_md:
                return draft_md
            usage_log['step1'] = usage1

            # Step2: 監査
            issues_json, usage2 = self.audit_draft(design_md, draft_md)
            if "ERROR:" in issues_json:
                logger.warning(f"[WARNING] 監査に失敗。初稿をそのまま使用します: {issues_json}")
                return draft_md
            usage_log['step2'] = usage2

            # Step3: 修正
            final_md, usage3 = self.refine_draft(draft_md, issues_json)
            if "ERROR:" in final_md:
                logger.warning(f"[WARNING] 修正に失敗。初稿をそのまま使用します: {final_md}")
                return draft_md
            usage_log['step3'] = usage3

            # 文字数チェック（Google Docs保存時に約5%減る可能性があるため、余裕を持たせる）
            char_count = len(final_md)
            target_min = 5000  # 目標5500字
            target_max = 6000

            logger.info(f"[最終] 文字数: {char_count}字（目標: {target_min}～{target_max}字）")

            # 文字数追記（必要時のみ、最大3回まで試行）
            append_attempt = 0
            max_append_attempts = 3
            while char_count < target_min and append_attempt < max_append_attempts:
                append_attempt += 1
                missing_chars = target_min - char_count
                logger.warning(f"[WARNING] 文字数不足（-{missing_chars}字）。追記を実行します...（試行 {append_attempt}/{max_append_attempts}）")
                final_md, usage4 = self.append_content_if_needed(final_md, issues_json, missing_chars)
                usage_log[f'step4_{append_attempt}'] = usage4
                char_count = len(final_md)
                logger.info(f"[追記後 {append_attempt}回目] 文字数: {char_count}字")

                # 追記しても文字数が増えなかった場合は終了
                if usage4 is None:
                    logger.warning(f"[WARNING] 追記に失敗しました。処理を終了します。")
                    break

            if char_count < target_min:
                logger.warning(f"[WARNING] {max_append_attempts}回追記しても目標文字数に達しませんでした（{char_count}字）")

            # トークン使用量をログ出力
            total_tokens = sum([u.total_tokens for u in usage_log.values() if u])
            logger.info(f"[トークン使用量] 合計: {total_tokens} tokens")
            for step, usage in usage_log.items():
                if usage:
                    logger.info(f"  {step}: {usage.total_tokens} tokens (入力: {usage.prompt_tokens}, 出力: {usage.completion_tokens})")

            return final_md

        except Exception as e:
            logger.error(f"エラー: 記事の生成に失敗 - {e}")
            return f"ERROR: {str(e)}"

    def audit_draft(self, design_md, draft_md):
        """Step2: 監査（問題点をJSONで返す、本文は変更しない）"""
        try:
            logger.info(f"[Step2] 監査を実行中...")

            response = self.openai_client.chat.completions.create(
                model="gpt-4o-mini",
                messages=[
                    {"role": "system", "content": """本文は一切変更しない。問題点だけをJSONで返す。
観点：見出し逸脱/漏れ、重複、矛盾、用語ゆれ、薄い（具体例/手順不足）

出力(JSONのみ)：
{
  "offtrack":[{"location":"","issue":"","fix":""}],
  "repetition":[{"location":"","issue":"","fix":""}],
  "contradiction":[{"location":"","issue":"","fix":""}],
  "term":[{"preferred":"","found":["",""],"locations":[""]}],
  "thin":[{"h2":"","issue":"","fix":""}]
}"""},
                    {"role": "user", "content": f"""# 全体設計
{design_md}

本文：
{draft_md}"""}
                ],
                max_completion_tokens=2000
            )

            issues_json = response.choices[0].message.content
            usage = response.usage
            logger.info(f"[Step2] 監査完了（tokens: {usage.total_tokens}）")

            return issues_json, usage

        except Exception as e:
            logger.error(f"エラー: 監査に失敗 - {e}")
            return f"ERROR: {str(e)}", None

    def refine_draft(self, draft_md, issues_json):
        """Step3: 最小修正（JSON指摘箇所のみ修正）"""
        try:
            logger.info(f"[Step3] 最小修正を実行中...")

            current_chars = len(draft_md)

            response = self.openai_client.chat.completions.create(
                model="gpt-4.1",
                messages=[
                    {"role": "system", "content": f"""あなたはクライアント企業メディアの記事編集者です。指摘JSONに従って記事を修正してください。

# ★最重要ルール★ 文字数調整
- 目標: 5,000〜6,000字の範囲内（絶対厳守）
- 現在: {current_chars}字
- 調整方針:
  - 5,000字未満の場合 → 具体例や説明を追加して5,000字以上にする
  - 6,000字超過の場合 → 冗長な部分を削って6,000字以下にする
  - 範囲内の場合 → 文字数を維持しつつ品質向上のみ行う

# 修正ルール
1. 指摘された箇所のみを修正する
2. 見出しは絶対に追加・変更・削除しない
3. 新しいH2、H3、H4を追加してはいけない
4. 「働」のみ平仮名（はたらく等）。履歴書・自己PR・面接・職場等は漢字のまま。
5. 太字マークダウン（**）は使わない

# 文体ルール
- です・ます調で統一
- 簡潔に書く。冗長な説明は削除

# 出力
- 記事全体を出力してください
- ★5,000字未満または6,000字超過は不合格★

# 絶対禁止（出力に含めないこと）
- 文字数カウント（例：「文字数：5,000字」「(5,162字)」等）
- チェック結果やコメント
- 説明文や前置き"""},
                    {"role": "user", "content": f"""# 指摘JSON
{issues_json}

# 現在の記事（{current_chars}字）
{draft_md}

上記の指摘に従って修正し、記事全体を出力してください。
★文字数は5,000〜6,000字の範囲内に収めてください★"""}
                ],
                max_completion_tokens=8000
            )

            final_md = response.choices[0].message.content
            usage = response.usage
            char_count = len(final_md)
            logger.info(f"[Step3] 修正完了（文字数: {char_count}字、tokens: {usage.total_tokens}）")

            return final_md, usage

        except Exception as e:
            logger.error(f"エラー: 修正に失敗 - {e}")
            return f"ERROR: {str(e)}", None

    def append_content_if_needed(self, final_md, issues_json, missing_chars):
        """文字数追記（必要時のみ）"""
        try:
            logger.info(f"[Step4] 文字数追記を実行中...")

            # issues_jsonから薄いH2を抽出（簡易的な実装）
            import json
            try:
                issues = json.loads(issues_json)
                thin_h2s = [item['h2'] for item in issues.get('thin', [])]
                target_h2s = ', '.join(thin_h2s) if thin_h2s else "内容が薄いセクション"
            except:
                target_h2s = "内容が薄いセクション"

            current_chars = len(final_md)
            target_chars = 5500

            # 追記すべき文字数を計算（目標5500字）
            chars_to_add = max(missing_chars, 5500 - current_chars)

            response = self.openai_client.chat.completions.create(
                model="gpt-4o-mini",
                messages=[
                    {"role": "system", "content": f"""あなたはクライアント企業メディアの記事編集者です。

# タスク
現在の記事は{current_chars}字で、目標の5,500字に対して約{chars_to_add}字不足しています。
内容を追記して、必ず5,000～6,000字の範囲に収めてください。

# 追記の方法（{chars_to_add}字分を追加）
1. 各H2セクションに具体例を1〜2個追加
2. 統計データや数字を追加（「約○○%」「○○万人」など）
3. 表を1つ追加（Markdown形式で）
4. 「例えば〜」「具体的には〜」で具体例を導入
5. 自然な締めの文を追加（毎回同じパターンの呼びかけは避ける）

# 表はMarkdown形式で書くこと
| 項目 | 内容 |
|------|------|
| 例 | 説明 |
↑ 各行は | で始まり | で終わること

# 文体ルール
- です・ます調で統一
- 1段落は3〜4文で改行
- 親しみやすく、読者に寄り添う語り口

# 禁止事項（絶対厳守）
- 見出しの追加は絶対禁止（新しいH2、H3、H4を作らない）
- 見出しの変更は絶対禁止（1文字も変えない）
- 見出しの削除は絶対禁止
- 「働」の漢字は全て平仮名（共働き→共ばたらき も含む）
- 太字マークダウン（**）
- 既存の文章の削除

# 重要
- 必ず記事全体を出力してください（追記部分だけでなく、全文）
- 出力する記事は5,000字以上にしてください
- 見出しは追加も変更もしないでください（既存の見出しをそのまま維持）

# 絶対禁止（出力に含めないこと）
- 文字数カウント（例：「文字数：5,000字」「(5,162字)」等）
- チェック結果やコメント
- 説明文や前置き"""},
                    {"role": "user", "content": f"""# 追記が必要なセクション
{target_h2s}

# 現在の記事（{current_chars}字）
{final_md}

上記の記事に約{chars_to_add}字分の内容を追記して、5,000～6,000字にしてください。
見出しは1文字も変更しないでください。

追記のアイデア:
- 具体例を追加
- 統計データを追加
- 表を追加（Markdown形式で。| 項目 | 内容 | の形式）
- 自然な締めの文を追加（同じパターンの連発は避ける）

記事全体を出力してください（追記部分だけでなく、全文を出力すること）。"""}
                ],
                max_completion_tokens=16000
            )

            appended_md = response.choices[0].message.content
            usage = response.usage
            char_count = len(appended_md)
            logger.info(f"[Step4] 追記完了（文字数: {char_count}字、tokens: {usage.total_tokens}）")

            return appended_md, usage

        except Exception as e:
            logger.error(f"エラー: 文字数追記に失敗 - {e}")
            return final_md, None

    def _format_headings_md(self, headings):
        """見出し構造をMarkdown形式の文字列に変換"""
        headings_formatted = []
        for heading in headings:
            level = heading['level']
            text = heading['text']
            indent = '  ' * (int(level[1]) - 2)  # H2=インデントなし、H3=2スペース、H4=4スペース
            headings_formatted.append(f"{indent}{level}: {text}")
        return '\n'.join(headings_formatted)

    def _create_prompt(self, keyword, h1_title, headings):
        """プロンプトを作成"""
        try:
            # Cloud Run環境でも動作するようにファイルパスを取得
            script_dir = os.path.dirname(os.path.abspath(__file__))
            prompt_file = os.path.join(script_dir, 'article_generation_prompt.txt')

            logger.info(f"[DEBUG] プロンプトファイルパス: {prompt_file}")

            with open(prompt_file, 'r', encoding='utf-8') as f:
                template = f.read()
        except FileNotFoundError as e:
            logger.error(f"エラー: プロンプトファイルが見つかりません - {e}")
            raise Exception(f"プロンプトファイル 'article_generation_prompt.txt' が見つかりません")

        # 見出し構造をフォーマット
        headings_str = self._format_headings_md(headings)

        # H2の数をカウント
        h2_count = len([h for h in headings if h['level'] == 'H2'])
        
        # 各セクションの目標文字数を計算
        # 目標総文字数(5500) - 導入・まとめ(約1000) = 本文(4500)
        # 本文(4500) / H2の数
        target_total_body = 4500
        target_per_section = target_total_body // max(1, h2_count)
        
        # 目安として「○○～○○文字」の形式にする
        min_target = target_per_section - 100
        max_target = target_per_section + 100
        target_str = f"{min_target}～{max_target}"

        logger.info(f"[DEBUG] H2数: {h2_count}, セクション目標: {target_str}文字")

        prompt = template.replace('{keyword}', keyword)
        prompt = prompt.replace('{h1_title}', h1_title)
        prompt = prompt.replace('{headings}', headings_str)
        prompt = prompt.replace('{target_per_section}', target_str)

        return prompt

    def get_or_create_monthly_doc_folder(self, year, month):
        """月別ドキュメントフォルダを取得または作成"""
        folder_name = f"{year}年{month}月"
        parent_folder_id = os.environ.get('DOCUMENT_FOLDER_ID', '0AJqmuE7ZYVocUk9PVA')

        # フォルダ内で既存のサブフォルダを検索
        query = f"name = '{folder_name}' and mimeType = 'application/vnd.google-apps.folder' and '{parent_folder_id}' in parents and trashed = false"

        try:
            results = self.drive_service.files().list(
                q=query,
                spaces='drive',
                fields='files(id, name)',
                supportsAllDrives=True,
                includeItemsFromAllDrives=True
            ).execute()

            files = results.get('files', [])

            if files:
                folder_id = files[0]['id']
                logger.info(f"✓ 既存の月別フォルダ「{folder_name}」を使用: {folder_id}")
                return folder_id

            # 新規作成
            logger.info(f"月別フォルダ「{folder_name}」を新規作成中...")
            file_metadata = {
                'name': folder_name,
                'mimeType': 'application/vnd.google-apps.folder',
                'parents': [parent_folder_id]
            }

            folder = self.drive_service.files().create(
                body=file_metadata,
                supportsAllDrives=True,
                fields='id'
            ).execute()

            folder_id = folder.get('id')
            logger.info(f"✓ 月別フォルダ「{folder_name}」を作成: {folder_id}")
            return folder_id

        except Exception as e:
            logger.error(f"エラー: 月別フォルダの取得/作成に失敗 - {e}")
            # フォールバック: 親フォルダを使用
            return parent_folder_id

    def get_year_month_from_spreadsheet(self):
        """スプレッドシート名から年月を抽出"""
        try:
            # スプレッドシートのメタデータを取得
            spreadsheet = self.sheets_service.spreadsheets().get(
                spreadsheetId=self.spreadsheet_id
            ).execute()

            spreadsheet_name = spreadsheet.get('properties', {}).get('title', '')
            logger.info(f"[DEBUG] スプレッドシート名: {spreadsheet_name}")

            # 「YYYY年M月」形式から年月を抽出
            import re
            match = re.match(r'(\d{4})年(\d{1,2})月', spreadsheet_name)
            if match:
                year = int(match.group(1))
                month = int(match.group(2))
                return year, month

            # フォールバック: 現在の年月を使用
            from datetime import datetime
            now = datetime.now()
            return now.year, now.month

        except Exception as e:
            logger.error(f"エラー: スプレッドシート名の取得に失敗 - {e}")
            from datetime import datetime
            now = datetime.now()
            return now.year, now.month

    def save_to_google_docs(self, article, title):
        """Googleドキュメントに保存（見出しスタイル付き、表対応、月別フォルダ）"""
        try:
            # 保存前の文字数をログ出力
            logger.info(f"[DEBUG] Google Docs保存前の文字数: {len(article)}字")

            # 年月を取得してフォルダを決定
            year, month = self.get_year_month_from_spreadsheet()
            FOLDER_ID = self.get_or_create_monthly_doc_folder(year, month)
            logger.info(f"[DEBUG] 保存先フォルダID: {FOLDER_ID} ({year}年{month}月)")

            # Drive APIを使って指定フォルダにファイルを作成
            file_metadata = {
                'name': title,
                'mimeType': 'application/vnd.google-apps.document',
                'parents': [FOLDER_ID]
            }

            if not self.drive_service:
                logger.error("Drive Serviceが初期化されていません")
                return None

            # ファイル作成（空のドキュメント）共有ドライブ対応
            doc = self.drive_service.files().create(
                body=file_metadata,
                supportsAllDrives=True
            ).execute()
            document_id = doc.get('id')

            # 記事からMarkdownの表を抽出
            tables = self._extract_tables_from_article(article)
            logger.info(f"[DEBUG] 抽出された表の数: {len(tables)}")

            # マークダウンを解析してGoogleドキュメントのリクエストに変換（表はプレースホルダー）
            requests, table_count = self._convert_markdown_to_docs_requests(article)

            # テキストを挿入
            self.docs_service.documents().batchUpdate(
                documentId=document_id,
                body={'requests': requests}
            ).execute()

            # 表データを返す（画像挿入後にテーブルを挿入するため）
            doc_url = f"https://docs.google.com/document/d/{document_id}/edit"
            return doc_url, document_id, tables

        except HttpError as err:
            logger.error(f"エラー: ドキュメントの作成に失敗 - {err}")
            return None, None, []

    def _parse_markdown_table(self, table_lines):
        """Markdownの表をパースして2次元配列で返す"""
        rows = []
        for line in table_lines:
            # 区切り行（|---|---|）はスキップ
            if re.match(r'^\|[\s\-:]+\|$', line.replace(' ', '').replace('-', '-')):
                continue
            if '---' in line and '|' in line:
                continue

            # セルを分割
            cells = [cell.strip() for cell in line.split('|')]
            # 先頭と末尾の空要素を削除
            if cells and cells[0] == '':
                cells = cells[1:]
            if cells and cells[-1] == '':
                cells = cells[:-1]

            if cells:
                rows.append(cells)

        return rows

    def _is_table_line(self, line):
        """行がMarkdownの表の一部かどうかを判定"""
        line = line.strip()
        if not line:
            return False
        # |で始まり|で終わる、または区切り行
        if line.startswith('|') and line.endswith('|'):
            return True
        return False

    def _convert_markdown_table_to_html(self, table_lines):
        """MarkdownテーブルをHTMLテーブルに変換"""
        if not table_lines:
            return ''

        html_rows = []
        is_header = True

        for line in table_lines:
            line = line.strip()

            # 区切り行（|---|---|）はスキップ
            if re.match(r'^\|[\s\-:]+\|$', line.replace(' ', '')):
                continue
            if '---' in line and '|' in line:
                is_header = False
                continue

            # セルを分割
            cells = [cell.strip() for cell in line.split('|')[1:-1]]

            if is_header:
                # ヘッダー行
                row_html = '<tr>' + ''.join(f'<th>{cell}</th>' for cell in cells) + '</tr>'
                is_header = False
            else:
                # データ行
                row_html = '<tr>' + ''.join(f'<td>{cell}</td>' for cell in cells) + '</tr>'

            html_rows.append(row_html)

        return '<table>\n' + '\n'.join(html_rows) + '\n</table>'

    def _convert_markdown_to_docs_requests(self, article):
        """マークダウンをGoogleドキュメントのリクエストに変換（表はプレースホルダー）"""
        import re

        requests = []
        current_index = 1

        # 記事を行ごとに処理
        lines = article.split('\n')

        # デバッグ：最初の20行をログ出力
        logger.info(f"[DEBUG] 記事の最初の20行:")
        for i, line in enumerate(lines[:20]):
            logger.info(f"[DEBUG] 行{i+1}: '{line[:100]}'" if len(line) > 100 else f"[DEBUG] 行{i+1}: '{line}'")

        heading_count = 0
        table_count = 0
        i = 0

        while i < len(lines):
            line = lines[i]

            # 表の検出 - マークダウン形式でそのまま挿入
            if self._is_table_line(line):
                # 表の開始
                table_lines = []
                while i < len(lines) and self._is_table_line(lines[i]):
                    table_lines.append(lines[i])
                    i += 1

                # マークダウン形式の表をそのまま挿入
                table_count += 1
                table_text = '\n'.join(table_lines) + '\n\n'
                requests.append({
                    'insertText': {
                        'location': {'index': current_index},
                        'text': table_text
                    }
                })
                current_index += len(table_text)
                logger.info(f"[DEBUG] マークダウン表{table_count}を挿入: {len(table_lines)}行")
                continue

            # 見出しを検出（スペースなしのパターンも対応）
            heading_match = re.match(r'^(#{1,4})\s*(.+)$', line.strip())

            if heading_match:
                # 見出しレベルを取得
                hashes = heading_match.group(1)
                heading_text = heading_match.group(2).strip()
                heading_level = len(hashes)
                heading_count += 1

                logger.info(f"[DEBUG] 見出し検出 H{heading_level}: '{heading_text}'")

                # 見出しテキストを挿入
                requests.append({
                    'insertText': {
                        'location': {'index': current_index},
                        'text': heading_text + '\n'
                    }
                })

                # 見出しスタイルを適用
                named_style = f'HEADING_{heading_level}'
                requests.append({
                    'updateParagraphStyle': {
                        'range': {
                            'startIndex': current_index,
                            'endIndex': current_index + len(heading_text)
                        },
                        'paragraphStyle': {
                            'namedStyleType': named_style
                        },
                        'fields': 'namedStyleType'
                    }
                })

                current_index += len(heading_text) + 1
            else:
                # 通常のテキスト
                if line:
                    requests.append({
                        'insertText': {
                            'location': {'index': current_index},
                            'text': line + '\n'
                        }
                    })
                    current_index += len(line) + 1
                else:
                    # 空行
                    requests.append({
                        'insertText': {
                            'location': {'index': current_index},
                            'text': '\n'
                        }
                    })
                    current_index += 1

            i += 1

        logger.info(f"[DEBUG] マークダウン変換完了: 見出し総数={heading_count}, 表数={table_count}")
        return requests, table_count

    def _extract_tables_from_article(self, article):
        """記事からMarkdownの表を抽出"""
        tables = []
        lines = article.split('\n')
        i = 0

        while i < len(lines):
            if self._is_table_line(lines[i]):
                table_lines = []
                while i < len(lines) and self._is_table_line(lines[i]):
                    table_lines.append(lines[i])
                    i += 1

                parsed_table = self._parse_markdown_table(table_lines)
                if parsed_table:
                    tables.append(parsed_table)
            else:
                i += 1

        return tables

    def _insert_tables_into_doc(self, document_id, tables):
        """ドキュメント内のプレースホルダーを実際の表に置換"""
        if not tables:
            return

        # 表を逆順で処理（後ろから処理することでインデックスのずれを防ぐ）
        for table_idx in range(len(tables), 0, -1):
            placeholder = f"[[TABLE_PLACEHOLDER_{table_idx}]]"
            table_data = tables[table_idx - 1]

            if not table_data or len(table_data) == 0:
                continue

            # ドキュメントを取得してプレースホルダーの位置を特定
            doc = self.docs_service.documents().get(documentId=document_id).execute()
            content = doc.get('body', {}).get('content', [])

            placeholder_index = None
            placeholder_end_index = None

            for element in content:
                if 'paragraph' in element:
                    paragraph = element['paragraph']
                    for elem in paragraph.get('elements', []):
                        if 'textRun' in elem:
                            text = elem['textRun'].get('content', '')
                            if placeholder in text:
                                placeholder_index = elem.get('startIndex')
                                placeholder_end_index = elem.get('endIndex')
                                break
                    if placeholder_index:
                        break

            if placeholder_index is None:
                logger.warning(f"[WARNING] プレースホルダー {placeholder} が見つかりません")
                continue

            # 行数と列数
            num_rows = len(table_data)
            num_cols = max(len(row) for row in table_data) if table_data else 1

            logger.info(f"[DEBUG] 表{table_idx}を挿入: {num_rows}行 x {num_cols}列, index={placeholder_index}")

            # プレースホルダーを削除
            requests = [{
                'deleteContentRange': {
                    'range': {
                        'startIndex': placeholder_index,
                        'endIndex': placeholder_end_index
                    }
                }
            }]

            self.docs_service.documents().batchUpdate(
                documentId=document_id,
                body={'requests': requests}
            ).execute()

            # 表を挿入
            requests = [{
                'insertTable': {
                    'rows': num_rows,
                    'columns': num_cols,
                    'location': {'index': placeholder_index}
                }
            }]

            self.docs_service.documents().batchUpdate(
                documentId=document_id,
                body={'requests': requests}
            ).execute()

            # 表のセルにテキストを挿入（ドキュメントを再取得してセルのインデックスを取得）
            doc = self.docs_service.documents().get(documentId=document_id).execute()
            content = doc.get('body', {}).get('content', [])

            # 表を探す
            table_element = None
            for element in content:
                if 'table' in element:
                    start_idx = element.get('startIndex', 0)
                    if start_idx >= placeholder_index - 5 and start_idx <= placeholder_index + 5:
                        table_element = element['table']
                        break

            if not table_element:
                logger.warning(f"[WARNING] 挿入した表が見つかりません")
                continue

            # 各セルにテキストを挿入（逆順で処理）
            cell_requests = []
            table_rows = table_element.get('tableRows', [])

            for row_idx in range(len(table_rows) - 1, -1, -1):
                row = table_rows[row_idx]
                cells = row.get('tableCells', [])

                for col_idx in range(len(cells) - 1, -1, -1):
                    cell = cells[col_idx]
                    cell_content = cell.get('content', [])

                    if cell_content:
                        # セルの最初の段落のインデックスを取得
                        first_para = cell_content[0]
                        if 'paragraph' in first_para:
                            cell_start = first_para['paragraph']['elements'][0]['startIndex']

                            # テーブルデータから対応するテキストを取得
                            if row_idx < len(table_data) and col_idx < len(table_data[row_idx]):
                                cell_text = table_data[row_idx][col_idx]
                                if cell_text:
                                    cell_requests.append({
                                        'insertText': {
                                            'location': {'index': cell_start},
                                            'text': cell_text
                                        }
                                    })

            if cell_requests:
                self.docs_service.documents().batchUpdate(
                    documentId=document_id,
                    body={'requests': cell_requests}
                ).execute()

            logger.info(f"[DEBUG] 表{table_idx}の挿入完了")

    def generate_image_with_vertex(self, h2_heading, keyword, max_retries=3):
        """Vertex AI Imagenで画像を生成（リトライ処理付き）"""
        # 日本語プロンプトを作成
        prompt = f"{keyword}、{h2_heading}に関連する明るく親しみやすいイラスト"

        logger.info(f"画像生成中: {prompt}")

        # Vertex AIの初期化（認証情報を使用）
        if self.credentials:
            aiplatform.init(
                project=self.project_id,
                location='us-central1',
                credentials=self.credentials
            )
        else:
            logger.error("認証情報が設定されていません")
            return None

        # Imagen 3を使用して画像を生成
        from vertexai.preview.vision_models import ImageGenerationModel

        model = ImageGenerationModel.from_pretrained("imagen-3.0-generate-001")

        # リトライ処理（指数バックオフ）
        for attempt in range(max_retries):
            try:
                images = model.generate_images(
                    prompt=prompt,
                    number_of_images=1,
                )

                # 画像が生成されたか確認
                if images is None:
                    logger.warning(f"[VERTEX] 画像が生成されませんでした（None応答）: {prompt}")
                    return "ERROR: 画像が生成されませんでした（None応答）"

                # イテレートして画像を取得
                image_list = list(images)
                if not image_list:
                    logger.warning(f"[VERTEX] 画像が生成されませんでした（空のリスト）: {prompt}")
                    return "ERROR: 画像が生成されませんでした（コンテンツポリシー等）"

                # 最初の画像を取得
                image = image_list[0]

                # 画像データをバイト形式で取得
                image_bytes = image._image_bytes

                logger.info(f"✓ 画像生成完了: {len(image_bytes)} bytes")
                return image_bytes

            except IndexError as e:
                logger.warning(f"[VERTEX] 画像リストが空でした: {e}")
                return "ERROR: 画像が生成されませんでした（IndexError）"
            except Exception as e:
                error_str = str(e)

                # クォータエラー（429）の検出
                if '429' in error_str or 'quota' in error_str.lower() or 'resource exhausted' in error_str.lower():
                    if attempt < max_retries - 1:
                        # 指数バックオフ: 30秒、60秒、120秒
                        wait_time = 30 * (2 ** attempt)
                        logger.warning(f"[QUOTA] クォータ制限検出。{wait_time}秒待機後にリトライ ({attempt + 1}/{max_retries})")
                        time.sleep(wait_time)
                        continue
                    else:
                        logger.error(f"[QUOTA] クォータ制限: {max_retries}回リトライ後も失敗。この画像をスキップします")
                        return "ERROR:QUOTA_EXCEEDED"

                # その他のエラー
                logger.error(f"エラー: 画像の生成に失敗しました - {e}")
                return f"ERROR: {error_str}"

        return "ERROR: リトライ上限到達"

    def upload_image_to_drive(self, image_bytes, filename, max_retries=3):
        """生成した画像をGoogle Driveにアップロード（リトライ処理付き）"""
        if not self.image_folder_id:
            logger.error("画像フォルダーIDが設定されていません")
            return None

        # 画像をPNG形式に変換
        from io import BytesIO
        image = Image.open(BytesIO(image_bytes))

        # PNG形式で保存
        output = BytesIO()
        image.save(output, format='PNG')

        # Google Driveにアップロード
        file_metadata = {
            'name': f"{filename}.png",
            'parents': [self.image_folder_id],
            'mimeType': 'image/png'
        }

        from googleapiclient.http import MediaIoBaseUpload

        # リトライ処理
        for attempt in range(max_retries):
            try:
                output.seek(0)  # リトライ時にストリームを先頭に戻す
                media = MediaIoBaseUpload(output, mimetype='image/png', resumable=True)

                uploaded_file = self.drive_service.files().create(
                    body=file_metadata,
                    media_body=media,
                    fields='id, webViewLink, webContentLink',
                    supportsAllDrives=True
                ).execute()

                file_id = uploaded_file.get('id')

                # 組織内共有設定 - 親フォルダから継承されている場合はスキップ
                try:
                    self.drive_service.permissions().create(
                        fileId=file_id,
                        body={'type': 'domain', 'domain': 'vexum-ai.com', 'role': 'reader'},
                        supportsAllDrives=True
                    ).execute()
                except Exception as perm_error:
                    # 親フォルダから権限が継承されている場合はエラーになるが、問題ない
                    logger.warning(f"権限設定をスキップ（親フォルダから継承）: {perm_error}")

                # 直接アクセス可能なURLを生成
                image_url = f"https://drive.google.com/uc?export=view&id={file_id}"

                logger.info(f"✓ 画像アップロード完了: {image_url}")
                return image_url

            except Exception as e:
                error_str = str(e)
                if attempt < max_retries - 1:
                    # 500エラーやタイムアウトの場合はリトライ
                    if '500' in error_str or '503' in error_str or 'timeout' in error_str.lower():
                        wait_time = 10 * (attempt + 1)  # 10秒、20秒、30秒
                        logger.warning(f"[RETRY] アップロード失敗。{wait_time}秒待機後にリトライ ({attempt + 1}/{max_retries}): {e}")
                        time.sleep(wait_time)
                        continue
                logger.error(f"エラー: 画像のアップロードに失敗しました - {e}")
                return None

        return None

    def update_sheet_status(self, sheet_name, status="処理済み", doc_url=""):
        """ステータスを更新（リトライ機能付き）"""
        max_retries = 3
        for attempt in range(max_retries):
            try:
                range_name = f"'{sheet_name}'!F2:G2"
                values = [[status, doc_url]]
                
                logger.info(f"[UPDATE_STATUS] シート名: {sheet_name} (試行 {attempt+1}/{max_retries})")
                logger.info(f"[UPDATE_STATUS] 範囲: {range_name}")
                logger.info(f"[UPDATE_STATUS] ステータス: {status}")
                logger.info(f"[UPDATE_STATUS] URL: {doc_url}")

                result = self.sheets_service.spreadsheets().values().update(
                    spreadsheetId=self.spreadsheet_id,
                    range=range_name,
                    valueInputOption='RAW',
                    body={'values': values}
                ).execute()
                
                logger.info(f"[UPDATE_STATUS] ✓ 更新成功: {result.get('updatedCells', 0)}セル更新")
                return  # 成功したら終了

            except HttpError as err:
                logger.error(f"[UPDATE_STATUS] ✗ エラー (試行 {attempt+1}): {err}")
                if attempt < max_retries - 1:
                    import time
                    time.sleep(2 ** attempt)  # 指数バックオフ (1s, 2s, 4s...)
                else:
                    logger.error(f"[UPDATE_STATUS] 全てのリトライに失敗しました")
            except Exception as e:
                logger.error(f"[UPDATE_STATUS] ✗ 予期しないエラー: {e}")
                import traceback
                logger.error(f"[UPDATE_STATUS] トレースバック: {traceback.format_exc()}")
                break  # 予期しないエラーはリトライしない

    def update_master_sheet_article_url(self, master_spreadsheet_id, keyword, doc_url, keyword_column='G', url_column='N'):
        """マスターシートに初稿URLを書き込む

        Args:
            master_spreadsheet_id: マスターシートのスプレッドシートID
            keyword: キーワード（照合用）
            doc_url: 初稿のURL
            keyword_column: キーワード列（デフォルト: G）
            url_column: 初稿URL書き込み列（デフォルト: N）
        """
        try:
            logger.info(f"[MASTER_UPDATE] マスターシートに初稿URL書き込み中...")
            logger.info(f"[MASTER_UPDATE] キーワード: {keyword}")
            logger.info(f"[MASTER_UPDATE] URL: {doc_url}")

            # マスターシートのキーワード列を取得
            result = self.sheets_service.spreadsheets().values().get(
                spreadsheetId=master_spreadsheet_id,
                range=f'{keyword_column}:{keyword_column}'
            ).execute()

            values = result.get('values', [])

            if not values:
                logger.warning("[MASTER_UPDATE] マスターシートにデータがありません")
                return False

            # キーワードと行番号のマッピングを作成
            row_num = None
            for i, row in enumerate(values):
                if row and row[0]:
                    cell_keyword = row[0].strip()
                    if cell_keyword == keyword:
                        row_num = i + 1  # 1-indexed
                        break

            if not row_num:
                logger.warning(f"[MASTER_UPDATE] キーワード「{keyword}」がマスターシートに見つかりません")
                return False

            # URLを書き込む
            self.sheets_service.spreadsheets().values().update(
                spreadsheetId=master_spreadsheet_id,
                range=f'{url_column}{row_num}',
                valueInputOption='RAW',
                body={'values': [[doc_url]]}
            ).execute()

            logger.info(f"[MASTER_UPDATE] ✓ マスターシート {url_column}{row_num} にURL書き込み完了")
            return True

        except Exception as e:
            logger.error(f"[MASTER_UPDATE] エラー: {e}")
            return False

    def get_image_folders(self):
        """画像フォルダー内のサブフォルダーとその中の画像を取得"""
        if not self.image_folder_id:
            logger.info("画像フォルダーIDが設定されていません")
            return {}

        if self.image_cache is not None:
            logger.info("キャッシュから画像フォルダー情報を返します")
            return self.image_cache

        try:
            # サブフォルダーを取得（共有ドライブ対応）
            query = f"'{self.image_folder_id}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false"

            logger.info(f"[DEBUG] 画像フォルダーID: {self.image_folder_id}")
            logger.info(f"[DEBUG] クエリ: {query}")

            response = self.drive_service.files().list(
                q=query,
                spaces='drive',
                fields='files(id, name)',
                supportsAllDrives=True,
                includeItemsFromAllDrives=True,
                corpora='allDrives'  # 共有ドライブも検索対象に含める
            ).execute()

            subfolders = response.get('files', [])
            logger.info(f"サブフォルダー数: {len(subfolders)}")

            if subfolders:
                logger.info(f"[DEBUG] 検出されたサブフォルダー: {[f['name'] for f in subfolders]}")

            # 各サブフォルダー内の画像を取得
            folder_images = {}
            for folder in subfolders:
                folder_name = folder['name']
                folder_id = folder['id']

                # サブフォルダー内の画像を取得（共有ドライブ対応）
                image_query = f"'{folder_id}' in parents and (mimeType contains 'image/') and trashed=false"
                image_response = self.drive_service.files().list(
                    q=image_query,
                    spaces='drive',
                    fields='files(id, name, webContentLink)',
                    supportsAllDrives=True,
                    includeItemsFromAllDrives=True,
                    corpora='allDrives'  # 共有ドライブも検索対象に含める
                ).execute()

                images = image_response.get('files', [])
                if images:
                    folder_images[folder_name] = images
                    logger.info(f"フォルダー '{folder_name}': {len(images)}枚の画像")

            self.image_cache = folder_images
            return folder_images

        except HttpError as err:
            logger.error(f"エラー: 画像フォルダーの取得に失敗 - {err}")
            return {}

    def match_heading_to_folder(self, h2_text, folder_names):
        """H2見出しに最適なフォルダーをAIで選択"""
        if not folder_names:
            return None

        try:
            folder_list = '\n'.join([f"- {name}" for name in folder_names])
            prompt = f"""# フォルダーマッチングタスク

以下のH2見出しに最も関連性の高いフォルダー名を1つだけ選んでください。

## H2見出し
{h2_text}

## 利用可能なフォルダー
{folder_list}

## 出力ルール
- フォルダー名のみを回答してください
- フォルダーが見つからない場合は「なし」と回答してください
- 説明や理由は不要です"""

            response = self.openai_client.chat.completions.create(
                model="gpt-4o-mini",
                messages=[
                    {"role": "system", "content": "あなたは記事の見出しとフォルダー名をマッチングする専門家です。"},
                    {"role": "user", "content": prompt}
                ],
                max_completion_tokens=50
            )

            selected_folder = response.choices[0].message.content.strip()
            logger.info(f"H2 '{h2_text}' → フォルダー '{selected_folder}'")

            if selected_folder == "なし" or selected_folder not in folder_names:
                return None

            return selected_folder

        except Exception as e:
            logger.error(f"エラー: フォルダーマッチングに失敗 - {e}")
            return None

    def insert_images_into_doc(self, document_id, h2_headings):
        """H2見出しの後に画像を挿入"""
        logger.info(f"[DEBUG] 画像挿入開始 - document_id: {document_id}")

        if not self.image_folder_id:
            logger.warning("⚠️ 画像フォルダーID (IMAGE_FOLDER_ID) が設定されていないため、画像挿入をスキップします")
            return

        # 画像フォルダー情報を取得
        folder_images = self.get_image_folders()
        if not folder_images:
            logger.warning("⚠️ 画像フォルダー内に画像が見つからないため、画像挿入をスキップします")
            return

        folder_names = list(folder_images.keys())
        logger.info(f"利用可能なフォルダー: {folder_names}")

        try:
            # ドキュメントの内容を取得
            logger.info(f"[DEBUG] ドキュメント内容を取得中...")
            doc = self.docs_service.documents().get(documentId=document_id).execute()
            content = doc.get('body').get('content')
            logger.info(f"[DEBUG] ドキュメント要素数: {len(content)}")

            # H2見出しを検索（Googleドキュメントの見出しスタイル or マークダウン形式に対応）
            # スペースなしでもマッチするように修正
            h2_pattern = re.compile(r'^\s*##\s*(.+?)\s*$')

            # まず全てのH2見出しを検出
            h2_elements = []
            paragraph_count = 0
            for element in content:
                if 'paragraph' in element:
                    paragraph_count += 1
                    paragraph = element['paragraph']

                    # パラグラフのテキストを取得
                    para_text = ''
                    for text_element in paragraph.get('elements', []):
                        if 'textRun' in text_element:
                            para_text += text_element['textRun']['content']

                    para_text_stripped = para_text.strip()

                    # 見出しスタイルをチェック（HEADING_2）
                    paragraph_style = paragraph.get('paragraphStyle', {})
                    named_style = paragraph_style.get('namedStyleType', '')

                    if named_style == 'HEADING_2':
                        # Googleドキュメントの見出しスタイル
                        h2_text = para_text_stripped
                        h2_elements.append({
                            'text': h2_text,
                            'element': element,
                            'end_index': element['endIndex']
                        })
                        logger.info(f"[DEBUG] H2見出しスタイル検出: '{h2_text}' (end_index: {element['endIndex']})")
                    else:
                        # マークダウン形式もサポート（後方互換性）
                        if para_text_stripped.startswith('##'):
                            logger.info(f"[DEBUG] マークダウン形式のH2検出: '{para_text_stripped[:100]}'")

                        match = h2_pattern.match(para_text_stripped)
                        if match:
                            h2_text = match.group(1).strip()
                            h2_elements.append({
                                'text': h2_text,
                                'element': element,
                                'end_index': element['endIndex']
                            })
                            logger.info(f"[DEBUG] マークダウンH2マッチ成功: '{h2_text}' (end_index: {element['endIndex']})")
                        elif para_text_stripped.startswith('##'):
                            logger.warning(f"[DEBUG] ## で始まるがマッチせず: '{para_text_stripped[:100]}'")

            logger.info(f"[DEBUG] パラグラフ総数: {paragraph_count}")
            logger.info(f"[DEBUG] 検出されたH2見出し数: {len(h2_elements)}")

            # 検出された全H2を出力
            for idx, h2 in enumerate(h2_elements):
                logger.info(f"[DEBUG] H2 [{idx+1}]: '{h2['text']}' (end_index: {h2['end_index']})")

            # 使用済み画像IDを記録
            used_image_ids = set()
            insert_requests = []

            # 最後のH2を除いて処理（最後は「まとめ」なので画像不要）
            h2_to_process = h2_elements[:-1] if len(h2_elements) > 1 else []
            logger.info(f"[DEBUG] 画像を挿入するH2見出し数: {len(h2_to_process)} (最後のH2を除外)")

            if len(h2_elements) > 0:
                logger.info(f"[DEBUG] 最後のH2（スキップ予定）: '{h2_elements[-1]['text']}'")

            for i, h2_info in enumerate(h2_to_process):
                h2_text = h2_info['text']
                end_index = h2_info['end_index']

                logger.info(f"[DEBUG] H2見出し処理中 ({i+1}/{len(h2_to_process)}): '{h2_text}'")

                # 「まとめ」を含む見出しもスキップ
                if 'まとめ' in h2_text or 'まとめ' in h2_text.lower():
                    logger.warning(f"[SKIP] H2 '{h2_text}': 「まとめ」を含むためスキップ")
                    continue

                # H2見出しに対応するフォルダーを選択（AIマッチング、失敗時はランダム）
                selected_folder = None
                try:
                    selected_folder = self.match_heading_to_folder(h2_text, folder_names)

                    # フォルダー名の検証
                    if selected_folder and selected_folder in folder_names:
                        logger.info(f"H2 '{h2_text}': AIマッチング成功 → フォルダー '{selected_folder}' を選択")
                    else:
                        logger.info(f"H2 '{h2_text}': AIマッチング結果が無効 ('{selected_folder}')、ランダム選択に切り替え")
                        selected_folder = None
                except Exception as e:
                    logger.error(f"H2 '{h2_text}': AIマッチング中にエラー発生 - {e}")
                    selected_folder = None

                # マッチングに失敗した場合は必ずランダムにフォルダーを選択
                if not selected_folder:
                    selected_folder = random.choice(folder_names)
                    logger.info(f"H2 '{h2_text}': ランダムにフォルダー '{selected_folder}' を選択")

                # フォルダーから未使用の画像をランダムに選択
                images = folder_images.get(selected_folder, [])
                if not images:
                    logger.error(f"H2 '{h2_text}': フォルダー '{selected_folder}' に画像がありません！別のフォルダーを選択")
                    # 別のフォルダーから選ぶ
                    for folder_name in folder_names:
                        if folder_images.get(folder_name):
                            selected_folder = folder_name
                            images = folder_images[folder_name]
                            logger.info(f"H2 '{h2_text}': フォルダー '{selected_folder}' に変更")
                            break

                available_images = [img for img in images if img['id'] not in used_image_ids]

                if not available_images:
                    logger.warning(f"H2 '{h2_text}': フォルダー '{selected_folder}' に未使用の画像がありません")
                    # 全ての画像が使用済みの場合は、全フォルダーから未使用の画像を探す
                    all_images = []
                    for folder_imgs in folder_images.values():
                        all_images.extend(folder_imgs)
                    available_images = [img for img in all_images if img['id'] not in used_image_ids]

                    # それでも未使用の画像がない場合は、画像なしでスキップ
                    if not available_images:
                        logger.warning(f"[SKIP] H2 '{h2_text}': 全フォルダーの画像が使用済みのため、画像挿入をスキップします（同じ記事内で画像重複を防止）")
                        continue

                if not available_images:
                    logger.error(f"[SKIP] H2 '{h2_text}': 利用可能な画像が全くありません！スキップします")
                    continue

                selected_image = random.choice(available_images)
                image_id = selected_image['id']
                used_image_ids.add(image_id)

                logger.info(f"H2 '{h2_text}': ✓ 画像 '{selected_image['name']}' (ID: {image_id[:20]}...) を挿入予定")

                # 画像を「リンクを知っている全員」に設定（Google Docs APIからアクセス可能にするため）
                try:
                    # 既存の権限を確認
                    permissions = self.drive_service.permissions().list(
                        fileId=image_id,
                        supportsAllDrives=True,
                        fields='permissions(id, type, role)'
                    ).execute()

                    # 「リンクを知っている全員」の権限があるか確認
                    has_public_link = any(
                        p.get('type') == 'anyone' for p in permissions.get('permissions', [])
                    )

                    # なければ追加（組織内のみに制限）
                    if not has_public_link:
                        self.drive_service.permissions().create(
                            fileId=image_id,
                            supportsAllDrives=True,
                            body={
                                'type': 'domain',
                                'domain': 'vexum-ai.com',
                                'role': 'reader'
                            }
                        ).execute()
                        logger.info(f"[DEBUG] 画像を組織内共有に設定: {image_id}")

                except Exception as e:
                    logger.error(f"権限設定エラー: {e}")

                # 直接アクセス可能なURL形式を使用
                image_url = f"https://lh3.googleusercontent.com/d/{image_id}"

                # 画像挿入リクエストを追加
                insert_requests.append({
                    'insertInlineImage': {
                        'uri': image_url,
                        'location': {'index': end_index},
                        'objectSize': {
                            'height': {'magnitude': 300, 'unit': 'PT'},
                            'width': {'magnitude': 400, 'unit': 'PT'}
                        }
                    }
                })

            # 画像を挿入（逆順で実行してインデックスのずれを防ぐ）
            logger.info(f"[DEBUG] ========================================")
            logger.info(f"[DEBUG] H2見出し検出数: {len(h2_elements)}")
            logger.info(f"[DEBUG] 処理対象H2数（最後を除く）: {len(h2_to_process)}")
            logger.info(f"[DEBUG] 画像挿入リクエスト数: {len(insert_requests)}")
            logger.info(f"[DEBUG] ========================================")

            if insert_requests:
                logger.info(f"[DEBUG] 画像挿入を実行します（{len(insert_requests)}個のリクエスト）...")

                # 各リクエストの詳細をログ出力
                for idx, req in enumerate(insert_requests):
                    img_location = req['insertInlineImage']['location']['index']
                    img_uri = req['insertInlineImage']['uri']
                    logger.info(f"[DEBUG] リクエスト[{idx+1}]: 位置={img_location}, URI={img_uri[:60]}...")

                insert_requests.reverse()

                try:
                    result = self.docs_service.documents().batchUpdate(
                        documentId=document_id,
                        body={'requests': insert_requests}
                    ).execute()
                    logger.info(f"✓ {len(insert_requests)}枚の画像を挿入しました")
                    logger.info(f"[DEBUG] batchUpdate結果: {result}")
                except Exception as batch_error:
                    logger.error(f"[ERROR] batchUpdate実行中にエラー発生: {batch_error}")
                    import traceback
                    logger.error(f"[ERROR] トレースバック: {traceback.format_exc()}")
            else:
                logger.warning(f"[WARNING] 挿入する画像がありません！")

        except Exception as e:
            logger.error(f"エラー: 画像挿入に失敗 - {e}")
            import traceback
            logger.error(f"詳細: {traceback.format_exc()}")

    def insert_generated_images_into_doc(self, document_id, h2_headings, keyword):
        """Vertex AIで画像を生成してH2見出しの後に挿入"""
        logger.info(f"[DEBUG] 画像生成・挿入開始 - document_id: {document_id}")

        if not self.image_folder_id:
            logger.warning("⚠️ 画像フォルダーID (IMAGE_FOLDER_ID) が設定されていないため、画像挿入をスキップします")
            return

        try:
            # ドキュメントの内容を取得
            logger.info(f"[DEBUG] ドキュメント内容を取得中...")
            doc = self.docs_service.documents().get(documentId=document_id).execute()
            content = doc.get('body').get('content')
            logger.info(f"[DEBUG] ドキュメント要素数: {len(content)}")

            # H2見出しを検索（Googleドキュメントの見出しスタイル or マークダウン形式に対応）
            # スペースなしでもマッチするように修正
            h2_pattern = re.compile(r'^\s*##\s*(.+?)\s*$')

            # 全てのH2見出しを検出
            h2_elements = []
            for element in content:
                if 'paragraph' in element:
                    paragraph = element['paragraph']
                    para_text = ''
                    for text_element in paragraph.get('elements', []):
                        if 'textRun' in text_element:
                            para_text += text_element['textRun']['content']

                    para_text_stripped = para_text.strip()

                    # 見出しスタイルをチェック（HEADING_2）
                    paragraph_style = paragraph.get('paragraphStyle', {})
                    named_style = paragraph_style.get('namedStyleType', '')

                    if named_style == 'HEADING_2':
                        # Googleドキュメントの見出しスタイル
                        h2_text = para_text_stripped
                        h2_elements.append({
                            'text': h2_text,
                            'element': element,
                            'end_index': element['endIndex']
                        })
                        logger.info(f"[DEBUG] H2見出しスタイル検出: '{h2_text}' (end_index: {element['endIndex']})")
                    else:
                        # マークダウン形式もサポート（後方互換性）
                        match = h2_pattern.match(para_text_stripped)
                        if match:
                            h2_text = match.group(1).strip()
                            h2_elements.append({
                                'text': h2_text,
                                'element': element,
                                'end_index': element['endIndex']
                            })
                            logger.info(f"[DEBUG] マークダウンH2マッチ: '{h2_text}' (end_index: {element['endIndex']})")

            logger.info(f"[DEBUG] 検出されたH2見出し数: {len(h2_elements)}")

            insert_requests = []
            image_errors = []

            # 最後のH2を除いて処理（最後は「まとめ」なので画像不要）
            h2_to_process = h2_elements[:-1] if len(h2_elements) > 1 else []
            logger.info(f"[DEBUG] 画像を挿入するH2見出し数: {len(h2_to_process)}")

            for i, h2_info in enumerate(h2_to_process):
                h2_text = h2_info['text']
                end_index = h2_info['end_index']

                logger.info(f"[DEBUG] H2見出し処理中 ({i+1}/{len(h2_to_process)}): '{h2_text}'")

                # 「まとめ」を含む見出しをスキップ
                if 'まとめ' in h2_text:
                    logger.warning(f"[SKIP] H2 '{h2_text}': 「まとめ」を含むためスキップ")
                    continue

                # Vertex AIで画像を生成
                image_bytes = self.generate_image_with_vertex(h2_text, keyword)

                if isinstance(image_bytes, str) and image_bytes.startswith("ERROR:"):
                    logger.warning(f"[SKIP] H2 '{h2_text}': 画像生成失敗 - {image_bytes}")
                    image_errors.append(f"H2 '{h2_text}': {image_bytes}")
                    continue

                if not image_bytes:
                    logger.warning(f"[SKIP] H2 '{h2_text}': 画像生成に失敗")
                    continue

                # Google Driveにアップロード
                import time
                timestamp = int(time.time())
                filename = f"{keyword}_{h2_text[:20]}_{timestamp}"
                image_url = self.upload_image_to_drive(image_bytes, filename)

                if not image_url:
                    logger.warning(f"[SKIP] H2 '{h2_text}': 画像アップロードに失敗")
                    continue

                # 画像挿入リクエストを追加
                insert_requests.append({
                    'insertInlineImage': {
                        'uri': image_url,
                        'location': {'index': end_index},
                        'objectSize': {
                            'height': {'magnitude': 300, 'unit': 'PT'},
                            'width': {'magnitude': 400, 'unit': 'PT'}
                        }
                    }
                })

                logger.info(f"✓ H2 '{h2_text}': 画像生成・アップロード完了")

            # 画像を挿入（逆順で実行してインデックスのずれを防ぐ）
            logger.info(f"[DEBUG] 画像挿入リクエスト数: {len(insert_requests)}")

            if insert_requests:
                logger.info(f"[DEBUG] 画像挿入を実行します（{len(insert_requests)}個のリクエスト）...")
                insert_requests.reverse()

                try:
                    result = self.docs_service.documents().batchUpdate(
                        documentId=document_id,
                        body={'requests': insert_requests}
                    ).execute()
                    logger.info(f"✓ {len(insert_requests)}枚の生成画像を挿入しました")
                except Exception as batch_error:
                    logger.error(f"[ERROR] batchUpdate実行中にエラー発生: {batch_error}")
                    import traceback
                    logger.error(f"[ERROR] トレースバック: {traceback.format_exc()}")
                    image_errors.append(f"画像挿入エラー: {str(batch_error)}")
            else:
                logger.warning(f"[WARNING] 挿入する画像がありません！")

            return image_errors

        except Exception as e:
            logger.error(f"エラー: 画像生成・挿入に失敗 - {e}")
            import traceback
            logger.error(f"詳細: {traceback.format_exc()}")
            return [f"画像処理全体エラー: {str(e)}"]

    def insert_both_images_into_doc(self, document_id, h2_headings, keyword):
        """フォルダ画像とVertex AI生成画像の両方をH2見出しの後に挿入（並列処理版）

        人間が最終チェックでどちらか選んで不要な方を削除する想定
        """
        from concurrent.futures import ThreadPoolExecutor, as_completed
        import time

        logger.info(f"[BOTH] 両方の画像挿入開始（並列処理） - document_id: {document_id}")

        if not self.image_folder_id:
            logger.warning("⚠️ 画像フォルダーID (IMAGE_FOLDER_ID) が設定されていないため、画像挿入をスキップします")
            return []

        # 画像フォルダー情報を取得
        folder_images = self.get_image_folders()
        folder_names = list(folder_images.keys()) if folder_images else []
        logger.info(f"[BOTH] 利用可能なフォルダー: {folder_names}")

        try:
            # ドキュメントの内容を取得
            doc = self.docs_service.documents().get(documentId=document_id).execute()
            content = doc.get('body').get('content')

            # H2見出しを検索
            h2_pattern = re.compile(r'^\s*##\s*(.+?)\s*$')
            h2_elements = []

            for element in content:
                if 'paragraph' in element:
                    paragraph = element['paragraph']
                    para_text = ''
                    for text_element in paragraph.get('elements', []):
                        if 'textRun' in text_element:
                            para_text += text_element['textRun']['content']

                    para_text_stripped = para_text.strip()
                    paragraph_style = paragraph.get('paragraphStyle', {})
                    named_style = paragraph_style.get('namedStyleType', '')

                    if named_style == 'HEADING_2':
                        h2_elements.append({
                            'text': para_text_stripped,
                            'element': element,
                            'end_index': element['endIndex']
                        })
                    else:
                        match = h2_pattern.match(para_text_stripped)
                        if match:
                            h2_elements.append({
                                'text': match.group(1).strip(),
                                'element': element,
                                'end_index': element['endIndex']
                            })

            logger.info(f"[BOTH] 検出されたH2見出し数: {len(h2_elements)}")

            image_errors = []
            used_folder_image_ids = set()

            # 最後のH2を除いて処理（「まとめ」には画像不要）
            h2_to_process = h2_elements[:-1] if len(h2_elements) > 1 else []
            # 「まとめ」を含むH2も除外
            h2_to_process = [h for h in h2_to_process if 'まとめ' not in h['text']]
            logger.info(f"[BOTH] 画像を挿入するH2見出し数: {len(h2_to_process)}")

            # フォルダ画像を先に全て割り当て（順番を保持するため）
            folder_assignments = {}
            for i, h2_info in enumerate(h2_to_process):
                h2_text = h2_info['text']
                folder_image_url = None

                if folder_images:
                    # フォルダ名でマッチングを試行
                    matched_folder = None
                    for folder_name in folder_names:
                        if folder_name in h2_text or h2_text in folder_name:
                            matched_folder = folder_name
                            break

                    if matched_folder and folder_images[matched_folder]:
                        for img in folder_images[matched_folder]:
                            if img['id'] not in used_folder_image_ids:
                                folder_image_url = f"https://drive.google.com/uc?id={img['id']}"
                                used_folder_image_ids.add(img['id'])
                                logger.info(f"[BOTH] フォルダ画像取得: {img['name']}")
                                break

                    if not folder_image_url:
                        all_images = []
                        for folder_name, images in folder_images.items():
                            for img in images:
                                if img['id'] not in used_folder_image_ids:
                                    all_images.append(img)
                        if all_images:
                            selected_img = random.choice(all_images)
                            folder_image_url = f"https://drive.google.com/uc?id={selected_img['id']}"
                            used_folder_image_ids.add(selected_img['id'])
                            logger.info(f"[BOTH] フォルダ画像（ランダム）: {selected_img['name']}")

                folder_assignments[i] = folder_image_url

            # Vertex AI画像生成を並列実行するための関数
            def generate_ai_image_task(index, h2_info):
                """並列実行用のタスク"""
                h2_text = h2_info['text']
                try:
                    logger.info(f"[PARALLEL] AI画像生成開始 ({index+1}/{len(h2_to_process)}): '{h2_text[:30]}...'")
                    image_bytes = self.generate_image_with_vertex(h2_text, keyword)

                    if image_bytes and not (isinstance(image_bytes, str) and image_bytes.startswith("ERROR:")):
                        timestamp = int(time.time())
                        filename = f"AI_{keyword}_{h2_text[:20]}_{timestamp}"
                        ai_image_url = self.upload_image_to_drive(image_bytes, filename)

                        if ai_image_url:
                            logger.info(f"[PARALLEL] AI画像生成完了 ({index+1}): '{h2_text[:30]}...'")
                            return (index, ai_image_url, None)
                        else:
                            return (index, None, f"H2 '{h2_text}': AI画像アップロード失敗")
                    else:
                        error_msg = image_bytes if isinstance(image_bytes, str) else "生成失敗"
                        return (index, None, f"H2 '{h2_text}': AI画像生成失敗 - {error_msg}")
                except Exception as e:
                    return (index, None, f"H2 '{h2_text}': 例外発生 - {str(e)}")

            # AI画像生成を順次実行（クォータ超過対策）
            ai_image_results = {}
            logger.info(f"[SEQUENTIAL] {len(h2_to_process)}個のAI画像を順次生成開始")
            start_time = time.time()
            quota_exceeded = False  # クォータ超過フラグ

            # 順次処理で1枚ずつ生成（各生成後に30秒待機）
            for i, h2_info in enumerate(h2_to_process):
                # クォータ超過済みなら残りをスキップ
                if quota_exceeded:
                    logger.warning(f"[SEQUENTIAL] クォータ超過のため残り{len(h2_to_process) - i}枚をスキップ")
                    break

                index, ai_url, error = generate_ai_image_task(i, h2_info)
                ai_image_results[index] = ai_url

                if error:
                    image_errors.append(error)
                    logger.warning(f"[SEQUENTIAL] {error}")
                    # クォータ超過エラーの場合、フラグを立てる
                    if "QUOTA_EXCEEDED" in error:
                        quota_exceeded = True
                        logger.warning("[SEQUENTIAL] クォータ超過を検出。以降のAI画像生成をスキップします")
                        continue

                # 次の生成まで30秒待機（クォータ制限対策・強化版）
                if i < len(h2_to_process) - 1 and not quota_exceeded:
                    logger.info(f"[SEQUENTIAL] 次の画像生成まで30秒待機...")
                    time.sleep(30)

            elapsed_time = time.time() - start_time
            logger.info(f"[SEQUENTIAL] AI画像生成完了: {elapsed_time:.1f}秒 (クォータ超過: {quota_exceeded})")

            # 画像挿入リクエストを構築
            insert_requests = []
            for i, h2_info in enumerate(h2_to_process):
                end_index = h2_info['end_index']
                folder_url = folder_assignments.get(i)
                ai_url = ai_image_results.get(i)

                # AI生成画像を先に追加（下に表示される）
                if ai_url:
                    insert_requests.append({
                        'insertText': {
                            'location': {'index': end_index},
                            'text': '\n'
                        }
                    })
                    insert_requests.append({
                        'insertInlineImage': {
                            'location': {'index': end_index},
                            'uri': ai_url,
                            'objectSize': {
                                'height': {'magnitude': 300, 'unit': 'PT'},
                                'width': {'magnitude': 400, 'unit': 'PT'}
                            }
                        }
                    })

                # フォルダ画像を後に追加（上に表示される）
                if folder_url:
                    insert_requests.append({
                        'insertText': {
                            'location': {'index': end_index},
                            'text': '\n'
                        }
                    })
                    insert_requests.append({
                        'insertInlineImage': {
                            'location': {'index': end_index},
                            'uri': folder_url,
                            'objectSize': {
                                'height': {'magnitude': 300, 'unit': 'PT'},
                                'width': {'magnitude': 400, 'unit': 'PT'}
                            }
                        }
                    })

            # リクエストを実行（全て一度に送信）
            if insert_requests:
                insert_requests.reverse()
                logger.info(f"[BOTH] 画像挿入リクエスト数: {len(insert_requests)}")

                # 全リクエストを一度に送信（リトライ処理付き）
                max_retries = 3
                for attempt in range(max_retries):
                    try:
                        self.docs_service.documents().batchUpdate(
                            documentId=document_id,
                            body={'requests': insert_requests}
                        ).execute()
                        logger.info(f"[BOTH] 全{len(insert_requests)}件の画像挿入完了")
                        break
                    except Exception as batch_error:
                        error_str = str(batch_error)
                        if attempt < max_retries - 1:
                            wait_time = 5 * (attempt + 1)  # 5秒、10秒、15秒
                            logger.warning(f"[BOTH] 画像挿入失敗（リトライ {attempt + 1}/{max_retries}、{wait_time}秒待機）: {batch_error}")
                            time.sleep(wait_time)
                        else:
                            logger.error(f"[BOTH] 画像挿入最終失敗: {batch_error}")
                            # フォールバック: 2リクエストずつペアで挿入（改行+画像）
                            logger.info("[BOTH] フォールバック: ペアで画像を挿入します")
                            success_count = 0
                            total_pairs = len(insert_requests) // 2
                            for i in range(0, len(insert_requests), 2):
                                pair = insert_requests[i:i+2]
                                pair_num = i // 2 + 1
                                try:
                                    self.docs_service.documents().batchUpdate(
                                        documentId=document_id,
                                        body={'requests': pair}
                                    ).execute()
                                    success_count += 1
                                    logger.info(f"[BOTH] ペア{pair_num}/{total_pairs}挿入成功")
                                except Exception as single_error:
                                    logger.warning(f"[BOTH] ペア{pair_num}挿入失敗: {single_error}")
                                time.sleep(2)  # ペアごとに2秒待機
                            logger.info(f"[BOTH] フォールバック完了: {success_count}/{total_pairs}ペア成功")
                            if success_count < total_pairs:
                                image_errors.append(f"画像挿入: {success_count}/{total_pairs}ペアのみ成功")

                logger.info(f"[BOTH] 画像挿入完了")

            return image_errors

        except Exception as e:
            logger.error(f"[BOTH] エラー: 両方の画像挿入に失敗 - {e}")
            import traceback
            logger.error(f"詳細: {traceback.format_exc()}")
            return [f"画像処理全体エラー: {str(e)}"]

    def process_single_sheet(self, sheet_name, force=False):
        """指定されたシート1つを処理

        Args:
            sheet_name: シート名
            force: Trueの場合、処理済みでも再生成する
        """
        logger.info(f"[SINGLE] シート '{sheet_name}' の処理を開始 (force={force})")
        self.authenticate_google()

        heading_data = self.get_headings_from_sheet(sheet_name, force=force)

        if not heading_data:
            logger.error(f"[SINGLE] シート '{sheet_name}': 見出しデータなし")
            return {'status': 'error', 'error': '見出しデータなし'}

        h2_count = len([h for h in heading_data['headings'] if h['level'] == 'H2'])
        logger.info(f"[SINGLE] シート '{sheet_name}': H1='{heading_data['h1_title']}', H2数={h2_count}")

        try:
            # 記事生成（初稿）
            article_draft = self.generate_article(
                heading_data['keyword'],
                heading_data['h1_title'],
                heading_data['headings']
            )

            if not article_draft or article_draft.startswith("ERROR:"):
                error_detail = article_draft if article_draft else "Unknown Error"
                logger.error(f"[SINGLE] 記事生成失敗: {error_detail}")
                return {'status': 'error', 'error': error_detail}

            article = article_draft

            # Googleドキュメントに保存
            doc_url, document_id, tables = self.save_to_google_docs(article, heading_data['h1_title'])

            if not doc_url:
                logger.error(f"[SINGLE] ドキュメント保存失敗")
                return {'status': 'error', 'error': 'ドキュメント保存失敗'}

            logger.info(f"[SINGLE] ドキュメント生成成功: {doc_url}")
            self.update_sheet_status(sheet_name, "画像処理中...", doc_url)

            # 画像挿入処理
            h2_headings = [h for h in heading_data['headings'] if h['level'] == 'H2']
            if self.image_generation_method == 'both':
                logger.info("[SINGLE] 両方の画像（フォルダ + Vertex AI）を挿入します")
                try:
                    self.insert_both_images_into_doc(document_id, h2_headings, heading_data['keyword'])
                except Exception as e:
                    logger.error(f"[SINGLE] 両方の画像挿入エラー（継続）: {e}")
            elif self.image_generation_method == 'vertex_ai':
                logger.info("[SINGLE] Vertex AIで画像を生成します")
                self.insert_generated_images_into_doc(document_id, h2_headings, heading_data['keyword'])
            else:
                logger.info("[SINGLE] 既存の画像フォルダから画像を取得します")
                try:
                    self.insert_images_into_doc(document_id, h2_headings)
                except Exception as e:
                    logger.error(f"[SINGLE] 画像挿入エラー（継続）: {e}")

            # 表はマークダウン形式で既に挿入済み（save_to_google_docs内で処理）
            if tables:
                logger.info(f"[SINGLE] 表は既にマークダウン形式で挿入済み: {len(tables)}個")

            # 最終ステータス更新
            self.update_sheet_status(sheet_name, "処理済み", doc_url)

            # マスターシートに初稿URLを書き込む（設定されている場合）
            if self.master_spreadsheet_id:
                self.update_master_sheet_article_url(
                    self.master_spreadsheet_id,
                    heading_data['keyword'],
                    doc_url,
                    self.keyword_column,
                    self.article_url_column
                )

            logger.info(f"[SINGLE] 処理完了: {sheet_name}")

            # Slack通知を送信
            self.send_article_notification(
                title=heading_data['h1_title'],
                url=doc_url,
                keyword=heading_data['keyword']
            )

            return {
                'status': 'success',
                'title': heading_data['h1_title'],
                'url': doc_url
            }

        except Exception as e:
            logger.error(f"[SINGLE] 例外発生: {e}")
            import traceback
            logger.error(f"[SINGLE] トレースバック: {traceback.format_exc()}")
            return {'status': 'error', 'error': str(e)}

    def process_all_sheets(self, max_articles=None):
        """すべての未処理シートを処理"""
        logger.info(f"[DEBUG] 処理開始 - max_articles: {max_articles}")
        self.authenticate_google()

        sheets = self.get_all_sheets()
        logger.info(f"[DEBUG] シート一覧取得: {len(sheets)}個のシート - {sheets}")

        if not sheets:
            logger.error("[DEBUG] シート一覧が空です！")
            return {
                'processed': [],
                'errors': [{'error': 'シート一覧が取得できませんでした'}],
                'skipped': [],
                'total': 0,
                'total_sheets': 0,
                'all_sheets_status': []
            }

        processed = []
        errors = []
        skipped = []  # スキップされたシートの情報を追加
        all_sheets_status = []  # 全シートの状態を記録

        count = 0
        logger.info(f"[DEBUG] ループ開始: {len(sheets)}シートを処理します")

        for sheet_name in sheets:
            logger.info(f"[DEBUG] ループ内: シート '{sheet_name}' を確認中...")
            if max_articles and count >= max_articles:
                logger.info(f"[DEBUG] 最大記事数 {max_articles} に到達。処理終了。")
                break

            heading_data = self.get_headings_from_sheet(sheet_name)

            if not heading_data:
                reason = "見出しデータなし"
                logger.info(f"[DEBUG] シート '{sheet_name}': {reason}。スキップ。")
                skipped.append({
                    'sheet': sheet_name,
                    'reason': reason
                })
                all_sheets_status.append({
                    'sheet': sheet_name,
                    'status': 'skipped',
                    'reason': reason
                })
                continue

            h2_count = len([h for h in heading_data['headings'] if h['level'] == 'H2'])
            logger.info(f"[DEBUG] シート '{sheet_name}': 処理対象として検出。H1='{heading_data['h1_title']}', 見出し総数={len(heading_data['headings'])}, H2数={h2_count}")

            try:
                logger.info(f"処理中: {heading_data['h1_title']}")

                # 記事生成（初稿）
                article_draft = self.generate_article(
                    heading_data['keyword'],
                    heading_data['h1_title'],
                    heading_data['headings']
                )

                if not article_draft or article_draft.startswith("ERROR:"):
                    # 初稿生成失敗
                    error_detail = article_draft if article_draft else "Unknown Error"
                    error_msg = f"記事生成に失敗しました: {error_detail}"
                    errors.append({'sheet': sheet_name, 'error': error_msg})
                    all_sheets_status.append({'sheet': sheet_name, 'status': 'error', 'error': error_msg})
                    continue

                # generate_articleで既にStep0〜Step4の処理済み
                article = article_draft

                if article:
                    # Googleドキュメントに保存
                    doc_url, document_id, tables = self.save_to_google_docs(article, heading_data['h1_title'])

                    if doc_url:
                        # 【重要】URL生成直後にスプレッドシートに書き込む（最優先）
                        logger.info(f"ドキュメント生成成功: {doc_url}")
                        self.update_sheet_status(sheet_name, "画像処理中...", doc_url)

                        # 画像生成方法に応じて処理を切り替え
                        h2_headings = [h for h in heading_data['headings'] if h['level'] == 'H2']
                        if self.image_generation_method == 'both':
                            # 両方の画像（フォルダ + Vertex AI）を挿入
                            logger.info("両方の画像（フォルダ + Vertex AI）を挿入します")
                            try:
                                self.insert_both_images_into_doc(document_id, h2_headings, heading_data['keyword'])
                            except Exception as e:
                                logger.error(f"両方の画像挿入中にエラーが発生しましたが、処理を継続します: {e}")
                        elif self.image_generation_method == 'vertex_ai':
                            # Vertex AIで画像を生成して挿入
                            logger.info("Vertex AIで画像を生成します")
                            img_errors = self.insert_generated_images_into_doc(document_id, h2_headings, heading_data['keyword'])
                            if img_errors:
                                error_msg = "; ".join(img_errors)
                                all_sheets_status.append({'sheet': sheet_name, 'status': 'warning', 'error': f"画像生成エラー: {error_msg}"})
                        else:
                            # 既存の画像フォルダから取得して挿入
                            logger.info("既存の画像フォルダから画像を取得します")
                            # 画像挿入処理（エラーが出ても停止しないようにtry-exceptで囲む手もあるが、insert_images_into_doc内でログが出ているはず）
                            try:
                                self.insert_images_into_doc(document_id, h2_headings)
                            except Exception as e:
                                logger.error(f"画像挿入中にエラーが発生しましたが、処理を継続します: {e}")

                        # 表はマークダウン形式で既に挿入済み（save_to_google_docs内で処理）
                        if tables:
                            logger.info(f"表は既にマークダウン形式で挿入済み: {len(tables)}個")

                        # 最終ステータス更新
                        self.update_sheet_status(sheet_name, "処理済み", doc_url)

                        # マスターシートに初稿URLを書き込む（設定されている場合）
                        if self.master_spreadsheet_id:
                            self.update_master_sheet_article_url(
                                self.master_spreadsheet_id,
                                heading_data['keyword'],
                                doc_url,
                                self.keyword_column,
                                self.article_url_column
                            )

                        # Slack通知を送信
                        self.send_article_notification(
                            title=heading_data['h1_title'],
                            url=doc_url,
                            keyword=heading_data['keyword']
                        )

                        processed.append({
                            'sheet': sheet_name,
                            'title': heading_data['h1_title'],
                            'url': doc_url
                        })
                        all_sheets_status.append({
                            'sheet': sheet_name,
                            'status': 'processed',
                            'url': doc_url
                        })
                        count += 1
                    else:
                        # ドキュメント保存失敗
                        error_msg = "ドキュメント保存に失敗しました"
                        errors.append({'sheet': sheet_name, 'error': error_msg})
                        all_sheets_status.append({'sheet': sheet_name, 'status': 'error', 'error': error_msg})
                # else句は不要（初稿生成失敗は上で処理済み）

            except Exception as e:
                error_msg = str(e)
                errors.append({
                    'sheet': sheet_name,
                    'error': error_msg
                })
                all_sheets_status.append({
                    'sheet': sheet_name,
                    'status': 'error',
                    'error': error_msg
                })

        return {
            'processed': processed,
            'errors': errors,
            'skipped': skipped,  # スキップ情報を追加
            'total': count,
            'total_sheets': len(sheets),  # 全シート数を追加
            'all_sheets_status': all_sheets_status  # 全シートの詳細状態
        }

    def get_unprocessed_sheets(self):
        """未処理のシートのみを取得"""
        self.authenticate_google()
        all_sheets = self.get_all_sheets()
        unprocessed = []

        for sheet_name in all_sheets:
            heading_data = self.get_headings_from_sheet(sheet_name)
            if heading_data:  # 処理済みでない場合のみheading_dataが返る
                unprocessed.append({
                    'sheet_name': sheet_name,
                    'keyword': heading_data.get('keyword', ''),
                    'h1_title': heading_data.get('h1_title', '')
                })

        logger.info(f"[PARALLEL] 未処理シート: {len(unprocessed)}件")
        return unprocessed

    def send_article_notification(self, title, url, keyword=""):
        """1件の記事生成完了時にSlack通知"""
        slack_webhook_url = os.environ.get('SLACK_WEBHOOK_URL')
        if not slack_webhook_url:
            return

        try:
            message = f"""✅ *初稿生成完了*

*{title}*
{url}
"""
            if keyword:
                message += f"\nキーワード: {keyword}"

            payload = {'text': message}
            response = requests.post(slack_webhook_url, json=payload, timeout=10)
            response.raise_for_status()
            logger.info(f"[SLACK] 記事通知送信: {title}")
        except Exception as e:
            logger.error(f"[SLACK] 記事通知エラー: {e}")

    def enqueue_articles_to_cloud_tasks(self, cloud_run_url):
        """未処理の全記事をCloud Tasksにキュー登録"""
        unprocessed = self.get_unprocessed_sheets()

        if not unprocessed:
            logger.info("[QUEUE] 未処理のシートがありません")
            return {
                'status': 'no_work',
                'message': '未処理のシートがありません',
                'queued': 0
            }

        total = len(unprocessed)
        logger.info(f"[QUEUE] {total}件をCloud Tasksにキュー登録します")

        # Cloud Tasks クライアント
        client = tasks_v2.CloudTasksClient()

        # キューのパス
        project = 'YOUR_GCP_PROJECT_ID'
        location = 'asia-northeast1'
        queue = 'article-generation-queue'
        parent = client.queue_path(project, location, queue)

        queued_count = 0

        for i, sheet in enumerate(unprocessed):
            try:
                # タスクのペイロード
                payload = {
                    'spreadsheet_id': self.spreadsheet_id,
                    'sheet_name': sheet['sheet_name'],
                    'image_generation_method': self.image_generation_method,
                    'master_spreadsheet_id': self.master_spreadsheet_id,
                    'keyword_column': self.keyword_column,
                    'article_url_column': self.article_url_column,
                    'task_index': i + 1,
                    'total_tasks': total
                }

                # Cloud Runエンドポイント
                url = f"{cloud_run_url}/process-article-task"

                # タスクを作成
                task = {
                    'http_request': {
                        'http_method': tasks_v2.HttpMethod.POST,
                        'url': url,
                        'headers': {'Content-Type': 'application/json'},
                        'body': json.dumps(payload).encode()
                    }
                }

                # キューに追加（5秒間隔でスケジュール）
                schedule_time = datetime.datetime.utcnow() + datetime.timedelta(seconds=i * 5)
                timestamp = timestamp_pb2.Timestamp()
                timestamp.FromDatetime(schedule_time)
                task['schedule_time'] = timestamp

                response = client.create_task(parent=parent, task=task)
                queued_count += 1
                logger.info(f"[QUEUE] タスク登録: {sheet['sheet_name']} ({i+1}/{total})")

            except Exception as e:
                logger.error(f"[QUEUE] タスク登録エラー: {sheet['sheet_name']} - {e}")

        # 開始通知
        self.send_batch_start_notification(total, queued_count)

        return {
            'status': 'queued',
            'total': total,
            'queued': queued_count,
            'message': f'{queued_count}件のタスクをキューに登録しました'
        }

    def send_batch_start_notification(self, total, queued):
        """一括生成開始時のSlack通知"""
        slack_webhook_url = os.environ.get('SLACK_WEBHOOK_URL')
        if not slack_webhook_url:
            return

        try:
            # 予想時間を計算（1件約5分）
            estimated_minutes = queued * 5
            estimated_hours = estimated_minutes // 60
            estimated_mins = estimated_minutes % 60

            if estimated_hours > 0:
                time_str = f"約{estimated_hours}時間{estimated_mins}分"
            else:
                time_str = f"約{estimated_mins}分"

            message = f"""🚀 *記事一括生成を開始しました*

• 対象記事数: {queued}件
• 予想所要時間: {time_str}
• 同時処理数: 3件

各記事の完了時に個別通知します。"""

            payload = {'text': message}
            response = requests.post(slack_webhook_url, json=payload, timeout=10)
            response.raise_for_status()
            logger.info("[QUEUE] 開始通知を送信しました")
        except Exception as e:
            logger.error(f"[QUEUE] 開始通知エラー: {e}")


class SearchConsoleKeywordFetcher:
    """Search Console APIからキーワードデータを取得してスプレッドシートに書き込む"""

    def __init__(self, site_url, spreadsheet_id):
        self.site_url = site_url
        self.spreadsheet_id = spreadsheet_id
        self.search_console_service = None
        self.sheets_service = None

    def authenticate_google(self):
        """サービスアカウントで認証"""
        service_account_key = os.environ.get('GOOGLE_SERVICE_ACCOUNT_KEY')

        if not service_account_key:
            raise ValueError("GOOGLE_SERVICE_ACCOUNT_KEY environment variable is not set")

        try:
            service_account_info = json.loads(service_account_key)
            logger.info(f"認証情報を読み込みました: {service_account_info.get('client_email', 'N/A')}")
        except json.JSONDecodeError as e:
            raise ValueError(f"Invalid JSON in GOOGLE_SERVICE_ACCOUNT_KEY: {e}")

        credentials = service_account.Credentials.from_service_account_info(
            service_account_info,
            scopes=SCOPES
        )

        self.search_console_service = build('searchconsole', 'v1', credentials=credentials)
        self.sheets_service = build('sheets', 'v4', credentials=credentials)
        logger.info("認証成功")

    def fetch_keywords(self, days=30, row_limit=1000):
        """Search Console APIからキーワードデータを取得"""
        try:
            from datetime import datetime, timedelta

            # 日付範囲を計算（Search Console APIは3日前までのデータしか取得できない）
            end_date = datetime.now() - timedelta(days=3)
            start_date = end_date - timedelta(days=days)

            logger.info(f"データ取得期間: {start_date.date()} ～ {end_date.date()}")

            # Search Console APIリクエスト
            request_body = {
                'startDate': start_date.strftime('%Y-%m-%d'),
                'endDate': end_date.strftime('%Y-%m-%d'),
                'dimensions': ['query'],
                'rowLimit': row_limit,
                'startRow': 0
            }

            logger.info(f"Search Console APIにリクエスト送信中... (site: {self.site_url})")
            response = self.search_console_service.searchanalytics().query(
                siteUrl=self.site_url,
                body=request_body
            ).execute()

            rows = response.get('rows', [])
            logger.info(f"✓ {len(rows)}件のキーワードデータを取得しました")

            # データ整形
            keywords_data = []
            for row in rows:
                keywords_data.append({
                    'keyword': row['keys'][0],
                    'clicks': int(row.get('clicks', 0)),
                    'impressions': int(row.get('impressions', 0)),
                    'ctr': round(row.get('ctr', 0) * 100, 2),
                    'position': round(row.get('position', 0), 1)
                })

            return keywords_data

        except HttpError as err:
            logger.error(f"Search Console APIエラー: {err}")
            logger.error(f"詳細: {err.content}")
            raise

    def write_to_spreadsheet(self, keywords_data, sheet_name="キーワード分析"):
        """スプレッドシートにキーワードデータを書き込む"""
        try:
            # シートが存在するか確認、なければ作成
            sheet_metadata = self.sheets_service.spreadsheets().get(
                spreadsheetId=self.spreadsheet_id
            ).execute()

            sheets = sheet_metadata.get('sheets', [])
            sheet_exists = any(sheet['properties']['title'] == sheet_name for sheet in sheets)

            if not sheet_exists:
                logger.info(f"シート '{sheet_name}' を作成中...")
                batch_update_request = {
                    'requests': [{
                        'addSheet': {
                            'properties': {
                                'title': sheet_name
                            }
                        }
                    }]
                }
                self.sheets_service.spreadsheets().batchUpdate(
                    spreadsheetId=self.spreadsheet_id,
                    body=batch_update_request
                ).execute()
                logger.info(f"✓ シート '{sheet_name}' を作成しました")
            else:
                # 既存のシートをクリア
                logger.info(f"シート '{sheet_name}' をクリア中...")
                self.sheets_service.spreadsheets().values().clear(
                    spreadsheetId=self.spreadsheet_id,
                    range=f"'{sheet_name}'!A1:Z1000"
                ).execute()

            # ヘッダー行 + データ行
            from datetime import datetime
            values = [
                ['取得日時', datetime.now().strftime('%Y/%m/%d %H:%M:%S')],
                [],
                ['キーワード', 'クリック数', '表示回数', 'CTR (%)', '平均掲載順位']
            ]

            for kw in keywords_data:
                values.append([
                    kw['keyword'],
                    kw['clicks'],
                    kw['impressions'],
                    kw['ctr'],
                    kw['position']
                ])

            # データ書き込み
            range_name = f"'{sheet_name}'!A1"
            logger.info(f"スプレッドシートにデータ書き込み中... ({len(keywords_data)}件)")

            self.sheets_service.spreadsheets().values().update(
                spreadsheetId=self.spreadsheet_id,
                range=range_name,
                valueInputOption='RAW',
                body={'values': values}
            ).execute()

            logger.info(f"✓ スプレッドシートへの書き込み完了")

        except HttpError as err:
            logger.error(f"スプレッドシート書き込みエラー: {err}")
            raise

    def run(self, days=30, row_limit=1000):
        """キーワード取得からスプレッドシート書き込みまで一括実行"""
        logger.info("=" * 50)
        logger.info("Search Console キーワード取得開始")
        logger.info("=" * 50)

        # 認証
        self.authenticate_google()

        # キーワード取得
        keywords_data = self.fetch_keywords(days=days, row_limit=row_limit)

        if not keywords_data:
            logger.warning("取得できたキーワードがありません")
            return {
                'success': False,
                'message': 'キーワードデータが取得できませんでした',
                'count': 0
            }

        # スプレッドシートに書き込み
        self.write_to_spreadsheet(keywords_data)

        logger.info("=" * 50)
        logger.info("完了")
        logger.info("=" * 50)

        return {
            'success': True,
            'message': f'{len(keywords_data)}件のキーワードを取得しました',
            'count': len(keywords_data),
            'spreadsheet_url': f'https://docs.google.com/spreadsheets/d/{self.spreadsheet_id}'
        }


class OutlineGenerator:
    """キーワードから構成案を生成してスプレッドシートに書き込む"""

    def __init__(self, spreadsheet_id, openai_api_key, custom_search_api_key=None, custom_search_cx=None, anthropic_api_key=None):
        self.spreadsheet_id = spreadsheet_id
        self.openai_api_key = openai_api_key
        self.custom_search_api_key = custom_search_api_key or os.environ.get('GOOGLE_CUSTOM_SEARCH_API_KEY')
        self.custom_search_cx = custom_search_cx or os.environ.get('GOOGLE_CUSTOM_SEARCH_CX')
        self.sheets_service = None
        # OpenAI クライアントを初期化
        self.openai_client = OpenAI(
            api_key=openai_api_key
        )
        # Claude クライアントを初期化（APIキーがあれば）
        self.anthropic_api_key = anthropic_api_key or os.environ.get('ANTHROPIC_API_KEY')
        self.claude_client = None
        if self.anthropic_api_key:
            self.claude_client = anthropic.Anthropic(api_key=self.anthropic_api_key)

    def authenticate_google(self):
        """サービスアカウントで認証"""
        service_account_key = os.environ.get('GOOGLE_SERVICE_ACCOUNT_KEY')

        if not service_account_key:
            raise ValueError("GOOGLE_SERVICE_ACCOUNT_KEY environment variable is not set")

        try:
            service_account_info = json.loads(service_account_key)
            logger.info(f"認証情報を読み込みました: {service_account_info.get('client_email', 'N/A')}")
        except json.JSONDecodeError as e:
            raise ValueError(f"Invalid JSON in GOOGLE_SERVICE_ACCOUNT_KEY: {e}")

        credentials = service_account.Credentials.from_service_account_info(
            service_account_info,
            scopes=SCOPES
        )

        self.sheets_service = build('sheets', 'v4', credentials=credentials)
        self.drive_service = build('drive', 'v3', credentials=credentials)
        logger.info("認証成功")

    def get_or_create_monthly_spreadsheet(self, year, month, folder_id=None):
        """月別スプレッドシートを取得または作成（テンプレートからコピー）"""
        spreadsheet_name = f"{year}年{month}月"
        folder_id = folder_id or os.environ.get('OUTLINE_FOLDER_ID')
        template_id = os.environ.get('TEMPLATE_SPREADSHEET_ID')

        if not folder_id:
            raise ValueError("OUTLINE_FOLDER_ID environment variable is not set")

        # フォルダ内で既存のスプレッドシートを検索
        query = f"name = '{spreadsheet_name}' and mimeType = 'application/vnd.google-apps.spreadsheet' and '{folder_id}' in parents and trashed = false"

        try:
            results = self.drive_service.files().list(
                q=query,
                spaces='drive',
                fields='files(id, name)',
                supportsAllDrives=True,
                includeItemsFromAllDrives=True
            ).execute()

            files = results.get('files', [])

            if files:
                # 既存のスプレッドシートを使用
                spreadsheet_id = files[0]['id']
                logger.info(f"✓ 既存の月別スプレッドシート「{spreadsheet_name}」を使用: {spreadsheet_id}")
                return spreadsheet_id

            # 新規作成（テンプレートからコピー or 空で作成）
            if template_id:
                logger.info(f"テンプレートから月別スプレッドシート「{spreadsheet_name}」を作成中...")

                # テンプレートをコピー
                copied_file = self.drive_service.files().copy(
                    fileId=template_id,
                    body={
                        'name': spreadsheet_name,
                        'parents': [folder_id]
                    },
                    supportsAllDrives=True,
                    fields='id'
                ).execute()

                spreadsheet_id = copied_file.get('id')
                logger.info(f"✓ テンプレートから月別スプレッドシート「{spreadsheet_name}」を作成: {spreadsheet_id}")
            else:
                logger.info(f"月別スプレッドシート「{spreadsheet_name}」を新規作成中...")

                file_metadata = {
                    'name': spreadsheet_name,
                    'mimeType': 'application/vnd.google-apps.spreadsheet',
                    'parents': [folder_id]
                }

                spreadsheet = self.drive_service.files().create(
                    body=file_metadata,
                    supportsAllDrives=True,
                    fields='id'
                ).execute()

                spreadsheet_id = spreadsheet.get('id')
                logger.info(f"✓ 月別スプレッドシート「{spreadsheet_name}」を作成: {spreadsheet_id}")

            return spreadsheet_id

        except Exception as e:
            logger.error(f"エラー: 月別スプレッドシートの取得/作成に失敗 - {e}")
            raise

    def generate_related_keywords(self, keyword):
        """メインキーワードに関連するキーワード10個を生成"""
        try:
            prompt = f"""# 関連キーワード生成タスク

メインキーワードと一緒に検索されやすい関連キーワードを10個提案してください。

## メインキーワード
{keyword}

## 条件
- ✅ 検索ボリュームがありそうなキーワード
- ✅ ユーザーの検索意図に沿ったキーワード
- ✅ メインキーワードと組み合わせやすいキーワード
- ✅ 記事の見出し（H2/H3）に使えるキーワード

## 出力形式
- キーワードのみを1行に1つずつ出力
- 10個まで
- 説明文は不要

例:
キーワード1
キーワード2
キーワード3
"""

            response = self.openai_client.chat.completions.create(
                model="gpt-4.1-mini",
                messages=[
                    {"role": "system", "content": "あなたはSEOキーワードリサーチの専門家です。"},
                    {"role": "user", "content": prompt}
                ],
                max_completion_tokens=300
            )

            keywords_text = response.choices[0].message.content.strip()
            # 改行で分割してリストに変換
            related_keywords = [kw.strip() for kw in keywords_text.split('\n') if kw.strip()]
            # 10個に制限
            related_keywords = related_keywords[:10]

            logger.info(f"✓ キーワード「{keyword}」の関連キーワード{len(related_keywords)}個を生成しました")
            return related_keywords

        except Exception as e:
            logger.error(f"エラー: 関連キーワード生成に失敗 - {e}")
            return []

    def extract_cooccurrence_keywords(self, keyword, num_urls=20, min_df=2, top_n=30):
        """上位ページから共起語を抽出（TF-DF分析）

        Args:
            keyword: 検索キーワード
            num_urls: 分析する上位ページ数（デフォルト20）
            min_df: 最低出現ページ数（デフォルト2）
            top_n: 返す共起語の数（デフォルト30）

        Returns:
            共起語リスト（出現頻度順）
        """
        try:
            logger.info(f"共起語抽出開始: キーワード「{keyword}」")

            # 1. 上位URLを取得
            urls = self.fetch_top_urls(keyword, num_results=num_urls)
            if not urls:
                logger.warning("URLを取得できませんでした。GPT生成にフォールバック")
                return self.generate_related_keywords(keyword)

            # 2. 各URLから本文を取得
            all_doc_words = []  # 各ドキュメントの単語セット
            tokenizer = Tokenizer()

            # 除外する単語（一般的すぎる語、記号など）
            stop_words = {
                'こと', 'もの', 'ため', 'よう', 'それ', 'これ', 'ここ', 'そこ',
                'の', 'に', 'は', 'を', 'が', 'と', 'で', 'て', 'から', 'まで',
                '年', '月', '日', '時', '分', '人', '方', '的', '性', '化', '中',
                '等', '用', '上', '下', '内', '外', '間', '前', '後', '以', '円',
                '万', '億', '件', '回', '個', '本', '点', '度', '部', '名', '数',
                'さん', 'とき', 'ところ', 'あと', 'まま', 'ほう', 'わけ', 'はず',
                'つもり', 'みたい', 'やつ', 'ひと', 'なか', 'うち', 'そば', 'へん',
                'あたり', 'まわり', 'あいだ', 'ぶん', 'ほか', 'べつ', 'ほど'
            }

            # メインキーワードを分解して除外リストに追加
            for token in tokenizer.tokenize(keyword):
                stop_words.add(token.surface)

            for url in urls:
                try:
                    content = self.fetch_article_content(url)
                    if not content or not content.get('body'):
                        continue

                    body_text = content['body']
                    doc_words = set()

                    # 形態素解析で名詞を抽出
                    for token in tokenizer.tokenize(body_text):
                        pos = token.part_of_speech.split(',')[0]
                        surface = token.surface

                        # 名詞のみ抽出（ただし数詞、非自立、接尾は除外）
                        if pos == '名詞':
                            sub_pos = token.part_of_speech.split(',')[1] if len(token.part_of_speech.split(',')) > 1 else ''
                            if sub_pos in ['数', '非自立', '接尾']:
                                continue
                            # 1文字の語、数字のみ、stop_wordsは除外
                            if len(surface) < 2:
                                continue
                            if surface.isdigit():
                                continue
                            if surface in stop_words:
                                continue
                            doc_words.add(surface)

                    all_doc_words.append(doc_words)
                    logger.info(f"  - {url[:50]}... から {len(doc_words)} 語抽出")

                except Exception as e:
                    logger.warning(f"URL処理エラー: {url[:50]}... - {e}")
                    continue

            if not all_doc_words:
                logger.warning("本文を取得できませんでした。GPT生成にフォールバック")
                return self.generate_related_keywords(keyword)

            # 3. DF（Document Frequency）を計算
            df_count = Counter()
            for doc_words in all_doc_words:
                for word in doc_words:
                    df_count[word] += 1

            # 4. TF（Term Frequency）を計算（全ドキュメント合計）
            tf_count = Counter()
            for doc_words in all_doc_words:
                tf_count.update(doc_words)

            # 5. DF >= min_df でフィルタ
            filtered_words = {word for word, df in df_count.items() if df >= min_df}

            # 6. TF-DFスコアで並び替え（DF * TF）
            scored_words = []
            for word in filtered_words:
                score = df_count[word] * tf_count[word]
                scored_words.append((word, score, df_count[word]))

            scored_words.sort(key=lambda x: x[1], reverse=True)

            # 7. 上位N語を返す
            result = [word for word, score, df in scored_words[:top_n]]

            logger.info(f"✓ 共起語抽出完了: {len(result)}語（{len(all_doc_words)}ページから）")
            logger.info(f"  上位10語: {', '.join(result[:10])}")

            return result

        except Exception as e:
            logger.error(f"エラー: 共起語抽出に失敗 - {e}")
            # フォールバック: GPT生成
            return self.generate_related_keywords(keyword)

    def fetch_top_urls(self, keyword, num_results=10):
        """Google Custom Search JSON APIで上位URLを取得

        Args:
            keyword: 検索キーワード
            num_results: 取得件数（最大20件、10件ごとにAPIリクエスト）
        """
        if not self.custom_search_api_key or not self.custom_search_cx:
            logger.warning("Google Custom Search APIキーまたは検索エンジンIDが設定されていません")
            return []

        try:
            urls = []
            url = "https://www.googleapis.com/customsearch/v1"

            # 10件ずつ取得（Custom Search APIは1リクエスト最大10件）
            for start in range(1, min(num_results + 1, 21), 10):
                params = {
                    'key': self.custom_search_api_key,
                    'cx': self.custom_search_cx,
                    'q': keyword,
                    'num': min(10, num_results - len(urls)),
                    'start': start
                }

                response = requests.get(url, params=params, timeout=10)
                response.raise_for_status()
                data = response.json()

                for item in data.get('items', []):
                    urls.append(item.get('link'))
                    if len(urls) >= num_results:
                        break

                if len(urls) >= num_results:
                    break

            logger.info(f"✓ キーワード「{keyword}」の上位{len(urls)}件のURLを取得しました")
            return urls

        except requests.exceptions.RequestException as e:
            logger.error(f"エラー: Google Custom Search APIリクエスト失敗 - {e}")
            return []
        except Exception as e:
            logger.error(f"エラー: URL取得に失敗 - {e}")
            return []

    def fetch_article_content(self, url):
        """URLから記事の見出し構造とメインコンテンツを抽出"""
        try:
            headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
            }
            response = requests.get(url, headers=headers, timeout=10)
            response.raise_for_status()
            response.encoding = response.apparent_encoding

            soup = BeautifulSoup(response.text, 'html.parser')

            # 不要な要素を削除
            for tag in soup.find_all(['script', 'style', 'nav', 'header', 'footer', 'aside', 'form', 'iframe']):
                tag.decompose()

            # タイトル（H1）を取得
            h1 = soup.find('h1')
            title = h1.get_text(strip=True) if h1 else ""

            # 見出し構造を取得（H2, H3）
            headings = []
            for tag in soup.find_all(['h2', 'h3']):
                text = tag.get_text(strip=True)
                if text and len(text) < 100:  # 長すぎる見出しは除外
                    headings.append(f"{tag.name.upper()}: {text}")

            # メタディスクリプションを取得
            meta_desc = soup.find('meta', attrs={'name': 'description'})
            description = meta_desc.get('content', '') if meta_desc else ""

            # 本文を取得（article, main, またはbodyから）
            body_content = ""
            main_content = soup.find('article') or soup.find('main') or soup.find('body')
            if main_content:
                # 段落テキストを収集
                paragraphs = main_content.find_all(['p', 'li'])
                body_texts = []
                for p in paragraphs:
                    text = p.get_text(strip=True)
                    if text and len(text) > 20:  # 短すぎるテキストは除外
                        body_texts.append(text)
                body_content = '\n'.join(body_texts)  # 全段落取得

            return {
                'url': url,
                'title': title,
                'description': description[:200] if description else "",
                'headings': headings[:20],  # 最大20個の見出し
                'body': body_content  # 本文を追加
            }

        except Exception as e:
            logger.warning(f"記事取得失敗 ({url}): {e}")
            return {'url': url, 'title': '', 'description': '', 'headings': [], 'body': ''}

    def fetch_top_articles(self, keyword):
        """上位3記事のURLと内容を取得（5件取得して有効な上位3件を返す）"""
        urls = self.fetch_top_urls(keyword)
        if not urls:
            return []

        # 並列で記事内容を取得（順序を保持）
        articles_dict = {}
        with ThreadPoolExecutor(max_workers=10) as executor:
            future_to_url = {executor.submit(self.fetch_article_content, url): url for url in urls}
            for future in as_completed(future_to_url):
                url = future_to_url[future]
                try:
                    article = future.result()
                    if article['title'] or article['headings']:
                        articles_dict[url] = article
                except Exception as e:
                    logger.warning(f"記事内容取得エラー: {e}")

        # 検索順位順に並べて上位3件を取得
        articles = []
        for url in urls:
            if url in articles_dict:
                articles.append(articles_dict[url])
                if len(articles) >= 3:
                    break

        if len(articles) < 3:
            logger.warning(f"⚠ キーワード「{keyword}」: 3件中{len(articles)}件のみ取得（10件中{len(articles_dict)}件が有効）")
        else:
            logger.info(f"✓ キーワード「{keyword}」の上位3件の記事内容を取得しました")
        return articles


    def generate_outline_for_keyword(self, keyword):
        """1つのキーワードに対して構成案を生成（共起語 + 上位URL込み）"""
        try:
            # ステップ1: 共起語（TF-DF分析）と上位URLを並列取得
            with ThreadPoolExecutor(max_workers=2) as executor:
                future_keywords = executor.submit(self.extract_cooccurrence_keywords, keyword)
                future_articles = executor.submit(self.fetch_top_articles, keyword)

                related_keywords = future_keywords.result()
                top_articles = future_articles.result()

            related_keywords_text = '\n'.join([f"- {kw}" for kw in related_keywords]) if related_keywords else "（なし）"
            # 上位記事の見出し構造と本文をフォーマット
            top_articles_text = ""
            top_urls = []
            for i, article in enumerate(top_articles, 1):
                top_urls.append(article['url'])
                headings_text = '\n'.join([f"    {h}" for h in article['headings'][:15]]) if article['headings'] else "    （見出し取得不可）"
                body_text = article.get('body', '') if article.get('body') else "（本文取得不可）"
                top_articles_text += f"""
【記事{i}】
- URL: {article['url']}
- タイトル: {article['title']}
- 概要: {article['description']}
- 見出し構造:
{headings_text}
- 本文（抜粋）:
{body_text}
"""
            if not top_articles_text:
                top_articles_text = "（上位記事の取得に失敗）"

            # ステップ2: メインキーワード + 共起語 + 上位URLで構成案を生成
            prompt = f"""# 記事構成案作成タスク

検索意図を満たす包括的な記事構成案を作成してください。

---

## 基本情報

### メインキーワード
{keyword}

### 共起語（上位ページから抽出・構成案に自然に組み込む）
{related_keywords_text}

### 検索上位記事の分析（重要：この見出し構造を参考に構成案を作成）
{top_articles_text}

## 【最重要】記事仕様

### 文字数（絶対厳守）
- ✅ 記事全体で **5500字（5,000～6,000字）** になる構成
- ❌ この範囲を外れる構成は不可

### 見出し構造（絶対厳守）
- ✅ **H2: 6～8個**（最後の「まとめ」を含む）
- ✅ **各H2（まとめを除く）: H3を2〜3個含める**
- ✅ **H4: ランキングやカテゴリ別展開が必要な場合に使用**
- ❌ H3が1つ以下は不可

### 文字数配分の目安
| セクション | 文字数 |
|-----------|--------|
| 導入部（H1直後） | 200-300字 |
| 各H2セクション | 500-600字（H3含む） |
| まとめ（最後のH2） | 400-500字 |

### 文字数調整の計算式
- H2が6個: 導入300字 + H2×6（各700字） + まとめ500字 = 5,500字
- H2が7個: 導入300字 + H2×7（各650字） + まとめ500字 = 5,500字
- H2が8個: 導入300字 + H2×8（各600字） + まとめ500字 = 5,500字

---

## 構成作成の指示

### 読者目線の構成（最重要）
記事を読む人が「知りたい順番」で情報を配置する：

1. **自然な流れで構成**: 基本→詳細→応用の順で読者が理解しやすい流れ
2. **基本→応用の流れ**: 初心者でも理解できる順序で構成
3. **見出しだけで内容がわかる**: H2を読むだけで記事全体が把握できる
4. **1つのH2に1つのテーマ**: 情報を詰め込みすぎない

**構成の基本パターン（参考）**:
```
H2: [〇〇とは？]（基本的な説明・定義）
H2: [基本知識]（前提となる情報）
H2: [具体的な方法・手順]
H2: [メリット・デメリット]
H2: [注意点・ポイント]
H2: [おすすめ・比較]（該当する場合）
H2: [よくある質問]（該当する場合）
H2: まとめ
```

### 必須事項
1. ✅ **H1はメインキーワードをそのまま使わず、自然な日本語の文章にする**
   - ❌ 悪い例：「履歴書 書き方 パート完全ガイド」
   - ✅ 良い例：「パートの履歴書の書き方完全ガイド」

### H1の読みやすさルール（重要）
- **1つのH1に1つのメッセージ**（詰め込みすぎない）
- **「〇〇と△△と□□」のような並列は避ける**（1つに絞る）
- **シンプルで一目でわかる表現**にする
- ❌ 悪い例：「看護師がムリなく副業で収入アップするはじめ方と選び方のコツ」
  - 問題：「はじめ方」「選び方」「コツ」と詰め込みすぎ、文構造が複雑
- ✅ 良い例：「看護師の副業おすすめ10選｜資格を活かして無理なく稼ぐ」
- ✅ 良い例：「看護師が副業で稼ぐには？始め方とおすすめの仕事を紹介」
- ポイント：読者が**一瞬で内容を理解できる**かどうか
2. ✅ 共起語をH2/H3の見出しに自然に組み込む
3. ✅ **上位記事の見出し構造を分析し、共通するトピックは必ず含める**
4. ✅ **上位記事にない独自の視点・情報も1-2個追加して差別化**
5. ✅ 最後のH2「まとめ」にはH3を付けない
6. ✅ **H2とH3は論理的な親子関係を持たせる**（H3はH2を詳細化）
7. ✅ **H3同士は同じ粒度・レベル感で揃える**（バラつきNG）
8. ✅ **H2同士の内容が重複しないようにする**（各H2は独立したテーマ）
   - ❌ 悪い例：「副業のメリット」と「副業をするとこんな良いことがある」（内容重複）
   - ✅ 良い例：「副業のメリット」と「副業を始める前の注意点」（テーマが別）

### 「働く」の平仮名統一（絶対厳守）
- ❌ 「働く」「働き」「働け」「働こう」「働いて」
- ✅ 「はたらく」「はたらき」「はたらけ」「はたらこう」「はたらいて」
- 見出しでも本文でも必ず平仮名で表記

### 禁止コンテンツ（絶対厳守）
- ❌ **副業がバレない方法・隠す方法**を示唆する見出しや内容
  - ❌「会社にバレずに副業する方法」
  - ❌「副業がバレないためのポイント」
  - ❌「確定申告でバレない方法」
- ✅ 代わりに「会社の就業規則を確認する」「届出が必要な場合もある」など正しい情報を

### Q&Aセクション（基本入れる）
- **基本的に全記事に入れる**（まとめの直前に配置）
- 読者が持ちそうな疑問をQ&A形式でまとめる
- **Q&Aは4つ以上**作成する（4〜6個が目安）
- H2の例：「〇〇に関するよくある質問」「Q&A：〇〇でよくある疑問」
- H3は「Q1：〜？」「Q2：〜？」「Q3：〜？」「Q4：〜？」と番号付きで記載
  - ✅「Q1：着ぐるみバイトはどれくらい暑い？」
  - ✅「Q2：未経験でもできる？」
  - ✅「Q3：時給はどれくらい？」
  - ✅「Q4：どうやって探せばいい？」

### クライアント企業紹介（必須・さりげなく自然に）
- **必ず入れる**が、**さりげなく・自然に**組み込む
- ✅ クライアント企業の名前は出してOK（ただし自然な文脈で）
- ❌ 露骨な宣伝・前面に出しすぎる紹介はNG
- ❌ Q&Aセクション内には入れない
- ✅ 記事の文脈に沿って、読者の役に立つ情報として紹介
- Q&Aがある場合はQ&Aの前に、ない場合はまとめの前に入れる

**さりげない入れ方のパターン**：

1. **バイト・仕事探し系の記事** → 探し方H2内のH3として
   - H3「スキマ時間で探すなら「クライアント企業」も便利」
   - H3「クライアント企業で気軽に始める」

2. **お金・節約・生活費系の記事** → 副収入H2内のH3として
   - H3「空いた時間にできる単発バイトで収入を補う」
   - H3「「クライアント企業」などのアプリで無理なく稼ぐ」

3. **働き方・キャリア系の記事** → はたらき方H2内のH3として
   - H3「スキマ時間を活用した柔軟なはたらき方」
   - H3「「クライアント企業」で自分のペースではたらく」

4. **その他の記事** → 記事テーマに関連づけて自然に
   - H3「ちょっとした収入で○○費をまかなう」
   - H3「忙しい人でも「クライアント企業」で短時間副収入」

### ランキング・〇〇選の記事（該当する場合）
- 記事内容が「おすすめランキング」「〇〇選」に適する場合に適用
- H2で「おすすめ〇選」「ランキング」を設定した場合：
  - H3がカテゴリ別の場合 → H4で①②③など番号付きで個別項目を展開
  - H3が直接ランキング項目の場合 → H3に「①」「②」「③」と番号を付ける
- **数の整合性を必ず担保**：「5選」なら必ず5個、「10選」なら必ず10個

### H4の使用条件
- ランキング内でH3がカテゴリの場合、H4で個別項目を展開
- **H4を使う場合、1つのH3につきH4を3つ以上生成すること**
- 例：
  ```
  H2: おすすめの副業6選
  H3: 在宅でできる副業
  H4: ①Webライティング
  H4: ②データ入力
  H4: ③アンケートモニター
  H3: 外出してできる副業
  H4: ④配達パートナー
  H4: ⑤イベントスタッフ
  H4: ⑥試食販売
  ```

### 見出しの文字数（目安）
- **H1: 30〜45文字**（メインタイトル・疑問形や「〜の方法」「〜を解説」で締める）
- **H2: 20〜35文字**（内容が伝わる程度に具体的に）
- **H3: 15〜20文字**（具体的な内容・答えを含める）
- **H4: 10〜20文字**（ランキング項目など）

### ★読みやすさ最重要★ 見出しのポイント

**■ 例文の使い方（重要）**
- 以下の例は**参考程度**に使う。そのままコピーしない
- **言い回しを使い回さない**（同じ表現パターンの繰り返しはNG）
- 例文から学ぶのは「雰囲気・トーン・構成の考え方」であり、文言ではない
- 記事テーマに合わせて、毎回オリジナルの表現で見出しを作成する

**■ 良い見出しの例（実際の投稿スタイル）**

【生活・節約系】
✅ H2の例：
- 「冬の暖房設定温度は何度がベスト？快適さと節約を両立する方法」
- 「暖房の設定温度を上げすぎると電気代はいくら変わる？」
- 「暖房費を節約するための具体的テクニック」

✅ H3の例：
- 「環境省が推奨する暖房の設定温度「20℃」の理由」
- 「1℃上げると電気代がどれくらい増えるのか」
- 「窓・ドアのすきま風対策で熱を逃がさない」

【バイト紹介系（仕事紹介・評判）】
✅ H2の例：
- 「郵便局バイトが「きつい」と言われる理由」
- 「郵便局バイトが向いている人・向いていない人の特徴」
- 「繁忙期（年末年始）はどれくらいきつい？対策と乗り切り方」

✅ H3の例：
- 「仕分け作業はスピードと正確さが求められる」
- 「コツコツ作業が得意な人は向いている」
- 「「クライアント企業」で柔軟にシフト調整する」

【バイト紹介系（探し方）】
✅ H2の例：
- 「着ぐるみバイトとは？仕事内容・向いてる人・1日の流れをわかりやすく解説」
- 「着ぐるみバイトの探し方｜初心者でも見つけやすい求人サイト・アプリ紹介」

✅ H3の例：
- 「どんな場所で働く？（テーマパーク・イベント・商業施設など）」
- 「衣装の中は想像以上に暑い！熱中症対策が必須」

【季節・イベント系】
✅ H2の例：
- 「おせちはいつ食べるのが正しい？基本の考え方」
- 「現代のおせち事情｜いつ・どう食べる？みんなの実態調査」

✅ H3の例：
- 「本来おせちは「元日（1月1日）の朝」に食べるのが基本」
- 「「大晦日に食べ始める」家庭が増えた理由」

【副業・お金系（〇〇選まとめ）】
✅ H2の例：
- 「初心者にも人気のすきま時間副業10選【2025年版】」
- 「1日単位でできる単発バイトで気軽に稼ぐ」

✅ H3の例：
- 「短時間バイト・単発バイト系（即金タイプ）」
- 「「クライアント企業」などアプリで手軽に仕事探し」

【貯金・家計管理系】
✅ H2の例：
- 「なぜお金が貯まらない？まずは原因を知ろう」
- 「節約でお金を貯める方法【固定費・変動費別】」
- 「すきま時間で収入をプラスにする方法」

✅ H3の例：
- 「先取り貯金で無理なく貯まる仕組みをつくる」
- 「「クライアント企業」で空き時間に収入を補う」

**■ 守るべきルール**

1. **見出しで内容が伝わること**
   - 読者が見出しを見ただけで「何が書いてあるか」わかる
   - 具体的な数字や方法を含める
   - ❌「ポイント」「コツ」だけで終わらない
   - ✅「〇〇で△△する方法」「〇〇の理由」など具体的に

2. **H2は疑問形を適度に使う**
   - 全体の3〜4割程度は疑問形にする
   - ✅「暖房は何度に設定すればいい？」「電気代はいくら変わる？」
   - ただし全部疑問形にしない（単調になる）

3. **H3は具体的な答え・方法を示す**
   - ✅「風向きと風量を変えるだけで節電できる理由」
   - ✅「サーキュレーターや扇風機で空気を循環させよう」
   - ✅「加湿器・厚手カーテン・ラグなどプラスαの工夫」

4. **記事テーマから外れたH2は禁止**
   - 本題に集中する
   - クライアント企業紹介はH3として自然に入れる（H2単独にしない）

5. **同じ語尾・表現の連続を避ける**
   - ❌「〜の理由」「〜の方法」「〜のコツ」が連続
   - ✅ 語尾を変える：「〜の理由」「〜で節約」「〜してみよう」
   - **「コツ」は使いすぎない**（1記事で1回程度、他は「方法」「ポイント」「やり方」で言い換え）

### 自然な日本語表現（最重要）
- 見出しは文法的に正しく、自然な日本語にすること
- ❌「向く人の特徴」→ ✅「向いている人の特徴」
- ❌ キーワード羅列 → ✅ 文章として読める形に
- まとめは「まとめ｜」の後にキーワード要約＋行動を促す文を入れる

### 共起語の扱い（絶対厳守）
- **共起語をスペース区切りのまま見出しに入れない**
- 共起語は「意味」を参考にし、自然な日本語に言い換えて使う
- ❌ 悪い例（そのまま入れている）：
  - 「夫婦 生活費 折半 メリットは〜」
  - 「夫婦 生活費 折半 トラブルが起きたときは？」
  - 「看護師 副業 在宅でできる〜」
- ✅ 良い例（自然な日本語に）：
  - 「生活費折半のメリットは〜」
  - 「折半でトラブルが起きたときは？」
  - 「在宅でできる看護師の副業〜」
- Q&Aの質問文も同様：共起語をそのまま入れず自然な日本語で

### 品質基準
- **見出しは具体的に**（何が書いてあるか一目でわかる）
- **読者の疑問に答える流れ**にする
- **H2の順序に意味を持たせる**（読み進めると理解が深まる構成）
- 共起語を自然に組み込む
- **読みやすさを最優先**（堅すぎず、親しみやすい表現）

---

## 出力形式

以下のフォーマットで出力してください：

```
H1: [メインタイトル（30-45文字）]

H2: [大見出し（20-35文字）]
H3: [小見出し（15-20文字）]
H3: [小見出し（15-20文字）]
H3: [小見出し（15-20文字）]

H2: [大見出し2]
H3: [小見出し2-1]
H3: [小見出し2-2]
H3: [小見出し2-3]

（ランキング系の場合）
H2: [おすすめ〇選]
H3: [カテゴリA]
H4: ①[項目1]
H4: ②[項目2]
H3: [カテゴリB]
H4: ③[項目3]
H4: ④[項目4]

...（H2は6～8個）

H2: よくある質問（Q&A）
H3: Q1：[質問1]？
H3: Q2：[質問2]？
H3: Q3：[質問3]？
H3: Q4：[質問4]？

H2: まとめ
```

---

優先順位: 1)H2数6-8個 2)各H2にH3を2〜3個 3)5,500字達成 4)見出しの自然な文章化 5)関連KW組み込み 6)検索意図

最終行に自己チェック（OK/NG）:
- [ ] H1: キーワードを自然な文章で含む（キーワード羅列ではない）
- [ ] H1: 詰め込みすぎていない（1つのメッセージ、並列表現を避ける）
- [ ] H2: 6～8個（まとめ含む）
- [ ] H3: 各H2に2〜3個（まとめ除く）
- [ ] H2とH3の論理的対応関係
- [ ] 「働」→「はたらく」等の平仮名統一
- [ ] クライアント企業紹介: まとめ前に含む
- [ ] Q&A: 該当記事にはまとめ前に含む
- [ ] ランキング系: 数の整合性OK
- [ ] 関連KW: H2/H3に自然に組み込み
- [ ] 共起語: スペース区切りのまま入れていない（自然な日本語に言い換え済み）
- [ ] 文字数: 構成で5,500字達成可能"""

            response = self.openai_client.chat.completions.create(
                model="gpt-5.2",
                messages=[
                    {"role": "system", "content": """あなたはSEO記事構成案の専門家として振る舞う。

【絶対厳守】
- H2は6～8個（最後の「まとめ」含む）。6個未満・8個超過は不可。
- 各H2（まとめを除く）にH3を2〜3個。1つ以下は不可。
- H4はランキング系でカテゴリ別展開が必要な場合のみ使用。
- 記事全体で5,500字（5,000～6,000字）になる構成。範囲外は不可。
- 出力は構成案のみ。説明・注釈・前置きは禁止。
- 「働く」は必ず「はたらく」と平仮名で表記（働き→はたらき、働いて→はたらいて等）。

【★最重要★ 読みやすい見出し】

■ 例文の使い方
- 例はあくまで**参考程度**。そのままコピー禁止
- **同じ言い回しを使い回さない**（表現パターンの繰り返しNG）
- 例から学ぶのは「雰囲気・トーン」であり、文言ではない
- 毎回オリジナルの表現で見出しを作成する

■ 見出しの文字数目安
- H1: 30〜45文字（疑問形や「〜の方法」「〜を解説」で締める）
- H2: 20〜35文字（内容が伝わる程度に具体的に）
- H3: 15〜20文字（具体的な内容・答えを含める）
- H4: 10〜20文字（ランキング項目など）

■ 良い見出しの例（このスタイルで書く）

【例1：生活・節約系】
H2の例：
- 「冬の暖房設定温度は何度がベスト？快適さと節約を両立する方法」
- 「暖房の設定温度を上げすぎると電気代はいくら変わる？」
- 「暖房費を節約するための具体的テクニック」
- 「ライフスタイル別・おすすめ暖房術」

H3の例：
- 「環境省が推奨する暖房の設定温度「20℃」の理由」
- 「リビング・寝室・オフィスなど場所別の目安温度」
- 「1℃上げると電気代がどれくらい増えるのか」
- 「窓・ドアのすきま風対策で熱を逃がさない」

【例2：バイト紹介系（仕事紹介・評判）】
H2の例：
- 「郵便局バイトが「きつい」と言われる理由」
- 「郵便局バイトの仕事内容と1日の流れ」
- 「郵便局バイトが向いている人・向いていない人の特徴」
- 「繁忙期（年末年始）はどれくらいきつい？対策と乗り切り方」
- 「「郵便局はきついからやめとけ」は本当？他バイトとの比較」

H3の例：
- 「仕分け作業はスピードと正確さが求められる」
- 「短期間で集中して稼げるメリットもある」
- 「コツコツ作業が得意な人は向いている」
- 「スキマ時間を使って柔軟にはたらく方法」
- 「クライアント企業で自分のペースではたらく」

【例2-2：バイト紹介系（探し方）】
H2の例：
- 「着ぐるみバイトとは？仕事内容・向いてる人・1日の流れをわかりやすく解説」
- 「着ぐるみバイトはきつい？大変な点と注意すべきポイント」
- 「着ぐるみバイトの探し方｜初心者でも見つけやすい求人サイト・アプリ紹介」

H3の例：
- 「どんな場所ではたらく？（テーマパーク・イベント・商業施設など）」
- 「衣装の中は想像以上に暑い！熱中症対策が必須」
- 「クライアント企業で気軽に探してみる」

【例3：季節・イベント系】
H2の例：
- 「おせちはいつ食べるのが正しい？基本の考え方」
- 「なぜおせちは正月に食べるの？料理の意味と由来」
- 「現代のおせち事情｜いつ・どう食べる？みんなの実態調査」

H3の例：
- 「本来おせちは「元日（1月1日）の朝」に食べるのが基本」
- 「料理一品一品に込められた願い（例：黒豆＝健康、数の子＝子孫繁栄）」
- 「「大晦日に食べ始める」家庭が増えた理由」

H4の例（トレンド系の展開）：
- 「2026年はSNS映えするおせちが話題！盛り付け・写真映えの工夫」
- 「コンビニおせちが手軽で人気上昇中！年末の救世主に」

【例4：副業・お金系（〇〇選まとめ）】
H2の例：
- 「すきま時間を使った副業が注目される理由」
- 「初心者にも人気のすきま時間副業10選【2025年版】」
- 「1日単位でできる副業なら「クライアント企業」がおすすめ」

H3の例：
- 「スキルの有無で分類して考える」
- 「短時間バイト・単発バイト系（即金タイプ）」
- 「登録無料で、スマホから仕事を探せる」

【例5：貯金・家計管理系】
H2の例：
- 「なぜお金が貯まらない？まずは原因を知ろう」
- 「今日からできる！お金を貯める基本のステップ」
- 「節約でお金を貯める方法【固定費・変動費別】」
- 「お金を増やすための「収入アップ」アプローチ」
- 「すきま時間で収入をプラスにするなら「クライアント企業」」

H3の例：
- 「収支のバランスが把握できていない」
- 「先取り貯金で無理なく貯まる仕組みをつくる」
- 「通信費・保険・サブスクなど固定費の見直し」
- 「副業・すきまバイトで収入を増やす」
- 「家計を助ける「ちょっと稼ぐ」新しい習慣」

■ 守るべきルール

1. 見出しで内容が伝わること
   - 読者が見出しを見ただけで「何が書いてあるか」わかる
   - 具体的な数字や方法を含める
   - ❌「ポイント」「コツ」だけで終わらない

2. H2は疑問形を適度に使う（3〜4割程度）
   - ✅「暖房は何度に設定すればいい？」「電気代はいくら変わる？」
   - ただし全部疑問形にしない（単調になる）

3. H3は具体的な答え・方法を示す
   - ✅「風向きと風量を変えるだけで節電できる理由」
   - ✅「サーキュレーターや扇風機で空気を循環させよう」

4. 同じ語尾・表現の連続を避ける
   - ❌「〜の理由」「〜の方法」「〜のコツ」が連続
   - ✅ 語尾を変える：「〜の理由」「〜で節約」「〜してみよう」
   - **「コツ」は使いすぎない**（1記事1回程度、他は「方法」「ポイント」で言い換え）

5. 記事テーマから外れたH2は禁止
   - クライアント企業紹介はH3として自然に入れる（H2にしない）

【必須セクション】
- Q&A：**基本的に全記事に入れる**。まとめの直前に配置。**4つ以上**作成。H3は「Q1：〜？」「Q2：〜？」「Q3：〜？」「Q4：〜？」と番号付き。
- クライアント企業紹介：**必ず入れる**が、**さりげなく自然に**。クライアント企業の名前は出してOKだが、露骨な宣伝NG。例：「スキマ時間で探すなら「クライアント企業」も便利」「「クライアント企業」で自分のペースではたらく」など自然な形で。

【ランキング・〇〇選】
- 該当する場合、H3がカテゴリ別ならH4で①②③と番号展開。
- **H4を使う場合、1つのH3につきH4を3つ以上**生成する。
- 数の整合性を必ず担保（「5選」なら5個、「10選」なら10個）。

【読者目線の構成】
- 自然な流れで構成：基本→詳細→応用の順で読者が理解しやすい流れ。
- 基本→応用の流れ：初心者でも理解できる順序。
- 見出しだけで内容がわかる：H2を読むだけで記事全体が把握できる。
- H3同士は同じ粒度で揃える。
- **H2同士の内容が重複しないようにする**（各H2は独立したテーマ）
- まとめ: キーワード要約＋行動を促す文（「まとめ｜」「結論｜」などの表記は避ける）

【自然な日本語】
- ❌「向く人」→ ✅「向いている人」
- ❌ キーワード羅列 → ✅ 文章として読める形に

【H1の読みやすさ（重要）】
- **1つのH1に1つのメッセージ**（詰め込みすぎない）
- **「〇〇と△△と□□」のような並列は避ける**
- ❌「看護師がムリなく副業で収入アップするはじめ方と選び方のコツ」（詰め込みすぎ）
- ✅「看護師の副業おすすめ10選｜資格を活かして無理なく稼ぐ」
- ✅「看護師が副業で稼ぐには？始め方とおすすめの仕事を紹介」

【共起語の扱い（絶対厳守）】
- **共起語をスペース区切りのまま見出しに入れるのは禁止**
- 共起語は「意味」だけを参考にし、自然な日本語に言い換える
- ❌「夫婦 生活費 折半 メリットは〜」→ ✅「生活費折半のメリットは〜」
- ❌「夫婦 生活費 折半 トラブル」→ ✅「折半でトラブルが起きたときは？」
- Q&Aの質問文も同様：共起語そのままはNG

【禁止コンテンツ（絶対厳守）】
- **副業がバレない方法・隠す方法**を示唆する見出しや内容は禁止
- ❌「会社にバレずに副業する方法」
- ❌「副業がバレないためのポイント」
- ❌「確定申告でバレない方法」
- ✅ 代わりに「会社の就業規則を確認する」「届出が必要な場合もある」など正しい情報を提示

【評価基準】優先順
1. H2数の厳守（6～8個）
2. H3数の厳守（各H2に2〜3個）
3. 見出しが具体的で内容が伝わる
4. 読みやすさ（自然な日本語、堅すぎない表現）
5. 共起語の自然な組み込み"""},
                    {"role": "user", "content": prompt}
                ],
                max_completion_tokens=2500
            )

            outline_text = response.choices[0].message.content.strip()

            logger.info(f"✓ キーワード「{keyword}」の構成案を生成しました")
            return {
                'keyword': keyword,
                'related_keywords': related_keywords,
                'top_urls': top_urls,
                'outline': outline_text,
                'success': True
            }

        except Exception as e:
            logger.error(f"エラー: キーワード「{keyword}」の構成案生成に失敗 - {e}")
            return {
                'keyword': keyword,
                'related_keywords': [],
                'top_urls': [],
                'outline': None,
                'success': False,
                'error': str(e)
            }

    def generate_outline_with_claude(self, keyword):
        """Claude APIを使用して構成案を生成（共起語・上位記事分析付き）"""
        try:
            if not self.claude_client:
                return {
                    'keyword': keyword,
                    'related_keywords': [],
                    'top_urls': [],
                    'outline': None,
                    'success': False,
                    'error': 'Claude APIキーが設定されていません'
                }

            logger.info(f"[Claude] キーワード「{keyword}」の構成案を生成中...")

            # ステップ1: 共起語と上位記事を並列取得（GPT版と同じ）
            with ThreadPoolExecutor(max_workers=2) as executor:
                future_keywords = executor.submit(self.extract_cooccurrence_keywords, keyword)
                future_articles = executor.submit(self.fetch_top_articles, keyword)

                related_keywords = future_keywords.result()
                top_articles = future_articles.result()

            related_keywords_text = '\n'.join([f"- {kw}" for kw in related_keywords]) if related_keywords else "（なし）"

            # 上位記事の見出し構造と本文をフォーマット
            top_articles_text = ""
            top_urls = []
            for i, article in enumerate(top_articles, 1):
                top_urls.append(article['url'])
                headings_text = '\n'.join([f"    {h}" for h in article['headings'][:15]]) if article['headings'] else "    （見出し取得不可）"
                body_text = article.get('body', '') if article.get('body') else "（本文取得不可）"
                top_articles_text += f"""
【記事{i}】
- URL: {article['url']}
- タイトル: {article['title']}
- 概要: {article['description']}
- 見出し構造:
{headings_text}
- 本文（抜粋）:
{body_text}
"""
            if not top_articles_text:
                top_articles_text = "（上位記事の取得に失敗）"

            logger.info(f"[Claude] 共起語: {len(related_keywords)}個、上位記事: {len(top_articles)}件取得")

            # ステップ2: Claudeで構成案を生成
            system_prompt = """あなたはSEO記事構成案の専門家です。

【絶対厳守】
- H2は6～8個（最後の「まとめ」含む）。6個未満・8個超過は不可。
- 各H2（まとめを除く）にH3を2〜3個。1つ以下は不可。
- 記事全体で5,500字（5,000～6,000字）になる構成。
- 出力は構成案のみ。説明・注釈・前置きは禁止。
- 「働く」は必ず「はたらく」と平仮名で表記。

【見出しルール】
- H1: 30〜45文字（メインタイトル）
- H2: 20〜35文字
- H3: 15〜20文字
- 見出しで内容が伝わること
- 自然な日本語にすること
- 共起語をH2/H3に自然に組み込む
- 上位記事の見出し構造を分析し、共通トピックは必ず含める
- 上位記事にない独自の視点も1-2個追加して差別化

【Q&Aセクション】
- まとめの直前に「よくある質問」を入れる
- Q1〜Q4の4つ以上

【クライアント企業紹介】
- まとめ前に自然にクライアント企業（クライアント企業）の紹介を入れる
- H3として「クライアント企業で気軽に始める」等

【出力形式】
H1: [タイトル]

H2: [見出し]
H3: [小見出し]
H3: [小見出し]

H2: [見出し]
H3: [小見出し]
H3: [小見出し]

...

H2: よくある質問（Q&A）
H3: Q1：[質問]？
H3: Q2：[質問]？
H3: Q3：[質問]？
H3: Q4：[質問]？

H2: まとめ"""

            user_prompt = f"""以下のキーワードで記事構成案を作成してください。

## メインキーワード
{keyword}

## 共起語（上位ページから抽出・構成案に自然に組み込む）
{related_keywords_text}

## 検索上位記事の分析（重要：この見出し構造を参考に構成案を作成）
{top_articles_text}

## 読者層
バイト・副業・生活情報を探す20～40代（専門知識なし）

構成案のみを出力してください。"""

            response = self.claude_client.messages.create(
                model="claude-sonnet-4-20250514",
                max_tokens=4000,
                messages=[
                    {"role": "user", "content": user_prompt}
                ],
                system=system_prompt
            )

            outline = response.content[0].text

            # 使用量情報
            usage_info = {
                'input_tokens': response.usage.input_tokens,
                'output_tokens': response.usage.output_tokens,
                'total_tokens': response.usage.input_tokens + response.usage.output_tokens
            }

            logger.info(f"[Claude] 構成案生成完了（tokens: {usage_info['total_tokens']}）")

            return {
                'keyword': keyword,
                'related_keywords': related_keywords,
                'top_urls': top_urls,
                'outline': outline,
                'usage': usage_info,
                'success': True
            }

        except Exception as e:
            logger.error(f"エラー: Claude構成案生成に失敗 - {e}")
            return {
                'keyword': keyword,
                'related_keywords': [],
                'top_urls': [],
                'outline': None,
                'success': False,
                'error': str(e)
            }

    def generate_outlines_parallel(self, keywords, max_workers=10):
        """複数のキーワードに対して並列で構成案を生成"""
        results = []

        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            future_to_keyword = {
                executor.submit(self.generate_outline_for_keyword, keyword): keyword
                for keyword in keywords
            }

            for future in as_completed(future_to_keyword):
                keyword = future_to_keyword[future]
                try:
                    result = future.result()
                    results.append(result)
                    logger.info(f"進捗: {len(results)}/{len(keywords)} 完了")
                except Exception as e:
                    logger.error(f"キーワード「{keyword}」の処理中にエラー: {e}")
                    results.append({
                        'keyword': keyword,
                        'related_keywords': [],
                        'top_urls': [],
                        'outline': None,
                        'success': False,
                        'error': str(e)
                    })

        return results

    def apply_formatting(self, sheet_id, outline_row_count):
        """スプレッドシートに書式設定を適用（見やすくする）"""
        requests = []

        # 0. 列幅を調整（見やすくする）
        # A列: 見出しレベル（150ピクセル）
        requests.append({
            'updateDimensionProperties': {
                'range': {
                    'sheetId': sheet_id,
                    'dimension': 'COLUMNS',
                    'startIndex': 0,
                    'endIndex': 1
                },
                'properties': {
                    'pixelSize': 150
                },
                'fields': 'pixelSize'
            }
        })

        # B列: タイトル（500ピクセル）
        requests.append({
            'updateDimensionProperties': {
                'range': {
                    'sheetId': sheet_id,
                    'dimension': 'COLUMNS',
                    'startIndex': 1,
                    'endIndex': 2
                },
                'properties': {
                    'pixelSize': 500
                },
                'fields': 'pixelSize'
            }
        })

        # C列: 内容（400ピクセル）
        requests.append({
            'updateDimensionProperties': {
                'range': {
                    'sheetId': sheet_id,
                    'dimension': 'COLUMNS',
                    'startIndex': 2,
                    'endIndex': 3
                },
                'properties': {
                    'pixelSize': 400
                },
                'fields': 'pixelSize'
            }
        })

        # 1. 全セルのフォントサイズを10に設定
        requests.append({
            'repeatCell': {
                'range': {
                    'sheetId': sheet_id
                },
                'cell': {
                    'userEnteredFormat': {
                        'textFormat': {
                            'fontSize': 10
                        }
                    }
                },
                'fields': 'userEnteredFormat.textFormat.fontSize'
            }
        })

        # 2. タイトル案（B2セル）のフォントサイズを11に設定
        requests.append({
            'repeatCell': {
                'range': {
                    'sheetId': sheet_id,
                    'startRowIndex': 1,  # 2行目
                    'endRowIndex': 2,    # 2行目のみ
                    'startColumnIndex': 1,  # B列
                    'endColumnIndex': 2     # B列のみ
                },
                'cell': {
                    'userEnteredFormat': {
                        'textFormat': {
                            'fontSize': 11
                        }
                    }
                },
                'fields': 'userEnteredFormat.textFormat.fontSize'
            }
        })

        # 3. B列とC列にテキスト折り返しを設定（全行）
        requests.append({
            'repeatCell': {
                'range': {
                    'sheetId': sheet_id,
                    'startColumnIndex': 1,  # B列
                    'endColumnIndex': 3     # C列まで
                },
                'cell': {
                    'userEnteredFormat': {
                        'wrapStrategy': 'WRAP'
                    }
                },
                'fields': 'userEnteredFormat.wrapStrategy'
            }
        })

        # 4. ヘッダー行（A2:A5）の背景色をパステルブルーに設定
        requests.append({
            'repeatCell': {
                'range': {
                    'sheetId': sheet_id,
                    'startRowIndex': 1,  # 2行目
                    'endRowIndex': 5,    # 5行目まで（タイトル案、メインKW、上位記事URL、共起語）
                    'startColumnIndex': 0,  # A列
                    'endColumnIndex': 1     # A列のみ
                },
                'cell': {
                    'userEnteredFormat': {
                        'backgroundColor': {
                            'red': 0.7,
                            'green': 0.85,
                            'blue': 0.95
                        },
                        'textFormat': {
                            'bold': True,
                            'foregroundColor': {
                                'red': 0.2,
                                'green': 0.2,
                                'blue': 0.2
                            }
                        }
                    }
                },
                'fields': 'userEnteredFormat(backgroundColor,textFormat)'
            }
        })

        # 4-2. カラムヘッダー行（A8:C8）の背景色をパステルオレンジに設定
        requests.append({
            'repeatCell': {
                'range': {
                    'sheetId': sheet_id,
                    'startRowIndex': 7,  # 8行目
                    'endRowIndex': 8,    # 8行目のみ
                    'startColumnIndex': 0,  # A列
                    'endColumnIndex': 3     # C列まで
                },
                'cell': {
                    'userEnteredFormat': {
                        'backgroundColor': {
                            'red': 1.0,
                            'green': 0.80,
                            'blue': 0.60
                        },
                        'textFormat': {
                            'bold': True,
                            'foregroundColor': {
                                'red': 0.2,
                                'green': 0.2,
                                'blue': 0.2
                            }
                        }
                    }
                },
                'fields': 'userEnteredFormat(backgroundColor,textFormat)'
            }
        })

        # 5. 構成案の各行（9行目以降）のA列にドロップダウンを設定
        if outline_row_count > 0:
            requests.append({
                'setDataValidation': {
                    'range': {
                        'sheetId': sheet_id,
                        'startRowIndex': 8,  # 9行目から
                        'endRowIndex': 8 + outline_row_count,  # 構成案の行数分
                        'startColumnIndex': 0,  # A列
                        'endColumnIndex': 1     # A列のみ
                    },
                    'rule': {
                        'condition': {
                            'type': 'ONE_OF_LIST',
                            'values': [
                                {'userEnteredValue': 'H1'},
                                {'userEnteredValue': 'H2'},
                                {'userEnteredValue': 'H3'}
                            ]
                        },
                        'showCustomUi': True,
                        'strict': False
                    }
                }
            })

            # 6. 条件付き書式を設定（H2 = パステル赤、H3 = パステル青）
            # H2の場合（パステル赤背景、濃いテキスト）
            requests.append({
                'addConditionalFormatRule': {
                    'rule': {
                        'ranges': [{
                            'sheetId': sheet_id,
                            'startRowIndex': 8,
                            'endRowIndex': 8 + outline_row_count,
                            'startColumnIndex': 0,
                            'endColumnIndex': 1
                        }],
                        'booleanRule': {
                            'condition': {
                                'type': 'TEXT_EQ',
                                'values': [{'userEnteredValue': 'H2'}]
                            },
                            'format': {
                                'backgroundColor': {
                                    'red': 1.0,
                                    'green': 0.78,
                                    'blue': 0.78
                                },
                                'textFormat': {
                                    'foregroundColor': {
                                        'red': 0.2,
                                        'green': 0.2,
                                        'blue': 0.2
                                    },
                                    'bold': True
                                }
                            }
                        }
                    },
                    'index': 0
                }
            })

            # H3の場合（パステル青背景、濃いテキスト）
            requests.append({
                'addConditionalFormatRule': {
                    'rule': {
                        'ranges': [{
                            'sheetId': sheet_id,
                            'startRowIndex': 8,
                            'endRowIndex': 8 + outline_row_count,
                            'startColumnIndex': 0,
                            'endColumnIndex': 1
                        }],
                        'booleanRule': {
                            'condition': {
                                'type': 'TEXT_EQ',
                                'values': [{'userEnteredValue': 'H3'}]
                            },
                            'format': {
                                'backgroundColor': {
                                    'red': 0.67,
                                    'green': 0.84,
                                    'blue': 0.96
                                },
                                'textFormat': {
                                    'foregroundColor': {
                                        'red': 0.2,
                                        'green': 0.2,
                                        'blue': 0.2
                                    },
                                    'bold': True
                                }
                            }
                        }
                    },
                    'index': 1
                }
            })

        # バッチ更新を実行
        batch_update_request = {'requests': requests}
        self.sheets_service.spreadsheets().batchUpdate(
            spreadsheetId=self.spreadsheet_id,
            body=batch_update_request
        ).execute()

    def parse_outline_to_sheet_format(self, outline_text):
        """構成案テキストをスプレッドシート用のフォーマットに変換"""
        lines = outline_text.strip().split('\n')
        h1_title = None
        rows = []

        for line in lines:
            line = line.strip()
            if not line:
                continue

            # H1, H2, H3, H4を検出
            if line.startswith('H1:'):
                h1_title = line.replace('H1:', '').strip()
            elif line.startswith('H2:'):
                rows.append(['H2', line.replace('H2:', '').strip(), ''])
            elif line.startswith('H3:'):
                rows.append(['H3', line.replace('H3:', '').strip(), ''])
            elif line.startswith('H4:'):
                rows.append(['H4', line.replace('H4:', '').strip(), ''])

        return h1_title, rows

    def create_sheet_for_keyword(self, keyword, h1_title, outline_rows, related_keywords=None, top_urls=None):
        """キーワードごとに新しいシートを作成して構成案を書き込む

        Returns:
            dict: {'sheet_name': str, 'sheet_id': int, 'sheet_url': str}
        """
        try:
            # シート名を作成（最大100文字、特殊文字を削除）
            sheet_name = keyword[:80]
            # スプレッドシートで使えない文字を削除
            sheet_name = re.sub(r'[\\\/\?\*\[\]:]', '', sheet_name)

            logger.info(f"シート '{sheet_name}' を作成中...")

            # シートを作成
            batch_update_request = {
                'requests': [{
                    'addSheet': {
                        'properties': {
                            'title': sheet_name
                        }
                    }
                }]
            }

            response = self.sheets_service.spreadsheets().batchUpdate(
                spreadsheetId=self.spreadsheet_id,
                body=batch_update_request
            ).execute()

            # シートIDを取得（書式設定で使用）
            sheet_id = response['replies'][0]['addSheet']['properties']['sheetId']

            # シートURLを生成
            sheet_url = f"https://docs.google.com/spreadsheets/d/{self.spreadsheet_id}/edit#gid={sheet_id}"

            logger.info(f"✓ シート '{sheet_name}' を作成しました (gid={sheet_id})")

            # データを書き込む（既存フォーマットに完全対応）
            # 共起語をカンマ区切りで1セルに入れる
            related_keywords_text = ', '.join(related_keywords) if related_keywords else ''
            # 上位URLをカンマ区切りで1セルに入れる
            top_urls_text = ', '.join(top_urls) if top_urls else ''

            values = [
                [''],  # 1行目: 空
                ['タイトル案（H1）', h1_title or keyword],  # 2行目: H1タイトル
                ['メインKW：', keyword],  # 3行目: キーワード
                ['上位記事URL', top_urls_text],  # 4行目: 上位記事URL（カンマ区切り）
                ['共起語：', related_keywords_text],  # 5行目: 共起語（カンマ区切り）
                ['担当', ''],  # 6行目: 担当
                ['▼構成案', ''],  # 7行目: 構成案ヘッダー
                ['見出しレベル', 'タイトル', '内容']  # 8行目: カラムヘッダー
            ]

            # 構成案データを追加（9行目から）
            values.extend(outline_rows)

            range_name = f"'{sheet_name}'!A1"

            self.sheets_service.spreadsheets().values().update(
                spreadsheetId=self.spreadsheet_id,
                range=range_name,
                valueInputOption='RAW',
                body={'values': values}
            ).execute()

            logger.info(f"✓ シート '{sheet_name}' にデータを書き込みました")

            # 書式設定を適用（見やすくする）
            self.apply_formatting(sheet_id, len(outline_rows))
            logger.info(f"✓ シート '{sheet_name}' に書式設定を適用しました")

            return {
                'sheet_name': sheet_name,
                'sheet_id': sheet_id,
                'sheet_url': sheet_url
            }

        except HttpError as err:
            logger.error(f"シート作成エラー: {err}")
            raise

    def update_master_sheet_urls(self, master_spreadsheet_id, keyword_data_map, keyword_column='G', url_column='M', title_column='L'):
        """マスターシートのURL列とタイトル列を更新

        Args:
            master_spreadsheet_id: マスターシートのスプレッドシートID
            keyword_data_map: {キーワード: {'url': URL, 'title': タイトル}} の辞書
            keyword_column: キーワードが入っている列（デフォルト: G）
            url_column: URLを書き込む列（デフォルト: M）
            title_column: タイトルを書き込む列（デフォルト: L）
        """
        try:
            logger.info(f"マスターシートに書き込み中... ({len(keyword_data_map)}件)")

            # マスターシートのデータを取得（最初のシート）
            result = self.sheets_service.spreadsheets().values().get(
                spreadsheetId=master_spreadsheet_id,
                range=f'{keyword_column}:{keyword_column}'
            ).execute()

            values = result.get('values', [])

            if not values:
                logger.warning("マスターシートにデータがありません")
                return 0

            # キーワードと行番号のマッピングを作成
            keyword_to_row = {}
            for i, row in enumerate(values):
                if row and row[0]:
                    keyword = row[0].strip()
                    keyword_to_row[keyword] = i + 1  # 1-indexed

            # URL・タイトルを書き込むリクエストを作成
            update_count = 0
            requests_data = []

            for keyword, data in keyword_data_map.items():
                # キーワードの完全一致または部分一致を試行
                row_num = keyword_to_row.get(keyword)

                if row_num:
                    # URLを書き込み
                    url = data.get('url') if isinstance(data, dict) else data
                    requests_data.append({
                        'range': f'{url_column}{row_num}',
                        'values': [[url]]
                    })

                    # タイトルを書き込み（存在する場合）
                    if isinstance(data, dict) and data.get('title'):
                        requests_data.append({
                            'range': f'{title_column}{row_num}',
                            'values': [[data['title']]]
                        })

                    update_count += 1
                    logger.info(f"  ✓ 「{keyword}」→ 行{row_num}にURL・タイトル書き込み")
                else:
                    logger.warning(f"  ⚠ 「{keyword}」がマスターシートに見つかりません")

            # バッチ更新
            if requests_data:
                self.sheets_service.spreadsheets().values().batchUpdate(
                    spreadsheetId=master_spreadsheet_id,
                    body={
                        'valueInputOption': 'RAW',
                        'data': requests_data
                    }
                ).execute()

            logger.info(f"✓ マスターシートに{update_count}件のURL・タイトルを書き込みました")
            return update_count

        except Exception as e:
            logger.error(f"マスターシートURL更新エラー: {e}")
            return 0

    def run(self, keywords, max_workers=10, master_spreadsheet_id=None, keyword_column='G', url_column='M'):
        """構成案生成からスプレッドシート書き込みまで一括実行

        Args:
            keywords: キーワードのリスト
            max_workers: 並列数
            master_spreadsheet_id: マスターシートのID（URLを書き込む場合）
            keyword_column: マスターシートのキーワード列
            url_column: マスターシートのURL書き込み列
        """
        logger.info("=" * 50)
        logger.info(f"構成案生成開始（{len(keywords)}件のキーワード）")
        logger.info("=" * 50)

        # 認証
        self.authenticate_google()

        # 並列で構成案を生成
        logger.info(f"並列処理で構成案を生成中（最大{max_workers}並列）...")
        results = self.generate_outlines_parallel(keywords, max_workers=max_workers)

        # 成功・失敗を集計
        success_count = sum(1 for r in results if r['success'])
        failed_count = len(results) - success_count

        logger.info(f"構成案生成完了: 成功{success_count}件、失敗{failed_count}件")

        # スプレッドシートに書き込む
        created_sheets = []
        errors = []
        keyword_data_map = {}  # マスターシート更新用（URL + タイトル）

        for result in results:
            if not result['success']:
                errors.append({
                    'keyword': result['keyword'],
                    'error': result.get('error', 'Unknown error')
                })
                continue

            try:
                # 構成案をパース
                h1_title, outline_rows = self.parse_outline_to_sheet_format(result['outline'])

                if not outline_rows:
                    logger.warning(f"キーワード「{result['keyword']}」: 構成案のパースに失敗")
                    errors.append({
                        'keyword': result['keyword'],
                        'error': '構成案のパースに失敗'
                    })
                    continue

                # シートを作成して書き込み（共起語 + 上位URLも渡す）
                related_keywords = result.get('related_keywords', [])
                top_urls = result.get('top_urls', [])
                sheet_result = self.create_sheet_for_keyword(result['keyword'], h1_title, outline_rows, related_keywords, top_urls)

                created_sheets.append({
                    'keyword': result['keyword'],
                    'sheet_name': sheet_result['sheet_name'],
                    'sheet_url': sheet_result['sheet_url'],
                    'title': h1_title
                })

                # マスターシート更新用にURL・タイトルを保存
                keyword_data_map[result['keyword']] = {
                    'url': sheet_result['sheet_url'],
                    'title': h1_title
                }

            except Exception as e:
                logger.error(f"キーワード「{result['keyword']}」: シート作成エラー - {e}")
                errors.append({
                    'keyword': result['keyword'],
                    'error': str(e)
                })

        # マスターシートにURL・タイトルを書き込む（指定されている場合）
        master_update_count = 0
        if master_spreadsheet_id and keyword_data_map:
            master_update_count = self.update_master_sheet_urls(
                master_spreadsheet_id,
                keyword_data_map,
                keyword_column,
                url_column,
                title_column='L'
            )

        logger.info("=" * 50)
        logger.info(f"完了: {len(created_sheets)}個のシートを作成")
        if master_spreadsheet_id:
            logger.info(f"マスターシート: {master_update_count}件のURLを更新")
        logger.info("=" * 50)

        # Slack通知を送信（キーワードとURLのみ）
        spreadsheet_url = f'https://docs.google.com/spreadsheets/d/{self.spreadsheet_id}'
        if created_sheets:
            slack_message = "📝 *構成案生成完了*\n\n"
            for sheet in created_sheets:
                slack_message += f"• {sheet['keyword']}\n  {sheet['sheet_url']}\n\n"

            send_slack_notification(slack_message)

        return {
            'success': True,
            'total_keywords': len(keywords),
            'created_sheets': len(created_sheets),
            'failed': len(errors),
            'sheets': created_sheets,
            'errors': errors,
            'spreadsheet_url': spreadsheet_url,
            'master_updated': master_update_count
        }


@app.route('/health', methods=['GET'])
def health():
    """ヘルスチェック"""
    return jsonify({'status': 'ok'})


@app.route('/generate-articles', methods=['POST'])
def generate_articles():
    """記事生成エンドポイント"""
    try:
        data = request.get_json()
        logger.info(f"[DEBUG] リクエストデータ: {data}")

        spreadsheet_id = data.get('spreadsheet_id')
        max_articles = data.get('max_articles')  # None = すべて処理
        image_generation_method = data.get('image_generation_method', 'both')  # デフォルトは両方（フォルダ + AI生成）

        # マスターシート関連のパラメータ
        master_spreadsheet_id = data.get('master_spreadsheet_id')
        keyword_column = data.get('keyword_column', 'G')
        article_url_column = data.get('article_url_column', 'N')

        logger.info(f"[DEBUG] スプレッドシートID: {spreadsheet_id}")
        logger.info(f"[DEBUG] 最大記事数: {max_articles}")
        logger.info(f"[DEBUG] 画像生成方法: {image_generation_method}")
        logger.info(f"[DEBUG] マスターシートID: {master_spreadsheet_id}")

        if not spreadsheet_id:
            logger.error("[DEBUG] スプレッドシートIDが指定されていません")
            return jsonify({'error': 'spreadsheet_id is required'}), 400

        # OpenAI APIキーを環境変数から取得
        openai_api_key = os.environ.get('OPENAI_API_KEY')
        if not openai_api_key:
            return jsonify({'error': 'OPENAI_API_KEY not set'}), 500

        # 画像フォルダーIDを環境変数から取得
        image_folder_id = os.environ.get('IMAGE_FOLDER_ID')

        # バックグラウンドで記事生成処理を実行（非同期）
        def process_articles_background():
            try:
                logger.info("[BACKGROUND] 記事生成処理を開始します")
                automation = ArticleAutomation(
                    spreadsheet_id,
                    openai_api_key,
                    image_folder_id,
                    image_generation_method=image_generation_method,
                    master_spreadsheet_id=master_spreadsheet_id,
                    keyword_column=keyword_column,
                    article_url_column=article_url_column
                )
                result = automation.process_all_sheets(max_articles)
                logger.info(f"[BACKGROUND] 記事生成処理が完了しました: {result}")
            except Exception as e:
                logger.error(f"[BACKGROUND] エラーが発生しました: {e}")
                import traceback
                logger.error(f"[BACKGROUND] トレースバック: {traceback.format_exc()}")

        # バックグラウンドスレッドで処理を開始
        import threading
        thread = threading.Thread(target=process_articles_background)
        thread.daemon = True
        thread.start()

        logger.info("[ASYNC] バックグラウンドで記事生成処理を開始しました")

        # 即座にレスポンスを返す（202 Accepted）
        articles_count = max_articles if max_articles else "すべて"
        return jsonify({
            'message': f'{articles_count}件の記事生成をバックグラウンドで開始しました',
            'status': 'processing',
            'spreadsheet_id': spreadsheet_id
        }), 202

    except Exception as e:
        logger.error(f"エラー: {e}")
        return jsonify({'error': str(e)}), 500


@app.route('/generate-single-article', methods=['POST'])
def generate_single_article():
    """単一記事生成エンドポイント（バッチ処理用）"""
    try:
        data = request.get_json()
        logger.info(f"[SINGLE] リクエストデータ: {data}")

        spreadsheet_id = data.get('spreadsheet_id')
        sheet_name = data.get('sheet_name')
        image_generation_method = data.get('image_generation_method', 'both')  # デフォルトは両方（フォルダ + AI生成）
        force = data.get('force', False)  # 処理済みでも強制再生成

        # マスターシート関連のパラメータ
        master_spreadsheet_id = data.get('master_spreadsheet_id')
        keyword_column = data.get('keyword_column', 'G')
        article_url_column = data.get('article_url_column', 'N')

        if not spreadsheet_id:
            return jsonify({'error': 'spreadsheet_id is required'}), 400

        if not sheet_name:
            return jsonify({'error': 'sheet_name is required'}), 400

        # OpenAI APIキーを環境変数から取得
        openai_api_key = os.environ.get('OPENAI_API_KEY')
        if not openai_api_key:
            return jsonify({'error': 'OPENAI_API_KEY not set'}), 500

        # 画像フォルダーIDを環境変数から取得
        image_folder_id = os.environ.get('IMAGE_FOLDER_ID')

        logger.info(f"[SINGLE] 記事生成開始: {sheet_name}")

        # 記事生成処理を実行（同期）
        automation = ArticleAutomation(
            spreadsheet_id,
            openai_api_key,
            image_folder_id,
            image_generation_method=image_generation_method,
            master_spreadsheet_id=master_spreadsheet_id,
            keyword_column=keyword_column,
            article_url_column=article_url_column
        )
        result = automation.process_single_sheet(sheet_name, force=force)

        logger.info(f"[SINGLE] 記事生成完了: {sheet_name} - {result}")

        return jsonify({
            'status': 'success',
            'sheet_name': sheet_name,
            'result': result
        }), 200

    except Exception as e:
        logger.error(f"[SINGLE] エラー: {e}")
        import traceback
        logger.error(f"[SINGLE] トレースバック: {traceback.format_exc()}")
        return jsonify({'error': str(e)}), 500


@app.route('/enqueue-all-articles', methods=['POST'])
def enqueue_all_articles():
    """全未処理シートをCloud Tasksにキュー登録"""
    try:
        data = request.get_json()
        logger.info(f"[ENQUEUE API] リクエストデータ: {data}")

        spreadsheet_id = data.get('spreadsheet_id')
        image_generation_method = data.get('image_generation_method', 'both')  # デフォルトは両方（フォルダ + AI生成）

        # マスターシート関連のパラメータ
        master_spreadsheet_id = data.get('master_spreadsheet_id')
        keyword_column = data.get('keyword_column', 'G')
        article_url_column = data.get('article_url_column', 'N')

        if not spreadsheet_id:
            return jsonify({'error': 'spreadsheet_id is required'}), 400

        # OpenAI APIキーを環境変数から取得
        openai_api_key = os.environ.get('OPENAI_API_KEY')
        if not openai_api_key:
            return jsonify({'error': 'OPENAI_API_KEY not set'}), 500

        # 画像フォルダーIDを環境変数から取得
        image_folder_id = os.environ.get('IMAGE_FOLDER_ID')

        # Cloud Run URL
        cloud_run_url = os.environ.get('CLOUD_RUN_URL', 'https://your-cloud-run-url.run.app')

        automation = ArticleAutomation(
            spreadsheet_id,
            openai_api_key,
            image_folder_id,
            image_generation_method=image_generation_method,
            master_spreadsheet_id=master_spreadsheet_id,
            keyword_column=keyword_column,
            article_url_column=article_url_column
        )

        result = automation.enqueue_articles_to_cloud_tasks(cloud_run_url)
        logger.info(f"[ENQUEUE API] キュー登録完了: {result}")

        return jsonify(result), 200

    except Exception as e:
        logger.error(f"[ENQUEUE API] エラー: {e}")
        import traceback
        logger.error(f"[ENQUEUE API] トレースバック: {traceback.format_exc()}")
        return jsonify({'error': str(e)}), 500


@app.route('/process-article-task', methods=['POST'])
def process_article_task():
    """Cloud Tasksから呼び出される個別記事処理エンドポイント"""
    try:
        data = request.get_json()
        logger.info(f"[TASK] タスク受信: {data.get('sheet_name')} ({data.get('task_index')}/{data.get('total_tasks')})")

        spreadsheet_id = data.get('spreadsheet_id')
        sheet_name = data.get('sheet_name')
        image_generation_method = data.get('image_generation_method', 'both')  # デフォルトは両方（フォルダ + AI生成）
        master_spreadsheet_id = data.get('master_spreadsheet_id')
        keyword_column = data.get('keyword_column', 'G')
        article_url_column = data.get('article_url_column', 'N')
        task_index = data.get('task_index', 0)
        total_tasks = data.get('total_tasks', 0)

        if not spreadsheet_id or not sheet_name:
            return jsonify({'error': 'spreadsheet_id and sheet_name are required'}), 400

        # OpenAI APIキーを環境変数から取得
        openai_api_key = os.environ.get('OPENAI_API_KEY')
        if not openai_api_key:
            return jsonify({'error': 'OPENAI_API_KEY not set'}), 500

        # 画像フォルダーIDを環境変数から取得
        image_folder_id = os.environ.get('IMAGE_FOLDER_ID')

        automation = ArticleAutomation(
            spreadsheet_id,
            openai_api_key,
            image_folder_id,
            image_generation_method=image_generation_method,
            master_spreadsheet_id=master_spreadsheet_id,
            keyword_column=keyword_column,
            article_url_column=article_url_column
        )

        # 記事を生成
        result = automation.process_single_sheet(sheet_name)

        # 成功時は進捗付きでSlack通知
        if result.get('status') == 'success':
            slack_webhook_url = os.environ.get('SLACK_WEBHOOK_URL')
            if slack_webhook_url:
                try:
                    message = f"""✅ *初稿生成完了* ({task_index}/{total_tasks})

*{result.get('title', sheet_name)}*
{result.get('url', '')}

キーワード: {sheet_name}"""

                    payload = {'text': message}
                    requests.post(slack_webhook_url, json=payload, timeout=10)
                except Exception as e:
                    logger.error(f"[TASK] Slack通知エラー: {e}")

        logger.info(f"[TASK] タスク完了: {sheet_name} - {result.get('status')}")

        return jsonify({
            'status': result.get('status'),
            'sheet_name': sheet_name,
            'title': result.get('title'),
            'url': result.get('url'),
            'task_index': task_index,
            'total_tasks': total_tasks
        }), 200

    except Exception as e:
        logger.error(f"[TASK] エラー: {e}")
        import traceback
        logger.error(f"[TASK] トレースバック: {traceback.format_exc()}")

        # エラー時もSlack通知
        slack_webhook_url = os.environ.get('SLACK_WEBHOOK_URL')
        if slack_webhook_url:
            try:
                sheet_name = data.get('sheet_name', '不明')
                message = f"""❌ *記事生成エラー*

シート: {sheet_name}
エラー: {str(e)}"""
                payload = {'text': message}
                requests.post(slack_webhook_url, json=payload, timeout=10)
            except:
                pass

        return jsonify({'error': str(e)}), 500


@app.route('/get-unprocessed-count', methods=['POST'])
def get_unprocessed_count():
    """未処理シートの件数を取得"""
    try:
        data = request.get_json()
        spreadsheet_id = data.get('spreadsheet_id')

        if not spreadsheet_id:
            return jsonify({'error': 'spreadsheet_id is required'}), 400

        openai_api_key = os.environ.get('OPENAI_API_KEY')
        image_folder_id = os.environ.get('IMAGE_FOLDER_ID')

        automation = ArticleAutomation(
            spreadsheet_id,
            openai_api_key,
            image_folder_id
        )
        unprocessed = automation.get_unprocessed_sheets()

        return jsonify({
            'count': len(unprocessed),
            'sheets': [s['sheet_name'] for s in unprocessed]
        }), 200

    except Exception as e:
        logger.error(f"[COUNT] エラー: {e}")
        return jsonify({'error': str(e)}), 500


@app.route('/fetch-keywords', methods=['POST'])
def fetch_keywords():
    """Search Consoleからキーワード取得エンドポイント"""
    try:
        data = request.get_json()
        logger.info(f"[DEBUG] リクエストデータ: {data}")

        site_url = data.get('site_url')
        spreadsheet_id = data.get('spreadsheet_id')
        days = data.get('days', 30)  # デフォルト30日

        logger.info(f"[DEBUG] サイトURL: {site_url}")
        logger.info(f"[DEBUG] スプレッドシートID: {spreadsheet_id}")
        logger.info(f"[DEBUG] 取得期間: {days}日")

        if not site_url:
            logger.error("[DEBUG] サイトURLが指定されていません")
            return jsonify({'error': 'site_url is required'}), 400

        if not spreadsheet_id:
            logger.error("[DEBUG] スプレッドシートIDが指定されていません")
            return jsonify({'error': 'spreadsheet_id is required'}), 400

        # キーワード取得処理
        fetcher = SearchConsoleKeywordFetcher(site_url, spreadsheet_id)
        result = fetcher.run(days=days, row_limit=1000)

        return jsonify(result), 200

    except Exception as e:
        logger.error(f"エラー: {e}")
        return jsonify({'error': str(e)}), 500


@app.route('/generate-outlines', methods=['POST'])
def generate_outlines():
    """キーワードから構成案を生成してスプレッドシートに書き込むエンドポイント

    リクエストパラメータ:
    - keywords: キーワードのリスト（必須）
    - year: 年（必須）例: 2025
    - month: 月（必須）例: 11
    - max_workers: 並列数（オプション、デフォルト10）
    - spreadsheet_id: 直接指定する場合のスプレッドシートID（オプション、year/monthより優先）
    - master_spreadsheet_id: マスターシートのID（オプション、構成案URLを書き込む）
    - keyword_column: マスターシートのキーワード列（オプション、デフォルト'G'）
    - url_column: マスターシートのURL書き込み列（オプション、デフォルト'M'）
    """
    try:
        data = request.get_json()
        logger.info(f"[DEBUG] リクエストデータ: {data}")

        keywords = data.get('keywords', [])
        year = data.get('year')
        month = data.get('month')
        max_workers = data.get('max_workers', 10)  # デフォルト10並列
        spreadsheet_id = data.get('spreadsheet_id')  # 直接指定（オプション）

        # マスターシート関連パラメータ
        master_spreadsheet_id = data.get('master_spreadsheet_id')  # URLを書き込むマスターシート
        keyword_column = data.get('keyword_column', 'G')  # キーワード列
        url_column = data.get('url_column', 'M')  # URL書き込み列

        logger.info(f"[DEBUG] キーワード数: {len(keywords)}")
        logger.info(f"[DEBUG] 年月: {year}年{month}月")
        logger.info(f"[DEBUG] 並列数: {max_workers}")
        if master_spreadsheet_id:
            logger.info(f"[DEBUG] マスターシートID: {master_spreadsheet_id}")
            logger.info(f"[DEBUG] キーワード列: {keyword_column}, URL列: {url_column}")

        if not keywords or len(keywords) == 0:
            logger.error("[DEBUG] キーワードが指定されていません")
            return jsonify({'error': 'keywords are required'}), 400

        # year, monthが指定されていなく、spreadsheet_idも無い場合はエラー
        if not spreadsheet_id and (not year or not month):
            logger.error("[DEBUG] year, month または spreadsheet_id が必要です")
            return jsonify({'error': 'year and month are required (or spreadsheet_id)'}), 400

        # OpenAI APIキーを環境変数から取得
        openai_api_key = os.environ.get('OPENAI_API_KEY')
        if not openai_api_key:
            logger.error("[DEBUG] OPENAI_API_KEYが設定されていません")
            return jsonify({'error': 'OPENAI_API_KEY not set'}), 500

        # spreadsheet_idが直接指定されていなければ、月別スプシを取得/作成
        if spreadsheet_id:
            output_spreadsheet_id = spreadsheet_id
            logger.info(f"[DEBUG] 直接指定されたスプレッドシートID: {output_spreadsheet_id}")
        else:
            # 認証してから月別スプシを取得/作成
            generator = OutlineGenerator(None, openai_api_key)
            generator.authenticate_google()
            output_spreadsheet_id = generator.get_or_create_monthly_spreadsheet(year, month)
            logger.info(f"[DEBUG] 月別スプレッドシートID: {output_spreadsheet_id}")

        # 構成案生成処理（出力先スプレッドシートに書き込む）
        generator = OutlineGenerator(output_spreadsheet_id, openai_api_key)
        result = generator.run(
            keywords,
            max_workers=max_workers,
            master_spreadsheet_id=master_spreadsheet_id,
            keyword_column=keyword_column,
            url_column=url_column
        )

        # レスポンスに年月情報を追加
        result['year'] = year
        result['month'] = month
        result['spreadsheet_name'] = f"{year}年{month}月" if year and month else None

        return jsonify(result), 200

    except Exception as e:
        logger.error(f"エラー: {e}")
        import traceback
        logger.error(f"トレースバック: {traceback.format_exc()}")
        return jsonify({'error': str(e)}), 500


@app.route('/generate-outline-claude', methods=['POST'])
def generate_outline_claude():
    """Claude APIを使用して構成案を生成するエンドポイント（性能テスト用）

    リクエストパラメータ:
    - keyword: キーワード（必須）
    """
    try:
        data = request.get_json()
        keyword = data.get('keyword')

        if not keyword:
            return jsonify({'error': 'keyword is required'}), 400

        # OpenAI APIキーは必要（OutlineGeneratorの初期化に必要）
        openai_api_key = os.environ.get('OPENAI_API_KEY')
        if not openai_api_key:
            return jsonify({'error': 'OPENAI_API_KEY not set'}), 500

        # Anthropic APIキーを確認
        anthropic_api_key = os.environ.get('ANTHROPIC_API_KEY')
        if not anthropic_api_key:
            return jsonify({'error': 'ANTHROPIC_API_KEY not set'}), 500

        # OutlineGeneratorを作成（spreadsheet_idはNullで可）
        generator = OutlineGenerator(
            spreadsheet_id=None,
            openai_api_key=openai_api_key,
            anthropic_api_key=anthropic_api_key
        )

        # Claude APIで構成案を生成
        result = generator.generate_outline_with_claude(keyword)

        return jsonify(result), 200

    except Exception as e:
        logger.error(f"エラー: {e}")
        import traceback
        logger.error(f"トレースバック: {traceback.format_exc()}")
        return jsonify({'error': str(e)}), 500


@app.route('/generate-draft-claude', methods=['POST'])
def generate_draft_claude():
    """Claude APIを使用して初稿を生成するエンドポイント（性能テスト用）

    リクエストパラメータ:
    - keyword: キーワード（必須）
    - h1_title: H1タイトル（必須）
    - headings: 見出し構造のリスト（必須）
      例: [{"level": "H2", "text": "見出し1"}, {"level": "H3", "text": "小見出し1"}]
    """
    try:
        data = request.get_json()
        keyword = data.get('keyword')
        h1_title = data.get('h1_title')
        headings = data.get('headings')

        if not keyword:
            return jsonify({'error': 'keyword is required'}), 400
        if not h1_title:
            return jsonify({'error': 'h1_title is required'}), 400
        if not headings:
            return jsonify({'error': 'headings is required'}), 400

        # OpenAI APIキーを環境変数から取得
        openai_api_key = os.environ.get('OPENAI_API_KEY')
        if not openai_api_key:
            return jsonify({'error': 'OPENAI_API_KEY not set'}), 500

        # Anthropic APIキーを確認
        anthropic_api_key = os.environ.get('ANTHROPIC_API_KEY')
        if not anthropic_api_key:
            return jsonify({'error': 'ANTHROPIC_API_KEY not set'}), 500

        # ArticleAutomationを作成
        automation = ArticleAutomation(
            spreadsheet_id=None,
            openai_api_key=openai_api_key,
            anthropic_api_key=anthropic_api_key
        )

        # Claude APIで初稿を生成
        draft_md, usage = automation.generate_draft_with_claude(keyword, h1_title, headings)

        if draft_md.startswith("ERROR:"):
            return jsonify({'error': draft_md, 'success': False}), 500

        return jsonify({
            'keyword': keyword,
            'h1_title': h1_title,
            'draft': draft_md,
            'char_count': len(draft_md),
            'usage': usage,
            'success': True
        }), 200

    except Exception as e:
        logger.error(f"エラー: {e}")
        import traceback
        logger.error(f"トレースバック: {traceback.format_exc()}")
        return jsonify({'error': str(e)}), 500


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 8080))
    app.run(host='0.0.0.0', port=port)
