"""
クライアント企業メディア記事自動生成スクリプト

このスクリプトは以下の処理を自動化します：
1. Google スプレッドシートから見出しデータを取得
2. OpenAI API を使用して記事を生成
3. 生成した記事を Google ドキュメントに保存
"""

import os
from dotenv import load_dotenv
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from openai import OpenAI
import json

# .envファイルから環境変数を読み込む
load_dotenv()

# Google API のスコープ
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',  # 読み書き両方必要
    'https://www.googleapis.com/auth/documents'
]

class ArticleAutomation:
    def __init__(self, spreadsheet_id, openai_api_key):
        """
        初期化

        Args:
            spreadsheet_id (str): Google スプレッドシートのID
            openai_api_key (str): OpenAI の APIキー
        """
        self.spreadsheet_id = spreadsheet_id
        self.openai_client = OpenAI(api_key=openai_api_key)
        self.sheets_service = None
        self.docs_service = None
        self.prompt_template = self._load_prompt_template()

    def _load_prompt_template(self):
        """プロンプトテンプレートを読み込む"""
        with open('article_generation_prompt.txt', 'r', encoding='utf-8') as f:
            return f.read()

    def authenticate_google(self):
        """Google API の認証を行う"""
        creds = None
        # token.json には前回の認証情報が保存される
        if os.path.exists('token.json'):
            creds = Credentials.from_authorized_user_file('token.json', SCOPES)

        # 認証情報がない、または無効な場合は再認証
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    'credentials.json', SCOPES)
                creds = flow.run_local_server(port=0)

            # 認証情報を保存
            with open('token.json', 'w') as token:
                token.write(creds.to_json())

        # サービスオブジェクトを作成
        self.sheets_service = build('sheets', 'v4', credentials=creds)
        self.docs_service = build('docs', 'v1', credentials=creds)

        print("✓ Google API 認証完了")

    def get_all_sheets(self):
        """すべてのシート名を取得"""
        try:
            sheet_metadata = self.sheets_service.spreadsheets().get(
                spreadsheetId=self.spreadsheet_id
            ).execute()
            sheets = sheet_metadata.get('sheets', [])
            return [sheet['properties']['title'] for sheet in sheets]
        except HttpError as err:
            print(f"エラー: シート一覧の取得に失敗しました - {err}")
            return []

    def get_headings_from_sheet(self, sheet_name=None):
        """
        スプレッドシートから見出しデータを取得

        実際のスプレッドシート構造:
        - 2行目 B列: タイトル案（H1）
        - 3行目 B列: メインKW（キーワード）
        - 8行目: ヘッダー行（H階層 / 見出し文）
        - 9行目～: 見出し構造
          - A列: H1, H2, H3
          - B列: 見出し文
        - F列: ステータス（処理済み/未処理）

        Args:
            sheet_name (str): シート名（Noneの場合は最初のシートを使用）

        Returns:
            list: 見出しデータのリスト
        """
        try:
            # シート名が指定されていない場合は、すべてのシートを取得
            if sheet_name is None:
                all_sheets = self.get_all_sheets()
                if not all_sheets:
                    print("シートが見つかりません")
                    return []
                # 最初のシート以外の全シートを処理対象とする
                # （1つ目のシートは通常、目次や説明シートのことが多いため）
                sheets_to_process = all_sheets
            else:
                sheets_to_process = [sheet_name]

            all_headings = []

            for sheet in sheets_to_process:
                # シート全体を読み込む（A1:F100）
                range_name = f"'{sheet}'!A1:F100"
                result = self.sheets_service.spreadsheets().values().get(
                    spreadsheetId=self.spreadsheet_id,
                    range=range_name
                ).execute()

                values = result.get('values', [])

                if not values or len(values) < 9:
                    print(f"シート '{sheet}' にデータが不足しています")
                    continue

                # 3行目 B列: メインKW（キーワード）
                keyword = values[2][1] if len(values) > 2 and len(values[2]) > 1 else ""

                # ステータスを確認（F列、仮に2行目とする）
                status = values[1][5] if len(values) > 1 and len(values[1]) > 5 else ""

                if status == "処理済み":
                    print(f"シート '{sheet}' は処理済みです")
                    continue

                # 8行目以降から見出しを抽出
                h1_title = ""
                h2_headings = []

                for i in range(7, len(values)):  # 7行目（インデックス7、実際は8行目）から
                    row = values[i]
                    if len(row) < 2:
                        continue

                    hierarchy = row[0].strip() if row[0] else ""
                    heading_text = row[1].strip() if row[1] else ""

                    if not heading_text:
                        continue

                    if hierarchy == "H1":
                        h1_title = heading_text
                    elif hierarchy == "H2":
                        h2_headings.append(heading_text)

                # H1が見つからない場合は2行目のタイトル案を使用
                if not h1_title and len(values) > 1 and len(values[1]) > 1:
                    h1_title = values[1][1]

                if h1_title and h2_headings:
                    all_headings.append({
                        'sheet_name': sheet,
                        'keyword': keyword,
                        'h1_title': h1_title,
                        'h2_headings': h2_headings
                    })
                    print(f"✓ シート '{sheet}' から見出しを取得: H1='{h1_title[:30]}...', H2={len(h2_headings)}個")

            print(f"\n✓ 合計 {len(all_headings)} 件の未処理記事を取得しました")
            return all_headings

        except HttpError as err:
            print(f"エラー: スプレッドシートの取得に失敗しました - {err}")
            return []

    def generate_article(self, keyword, h1_title, h2_headings):
        """
        OpenAI API を使用して記事を生成

        Args:
            keyword (str): キーワード
            h1_title (str): H1見出し
            h2_headings (list): H2見出しのリスト

        Returns:
            str: 生成された記事
        """
        # H2見出しをフォーマット
        h2_formatted = '\n'.join([f"{i+1}. {heading}" for i, heading in enumerate(h2_headings)])

        # プロンプトに値を埋め込む
        prompt = self.prompt_template.replace('{keyword}', keyword)
        prompt = prompt.replace('{h1_title}', h1_title)
        prompt = prompt.replace('{h2_headings}', h2_formatted)

        try:
            print(f"記事生成中: {h1_title}")

            response = self.openai_client.chat.completions.create(
                model="gpt-4o",  # または "gpt-3.5-turbo"
                messages=[
                    {"role": "system", "content": "あなたはクライアント企業メディアの専門ライターです。記事には必ずマークダウン形式の表を1つ以上含めてください。表がない記事は不合格です。"},
                    {"role": "user", "content": prompt}
                ],
                temperature=0.7,
                max_tokens=4000
            )

            article = response.choices[0].message.content
            print("✓ 記事生成完了")
            return article

        except Exception as e:
            print(f"エラー: 記事の生成に失敗しました - {e}")
            return None

    def save_to_google_docs(self, article, title):
        """
        生成した記事を Google ドキュメントに保存

        Args:
            article (str): 記事本文
            title (str): ドキュメントのタイトル

        Returns:
            str: 作成されたドキュメントのURL
        """
        try:
            # 新しいドキュメントを作成
            document = self.docs_service.documents().create(
                body={'title': title}
            ).execute()

            document_id = document.get('documentId')
            print(f"✓ ドキュメント作成: {title}")

            # 記事本文を挿入
            requests = [
                {
                    'insertText': {
                        'location': {'index': 1},
                        'text': article
                    }
                }
            ]

            self.docs_service.documents().batchUpdate(
                documentId=document_id,
                body={'requests': requests}
            ).execute()

            document_url = f"https://docs.google.com/document/d/{document_id}/edit"
            print(f"✓ 記事を保存しました: {document_url}")

            return document_url

        except HttpError as err:
            print(f"エラー: ドキュメントの作成に失敗しました - {err}")
            return None

    def update_sheet_status(self, sheet_name, status="処理済み", doc_url=""):
        """
        スプレッドシートのステータスを更新

        Args:
            sheet_name (str): シート名
            status (str): ステータス
            doc_url (str): ドキュメントURL
        """
        try:
            # F2列（ステータス）とG2列（URL）を更新
            range_name = f"'{sheet_name}'!F2:G2"
            values = [[status, doc_url]]

            body = {'values': values}

            self.sheets_service.spreadsheets().values().update(
                spreadsheetId=self.spreadsheet_id,
                range=range_name,
                valueInputOption='RAW',
                body=body
            ).execute()

            print(f"✓ シート '{sheet_name}' のステータスを更新しました")

        except HttpError as err:
            print(f"エラー: ステータスの更新に失敗しました - {err}")

    def process_all_articles(self):
        """すべての未処理記事を処理"""
        # 認証
        self.authenticate_google()

        # 見出しデータを取得
        headings_list = self.get_headings_from_sheet()

        if not headings_list:
            print("処理する記事がありません")
            return

        # 各記事を処理
        for heading_data in headings_list:
            print(f"\n{'='*60}")
            print(f"処理中: {heading_data['h1_title']}")
            print(f"シート: {heading_data['sheet_name']}")
            print(f"{'='*60}")

            # 記事生成
            article = self.generate_article(
                heading_data['keyword'],
                heading_data['h1_title'],
                heading_data['h2_headings']
            )

            if article:
                # Google ドキュメントに保存
                doc_url = self.save_to_google_docs(
                    article,
                    heading_data['h1_title']
                )

                if doc_url:
                    # ステータスを更新
                    self.update_sheet_status(
                        heading_data['sheet_name'],
                        "処理済み",
                        doc_url
                    )

            print()

        print(f"\n{'='*60}")
        print("すべての記事の処理が完了しました！")
        print(f"{'='*60}")


def main():
    """メイン処理"""
    # 設定ファイルから値を読み込む（または直接指定）
    SPREADSHEET_ID = os.getenv('SPREADSHEET_ID', 'YOUR_SPREADSHEET_ID_HERE')
    OPENAI_API_KEY = os.getenv('OPENAI_API_KEY', 'YOUR_OPENAI_API_KEY_HERE')

    # 自動化処理を実行
    automation = ArticleAutomation(SPREADSHEET_ID, OPENAI_API_KEY)
    automation.process_all_articles()


if __name__ == '__main__':
    main()
