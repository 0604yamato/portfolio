/**
 * クライアント企業メディア記事自動生成 GAS版
 *
 * セットアップ手順：
 * 1. スプレッドシートを開く
 * 2. 「拡張機能」→「Apps Script」でスクリプトエディタを開く
 * 3. このコードを貼り付け
 * 4. 「プロジェクトの設定」→「スクリプトプロパティ」で以下を設定：
 *    - OPENAI_API_KEY: あなたのOpenAI APIキー
 * 5. 保存して実行
 */

// グローバル定数
const MODEL = "gpt-4o";
const MAX_TOKENS = 10000;
const TEMPERATURE = 0.7;
const CLOUD_RUN_URL = 'https://your-cloud-run-url.run.app';
const MASTER_SPREADSHEET_ID = 'YOUR_MASTER_SPREADSHEET_ID';  // マスターシートID

/**
 * スプレッドシートを開いたときに実行（カスタムメニュー追加）
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('記事自動生成')
    .addItem('📅 月別構成案生成', 'generateOutlinesByMonth')
    .addItem('📅 月別構成案生成（1件テスト）', 'generateOutlinesByMonthTest')
    .addSeparator()
    .addItem('📄 月別初稿生成', 'generateArticlesByMonth')
    .addItem('📄 月別初稿生成（1件テスト）', 'generateArticlesByMonthTest')
    .addSeparator()
    .addItem('🤖 月別構成案生成（Claude）', 'generateOutlinesByMonthClaude')
    .addItem('🤖 月別構成案生成（Claude・1件テスト）', 'generateOutlinesByMonthClaudeTest')
    .addItem('🤖 月別初稿生成（Claude）', 'generateArticlesByMonthClaude')
    .addItem('🤖 月別初稿生成（Claude・1件テスト）', 'generateArticlesByMonthClaudeTest')
    .addSeparator()
    .addItem('🚀 全記事一括生成（100件対応）', 'generateAllArticlesBulk')
    .addItem('🛑 生成を停止', 'stopBatchGeneration')
    .addSeparator()
    .addItem('⚙️ 画像フォルダ設定', 'setImageFolderId')
    .addItem('⚙️ APIキーを設定', 'setApiKey')
    .addToUi();
}

/**
 * OpenAI APIキーを設定するダイアログ
 */
function setApiKey() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt(
    'OpenAI APIキーの設定',
    'OpenAI APIキーを入力してください：',
    ui.ButtonSet.OK_CANCEL
  );

  if (result.getSelectedButton() == ui.Button.OK) {
    const apiKey = result.getResponseText();
    PropertiesService.getScriptProperties().setProperty('OPENAI_API_KEY', apiKey);
    ui.alert('APIキーを保存しました！');
  }
}


/**
 * 画像フォルダIDを設定するダイアログ
 */
function setImageFolderId() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();

  const result = ui.prompt(
    '画像フォルダ設定',
    'Googleドライブの画像フォルダIDを入力してください\n（フォルダを開いたときのURLの最後の部分）：',
    ui.ButtonSet.OK_CANCEL
  );

  if (result.getSelectedButton() == ui.Button.OK) {
    const folderId = result.getResponseText().trim();
    props.setProperty('IMAGE_FOLDER_ID', folderId);
    ui.alert('画像フォルダIDを保存しました！');
  }
}


/**
 * 現在のシートの記事を生成
 */
function generateCurrentSheetArticle() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const sheetName = sheet.getName();

  const ui = SpreadsheetApp.getUi();

  // ステータス確認（処理済みの場合は再確認）
  const status = sheet.getRange('F2').getValue();
  if (status === '処理済み') {
    const confirmResult = ui.alert(
      '再生成確認',
      `シート「${sheetName}」は既に処理済みです。\n再度記事を生成しますか？\n（前の記事は上書きされます）`,
      ui.ButtonSet.YES_NO
    );

    if (confirmResult != ui.Button.YES) return;
  } else {
    const result = ui.alert(
      '記事生成確認',
      `シート「${sheetName}」の記事を生成しますか？`,
      ui.ButtonSet.YES_NO
    );

    if (result != ui.Button.YES) return;
  }

  try {
    processSheet(sheet);
    ui.alert('✓ 記事の生成が完了しました！');
  } catch (error) {
    ui.alert('エラー', `記事の生成に失敗しました：\n${error.message}`, ui.ButtonSet.OK);
    Logger.log(error);
  }
}


/**
 * すべての未処理シートの記事を生成
 */
function generateAllArticles() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.alert(
    '一括生成確認',
    'すべての未処理シートの記事を生成しますか？\n（処理に時間がかかる場合があります）',
    ui.ButtonSet.YES_NO
  );

  if (result == ui.Button.YES) {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = spreadsheet.getSheets();
    let processedCount = 0;
    let errorCount = 0;

    for (let sheet of sheets) {
      try {
        // ステータスを確認
        const status = sheet.getRange('F2').getValue();
        if (status === '処理済み') {
          Logger.log(`シート「${sheet.getName()}」はスキップ（処理済み）`);
          continue;
        }

        processSheet(sheet);
        processedCount++;

        // API制限対策：各リクエスト後に少し待機
        Utilities.sleep(2000);

      } catch (error) {
        Logger.log(`シート「${sheet.getName()}」でエラー: ${error.message}`);
        errorCount++;
      }
    }

    ui.alert(
      '一括生成完了',
      `処理完了: ${processedCount}件\nエラー: ${errorCount}件`,
      ui.ButtonSet.OK
    );
  }
}

/**
 * 1つのシートを処理
 */
function processSheet(sheet) {
  const sheetName = sheet.getName();
  Logger.log(`処理開始: ${sheetName}`);

  // 見出しデータを取得
  const headingData = extractHeadingsFromSheet(sheet);

  if (!headingData.h1Title || headingData.headings.length === 0) {
    throw new Error('見出しデータが不足しています');
  }

  Logger.log(`キーワード: ${headingData.keyword}`);
  Logger.log(`H1: ${headingData.h1Title}`);
  Logger.log(`H2数: ${headingData.headings.length}`);

  // 記事を生成
  const article = generateArticle(
    headingData.keyword,
    headingData.h1Title,
    headingData.headings
  );

  // Googleドキュメントに保存
  const docUrl = saveToGoogleDocs(article, headingData.h1Title);

  // ステータスを更新
  sheet.getRange('F2').setValue('処理済み');
  sheet.getRange('G2').setValue(docUrl);

  Logger.log(`✓ 完了: ${docUrl}`);
}

/**
 * シートから見出しデータを抽出
 */
function extractHeadingsFromSheet(sheet) {
  // 3行目 B列: キーワード
  const keyword = sheet.getRange('B3').getValue() || '';

  // 2行目 B列: H1タイトル（予備）
  const h1TitleBackup = sheet.getRange('B2').getValue() || '';

  // 8行目以降: 見出し構造
  const lastRow = sheet.getLastRow();
  let h1Title = '';
  const headings = []; // H2とH3を階層構造で保持

  for (let row = 8; row <= lastRow; row++) {
    const hierarchy = sheet.getRange(row, 1).getValue();
    const headingText = sheet.getRange(row, 2).getValue();

    if (!headingText) continue;

    if (hierarchy === 'H1') {
      h1Title = headingText;
    } else if (hierarchy === 'H2') {
      headings.push({
        level: 'H2',
        text: headingText,
        children: [] // H3見出しを格納
      });
    } else if (hierarchy === 'H3') {
      // 最後のH2の子としてH3を追加
      if (headings.length > 0) {
        headings[headings.length - 1].children.push({
          level: 'H3',
          text: headingText
        });
      }
    }
  }

  // H1が見つからない場合は2行目のタイトル案を使用
  if (!h1Title) {
    h1Title = h1TitleBackup;
  }

  return {
    keyword: keyword,
    h1Title: h1Title,
    headings: headings // 階層構造で返す
  };
}

/**
 * OpenAI APIを使用して記事を生成
 */
function generateArticle(keyword, h1Title, headings) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');

  if (!apiKey) {
    throw new Error('OpenAI APIキーが設定されていません。メニューから「APIキーを設定」を実行してください。');
  }

  // プロンプトを作成
  const prompt = createPrompt(keyword, h1Title, headings);

  // OpenAI APIにリクエスト
  const url = 'https://api.openai.com/v1/chat/completions';
  const payload = {
    model: MODEL,
    messages: [
      {
        role: 'system',
        content: 'あなたはクライアント企業メディアの専門ライターです。記事には必ずマークダウン形式の表を1つ以上含めてください。表がない記事は不合格です。'
      },
      {
        role: 'user',
        content: prompt
      }
    ],
    temperature: TEMPERATURE,
    max_tokens: MAX_TOKENS
  };

  const options = {
    method: 'post',
    headers: {
      'Authorization': `Bearer ${apiKey}`,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  Logger.log('OpenAI APIを呼び出し中...');
  const response = UrlFetchApp.fetch(url, options);
  const responseCode = response.getResponseCode();
  const responseText = response.getContentText();

  if (responseCode !== 200) {
    throw new Error(`OpenAI API エラー (${responseCode}): ${responseText}`);
  }

  const json = JSON.parse(responseText);
  const article = json.choices[0].message.content;

  Logger.log(`✓ 記事生成完了（${article.length}文字）`);
  return article;
}

/**
 * プロンプトを作成
 */
function createPrompt(keyword, h1Title, headings) {
  // 見出しリストを階層構造で作成
  let headingsList = '';
  headings.forEach((h2, i) => {
    headingsList += `${i + 1}. H2: ${h2.text}\n`;
    h2.children.forEach((h3, j) => {
      headingsList += `   ${i + 1}-${j + 1}. H3: ${h3.text}\n`;
    });
  });

  return `# クライアント企業メディア記事生成プロンプト

あなたはクライアント企業メディアの記事ライターです。以下の指示に従って、読者にとって有益で実用的な記事を作成してください。

## 【最重要】表について - 必ず守ること
**この記事には必ず1つ以上（推奨2～3個）のマークダウン形式の表を含めてください。**
表がない記事は絶対に不合格です。以下のような表を必ず作成してください：
- 比較表（時給比較、プラン比較など）
- メリット・デメリット表
- 手順・チェックリスト表
- 料金表、分類表など

**表の記載例：**
| 項目 | 内容A | 内容B |
|------|-------|-------|
| 特徴1 | 説明 | 説明 |

## 【重要】文字数について
この記事は必ず6,000文字（5,800～6,200文字）で作成してください。短くしないでください。各セクションで十分な説明と具体例を含め、ボリュームのある記事を書いてください。

## 【重要】見出しについて
**絶対に勝手に見出しを追加しないでください。** 以下に提供された見出しのみを使用してください。見出しを省略したり、変更したりしないでください。

## 記事の基本情報
- キーワード: ${keyword}
- H1見出し: ${h1Title}
- 見出し構造（これらのみを使用）:
${headingsList}

## 記事作成の指示

### 1. 全体構成
- H1（メインタイトル）: ${h1Title} をそのまま使用
- H2（主要セクション）: **提供されたH2見出しを必ず全て使用**
  - 上記の見出し構造に記載されたH2見出しを、順番通りにマークダウン形式（## 見出し）で記事内に記載する
  - 見出しを省略したり、変更したりしない
  - 勝手に見出しを追加しない
- H3（小見出し）: **提供されたH3見出しを必ず全て使用**
  - 上記の見出し構造に記載されたH3見出しを、対応するH2セクション内にマークダウン形式（### 見出し）で記事内に記載する
  - 見出しを省略したり、変更したりしない
  - 勝手にH3見出しを追加しない
  - 各H3見出しの後には150～200文字程度の説明文を書く
- **全体の文字数: 必ず6,000文字（5,800～6,200文字）で作成**
- 画像やバナーの挿入は不要（テキストのみで作成）

### 2. 文体・トーン
- 親切で丁寧: 「～ですか？」「～でしょう」と問いかけ、読者に共感を誘発
- 具体的かつ実践的: 抽象的な表現を避け、具体例やデータを含める
- 権威性と信頼性: 必要に応じて統計データや専門機関の情報を参照
- 呼びかけ型: 「ぜひ参考にしてください」「検討してみましょう」など読者への声かけを含める
- 初心者向け: 専門用語には必ず解説を付ける
- 「働く」は「はたらく」と平仮名で表記する

### 3. コンテンツの特徴
- 導入部（300～400文字）: 疑問形で読者の課題を提示し、関心を引く
- 定義・概要: テーマの基本的な定義や背景を多角的に説明
- 具体例・リスト: 7～10個程度の具体的な例やカテゴリーを提示
- 実践的アドバイス: 読者が実際に行動できる具体的な方法を提示
- 具体的な計算例: 金額や数字が関係する場合、必ず計算例を示す
- **【必須】表の使用**: 記事内に必ず1つ以上（できれば2～3個）のマークダウン形式の表を含めること
  - 比較データ、料金表、手順、分類、チェックリスト、メリット・デメリットなどは必ず表で視覚化する
  - 例: 料金比較表、メリット・デメリット比較表、手順表、チェックリスト、分類表、時給比較表など
  - 表がない記事は不合格
- まとめ（300～400文字）: クライアント企業への誘導を含む

### 4. 各セクションの書き方

**【重要】記事の構成例:**

例）
# H1タイトル（${h1Title}）

導入文（300～400文字）

## H2見出し1（スプレッドシートに記載された見出しをそのまま使用）
概要文（200～300文字）

### H3見出し1-1（スプレッドシートに記載された見出しをそのまま使用）
説明文（150～200文字）

### H3見出し1-2（スプレッドシートに記載された見出しをそのまま使用）
説明文（150～200文字）

## H2見出し2（スプレッドシートに記載された見出しをそのまま使用）
概要文（200～300文字）

### H3見出し2-1（スプレッドシートに記載された見出しをそのまま使用）
説明文（150～200文字）

（以下、提供された見出し構造に従って全て使用）

## まとめ
まとめ文（300～400文字）

各H2セクションは以下の構成で**必ず900～1,100文字**書いてください：

**【重要】H2セクションの構成:**
1. **H2見出しをマークダウン形式で記載**: ## 見出し（スプレッドシートに記載された通り）
2. **H2見出しの直後に概要文**: まず200～300文字程度の概要文・導入文を書く
   - このH2セクション全体で何を説明するか
   - なぜこのトピックが重要か
   - 読者にどんなメリットがあるか
3. **H3見出しとその内容**: スプレッドシートに記載されたH3見出しを使用
   - H3見出しをマークダウン形式で記載: ### 見出し（スプレッドシートに記載された通り）
   - 各H3見出しの後に150～200文字程度の説明文を書く
   - H3見出しを勝手に追加したり、省略したりしない
4. **本文**: 具体的な情報、データ、例を豊富に含める
   - 箇条書き（・）を適宜使用
   - 「問題点」→「解決策」のセットで提示
   - 複数の具体例を詳しく説明する
   - 読者が理解しやすいように丁寧に解説する
   - 計算例や実例を含めて詳しく説明する
   - **表や図が効果的な場合は積極的に使用する**（比較データ、手順、分類など）
5. **段落は2～4文で短く区切る**（非常に重要）
6. **まとめ・移行**: 次のセクションへの自然な移行

**注意: スプレッドシートに記載された見出しのみを使用する。勝手に見出しを追加したり、省略したりしない。**

**文字数配分の目安（必ず守ること）:**
- 導入文: 300～400文字
- 各H2セクション: **900～1,100文字** × H2の数（短くしないこと）
  - H2直後の概要文: 200～300文字
  - 本文: 600～800文字（具体例を3～5個含む）
- まとめ: 300～400文字
- **合計: 必ず6,000文字程度（5,800～6,200文字）**

### 5. 強調の使い方（非常に重要）
- 数字は必ず鉤括弧で強調：「103万円」「20選」「3つの方法」など
- 重要なキーワードを鉤括弧で強調：「〜の壁」「源泉徴収」など
- 1段落に1～2箇所の強調を目安に
- 強調は「」（鉤括弧）を使用する

### 6. 見出しの作成ルール
- 読者の疑問に答える形式（「どのように～？」「なぜ～？」）
- 具体性を持たせる（「3つの方法」「5つのポイント」など数字を使う）
- キーワードを自然に含める

### 7. 記事の終わり方
最終セクションの後に、以下の内容を追加（300～400文字程度）：

## まとめ

この記事では、${keyword}について詳しく解説しました。

まずはスキマ時間を活用して、さまざまな仕事を体験してみるのもおすすめです。「クライアント企業」なら、面接なし・履歴書なしで、1日単位からお仕事を探せます。

短期バイトから自分に合ったはたらき方を見つけ、将来のキャリアを考えるきっかけにしてみてください。

## 記事生成時の注意事項

1. **【最重要】文字数厳守**: 必ず5,800～6,200文字の範囲内で作成（理想は6,000文字程度）。3,000文字や4,000文字などの短い記事は絶対に避けてください。
2. **【最重要】見出しはスプレッドシートのもののみ使用**:
   - 提供されたH2見出しとH3見出しを必ず全て使用する
   - 見出しを省略したり、変更したり、追加したりしない
   - H2はマークダウン形式（## 見出し）、H3は（### 見出し）で記事内に記載する
   - スプレッドシートに記載されていない見出しを勝手に追加しない
3. **H2見出しの直後に概要文を書く**: 各H2見出しの直後に200～300文字の概要文・導入文を必ず書く
4. **H3見出しの後に説明文を書く**: 各H3見出しの後に150～200文字の説明文を必ず書く
5. **本文を充実させる**: 具体例を詳しく解説する
6. 段落は2～4文で短く: 読みやすさを最優先、スマホ対応
7. 鉤括弧で強調: 数字・キーワードは必ず「」で強調
8. **【最重要】表を必ず含める**: 記事内に必ず1つ以上（できれば2～3個）のマークダウン形式の表を含めること。比較データ、手順、分類、料金表、メリット・デメリットなど、表で表現できるものは必ず表で視覚化する。**表がない記事は不合格**
9. 具体的な計算例: 金額や数字の説明には必ず計算例を含める
10. 箇条書き多用: ほぼすべてのセクションで箇条書きを活用
11. 詳しく解説: 各H2セクションで具体例を複数挙げ、詳細に説明する
12. SEO対策: キーワードを見出しや本文に自然に含める（詰め込みすぎない）
13. 正確性: 具体的な統計やデータを使用する場合は「○○調査によると」などと出典を明記
14. 画像不要: テキストのみで構成し、画像の挿入指示や代替テキストは含めない
15. 平仮名表記: 「働く」は「はたらく」と平仮名で表記する

## 【絶対に守ること】

記事全体が**必ず5,800～6,200文字（理想は6,000文字程度）**になるよう、各セクションの文字数を調整してください。短い記事は不合格です。

**文字数の内訳:**
- 導入文: 300～400文字
- 各H2セクション: **900～1,100文字**（例を豊富に、詳しく説明する）
  - **H2見出しの直後に概要文: 200～300文字**
  - **本文: 600～800文字**（具体例を3～5個含む）
- まとめ: 300～400文字

**重要な注意:**
- **提供されたH2見出しとH3見出しを必ず全て使用する**（省略・変更・追加しない）
- **H2見出しはマークダウン形式（## 見出し）で記事内に記載する**
- **H3見出しはマークダウン形式（### 見出し）で記事内に記載する**
- **スプレッドシートに記載されていない見出しを勝手に追加しない**
- **スプレッドシートに記載された見出しのみを使用する**
- **【最重要】表を必ず1つ以上含める**: 記事内に必ず1つ以上（できれば2～3個）のマークダウン形式の表を含めること。比較データ、料金、手順、メリット・デメリットなどは必ず表で視覚化。**表がない記事は不合格**
- 文字数を増やすため、各セクションの本文を充実させる

段落を短く（2～4文）、鉤括弧「」で強調し、クライアント企業メディアの既存記事のスタイルに合わせてください。

**【重要】表の使用例（記事内に必ず1つ以上含めること）:**

**例1: 比較表**
| 項目 | プランA | プランB | プランC |
|------|---------|---------|---------|
| 時給 | 1,200円 | 1,500円 | 1,800円 |
| 勤務時間 | 4時間～ | 6時間～ | 8時間～ |
| 特徴 | 短時間OK | 高時給 | フルタイム |

**例2: メリット・デメリット表**
| 働き方 | メリット | デメリット |
|--------|----------|------------|
| 短期バイト | 自由度が高い、様々な経験 | 収入が不安定 |
| 長期バイト | 安定収入、スキル習得 | 拘束時間が長い |

**例3: 手順・チェックリスト表**
| 手順 | 内容 | 所要時間 |
|------|------|----------|
| 1 | アプリをダウンロード | 3分 |
| 2 | プロフィール登録 | 5分 |
| 3 | お仕事を検索 | 10分 |

上記の指示に従って、提供された見出しから**6,000文字程度の高品質な記事**を生成してください。

**【再確認】記事には必ず1つ以上のマークダウン形式の表を含めてください。表がない場合は不合格です。**`;
}

/**
 * Googleドキュメントに記事を保存
 */
function saveToGoogleDocs(article, title) {
  // 新しいドキュメントを作成
  const doc = DocumentApp.create(title);
  const body = doc.getBody();

  // 記事本文を挿入
  body.setText(article);

  // ドキュメントのURLを取得
  const docUrl = doc.getUrl();

  Logger.log(`✓ ドキュメント作成: ${docUrl}`);
  return docUrl;
}

// ============================================
// Google Drive画像取得
// ============================================

/**
 * Googleドライブのカテゴリー別フォルダから全画像を取得
 * @return {Array<Object>} 画像情報の配列 [{id, name, url, category, blob}]
 */
function getAllImagesFromDrive() {
  const props = PropertiesService.getScriptProperties();
  const folderId = props.getProperty('IMAGE_FOLDER_ID');

  if (!folderId) {
    throw new Error('画像フォルダIDが設定されていません。メニューから「画像フォルダ設定」を実行してください。');
  }

  const images = [];
  const rootFolder = DriveApp.getFolderById(folderId);

  // カテゴリー別サブフォルダを探索
  const subFolders = rootFolder.getFolders();
  while (subFolders.hasNext()) {
    const folder = subFolders.next();
    const category = folder.getName();

    // フォルダ内の画像ファイルを取得
    const files = folder.getFiles();
    while (files.hasNext()) {
      const file = files.next();
      const mimeType = file.getMimeType();

      // 画像ファイルのみ対象
      if (mimeType.startsWith('image/')) {
        images.push({
          id: file.getId(),
          name: file.getName(),
          url: file.getUrl(),
          category: category,
          blob: file.getBlob()
        });
      }
    }
  }

  Logger.log(`✓ 画像を${images.length}件取得しました`);
  return images;
}

/**
 * 見出しに最もマッチする画像を見つける
 * @param {string} heading - H2見出しテキスト
 * @param {Array<Object>} images - 画像情報の配列
 * @return {Object|null} マッチした画像情報
 */
function findBestMatchingImage(heading, images) {
  if (!images || images.length === 0) return null;

  // キーワード抽出（簡易版：平仮名・カタカナ・漢字のみ）
  const keywords = heading.match(/[ぁ-んァ-ヶー一-龠]+/g) || [];

  let bestMatch = null;
  let bestScore = 0;

  for (const image of images) {
    let score = 0;

    // カテゴリー名とのマッチング
    for (const keyword of keywords) {
      if (image.category.includes(keyword)) {
        score += 3; // カテゴリー名マッチは高得点
      }
      if (image.name.includes(keyword)) {
        score += 2; // ファイル名マッチ
      }
    }

    if (score > bestScore) {
      bestScore = score;
      bestMatch = image;
    }
  }

  // スコアが0の場合はランダムに選択
  if (bestScore === 0 && images.length > 0) {
    const randomIndex = Math.floor(Math.random() * images.length);
    bestMatch = images[randomIndex];
    Logger.log(`マッチなし: ランダム選択 (${bestMatch.name})`);
  } else {
    Logger.log(`✓ 見出し「${heading}」に画像「${bestMatch.name}」をマッチング (スコア: ${bestScore})`);
  }

  return bestMatch;
}



// ==========================================
// Cloud Run版記事生成
// ==========================================

/**
 * Cloud Run版で記事を一括生成（フォルダ画像）
 */
function generateArticlesCloudRunExistingFolder() {
  generateArticlesCloudRunWithMethod('existing_folder', 'フォルダ画像');
}

/**
 * Cloud Run版で記事を一括生成（共通処理）
 */
function generateArticlesCloudRunWithMethod(imageGenerationMethod, methodName) {
  const ui = SpreadsheetApp.getUi();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const spreadsheetId = spreadsheet.getId();

  // 生成する記事数を入力
  const result = ui.prompt(
    '記事生成数',
    '何記事生成しますか？（空白=すべての未処理記事）',
    ui.ButtonSet.OK_CANCEL
  );

  if (result.getSelectedButton() != ui.Button.OK) {
    return;
  }

  const maxArticles = result.getResponseText() ? parseInt(result.getResponseText()) : null;

  try {
    // Cloud Runを呼び出し
    Logger.log(`Cloud Runに接続: ${CLOUD_RUN_URL}`);
    Logger.log(`スプレッドシートID: ${spreadsheetId}`);
    Logger.log(`画像生成方法: ${imageGenerationMethod}`);

    const response = UrlFetchApp.fetch(`${CLOUD_RUN_URL}/generate-articles`, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({
        spreadsheet_id: spreadsheetId,
        max_articles: maxArticles,
        image_generation_method: imageGenerationMethod,
        master_spreadsheet_id: MASTER_SPREADSHEET_ID,
        keyword_column: 'G',
        article_url_column: 'N'
      }),
      muteHttpExceptions: true,
      timeout: 300  // タイムアウトを300秒（5分）に延長
    });

    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    Logger.log(`レスポンスコード: ${responseCode}`);
    Logger.log(`レスポンス: ${responseText}`);

    if (responseCode === 200) {
      const data = JSON.parse(responseText);

      // 全シート状態の詳細を作成
      let allSheetsDetails = '';
      if (data.all_sheets_status && data.all_sheets_status.length > 0) {
        allSheetsDetails = '\n\n全シートの状態:\n';
        data.all_sheets_status.forEach(s => {
          if (s.status === 'processed') {
            allSheetsDetails += `✓ ${s.sheet}: 処理成功\n`;
          } else if (s.status === 'skipped') {
            allSheetsDetails += `- ${s.sheet}: スキップ (${s.reason})\n`;
          } else if (s.status === 'error') {
            allSheetsDetails += `✗ ${s.sheet}: エラー (${s.error})\n`;
          } else if (s.status === 'warning') {
            allSheetsDetails += `⚠ ${s.sheet}: 警告 (${s.error})\n`;
          }
        });
      }

      ui.alert(
        '✓ 記事生成完了',
        `全シート数: ${data.total_sheets}件\n処理完了: ${data.total}件\n成功: ${data.processed.length}件\nスキップ: ${data.skipped ? data.skipped.length : 0}件\nエラー: ${data.errors.length}件${allSheetsDetails}`,
        ui.ButtonSet.OK
      );
    } else if (responseCode === 202) {
      // 非同期処理（バックグラウンド実行 + 完了通知）
      const data = JSON.parse(responseText);

      // 推定完了時間を計算（1記事あたり3-5分）
      const articlesToGenerate = maxArticles || 1;
      const estimatedMinutes = Math.ceil(articlesToGenerate * 4); // 1記事あたり平均4分

      ui.alert(
        '処理開始',
        `バックグラウンドで処理中です。\n完了したら通知します。\n\n推定時間: 約${estimatedMinutes}分`,
        ui.ButtonSet.OK
      );

      // スプレッドシートを監視して完了を通知
      monitorArticleGeneration(spreadsheet, articlesToGenerate, estimatedMinutes);
    } else {
      ui.alert('✗ エラー', `記事生成に失敗しました (${responseCode}):\n\n${responseText}`, ui.ButtonSet.OK);
    }

  } catch (error) {
    Logger.log(`エラー: ${error}`);
    ui.alert('✗ エラー', `通信エラー: ${error.message}`, ui.ButtonSet.OK);
  }
}

/**
 * 記事生成の完了を監視してアラート表示
 */
function monitorArticleGeneration(spreadsheet, expectedArticles, estimatedMinutes) {
  const ui = SpreadsheetApp.getUi();
  const maxChecks = 40; // 最大40回チェック（20分）
  const checkInterval = 30; // 30秒ごとにチェック

  let completedSheets = [];
  let checksCount = 0;

  Logger.log(`記事生成の監視を開始（最大${maxChecks}回、${checkInterval}秒間隔）`);

  // 初回チェックまで少し待つ（Cloud Runの起動時間を考慮）
  Utilities.sleep(15000); // 15秒

  while (checksCount < maxChecks) {
    checksCount++;
    Logger.log(`チェック ${checksCount}/${maxChecks}`);

    // 全シートをチェック
    const sheets = spreadsheet.getSheets();
    const newlyCompleted = [];

    for (const sheet of sheets) {
      const sheetName = sheet.getName();

      // すでに通知済みのシートはスキップ
      if (completedSheets.includes(sheetName)) continue;

      // ステータス列（B列）をチェック
      const statusCell = sheet.getRange('B1');
      const status = statusCell.getValue();

      // ドキュメントURL列（C列）をチェック
      const urlCell = sheet.getRange('C1');
      const docUrl = urlCell.getValue();

      // 「処理済み」かつURLが存在する場合
      if (status === '処理済み' && docUrl && docUrl.toString().trim() !== '') {
        newlyCompleted.push({
          name: sheetName,
          url: docUrl
        });
        completedSheets.push(sheetName);
        Logger.log(`✓ 完了検出: ${sheetName} - ${docUrl}`);
      }
    }

    // 新しく完了したシートがあれば通知
    if (newlyCompleted.length > 0) {
      let message = `${newlyCompleted.length}件の記事が生成されました！\n\n`;
      newlyCompleted.forEach(sheet => {
        message += `✓ ${sheet.name}\n   ${sheet.url}\n\n`;
      });
      message += `合計完了: ${completedSheets.length}件`;

      if (completedSheets.length >= expectedArticles) {
        message += `\n\nすべての記事生成が完了しました。`;
      } else {
        message += `\n\n残り: ${expectedArticles - completedSheets.length}件（推定）`;
      }

      ui.alert('✓ 記事生成完了通知', message, ui.ButtonSet.OK);

      // すべて完了したら監視終了
      if (completedSheets.length >= expectedArticles) {
        Logger.log('すべての記事生成が完了しました');
        return;
      }
    }

    // 最後のチェックでなければ待機
    if (checksCount < maxChecks) {
      Utilities.sleep(checkInterval * 1000);
    }
  }

  // タイムアウト
  if (completedSheets.length > 0) {
    ui.alert(
      '⚠ 監視タイムアウト',
      `監視時間が終了しました。\n\n完了: ${completedSheets.length}件\n\n一部の記事がまだ生成中の可能性があります。\nスプレッドシートを手動で確認してください。`,
      ui.ButtonSet.OK
    );
  } else {
    ui.alert(
      '⚠ 監視タイムアウト',
      `監視時間が終了しましたが、まだ記事が生成されていません。\n\nCloud Runのログを確認するか、しばらく待ってからスプレッドシートを確認してください。`,
      ui.ButtonSet.OK
    );
  }
}


/**
 * キーワードから構成案を一括生成（Cloud Run版）
 */
function generateOutlinesFromKeywords() {
  const ui = SpreadsheetApp.getUi();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = SpreadsheetApp.getActiveSheet();
  const spreadsheetId = spreadsheet.getId();

  // 1. 選択範囲を取得してキーワードリストを抽出
  const range = sheet.getActiveRange();
  if (!range) {
    ui.alert('⚠️ エラー', 'キーワードが含まれる範囲を選択してください。\n\n例: A列のキーワードリストを選択', ui.ButtonSet.OK);
    return;
  }

  const values = range.getValues();
  const keywords = [];

  // 選択範囲からキーワードを抽出（空白行を除く）
  for (let i = 0; i < values.length; i++) {
    const keyword = values[i][0];
    if (keyword && keyword.toString().trim() !== '') {
      keywords.push(keyword.toString().trim());
    }
  }

  if (keywords.length === 0) {
    ui.alert('⚠️ エラー', 'キーワードが見つかりませんでした。\n\nキーワードが含まれるセル範囲を選択してください。', ui.ButtonSet.OK);
    return;
  }

  // 2. 確認ダイアログ
  const result = ui.alert(
    '構成案生成確認',
    `${keywords.length}件のキーワードから構成案を生成します。\n\n処理時間: 約${Math.ceil(keywords.length / 10)}～${Math.ceil(keywords.length / 5)}分\n（10並列処理）\n\nよろしいですか？`,
    ui.ButtonSet.YES_NO
  );

  if (result != ui.Button.YES) return;

  ui.alert('生成開始', `${keywords.length}件のキーワードから構成案を生成しています。\n\n処理が完了するまでお待ちください。\n（バックグラウンドで実行されます）`, ui.ButtonSet.OK);

  try {
    // Cloud Runを呼び出し
    Logger.log(`Cloud Runに接続: ${CLOUD_RUN_URL}`);
    Logger.log(`スプレッドシートID: ${spreadsheetId}`);
    Logger.log(`キーワード数: ${keywords.length}`);
    Logger.log(`キーワード: ${keywords.slice(0, 5).join(', ')}...`);

    const response = UrlFetchApp.fetch(`${CLOUD_RUN_URL}/generate-outlines`, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({
        spreadsheet_id: spreadsheetId,
        keywords: keywords,
        max_workers: 10
      }),
      muteHttpExceptions: true,
      timeout: 300  // タイムアウトを300秒（5分）に延長
    });

    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    Logger.log(`レスポンスコード: ${responseCode}`);
    Logger.log(`レスポンス: ${responseText}`);

    if (responseCode === 200) {
      const data = JSON.parse(responseText);

      if (data.success) {
        let message = `全体: ${data.total_keywords}件\n成功: ${data.created_sheets}件\n失敗: ${data.failed}件\n\n`;

        if (data.created_sheets > 0) {
          message += '作成されたシート:\n';
          data.sheets.slice(0, 5).forEach(s => {
            message += `  - ${s.sheet_name}\n`;
          });
          if (data.created_sheets > 5) {
            message += `  ... 他${data.created_sheets - 5}件\n`;
          }
        }

        if (data.failed > 0 && data.errors.length > 0) {
          message += '\nエラー:\n';
          data.errors.slice(0, 3).forEach(e => {
            message += `  - ${e.keyword}: ${e.error}\n`;
          });
          if (data.errors.length > 3) {
            message += `  ... 他${data.errors.length - 3}件\n`;
          }
        }

        ui.alert('✓ 構成案生成完了', message, ui.ButtonSet.OK);
      } else {
        ui.alert('⚠️ エラー', data.message || '構成案生成に失敗しました', ui.ButtonSet.OK);
      }
    } else {
      ui.alert('✗ エラー', `構成案生成に失敗しました (${responseCode}):\n\n${responseText}`, ui.ButtonSet.OK);
    }

  } catch (error) {
    Logger.log(`エラー: ${error}`);
    ui.alert('✗ エラー', `通信エラー: ${error.message}`, ui.ButtonSet.OK);
  }
}

/**
 * ボリュームの多いキーワードから構成案を生成（共通処理）
 */
function generateOutlinesFromTopKeywordsCore(numArticles, minVolume) {
  const ui = SpreadsheetApp.getUi();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = SpreadsheetApp.getActiveSheet();
  const spreadsheetId = spreadsheet.getId();

  ui.alert('データ取得中', 'キーワードデータを取得しています...', ui.ButtonSet.OK);

  try {
    // シートからキーワードとボリュームを取得
    const lastRow = sheet.getLastRow();
    const allData = sheet.getRange(4, 1, lastRow - 3, 10).getValues(); // 4行目から全データ取得

    const keywordData = [];

    // 左側のデータ（C列:キーワード, D列:ボリューム）
    for (let i = 0; i < allData.length; i++) {
      const keyword = allData[i][2]; // C列（インデックス2）
      const volume = allData[i][3];  // D列（インデックス3）

      if (keyword && keyword.toString().trim() !== '') {
        const vol = parseInt(volume) || 0;
        if (vol >= minVolume) {
          keywordData.push({
            keyword: keyword.toString().trim(),
            volume: vol
          });
        }
      }
    }

    // 右側のデータ（H列:キーワード, I列:ボリューム）
    for (let i = 0; i < allData.length; i++) {
      const keyword = allData[i][7]; // H列（インデックス7）
      const volume = allData[i][8];  // I列（インデックス8）

      if (keyword && keyword.toString().trim() !== '') {
        const vol = parseInt(volume) || 0;
        if (vol >= minVolume) {
          keywordData.push({
            keyword: keyword.toString().trim(),
            volume: vol
          });
        }
      }
    }

    if (keywordData.length === 0) {
      ui.alert('⚠️ エラー', 'キーワードが見つかりませんでした。\n\nC列・H列にキーワード、D列・I列にボリュームが入力されていることを確認してください。', ui.ButtonSet.OK);
      return;
    }

    // ボリュームでソート（降順）
    keywordData.sort((a, b) => b.volume - a.volume);

    // 上位N件を選択
    const topKeywords = keywordData.slice(0, Math.min(numArticles, keywordData.length));
    const keywords = topKeywords.map(k => k.keyword);

    Logger.log(`取得したキーワード数: ${keywordData.length}`);
    Logger.log(`選択したキーワード数: ${keywords.length}`);
    Logger.log(`上位キーワード: ${topKeywords.slice(0, 5).map(k => `${k.keyword}(${k.volume})`).join(', ')}...`);

    // 確認ダイアログ
    let confirmMessage = `以下の設定で構成案を生成します：\n\n`;
    confirmMessage += `生成数: ${keywords.length}件\n`;
    if (minVolume > 0) {
      confirmMessage += `最小ボリューム: ${minVolume}以上\n`;
    }
    confirmMessage += `処理時間: 約${Math.ceil(keywords.length / 10)}～${Math.ceil(keywords.length / 5)}分\n\n`;
    confirmMessage += `上位${Math.min(5, keywords.length)}件のキーワード:\n`;
    topKeywords.slice(0, 5).forEach((k, i) => {
      confirmMessage += `${i + 1}. ${k.keyword} (vol: ${k.volume})\n`;
    });
    if (keywords.length > 5) {
      confirmMessage += `... 他${keywords.length - 5}件\n`;
    }
    confirmMessage += `\nよろしいですか？`;

    const result = ui.alert('構成案生成確認', confirmMessage, ui.ButtonSet.YES_NO);

    if (result != ui.Button.YES) return;

    ui.alert('生成開始', `${keywords.length}件のキーワードから構成案を生成しています。\n\n処理が完了するまでお待ちください。\n（バックグラウンドで実行されます）`, ui.ButtonSet.OK);

    // Cloud Runを呼び出し
    Logger.log(`Cloud Runに接続: ${CLOUD_RUN_URL}`);
    Logger.log(`スプレッドシートID: ${spreadsheetId}`);
    Logger.log(`キーワード数: ${keywords.length}`);

    const response = UrlFetchApp.fetch(`${CLOUD_RUN_URL}/generate-outlines`, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({
        spreadsheet_id: spreadsheetId,
        keywords: keywords,
        max_workers: 10
      }),
      muteHttpExceptions: true,
      timeout: 300  // タイムアウトを300秒（5分）に延長
    });

    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    Logger.log(`レスポンスコード: ${responseCode}`);
    Logger.log(`レスポンス: ${responseText}`);

    if (responseCode === 200) {
      const data = JSON.parse(responseText);

      if (data.success) {
        let message = `全体: ${data.total_keywords}件\n成功: ${data.created_sheets}件\n失敗: ${data.failed}件\n\n`;

        if (data.created_sheets > 0) {
          message += '作成されたシート:\n';
          data.sheets.slice(0, 5).forEach(s => {
            message += `  - ${s.sheet_name}\n`;
          });
          if (data.created_sheets > 5) {
            message += `  ... 他${data.created_sheets - 5}件\n`;
          }
        }

        if (data.failed > 0 && data.errors.length > 0) {
          message += '\nエラー:\n';
          data.errors.slice(0, 3).forEach(e => {
            message += `  - ${e.keyword}: ${e.error}\n`;
          });
          if (data.errors.length > 3) {
            message += `  ... 他${data.errors.length - 3}件\n`;
          }
        }

        ui.alert('✓ 構成案生成完了', message, ui.ButtonSet.OK);
      } else {
        ui.alert('⚠️ エラー', data.message || '構成案生成に失敗しました', ui.ButtonSet.OK);
      }
    } else {
      ui.alert('✗ エラー', `構成案生成に失敗しました (${responseCode}):\n\n${responseText}`, ui.ButtonSet.OK);
    }

  } catch (error) {
    Logger.log(`エラー: ${error}`);
    ui.alert('✗ エラー', `通信エラー: ${error.message}`, ui.ButtonSet.OK);
  }
}

/**
 * ボリューム上位1件から構成案を生成（精度確認用）
 */
function generateOutlinesTop1() {
  generateOutlinesFromTopKeywordsCore(1, 0);
}

/**
 * ボリューム上位3件から構成案を生成
 */
function generateOutlinesTop3() {
  generateOutlinesFromTopKeywordsCore(3, 0);
}

/**
 * ボリューム上位5件から構成案を生成
 */
function generateOutlinesTop5() {
  generateOutlinesFromTopKeywordsCore(5, 0);
}

/**
 * ボリューム上位10件から構成案を生成
 */
function generateOutlinesTop10() {
  generateOutlinesFromTopKeywordsCore(10, 0);
}

/**
 * ボリューム上位20件から構成案を生成
 */
function generateOutlinesTop20() {
  generateOutlinesFromTopKeywordsCore(20, 0);
}

/**
 * ボリューム上位50件から構成案を生成
 */
function generateOutlinesTop50() {
  generateOutlinesFromTopKeywordsCore(50, 0);
}

/**
 * ボリューム上位100件から構成案を生成
 */
function generateOutlinesTop100() {
  generateOutlinesFromTopKeywordsCore(100, 0);
}

/**
 * ボリュームの多いキーワードから構成案を生成（カスタム数値入力）
 */
function generateOutlinesFromTopKeywords() {
  const ui = SpreadsheetApp.getUi();

  // 1. 生成する記事数を入力
  const numResult = ui.prompt(
    '生成数の入力',
    '何件の構成案を生成しますか？\n（検索ボリュームの多い順に選択されます）',
    ui.ButtonSet.OK_CANCEL
  );

  if (numResult.getSelectedButton() != ui.Button.OK) return;

  const numArticles = parseInt(numResult.getResponseText());
  if (isNaN(numArticles) || numArticles <= 0) {
    ui.alert('⚠️ エラー', '有効な数値を入力してください。', ui.ButtonSet.OK);
    return;
  }

  // 2. 最小ボリュームの閾値を入力（オプション）
  const minVolResult = ui.prompt(
    '最小ボリュームの設定（オプション）',
    '最小検索ボリュームを入力してください\n（これ以下のキーワードは除外されます。0で全て含める）',
    ui.ButtonSet.OK_CANCEL
  );

  if (minVolResult.getSelectedButton() != ui.Button.OK) return;

  const minVolume = parseInt(minVolResult.getResponseText()) || 0;

  // 共通処理を呼び出し
  generateOutlinesFromTopKeywordsCore(numArticles, minVolume);
}

// ============================================================
// 一括生成機能（トリガーベース）
// ============================================================

/**
 * 10記事一括生成を開始
 */
function startBatchGeneration10() {
  startBatchGeneration(10);
}

/**
 * カスタム数の一括生成を開始
 */
function startBatchGenerationCustom() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt(
    '一括生成',
    '生成する記事数を入力してください:',
    ui.ButtonSet.OK_CANCEL
  );

  if (result.getSelectedButton() != ui.Button.OK) return;

  const count = parseInt(result.getResponseText());
  if (isNaN(count) || count < 1) {
    ui.alert('エラー', '1以上の数値を入力してください', ui.ButtonSet.OK);
    return;
  }

  startBatchGeneration(count);
}

/**
 * 一括生成を開始（メイン関数）
 */
function startBatchGeneration(totalArticles) {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // 未処理のシート名リストを取得
  const pendingSheets = [];
  const sheets = spreadsheet.getSheets();

  for (const sheet of sheets) {
    const sheetName = sheet.getName();
    // キーワードシートや設定シートはスキップ
    if (sheetName === 'キーワード' || sheetName === '設定' || sheetName === 'keywords') continue;

    // ステータスをチェック（B1セル）
    const status = sheet.getRange('B1').getValue();
    if (!status || status === '' || status === '未処理') {
      pendingSheets.push(sheetName);
    }
  }

  if (pendingSheets.length === 0) {
    ui.alert('情報', '未処理の記事がありません', ui.ButtonSet.OK);
    return;
  }

  // 生成数を調整
  const actualCount = Math.min(totalArticles, pendingSheets.length);
  const sheetsToProcess = pendingSheets.slice(0, actualCount);

  // 状態を保存
  props.setProperty('BATCH_SHEETS', JSON.stringify(sheetsToProcess));
  props.setProperty('BATCH_CURRENT_INDEX', '0');
  props.setProperty('BATCH_TOTAL', actualCount.toString());
  props.setProperty('BATCH_SPREADSHEET_ID', spreadsheet.getId());
  props.setProperty('BATCH_RUNNING', 'true');

  // 既存のトリガーを削除
  deleteBatchTriggers();

  // 1分ごとのトリガーを作成
  ScriptApp.newTrigger('processBatchArticle')
    .timeBased()
    .everyMinutes(1)
    .create();

  ui.alert(
    '一括生成開始',
    `${actualCount}記事の生成を開始します。\n\n` +
    `約${actualCount * 4}分で完了予定です。\n` +
    `進捗はスプレッドシートで確認できます。\n\n` +
    `停止する場合は「一括生成を停止」を選択してください。`,
    ui.ButtonSet.OK
  );

  // 最初の1記事をすぐに開始
  processBatchArticle();
}

/**
 * トリガーから呼ばれる：1記事を生成
 */
function processBatchArticle() {
  const props = PropertiesService.getScriptProperties();

  // 実行中かチェック
  if (props.getProperty('BATCH_RUNNING') !== 'true') {
    deleteBatchTriggers();
    return;
  }

  const sheetsJson = props.getProperty('BATCH_SHEETS');
  const currentIndex = parseInt(props.getProperty('BATCH_CURRENT_INDEX') || '0');
  const spreadsheetId = props.getProperty('BATCH_SPREADSHEET_ID');

  if (!sheetsJson || !spreadsheetId) {
    Logger.log('バッチ設定が見つかりません');
    stopBatchGeneration();
    return;
  }

  const sheets = JSON.parse(sheetsJson);

  // 全て完了したか確認
  if (currentIndex >= sheets.length) {
    Logger.log('全記事の生成が完了しました');
    stopBatchGeneration();
    return;
  }

  const sheetName = sheets[currentIndex];
  Logger.log(`[${currentIndex + 1}/${sheets.length}] 記事生成開始: ${sheetName}`);

  try {
    // Cloud Runで1記事生成
    const response = UrlFetchApp.fetch(`${CLOUD_RUN_URL}/generate-single-article`, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({
        spreadsheet_id: spreadsheetId,
        sheet_name: sheetName,
        image_generation_method: 'existing_folder',
        master_spreadsheet_id: MASTER_SPREADSHEET_ID,
        keyword_column: 'G',
        article_url_column: 'N'
      }),
      muteHttpExceptions: true,
      timeout: 300
    });

    const responseCode = response.getResponseCode();
    Logger.log(`レスポンス: ${responseCode} - ${response.getContentText().substring(0, 200)}`);

    if (responseCode === 200) {
      Logger.log(`✓ 記事生成成功: ${sheetName}`);
    } else {
      Logger.log(`✗ 記事生成エラー: ${sheetName} - ${response.getContentText()}`);
    }

  } catch (error) {
    Logger.log(`✗ 記事生成例外: ${sheetName} - ${error.message}`);
  }

  // 次のインデックスに進む
  props.setProperty('BATCH_CURRENT_INDEX', (currentIndex + 1).toString());

  // 完了したか確認
  if (currentIndex + 1 >= sheets.length) {
    Logger.log('全記事の生成が完了しました');
    stopBatchGeneration();
  }
}

/**
 * 一括生成を停止
 */
function stopBatchGeneration() {
  const props = PropertiesService.getScriptProperties();
  const ui = SpreadsheetApp.getUi();

  // 状態をクリア
  props.deleteProperty('BATCH_SHEETS');
  props.deleteProperty('BATCH_CURRENT_INDEX');
  props.deleteProperty('BATCH_TOTAL');
  props.deleteProperty('BATCH_SPREADSHEET_ID');
  props.setProperty('BATCH_RUNNING', 'false');

  // トリガーを削除
  deleteBatchTriggers();

  Logger.log('一括生成を停止しました');

  try {
    ui.alert('停止完了', '一括生成を停止しました', ui.ButtonSet.OK);
  } catch (e) {
    // トリガーから呼ばれた場合はUIがないのでスキップ
  }
}

/**
 * バッチ生成用のトリガーを削除
 */
function deleteBatchTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'processBatchArticle') {
      ScriptApp.deleteTrigger(trigger);
      Logger.log('トリガーを削除しました');
    }
  }
}

// ============================================================
// 月別構成案生成機能
// ============================================================

/**
 * 入力された月文字列から年と月を抽出
 * @param {string} inputMonth - 入力された月（例：「12月」「2025年1月」「2025年12月」）
 * @returns {Object} {year: number, month: number} または null
 */
function parseYearMonth(inputMonth) {
  // 「2025年1月」「2025年12月」形式
  const fullMatch = inputMonth.match(/(\d{4})年(\d{1,2})月/);
  if (fullMatch) {
    return {
      year: parseInt(fullMatch[1], 10),
      month: parseInt(fullMatch[2], 10)
    };
  }

  // 「1月」「12月」形式（年なし→現在の年を使用）
  const monthOnlyMatch = inputMonth.match(/(\d{1,2})月/);
  if (monthOnlyMatch) {
    const currentYear = new Date().getFullYear();
    return {
      year: currentYear,
      month: parseInt(monthOnlyMatch[1], 10)
    };
  }

  return null;
}

/**
 * 年月選択ダイアログを表示
 */
function showYearMonthDialog(isTest) {
  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 20px; }
      select { font-size: 16px; padding: 8px; margin: 5px; }
      button { font-size: 16px; padding: 10px 20px; margin: 10px 5px; cursor: pointer; }
      .ok { background: #4285f4; color: white; border: none; border-radius: 4px; }
      .cancel { background: #f1f1f1; border: 1px solid #ccc; border-radius: 4px; }
    </style>
    <h3>📅 公開予定月を選択</h3>
    <p>
      <select id="year">
        ${(() => {
          const currentYear = new Date().getFullYear();
          let options = '';
          for (let y = currentYear; y <= currentYear + 2; y++) {
            options += '<option value="' + y + '">' + y + '年</option>';
          }
          return options;
        })()}
      </select>
      <select id="month">
        ${(() => {
          let options = '';
          for (let m = 1; m <= 12; m++) {
            const selected = m === new Date().getMonth() + 1 ? ' selected' : '';
            options += '<option value="' + m + '"' + selected + '>' + m + '月</option>';
          }
          return options;
        })()}
      </select>
    </p>
    <p>
      <button class="ok" onclick="submit()">OK</button>
      <button class="cancel" onclick="google.script.host.close()">キャンセル</button>
    </p>
    <script>
      function submit() {
        const year = document.getElementById('year').value;
        const month = document.getElementById('month').value;
        google.script.run.withSuccessHandler(() => google.script.host.close())
          .processYearMonthSelection(parseInt(year), parseInt(month), ${isTest});
      }
    </script>
  `)
  .setWidth(300)
  .setHeight(200);

  SpreadsheetApp.getUi().showModalDialog(html, '月別構成案生成');
}

/**
 * 年月選択後の処理
 */
function processYearMonthSelection(year, month, isTest) {
  const ui = SpreadsheetApp.getUi();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = SpreadsheetApp.getActiveSheet();

  Logger.log(`年月を選択: ${year}年${month}月 (テスト: ${isTest})`);

  // シートからデータを取得
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    ui.alert('⚠️ エラー', 'データがありません。', ui.ButtonSet.OK);
    return;
  }

  const dataRange = sheet.getRange(2, 1, lastRow - 1, 9);
  const allData = dataRange.getValues();

  const targetMonth = `${month}月`;
  const targetMonthFull = `${year}年${month}月`;
  const keywords = [];

  // デバッグ: 最初の5行のデータを表示
  Logger.log(`=== デバッグ情報 ===`);
  Logger.log(`検索条件: targetMonth="${targetMonth}", targetMonthFull="${targetMonthFull}"`);
  Logger.log(`データ行数: ${allData.length}`);
  for (let i = 0; i < Math.min(5, allData.length); i++) {
    Logger.log(`行${i+2}: D列="${allData[i][3]}", G列="${allData[i][6]}", I列="${allData[i][8]}"`);
  }
  Logger.log(`===================`);

  for (let i = 0; i < allData.length; i++) {
    const rawPubMonth = allData[i][3];
    const keyword = String(allData[i][6]).trim();
    const status = String(allData[i][8]).trim();

    // D列の値を年月に変換（日付型とテキスト型の両方に対応）
    let pubYear = null;
    let pubMonthNum = null;

    if (rawPubMonth instanceof Date) {
      // 日付型の場合
      pubYear = rawPubMonth.getFullYear();
      pubMonthNum = rawPubMonth.getMonth() + 1; // 0始まりなので+1
    } else {
      // テキスト型の場合（「2026年1月」や「1月」形式）
      const pubMonthStr = String(rawPubMonth).trim();
      const fullMatch = pubMonthStr.match(/(\d{4})年(\d{1,2})月/);
      const shortMatch = pubMonthStr.match(/^(\d{1,2})月$/);

      if (fullMatch) {
        pubYear = parseInt(fullMatch[1]);
        pubMonthNum = parseInt(fullMatch[2]);
      } else if (shortMatch) {
        pubMonthNum = parseInt(shortMatch[1]);
      }
    }

    // デバッグ: 各行の判定結果
    if (i < 5) {
      Logger.log(`行${i+2}チェック: pubYear=${pubYear}, pubMonthNum=${pubMonthNum}, keyword="${keyword}", status="${status}"`);
    }

    if (!keyword) continue;
    if (status !== '') continue;

    // 年月が一致するかチェック
    const monthMatches = pubMonthNum === month;
    const yearMatches = pubYear === null || pubYear === year;

    if (monthMatches && yearMatches) {
      keywords.push(keyword);
      Logger.log(`✓ 行${i+2}をキーワードに追加: "${keyword}"`);
      if (isTest && keywords.length >= 1) break;
    }
  }

  if (keywords.length === 0) {
    ui.alert('⚠️ 対象なし', `${year}年${month}月の未処理キーワードがありません。`, ui.ButtonSet.OK);
    return;
  }

  const message = isTest
    ? `${year}年${month}月のキーワード1件をテスト生成します:\n\n• ${keywords[0]}`
    : `${year}年${month}月のキーワード${keywords.length}件の構成案を生成します。`;

  const confirm = ui.alert('確認', message + '\n\n続行しますか？', ui.ButtonSet.OK_CANCEL);
  if (confirm != ui.Button.OK) return;

  // Cloud Run呼び出し
  try {
    const response = UrlFetchApp.fetch(`${CLOUD_RUN_URL}/generate-outlines`, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({
        year: year,
        month: month,
        keywords: keywords,
        max_workers: isTest ? 1 : 10,
        master_spreadsheet_id: spreadsheet.getId(),
        keyword_column: 'G',
        url_column: 'M'
      }),
      muteHttpExceptions: true,
      timeout: 300
    });

    const responseCode = response.getResponseCode();
    if (responseCode === 200) {
      const result = JSON.parse(response.getContentText());
      ui.alert('✅ 完了', `構成案生成が完了しました。\n\n成功: ${result.created_sheets}件\n失敗: ${result.failed}件`, ui.ButtonSet.OK);
    } else {
      ui.alert('❌ エラー', `構成案生成に失敗しました (${responseCode}):\n${response.getContentText()}`, ui.ButtonSet.OK);
    }
  } catch (e) {
    ui.alert('❌ エラー', `エラーが発生しました:\n${e.message}`, ui.ButtonSet.OK);
  }
}

/**
 * 月別構成案生成（G列KW・D列公開予定月・I列ステータス）
 */
function generateOutlinesByMonth() {
  showYearMonthDialog(false);
}

/**
 * 月別構成案生成（1件テスト）
 */
function generateOutlinesByMonthTest() {
  showYearMonthDialog(true);
}

/**
 * D列の値と入力された月が一致するかチェック
 * 対応形式：
 * - Date オブジェクト（2026/01/01 など）
 * - 文字列（"2026年1月"、"1月"、"2026/1/1" など）
 * @param {Date|string} cellValue - D列の値
 * @param {string} inputMonth - 入力された月（例：1月、2026年1月）
 * @returns {boolean} 一致するかどうか
 */
function isMonthMatch(cellValue, inputMonth) {
  if (!cellValue) return false;

  // 入力から年と月を抽出
  let targetYear = null;
  let targetMonth = null;

  // "2026年1月" 形式
  const yearMonthMatch = inputMonth.match(/(\d{4})年(\d{1,2})月?/);
  if (yearMonthMatch) {
    targetYear = parseInt(yearMonthMatch[1]);
    targetMonth = parseInt(yearMonthMatch[2]);
  } else {
    // "1月" 形式（年指定なし）
    const monthMatch = inputMonth.match(/(\d{1,2})月?/);
    if (monthMatch) {
      targetMonth = parseInt(monthMatch[1]);
    }
  }

  if (!targetMonth) return false;

  // D列の値から年と月を抽出
  let cellYear = null;
  let cellMonth = null;

  if (cellValue instanceof Date) {
    // Dateオブジェクトの場合
    cellYear = cellValue.getFullYear();
    cellMonth = cellValue.getMonth() + 1; // 0-indexed なので +1
  } else {
    // 文字列の場合
    const cellStr = cellValue.toString();

    // "2026年1月" 形式
    const cellYearMonthMatch = cellStr.match(/(\d{4})年(\d{1,2})月?/);
    if (cellYearMonthMatch) {
      cellYear = parseInt(cellYearMonthMatch[1]);
      cellMonth = parseInt(cellYearMonthMatch[2]);
    } else {
      // "2026/1/1" や "2026-01-01" 形式
      const cellDateMatch = cellStr.match(/(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
      if (cellDateMatch) {
        cellYear = parseInt(cellDateMatch[1]);
        cellMonth = parseInt(cellDateMatch[2]);
      } else {
        // "1月" 形式
        const cellMonthMatch = cellStr.match(/(\d{1,2})月/);
        if (cellMonthMatch) {
          cellMonth = parseInt(cellMonthMatch[1]);
        }
      }
    }
  }

  if (!cellMonth) return false;

  // 月の一致チェック
  if (targetMonth !== cellMonth) return false;

  // 年が指定されている場合は年もチェック
  if (targetYear && cellYear && targetYear !== cellYear) return false;

  return true;
}

/**
 * 月別スプレッドシートを取得または作成
 * @param {string} month - 月（例：12月、2025年1月）
 * @returns {string} スプレッドシートID
 */
function getOrCreateMonthlySpreadsheet(month) {
  // 年月を正規化（例：12月 → 2025年12月）
  let normalizedMonth = month;
  if (!month.includes('年')) {
    const currentYear = new Date().getFullYear();
    normalizedMonth = `${currentYear}年${month}`;
  }

  // スプレッドシート名
  const spreadsheetName = `${normalizedMonth}_構成案`;

  // 既存のスプレッドシートを検索
  const files = DriveApp.getFilesByName(spreadsheetName);
  if (files.hasNext()) {
    const file = files.next();
    Logger.log(`既存のスプレッドシートを使用: ${spreadsheetName} (${file.getId()})`);
    return file.getId();
  }

  // 新規作成
  try {
    const newSpreadsheet = SpreadsheetApp.create(spreadsheetName);
    const spreadsheetId = newSpreadsheet.getId();
    Logger.log(`新しいスプレッドシートを作成: ${spreadsheetName} (${spreadsheetId})`);

    // 現在のスプレッドシートと同じフォルダに移動（オプション）
    const currentSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const currentFile = DriveApp.getFileById(currentSpreadsheet.getId());
    const parents = currentFile.getParents();
    if (parents.hasNext()) {
      const folder = parents.next();
      const newFile = DriveApp.getFileById(spreadsheetId);
      folder.addFile(newFile);
      DriveApp.getRootFolder().removeFile(newFile);
      Logger.log(`スプレッドシートを移動: ${folder.getName()}`);
    }

    return spreadsheetId;
  } catch (error) {
    Logger.log(`スプレッドシート作成エラー: ${error}`);
    return null;
  }
}

// ========================================
// 月別初稿生成（マスターシートから実行）
// ========================================

/**
 * 月別初稿生成（通常）
 */
function generateArticlesByMonth() {
  showArticleYearMonthDialog(false);
}

/**
 * 月別初稿生成（1件テスト）
 */
function generateArticlesByMonthTest() {
  showArticleYearMonthDialog(true);
}

/**
 * 年月選択ダイアログを表示（初稿生成用）
 */
function showArticleYearMonthDialog(isTest) {
  const currentYear = new Date().getFullYear();
  const years = [];
  for (let y = currentYear; y <= currentYear + 2; y++) {
    years.push(y);
  }

  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 20px; }
      select { font-size: 16px; padding: 5px; margin: 5px; }
      button { font-size: 16px; padding: 10px 20px; margin-top: 15px; cursor: pointer; }
      .container { text-align: center; }
    </style>
    <div class="container">
      <p>初稿を生成する構成案スプシの年月を選択：</p>
      <select id="year">
        ${years.map(y => `<option value="${y}">${y}年</option>`).join('')}
      </select>
      <select id="month">
        ${[...Array(12)].map((_, i) => `<option value="${i + 1}">${i + 1}月</option>`).join('')}
      </select>
      <br>
      <button onclick="submit()">初稿生成開始</button>
    </div>
    <script>
      function submit() {
        const year = document.getElementById('year').value;
        const month = document.getElementById('month').value;
        google.script.run
          .withSuccessHandler(() => google.script.host.close())
          .withFailureHandler(err => alert('エラー: ' + err))
          .processArticleYearMonthSelection(parseInt(year), parseInt(month), ${isTest});
      }
    </script>
  `)
  .setWidth(300)
  .setHeight(200);

  SpreadsheetApp.getUi().showModalDialog(html, isTest ? '月別初稿生成（1件テスト）' : '月別初稿生成');
}

/**
 * 年月選択後の処理（初稿生成）
 */
function processArticleYearMonthSelection(year, month, isTest) {
  const ui = SpreadsheetApp.getUi();
  const spreadsheetName = `${year}年${month}月`;

  Logger.log(`初稿生成: ${spreadsheetName} (テスト: ${isTest})`);

  // 構成案フォルダID（Cloud Runと同じフォルダ）
  const outlineFolderId = 'YOUR_OUTLINE_FOLDER_ID';

  // フォルダ内で構成案スプシを検索
  let outlineSpreadsheetId = null;
  try {
    const folder = DriveApp.getFolderById(outlineFolderId);
    const files = folder.getFilesByName(spreadsheetName);

    if (files.hasNext()) {
      outlineSpreadsheetId = files.next().getId();
      Logger.log(`構成案スプシを発見: ${spreadsheetName} (${outlineSpreadsheetId})`);
    }
  } catch (error) {
    Logger.log(`フォルダ検索エラー: ${error}`);
  }

  if (!outlineSpreadsheetId) {
    ui.alert('⚠️ 構成案なし', `「${spreadsheetName}」の構成案スプシが見つかりません。\n\n先に構成案を生成してください。`, ui.ButtonSet.OK);
    return;
  }

  // 確認ダイアログ
  const message = isTest
    ? `「${spreadsheetName}」の構成案から1件だけ初稿を生成します。`
    : `「${spreadsheetName}」の構成案から初稿を一括生成します。`;

  const confirm = ui.alert('確認', message + '\n\n続行しますか？', ui.ButtonSet.OK_CANCEL);
  if (confirm != ui.Button.OK) return;

  // Cloud Run呼び出し
  try {
    const endpoint = '/generate-articles';  // 常にgenerate-articlesを使用
    const payload = {
      spreadsheet_id: outlineSpreadsheetId,
      master_spreadsheet_id: MASTER_SPREADSHEET_ID,
      keyword_column: 'G',
      article_url_column: 'N',
      max_articles: isTest ? 1 : null  // テスト時は1件のみ
    };

    Logger.log(`Cloud Run呼び出し: ${CLOUD_RUN_URL}${endpoint}`);
    Logger.log(`ペイロード: ${JSON.stringify(payload)}`);

    const response = UrlFetchApp.fetch(`${CLOUD_RUN_URL}${endpoint}`, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    Logger.log(`レスポンスコード: ${responseCode}`);
    Logger.log(`レスポンス: ${responseText}`);

    if (responseCode === 200) {
      const result = JSON.parse(responseText);
      const processedCount = result.processed ? result.processed.length : 0;
      const errorCount = result.errors ? result.errors.length : 0;

      let resultMessage = `✅ 初稿生成完了\n\n`;
      resultMessage += `処理件数: ${processedCount}件\n`;
      if (errorCount > 0) {
        resultMessage += `エラー: ${errorCount}件\n`;
      }

      if (result.processed && result.processed.length > 0) {
        resultMessage += `\n生成された記事:\n`;
        result.processed.slice(0, 5).forEach(item => {
          resultMessage += `• ${item.title || item.sheet}\n`;
        });
        if (result.processed.length > 5) {
          resultMessage += `... 他 ${result.processed.length - 5}件`;
        }
      }

      ui.alert('完了', resultMessage, ui.ButtonSet.OK);
    } else if (responseCode === 202) {
      // バックグラウンド処理開始
      const result = JSON.parse(responseText);
      ui.alert('🚀 処理開始', `初稿生成をバックグラウンドで開始しました。\n\n完了までしばらくお待ちください。\n（処理状況はCloud Runログで確認できます）`, ui.ButtonSet.OK);
    } else {
      ui.alert('❌ エラー', `初稿生成に失敗しました (${responseCode}):\n\n${responseText}`, ui.ButtonSet.OK);
    }

  } catch (error) {
    Logger.log(`エラー: ${error}`);
    ui.alert('❌ エラー', `初稿生成中にエラーが発生しました:\n\n${error}`, ui.ButtonSet.OK);
  }
}

// ========================================
// 全未処理シート並列生成
// ========================================

/**
 * 全記事一括生成（Cloud Tasks使用・100件対応）
 */
function generateAllArticlesBulk() {
  const ui = SpreadsheetApp.getUi();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const spreadsheetId = spreadsheet.getId();

  // まず未処理件数を確認
  try {
    const countResponse = UrlFetchApp.fetch(`${CLOUD_RUN_URL}/get-unprocessed-count`, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({
        spreadsheet_id: spreadsheetId
      }),
      muteHttpExceptions: true,
      timeout: 60
    });

    const countCode = countResponse.getResponseCode();
    if (countCode !== 200) {
      ui.alert('❌ エラー', `未処理件数の取得に失敗しました (${countCode})`, ui.ButtonSet.OK);
      return;
    }

    const countData = JSON.parse(countResponse.getContentText());
    const unprocessedCount = countData.count;
    const sheetNames = countData.sheets || [];

    if (unprocessedCount === 0) {
      ui.alert('ℹ️ 情報', '未処理のシートがありません。\n\nすべての構成案は既に初稿生成済みです。', ui.ButtonSet.OK);
      return;
    }

    // 予想時間を計算（1件約5分、同時3件処理）
    const estimatedMinutes = Math.ceil(unprocessedCount * 5 / 3);
    const estimatedHours = Math.floor(estimatedMinutes / 60);
    const estimatedMins = estimatedMinutes % 60;
    let timeStr = '';
    if (estimatedHours > 0) {
      timeStr = `約${estimatedHours}時間${estimatedMins}分`;
    } else {
      timeStr = `約${estimatedMins}分`;
    }

    // 確認ダイアログ
    let confirmMessage = `🚀 一括生成を開始しますか？\n\n`;
    confirmMessage += `📝 未処理記事数: ${unprocessedCount}件\n`;
    confirmMessage += `⏱️ 予想所要時間: ${timeStr}\n`;
    confirmMessage += `🔄 同時処理数: 3件\n\n`;

    if (sheetNames.length > 0) {
      confirmMessage += `対象シート（最初の5件）:\n`;
      sheetNames.slice(0, 5).forEach(name => {
        confirmMessage += `  • ${name}\n`;
      });
      if (sheetNames.length > 5) {
        confirmMessage += `  ... 他${sheetNames.length - 5}件\n`;
      }
    }

    confirmMessage += `\n✅ 各記事完了時にSlack通知が届きます。`;
    confirmMessage += `\n💡 このダイアログを閉じても処理は継続します。`;

    const confirm = ui.alert('確認', confirmMessage, ui.ButtonSet.OK_CANCEL);
    if (confirm != ui.Button.OK) return;

    // Cloud Tasksにキュー登録
    Logger.log(`一括生成開始: ${unprocessedCount}件をCloud Tasksに登録`);

    const response = UrlFetchApp.fetch(`${CLOUD_RUN_URL}/enqueue-all-articles`, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({
        spreadsheet_id: spreadsheetId,
        master_spreadsheet_id: MASTER_SPREADSHEET_ID,
        keyword_column: 'G',
        article_url_column: 'N'
      }),
      muteHttpExceptions: true,
      timeout: 120  // キュー登録に時間がかかる場合あり
    });

    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    Logger.log(`レスポンスコード: ${responseCode}`);
    Logger.log(`レスポンス: ${responseText}`);

    if (responseCode === 200) {
      const result = JSON.parse(responseText);

      if (result.status === 'no_work') {
        ui.alert('ℹ️ 情報', '未処理のシートがありません。', ui.ButtonSet.OK);
        return;
      }

      ui.alert(
        '🚀 一括生成開始',
        `${result.queued}件の初稿生成をキューに登録しました！\n\n` +
        `⏱️ 予想所要時間: ${timeStr}\n` +
        `🔄 同時処理数: 3件\n\n` +
        `✅ 各記事完了時にSlackに通知が届きます。\n` +
        `💡 このスプレッドシートを閉じても処理は継続します。`,
        ui.ButtonSet.OK
      );
    } else {
      ui.alert('❌ エラー', `一括生成の開始に失敗しました (${responseCode}):\n\n${responseText}`, ui.ButtonSet.OK);
    }

  } catch (error) {
    Logger.log(`エラー: ${error}`);
    ui.alert('❌ エラー', `一括生成中にエラーが発生しました:\n\n${error}`, ui.ButtonSet.OK);
  }
}

// ============================================
// Claude API 月別生成機能
// ============================================

/**
 * 月別構成案生成（Claude）
 */
function generateOutlinesByMonthClaude() {
  showYearMonthDialogClaude(false);
}

/**
 * 月別構成案生成（Claude・1件テスト）
 */
function generateOutlinesByMonthClaudeTest() {
  showYearMonthDialogClaude(true);
}

/**
 * 年月選択ダイアログを表示（Claude構成案用）
 */
function showYearMonthDialogClaude(isTest) {
  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 20px; }
      select { font-size: 16px; padding: 8px; margin: 5px; }
      button { font-size: 16px; padding: 10px 20px; margin: 10px 5px; cursor: pointer; }
      .ok { background: #8B5CF6; color: white; border: none; border-radius: 4px; }
      .cancel { background: #f1f1f1; border: 1px solid #ccc; border-radius: 4px; }
      h3 { color: #8B5CF6; }
    </style>
    <h3>🤖 Claude構成案生成</h3>
    <p>公開予定月を選択：</p>
    <p>
      <select id="year">
        ${(() => {
          const currentYear = new Date().getFullYear();
          let options = '';
          for (let y = currentYear; y <= currentYear + 2; y++) {
            options += '<option value="' + y + '">' + y + '年</option>';
          }
          return options;
        })()}
      </select>
      <select id="month">
        ${(() => {
          let options = '';
          for (let m = 1; m <= 12; m++) {
            const selected = m === new Date().getMonth() + 1 ? ' selected' : '';
            options += '<option value="' + m + '"' + selected + '>' + m + '月</option>';
          }
          return options;
        })()}
      </select>
    </p>
    <p>
      <button class="ok" onclick="submit()">生成開始</button>
      <button class="cancel" onclick="google.script.host.close()">キャンセル</button>
    </p>
    <script>
      function submit() {
        const year = document.getElementById('year').value;
        const month = document.getElementById('month').value;
        google.script.run.withSuccessHandler(() => google.script.host.close())
          .processYearMonthSelectionClaude(parseInt(year), parseInt(month), ${isTest});
      }
    </script>
  `)
  .setWidth(320)
  .setHeight(220);

  SpreadsheetApp.getUi().showModalDialog(html, '🤖 Claude構成案生成');
}

/**
 * 年月選択後の処理（Claude構成案）
 */
function processYearMonthSelectionClaude(year, month, isTest) {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();

  Logger.log(`[Claude] 年月を選択: ${year}年${month}月 (テスト: ${isTest})`);

  // シートからデータを取得
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    ui.alert('⚠️ エラー', 'データがありません。', ui.ButtonSet.OK);
    return;
  }

  const dataRange = sheet.getRange(2, 1, lastRow - 1, 9);
  const allData = dataRange.getValues();

  const keywords = [];

  for (let i = 0; i < allData.length; i++) {
    const rawPubMonth = allData[i][3]; // D列：公開予定月
    const keyword = String(allData[i][6]).trim(); // G列：キーワード
    const status = String(allData[i][8]).trim(); // I列：ステータス

    // 公開予定月の判定
    let pubYear = null;
    let pubMonthNum = null;

    if (rawPubMonth instanceof Date) {
      pubYear = rawPubMonth.getFullYear();
      pubMonthNum = rawPubMonth.getMonth() + 1;
    } else {
      const pubMonthStr = String(rawPubMonth).trim();
      const fullMatch = pubMonthStr.match(/(\d{4})年(\d{1,2})月/);
      const shortMatch = pubMonthStr.match(/^(\d{1,2})月$/);

      if (fullMatch) {
        pubYear = parseInt(fullMatch[1]);
        pubMonthNum = parseInt(fullMatch[2]);
      } else if (shortMatch) {
        pubMonthNum = parseInt(shortMatch[1]);
        pubYear = year; // 年指定なしの場合は選択した年を使用
      }
    }

    // 条件判定
    const monthMatch = pubMonthNum === month && (pubYear === null || pubYear === year);
    const hasKeyword = keyword && keyword !== '';
    const notProcessed = status === '' || status === '未処理';

    if (monthMatch && hasKeyword && notProcessed) {
      keywords.push(keyword);
      if (isTest && keywords.length >= 1) break;
    }
  }

  if (keywords.length === 0) {
    ui.alert('⚠️ 対象なし', `${year}年${month}月の未処理キーワードがありません。`, ui.ButtonSet.OK);
    return;
  }

  const confirmMsg = isTest
    ? `${year}年${month}月のキーワード1件をClaudeで構成案生成します。\n\nキーワード: ${keywords[0]}`
    : `${year}年${month}月のキーワード${keywords.length}件をClaudeで構成案生成します。`;

  const result = ui.alert('🤖 Claude構成案生成', confirmMsg, ui.ButtonSet.YES_NO);
  if (result !== ui.Button.YES) return;

  // 各キーワードを順番に処理
  let successCount = 0;
  let failCount = 0;
  const results = [];

  for (const keyword of keywords) {
    try {
      Logger.log(`[Claude] 構成案生成: ${keyword}`);

      const response = UrlFetchApp.fetch(`${CLOUD_RUN_URL}/generate-outline-claude`, {
        method: 'POST',
        contentType: 'application/json',
        payload: JSON.stringify({ keyword: keyword }),
        muteHttpExceptions: true
      });

      const responseCode = response.getResponseCode();
      const data = JSON.parse(response.getContentText());

      if (responseCode === 200 && data.success) {
        // 結果をシートに出力
        const resultSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(`Claude_${keyword.substring(0, 20)}`);
        resultSheet.getRange('A1').setValue('キーワード');
        resultSheet.getRange('B1').setValue(data.keyword);
        resultSheet.getRange('A2').setValue('共起語');
        resultSheet.getRange('B2').setValue((data.related_keywords || []).join(', '));
        resultSheet.getRange('A3').setValue('上位URL');
        resultSheet.getRange('B3').setValue((data.top_urls || []).join('\n'));
        resultSheet.getRange('A4').setValue('トークン');
        resultSheet.getRange('B4').setValue(data.usage ? data.usage.total_tokens : 'N/A');
        resultSheet.getRange('A6').setValue('構成案');
        resultSheet.getRange('A7').setValue(data.outline);
        resultSheet.getRange('A7').setWrap(true);

        successCount++;
        results.push(`✅ ${keyword}`);
      } else {
        failCount++;
        results.push(`❌ ${keyword}: ${data.error || 'エラー'}`);
      }
    } catch (error) {
      failCount++;
      results.push(`❌ ${keyword}: ${error}`);
    }
  }

  ui.alert('🤖 Claude構成案生成完了',
    `成功: ${successCount}件\n失敗: ${failCount}件\n\n${results.join('\n')}`,
    ui.ButtonSet.OK);
}

/**
 * 月別初稿生成（Claude）
 */
function generateArticlesByMonthClaude() {
  showYearMonthDialogClaudeDraft(false);
}

/**
 * 月別初稿生成（Claude・1件テスト）
 */
function generateArticlesByMonthClaudeTest() {
  showYearMonthDialogClaudeDraft(true);
}

/**
 * 年月選択ダイアログを表示（Claude初稿用）
 */
function showYearMonthDialogClaudeDraft(isTest) {
  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 20px; }
      select { font-size: 16px; padding: 8px; margin: 5px; }
      button { font-size: 16px; padding: 10px 20px; margin: 10px 5px; cursor: pointer; }
      .ok { background: #8B5CF6; color: white; border: none; border-radius: 4px; }
      .cancel { background: #f1f1f1; border: 1px solid #ccc; border-radius: 4px; }
      h3 { color: #8B5CF6; }
    </style>
    <h3>🤖 Claude初稿生成</h3>
    <p>構成案の年月を選択：</p>
    <p>
      <select id="year">
        ${(() => {
          const currentYear = new Date().getFullYear();
          let options = '';
          for (let y = currentYear; y <= currentYear + 2; y++) {
            options += '<option value="' + y + '">' + y + '年</option>';
          }
          return options;
        })()}
      </select>
      <select id="month">
        ${(() => {
          let options = '';
          for (let m = 1; m <= 12; m++) {
            const selected = m === new Date().getMonth() + 1 ? ' selected' : '';
            options += '<option value="' + m + '"' + selected + '>' + m + '月</option>';
          }
          return options;
        })()}
      </select>
    </p>
    <p>
      <button class="ok" onclick="submit()">生成開始</button>
      <button class="cancel" onclick="google.script.host.close()">キャンセル</button>
    </p>
    <script>
      function submit() {
        const year = document.getElementById('year').value;
        const month = document.getElementById('month').value;
        google.script.run.withSuccessHandler(() => google.script.host.close())
          .processYearMonthSelectionClaudeDraft(parseInt(year), parseInt(month), ${isTest});
      }
    </script>
  `)
  .setWidth(320)
  .setHeight(220);

  SpreadsheetApp.getUi().showModalDialog(html, '🤖 Claude初稿生成');
}

/**
 * 年月選択後の処理（Claude初稿）
 */
function processYearMonthSelectionClaudeDraft(year, month, isTest) {
  const ui = SpreadsheetApp.getUi();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  Logger.log(`[Claude] 初稿生成: ${year}年${month}月 (テスト: ${isTest})`);

  // Claude構成案シートを検索（現在のスプレッドシート内）
  const sheets = spreadsheet.getSheets();
  const claudeSheets = sheets.filter(s => s.getName().startsWith('Claude_'));

  if (claudeSheets.length === 0) {
    ui.alert('⚠️ 構成案なし', 'Claude構成案シートが見つかりません。\n\n先に「🤖 月別構成案生成（Claude）」を実行してください。', ui.ButtonSet.OK);
    return;
  }

  // 未処理のシートを取得（初稿シートが存在しないもの）
  const pendingSheets = [];
  for (const sheet of claudeSheets) {
    const sheetName = sheet.getName();
    const keyword = sheet.getRange('B1').getValue();
    const draftSheetName = `Claude初稿_${keyword}`;

    // 初稿シートが存在しないものを対象
    if (!spreadsheet.getSheetByName(draftSheetName)) {
      pendingSheets.push({ sheet, keyword, sheetName });
      if (isTest && pendingSheets.length >= 1) break;
    }
  }

  if (pendingSheets.length === 0) {
    ui.alert('⚠️ 対象なし', '未処理のClaude構成案シートがありません。', ui.ButtonSet.OK);
    return;
  }

  const confirmMsg = isTest
    ? `1件の構成案からClaude初稿を生成します。\n\nシート: ${pendingSheets[0].sheetName}`
    : `${pendingSheets.length}件の構成案からClaude初稿を生成します。`;

  const result = ui.alert('🤖 Claude初稿生成', confirmMsg, ui.ButtonSet.YES_NO);
  if (result !== ui.Button.YES) return;

  // 各構成案シートを処理
  let successCount = 0;
  let failCount = 0;
  const results = [];

  for (const { sheet, keyword, sheetName } of pendingSheets) {
    try {
      Logger.log(`[Claude] 初稿生成: ${keyword}`);

      // 構成案シートからデータを取得
      const outlineText = sheet.getRange('A7').getValue();
      const lines = outlineText.split('\n');

      // H1タイトルを取得
      let h1Title = '';
      for (const line of lines) {
        const trimmed = line.trim();
        if (trimmed.startsWith('H1:') || trimmed.startsWith('H1：')) {
          h1Title = trimmed.replace(/^H1[:：]\s*/, '');
          break;
        }
      }

      if (!h1Title) {
        h1Title = keyword; // H1がなければキーワードを使用
      }

      // 見出し構造を解析
      const headings = [];
      for (const line of lines) {
        const trimmed = line.trim();
        if (trimmed.startsWith('H2:') || trimmed.startsWith('H2：')) {
          headings.push({ level: 'H2', text: trimmed.replace(/^H2[:：]\s*/, '') });
        } else if (trimmed.startsWith('H3:') || trimmed.startsWith('H3：')) {
          headings.push({ level: 'H3', text: trimmed.replace(/^H3[:：]\s*/, '') });
        } else if (trimmed.startsWith('H4:') || trimmed.startsWith('H4：')) {
          headings.push({ level: 'H4', text: trimmed.replace(/^H4[:：]\s*/, '') });
        }
      }

      if (headings.length === 0) {
        failCount++;
        results.push(`❌ ${keyword}: 見出しなし`);
        continue;
      }

      // Claude APIで初稿生成
      const response = UrlFetchApp.fetch(`${CLOUD_RUN_URL}/generate-draft-claude`, {
        method: 'POST',
        contentType: 'application/json',
        payload: JSON.stringify({
          keyword: keyword,
          h1_title: h1Title,
          headings: headings
        }),
        muteHttpExceptions: true
      });

      const responseCode = response.getResponseCode();
      const data = JSON.parse(response.getContentText());

      if (responseCode === 200 && data.success) {
        // 結果をシートに出力
        const resultSheet = spreadsheet.insertSheet(`Claude初稿_${keyword.substring(0, 15)}`);
        resultSheet.getRange('A1').setValue('キーワード');
        resultSheet.getRange('B1').setValue(data.keyword);
        resultSheet.getRange('A2').setValue('H1タイトル');
        resultSheet.getRange('B2').setValue(data.h1_title);
        resultSheet.getRange('A3').setValue('文字数');
        resultSheet.getRange('B3').setValue(data.char_count);
        resultSheet.getRange('A4').setValue('トークン');
        resultSheet.getRange('B4').setValue(data.usage ? data.usage.total_tokens : 'N/A');
        resultSheet.getRange('A6').setValue('初稿');
        resultSheet.getRange('A7').setValue(data.draft);
        resultSheet.getRange('A7').setWrap(true);

        successCount++;
        results.push(`✅ ${keyword} (${data.char_count}字)`);
      } else {
        failCount++;
        results.push(`❌ ${keyword}: ${data.error || 'エラー'}`);
      }
    } catch (error) {
      failCount++;
      results.push(`❌ ${keyword}: ${error}`);
    }
  }

  ui.alert('🤖 Claude初稿生成完了',
    `成功: ${successCount}件\n失敗: ${failCount}件\n\n${results.join('\n')}`,
    ui.ButtonSet.OK);
}
