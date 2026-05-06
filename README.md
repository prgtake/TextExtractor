# TextExtractor1 v1.0.0

Google Gemini APIのマルチモーダル能力を最大限に引き出し、あらゆる書類から「正確な事実」と「高度なAI推論」を同時に取り出すためのデータ抽出・解析プラットフォームです。

## 1. 主な機能と特徴
- **事実と推論のハイブリッド抽出**: 提示された資料から「そのまま書き写す」項目と、AIが「知識や検索を元に考える」項目を明確に分離。
- **圧倒的なファイル対応力**: 画像、PDF、Office製品（新旧）、メール、一太郎（.jtd）に対応。
- **スマートな設定管理**: APIキーやモデル名を `app_config.json` に自動保存。
- **最新情報の補完**: Google検索やRAG（File Search Store）との連携。

## 2. 対応ファイル形式
- **画像/PDF**: .png, .jpg, .jpeg, .webp, .bmp, .pdf
- **Office文書**: .docx, .xlsx, .xls, .pptx, .doc, .ppt
- **一太郎**: .jtd
- **メール**: .eml, .msg (Outlook)
- **テキスト**: .txt, .csv, .log

## 3. 指示ファイル（プロンプト）の書き方
指示ファイル（.txt）は以下の構成にしてください。

**※重要：文字コードは「UTF-8 (BOMなし/UTF-8N)」で保存してください。** Shift-JISなど他の形式では文字化けが発生し、AIが正しく解析できない場合があります。

1. **概要（任意）**: 処理対象の説明。
2. **出力項目（必須）**: `出力項目: 項目1, 項目2` の形式で記述。
3. **処理指示（必須）**:
   - `【抽出】`: 資料の内容をそのまま書き写す（検索なし）。
   - `【生成】`: AIが推論やWEB検索を用いて生成する。

### 記入例
```text
出力項目: 会社名, 氏名, 性格の予想, 会社の最新トピック

・会社名、氏名：【抽出】
・性格の予想：【生成】
・会社の最新トピック：【生成】
```

## 4. 利用上の大前提・限界（重要）
本ツールは「ファイル単位」で処理を行うため、以下の仕様・制限があります。
- **1ファイル内での完結**: 「【抽出】」項目は、そのファイルの中に全ての情報が含まれている必要があります。複数ファイル（例：名刺の表と裏が別画像）にまたがる情報を1つのレコードとして統合して「抽出」することはできません。
  - *対策*: 複数枚にまたがる場合は、あらかじめ1つのPDFに結合してから処理してください。
- **ファイル間の名寄せ**: 異なるファイル同士の情報を突き合わせて「名寄せ」を行う機能はありません。
- **情報の統合（生成）**: 「【生成】」項目においてRAG機能を併用する場合のみ、ナレッジベース内の複数資料を横断的に参照して回答を生成することが可能です。

## 5. 操作手順
1. アプリを起動し、APIキーを入力（初回のみ）。
2. 「対象フォルダを選択」でファイル群を選択。
3. 「指示ファイルを選択」でプロンプトを選択。
4. 「指示ファイルを実行」で一括処理を開始。
   - 実行後、選択した**対象フォルダ内**に `[指示ファイル名]_DB.db` という名前でSQLiteデータベースが自動生成されます。
5. 「DBをCSV出力」で結果を書き出し。

5. 再実行とエラーへの対応
- **自動スキップ機能**: 実行時にデータベース（.db）をチェックし、既に結果が格納されているファイルは自動的にスキップします。途中でエラーが起きた場合や、後からファイルを追加した場合でも、重複を気にせず再実行できます。
- **堅牢なリトライ処理**: APIの過負荷（503）や回数制限（429）が発生した場合、指数バックオフ（待ち時間を段階的に増やす仕組み）を用いて、最大5回まで自動で再試行します。

## 6. トークンの消費とコストについて
本アプリはGoogle Gemini APIを使用するため、処理内容に応じてトークンが消費されます。

- **画像・PDF**: ページ数や解像度に応じてトークンを消費します。
- **Office・テキスト系ファイル**: 事前にテキスト抽出を行うため、画像として送るよりもトークンを大幅に節約できます。
- **WEB検索・RAG**: 有効にすると、検索結果の読み込みにより消費量が増加します。

Gemini 1.5 Flashなどの軽量モデルを使用することで、高い処理能力を維持しつつコストを最小限に抑えることが可能です。

## 6. ライセンス / License
Copyright (c) 2026 Datan (データン)  
Licensed under the MIT License.

---

# English Manual

**TextExtractor1** is a data extraction and analysis platform designed to leverage the full multimodal capabilities of the Google Gemini API to simultaneously retrieve "accurate facts" and "advanced AI insights" from any document.

## 1. Key Features
- **Hybrid Extraction of Facts and Reasoning**: Clearly separates items to be "transcribed as-is" from those the AI "thinks about based on knowledge or search."
- **Unmatched File Compatibility**: Extracts text information from images, PDFs, Office products (new and old), emails, and even "Ichitaro" (.jtd) files.
- **Smart Settings Management**: Automatically saves your API key and preferred Gemini model name, allowing you to start analysis instantly.
- **Latest Information Supplementation**: By checking a box in the GUI, the AI autonomously references Google Search or specific knowledge bases (RAG) to generate enriched information.

## 2. Supported File Formats
- **Images/PDF**: .png, .jpg, .jpeg, .webp, .bmp, .pdf
- **Office Documents**: .docx, .xlsx, .xls, .pptx, .doc, .ppt
- **Ichitaro**: .jtd
- **Email**: .eml, .msg (Outlook)
- **Text/Data**: .txt, .csv, .log

## 3. Golden Rules for Instruction Files (Prompts)
The "Instruction File (.txt)" that determines the AI's behavior should follow this structure:

**Important: Save the file using "UTF-8 (without BOM)" encoding.** Using other formats like Shift-JIS may cause character corruption, preventing the AI from processing the instructions correctly.

1. **Overview (Optional)**: A brief description of the task.
2. **Output Items (Required)**: List item names after "**出力項目:**" separated by commas. These will become the columns in your database.
3. **Processing Instructions (Required)**: Specify each item using one of the following:
   - `【抽出】` (Extraction): Transcribes text directly from the material (no search performed).
   - `【生成】` (Generation): AI performs reasoning or uses Web Search to supplement information.

### Prompt Example
```text
Analyze this business card.

出力項目: Company, Name, Personality Forecast, Latest Company News

- Company, Name: 【抽出】 (Extract exactly as seen.)
- Personality Forecast: 【生成】 (AI infers based on title and design.)
- Latest Company News: 【生成】 (Use Google Search to summarize recent topics.)
```

## 4. Key Limitations & Prerequisites (Important)
Since this tool processes documents on a **"per-file"** basis, the following limitations apply:
- **Single File Completion**: "【Extraction】" items must have all their information contained within a single file. The tool cannot integrate data spread across multiple files (e.g., separate images for the front and back of a business card) into a single record.
  - *Workaround*: If data is spread across multiple pages/images, merge them into a single PDF before processing.
- **No Cross-File Record Matching**: There is no built-in function to match or merge records between different files.
- **Cross-File Information Integration (Generation)**: Answer generation across multiple documents is only possible for "【Generation】" items when using the RAG feature with a pre-configured knowledge base.

## 5. Operating Procedures
1. Launch the app and enter your API key when prompted (first time only).
2. Click **"Select Target Folder"** and choose the folder containing the files you want to analyze.
3. Click **"Select Instruction File"** and choose your prepared prompt (.txt).
4. Verify the model name and turn on **"Web Search"** or **"RAG"** if needed.
5. Click **"Run Instruction File"** to start batch processing of all files.
6. After processing, click **"Export DB to CSV"** to save the accumulated data for use in Excel.

5. Resumption and Error Handling
- **Automatic Skip Function**: Checks the database (.db) at runtime and automatically skips files that already have results stored. You can restart without worrying about duplicates if an error occurs mid-process or if you add files later.
- **Robust Retry Processing**: If API overload (503) or rate limits (429) occur, the app automatically retries up to 5 times using exponential backoff (gradually increasing wait times).

## 6. License
Copyright (c) 2026 Datan (データン)

### MIT License
Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
