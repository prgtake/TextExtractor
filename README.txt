TextExtractor (by.gemini) v1.1.0 完全マニュアル

Google Gemini APIのマルチモーダル能力を最大限に引き出し、あらゆる書類から「正確な事実」と「高度なAI推論」を同時に取り出すためのデータ抽出・解析プラットフォームです。

--------------------------------------------------
1. 主な機能と特徴
--------------------------------------------------
■ 事実と推論のハイブリッド抽出
   提示された資料から「そのまま書き写す」項目と、AIが「知識や検索を元に考える」項目を明確に分離して処理します。

■ 圧倒的なファイル対応力
   画像やPDFはもちろん、Office製品（新旧）、メール、さらに「一太郎」まで、あらゆる形式の文書からテキスト情報を吸い上げます。

■ スマートな設定管理
   APIキーや使用するGeminiモデル名は一度入力すれば自動保存。専門知識がなくても、起動してすぐに解析を始められます。

■ 最新情報の補完（WEB検索/RAG）
   GUIのチェックを入れるだけで、AIが自律的にGoogle検索や特定の知識ベース（RAG）を参照し、情報を豊かに生成します。

--------------------------------------------------
2. 対応ファイル形式
--------------------------------------------------
・画像/PDF: .png, .jpg, .jpeg, .webp, .bmp, .pdf
・Office文書: .docx, .xlsx, .xls, .pptx, .doc, .ppt
・一太郎: .jtd
・メール: .eml, .msg (Outlook)
・テキスト: .txt, .csv, .log

--------------------------------------------------
3. コラム：AIアシスタント（Gemini CLI等）との違い
--------------------------------------------------
「Gemini CLIなどのAIアシスタントにフォルダを渡すのと何が違うの？」という疑問への回答です。

■ AIアシスタント（研究者型）
   フォルダ全体の構造を把握し、対話を通じて特定のファイルの内容を調べたり、コードを書いたりするのが得意です。しかし、数百個のファイルを「すべて、正確に、漏れなく」データ化し続ける単純作業は、記憶の限界やコスト効率の面で不向きです。

■ 本ツール（実務作業員型）
   フォルダ内の「全ファイルの中身」を1つずつ（または数個ずつ）読み込み、指定された形式でデータベースに書き出すことに特化しています。APIエラー時の自動再試行や、処理済みファイルのスキップ機能を備えており、大量のデータを確実に仕分ける「自動化マシン」として設計されています。

【一括モードの特殊な役割】
   本アプリの「一括モード」は、単なる自動化（順次処理）ではなく、「複数のファイルをAIに同時に見せて、全体を統合した判断をさせる」ための機能です。
   ・通常モード：1枚1枚を正確にデータ化（例：100枚の名刺から100行のリスト作成）。
   ・一括モード：複数枚をまとめて解析（例：5枚の領収書を合算して1つの支払いデータとして統合）。

--------------------------------------------------
4. 指示ファイル（プロンプト）の黄金律
--------------------------------------------------
AIの動作を決定づける「指示ファイル（.txt）」は、以下の構成にしてください。
※アプリ内の「書き方ヘルプ」ボタンでも確認できます。

**※重要：文字コードは「UTF-8 (BOMなし/UTF-8N)」で保存してください。** Shift-JISなど他の形式では文字化けが発生し、AIが正しく解析できない場合があります。

① 概要（任意）
   例：名刺を解析して、人物像と会社の動向をまとめます。

② SQLite利用設定（任意）
   外部DBを参照する場合、「###使用するSQLiteのパス」に続けて【絶対パス】を記述してください。
   例： "C:\data\MasterData.db"

③ 出力項目（必須）
   「出力項目:」の後に項目名をカンマ区切りで。これがDBの列になります。
   例：出力項目: 会社名, 氏名, 業界の評判

④ 処理指示（必須）
   各項目を以下のどちらかで指定してください。
   ・【抽出】: 資料にある文字をそのまま書き写します（検索はしません）。
   ・【生成】: AIが推論したり、WEB検索を使って調べたりします。

--- 指示ファイルの記入例 ---
###使用するSQLiteのパス
"C:\data\MasterData.db"

出力項目: 会社名, 氏名, 役職, 取引ランク, 会社の最新トピック

・会社名、氏名、役職：【抽出】
・取引ランク：【生成】SQLiteツールでMasterData.顧客マスタを参照してください。
・会社の最新トピック：【生成】（Google検索を利用して最新情報を要約する）
-------------------------

--------------------------------------------------
5. 操作手順
--------------------------------------------------
① アプリを起動し、APIキーを求められたら入力します（初回のみ）。
② 「対象フォルダを選択」で、解析したいファイル群が入ったフォルダを選びます。
③ 「指示ファイルを選択」で、作成したプロンプトファイルを選びます。
④ 使用したいモデル名を確認し、必要なら「WEB検索」「RAG」をオンにします。
⑤ 「指示ファイルを実行」を押すと、全ファイルの一括処理が始まります。
   ※処理結果は、選択した対象フォルダ内に「[指示ファイル名]_DB.db」として保存されます。
⑥ 処理完了後、「DBをCSV出力」でExcel等で開けるデータとして保存できます。

--------------------------------------------------
6. トークンの消費について
--------------------------------------------------
本アプリは処理のたびにGemini APIのトークンを消費します。
・画像/PDF：画像データとして送信するため、一定のトークンを消費します。
・Office文書/テキスト：事前に「文字情報」のみを抜き出して送信するため、画像として送るよりもトークン消費を抑えられます。
・WEB検索/RAG：検索結果の読み込みにより、通常よりも消費量が増えます。

コストを抑えたい場合は、gemini-1.5-flash等の軽量モデルの利用を推奨します。

--------------------------------------------------
7. 困ったときは
--------------------------------------------------
・文字化けや誤読がある：
  AIは不自然な文字を自動で日本語として再解釈するように設計されていますが、元の画像が極端に不鮮明な場合は【抽出】できないことがあります。
・WEB検索が動かない：
  【生成】指示とGUIのチェックボックスが両方有効になっているか確認してください。
・429エラーが出る：
  APIの利用制限です。プログラムは自動的に30秒待機して再開しますが、無料枠の場合は少し時間を置いてから再実行してください。

--------------------------------------------------
ライセンス / License
--------------------------------------------------
Copyright (c) 2026 Datan (データン)

本ソフトウェアは MIT ライセンスの下で公開されています。



==================================================
TextExtractor (by.gemini) Complete Manual
==================================================

A data extraction and analysis platform designed to leverage the full multimodal capabilities of the Google Gemini API to simultaneously retrieve "accurate facts" and "advanced AI insights" from any document.

--------------------------------------------------
1. Key Features and Characteristics
--------------------------------------------------
- Hybrid Extraction of Facts and Reasoning:
  Clearly separates items to be "transcribed as-is" from those the AI "thinks about based on knowledge or search."

- Unmatched File Compatibility:
  Extracts text information from images, PDFs, Office products (new and old), emails, and even "Ichitaro" (.jtd) files.

- Smart Settings Management:
  Automatically saves your API key and preferred Gemini model name after the first entry, allowing you to start analysis instantly.

- Latest Information Supplementation (Web Search/RAG):
  By simply checking a box in the GUI, the AI autonomously references Google Search or specific knowledge bases (RAG) to generate enriched information.

--------------------------------------------------
2. Supported File Formats
--------------------------------------------------
- Images/PDF: .png, .jpg, .jpeg, .webp, .bmp, .pdf
- Office Documents: .docx, .xlsx, .xls, .pptx, .doc, .ppt
- Ichitaro: .jtd
- Email: .eml, .msg (Outlook)
- Text/Data: .txt, .csv, .log

--------------------------------------------------
3. Golden Rules for Instruction Files (Prompts)
--------------------------------------------------
The "Instruction File (.txt)" that determines the AI's behavior should follow this structure.
*You can also check this via the "How to Write Help" button in the app.

**Important: Save the file using "UTF-8 (without BOM)" encoding.** Using other formats like Shift-JIS may cause character corruption, preventing the AI from processing the instructions correctly.

(1) Overview (Optional)
    Example: Analyze this business card to summarize the person's profile and company trends.

(2) Output Items (Required)
    List item names after "出力項目:" separated by commas. These will become the columns in your database.
    Example: 出力項目: Company, Name, Industry Reputation

(3) Processing Instructions (Required)
    Specify each item using one of the following:
    - 【抽出】 (Extraction): Transcribes text directly from the material (no search performed).
    - 【生成】 (Generation): AI performs reasoning or uses Web Search to supplement information.

--- Example of an Instruction File ---
出力項目: Company, Name, Title, Personality Forecast, Latest Company News

- Company, Name, Title: 【抽出】 (Extract exactly as seen. Use empty string if missing.)
- Personality Forecast: 【生成】 (AI infers based on title and design.)
- Latest Company News: 【生成】 (Use Google Search to summarize recent topics.)
-------------------------

--------------------------------------------------
4. Operating Procedures
--------------------------------------------------
1. Launch the app and enter your API key when prompted (first time only).
2. Click "Select Target Folder" and choose the folder containing the files you want to analyze.
3. Click "Select Instruction File" and choose your prepared prompt (.txt).
4. Verify the model name and turn on "Web Search" or "RAG" if needed.
5. Click "Run Instruction File" to start batch processing of all files.
6. After processing, click "Export DB to CSV" to save the accumulated data for use in Excel.

--------------------------------------------------
5. Troubleshooting
--------------------------------------------------
- Garbled Text or Misreading:
  The AI is designed to automatically re-interpret unnatural strings as meaningful text. However, extraction may fail if the original image is extremely blurry.
- Web Search is Not Working:
  Ensure both the 【生成】 instruction is present in the prompt and the GUI checkbox is enabled.
- 429 Error Occurs:
  This is an API rate limit. The program will automatically wait for 30 seconds and retry. If you are on the free tier, you may need to wait a bit longer before restarting.

--------------------------------------------------
 License
--------------------------------------------------
Copyright (c) 2026 Datan (データン)


[MIT License]
Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
