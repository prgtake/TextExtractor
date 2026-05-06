# -*- coding: utf-8 -*-
# =====================================================
#  TextExtractor1
#  Copyright (c) 2026 Datan (データン)
#  Licensed under the MIT License.
# =====================================================
import tkinter as tk
from tkinter import filedialog 
from tkinter import messagebox
import os
import re
import sqlite3
import json
import datetime
import time
import csv
import pandas as pd
import docx
from pptx import Presentation
import extract_msg
from email import policy
from email.parser import BytesParser
import tkinter.simpledialog as sd # 追加

from google import genai  # 最新元 Google Gen AI SDK
from google.genai import types  # 型定義のために追加
from google.genai.errors import APIError  # APIエラー処理用に追加

# =====================================================
#  APIキーとモデルの設定（app_tkinter.py 互換）
# =====================================================
# スクリプトまたはEXEのある場所を取得
base_dir = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(base_dir, "app_config.json")
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY")
DEFAULT_MODEL = "gemini-2.5-flash-lite"
MODEL_NAME_VAL = DEFAULT_MODEL

# 設定ファイルから読み込み
if os.path.exists(CONFIG_FILE):
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            config_data = json.load(f)
            if not GEMINI_API_KEY:
                GEMINI_API_KEY = config_data.get("GEMINI_API_KEY")
            MODEL_NAME_VAL = config_data.get("MODEL_NAME", DEFAULT_MODEL)
    except Exception:
        pass

# 2. それでもなければダイアログで入力を求める
if not GEMINI_API_KEY or GEMINI_API_KEY == "YOUR_API_KEY_HERE":
    # 一時的なルート窓を作成してダイアログを表示
    temp_root = tk.Tk()
    temp_root.withdraw()
    temp_root.attributes("-topmost", True)
    key = sd.askstring("API キー", "Gemini API キーを入力してください：\n（次回から入力不要になります）", parent=temp_root)
    temp_root.destroy()

    if not key:
        # キーが入力されなかった場合は終了
        print("APIキーが設定されていません。終了します。")
        exit()
    
    GEMINI_API_KEY = key
    # 次回のために保存
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump({"GEMINI_API_KEY": GEMINI_API_KEY, "MODEL_NAME": MODEL_NAME_VAL}, f)
    except Exception as e:
        print(f"キーの保存に失敗しました: {e}")

# クライアントの初期化
# http_options を設定して SDK レベルでの自動リトライを有効化（503, 429等に対応）
client = genai.Client(
    api_key=GEMINI_API_KEY,
    http_options={
        'retry_options': {
            'attempts': 5,
            'initial_delay': 2.0,
            'max_delay': 60.0,
            'exp_base': 2.0,
        }
    }
)

# アプリバージョン
APP_VERSION = "1.0.0"


def extract_text_from_any_file(file_path):
    """
    PDF/画像以外のファイルからテキストを抽出する補助関数
    """
    ext = os.path.splitext(file_path)[1].lower()
    content = ""
    
    try:
        # --- テキスト / CSV ---
        if ext in [".txt", ".csv", ".log"]:
            with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
                content = f.read()

        # --- Excel ---
        elif ext in [".xlsx", ".xls"]:
            dict_df = pd.read_excel(file_path, sheet_name=None)
            for sheet_name, df in dict_df.items():
                content += f"\n[Sheet: {sheet_name}]\n{df.to_csv(index=False)}"

        # --- Word ---
        elif ext == ".docx":
            doc = docx.Document(file_path)
            content = "\n".join([para.text for para in doc.paragraphs])

        # --- PowerPoint ---
        elif ext == ".pptx":
            prs = Presentation(file_path)
            for i, slide in enumerate(prs.slides):
                content += f"\n[Slide {i+1}]\n"
                for shape in slide.shapes:
                    if hasattr(shape, "text"):
                        content += shape.text + "\n"

        # --- 古いOffice形式 (.doc / .ppt) & 一太郎 (.jtd) ---
        elif ext in [".doc", ".ppt", ".jtd"]:
            import olefile
            if olefile.isOleFile(file_path):
                with olefile.OleFileIO(file_path) as ole:
                    for stream_name_list in ole.listdir():
                        # ストリーム名をパス形式に変換
                        stream_full_path = "/".join(stream_name_list)
                        
                        # 対応するストリームのキーワードをチェック
                        # JS-Main, Body は一太郎でよく使われる
                        if any(
                            key in stream_full_path
                            for key in ["WordDocument", "PowerPoint Document", "Current User", "JS-Main", "Body"]
                        ):
                            with ole.openstream(stream_name_list) as stream:
                                data = stream.read()
                                try:
                                    # 日本語(CP932)でのデコードを試みる
                                    decoded_text = data.decode('cp932', errors='ignore')
                                    # 制御文字を除去し、表示可能な文字のみを抽出
                                    content += "".join(c for c in decoded_text if c.isprintable() or c in "\n\r\t")
                                except:
                                    # デコード失敗時はバイナリから直接読める文字を拾う
                                    content += "".join(
                                        chr(b) if 32 <= b <= 126 or b >= 128 else " "
                                        for b in data
                                    )



        # --- メール (.eml) ---
        elif ext == ".eml":
            with open(file_path, 'rb') as f:
                msg = BytesParser(policy=policy.default).parse(f)
            content = f"Subject: {msg['subject']}\nFrom: {msg['from']}\nDate: {msg['date']}\n\n"
            content += msg.get_body(preferencelist=('plain')).get_content()

        # --- メール (.msg / Outlook) ---
        elif ext == ".msg":
            msg = extract_msg.Message(file_path)
            content = f"Subject: {msg.subject}\nFrom: {msg.sender}\nDate: {msg.date}\n\n{msg.body}"
            msg.close()
            
    except Exception as e:
        print(f"テキスト抽出エラー ({ext}): {e}")
        
    return content



def select_folder():
    folder_path = filedialog.askdirectory()
    if folder_path:
        folder_label.config(text=f"選択フォルダ: {folder_path}")
    else:
        folder_label.config(text="フォルダが選択されていません")

def select_prompt_file():
    prompt_file_path = filedialog.askopenfilename(filetypes=[("Text Files", "*.txt")])
    if prompt_file_path:
        prompt_label.config(text=f"選択ファイル: {prompt_file_path}")
    else:
        messagebox.showwarning("警告", "プロンプトファイルが選択されていません")

def show_prompt_help():
    help_text = (
        "【指示ファイル（プロンプト）の書き方・重要ルール】\n\n"
        "※文字コードは「UTF-8 (BOMなし/UTF-8N)」で保存してください。※\n\n"
        "以下の3つのセクションで構成してください。\n\n"
        "① 概要（任意）\n"
        "   AIに処理対象が何であるかを伝えます。\n\n"
        "② 出力項目（必須・1行で記述）\n"
        "   形式： 出力項目: 項目1, 項目2, 項目3\n"
        "   ※これがデータベースの「列名」になります。\n"
        "   ※項目は「,（半角カンマ）」または「、（全角読点）」で区切ってください。\n\n"
        "③ 処理指示（必須・項目ごとに指定）\n"
        "   各項目を【抽出】するか【生成】するかを明示してください。\n"
        "   ・【抽出】: ファイル内の情報をそのまま抜き出します。検索は行いません。\n"
        "   ・【生成】: AIが推論したり、WEB検索を使って補完したりします。\n\n"
        "【利用上の注意】\n"
        "・情報は「1ファイル内」で完結している必要があります。\n"
        "  複数ファイル（名刺の表と裏など）にまたがる情報を1つにまとめて\n"
        "  「抽出」することはできません。事前にPDF等に結合してください。\n"
        "・項目数を途中で変更した場合は、対象フォルダの.dbファイルを削除して再実行してください。"
    )
    messagebox.showinfo("TextExtractor1", help_text)

def process_files_with_prompt(folder_path, prompt_file_path):
    if not folder_path or not os.path.isdir(folder_path):
        messagebox.showerror("エラー", "有効なフォルダが選択されていません。")
        return
    if not prompt_file_path or not os.path.isfile(prompt_file_path):
        messagebox.showerror("エラー", "有効なプロンプトファイルが選択されていません。")
        return

    # ログ出力用補助関数
    def log_message(msg):
        log_area.config(state="normal")
        log_area.insert("end", msg + "\n")
        log_area.see("end")
        log_area.config(state="disabled")
        root.update()

    log_area.config(state="normal")
    log_area.delete("1.0", "end")
    log_area.config(state="disabled")
    log_message("=== 処理開始 ===")

    # 選択されたモデル名を取得し、設定ファイルに保存
    model_name = model_name_var.get().strip() or DEFAULT_MODEL
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump({"GEMINI_API_KEY": GEMINI_API_KEY, "MODEL_NAME": model_name}, f)
    except Exception:
        pass

    # 1. DB作成
    prompt_name = os.path.splitext(os.path.basename(prompt_file_path))[0]
    db_path = os.path.join(folder_path, f"{prompt_name}_DB.db")
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    # 2. プロンプト読み込み
    with open(prompt_file_path, "r", encoding="utf-8") as f:
        prompt_text = f.read()

    # 3. 出力項目の取得
    match = re.search(r"出力項目[:：]\s*(.+)", prompt_text)
    if not match:
        conn.close()
        log_message("[エラー] プロンプトファイルに『出力項目:』の記述が見つかりません。")
        return
    # 半角カンマ「,」と全角読点「、」の両方に対応
    columns = [c.strip() for c in re.split(r'[,、]', match.group(1))]
    log_message(f"検出項目 ({len(columns)}件): {', '.join(columns)}")

    # 4. テーブル作成
    table_name = "extracted_data"
    cursor.execute(f"""
        CREATE TABLE IF NOT EXISTS {table_name} (
            filename TEXT,
            {", ".join([f"'{col}' TEXT" for col in columns])}
        )
    """)

    # --- 利用可能なツールの確認（指示文に反映するため） ---
    available_tools = []
    if use_web_search.get():
        available_tools.append("Google検索")
    if use_file_search.get() and file_search_store.get():
        available_tools.append("RAG（File Search Store）")
    
    tools_info = "、".join(available_tools) if available_tools else "なし（内蔵知識のみ）"

    # --- システム共通ルール（ユーザーがプロンプトに書かなくて済むように自動付与） ---
    system_common_instruction = (
        f"\n\n### システムルール (厳守) ###\n"
        f"1. 出力は必ず以下のキーを持つ純粋なJSON形式（またはその配列）のみとしてください。\n"
        f"   キー名: {', '.join(columns)}\n"
        f"2. 各項目について、ユーザーの指示にある【抽出】と【生成】の区別を厳密に守ってください。\n"
        f"   ・【抽出】項目: 提示されたファイルの内容のみを使用してください。検索ツールは一切使用せず、見つからない場合は必ず空文字 (\"\") としてください。\n"
        f"   ・【生成】項目: 現在利用可能なツール（{tools_info}）を必要に応じて利用し、あなたの知見や推論に基づいた回答を生成してください。\n"
        f"3. 原則として、英字や記号が多い不自然な文字列はOCR失敗とみなし、日本語として意味が通じるよう再解釈してください。\n"
        f"4. 解説、挨拶、Markdownの装飾(```json等)は一切含めず、パース可能な純粋なJSONデータのみを返してください。\n"
        f"5. 複数の対象がある場合は、必ずJSON配列 [{{...}}, {{...}}] 形式で返してください。\n"
    )

    # 5. 各ファイルを処理
    files_to_process = []
    if use_recursive.get():
        for root_dir, dirs, files in os.walk(folder_path):
            for f in files:
                if not f.endswith((".db", ".txt")):
                    files_to_process.append(os.path.join(root_dir, f))
    else:
        files_to_process = [os.path.join(folder_path, f) for f in os.listdir(folder_path) 
                            if os.path.isfile(os.path.join(folder_path, f)) and not f.endswith((".db", ".txt"))]

    success_count = 0
    fail_count = 0

    for i, file_path in enumerate(files_to_process):
        filename = os.path.basename(file_path)
        
        # --- 既にDBに存在するかチェック（重複処理の回避） ---
        cursor.execute(f"SELECT COUNT(*) FROM {table_name} WHERE filename = ?", (filename,))
        if cursor.fetchone()[0] > 0:
            log_message(f"[{i+1}/{len(files_to_process)}] スキップ (処理済み): {filename}")
            success_count += 1
            continue

        log_message(f"[{i+1}/{len(files_to_process)}] 処理中: {filename}")
        
        ext = filename.lower()
        is_image = ext.endswith((".png", ".jpg", ".jpeg", ".webp", ".bmp"))
        is_pdf = ext.endswith(".pdf")

        try:
            # リトライ用のループを追加（指数バックオフ）
            max_retries = 5
            for attempt in range(max_retries):

                try:
                    # ===== Gemini 設定（WEB検索 / RAG）=====
                    config = {}
                    if use_web_search.get():
                        config["tools"] = [{"google_search": {}}]
                    if use_file_search.get() and file_search_store.get():
                        config["tools"] = config.get("tools", [])
                        config["tools"].append({
                            "file_search": {
                                "file_search_store_names": [
                                    file_search_store.get()
                                ]
                            }
                        })
                    # --- 解析パターンの切り分け ---
                    # A: PDFおよび画像（マルチモーダル解析）
                    if is_image or is_pdf:
                        with open(file_path, "rb") as f:
                            file_bytes = f.read()
                        mime_type = "application/pdf" if is_pdf else "image/jpeg"
                        
                        # 正しいデータ送信形式：types.Part.from_bytes を使用
                        response = client.models.generate_content(
                            model=model_name,
                            contents=[
                                f"{prompt_text}\n{system_common_instruction}",
                                types.Part.from_bytes(data=file_bytes, mime_type=mime_type)
                            ],
                            config=config
                        )
                    # B: それ以外（テキスト抽出解析）
                    else:
                        text_content = extract_text_from_any_file(file_path)
                        if not text_content.strip():
                            log_message(f"  -> スキップ: テキスト抽出不能")
                            fail_count += 1
                            break # attemptループを抜ける
                        response = client.models.generate_content(
                            model=model_name,
                            contents=f"{prompt_text}\n{system_common_instruction}\n\n対象内容:\n{text_content}",
                            config=config
                        )
                    
                    # 成功時のパースと保存
                    result_text = response.text.strip()
                    # JSON部分のみを抽出
                    json_match = re.search(r"(\[.*\]|\{.*\})", result_text, re.DOTALL)
                    if json_match:
                        result_text = json_match.group(1)

                    parsed = json.loads(result_text)
                    result_items = parsed if isinstance(parsed, list) else [parsed]

                    for item in result_items:
                        values = [filename] + [str(item.get(col, "")) for col in columns]
                        cursor.execute(f"INSERT INTO {table_name} VALUES ({','.join(['?']*len(values))})", values)
                    
                    conn.commit()
                    log_message(f"  -> 成功: {len(result_items)}件登録")
                    success_count += 1
                    break # 成功したらリトライループを抜ける
                    
                except Exception as e:
                    err_msg = str(e)
                    # 503 (UNAVAILABLE) や 429 (RESOURCE_EXHAUSTED) をリトライ対象とする
                    is_retryable = False
                    if isinstance(e, APIError) and (e.code in [429, 503, 500, 504]):
                        is_retryable = True
                    elif any(x in err_msg for x in ["429", "503", "500", "504", "UNAVAILABLE", "Resource has been exhausted"]):
                        is_retryable = True

                    if is_retryable and attempt < max_retries - 1:
                        # 指数バックオフ: 30秒, 60秒, 120秒, 240秒...
                        # SDK自体のリトライでも解消しない場合の最終手段として長めに待機
                        wait_time = 30 * (2 ** attempt)
                        reason = f"API制限/過負荷 ({e.code if hasattr(e, 'code') else 'Error'})"
                        log_message(f"  -> {reason}。{wait_time}秒待機して再試行します... ({attempt + 1}/{max_retries})")
                        time.sleep(wait_time)
                    else:
                        raise e 

        except Exception as e:
            log_message(f"  -> [エラー] 失敗: {e}")
            fail_count += 1
            time.sleep(4)

        # 無料枠等の制限回避のために待機
        if i < len(files_to_process) - 1:
            time.sleep(2)

    conn.close()
    log_message(f"=== 完了: 成功 {success_count} / 失敗 {fail_count} ===")
    messagebox.showinfo("完了", f"全てのファイルの処理が完了しました。\n成功: {success_count} / 失敗: {fail_count}\nデータベース: {db_path}")

def export_db_to_csv():
# 1. 現在選択されているフォルダとプロンプトのパスを取得
    folder_path = folder_label.cget("text").replace("選択フォルダ: ", "")
    prompt_file_path = prompt_label.cget("text").replace("選択ファイル: ", "")

    # 未選択状態のチェック
    if folder_path == "未選択" or prompt_file_path == "未選択":
        messagebox.showerror("エラー", "フォルダと指示ファイルを先に選択してください。")
        return

    # 2. DBパスを一意に特定（実行時と同じロジック）
    prompt_name = os.path.splitext(os.path.basename(prompt_file_path))[0]
    db_path = os.path.join(folder_path, f"{prompt_name}_DB.db")

    # 3. 存在確認
    if not os.path.exists(db_path):
        messagebox.showerror("エラー", f"対象DBが見つかりません。\nパス: {db_path}\n先に「指示ファイルを実行」してください。")
        return

    # 4. CSVの保存先だけをユーザーに選ばせる
    default_csv_name = datetime.datetime.now().strftime("%Y%m%d_%H%M%S") + ".csv"
    csv_path = filedialog.asksaveasfilename(
        defaultextension=".csv",
        initialfile=default_csv_name,
        filetypes=[("CSV Files", "*.csv")]
    )
    if not csv_path:
        return

    # 5. 出力処理
    try:
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        
        # テーブル名はプログラム固定 of "extracted_data"
        table_name = "extracted_data"
        cursor.execute(f"SELECT * FROM {table_name}")
        rows = cursor.fetchall()
        column_names = [d[0] for d in cursor.description]

        with open(csv_path, "w", encoding="utf_8_sig", newline="") as f:
            writer = csv.writer(f, delimiter=",", quoting=csv.QUOTE_ALL)
            writer.writerow(column_names)
            
            # 改行潰し処理
            cleaned_rows = []
            for row in rows:
                cleaned_rows.append([
                    (v.replace("\r\n", " ").replace("\n", " ").replace("\r", " ") 
                     if isinstance(v, str) else v) 
                    for v in row
                ])
            writer.writerows(cleaned_rows)

        conn.close()
        messagebox.showinfo("成功", f"CSVファイルを出力しました:\n{csv_path}")

    except Exception as e:
        messagebox.showerror("エラー", f"CSV出力に失敗しました:\n{e}")

# ===== GUI構成 =====
root = tk.Tk()

# ==== WEB / RAG / サブフォルダ 利用フラグ ====
use_web_search = tk.BooleanVar(value=False)
use_file_search = tk.BooleanVar(value=False)
use_recursive = tk.BooleanVar(value=False)

# RAGで使う File Search Store 名
file_search_store = tk.StringVar(value="")

# ==== モデル選択用変数 ====
model_name_var = tk.StringVar(value=MODEL_NAME_VAL)

root.title(f"TextExtractor1 v{APP_VERSION}")
root.geometry("500x600")
root.configure(bg="#f5f5f5")

# スタイル設定
btn_style = {"font": ("Meiryo", 10), "padx": 10, "pady": 3}


# 対象フォルダ選択用のフレーム
folder_select_frame = tk.Frame(root, bg="#f5f5f5")
folder_select_frame.pack(pady=10, anchor="w", padx=20)

tk.Button(folder_select_frame, text="対象フォルダを選択", command=select_folder, **btn_style).pack(side="left")
tk.Checkbutton(folder_select_frame, text="サブフォルダも含める", variable=use_recursive, bg="#f5f5f5").pack(side="left", padx=15)

folder_label = tk.Label(root, text="未選択", fg="blue", bg="#f5f5f5")
folder_label.pack(pady=3, anchor="w", padx=20)


# 指示ファイル選択用のフレーム
prompt_frame = tk.Frame(root, bg="#f5f5f5")
prompt_frame.pack(pady=3, anchor="w", padx=20)

tk.Button(prompt_frame, text="指示ファイルを選択", command=select_prompt_file, **btn_style).pack(side="left")
tk.Button(prompt_frame, text="書き方ヘルプ", command=show_prompt_help, bg="#ffffcc", **btn_style).pack(side="left", padx=10)

prompt_label = tk.Label(root, text="未選択", fg="blue", bg="#f5f5f5")
prompt_label.pack(pady=5, anchor="w", padx=20)

# Geminiモデル設定用のフレーム
model_frame_ui = tk.Frame(root, bg="#f5f5f5")
model_frame_ui.pack(pady=3, anchor="w", padx=20)
tk.Label(model_frame_ui, text="使用モデル:", font=("Meiryo", 10), bg="#f5f5f5").pack(side="left")
tk.Entry(model_frame_ui, textvariable=model_name_var, font=("Meiryo", 10), width=25, relief="solid").pack(side="left", padx=10)
tk.Label(model_frame_ui, text="※例: gemini-2.0-flash", font=("Meiryo", 8), bg="#f5f5f5", fg="gray").pack(side="left")

# オプション用フレーム
option_frame = tk.Frame(root, bg="#f5f5f5")
option_frame.pack(anchor="w", padx=20)

tk.Checkbutton(option_frame, text="WEB検索を利用", variable=use_web_search, bg="#f5f5f5").pack(side="left")
tk.Checkbutton(option_frame, text="RAGを利用", variable=use_file_search, bg="#f5f5f5").pack(side="left", padx=10)

tk.Label(root, text="使うRAGを選択", bg="#f5f5f5").pack(anchor="w", padx=20)

rag_dropdown = tk.OptionMenu(root, file_search_store, "")
rag_dropdown.pack(anchor="w", padx=20)

stores = client.file_search_stores.list()

menu = rag_dropdown["menu"]
menu.delete(0, "end")

for store in stores:
    store_id = store.name
    menu.add_command(
        label=store.display_name or store_id,
        command=lambda x=store_id: file_search_store.set(x)
    )

# ボタンを横に並べるための「箱（フレーム）」を作ります
button_frame = tk.Frame(root, bg="#f5f5f5")
button_frame.pack(anchor="w", padx=20, pady=5)



tk.Button(button_frame, text="指示ファイルを実行", bg="#ff9999", **btn_style, 
          command=lambda: process_files_with_prompt(
              folder_label.cget("text").replace("選択フォルダ: ", ""),
              prompt_label.cget("text").replace("選択ファイル: ", ""))
          ).pack(side="left")

tk.Button(button_frame, text="　 DBをCSV出力 　", bg="#99ff99", **btn_style,
          command=export_db_to_csv  # 引数なしで呼び出し
          ).pack(side="left", padx=20)

# ログ表示エリアの追加
log_frame = tk.Frame(root, bg="#f5f5f5")
log_frame.pack(pady=10, fill="both", expand=True, padx=20)
tk.Label(log_frame, text="実行ログ:", bg="#f5f5f5").pack(anchor="w")
log_area = tk.Text(log_frame, height=8, font=("Consolas", 9), state="disabled", bg="white")
log_area.pack(side="left", fill="both", expand=True)
scrollbar = tk.Scrollbar(log_frame, command=log_area.yview)
scrollbar.pack(side="right", fill="y")
log_area.config(yscrollcommand=scrollbar.set)

root.mainloop()
