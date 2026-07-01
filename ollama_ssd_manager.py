# -*- coding: utf-8 -*-
"""
Created on Fri Apr 24 16:56:09 2026

@author: Administrator
"""

# -*- coding: utf-8 -*-
"""
Ollama SSD Manager 完全版
E:\Ollama 専用
診断・爆速化・安全管理・AI分析 GUI
"""

import os
import json
import time
import queue
import shutil
import threading
import subprocess
import urllib.request
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog


# =========================
# 基本設定
# =========================

OLLAMA_EXE = r"E:\Ollama\ollama.exe"
OLLAMA_MODELS_DIR = r"E:\Ollama\models"
OLLAMA_HOST = "http://localhost:11434"

DEFAULT_MODELS = [
    "gemma3:4b",
    "qwen3.5:4b",
    "qwen3.5:9b",
    "qwen3-vl:4b",
    "qwen3-vl:8b",
    "deepseek-r1:8b",
]

msg_queue = queue.Queue()


# =========================
# 共通関数
# =========================

def make_env():
    env = os.environ.copy()
    env["OLLAMA_MODELS"] = OLLAMA_MODELS_DIR
    env["OLLAMA_KEEP_ALIVE"] = "1h"
    env["OLLAMA_MAX_LOADED_MODELS"] = "1"
    env["OLLAMA_NUM_PARALLEL"] = "1"
    env["OLLAMA_FLASH_ATTENTION"] = "1"
    return env


def put(text):
    msg_queue.put(("text", text))


def is_ollama_alive():
    try:
        with urllib.request.urlopen(OLLAMA_HOST + "/api/tags", timeout=3) as res:
            return res.status == 200
    except Exception:
        return False


def run_powershell(cmd):
    try:
        r = subprocess.run(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", cmd],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace"
        )
        return r.stdout + r.stderr
    except Exception as e:
        return str(e)


def format_size(size):
    units = ["B", "KB", "MB", "GB", "TB"]
    n = float(size)
    for u in units:
        if n < 1024:
            return f"{n:.2f} {u}"
        n /= 1024
    return f"{n:.2f} PB"


# =========================
# Ollama操作
# =========================

def start_ollama():
    if is_ollama_alive():
        messagebox.showinfo("確認", "Ollamaはすでに起動しています。")
        return

    try:
        subprocess.Popen(
            [OLLAMA_EXE, "serve"],
            env=make_env(),
            creationflags=subprocess.CREATE_NEW_CONSOLE
        )
        time.sleep(2)

        run_powershell("(Get-Process ollama -ErrorAction SilentlyContinue).PriorityClass='High'")

        if is_ollama_alive():
            status_var.set("Ollama起動中")
            put("\n[OK] Ollamaを起動しました。\n")
        else:
            status_var.set("起動確認中")
            put("\n[注意] Ollama起動確認に時間がかかっています。\n")

    except Exception as e:
        messagebox.showerror("起動エラー", str(e))


def stop_ollama():
    run_powershell("Stop-Process -Name ollama -Force -ErrorAction SilentlyContinue")
    status_var.set("Ollama停止")
    put("\n[OK] Ollamaを停止しました。\n")


def refresh_models():
    def worker():
        try:
            with urllib.request.urlopen(OLLAMA_HOST + "/api/tags", timeout=10) as res:
                data = json.loads(res.read().decode("utf-8"))

            models = [m.get("name", "") for m in data.get("models", []) if m.get("name")]

            if models:
                msg_queue.put(("models", models))
                put("\n[OK] モデル一覧を更新しました。\n")
                for m in models:
                    put(f"  - {m}\n")
            else:
                put("\n[注意] モデルが見つかりません。\n")

        except Exception as e:
            put(f"\n[ERROR] モデル一覧取得失敗: {e}\n")

    threading.Thread(target=worker, daemon=True).start()


# =========================
# 診断系
# =========================

def diagnose_ssd():
    def worker():
        put("\n========== SSD診断 ==========\n")

        try:
            total, used, free = shutil.disk_usage("E:\\")
            put(f"Eドライブ総容量: {format_size(total)}\n")
            put(f"Eドライブ使用量: {format_size(used)}\n")
            put(f"Eドライブ空き容量: {format_size(free)}\n")
            put(f"空き率: {free / total * 100:.1f}%\n")
        except Exception as e:
            put(f"容量取得エラー: {e}\n")

        put("\n--- Volume情報 ---\n")
        put(run_powershell("Get-Volume E | Format-List *"))

        put("\n--- PhysicalDisk情報 ---\n")
        put(run_powershell("Get-PhysicalDisk | Format-Table Number,FriendlyName,MediaType,OperationalStatus,HealthStatus,Size"))

        put("\n--- TRIM設定 ---\n")
        put(run_powershell("fsutil behavior query DisableDeleteNotify"))

        put("\n[診断メモ]\n")
        put("- exFATの場合、Optimize-Volume -ReTrim は失敗して正常です。\n")
        put("- スマホLLMと共用するSSDなら、NTFS化は慎重にしてください。\n")
        put("- 空き容量は最低でも20%以上あると安全です。\n")

        put("\n========== SSD診断完了 ==========\n")

    threading.Thread(target=worker, daemon=True).start()


def diagnose_ollama():
    def worker():
        put("\n========== Ollama診断 ==========\n")

        put(f"Ollama実行ファイル: {OLLAMA_EXE}\n")
        put(f"存在確認: {os.path.exists(OLLAMA_EXE)}\n")
        put(f"モデル保存先: {OLLAMA_MODELS_DIR}\n")
        put(f"存在確認: {os.path.exists(OLLAMA_MODELS_DIR)}\n")
        put(f"Ollama API応答: {is_ollama_alive()}\n")

        put("\n--- Ollamaプロセス ---\n")
        put(run_powershell("Get-Process ollama -ErrorAction SilentlyContinue | Select-Object Name,Id,CPU,WorkingSet64,PriorityClass | Format-Table"))

        put("\n--- Ollamaモデル一覧 ---\n")
        try:
            r = subprocess.run(
                [OLLAMA_EXE, "list"],
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                env=make_env()
            )
            put(r.stdout + r.stderr)
        except Exception as e:
            put(str(e))

        put("\n--- 環境変数案 ---\n")
        for k in ["OLLAMA_MODELS", "OLLAMA_KEEP_ALIVE", "OLLAMA_MAX_LOADED_MODELS", "OLLAMA_NUM_PARALLEL", "OLLAMA_FLASH_ATTENTION"]:
            put(f"{k} = {make_env().get(k)}\n")

        put("\n========== Ollama診断完了 ==========\n")

    threading.Thread(target=worker, daemon=True).start()


def apply_fast_settings():
    """
    安全な環境変数のみ設定。
    PATHは触らない。
    """
    cmds = [
        'setx OLLAMA_MODELS "E:\\Ollama\\models"',
        'setx OLLAMA_KEEP_ALIVE "1h"',
        'setx OLLAMA_MAX_LOADED_MODELS "1"',
        'setx OLLAMA_NUM_PARALLEL "1"',
        'setx OLLAMA_FLASH_ATTENTION "1"',
    ]

    put("\n========== 爆速設定適用 ==========\n")

    for c in cmds:
        put(f"> {c}\n")
        put(run_powershell(c))

    put("\n[OK] 爆速設定を適用しました。\n")
    put("PowerShell・Spyder・GUIを一度閉じて再起動すると確実に反映されます。\n")
    put("※ PATHは変更していません。\n")
    put("\n========== 完了 ==========\n")


def search_large_files():
    folder = filedialog.askdirectory(title="巨大ファイルを検索するフォルダを選択", initialdir="E:\\")
    if not folder:
        return

    def worker():
        put(f"\n========== 巨大ファイル検索: {folder} ==========\n")
        files = []

        for root_dir, dirs, filenames in os.walk(folder):
            for name in filenames:
                path = os.path.join(root_dir, name)
                try:
                    size = os.path.getsize(path)
                    if size >= 500 * 1024 * 1024:
                        files.append((size, path))
                except Exception:
                    pass

        files.sort(reverse=True)

        if not files:
            put("500MB以上のファイルは見つかりませんでした。\n")
        else:
            put("500MB以上のファイル一覧：\n")
            for size, path in files[:100]:
                put(f"{format_size(size)}  {path}\n")

        put("\n[安全注意]\n")
        put("このアプリは削除しません。削除判断は必ず手動で行ってください。\n")
        put("\n========== 検索完了 ==========\n")

    threading.Thread(target=worker, daemon=True).start()


# =========================
# AIチャット / 診断結果分析
# =========================

def ask_model(prompt_override=None):
    model = model_var.get().strip()
    prompt = prompt_override if prompt_override is not None else input_box.get("1.0", tk.END).strip()

    if not prompt:
        messagebox.showwarning("入力なし", "質問を入力してください。")
        return

    if not is_ollama_alive():
        start_ollama()
        time.sleep(2)

    output_box.insert(tk.END, f"\n\n【あなた】\n{prompt[:1000]}\n\n【{model}】\n")
    output_box.see(tk.END)

    send_button.config(state="disabled")
    status_var.set(f"{model} 応答中...")

    def worker():
        try:
            data = {
                "model": model,
                "prompt": prompt,
                "stream": True,
                "options": {
                    "num_predict": int(num_predict_var.get()),
                    "temperature": float(temp_var.get()),
                    "num_ctx": int(ctx_var.get())
                }
            }

            req = urllib.request.Request(
                OLLAMA_HOST + "/api/generate",
                data=json.dumps(data).encode("utf-8"),
                headers={"Content-Type": "application/json"}
            )

            with urllib.request.urlopen(req, timeout=900) as res:
                got = False
                for line in res:
                    if not line:
                        continue
                    try:
                        obj = json.loads(line.decode("utf-8"))
                    except Exception:
                        continue

                    if obj.get("response"):
                        got = True
                        msg_queue.put(("text", obj["response"]))

                    if obj.get("error"):
                        msg_queue.put(("text", "\n[ERROR] " + obj["error"] + "\n"))

                    if obj.get("done"):
                        if not got:
                            msg_queue.put(("text", "\n[注意] 応答本文が空でした。別モデルを試してください。\n"))
                        msg_queue.put(("text", "\n\n--- 完了 ---\n"))
                        break

        except Exception as e:
            msg_queue.put(("text", f"\n\n[ERROR] {e}\n"))

        finally:
            msg_queue.put(("done", None))

    threading.Thread(target=worker, daemon=True).start()


def analyze_with_ai():
    text = output_box.get("1.0", tk.END).strip()

    if not text:
        messagebox.showwarning("診断結果なし", "先にSSD診断またはOllama診断を実行してください。")
        return

    prompt = f"""
あなたはWindowsローカルLLM環境の診断アシスタントです。
以下の診断ログを読み、SSD上のOllama環境について、
安全性、速度、安定性、改善案を日本語で整理してください。

制約：
- ファイル削除を指示しない
- NTFS化を強制しない
- スマホ用LLMが同じSSDに入っている前提で壊さない
- 初心者でも実行できるPowerShellコマンドだけ提案
- 危険な操作には必ず注意書きを付ける

診断ログ：
{text[-12000:]}
"""
    ask_model(prompt)


def clear_output():
    output_box.delete("1.0", tk.END)


def copy_output():
    text = output_box.get("1.0", tk.END).strip()
    if text:
        root.clipboard_clear()
        root.clipboard_append(text)
        messagebox.showinfo("コピー", "ログ/回答をコピーしました。")


# =========================
# GUI更新
# =========================

def update_output():
    try:
        while True:
            kind, value = msg_queue.get_nowait()

            if kind == "text":
                output_box.insert(tk.END, value)
                output_box.see(tk.END)

            elif kind == "done":
                send_button.config(state="normal")
                status_var.set("待機中")

            elif kind == "models":
                model_combo["values"] = value
                if value:
                    model_var.set(value[0])

    except queue.Empty:
        pass

    root.after(100, update_output)


# =========================
# GUI本体
# =========================

root = tk.Tk()
root.title("Ollama SSD Manager 完全版 - Eドライブ専用")
root.geometry("1050x780")

status_var = tk.StringVar(value="待機中")

main = ttk.Frame(root, padding=10)
main.pack(fill="both", expand=True)

title = ttk.Label(main, text="Ollama SSD Manager 完全版", font=("Meiryo", 17, "bold"))
title.pack(anchor="w")

subtitle = ttk.Label(
    main,
    text="E:\\Ollama 専用 / 診断・爆速化・安全管理 / 削除なし",
    font=("Meiryo", 10)
)
subtitle.pack(anchor="w", pady=(0, 8))

# 操作ボタン
btn_frame = ttk.LabelFrame(main, text="操作")
btn_frame.pack(fill="x", pady=5)

ttk.Button(btn_frame, text="Ollama起動", command=start_ollama).pack(side="left", padx=4, pady=6)
ttk.Button(btn_frame, text="Ollama停止", command=stop_ollama).pack(side="left", padx=4)
ttk.Button(btn_frame, text="モデル一覧更新", command=refresh_models).pack(side="left", padx=4)
ttk.Button(btn_frame, text="SSD診断", command=diagnose_ssd).pack(side="left", padx=4)
ttk.Button(btn_frame, text="Ollama診断", command=diagnose_ollama).pack(side="left", padx=4)
ttk.Button(btn_frame, text="爆速設定適用", command=apply_fast_settings).pack(side="left", padx=4)
ttk.Button(btn_frame, text="巨大ファイル検索", command=search_large_files).pack(side="left", padx=4)

ttk.Label(btn_frame, textvariable=status_var).pack(side="right", padx=10)

# モデル設定
model_frame = ttk.LabelFrame(main, text="AI分析モデル")
model_frame.pack(fill="x", pady=5)

model_var = tk.StringVar(value="gemma3:4b")
model_combo = ttk.Combobox(
    model_frame,
    textvariable=model_var,
    values=DEFAULT_MODELS,
    state="readonly",
    width=30
)
model_combo.pack(side="left", padx=8, pady=8)

ttk.Label(model_frame, text="出力Token").pack(side="left")
num_predict_var = tk.StringVar(value="768")
ttk.Entry(model_frame, textvariable=num_predict_var, width=8).pack(side="left", padx=4)

ttk.Label(model_frame, text="温度").pack(side="left")
temp_var = tk.StringVar(value="0.4")
ttk.Entry(model_frame, textvariable=temp_var, width=8).pack(side="left", padx=4)

ttk.Label(model_frame, text="文脈").pack(side="left")
ctx_var = tk.StringVar(value="4096")
ttk.Entry(model_frame, textvariable=ctx_var, width=8).pack(side="left", padx=4)

# 入力欄
ttk.Label(main, text="AIへの質問 / 指示").pack(anchor="w", pady=(8, 0))

input_box = scrolledtext.ScrolledText(main, height=5, font=("Meiryo", 10))
input_box.pack(fill="x", pady=4)
input_box.insert(tk.END, "このOllama環境を安全に高速化する改善案を出して。")

send_frame = ttk.Frame(main)
send_frame.pack(fill="x", pady=4)

send_button = ttk.Button(send_frame, text="AIに質問", command=ask_model)
send_button.pack(side="left", fill="x", expand=True, padx=3)

ttk.Button(send_frame, text="診断結果をAI分析", command=analyze_with_ai).pack(side="left", padx=3)
ttk.Button(send_frame, text="ログ/回答クリア", command=clear_output).pack(side="left", padx=3)
ttk.Button(send_frame, text="ログ/回答コピー", command=copy_output).pack(side="left", padx=3)

# 出力欄
ttk.Label(main, text="診断ログ / AI回答").pack(anchor="w", pady=(8, 0))

output_box = scrolledtext.ScrolledText(main, height=28, font=("Meiryo", 10))
output_box.pack(fill="both", expand=True, pady=5)

output_box.insert(tk.END, "Ollama SSD Managerを起動しました。\n")
output_box.insert(tk.END, "推奨手順:\n")
output_box.insert(tk.END, "1. Ollama起動\n")
output_box.insert(tk.END, "2. モデル一覧更新\n")
output_box.insert(tk.END, "3. SSD診断\n")
output_box.insert(tk.END, "4. Ollama診断\n")
output_box.insert(tk.END, "5. 診断結果をAI分析\n\n")
output_box.insert(tk.END, "安全方針: 削除なし / PATH変更なし / exFAT維持 / スマホLLM領域を保護\n\n")

update_output()
root.mainloop()