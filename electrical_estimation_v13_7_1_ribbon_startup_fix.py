#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
電気設備積算アプリ 統合軽量版 + 画像PDF/Ollama学習
=====================================================

元アプリから残した機能:
- PDFテキスト抽出積算
- DXF解析（ezdxf優先 / ASCII簡易フォールバック）
- 単価マスター編集・CSV入出力
- 記号パターン編集
- CADライブラリ登録・URL/ZIP取込
- Ollama AIチャット / モデル一覧取得

追加機能:
- CAD情報なし画像PDF/図面画像の軽量解析
- 図面プレビューで図記号クリック/ダブルクリック
- Ollamaで図記号名を推定
- OK登録 / 訂正登録・学習 / 特徴量保存
- 学習データSQLite保存 / CSV出力 / 再学習・再推論エンジン
- LiteLLM: Ollama/OpenAI/Claude/OpenAI互換APIの統一呼び出し
- LangMem互換長期記憶: OK/NG修正・積算レビューを次回プロンプトへ反映
- フリーズ対策: 縮小画像・間引き走査・別スレッド・GUI更新はメインスレッドのみ

必要:
    py -m pip install pillow pymupdf pypdf ezdxf requests chardet

起動:
    py electrical_estimation_integrated_light_ollama.py
"""

import os, sys, re, csv, json, math, time, base64, queue, zipfile, shutil, sqlite3, threading, subprocess, traceback
from pathlib import Path
from dataclasses import dataclass
from collections import defaultdict
from datetime import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter import scrolledtext

# ---------------- Pydroid / Android safe helpers ----------------
def is_android_pydroid():
    return (
        "ANDROID_ROOT" in os.environ
        or "ANDROID_DATA" in os.environ
        or "PYDROID" in sys.executable.lower()
        or "/data/user/" in sys.executable
    )

def get_safe_base_dir():
    """
    Windowsでは E:\\electrical_estimation_ai を優先。
    Android/Pydroid3では /storage/emulated/0/electrical_estimation_ai を使う。
    """
    if is_android_pydroid():
        candidates = [
            Path("/storage/emulated/0/electrical_estimation_ai"),
            Path.home() / "electrical_estimation_ai",
            Path.cwd() / "electrical_estimation_ai",
        ]
        for p in candidates:
            try:
                p.mkdir(parents=True, exist_ok=True)
                test = p / ".write_test"
                test.write_text("ok", encoding="utf-8")
                test.unlink(missing_ok=True)
                return p
            except Exception:
                continue
        return Path.cwd() / "electrical_estimation_ai"
    p = Path(r"E:\electrical_estimation_ai")
    if not p.exists():
        p = Path.cwd() / "electrical_estimation_ai"
    return p

def write_crash_log(exc_text):
    try:
        base = get_safe_base_dir()
        base.mkdir(parents=True, exist_ok=True)
        log_path = base / "pydroid_crash_log.txt"
        log_path.write_text(exc_text, encoding="utf-8")
        print("クラッシュログ:", log_path)
    except Exception:
        print(exc_text)


try:
    from PIL import Image, ImageTk, ImageDraw
except Exception:
    Image = ImageTk = ImageDraw = None

def try_import_fitz():
    try:
        import fitz
        return fitz
    except Exception:
        return None

def try_import_pypdf():
    try:
        from pypdf import PdfReader
        return PdfReader
    except Exception:
        return None

def try_import_requests():
    try:
        import requests
        return requests
    except Exception:
        return None

def try_import_ezdxf():
    try:
        import ezdxf
        return ezdxf
    except Exception:
        return None

def try_import_chardet():
    try:
        import chardet
        return chardet
    except Exception:
        return None


def try_import_cv2():
    try:
        import cv2
        return cv2
    except Exception:
        return None

def try_import_numpy():
    try:
        import numpy as np
        return np
    except Exception:
        return None

def try_import_pytesseract():
    try:
        import pytesseract
        return pytesseract
    except Exception:
        return None

def try_import_ultralytics():
    try:
        from ultralytics import YOLO
        return YOLO
    except Exception:
        return None


def try_import_litellm():
    try:
        import litellm
        return litellm
    except Exception:
        return None

def try_import_langmem():
    try:
        import langmem
        return langmem
    except Exception:
        return None


def try_import_faiss():
    try:
        import faiss
        return faiss
    except Exception:
        return None

def try_import_numpy():
    try:
        import numpy as np
        return np
    except Exception:
        return None

def try_import_torch():
    try:
        import torch
        return torch
    except Exception:
        return None

def try_import_clip_stack():
    """
    transformers + torch によるCLIP画像埋め込み。
    無ければ None を返し、軽量特徴量埋め込みへフォールバック。
    """
    try:
        import torch
        from transformers import CLIPProcessor, CLIPModel
        return torch, CLIPProcessor, CLIPModel
    except Exception:
        return None

APP_TITLE = "ELECTimate AI v13.7.1 リボンUI起動修正版"
BASE_DIR = get_safe_base_dir()
BASE_DIR.mkdir(parents=True, exist_ok=True)
DB_PATH = str(BASE_DIR / "estimation_master_integrated.db")
CAD_LIBRARY_PATH = str(BASE_DIR / "cad_library")
OUTPUT_DIR = BASE_DIR / "outputs"
CROP_DIR = OUTPUT_DIR / "symbol_crops"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
CROP_DIR.mkdir(parents=True, exist_ok=True)

VECTOR_INDEX_DIR = BASE_DIR / "vector_index"
VECTOR_INDEX_DIR.mkdir(parents=True, exist_ok=True)
FAISS_INDEX_PATH = VECTOR_INDEX_DIR / "symbols.faiss"
VECTOR_META_PATH = VECTOR_INDEX_DIR / "symbols_meta.jsonl"
CLIP_MODEL_NAME_DEFAULT = "openai/clip-vit-base-patch32"
VECTOR_DIM_FALLBACK = 128

os.makedirs(CAD_LIBRARY_PATH, exist_ok=True)
DEFAULT_OLLAMA_URL = "http://127.0.0.1:11434"
DEFAULT_OLLAMA_MODEL = "deepseek-r1:8b"
DEFAULT_OPENAI_MODEL = "gpt-4o-mini"
DEFAULT_CLAUDE_MODEL = "claude-sonnet-4-5"
CLAUDE_MODEL_CHOICES = ["claude-sonnet-4-5","claude-haiku-4-5","claude-opus-4-1","claude-3-5-haiku-latest","claude-3-5-sonnet-latest"]
DEFAULT_CUSTOM_OPENAI_BASE_URL = "http://127.0.0.1:8000/v1"
DQN_ACTIONS = ["color_fast","color_strict","opencv_shape","template_match","ocr_assist","learned_first","llm_assist","manual_annotation"]
DQN_ACTION_DESCRIPTIONS = {
    "color_fast":"色抽出を広め設定で高速解析",
    "color_strict":"色抽出を厳しめ設定で巨大誤検出を抑制",
    "opencv_shape":"OpenCV形状特徴を優先",
    "template_match":"テンプレート/記号パターンを優先",
    "ocr_assist":"OCR文字情報を補助利用",
    "learned_first":"学習DB・記号データセットを最優先",
    "llm_assist":"LLM判定を補助利用",
    "manual_annotation":"自動判定を抑え、人間アノテーションを促す",
}
DQN_MODEL_PATH = str(BASE_DIR / "dqn_strategy_model.json")

# 2026-05時点の概算用。正確な最新価格は各社の公式価格ページを確認してください。
OPENAI_PRICE_USD_PER_MTOK = {
    "gpt-4o-mini": (0.15, 0.60),
    "gpt-4o": (5.00, 15.00),
    "gpt-4.1-mini": (0.40, 1.60),
    "gpt-4.1": (2.00, 8.00),
    "gpt-5-mini": (0.25, 2.00),
    "gpt-5": (1.25, 10.00),
}
ANTHROPIC_PRICE_USD_PER_MTOK = {
    "claude-sonnet-4-5": (3.00, 15.00),
    "claude-sonnet-4-6": (3.00, 15.00),
    "claude-haiku-4-5": (1.00, 5.00),
    "claude-opus-4-1": (15.00, 75.00),
    "claude-3-5-haiku-latest": (0.80, 4.00),
    "claude-3-5-sonnet-latest": (3.00, 15.00),
}



def diagnose_pydroid_environment():
    """
    Pydroid3で落ちる原因を起動時に診断する。
    GUI起動できない場合でも標準出力とログへ出す。
    """
    results = []
    results.append(f"Python: {sys.version}")
    results.append(f"Executable: {sys.executable}")
    results.append(f"Android/Pydroid: {is_android_pydroid()}")
    results.append(f"BASE_DIR: {BASE_DIR}")
    checks = [
        ("tkinter", lambda: __import__("tkinter")),
        ("Pillow", lambda: __import__("PIL")),
        ("requests", lambda: __import__("requests")),
        ("sqlite3", lambda: __import__("sqlite3")),
        ("PyMuPDF(fitz)", lambda: __import__("fitz")),
        ("pypdf", lambda: __import__("pypdf")),
        ("ezdxf", lambda: __import__("ezdxf")),
        ("cv2", lambda: __import__("cv2")),
        ("numpy", lambda: __import__("numpy")),
        ("FAISS", lambda: __import__("faiss")),
        ("transformers", lambda: __import__("transformers")),
        ("torch", lambda: __import__("torch")),
    ]
    for name, fn in checks:
        try:
            fn()
            results.append(f"[OK] {name}")
        except Exception as e:
            results.append(f"[NG] {name}: {type(e).__name__}: {e}")
    msg = "\n".join(results)
    try:
        (BASE_DIR / "pydroid_startup_check.txt").write_text(msg, encoding="utf-8")
    except Exception:
        pass
    print(msg)
    return msg


# ---------------- common ----------------
def add_paste_button(parent, widget, width=7):
    def paste():
        try: text = parent.clipboard_get()
        except Exception: text = ""
        if not text: return
        try:
            widget.delete(0, tk.END); widget.insert(0, text)
        except Exception:
            try: widget.delete("1.0", tk.END); widget.insert("1.0", text)
            except Exception: pass
    ttk.Button(parent, text="貼付", command=paste, width=width).pack(side=tk.LEFT, padx=2)

def run_pip_install(package, timeout_sec=220):
    cmd = [sys.executable, "-m", "pip", "install", package]
    proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, bufsize=1)
    lines, start = [], time.time()
    try:
        while True:
            if time.time() - start > timeout_sec:
                proc.kill(); return False, lines + [f"[TIMEOUT] {timeout_sec}s"]
            line = proc.stdout.readline() if proc.stdout else ""
            if line: lines.append(line.rstrip())
            elif proc.poll() is not None: break
            else: time.sleep(0.05)
        return proc.wait() == 0, lines
    except Exception as e:
        try: proc.kill()
        except Exception: pass
        return False, lines + [f"[EXCEPTION] {e}"]

def detect_encoding(path):
    ch = try_import_chardet()
    if ch:
        try:
            raw = Path(path).read_bytes()[:60000]
            r = ch.detect(raw)
            enc, conf = r.get("encoding"), r.get("confidence", 0)
            if enc and conf > 0.75:
                return "cp932" if "shift" in enc.lower() else enc
        except Exception: pass
    return "cp932"

def read_text_file_safe(path):
    encs = [detect_encoding(path), "cp932", "shift-jis", "utf-8", "utf-8-sig", "euc-jp", "latin1"]
    seen = []
    for enc in encs:
        if enc in seen: continue
        seen.append(enc)
        try:
            s = Path(path).read_text(encoding=enc, errors="ignore")
            if s.strip(): return s, enc
        except Exception: pass
    return "", "unknown"


# ---------------- rqlite adapter ----------------
class RqliteCursor:
    def __init__(self, conn):
        self.conn = conn
        self._last_rows = []

    def execute(self, sql, params=()):
        self._last_rows = self.conn.execute(sql, params)
        return self

    def executemany(self, sql, seq_of_params):
        for params in seq_of_params:
            self.conn.execute(sql, params)
        self._last_rows = []
        return self

    def fetchone(self):
        if not self._last_rows:
            return None
        return self._last_rows.pop(0)

    def fetchall(self):
        rows = self._last_rows
        self._last_rows = []
        return rows


class RqliteConnection:
    """
    sqlite3.Connection風の最小rqliteアダプタ。
    rqliteのHTTP API /status, /db/execute, /db/query を使う。

    注意:
    rqliteはPythonライブラリではなく、別プロセスで起動するDBサーバです。
    http://127.0.0.1:4001 に接続するには、先に rqlited.exe を起動しておく必要があります。
    """
    def __init__(self, base_url, timeout=12):
        self.base_url = self.normalize_url(base_url)
        self.timeout = timeout
        self.row_factory = None

    @staticmethod
    def normalize_url(url):
        url = (url or 'http://127.0.0.1:4001').strip()
        if not url.startswith(('http://','https://')):
            url = 'http://' + url
        return url.rstrip('/')

    def cursor(self):
        return RqliteCursor(self)

    def _is_select(self, sql):
        s = sql.strip().lower()
        return s.startswith('select') or s.startswith('pragma') or s.startswith('with')

    def _requests(self):
        requests = try_import_requests()
        if not requests:
            raise RuntimeError('rqlite接続には requests が必要です')
        return requests

    def _human_connection_error(self, e):
        return (
            f'rqliteサーバへ接続できません: {self.base_url}\n\n'
            '原因はほぼ次のどれかです。\n'
            '1. rqlited.exe がまだ起動していない\n'
            '2. ポート4001ではなく別ポートで起動している\n'
            '3. URL欄が 127.0.0.1:4001 だけで、http:// が無い\n'
            '4. ファイアウォール/セキュリティソフトがlocalhost接続を拒否している\n\n'
            '起動例 PowerShell:\n'
            '  E:\\rqlite\\rqlited.exe E:\\rqlite\\node1\n\n'
            'または rqlited.exe のあるフォルダで:\n'
            '  .\\rqlited.exe E:\\rqlite\\node1\n\n'
            f'詳細: {e}'
        )

    def status(self):
        requests = self._requests()
        try:
            r = requests.get(self.base_url + '/status', timeout=self.timeout)
            r.raise_for_status()
            return r.json()
        except Exception as e:
            raise RuntimeError(self._human_connection_error(e))

    def _post(self, endpoint, payload):
        requests = self._requests()
        try:
            r = requests.post(self.base_url + endpoint, json=payload, timeout=self.timeout)
            r.raise_for_status()
            data = r.json()
        except Exception as e:
            raise RuntimeError(self._human_connection_error(e))
        if isinstance(data, dict) and data.get('error'):
            raise RuntimeError(data.get('error'))
        return data

    def execute(self, sql, params=()):
        params = list(params or [])
        if self._is_select(sql):
            data = self._post('/db/query?level=strong', [[sql] + params])
            results = data.get('results', []) if isinstance(data, dict) else []
            if not results:
                return []
            res = results[0]
            if res.get('error'):
                raise RuntimeError(res.get('error'))
            cols = res.get('columns', [])
            vals = res.get('values', []) or []
            if self.row_factory is sqlite3.Row:
                return [dict(zip(cols, row)) for row in vals]
            return [tuple(row) for row in vals]
        data = self._post('/db/execute', [[sql] + params])
        if isinstance(data, dict):
            for res in data.get('results',[]) or []:
                if res.get('error'):
                    raise RuntimeError(res.get('error'))
        return []

    def executemany(self, sql, seq_of_params):
        stmts = [[sql] + list(params or []) for params in seq_of_params]
        if stmts:
            self._post('/db/execute', stmts)
        return []

    def commit(self):
        pass

    def close(self):
        pass

class DBBackendConfig:
    @staticmethod
    def backend():
        return os.environ.get('ESTIMATION_DB_BACKEND', 'sqlite').strip().lower()

    @staticmethod
    def rqlite_url():
        return os.environ.get('RQLITE_URL', 'http://127.0.0.1:4001').strip()


# ---------------- rqlite migration helpers ----------------
def normalize_rqlite_url(url):
    url=(url or '').strip()
    if not url:
        return 'http://127.0.0.1:4001'
    if not url.startswith(('http://','https://')):
        url='http://' + url
    return url.rstrip('/')

def sqlite_value_to_rqlite(v):
    if isinstance(v,(int,float)) or v is None:
        return v
    if isinstance(v, bytes):
        try:
            return v.decode('utf-8','ignore')
        except Exception:
            return str(v)
    return str(v)

# ---------------- DB ----------------
class DatabaseManager:
    def __init__(self, db_path):
        self.db_path = db_path
        self.init_db()
    def connect(self):
        if DBBackendConfig.backend() == 'rqlite':
            return RqliteConnection(DBBackendConfig.rqlite_url())
        return sqlite3.connect(self.db_path)
    def init_db(self):
        con = self.connect(); c = con.cursor()
        c.execute("""CREATE TABLE IF NOT EXISTS unit_prices(id INTEGER PRIMARY KEY AUTOINCREMENT, category TEXT, item_name TEXT, spec TEXT, unit TEXT, unit_price REAL, keywords TEXT)""")
        c.execute("""CREATE TABLE IF NOT EXISTS symbol_patterns(id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT, pattern_type TEXT, pattern_json TEXT, description TEXT, preset INTEGER DEFAULT 0, added_at TEXT DEFAULT CURRENT_TIMESTAMP)""")
        c.execute("""CREATE TABLE IF NOT EXISTS cad_library(id INTEGER PRIMARY KEY AUTOINCREMENT, manufacturer TEXT, model_number TEXT, category TEXT, item_name TEXT, file_path TEXT, pattern_signature TEXT, spec_json TEXT, url TEXT, added_at TEXT DEFAULT CURRENT_TIMESTAMP)""")
        c.execute("""CREATE TABLE IF NOT EXISTS image_color_rules(id INTEGER PRIMARY KEY AUTOINCREMENT, color_name TEXT UNIQUE, equipment TEXT, r_min INTEGER, r_max INTEGER, g_min INTEGER, g_max INTEGER, b_min INTEGER, b_max INTEGER, enabled INTEGER DEFAULT 1)""")
        c.execute("""CREATE TABLE IF NOT EXISTS app_settings(key TEXT PRIMARY KEY, value TEXT)""")
        c.execute("""CREATE TABLE IF NOT EXISTS annotation_samples(id INTEGER PRIMARY KEY AUTOINCREMENT, created_at TEXT, source_file TEXT, page INTEGER, x1 INTEGER, y1 INTEGER, x2 INTEGER, y2 INTEGER, color_name TEXT, llm_answer TEXT, final_answer TEXT, is_llm_correct INTEGER, crop_path TEXT, memo TEXT)""")
        c.execute("""CREATE TABLE IF NOT EXISTS annotation_features(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            created_at TEXT,
            source_file TEXT,
            page INTEGER,
            final_answer TEXT,
            crop_path TEXT,
            feature_json TEXT,
            memo TEXT
        )""")
        c.execute("""CREATE TABLE IF NOT EXISTS symbol_image_dataset(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            created_at TEXT,
            symbol_name TEXT,
            source_file TEXT,
            crop_path TEXT,
            feature_json TEXT,
            memo TEXT
        )""")
        c.execute("""CREATE TABLE IF NOT EXISTS manual_cables(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            created_at TEXT,
            source_file TEXT,
            page INTEGER,
            cable_type TEXT,
            x1 INTEGER,
            y1 INTEGER,
            x2 INTEGER,
            y2 INTEGER,
            length_px REAL,
            length_m REAL,
            memo TEXT
        )""")
        c.execute("""CREATE TABLE IF NOT EXISTS vector_symbol_items(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            created_at TEXT,
            label TEXT,
            source_file TEXT,
            crop_path TEXT,
            embedding_json TEXT,
            backend TEXT,
            dim INTEGER,
            meta_json TEXT
        )""")
        c.execute("""CREATE TABLE IF NOT EXISTS vector_search_events(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            created_at TEXT,
            query_file TEXT,
            query_bbox TEXT,
            predicted_label TEXT,
            score REAL,
            topk_json TEXT,
            backend TEXT
        )""")
        c.execute("""CREATE TABLE IF NOT EXISTS todo_mindmaps(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            created_at TEXT,
            title TEXT,
            root_qty REAL DEFAULT 1,
            total_minutes REAL DEFAULT 0,
            json_data TEXT,
            memo TEXT
        )""")
        c.execute("""CREATE TABLE IF NOT EXISTS visual_node_graphs(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            created_at TEXT,
            title TEXT,
            nodes_json TEXT,
            edges_json TEXT,
            memo TEXT
        )""")
        c.execute("""CREATE TABLE IF NOT EXISTS langmem_memories(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            created_at TEXT,
            namespace TEXT,
            kind TEXT,
            text TEXT,
            tags TEXT,
            source TEXT,
            weight REAL DEFAULT 1.0,
            meta_json TEXT
        )""")
        c.execute("""CREATE TABLE IF NOT EXISTS dqn_strategy_events(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            created_at TEXT,
            source_file TEXT,
            page INTEGER,
            state_json TEXT,
            action TEXT,
            reward REAL DEFAULT 0,
            next_state_json TEXT,
            note TEXT
        )""")
        c.execute("""CREATE TABLE IF NOT EXISTS dqn_strategy_summary(
            action TEXT PRIMARY KEY,
            trials INTEGER DEFAULT 0,
            reward_sum REAL DEFAULT 0,
            reward_avg REAL DEFAULT 0,
            last_used TEXT
        )""")
        con.commit(); con.close(); self.seed_defaults()
    def seed_defaults(self):
        con = self.connect(); c = con.cursor()
        c.execute("SELECT COUNT(*) FROM unit_prices")
        if c.fetchone()[0] == 0:
            rows = [
                ('照明器具','LEDダウンライト','埋込型 φ150','台',12000,'ダウンライト,DL,ＤＬ,purple'),
                ('照明器具','LEDベースライト','埋込型 40形','台',15000,'ベースライト,BL,ＢＬ,yellow'),
                ('照明器具','非常灯','LED 20形','台',28000,'非常灯,emergency'),
                ('照明器具','誘導灯','LED 避難口','台',25000,'誘導灯,exit'),
                ('コンセント','コンセント','2P 15A 接地極付','個',3500,'コンセント,CO,ＣＯ,green'),
                ('コンセント','防水コンセント','2P 15A','個',6500,'防水コンセント'),
                ('コンセント','フロアコンセント','2P 15A','個',8500,'フロアコンセント,FC,ＦＣ'),
                ('スイッチ','片切スイッチ','15A 埋込','個',2200,'スイッチ,SW,ＳＷ,blue,cyan'),
                ('スイッチ','3路スイッチ','15A 埋込','個',2800,'3路スイッチ'),
                ('分電盤','分電盤','20回路 壁埋込','面',85000,'分電盤,DB,ＤＢ,red'),
                ('分電盤','主配電盤','750A','面',450000,'主配電盤,MDB,ＭＤＢ'),
                ('配線','VVF','2.0-2C','m',250,'VVF,2.0-2C'),
                ('配線','VVF','2.0-3C','m',320,'VVF,2.0-3C'),
                ('配線','CV','60sq-3C','m',2400,'CV,60sq'),
                ('空調','換気扇','φ100','台',18000,'換気扇,VF,ＶＦ'),
                ('空調','エアコン','5.0kW','台',180000,'エアコン,AC,ＡＣ'),
                ('その他','未分類設備','画像記号','個',0,'unknown,未分類'),
            ]
            c.executemany("INSERT INTO unit_prices(category,item_name,spec,unit,unit_price,keywords) VALUES(?,?,?,?,?,?)", rows)
        c.execute("SELECT COUNT(*) FROM symbol_patterns")
        if c.fetchone()[0] == 0:
            rows = [
                ('ダウンライト','simple_circle',json.dumps({'type':'simple_circle','radius_min':40,'radius_max':120},ensure_ascii=False),'単純な円形記号',1),
                ('ベースライト','simple_circle',json.dumps({'type':'simple_circle','radius_min':120,'radius_max':220},ensure_ascii=False),'大きな円形記号',1),
                ('コンセント','circle_with_line',json.dumps({'type':'circle_with_line','radius_min':30,'radius_max':90},ensure_ascii=False),'円+線',1),
                ('スイッチ','circle_with_line',json.dumps({'type':'circle_with_line','radius_min':30,'radius_max':90},ensure_ascii=False),'円+線',1),
                ('分電盤','rectangle_with_cross',json.dumps({'type':'rectangle_with_cross','width_min':80,'width_max':220},ensure_ascii=False),'矩形+×印',1),
            ]
            c.executemany("INSERT INTO symbol_patterns(name,pattern_type,pattern_json,description,preset) VALUES(?,?,?,?,?)", rows)
        c.execute("SELECT COUNT(*) FROM image_color_rules")
        if c.fetchone()[0] == 0:
            rules = [('purple','LEDダウンライト',120,255,0,150,120,255,1),('yellow','LEDベースライト',160,255,125,255,0,140,1),('green','コンセント',0,150,100,255,0,160,1),('cyan','片切スイッチ',0,130,120,255,120,255,1),('blue','片切スイッチ',0,120,0,170,120,255,1),('red','分電盤',150,255,0,140,0,140,1)]
            c.executemany("INSERT INTO image_color_rules(color_name,equipment,r_min,r_max,g_min,g_max,b_min,b_max,enabled) VALUES(?,?,?,?,?,?,?,?,?)", rules)
        defaults = {"ollama_url":DEFAULT_OLLAMA_URL,"ollama_model":DEFAULT_OLLAMA_MODEL,"pdf_scale":"1.25","analysis_max_width":"1000","preview_max_width":"850","scan_step":"4","cluster_distance":"6","min_cluster_points":"3","min_box_w":"4","min_box_h":"4","max_box_w":"55","max_box_h":"55","exclude_legend":"1","legend_right_ratio":"0.24","legend_bottom_ratio":"0.28","max_detections":"140","analysis_engine":"color","template_enabled":"1","opencv_enabled":"0","ocr_enabled":"0","yolo_enabled":"0","sam_enabled":"0","cnn_vit_enabled":"0","use_litellm":"0","use_langmem":"1","memory_top_k":"6","learning_match_threshold":"0.72","dqn_enabled":"1","dqn_epsilon":"0.15","dqn_learning_rate":"0.08","dqn_gamma":"0.90","dqn_last_action":"","db_backend":"sqlite","rqlite_url":"http://127.0.0.1:4001","vector_backend":"auto","clip_model_name":"openai/clip-vit-base-patch32","vector_top_k":"5","vector_threshold":"0.72","use_vector_search":"1"}
        for k,v in defaults.items(): c.execute("INSERT OR IGNORE INTO app_settings(key,value) VALUES(?,?)", (k,v))
        con.commit(); con.close()
    def get_setting(self,k,d=""):
        con=self.connect(); row=con.execute("SELECT value FROM app_settings WHERE key=?",(k,)).fetchone(); con.close(); return row[0] if row else d
    def set_setting(self,k,v):
        con=self.connect(); con.execute("INSERT OR REPLACE INTO app_settings(key,value) VALUES(?,?)",(k,str(v))); con.commit(); con.close()
    def get_settings(self):
        con=self.connect(); rows=con.execute("SELECT key,value FROM app_settings").fetchall(); con.close(); return dict(rows)
    def get_all_unit_prices(self):
        con=self.connect(); rows=con.execute("SELECT id,category,item_name,spec,unit,unit_price,keywords FROM unit_prices ORDER BY category,item_name").fetchall(); con.close(); return rows
    def upsert_unit_price(self,data,rid=None):
        con=self.connect()
        if rid: con.execute("UPDATE unit_prices SET category=?,item_name=?,spec=?,unit=?,unit_price=?,keywords=? WHERE id=?", (data['category'],data['item_name'],data['spec'],data['unit'],data['unit_price'],data['keywords'],rid))
        else: con.execute("INSERT INTO unit_prices(category,item_name,spec,unit,unit_price,keywords) VALUES(?,?,?,?,?,?)", (data['category'],data['item_name'],data['spec'],data['unit'],data['unit_price'],data['keywords']))
        con.commit(); con.close()
    def delete_unit_price(self,rid):
        con=self.connect(); con.execute("DELETE FROM unit_prices WHERE id=?",(rid,)); con.commit(); con.close()
    def import_unit_csv(self,path):
        con=self.connect(); imported=skipped=0
        with open(path,'r',encoding='utf-8-sig') as f:
            for row in csv.DictReader(f):
                try:
                    price=float(str(row.get('単価','0')).replace(',','').strip() or 0)
                    con.execute("INSERT INTO unit_prices(category,item_name,spec,unit,unit_price,keywords) VALUES(?,?,?,?,?,?)", (row.get('カテゴリ',''),row.get('品名',''),row.get('仕様',''),row.get('単位',''),price,row.get('キーワード',''))); imported+=1
                except Exception: skipped+=1
        con.commit(); con.close(); return imported, skipped
    def export_unit_csv(self,path):
        rows=self.get_all_unit_prices()
        with open(path,'w',encoding='utf-8-sig',newline='') as f:
            w=csv.writer(f); w.writerow(['カテゴリ','品名','仕様','単位','単価','キーワード'])
            for r in rows: w.writerow([r[1],r[2],r[3],r[4],r[5],r[6]])
        return len(rows)
    def find_price(self,name):
        con=self.connect(); row=con.execute("SELECT category,item_name,spec,unit,unit_price FROM unit_prices WHERE item_name=? OR keywords LIKE ? ORDER BY CASE WHEN item_name=? THEN 0 ELSE 1 END LIMIT 1", (name,f'%{name}%',name)).fetchone(); con.close()
        return row if row else ('その他',name,'推定','個',0)
    def get_all_symbol_patterns(self):
        con=self.connect(); rows=con.execute("SELECT id,name,pattern_type,pattern_json,description,preset FROM symbol_patterns ORDER BY preset DESC,name").fetchall(); con.close(); return rows
    def upsert_symbol_pattern(self,data,rid=None):
        con=self.connect()
        if rid: con.execute("UPDATE symbol_patterns SET name=?,pattern_type=?,pattern_json=?,description=? WHERE id=?", (data['name'],data['pattern_type'],data['pattern_json'],data['description'],rid))
        else: con.execute("INSERT INTO symbol_patterns(name,pattern_type,pattern_json,description,preset) VALUES(?,?,?,?,0)", (data['name'],data['pattern_type'],data['pattern_json'],data['description']))
        con.commit(); con.close()
    def delete_symbol_pattern(self,rid):
        con=self.connect(); con.execute("DELETE FROM symbol_patterns WHERE id=? AND preset=0",(rid,)); con.commit(); con.close()
    def get_cad_library(self):
        con=self.connect(); rows=con.execute("SELECT id,manufacturer,model_number,category,item_name,file_path,url FROM cad_library ORDER BY manufacturer,category,item_name").fetchall(); con.close(); return rows
    def register_cad_file(self,data):
        con=self.connect(); con.execute("INSERT INTO cad_library(manufacturer,model_number,category,item_name,file_path,pattern_signature,spec_json,url) VALUES(?,?,?,?,?,?,?,?)", (data['manufacturer'],data['model_number'],data['category'],data['item_name'],data['file_path'],data['pattern_signature'],data['spec_json'],data.get('url',''))); con.commit(); con.close()
    def delete_cad_file(self,rid):
        con=self.connect(); con.execute("DELETE FROM cad_library WHERE id=?",(rid,)); con.commit(); con.close()
    def get_image_rules(self):
        con=self.connect(); con.row_factory=sqlite3.Row; rows=[dict(r) for r in con.execute("SELECT * FROM image_color_rules ORDER BY id").fetchall()]; con.close(); return rows
    def save_image_rule(self,r):
        con=self.connect(); con.execute("UPDATE image_color_rules SET equipment=?,r_min=?,r_max=?,g_min=?,g_max=?,b_min=?,b_max=?,enabled=? WHERE color_name=?", (r['equipment'],int(r['r_min']),int(r['r_max']),int(r['g_min']),int(r['g_max']),int(r['b_min']),int(r['b_max']),int(r['enabled']),r['color_name'])); con.commit(); con.close()
    def save_annotation(self,source_file,det,llm_answer,final_answer,correct,crop_path,memo=''):
        con=self.connect(); con.execute("INSERT INTO annotation_samples(created_at,source_file,page,x1,y1,x2,y2,color_name,llm_answer,final_answer,is_llm_correct,crop_path,memo) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)", (datetime.now().isoformat(timespec='seconds'),source_file,det.page,det.x1,det.y1,det.x2,det.y2,det.color_name,llm_answer,final_answer,1 if correct else 0,crop_path,memo)); con.commit(); con.close()
    def save_annotation_feature(self, source_file, det, final_answer, crop_path, feature_json, memo=''):
        con=self.connect()
        con.execute("""INSERT INTO annotation_features(created_at,source_file,page,final_answer,crop_path,feature_json,memo)
                       VALUES(?,?,?,?,?,?,?)""",
                    (datetime.now().isoformat(timespec='seconds'),source_file,getattr(det,'page',0),final_answer,crop_path,feature_json,memo))
        con.commit(); con.close()
    def get_annotation_features(self, limit=2000):
        con=self.connect(); con.row_factory=sqlite3.Row
        rows=[dict(r) for r in con.execute("SELECT * FROM annotation_features ORDER BY id DESC LIMIT ?",(limit,)).fetchall()]
        con.close(); return rows
    def annotation_feature_summary(self):
        con=self.connect()
        rows=con.execute("SELECT final_answer,COUNT(*) FROM annotation_features GROUP BY final_answer ORDER BY final_answer").fetchall()
        con.close(); return rows
    def rebuild_features_from_annotations(self):
        # crop_pathが残っている過去アノテーションからfeature tableを再構築する。
        con=self.connect(); con.row_factory=sqlite3.Row
        rows=[dict(r) for r in con.execute("SELECT * FROM annotation_samples WHERE crop_path IS NOT NULL AND crop_path<>'' ORDER BY id").fetchall()]
        con.close()
        added=0; skipped=0
        for r in rows:
            try:
                cp=r.get('crop_path','')
                if not cp or not Path(cp).exists():
                    skipped+=1; continue
                img=Image.open(cp).convert('RGB')
                feat=extract_symbol_feature(img)
                con=self.connect()
                con.execute("""INSERT INTO annotation_features(created_at,source_file,page,final_answer,crop_path,feature_json,memo)
                               VALUES(?,?,?,?,?,?,?)""",
                            (datetime.now().isoformat(timespec='seconds'),r.get('source_file',''),r.get('page',0),r.get('final_answer',''),cp,json.dumps(feat,ensure_ascii=False),'rebuilt from annotation_samples'))
                con.commit(); con.close()
                added+=1
            except Exception:
                skipped+=1
        return added, skipped
    def save_memory(self, namespace, kind, text_value, tags='', source='', weight=1.0, meta_json=''):
        con=self.connect()
        con.execute("""INSERT INTO langmem_memories(created_at,namespace,kind,text,tags,source,weight,meta_json)
                       VALUES(?,?,?,?,?,?,?,?)""",
                    (datetime.now().isoformat(timespec='seconds'),namespace,kind,text_value,tags,source,float(weight),meta_json))
        con.commit(); con.close()

    def search_memories(self, query, namespace='default', limit=6):
        """
        LangMem互換の軽量SQLiteメモリ検索。
        langmemパッケージが無い環境でも、修正履歴・積算癖・記号ルールをRAG風に呼び出す。
        """
        con=self.connect(); con.row_factory=sqlite3.Row
        rows=[dict(r) for r in con.execute(
            "SELECT * FROM langmem_memories WHERE namespace IN (?, 'global', 'default') ORDER BY id DESC LIMIT 1000",
            (namespace,)
        ).fetchall()]
        con.close()
        q_terms=[t for t in re.split(r'\W+', str(query).lower()) if t]
        scored=[]
        for r in rows:
            blob=(str(r.get('text',''))+' '+str(r.get('tags',''))+' '+str(r.get('kind',''))).lower()
            score=0.0
            for t in q_terms:
                if t and t in blob:
                    score += 1.0
            # 日本語向け: 完全語分割できないため文字列包含も見る
            q=str(query).strip().lower()
            if q and q in blob:
                score += 3.0
            score *= float(r.get('weight') or 1.0)
            if score>0:
                r['_score']=score
                scored.append(r)
        scored.sort(key=lambda x:(x.get('_score',0), x.get('id',0)), reverse=True)
        return scored[:int(limit)]

    def memory_summary(self):
        con=self.connect()
        rows=con.execute("SELECT namespace,kind,COUNT(*) FROM langmem_memories GROUP BY namespace,kind ORDER BY namespace,kind").fetchall()
        con.close(); return rows
    def save_symbol_image_dataset(self, symbol_name, source_file, crop_path, feature_json, memo=''):
        con=self.connect()
        con.execute("""INSERT INTO symbol_image_dataset(created_at,symbol_name,source_file,crop_path,feature_json,memo)
                       VALUES(?,?,?,?,?,?)""",
                    (datetime.now().isoformat(timespec='seconds'),symbol_name,source_file,crop_path,feature_json,memo))
        con.commit(); con.close()

    def get_symbol_image_dataset(self, limit=3000):
        con=self.connect(); con.row_factory=sqlite3.Row
        rows=[dict(r) for r in con.execute("SELECT * FROM symbol_image_dataset ORDER BY id DESC LIMIT ?",(limit,)).fetchall()]
        con.close(); return rows

    def symbol_dataset_summary(self):
        con=self.connect()
        rows=con.execute("SELECT symbol_name,COUNT(*) FROM symbol_image_dataset GROUP BY symbol_name ORDER BY symbol_name").fetchall()
        con.close(); return rows

    def save_manual_cable(self, source_file, page, cable_type, x1, y1, x2, y2, length_px, length_m, memo=''):
        con=self.connect()
        con.execute("""INSERT INTO manual_cables(created_at,source_file,page,cable_type,x1,y1,x2,y2,length_px,length_m,memo)
                       VALUES(?,?,?,?,?,?,?,?,?,?,?)""",
                    (datetime.now().isoformat(timespec='seconds'),source_file,page,cable_type,x1,y1,x2,y2,length_px,length_m,memo))
        con.commit(); con.close()

    def count_annotation_features(self):
        con=self.connect()
        n=con.execute("SELECT COUNT(*) FROM annotation_features").fetchone()[0]
        con.close(); return n

    def crop_path_exists_in_features(self, crop_path):
        con=self.connect()
        row=con.execute("SELECT id FROM annotation_features WHERE crop_path=? LIMIT 1",(str(crop_path),)).fetchone()
        con.close(); return bool(row)

    def save_dqn_event(self, source_file, page, state, action, reward=0.0, next_state=None, note=''):
        con=self.connect()
        con.execute("""INSERT INTO dqn_strategy_events(created_at,source_file,page,state_json,action,reward,next_state_json,note)
                       VALUES(?,?,?,?,?,?,?,?)""",
                    (datetime.now().isoformat(timespec='seconds'),source_file,int(page or 0),
                     json.dumps(state or {},ensure_ascii=False),action,float(reward or 0),
                     json.dumps(next_state or {},ensure_ascii=False),note))
        row=con.execute("SELECT trials,reward_sum FROM dqn_strategy_summary WHERE action=?",(action,)).fetchone()
        if row:
            trials=int(row[0])+1; reward_sum=float(row[1])+float(reward or 0)
            con.execute("UPDATE dqn_strategy_summary SET trials=?,reward_sum=?,reward_avg=?,last_used=? WHERE action=?",
                        (trials,reward_sum,reward_sum/max(1,trials),datetime.now().isoformat(timespec='seconds'),action))
        else:
            con.execute("INSERT INTO dqn_strategy_summary(action,trials,reward_sum,reward_avg,last_used) VALUES(?,?,?,?,?)",
                        (action,1,float(reward or 0),float(reward or 0),datetime.now().isoformat(timespec='seconds')))
        con.commit(); con.close()

    def get_dqn_summary(self):
        con=self.connect()
        rows=con.execute("SELECT action,trials,reward_sum,reward_avg,last_used FROM dqn_strategy_summary ORDER BY reward_avg DESC,trials DESC").fetchall()
        con.close(); return rows

    def get_recent_dqn_events(self, limit=80):
        con=self.connect(); con.row_factory=sqlite3.Row
        rows=[dict(r) for r in con.execute("SELECT * FROM dqn_strategy_events ORDER BY id DESC LIMIT ?",(int(limit),)).fetchall()]
        con.close(); return rows

    def add_dqn_reward(self, action, reward, note='manual reward'):
        self.save_dqn_event('',0,{},action,float(reward),{},note)

    def save_vector_symbol_item(self, label, source_file, crop_path, embedding, backend, meta=None):
        con=self.connect()
        emb_json=json.dumps([float(x) for x in embedding], ensure_ascii=False)
        con.execute("""INSERT INTO vector_symbol_items(created_at,label,source_file,crop_path,embedding_json,backend,dim,meta_json)
                       VALUES(?,?,?,?,?,?,?,?)""",
                    (datetime.now().isoformat(timespec='seconds'),label,source_file,crop_path,emb_json,backend,len(embedding),json.dumps(meta or {},ensure_ascii=False)))
        con.commit(); con.close()

    def get_vector_symbol_items(self, limit=20000):
        con=self.connect(); con.row_factory=sqlite3.Row
        rows=[dict(r) for r in con.execute("SELECT * FROM vector_symbol_items ORDER BY id ASC LIMIT ?",(int(limit),)).fetchall()]
        con.close(); return rows

    def count_vector_symbol_items(self):
        con=self.connect()
        n=con.execute("SELECT COUNT(*) FROM vector_symbol_items").fetchone()[0]
        con.close(); return n

    def vector_crop_exists(self, crop_path, label=None):
        con=self.connect()
        if label:
            row=con.execute("SELECT id FROM vector_symbol_items WHERE crop_path=? AND label=? LIMIT 1",(str(crop_path),label)).fetchone()
        else:
            row=con.execute("SELECT id FROM vector_symbol_items WHERE crop_path=? LIMIT 1",(str(crop_path),)).fetchone()
        con.close(); return bool(row)

    def save_vector_search_event(self, query_file, query_bbox, predicted_label, score, topk, backend):
        con=self.connect()
        con.execute("""INSERT INTO vector_search_events(created_at,query_file,query_bbox,predicted_label,score,topk_json,backend)
                       VALUES(?,?,?,?,?,?,?)""",
                    (datetime.now().isoformat(timespec='seconds'),query_file,query_bbox,predicted_label,float(score or 0),json.dumps(topk or [],ensure_ascii=False),backend))
        con.commit(); con.close()

    def vector_summary(self):
        con=self.connect()
        rows=con.execute("SELECT label,backend,COUNT(*) FROM vector_symbol_items GROUP BY label,backend ORDER BY label,backend").fetchall()
        con.close(); return rows

    def save_todo_mindmap(self, title, root_qty, total_minutes, json_data, memo=''):
        con=self.connect()
        con.execute("""INSERT INTO todo_mindmaps(created_at,title,root_qty,total_minutes,json_data,memo)
                       VALUES(?,?,?,?,?,?)""",
                    (datetime.now().isoformat(timespec='seconds'),title,float(root_qty or 1),float(total_minutes or 0),json.dumps(json_data,ensure_ascii=False),memo))
        con.commit(); con.close()

    def get_todo_mindmaps(self, limit=50):
        con=self.connect(); con.row_factory=sqlite3.Row
        rows=[dict(r) for r in con.execute("SELECT * FROM todo_mindmaps ORDER BY id DESC LIMIT ?",(int(limit),)).fetchall()]
        con.close(); return rows

    def save_visual_node_graph(self, title, nodes, edges, memo=''):
        con=self.connect()
        con.execute("""INSERT INTO visual_node_graphs(created_at,title,nodes_json,edges_json,memo)
                       VALUES(?,?,?,?,?)""",
                    (datetime.now().isoformat(timespec='seconds'),title,json.dumps(nodes,ensure_ascii=False),json.dumps(edges,ensure_ascii=False),memo))
        con.commit(); con.close()

    def annotation_summary(self):
        con=self.connect(); rows=con.execute("SELECT final_answer,COUNT(*),SUM(is_llm_correct) FROM annotation_samples GROUP BY final_answer ORDER BY final_answer").fetchall(); con.close(); return rows



# ---------------- FAISS / CLIP / vector memory ----------------
class ClipEmbedder:
    """
    CLIP画像埋め込み。
    transformers/torch が使えるPCではCLIPを使い、
    Pydroid3や軽量環境ではOpenCV/Pillow特徴量の固定長ベクトルにフォールバックする。
    """
    _model_cache = {}

    def __init__(self, model_name=CLIP_MODEL_NAME_DEFAULT, backend='auto', log_fn=None):
        self.model_name=model_name or CLIP_MODEL_NAME_DEFAULT
        self.backend=backend or 'auto'
        self.log_fn=log_fn or (lambda s: None)
        self.active_backend='fallback'
        self.model=None
        self.processor=None
        self.torch=None

        if self.backend in ('auto','clip'):
            stack=try_import_clip_stack()
            if stack is not None:
                try:
                    torch, CLIPProcessor, CLIPModel = stack
                    cache_key=self.model_name
                    if cache_key in ClipEmbedder._model_cache:
                        self.torch,self.processor,self.model=ClipEmbedder._model_cache[cache_key]
                    else:
                        self.torch=torch
                        self.processor=CLIPProcessor.from_pretrained(self.model_name)
                        self.model=CLIPModel.from_pretrained(self.model_name)
                        self.model.eval()
                        ClipEmbedder._model_cache[cache_key]=(self.torch,self.processor,self.model)
                    self.active_backend='clip'
                except Exception as e:
                    self.log_fn(f'CLIP初期化失敗。fallbackへ切替: {e}')
                    self.active_backend='fallback'

    def embed_image(self, image):
        if self.active_backend=='clip' and self.model is not None:
            try:
                with self.torch.no_grad():
                    inputs=self.processor(images=image.convert('RGB'), return_tensors='pt')
                    feats=self.model.get_image_features(**inputs)
                    vec=feats[0].detach().cpu().float().numpy()
                    return self._normalize(vec.tolist()), 'clip'
            except Exception as e:
                self.log_fn(f'CLIP埋め込み失敗。fallbackへ切替: {e}')
        return self.fallback_embedding(image), 'fallback'

    def fallback_embedding(self, image):
        """
        Pydroid3でも動く軽量画像埋め込み。
        extract_symbol_featureの結果を128次元程度に展開。
        """
        feat=extract_symbol_feature(image)
        vec=[]
        vec.extend(feat.get('hist64',[])[:64])
        ah=feat.get('ahash',[])
        vec.extend([float(x) for x in ah[:64]])
        vec.append(float(feat.get('aspect',1))/4.0)
        vec.append(float(feat.get('edge_density',0)))
        cc=feat.get('color_counts',{}) or {}
        total=max(1,sum(float(v) for v in cc.values()))
        for k in ['purple','yellow','green','cyan','blue','red','dark']:
            vec.append(float(cc.get(k,0))/total)
        op=feat.get('opencv',{}) or {}
        for k in ['circularity','aspect','area_ratio']:
            try: vec.append(float(op.get(k,0)))
            except Exception: vec.append(0.0)
        while len(vec)<VECTOR_DIM_FALLBACK:
            vec.append(0.0)
        vec=vec[:VECTOR_DIM_FALLBACK]
        return self._normalize(vec)

    def _normalize(self, vec):
        s=sum(float(x)*float(x) for x in vec) ** 0.5
        if s<=0: return [0.0 for _ in vec]
        return [float(x)/s for x in vec]


class SymbolVectorEngine:
    """
    FAISS + CLIP/fallback埋め込みによる図面記号ベクトル検索。
    DBはSQLite/rqliteどちらでもDatabaseManager経由で利用。
    """
    def __init__(self, db, model_name=CLIP_MODEL_NAME_DEFAULT, backend='auto', log_fn=None):
        self.db=db
        self.log_fn=log_fn or (lambda s: None)
        self.embedder=ClipEmbedder(model_name=model_name, backend=backend, log_fn=self.log_fn)
        self.faiss=try_import_faiss()
        self.np=try_import_numpy()
        self.index=None
        self.meta=[]
        self.dim=None

    def add_crop(self, crop_path, label, source_file='', meta=None):
        img=Image.open(crop_path).convert('RGB')
        emb,backend=self.embedder.embed_image(img)
        self.db.save_vector_symbol_item(label, source_file or str(crop_path), str(crop_path), emb, backend, meta or {})
        return backend, len(emb)

    def rebuild_from_db(self):
        rows=self.db.get_vector_symbol_items(limit=50000)
        self.meta=[]
        vectors=[]
        for r in rows:
            try:
                emb=json.loads(r.get('embedding_json') or '[]')
                if not emb: continue
                vectors.append([float(x) for x in emb])
                self.meta.append(r)
            except Exception:
                continue
        if not vectors:
            self.index=None; self.dim=None
            return 0, 'none'

        self.dim=len(vectors[0])
        if self.faiss is not None and self.np is not None:
            arr=self.np.array(vectors, dtype='float32')
            # cosine similarity via normalized vectors + inner product
            index=self.faiss.IndexFlatIP(self.dim)
            index.add(arr)
            self.index=index
            return len(vectors), 'faiss'
        else:
            self.index=vectors
            return len(vectors), 'python'

    def search_image(self, image, top_k=5):
        emb,backend=self.embedder.embed_image(image)
        if self.index is None:
            self.rebuild_from_db()
        if self.index is None:
            return [], backend

        top_k=max(1,int(top_k or 5))
        if self.faiss is not None and self.np is not None and hasattr(self.index,'search'):
            arr=self.np.array([emb], dtype='float32')
            scores, idxs=self.index.search(arr, top_k)
            results=[]
            for score,idx in zip(scores[0],idxs[0]):
                if idx<0 or idx>=len(self.meta): continue
                m=self.meta[int(idx)]
                results.append({'label':m.get('label'), 'score':float(score), 'meta':m})
            return results, 'faiss+'+backend
        else:
            scored=[]
            for i,v in enumerate(self.index):
                score=sum(float(a)*float(b) for a,b in zip(emb,v))
                scored.append((score,i))
            scored.sort(reverse=True)
            results=[]
            for score,i in scored[:top_k]:
                m=self.meta[i]
                results.append({'label':m.get('label'), 'score':float(score), 'meta':m})
            return results, 'python+'+backend

    def save_faiss_index(self):
        if self.faiss is None or self.index is None or not hasattr(self.index,'ntotal'):
            return False
        try:
            self.faiss.write_index(self.index, str(FAISS_INDEX_PATH))
            with open(VECTOR_META_PATH,'w',encoding='utf-8') as f:
                for m in self.meta:
                    f.write(json.dumps(m,ensure_ascii=False)+'\n')
            return True
        except Exception:
            return False

    def load_faiss_index(self):
        if self.faiss is None or not FAISS_INDEX_PATH.exists() or not VECTOR_META_PATH.exists():
            return False
        try:
            self.index=self.faiss.read_index(str(FAISS_INDEX_PATH))
            self.meta=[]
            for line in VECTOR_META_PATH.read_text(encoding='utf-8').splitlines():
                if line.strip(): self.meta.append(json.loads(line))
            return True
        except Exception:
            return False


# ---------------- DQN strategy agent ----------------
class DQNStrategyAgent:
    """軽量DQN風の解析戦略エージェント。PyTorch不要。"""
    def __init__(self, db, model_path=DQN_MODEL_PATH):
        self.db=db; self.model_path=Path(model_path); self.actions=list(DQN_ACTIONS)
        self.feature_names=["image_area_norm","aspect","color_rule_count","feature_count_norm","dataset_count_norm","annotation_count_norm","last_error_rate","manual_ratio","huge_box_penalty","unknown_ratio"]
        self.weights={a:[0.0 for _ in self.feature_names] for a in self.actions}; self.bias={a:0.0 for a in self.actions}; self.load()
    def load(self):
        try:
            if self.model_path.exists():
                data=json.loads(self.model_path.read_text(encoding='utf-8')); self.weights.update(data.get('weights',{})); self.bias.update(data.get('bias',{}))
        except Exception: pass
    def save(self):
        try:
            self.model_path.parent.mkdir(parents=True,exist_ok=True); self.model_path.write_text(json.dumps({'weights':self.weights,'bias':self.bias},ensure_ascii=False,indent=2),encoding='utf-8')
        except Exception: pass
    def state_from_image(self,img,db,detections=None):
        detections=detections or []
        try: rules=len(db.get_image_rules())
        except Exception: rules=0
        try: dataset_count=len(db.get_symbol_image_dataset(limit=5000)) if hasattr(db,'get_symbol_image_dataset') else 0
        except Exception: dataset_count=0
        try: feat_count=db.count_annotation_features() if hasattr(db,'count_annotation_features') else 0
        except Exception: feat_count=0
        area=(img.width*img.height) if img else 0; aspect=(img.width/max(1,img.height)) if img else 1
        unknown=sum(1 for d in detections if getattr(d,'equipment','')=='未分類設備')
        huge=0
        for d in detections:
            try:
                if d.w*d.h > max(1,area)*0.012: huge+=1
            except Exception: pass
        n=max(1,len(detections))
        return {"image_area_norm":min(area/3000000.0,3.0),"aspect":min(aspect/4.0,2.0),"color_rule_count":min(rules/10.0,2.0),"feature_count_norm":min(feat_count/500.0,5.0),"dataset_count_norm":min(dataset_count/200.0,5.0),"annotation_count_norm":min(feat_count/1000.0,5.0),"last_error_rate":0.0,"manual_ratio":0.0,"huge_box_penalty":min(huge/n,1.0),"unknown_ratio":min(unknown/n,1.0)}
    def vector(self,state): return [float(state.get(k,0.0) or 0.0) for k in self.feature_names]
    def q(self,state,action):
        v=self.vector(state); w=self.weights.get(action,[0.0]*len(v)); return float(self.bias.get(action,0.0))+sum(float(a)*float(b) for a,b in zip(w,v))
    def select_action(self,state,epsilon=0.15):
        import random
        if random.random()<float(epsilon): return random.choice(self.actions),'explore'
        scores={a:self.q(state,a) for a in self.actions}; best=max(scores,key=scores.get); return best,'exploit'
    def update(self,state,action,reward,next_state=None,lr=0.08,gamma=0.90):
        v=self.vector(state); current=self.q(state,action); next_best=max([self.q(next_state,a) for a in self.actions], default=0.0) if next_state else 0.0
        target=float(reward)+float(gamma)*next_best; td=target-current; w=self.weights.setdefault(action,[0.0]*len(v))
        for i,x in enumerate(v): w[i]=float(w[i])+float(lr)*td*float(x)
        self.bias[action]=float(self.bias.get(action,0.0))+float(lr)*td; self.save(); return td
    def apply_action_to_settings(self,settings,action):
        s=dict(settings)
        if action=='color_fast': s.update({'analysis_engine':'color','scan_step':'4','cluster_distance':'8','max_box_w':'80','max_box_h':'80'})
        elif action=='color_strict': s.update({'analysis_engine':'color','scan_step':'3','cluster_distance':'4','max_box_w':'45','max_box_h':'45','max_detections':'220'})
        elif action=='opencv_shape': s.update({'opencv_enabled':'1','template_enabled':'1','cluster_distance':'5'})
        elif action=='template_match': s.update({'template_enabled':'1','cluster_distance':'5','max_box_w':'65','max_box_h':'65'})
        elif action=='ocr_assist': s.update({'ocr_enabled':'1','template_enabled':'1'})
        elif action=='learned_first': s.update({'template_enabled':'1','cluster_distance':'4','learning_match_threshold':'0.60','max_box_w':'60','max_box_h':'60'})
        elif action=='llm_assist': s.update({'template_enabled':'1','cluster_distance':'5'})
        elif action=='manual_annotation': s.update({'scan_step':'5','cluster_distance':'3','max_box_w':'50','max_box_h':'50','max_detections':'80'})
        return s
    def reward_from_detections(self,dets):
        if not dets: return -0.8
        enabled=[d for d in dets if getattr(d,'enabled',True)]; unknown=sum(1 for d in enabled if getattr(d,'equipment','')=='未分類設備'); learned=sum(1 for d in enabled if getattr(d,'source','') in ('learned','dataset','symbol_paste','manual_edit','drag')); huge=sum(1 for d in enabled if 'huge' in getattr(d,'memo',''))
        r=min(len(enabled)/50.0,0.8)+min(learned/10.0,1.0)-min(unknown/max(1,len(enabled)),1.0)-huge*0.4
        if len(enabled)>250: r-=1.0
        return r

# ---------------- PDF/DXF ----------------
class AdvancedDXFParser:
    def __init__(self): self.entities=[]; self.circles=[]; self.lines=[]; self.hatches=[]; self.texts=[]; self.blocks=[]; self.encoding_used=None
    def parse_dxf(self,path):
        self.__init__(); ezdxf=try_import_ezdxf()
        if ezdxf:
            try:
                doc=ezdxf.readfile(path); self.encoding_used='ezdxf'
                for e in doc.modelspace():
                    et=e.dxftype(); layer=getattr(e.dxf,'layer','0') if hasattr(e,'dxf') else '0'
                    if et=='CIRCLE':
                        item={'type':'CIRCLE','layer':layer,'data':{'x':float(e.dxf.center.x),'y':float(e.dxf.center.y),'radius':float(e.dxf.radius)}}; self.entities.append(item); self.circles.append(item)
                    elif et=='LINE':
                        item={'type':'LINE','layer':layer,'data':{'x':float(e.dxf.start.x),'y':float(e.dxf.start.y),'x2':float(e.dxf.end.x),'y2':float(e.dxf.end.y)}}; self.entities.append(item); self.lines.append(item)
                    elif et in ('LWPOLYLINE','POLYLINE'):
                        pts=[]
                        try: pts=[(float(p[0]),float(p[1])) for p in e.get_points()]
                        except Exception: pass
                        x=y=x2=y2=0.0
                        if pts: x,y=pts[0]; x2,y2=pts[-1]
                        item={'type':et,'layer':layer,'data':{'x':x,'y':y,'x2':x2,'y2':y2}}; self.entities.append(item); self.lines.append(item)
                    elif et=='HATCH': item={'type':'HATCH','layer':layer,'data':{}}; self.entities.append(item); self.hatches.append(item)
                    elif et in ('TEXT','MTEXT'):
                        try: txt=e.dxf.text if et=='TEXT' else e.text
                        except Exception: txt=''
                        item={'type':et,'layer':layer,'data':{'text':txt or ''}}; self.entities.append(item); self.texts.append(txt or '')
                    elif et=='INSERT': item={'type':'INSERT','layer':layer,'data':{'block_name':getattr(e.dxf,'name','')}}; self.entities.append(item); self.blocks.append(item)
                return
            except Exception: pass
        content, enc = read_text_file_safe(path); self.encoding_used=enc
        if not content: raise RuntimeError('DXFを読み込めません')
        cur=None; code=None; in_ent=False
        for line in content.splitlines():
            s=line.strip()
            if s=='ENTITIES': in_ent=True; continue
            if s=='ENDSEC':
                if cur: self._classify(cur)
                cur=None; in_ent=False; continue
            if not in_ent: continue
            if s.lstrip('-').isdigit(): code=int(s); continue
            if code==0:
                if cur: self._classify(cur)
                cur={'type':s,'layer':'0','data':{}}
            elif cur:
                if code==8: cur['layer']=s
                elif code==1: cur['data']['text']=s; self.texts.append(s)
                elif code==2: cur['data']['block_name']=s
                elif code in (10,11):
                    try: cur['data']['x' if code==10 else 'x2']=float(s)
                    except Exception: pass
                elif code in (20,21):
                    try: cur['data']['y' if code==20 else 'y2']=float(s)
                    except Exception: pass
                elif code==40:
                    try: cur['data']['radius']=float(s)
                    except Exception: pass
            code=None
        if cur: self._classify(cur)
    def _classify(self,e):
        self.entities.append(e)
        if e['type']=='CIRCLE': self.circles.append(e)
        elif e['type'] in ('LINE','LWPOLYLINE','POLYLINE'): self.lines.append(e)
        elif e['type']=='HATCH': self.hatches.append(e)
        elif e['type']=='INSERT': self.blocks.append(e)
    def signature(self): return {'circle_count':len(self.circles),'line_count':len(self.lines),'hatch_count':len(self.hatches),'block_count':len(self.blocks),'text_count':len(self.texts),'texts':self.texts[:20],'encoding':self.encoding_used}
    def count_by_patterns(self,rules):
        counts=defaultdict(int)
        for rule in rules:
            name=rule['name']; pat=rule.get('pattern',{}); typ=pat.get('type','simple_circle')
            if typ in ('simple_circle','circle_with_line','circle_with_radial_lines'):
                rmin=pat.get('radius_min',0); rmax=pat.get('radius_max',999999)
                for c in self.circles:
                    r=c['data'].get('radius',0)
                    if rmin<=r<=rmax: counts[name]+=1
            elif typ=='rectangle_with_cross':
                n=len([e for e in self.entities if e['type'] in ('LWPOLYLINE','POLYLINE')]); counts[name]+=max(0,min(n//5,50))
        return dict(counts)

def extract_pdf_text(path):
    fitz=try_import_fitz()
    if fitz:
        try:
            doc=fitz.open(path); txt='\\n'.join([p.get_text() for p in doc]); doc.close()
            if txt.strip(): return txt,'PyMuPDF'
        except Exception: pass
    PdfReader=try_import_pypdf()
    if PdfReader:
        try:
            reader=PdfReader(path); txt='\\n'.join([(p.extract_text() or '') for p in reader.pages])
            if txt.strip(): return txt,'pypdf'
        except Exception: pass
    return '', 'none'

def extract_counts_from_text(text, db):
    counts=defaultdict(int)
    for _,cat,item,spec,unit,price,keywords in db.get_all_unit_prices():
        keys=[item]+[k.strip() for k in (keywords or '').split(',') if k.strip()]
        for k in keys:
            for m in re.finditer(re.escape(k)+r'\s*[×xX＊*]?\s*(\d+)', text):
                try: counts[item]+=int(m.group(1))
                except Exception: pass
            if k in text and counts[item]==0:
                counts[item]+=min(text.count(k),3)
    return dict(counts)

# ---------------- image analysis ----------------
@dataclass
class ImageDetection:
    det_id:int; page:int; x1:int; y1:int; x2:int; y2:int; color_name:str; equipment:str; score:float; pixel_count:int; enabled:bool=True; source:str='auto'; memo:str=''
    @property
    def w(self): return max(1,self.x2-self.x1+1)
    @property
    def h(self): return max(1,self.y2-self.y1+1)

@dataclass
class ToriiShin:
    """通り芯: X軸=縦線(Y-axis line), Y軸=横線(X-axis line)"""
    id: int
    name: str        # X1, X2, Y1, Y2 等
    axis: str        # 'X'=縦線  'Y'=横線
    img_pos: float   # 画像ピクセル座標 (X→x値, Y→y値)
    page: int
    color: str = '#CC0000'
    enabled: bool = True

@dataclass
class CADSymbol:
    """スナップ配置された図面記号"""
    id: int
    name: str
    page: int
    img_x: float
    img_y: float
    ref_torii: str = ''        # 基準交点 例:'X1×Y1'
    offset_x_mm: float = 0.0
    offset_y_mm: float = 0.0

def load_pdf_or_image(path,pdf_scale,log=None):
    if Image is None: raise RuntimeError('Pillowが必要です')
    if Path(path).suffix.lower()=='.pdf':
        fitz=try_import_fitz()
        if not fitz: raise RuntimeError('画像PDF解析にはPyMuPDFが必要です')
        pages=[]; doc=fitz.open(path)
        for i,p in enumerate(doc):
            pix=p.get_pixmap(matrix=fitz.Matrix(pdf_scale,pdf_scale),alpha=False)
            img=Image.frombytes('RGB',[pix.width,pix.height],pix.samples); pages.append(img)
            if log: log(f'PDF {i+1}ページ目を画像化: {img.width}x{img.height}')
        doc.close(); return pages
    img=Image.open(path).convert('RGB')
    if log: log(f'画像読込: {img.width}x{img.height}')
    return [img]

def shrink_image(img,max_width):
    if img.width<=max_width: return img.copy(),1.0
    r=max_width/img.width
    return img.resize((max_width,int(img.height*r)),Image.Resampling.BILINEAR),r

def match_rgb(rgb,rule):
    r,g,b=rgb
    return int(rule['r_min'])<=r<=int(rule['r_max']) and int(rule['g_min'])<=g<=int(rule['g_max']) and int(rule['b_min'])<=b<=int(rule['b_max'])

def in_legend(x,y,w,h,s):
    if s.get('exclude_legend','1')!='1': return False
    return x>=int(w*(1-float(s.get('legend_right_ratio','0.24')))) and y>=int(h*(1-float(s.get('legend_bottom_ratio','0.28'))))

def cluster_points(points,dist):
    if not points: return []
    cell=max(3,int(dist)); buckets={}
    for x,y in points: buckets.setdefault((x//cell,y//cell),[]).append((x,y))
    used=set(); clusters=[]
    for key in list(buckets.keys()):
        if key in used: continue
        stack=[key]; used.add(key); pts=[]
        while stack:
            kx,ky=stack.pop(); pts.extend(buckets.get((kx,ky),[]))
            for nx in (kx-1,kx,kx+1):
                for ny in (ky-1,ky,ky+1):
                    nk=(nx,ny)
                    if nk in buckets and nk not in used: used.add(nk); stack.append(nk)
        clusters.append(pts)
    return clusters

def near(a,b,dist): return not (a.x2+dist<b.x1 or b.x2+dist<a.x1 or a.y2+dist<b.y1 or b.y2+dist<a.y1)

def merge_dets(dets, dist, max_merged_w=None, max_merged_h=None):
    """近接する同種検出を統合。max_merged_w/hを超えるマージは行わない（過大ボックス防止）。"""
    out=[]; used=[False]*len(dets)
    for i,d in enumerate(dets):
        if used[i]: continue
        group=[d]; used[i]=True; changed=True
        while changed:
            changed=False
            for j,e in enumerate(dets):
                if used[j] or e.equipment!=d.equipment: continue
                if any(near(g,e,dist) for g in group):
                    all_g = group + [e]
                    nw = max(g.x2 for g in all_g) - min(g.x1 for g in all_g)
                    nh = max(g.y2 for g in all_g) - min(g.y1 for g in all_g)
                    if max_merged_w and nw > max_merged_w * 2: continue
                    if max_merged_h and nh > max_merged_h * 2: continue
                    group.append(e); used[j]=True; changed=True
        if len(group)==1: out.append(d)
        else:
            out.append(ImageDetection(0,d.page,min(g.x1 for g in group),min(g.y1 for g in group),max(g.x2 for g in group),max(g.y2 for g in group),d.color_name,d.equipment,max(g.score for g in group),sum(g.pixel_count for g in group),True,'merged',f'merged {len(group)}'))
    for idx,d in enumerate(out,1): d.det_id=idx
    return out


def split_large_detection_smart(d, img, rules, settings, exp_w=55, exp_h=55):
    """
    大型マージ検出を「実際の記号位置」で分割する。
    均等グリッドではなく、元画像上の色クラスタ位置を再検出して個別ボックスを生成する。
    """
    bw=d.x2-d.x1; bh=d.y2-d.y1
    if bw*bh<100: return [d]

    # この大型ボックスにマッチする色ルールを取得
    matching_rule=next((r for r in rules if r.get('color_name')==d.color_name and int(r.get('enabled',1))==1), None)
    if not matching_rule:
        d.enabled=False; d.memo=(d.memo+' / large box disabled (no matching rule)').strip()
        return [d]

    # 大型ボックス内を切り出してサブ解析
    cx1=max(0,d.x1); cy1=max(0,d.y1); cx2=min(img.width,d.x2); cy2=min(img.height,d.y2)
    crop=img.crop((cx1,cy1,cx2,cy2))

    sub_s=dict(settings)
    sub_s['max_box_w']=str(min(exp_w,max(8,bw//2)))
    sub_s['max_box_h']=str(min(exp_h,max(8,bh//2)))
    sub_s['cluster_distance']=str(max(3,int(settings.get('cluster_distance','6'))))
    sub_s['analysis_max_width']=str(min(1000,max(bw,100)))
    sub_s['max_detections']='60'
    sub_s['exclude_legend']='0'  # サブ画像内に凡例なし

    try:
        sub_dets=analyze_image_page_light(crop,d.page,sub_s,[matching_rule])
        if not sub_dets or len(sub_dets)==1:
            # サブ解析でも分割できなかった → 無効化して手動アノテーション促す
            d.enabled=False; d.memo=(d.memo+' / large box: please annotate manually').strip()
            return [d]
        result=[]
        for sd in sub_dets:
            result.append(ImageDetection(
                0,d.page,cx1+sd.x1,cy1+sd.y1,cx1+sd.x2,cy1+sd.y2,
                sd.color_name,sd.equipment,
                sd.score,sd.pixel_count,True,'smart_split',
                f'smart_split {len(sub_dets)} from {bw}x{bh}'
            ))
        return result
    except Exception:
        d.enabled=False
        return [d]

def analyze_image_page_light(img,page_no,settings,rules,log=None):
    max_w=int(settings.get('analysis_max_width','1000')); step=max(1,int(settings.get('scan_step','4'))); dist=int(settings.get('cluster_distance','18'))
    min_pts=int(settings.get('min_cluster_points','3')); min_w=int(settings.get('min_box_w','4')); min_h=int(settings.get('min_box_h','4')); max_bw=int(settings.get('max_box_w','120')); max_bh=int(settings.get('max_box_h','120')); max_dets=int(settings.get('max_detections','140'))
    small,ratio=shrink_image(img,max_w); sw,sh=small.size; px=small.load(); dets=[]
    for rule in rules:
        if int(rule.get('enabled',1))!=1: continue
        pts=[]
        for y in range(0,sh,step):
            for x in range(0,sw,step):
                if in_legend(x,y,sw,sh,settings): continue
                if match_rgb(px[x,y],rule): pts.append((x,y))
        if log: log(f"page {page_no}: {rule['color_name']} sampled_points={len(pts)}")
        for cl in cluster_points(pts,dist):
            if len(cl)<min_pts: continue
            xs=[p[0] for p in cl]; ys=[p[1] for p in cl]; sx1,sx2=min(xs),max(xs); sy1,sy2=min(ys),max(ys); bw=sx2-sx1+1; bh=sy2-sy1+1
            if bw<min_w or bh<min_h or bw>max_bw or bh>max_bh: continue
            asp=bw/max(1,bh)
            if asp>8 or asp<0.125: continue
            inv=1/ratio; pad=int(5*inv); x1=max(0,int(sx1*inv)-pad); y1=max(0,int(sy1*inv)-pad); x2=min(img.width-1,int(sx2*inv)+pad); y2=min(img.height-1,int(sy2*inv)+pad)
            dets.append(ImageDetection(0,page_no,x1,y1,x2,y2,rule['color_name'],rule['equipment'],min(0.99,max(0.1,len(cl)/max(1,bw*bh))),len(cl)))
    img_max_bw=int(max_bw/max(0.1,ratio)); img_max_bh=int(max_bh/max(0.1,ratio))
    dets=merge_dets(dets,int(dist/max(0.1,ratio)),max_merged_w=img_max_bw,max_merged_h=img_max_bh); dets=sorted(dets,key=lambda d:d.pixel_count,reverse=True)[:max_dets]; dets=sorted(dets,key=lambda d:(d.page,d.y1,d.x1))
    for idx,d in enumerate(dets,1): d.det_id=idx
    if log: log(f'page {page_no}: 検出候補 {len(dets)}件')
    return dets

def draw_preview(img,dets,selected_id,max_width):
    base=img.convert('RGB').copy(); draw=ImageDraw.Draw(base); cmap={'purple':(170,0,255),'yellow':(230,170,0),'green':(0,180,0),'cyan':(0,170,190),'blue':(0,80,255),'red':(255,0,0),'manual':(255,0,255)}
    for d in dets:
        col=(150,150,150) if not d.enabled else cmap.get(d.color_name,(255,0,255))
        if d.det_id==selected_id: col=(255,0,0)
        draw.rectangle([d.x1,d.y1,d.x2,d.y2],outline=col,width=4 if d.det_id==selected_id else 2); draw.text((d.x1,max(0,d.y1-14)),str(d.det_id),fill=col)
    if base.width>max_width:
        r=max_width/base.width; base=base.resize((max_width,int(base.height*r)),Image.Resampling.BILINEAR)
    return base


def draw_symbol_glyph_on_canvas(canvas, x, y, name, scale=1.0, tag='manual_symbol'):
    """
    新JIS C 0303準拠の電気図記号をcanvasに描画する。
    英字ラベルではなく形状で識別できる図記号を描く。
    """
    s=max(10, int(20*scale))
    lw=max(1, int(2*scale))
    ids=[]
    col='#003366'
    name=str(name or '')
    if 'ダウンライト' in name:
        # JIS: 引込形照明 = 円 + 中央に十字線（LED表現）
        ids.append(canvas.create_oval(x-s,y-s,x+s,y+s,outline=col,width=lw,tags=(tag,)))
        ids.append(canvas.create_line(x-s//2,y,x+s//2,y,fill=col,width=lw,tags=(tag,)))
        ids.append(canvas.create_line(x,y-s//2,x,y+s//2,fill=col,width=lw,tags=(tag,)))
    elif 'ベースライト' in name:
        # JIS: 直管形蛍光灯/LEDベースライト = 細長矩形 + 中心線
        ids.append(canvas.create_rectangle(x-int(s*1.8),y-s//3,x+int(s*1.8),y+s//3,outline=col,width=lw,tags=(tag,)))
        ids.append(canvas.create_line(x-int(s*1.8),y,x+int(s*1.8),y,fill=col,width=lw,tags=(tag,)))
    elif 'コンセント' in name:
        # JIS: 円 + 右側に幕板線 + 幕板先端に横線（接地極なし片用）
        ids.append(canvas.create_oval(x-s,y-s,x+s,y+s,outline=col,width=lw,tags=(tag,)))
        ids.append(canvas.create_line(x+s,y,x+s+s//2,y,fill=col,width=lw,tags=(tag,)))
        ids.append(canvas.create_line(x+s+s//2,y-s//3,x+s+s//2,y+s//3,fill=col,width=lw,tags=(tag,)))
    elif 'スイッチ' in name or 'switch' in name.lower():
        # JIS: 円 + 上方向に引き出し線
        ids.append(canvas.create_oval(x-s,y-s,x+s,y+s,outline=col,width=lw,tags=(tag,)))
        ids.append(canvas.create_line(x,y-s,x,y-s*2,fill=col,width=lw,tags=(tag,)))
        ids.append(canvas.create_line(x,y-s*2,x+s//2,y-s*2+s//2,fill=col,width=lw,tags=(tag,)))
    elif '分電盤' in name:
        # JIS: 正方形 + 縦横仕切り線（盤記号）
        ids.append(canvas.create_rectangle(x-s,y-s,x+s,y+s,outline=col,width=lw,tags=(tag,)))
        ids.append(canvas.create_line(x-s,y,x+s,y,fill=col,width=lw,tags=(tag,)))
        ids.append(canvas.create_line(x,y-s,x,y+s,fill=col,width=lw,tags=(tag,)))
    elif '非常灯' in name:
        # JIS: 長方形の左半分塗りつぶし（非常照明器具）
        ids.append(canvas.create_rectangle(x-int(s*1.4),y-s//2,x+int(s*1.4),y+s//2,outline=col,width=lw,tags=(tag,)))
        ids.append(canvas.create_rectangle(x-int(s*1.4),y-s//2,x,y+s//2,fill=col,outline=col,tags=(tag,)))
    elif '誘導灯' in name:
        # JIS: 長方形 + 矢印（誘導灯）
        ids.append(canvas.create_rectangle(x-int(s*1.4),y-s//2,x+int(s*1.4),y+s//2,outline=col,width=lw,tags=(tag,)))
        ids.append(canvas.create_line(x-s//2,y,x+s//2,y,fill='white',width=lw,arrow=tk.LAST,tags=(tag,)))
    elif '換気扇' in name:
        # JIS: 正方形 + 対角線X（換気扇）
        ids.append(canvas.create_rectangle(x-s,y-s,x+s,y+s,outline=col,width=lw,tags=(tag,)))
        ids.append(canvas.create_line(x-s+3,y-s+3,x+s-3,y+s-3,fill=col,width=lw,tags=(tag,)))
        ids.append(canvas.create_line(x+s-3,y-s+3,x-s+3,y+s-3,fill=col,width=lw,tags=(tag,)))
    else:
        # 汎用: 円に稲妻マーク
        ids.append(canvas.create_oval(x-s,y-s,x+s,y+s,outline=col,width=lw,tags=(tag,)))
        ids.append(canvas.create_text(x,y,text='⚡',fill=col,font=('Arial',max(8,int(11*scale)),'bold'),tags=(tag,)))
    # 機器名ラベル（記号の下）
    ids.append(canvas.create_text(x,y+s+13,text=name,fill=col,font=('Arial',max(7,int(9*scale))),tags=(tag,)))
    return ids


def crop_symbol(img,det,margin=28): return img.crop((max(0,det.x1-margin),max(0,det.y1-margin),min(img.width,det.x2+margin+1),min(img.height,det.y2+margin+1)))
def image_b64_png(img,max_size=320):
    import io
    im=img.convert('RGB').copy(); im.thumbnail((max_size,max_size)); buf=io.BytesIO(); im.save(buf,format='PNG'); return base64.b64encode(buf.getvalue()).decode('ascii')


def dxf_to_preview_image_and_detections(path, db, max_size=1600, log=None):
    """DXFを画像PDF解析タブ用のプレビュー画像と検出候補に変換する簡易レンダラ。"""
    parser=AdvancedDXFParser(); parser.parse_dxf(path)
    xs=[]; ys=[]
    for e in parser.entities:
        d=e.get('data',{})
        for k in ('x','x2'):
            if k in d: xs.append(float(d[k]))
        for k in ('y','y2'):
            if k in d: ys.append(float(d[k]))
        if e.get('type')=='CIRCLE':
            x=float(d.get('x',0)); y=float(d.get('y',0)); r=float(d.get('radius',1)); xs += [x-r,x+r]; ys += [y-r,y+r]
    if not xs or not ys:
        img=Image.new('RGB',(1000,700),'white'); return [img], []
    minx,maxx,miny,maxy=min(xs),max(xs),min(ys),max(ys)
    w=max(1,maxx-minx); h=max(1,maxy-miny); margin=40
    scale=min((max_size-2*margin)/w,(max_size-2*margin)/h,1.0)
    iw=int(w*scale+2*margin); ih=int(h*scale+2*margin)
    img=Image.new('RGB',(max(iw,600),max(ih,400)),'white'); draw=ImageDraw.Draw(img)
    def tx(x): return int((float(x)-minx)*scale+margin)
    def ty(y): return int((maxy-float(y))*scale+margin)
    dets=[]; did=1
    rules=[]
    try:
        for r in db.get_all_symbol_patterns(): rules.append({'name':r[1],'pattern':json.loads(r[3])})
    except Exception: pass
    for e in parser.entities:
        typ=e.get('type'); d=e.get('data',{})
        if typ=='LINE':
            draw.line([tx(d.get('x',0)),ty(d.get('y',0)),tx(d.get('x2',0)),ty(d.get('y2',0))],fill=(40,40,40),width=1)
        elif typ in ('LWPOLYLINE','POLYLINE'):
            draw.line([tx(d.get('x',0)),ty(d.get('y',0)),tx(d.get('x2',0)),ty(d.get('y2',0))],fill=(80,80,80),width=1)
        elif typ=='CIRCLE':
            x=float(d.get('x',0)); y=float(d.get('y',0)); r=float(d.get('radius',1)); x1=tx(x-r); x2=tx(x+r); y1=ty(y+r); y2=ty(y-r)
            draw.ellipse([x1,y1,x2,y2],outline=(170,0,255),width=2)
            equip='未分類設備'
            for rule in rules:
                p=rule.get('pattern',{})
                if p.get('type') in ('simple_circle','circle_with_line','circle_with_radial_lines') and p.get('radius_min',0)<=r<=p.get('radius_max',999999):
                    equip=rule.get('name','未分類設備'); break
            dets.append(ImageDetection(did,1,min(x1,x2),min(y1,y2),max(x1,x2),max(y1,y2),'dxf',equip,0.8,1,True,'dxf','DXF circle'))
            did+=1
        elif typ=='INSERT':
            name=d.get('block_name','') or 'BLOCK'; equip='未分類設備'
            for _,cat,item,spec,unit,price,kw in db.get_all_unit_prices():
                if item in name or any(k and k.lower() in name.lower() for k in (kw or '').split(',')):
                    equip=item; break
            x=tx(d.get('x',minx)); y=ty(d.get('y',maxy)); draw.rectangle([x-8,y-8,x+8,y+8],outline=(255,0,0),width=2)
            dets.append(ImageDetection(did,1,x-8,y-8,x+8,y+8,'dxf',equip,0.7,1,True,'dxf','DXF insert '+name)); did+=1
    if log: log('DXF画像化: '+json.dumps(parser.signature(),ensure_ascii=False))
    return [img], dets



def infer_label_from_crop_filename(path):
    """
    symbol_cropsのファイル名から正解ラベルを推定する。
    重要: 推定不能な画像は None を返す。
    Noneの画像を教師データとして使うと、未分類教師が増えて頓珍漢になるため。
    """
    name = Path(path).stem.lower()
    mapping = [
        ('ledダウンライト','LEDダウンライト'), ('ダウンライト','LEDダウンライト'), ('downlight','LEDダウンライト'), ('_dl','LEDダウンライト'), ('-dl','LEDダウンライト'),
        ('ledベースライト','LEDベースライト'), ('ベースライト','LEDベースライト'), ('baselight','LEDベースライト'), ('_bl','LEDベースライト'), ('-bl','LEDベースライト'),
        ('コンセント','コンセント'), ('outlet','コンセント'), ('socket','コンセント'), ('_co','コンセント'), ('-co','コンセント'),
        ('片切スイッチ','片切スイッチ'), ('スイッチ','片切スイッチ'), ('switch','片切スイッチ'), ('_sw','片切スイッチ'), ('-sw','片切スイッチ'),
        ('分電盤','分電盤'), ('panel','分電盤'), ('_db','分電盤'), ('-db','分電盤'),
        ('非常灯','非常灯'), ('emergency','非常灯'),
        ('誘導灯','誘導灯'), ('exit','誘導灯'),
        ('換気扇','換気扇'), ('fan','換気扇'),
    ]
    for key,label in mapping:
        if key and key in name:
            return label
    return None


def extract_symbol_feature(crop):
    """
    再学習・再推論用の軽量特徴量 v12.12。
    色だけに寄りすぎると誤判定が増えるため、形状寄りの特徴も入れる。
    """
    im = crop.convert('RGB')
    w, h = im.size
    small = im.resize((32,32), Image.Resampling.BILINEAR)
    px = small.load()

    hist = [0]*64
    gray_vals=[]
    color_counts = {'purple':0,'yellow':0,'green':0,'cyan':0,'blue':0,'red':0,'dark':0}
    for y in range(32):
        for x in range(32):
            r,g,b = px[x,y]
            gray=(r+g+b)//3
            gray_vals.append(gray)
            ri=min(3,r//64); gi=min(3,g//64); bi=min(3,b//64)
            hist[ri*16+gi*4+bi] += 1
            if r>120 and g<150 and b>120: color_counts['purple'] += 1
            if r>160 and g>120 and b<150: color_counts['yellow'] += 1
            if r<150 and g>100 and b<170: color_counts['green'] += 1
            if r<140 and g>120 and b>120: color_counts['cyan'] += 1
            if r<130 and g<170 and b>120: color_counts['blue'] += 1
            if r>150 and g<140 and b<140: color_counts['red'] += 1
            if r<95 and g<95 and b<95: color_counts['dark'] += 1

    total=max(1,sum(hist))
    hist=[v/total for v in hist]
    major_color=max(color_counts, key=color_counts.get)

    # 8x8 average hash
    tiny=im.convert('L').resize((8,8), Image.Resampling.BILINEAR)
    vals=list(tiny.getdata())
    avg=sum(vals)/max(1,len(vals))
    ahash=[1 if v<avg else 0 for v in vals]  # 黒線寄りを1

    # edge density approximation
    gray=small.convert('L')
    gp=gray.load()
    edge=0
    for y in range(1,31):
        for x in range(1,31):
            if abs(gp[x+1,y]-gp[x-1,y]) + abs(gp[x,y+1]-gp[x,y-1]) > 70:
                edge+=1
    edge_density=edge/(30*30)

    feat={
        'w':w,'h':h,'aspect':w/max(1,h),
        'hist64':hist,
        'major_color':major_color,
        'color_counts':color_counts,
        'ahash':ahash,
        'edge_density':edge_density,
    }
    cvfeat=opencv_shape_features(im)
    feat['opencv']=cvfeat if isinstance(cvfeat,dict) else {}
    return feat


def feature_similarity(a, b):
    """
    0.0〜1.0の類似度 v12.12。
    色だけでなく、8x8形状ハッシュと縦横比を重視する。
    """
    try:
        ha=a.get('hist64',[]); hb=b.get('hist64',[])
        hist_sim=sum(min(x,y) for x,y in zip(ha,hb)) if ha and hb and len(ha)==len(hb) else 0.0

        aa=a.get('ahash',[]); ab=b.get('ahash',[])
        if aa and ab and len(aa)==len(ab):
            hash_sim=sum(1 for x,y in zip(aa,ab) if x==y)/len(aa)
        else:
            hash_sim=0.0

        asp_a=float(a.get('aspect',1)); asp_b=float(b.get('aspect',1))
        aspect_sim=max(0.0,1.0-min(abs(asp_a-asp_b)/2.5,1.0))

        ed_a=float(a.get('edge_density',0)); ed_b=float(b.get('edge_density',0))
        edge_sim=max(0.0,1.0-min(abs(ed_a-ed_b)/0.5,1.0))

        color_sim=1.0 if a.get('major_color')==b.get('major_color') else 0.55

        cva=a.get('opencv',{}) or {}; cvb=b.get('opencv',{}) or {}
        circ_a=float(cva.get('circularity',0) or 0); circ_b=float(cvb.get('circularity',0) or 0)
        circ_sim=max(0.0,1.0-min(abs(circ_a-circ_b),1.0)) if (circ_a or circ_b) else 0.5

        return max(0.0,min(1.0,
            hash_sim*0.35 +
            aspect_sim*0.20 +
            edge_sim*0.15 +
            hist_sim*0.15 +
            color_sim*0.08 +
            circ_sim*0.07
        ))
    except Exception:
        return 0.0


def best_symbol_dataset_match_for_crop(crop, dataset_rows, threshold=0.72):
    """
    symbol_image_dataset の51件と crop の類似度を比較して最良ラベルを返す。
    best_learned_match_for_crop の dataset版。
    戻り値: (label, score, row)
    """
    try:
        feat = extract_symbol_feature(crop)
        best_label, best_score, best_row = None, 0.0, None
        for r in dataset_rows:
            try:
                rf = json.loads(r.get('feature_json') or '{}')
                if not rf:
                    continue
                score = feature_similarity(feat, rf)
                if score > best_score:
                    best_score = score
                    # symbol_image_dataset は symbol_name フィールドを使う
                    best_label = r.get('symbol_name') or r.get('final_answer') or r.get('label')
                    best_row = r
            except Exception:
                continue
        if best_label and best_score >= threshold:
            return best_label, best_score, best_row
        return None, best_score, best_row
    except Exception:
        return None, 0.0, None


def best_learned_match_for_crop(crop, feature_rows, threshold=0.82):
    """
    crop画像に最も近い学習済み記号を探す。
    戻り値: (label, score, row)
    """
    try:
        feat = extract_symbol_feature(crop)
        best_label, best_score, best_row = None, 0.0, None
        for r in feature_rows:
            try:
                rf = json.loads(r.get('feature_json') or '{}')
                score = feature_similarity(feat, rf)
                if score > best_score:
                    best_score = score
                    best_label = r.get('final_answer')
                    best_row = r
            except Exception:
                continue
        if best_label and best_score >= threshold:
            return best_label, best_score, best_row
        return None, best_score, best_row
    except Exception:
        return None, 0.0, None

def simple_template_features(crop):
    try:
        im = crop.convert('RGB')
        w, h = im.size
        px = im.resize((32, 32), Image.Resampling.BILINEAR).load()
        colors = {'purple':0,'yellow':0,'green':0,'cyan':0,'blue':0,'red':0,'dark':0}
        for y in range(32):
            for x in range(32):
                r,g,b = px[x,y]
                if r>120 and g<150 and b>120: colors['purple'] += 1
                if r>160 and g>120 and b<150: colors['yellow'] += 1
                if r<150 and g>100 and b<170: colors['green'] += 1
                if r<140 and g>120 and b>120: colors['cyan'] += 1
                if r<130 and g<170 and b>120: colors['blue'] += 1
                if r>150 and g<140 and b<140: colors['red'] += 1
                if r<80 and g<80 and b<80: colors['dark'] += 1
        major = max(colors, key=colors.get)
        aspect = w / max(1, h)
        if major == 'purple': return 'LEDダウンライト', {'major_color':major,'aspect':aspect,'colors':colors}
        if major == 'yellow': return 'LEDベースライト', {'major_color':major,'aspect':aspect,'colors':colors}
        if major == 'green': return 'コンセント', {'major_color':major,'aspect':aspect,'colors':colors}
        if major in ('cyan','blue'): return '片切スイッチ', {'major_color':major,'aspect':aspect,'colors':colors}
        if major == 'red': return '分電盤', {'major_color':major,'aspect':aspect,'colors':colors}
        if aspect > 2.2: return 'LEDベースライト', {'major_color':major,'aspect':aspect,'colors':colors}
        return '未分類設備', {'major_color':major,'aspect':aspect,'colors':colors}
    except Exception as e:
        return '未分類設備', {'error':str(e)}

def opencv_shape_features(crop):
    cv2 = try_import_cv2()
    np = try_import_numpy()
    if cv2 is None or np is None:
        return None
    try:
        arr = np.array(crop.convert('RGB'))
        gray = cv2.cvtColor(arr, cv2.COLOR_RGB2GRAY)
        _, th = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV+cv2.THRESH_OTSU)
        contours, _ = cv2.findContours(th, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        if not contours:
            return {'contours':0}
        c = max(contours, key=cv2.contourArea)
        area = float(cv2.contourArea(c))
        peri = float(cv2.arcLength(c, True))
        circularity = 4*math.pi*area/(peri*peri) if peri else 0
        x,y,w,h = cv2.boundingRect(c)
        return {'contours':len(contours),'area':area,'perimeter':peri,'circularity':circularity,'aspect':w/max(1,h)}
    except Exception as e:
        return {'error':str(e)}

def ocr_crop_text(crop):
    pytesseract = try_import_pytesseract()
    if pytesseract is None:
        return ''
    try:
        return pytesseract.image_to_string(crop, lang='jpn+eng').strip()
    except Exception:
        try:
            return pytesseract.image_to_string(crop, lang='eng').strip()
        except Exception:
            return ''

def yolo_detect_placeholder(crop, model_path=''):
    YOLO = try_import_ultralytics()
    if YOLO is None or not model_path:
        return None
    try:
        model = YOLO(model_path)
        return str(model(crop))
    except Exception as e:
        return 'YOLO_ERROR: '+str(e)

def quick_ollama_ping(url, timeout=2):
    requests = try_import_requests()
    if not requests:
        return False
    try:
        r = requests.get(url.rstrip('/') + '/api/tags', timeout=timeout)
        return r.status_code == 200
    except Exception:
        return False

class OllamaClient:
    def __init__(self,url): self.url=url.rstrip('/')
    def models(self):
        requests=try_import_requests()
        if not requests: raise RuntimeError('requestsが必要です')
        r=requests.get(self.url+'/api/tags',timeout=8)
        r.raise_for_status()
        return [m.get('name','') for m in r.json().get('models',[]) if m.get('name')]
    def show_model(self,model):
        requests=try_import_requests()
        if not requests: raise RuntimeError('requestsが必要です')
        r=requests.post(self.url+'/api/show',json={'model':model},timeout=20)
        r.raise_for_status()
        return r.json()

    def generate(self,model,prompt,images=None,timeout=60):
        requests=try_import_requests()
        if not requests: raise RuntimeError('requestsが必要です')
        payload={
            'model':model,
            'prompt':prompt,
            'stream':False,
            'keep_alive':'1m',
            'options':{'num_ctx':4096,'temperature':0.1,'num_predict':512}
        }
        if images:
            payload['images']=images
        try:
            r=requests.post(self.url+'/api/generate',json=payload,timeout=timeout)
            r.raise_for_status()
            return r.json().get('response','')
        except Exception as e:
            # OllamaのタイムアウトでTkinterを落とさない。
            # 呼び出し側でフォールバック処理できるようRuntimeErrorに統一。
            raise RuntimeError(f'Ollama応答タイムアウト/接続エラー: {e}')


def list_ollama_models_fallback(ollama_exe_candidates=None):
    """
    /api/tags が失敗した場合に ollama list からモデル一覧を取得するフォールバック。
    """
    candidates = ollama_exe_candidates or [
        r"E:\Ollama\ollama.exe",
        r"D:\Ollama\ollama.exe",
        "ollama"
    ]
    for exe in candidates:
        try:
            proc = subprocess.run(
                [exe, "list"],
                capture_output=True,
                text=True,
                timeout=15,
                encoding="utf-8",
                errors="ignore"
            )
            if proc.returncode != 0:
                continue
            lines = proc.stdout.splitlines()
            models = []
            for line in lines[1:]:
                parts = line.split()
                if parts:
                    name = parts[0].strip()
                    if name and ":" in name:
                        models.append(name)
            if models:
                return models
        except Exception:
            continue
    return []

def list_openai_models(api_key):
    requests=try_import_requests()
    if not requests or not api_key: return []
    r=requests.get('https://api.openai.com/v1/models',headers={'Authorization':f'Bearer {api_key}'},timeout=20)
    r.raise_for_status()
    ids=[m.get('id','') for m in r.json().get('data',[]) if m.get('id')]
    return sorted(ids)

def list_anthropic_models(api_key):
    requests=try_import_requests()
    if not requests or not api_key: return []
    r=requests.get('https://api.anthropic.com/v1/models',headers={'x-api-key':api_key,'anthropic-version':'2023-06-01'},timeout=20)
    r.raise_for_status()
    ids=[]
    for m in r.json().get('data',[]):
        if isinstance(m,dict): ids.append(m.get('id') or m.get('name') or '')
    return [x for x in ids if x]


class LongTermMemoryManager:
    """
    LangMem導入層。
    langmemパッケージが入っていれば将来的に差し替え可能。
    現状はSQLiteをバックエンドにしたLangMem互換の長期記憶として動作する。
    """
    def __init__(self, db):
        self.db = db
        self.langmem = try_import_langmem()

    def add(self, namespace, kind, text_value, tags='', source='', weight=1.0, meta=None):
        meta_json=json.dumps(meta or {}, ensure_ascii=False)
        self.db.save_memory(namespace, kind, text_value, tags, source, weight, meta_json)

    def context(self, query, namespace='default', top_k=6):
        rows=self.db.search_memories(query, namespace=namespace, limit=top_k)
        if not rows:
            return ''
        lines=['[長期記憶 / LangMem互換メモリ]']
        for r in rows:
            lines.append(f"- ({r.get('kind','memory')}, score={r.get('_score',0):.1f}) {r.get('text','')}")
        return '\n'.join(lines)

class AIResponse(str):
    def __new__(cls, text, provider='unknown', model='', input_tokens=0, output_tokens=0, cost_usd=0.0):
        obj = str.__new__(cls, text or '')
        obj.provider = provider
        obj.model = model
        obj.input_tokens = int(input_tokens or 0)
        obj.output_tokens = int(output_tokens or 0)
        obj.total_tokens = obj.input_tokens + obj.output_tokens
        obj.cost_usd = float(cost_usd or 0.0)
        return obj
    def usage_summary(self):
        return f"provider={self.provider} / model={self.model} / input={self.input_tokens} / output={self.output_tokens} / total={self.total_tokens} / cost≈${self.cost_usd:.6f}"

class UnifiedAIClient:
    """LiteLLM対応。dictスナップショット(スレッドセーフ)またはappインスタンスで初期化できる。"""
    def __init__(self, app_or_dict):
        if isinstance(app_or_dict, dict):
            self._d = app_or_dict; self.app = None
        else:
            self._d = None; self.app = app_or_dict
        self.last_response = AIResponse('', 'unknown')

    def _v(self, key, default=''):
        if self._d is not None:
            return self._d.get(key, default)
        m = {
            'provider':'ai_provider_var','ollama_url':'ollama_url_var','ollama_model':'ollama_model_var',
            'openai_api_key':'openai_api_key_var','openai_model':'openai_model_var',
            'anthropic_api_key':'anthropic_api_key_var','anthropic_model':'anthropic_model_var',
            'custom_openai_base_url':'custom_openai_base_url_var',
            'custom_openai_api_key':'custom_openai_api_key_var','custom_openai_model':'custom_openai_model_var',
            'use_litellm':'use_litellm_var','use_langmem':'use_langmem_var','memory_top_k':'memory_top_k_var'
        }
        a = m.get(key)
        return getattr(self.app, a).get().strip() if a and hasattr(self.app, a) else default

    def provider(self):
        p = self._v('provider')
        if not p and self.app:
            try: p = self.app.db.get_setting('ai_provider','ollama')
            except Exception: pass
        return p or 'ollama'

    def _cost(self, provider, model, in_tok, out_tok):
        table = OPENAI_PRICE_USD_PER_MTOK if provider in ('openai','custom_openai') else ANTHROPIC_PRICE_USD_PER_MTOK
        price = table.get(model)
        if price is None:
            for k,v in table.items():
                if model.startswith(k) or k in model:
                    price = v; break
        if price is None: return 0.0
        return (in_tok/1_000_000.0)*price[0] + (out_tok/1_000_000.0)*price[1]

    def use_litellm(self):
        return str(self._v('use_litellm','0')).lower() in ('1','true','yes','on')

    def _litellm_model_and_kwargs(self):
        p=self.provider()
        kwargs={}
        if p=='ollama':
            model='ollama/' + (self._v('ollama_model',DEFAULT_OLLAMA_MODEL) or DEFAULT_OLLAMA_MODEL)
            kwargs['api_base']=self._v('ollama_url',DEFAULT_OLLAMA_URL)
            raw_model=self._v('ollama_model',DEFAULT_OLLAMA_MODEL)
        elif p=='anthropic':
            raw_model=self._v('anthropic_model',DEFAULT_CLAUDE_MODEL) or DEFAULT_CLAUDE_MODEL
            model='anthropic/' + raw_model
            kwargs['api_key']=self._v('anthropic_api_key','')
        elif p=='openai':
            raw_model=self._v('openai_model',DEFAULT_OPENAI_MODEL) or DEFAULT_OPENAI_MODEL
            model=raw_model
            kwargs['api_key']=self._v('openai_api_key','')
        elif p=='custom_openai':
            raw_model=self._v('custom_openai_model','local-model') or 'local-model'
            model=raw_model
            kwargs['api_base']=self._v('custom_openai_base_url',DEFAULT_CUSTOM_OPENAI_BASE_URL).rstrip('/')
            key=self._v('custom_openai_api_key','')
            if key: kwargs['api_key']=key
        else:
            raw_model=self._v('ollama_model',DEFAULT_OLLAMA_MODEL)
            model='ollama/' + raw_model
            kwargs['api_base']=self._v('ollama_url',DEFAULT_OLLAMA_URL)
        return p, raw_model, model, kwargs

    def generate(self,prompt,timeout=30):
        # LiteLLMを優先。ただし未インストール/失敗時は既存の直接APIへ安全にフォールバック。
        if self.use_litellm():
            litellm=try_import_litellm()
            if litellm is not None:
                try:
                    return self.generate_litellm(prompt,timeout)
                except Exception as e:
                    # LiteLLM自体の失敗で全体停止させない
                    pass
        p=self.provider()
        if p=='openai': return self.generate_openai(prompt,timeout)
        if p=='anthropic': return self.generate_anthropic(prompt,timeout)
        if p=='custom_openai': return self.generate_custom_openai(prompt,timeout)
        return self.generate_ollama(prompt,timeout)

    def generate_litellm(self,prompt,timeout=30):
        litellm=try_import_litellm()
        if litellm is None:
            raise RuntimeError('litellmがインストールされていません')
        provider, raw_model, model, kwargs = self._litellm_model_and_kwargs()
        messages=[
            {'role':'system','content':'あなたは日本の電気設備図面と積算の補助AIです。回答は簡潔に。'},
            {'role':'user','content':prompt}
        ]
        resp=litellm.completion(model=model, messages=messages, temperature=0.1, max_tokens=800, timeout=timeout, **kwargs)
        txt=''
        try:
            txt=resp.choices[0].message.content or ''
        except Exception:
            txt=str(resp)
        usage={}
        try:
            usage=dict(resp.usage)
        except Exception:
            try: usage=resp.get('usage',{}) or {}
            except Exception: usage={}
        in_tok=int(usage.get('prompt_tokens') or usage.get('input_tokens') or 0)
        out_tok=int(usage.get('completion_tokens') or usage.get('output_tokens') or 0)
        if not in_tok: in_tok=max(1,len(prompt)//3)
        if not out_tok: out_tok=max(1,len(txt)//3)
        cost=0.0 if provider=='ollama' else self._cost(provider, raw_model, in_tok, out_tok)
        r=AIResponse(txt,provider,raw_model,in_tok,out_tok,cost)
        self.last_response=r; return r

    def generate_ollama(self,prompt,timeout=120):
        model=self._v('ollama_model',DEFAULT_OLLAMA_MODEL)
        txt=OllamaClient(self._v('ollama_url',DEFAULT_OLLAMA_URL)).generate(model,prompt,images=None,timeout=timeout)
        resp=AIResponse(txt,'ollama',model,len(prompt)//3,len(txt)//3,0.0)
        self.last_response=resp; return resp

    def generate_openai(self,prompt,timeout=30):
        requests=try_import_requests()
        if not requests: raise RuntimeError('requestsが必要です')
        key=self._v('openai_api_key'); model=self._v('openai_model') or DEFAULT_OPENAI_MODEL
        if not key: raise RuntimeError('OpenAI APIキーが未入力です')
        r=requests.post('https://api.openai.com/v1/chat/completions',
            headers={'Authorization':f'Bearer {key}','Content-Type':'application/json'},
            json={'model':model,'messages':[{'role':'system','content':'あなたは日本の電気設備図面と積算の補助AIです。回答は簡潔に。'},{'role':'user','content':prompt}], 'temperature':0.1, 'max_tokens':800},
            timeout=timeout); r.raise_for_status()
        data=r.json(); txt=data['choices'][0]['message']['content']
        u=data.get('usage',{}) or {}; in_tok=u.get('prompt_tokens',0); out_tok=u.get('completion_tokens',0)
        resp=AIResponse(txt,'openai',model,in_tok,out_tok,self._cost('openai',model,in_tok,out_tok))
        self.last_response=resp; return resp

    def generate_anthropic(self,prompt,timeout=30):
        requests=try_import_requests()
        if not requests: raise RuntimeError('requestsが必要です')
        key=self._v('anthropic_api_key'); model=self._v('anthropic_model') or DEFAULT_CLAUDE_MODEL
        if not key: raise RuntimeError('Anthropic APIキーが未入力です')
        r=requests.post('https://api.anthropic.com/v1/messages',
            headers={'x-api-key':key,'anthropic-version':'2023-06-01','content-type':'application/json'},
            json={'model':model,'max_tokens':800,'temperature':0.1,'system':'あなたは日本の電気設備図面と積算の補助AIです。回答は簡潔に。','messages':[{'role':'user','content':prompt}]},
            timeout=timeout); r.raise_for_status()
        data=r.json(); txt='\n'.join(p.get('text','') for p in data.get('content',[]) if isinstance(p,dict) and p.get('type')=='text').strip()
        u=data.get('usage',{}) or {}; in_tok=u.get('input_tokens',0); out_tok=u.get('output_tokens',0)
        resp=AIResponse(txt,'anthropic',model,in_tok,out_tok,self._cost('anthropic',model,in_tok,out_tok))
        self.last_response=resp; return resp

    def generate_custom_openai(self,prompt,timeout=30):
        requests=try_import_requests()
        if not requests: raise RuntimeError('requestsが必要です')
        base=self._v('custom_openai_base_url',DEFAULT_CUSTOM_OPENAI_BASE_URL).rstrip('/'); key=self._v('custom_openai_api_key'); model=self._v('custom_openai_model') or 'local-model'
        if not base: raise RuntimeError('OpenAI互換API Base URLが未入力です')
        hdrs={'Content-Type':'application/json'}
        if key: hdrs['Authorization']=f'Bearer {key}'
        r=requests.post(base+'/chat/completions',headers=hdrs,
            json={'model':model,'messages':[{'role':'system','content':'あなたは日本の電気設備図面と積算の補助AIです。回答は簡潔に。'},{'role':'user','content':prompt}], 'temperature':0.1, 'max_tokens':800},
            timeout=timeout); r.raise_for_status()
        data=r.json(); txt=data['choices'][0]['message']['content']
        u=data.get('usage',{}) or {}; in_tok=u.get('prompt_tokens',0); out_tok=u.get('completion_tokens',0)
        resp=AIResponse(txt,'custom_openai',model,in_tok,out_tok,self._cost('custom_openai',model,in_tok,out_tok))
        self.last_response=resp; return resp
# ---------------- dialogs ----------------
class UnitPriceDialog(tk.Toplevel):
    def __init__(self,parent,db,record=None):
        super().__init__(parent); self.db=db; self.record=record; self.result=False; self.title('単価マスター編集' if record else '単価マスター追加'); self.geometry('560x420'); self.vars={}
        f=ttk.Frame(self,padding=15); f.pack(fill=tk.BOTH,expand=True); labels=['カテゴリ','品名','仕様','単位','単価(円)','キーワード(カンマ区切り)']
        for i,l in enumerate(labels):
            ttk.Label(f,text=l).grid(row=i,column=0,sticky='e',pady=4); row=ttk.Frame(f); row.grid(row=i,column=1,sticky='ew',padx=8,pady=4); v=tk.StringVar(); e=ttk.Entry(row,textvariable=v); e.pack(side=tk.LEFT,fill=tk.X,expand=True); add_paste_button(row,e); self.vars[l]=v
        f.columnconfigure(1,weight=1)
        if record:
            vals=[record[1],record[2],record[3],record[4],str(record[5]),record[6]]
            for l,v in zip(labels,vals): self.vars[l].set(v)
        btn=ttk.Frame(f); btn.grid(row=len(labels),column=0,columnspan=2,pady=12); ttk.Button(btn,text='保存',command=self.save,width=15).pack(side=tk.LEFT,padx=8); ttk.Button(btn,text='キャンセル',command=self.destroy,width=15).pack(side=tk.LEFT,padx=8); self.grab_set()
    def save(self):
        try:
            data={'category':self.vars['カテゴリ'].get().strip(),'item_name':self.vars['品名'].get().strip(),'spec':self.vars['仕様'].get().strip(),'unit':self.vars['単位'].get().strip(),'unit_price':float(self.vars['単価(円)'].get().replace(',','').strip()),'keywords':self.vars['キーワード(カンマ区切り)'].get().strip()}
            if not data['item_name']: raise ValueError('品名が空です')
            self.db.upsert_unit_price(data,self.record[0] if self.record else None); self.result=True; self.destroy()
        except Exception as e: messagebox.showwarning('入力エラー',str(e),parent=self)

class SymbolPatternDialog(tk.Toplevel):
    def __init__(self,parent,db,record=None):
        super().__init__(parent); self.db=db; self.record=record; self.result=False; self.title('記号パターン編集' if record else '記号パターン追加'); self.geometry('640x520')
        f=ttk.Frame(self,padding=15); f.pack(fill=tk.BOTH,expand=True); self.v_name=tk.StringVar(); self.v_type=tk.StringVar(value='simple_circle'); self.v_desc=tk.StringVar(); self.v_min=tk.StringVar(value='40'); self.v_max=tk.StringVar(value='120')
        for i,(lab,var) in enumerate([('設備名',self.v_name),('最小半径/幅',self.v_min),('最大半径/幅',self.v_max),('説明',self.v_desc)]): ttk.Label(f,text=lab).grid(row=i,column=0,sticky='e',pady=3); ttk.Entry(f,textvariable=var).grid(row=i,column=1,sticky='ew')
        ttk.Label(f,text='タイプ').grid(row=4,column=0,sticky='e'); ttk.Combobox(f,textvariable=self.v_type,values=['simple_circle','circle_with_line','rectangle_with_cross','circle_with_radial_lines'],state='readonly').grid(row=4,column=1,sticky='w')
        f.columnconfigure(1,weight=1)
        if record:
            self.v_name.set(record[1]); self.v_type.set(record[2]); self.v_desc.set(record[4] or '')
            try:
                p=json.loads(record[3]); self.v_min.set(str(p.get('radius_min',p.get('width_min',40)))); self.v_max.set(str(p.get('radius_max',p.get('width_max',120))))
            except Exception: pass
        btn=ttk.Frame(f); btn.grid(row=5,column=0,columnspan=2,pady=12); ttk.Button(btn,text='保存',command=self.save).pack(side=tk.LEFT,padx=8); ttk.Button(btn,text='キャンセル',command=self.destroy).pack(side=tk.LEFT,padx=8); self.grab_set()
    def save(self):
        try:
            typ=self.v_type.get(); p={'type':typ}
            if typ=='rectangle_with_cross': p['width_min']=float(self.v_min.get()); p['width_max']=float(self.v_max.get())
            else: p['radius_min']=float(self.v_min.get()); p['radius_max']=float(self.v_max.get())
            data={'name':self.v_name.get().strip(),'pattern_type':typ,'pattern_json':json.dumps(p,ensure_ascii=False),'description':self.v_desc.get().strip()}
            if not data['name']: raise ValueError('設備名が空です')
            self.db.upsert_symbol_pattern(data,self.record[0] if self.record else None); self.result=True; self.destroy()
        except Exception as e: messagebox.showwarning('入力エラー',str(e),parent=self)

class CADRegisterDialog(tk.Toplevel):
    def __init__(self,parent,db,file_path):
        super().__init__(parent); self.db=db; self.file_path=file_path; self.result=False; self.signature={}; self.title('CADデータ登録'); self.geometry('650x560')
        f=ttk.Frame(self,padding=12); f.pack(fill=tk.BOTH,expand=True); self.v_mfr=tk.StringVar(); self.v_model=tk.StringVar(); self.v_cat=tk.StringVar(value='照明器具'); self.v_name=tk.StringVar()
        ttk.Label(f,text=os.path.basename(file_path),foreground='blue').grid(row=0,column=0,columnspan=2,sticky='w')
        for i,(lab,var) in enumerate([('メーカー',self.v_mfr),('型番',self.v_model),('カテゴリ',self.v_cat),('機器名',self.v_name)],1): ttk.Label(f,text=lab).grid(row=i,column=0,sticky='e',pady=3); ttk.Entry(f,textvariable=var).grid(row=i,column=1,sticky='ew')
        self.info=scrolledtext.ScrolledText(f,height=14); self.info.grid(row=5,column=0,columnspan=2,sticky='nsew',pady=8); f.columnconfigure(1,weight=1); f.rowconfigure(5,weight=1)
        btn=ttk.Frame(f); btn.grid(row=6,column=0,columnspan=2); ttk.Button(btn,text='登録',command=self.save).pack(side=tk.LEFT,padx=8); ttk.Button(btn,text='キャンセル',command=self.destroy).pack(side=tk.LEFT,padx=8); self.analyze(); self.grab_set()
    def analyze(self):
        try:
            p=AdvancedDXFParser(); p.parse_dxf(self.file_path); sig=p.signature(); self.signature=sig; self.info.insert('1.0',json.dumps(sig,ensure_ascii=False,indent=2)); text=(os.path.basename(self.file_path)+' '+' '.join(sig.get('texts',[]))).lower()
            if 'down' in text or 'dl' in text or 'ダウン' in text: self.v_cat.set('照明器具'); self.v_name.set('ダウンライト')
            elif 'コンセント' in text or 'co' in text: self.v_cat.set('コンセント'); self.v_name.set('コンセント')
        except Exception as e: self.info.insert('1.0',f'解析エラー: {e}')
    def save(self):
        data={'manufacturer':self.v_mfr.get(),'model_number':self.v_model.get(),'category':self.v_cat.get(),'item_name':self.v_name.get(),'file_path':self.file_path,'pattern_signature':json.dumps(self.signature,ensure_ascii=False),'spec_json':json.dumps({},ensure_ascii=False),'url':''}
        if not data['item_name']: messagebox.showwarning('入力エラー','機器名を入力してください',parent=self); return
        self.db.register_cad_file(data); self.result=True; self.destroy()


class ImageDetectionEditDialog(tk.Toplevel):
    def __init__(self, parent, det):
        super().__init__(parent)
        self.det=det; self.result=None
        self.title('画像検出候補の編集'); self.geometry('430x360')
        f=ttk.Frame(self,padding=12); f.pack(fill=tk.BOTH,expand=True)
        self.v_enabled=tk.BooleanVar(value=bool(det.enabled))
        self.v_equipment=tk.StringVar(value=det.equipment)
        self.v_color=tk.StringVar(value=det.color_name)
        self.v_x1=tk.StringVar(value=str(det.x1)); self.v_y1=tk.StringVar(value=str(det.y1))
        self.v_x2=tk.StringVar(value=str(det.x2)); self.v_y2=tk.StringVar(value=str(det.y2))
        self.v_score=tk.StringVar(value=str(det.score)); self.v_memo=tk.StringVar(value=getattr(det,'memo',''))
        ttk.Checkbutton(f,text='有効',variable=self.v_enabled).grid(row=0,column=0,columnspan=2,sticky='w',pady=4)
        rows=[('設備名',self.v_equipment),('色/種別',self.v_color),('x1',self.v_x1),('y1',self.v_y1),('x2',self.v_x2),('y2',self.v_y2),('score',self.v_score),('memo',self.v_memo)]
        for i,(lab,var) in enumerate(rows,1):
            ttk.Label(f,text=lab).grid(row=i,column=0,sticky='e',pady=3)
            if lab=='設備名':
                ttk.Combobox(f,textvariable=var,values=['LEDダウンライト','LEDベースライト','コンセント','片切スイッチ','分電盤','換気扇','非常灯','誘導灯','未分類設備'],width=28).grid(row=i,column=1,sticky='w')
            else:
                ttk.Entry(f,textvariable=var,width=30).grid(row=i,column=1,sticky='w')
        btn=ttk.Frame(f); btn.grid(row=10,column=0,columnspan=2,pady=12)
        ttk.Button(btn,text='保存',command=self.save).pack(side=tk.LEFT,padx=5)
        ttk.Button(btn,text='キャンセル',command=self.destroy).pack(side=tk.LEFT,padx=5)
        self.grab_set()
    def save(self):
        try:
            self.result={'enabled':self.v_enabled.get(),'equipment':self.v_equipment.get().strip() or '未分類設備','color_name':self.v_color.get().strip() or 'manual','x1':int(float(self.v_x1.get())),'y1':int(float(self.v_y1.get())),'x2':int(float(self.v_x2.get())),'y2':int(float(self.v_y2.get())),'score':float(self.v_score.get() or 0),'memo':self.v_memo.get().strip()}
            self.destroy()
        except Exception as e:
            messagebox.showwarning('入力エラー',str(e),parent=self)

class EstimateRowEditDialog(tk.Toplevel):
    def __init__(self, parent, values):
        super().__init__(parent)
        self.result=None
        self.title('画像積算行の編集'); self.geometry('460x320')
        f=ttk.Frame(self,padding=12); f.pack(fill=tk.BOTH,expand=True)
        labels=['カテゴリ','品名','単位','数量','単価','金額']; self.vars={}
        for i,lab in enumerate(labels):
            ttk.Label(f,text=lab).grid(row=i,column=0,sticky='e',pady=4)
            v=tk.StringVar(value=str(values[i] if i < len(values) else ''))
            ttk.Entry(f,textvariable=v,width=32).grid(row=i,column=1,sticky='w')
            self.vars[lab]=v
        btn=ttk.Frame(f); btn.grid(row=7,column=0,columnspan=2,pady=12)
        ttk.Button(btn,text='保存',command=self.save).pack(side=tk.LEFT,padx=5)
        ttk.Button(btn,text='キャンセル',command=self.destroy).pack(side=tk.LEFT,padx=5)
        self.grab_set()
    def save(self):
        try:
            qty=float(self.vars['数量'].get() or 0); price=float(self.vars['単価'].get() or 0)
            amount=float(self.vars['金額'].get() or qty*price)
            self.result=(self.vars['カテゴリ'].get(),self.vars['品名'].get(),self.vars['単位'].get(),qty,price,amount)
            self.destroy()
        except Exception as e:
            messagebox.showwarning('入力エラー',str(e),parent=self)


class SymbolImageDatasetDialog(tk.Toplevel):
    def __init__(self, parent, default_name='LEDダウンライト'):
        super().__init__(parent); self.result=None
        self.title('図面記号画像/PDFデータセット登録'); self.geometry('520x260')
        f=ttk.Frame(self,padding=12); f.pack(fill=tk.BOTH,expand=True)
        self.v_name=tk.StringVar(value=default_name); self.v_file=tk.StringVar(); self.v_memo=tk.StringVar()
        ttk.Label(f,text='記号名').grid(row=0,column=0,sticky='e',pady=4)
        ttk.Combobox(f,textvariable=self.v_name,values=['LEDダウンライト','LEDベースライト','コンセント','片切スイッチ','分電盤','換気扇','非常灯','誘導灯','未分類設備'],width=32).grid(row=0,column=1,sticky='w')
        ttk.Label(f,text='画像/PDF').grid(row=1,column=0,sticky='e',pady=4)
        ttk.Entry(f,textvariable=self.v_file,width=42).grid(row=1,column=1,sticky='w')
        ttk.Button(f,text='選択',command=self.select_file).grid(row=1,column=2,padx=4)
        ttk.Label(f,text='メモ').grid(row=2,column=0,sticky='e',pady=4)
        ttk.Entry(f,textvariable=self.v_memo,width=42).grid(row=2,column=1,sticky='w')
        ttk.Label(f,text='PDFの場合は1ページ目全体を記号サンプルとして登録します。\nより正確にするには画像として切り出した記号を登録してください。',foreground='blue').grid(row=3,column=0,columnspan=3,sticky='w',pady=8)
        btn=ttk.Frame(f); btn.grid(row=4,column=0,columnspan=3,pady=12)
        ttk.Button(btn,text='登録',command=self.save).pack(side=tk.LEFT,padx=6)
        ttk.Button(btn,text='キャンセル',command=self.destroy).pack(side=tk.LEFT,padx=6)
        self.grab_set()
    def select_file(self):
        p=filedialog.askopenfilename(filetypes=[('画像/PDF','*.png;*.jpg;*.jpeg;*.bmp;*.pdf'),('All','*.*')])
        if p: self.v_file.set(p)
    def save(self):
        if not self.v_name.get().strip() or not self.v_file.get().strip():
            messagebox.showwarning('未入力','記号名とファイルを指定してください',parent=self); return
        self.result={'name':self.v_name.get().strip(),'file':self.v_file.get().strip(),'memo':self.v_memo.get().strip()}
        self.destroy()

class CableInputDialog(tk.Toplevel):
    def __init__(self, parent, default_type='VVF2.0-3C', length_px=0.0, scale_init=None, auto_length_m=None):
        super().__init__(parent); self.result=None
        self.title('手書きケーブル登録'); self.geometry('430x300')
        f=ttk.Frame(self,padding=12); f.pack(fill=tk.BOTH,expand=True)
        scale_str=f'{scale_init:.2f}' if scale_init else '100'
        self.v_type=tk.StringVar(value=default_type); self.v_scale=tk.StringVar(value=scale_str)
        self.v_len_px=tk.StringVar(value=f'{length_px:.1f}'); self.v_memo=tk.StringVar()
        self._auto_length_m=auto_length_m
        ttk.Label(f,text='ケーブル種別').grid(row=0,column=0,sticky='e',pady=4)
        ttk.Combobox(f,textvariable=self.v_type,values=['VVF2.0-2C','VVF2.0-3C','VVF1.6-2C','VVF1.6-3C','CV5.5-3C','CV14-3C','CV60sq-3C'],width=26).grid(row=0,column=1,sticky='w')
        ttk.Label(f,text='縮尺 px/m（例:100px=1mなら100）').grid(row=1,column=0,sticky='e',pady=4)
        ttk.Entry(f,textvariable=self.v_scale,width=12).grid(row=1,column=1,sticky='w')
        ttk.Label(f,text='線長 px').grid(row=2,column=0,sticky='e',pady=4)
        ttk.Entry(f,textvariable=self.v_len_px,width=12,state='readonly').grid(row=2,column=1,sticky='w')
        ttk.Label(f,text='メモ').grid(row=3,column=0,sticky='e',pady=4)
        ttk.Entry(f,textvariable=self.v_memo,width=30).grid(row=3,column=1,sticky='w')
        if auto_length_m is not None:
            ttk.Label(f,text=f'※縮尺設定から自動計算: {auto_length_m:.2f}m  (確認して登録ボタンを押してください)',foreground='green',font=('Arial',9,'bold')).grid(row=4,column=0,columnspan=2,sticky='w',pady=4)
        ttk.Label(f,text='入力例: px/m=100 なら 100px を1mとして積算します。',foreground='blue').grid(row=5,column=0,columnspan=2,sticky='w',pady=4)
        btn=ttk.Frame(f); btn.grid(row=5,column=0,columnspan=2,pady=12)
        ttk.Button(btn,text='登録',command=self.save).pack(side=tk.LEFT,padx=6)
        ttk.Button(btn,text='キャンセル',command=self.destroy).pack(side=tk.LEFT,padx=6)
        self.grab_set()
    def save(self):
        try:
            scale=float(self.v_scale.get() or 100); length_px=float(self.v_len_px.get() or 0)
            # 縮尺設定からの自動計算値がある場合はそちらを使用
            length_m=self._auto_length_m if self._auto_length_m is not None else length_px/max(0.0001,scale)
            self.result={'type':self.v_type.get().strip(),'scale':scale,'length_px':length_px,'length_m':length_m,'memo':self.v_memo.get().strip()}
            self.destroy()
        except Exception as e:
            messagebox.showwarning('入力エラー',str(e),parent=self)


class SymbolPasteDialog(tk.Toplevel):
    def __init__(self, parent, names):
        super().__init__(parent)
        self.result=None
        self.title('図面記号貼り付け')
        self.geometry('360x180')
        f=ttk.Frame(self,padding=12); f.pack(fill=tk.BOTH,expand=True)
        self.v_name=tk.StringVar(value=(names[0] if names else 'LEDダウンライト'))
        ttk.Label(f,text='貼り付ける図面記号').pack(anchor='w')
        ttk.Combobox(f,textvariable=self.v_name,values=names or ['LEDダウンライト','LEDベースライト','コンセント','片切スイッチ','分電盤','換気扇','非常灯','誘導灯'],width=32).pack(fill=tk.X,pady=6)
        btn=ttk.Frame(f); btn.pack(pady=10)
        ttk.Button(btn,text='OK',command=self.ok).pack(side=tk.LEFT,padx=6)
        ttk.Button(btn,text='キャンセル',command=self.destroy).pack(side=tk.LEFT,padx=6)
        self.grab_set()
    def ok(self):
        self.result=self.v_name.get().strip()
        self.destroy()

class ManualCableEditDialog(tk.Toplevel):
    def __init__(self, parent, cable):
        super().__init__(parent)
        self.result=None
        self.title('手書きケーブル編集')
        self.geometry('430x300')
        f=ttk.Frame(self,padding=12); f.pack(fill=tk.BOTH,expand=True)
        self.v_type=tk.StringVar(value=cable.get('type','VVF2.0-3C'))
        self.v_len=tk.StringVar(value=str(round(float(cable.get('length_m',0)),2)))
        self.v_x1=tk.StringVar(value=str(cable.get('x1',0)))
        self.v_y1=tk.StringVar(value=str(cable.get('y1',0)))
        self.v_x2=tk.StringVar(value=str(cable.get('x2',0)))
        self.v_y2=tk.StringVar(value=str(cable.get('y2',0)))
        rows=[('ケーブル種別',self.v_type),('長さm',self.v_len),('x1',self.v_x1),('y1',self.v_y1),('x2',self.v_x2),('y2',self.v_y2)]
        for i,(lab,var) in enumerate(rows):
            ttk.Label(f,text=lab).grid(row=i,column=0,sticky='e',pady=4)
            if lab=='ケーブル種別':
                ttk.Combobox(f,textvariable=var,values=['VVF2.0-2C','VVF2.0-3C','VVF1.6-2C','VVF1.6-3C','CV5.5-3C','CV14-3C','CV60sq-3C'],width=26).grid(row=i,column=1,sticky='w')
            else:
                ttk.Entry(f,textvariable=var,width=18).grid(row=i,column=1,sticky='w')
        btn=ttk.Frame(f); btn.grid(row=7,column=0,columnspan=2,pady=12)
        ttk.Button(btn,text='保存',command=self.save).pack(side=tk.LEFT,padx=6)
        ttk.Button(btn,text='キャンセル',command=self.destroy).pack(side=tk.LEFT,padx=6)
        self.grab_set()
    def save(self):
        try:
            self.result={
                'type': self.v_type.get().strip() or 'VVF2.0-3C',
                'length_m': float(self.v_len.get() or 0),
                'x1': int(float(self.v_x1.get() or 0)),
                'y1': int(float(self.v_y1.get() or 0)),
                'x2': int(float(self.v_x2.get() or 0)),
                'y2': int(float(self.v_y2.get() or 0)),
            }
            self.destroy()
        except Exception as e:
            messagebox.showwarning('入力エラー',str(e),parent=self)


# ---------------- Mindmap / NodeCAD helpers ----------------
def default_mindmap_json(task, qty=1):
    task = task or "電気設備工事"
    is_outlet = ("コンセント" in task) or ("outlet" in task.lower())
    is_light = ("照明" in task) or ("ダウンライト" in task) or ("ライト" in task)
    if is_outlet:
        mats=[("埋込ダブルコンセント本体",1,"個"),("埋込コンセントプレート",1,"枚"),("埋込取付枠",1,"個"),("スイッチボックス",1,"個"),("VVF2.0-3C",5,"m")]
        tools=["電工ドライバー","VVFストリッパー","電工ナイフ","検電器","水平器"]
        steps=[("壁面穴あけ・ボックス設置",10),("ケーブル通線・被覆ストリップ",5),("器具結線・枠取付",7),("プレート取付・電圧確認",5)]
    elif is_light:
        mats=[("LEDダウンライト",1,"台"),("VVF1.6-2C",5,"m"),("差込コネクタ",2,"個"),("支持金物",1,"式")]
        tools=["ホルソー","電工ドライバー","VVFストリッパー","脚立","検電器"]
        steps=[("開口位置確認・墨出し",5),("天井開口",8),("配線・結線",7),("器具取付・点灯確認",5)]
    else:
        mats=[("主要機器",1,"式"),("ケーブル",5,"m"),("取付材料",1,"式")]
        tools=["電工工具一式","検電器","脚立"]
        steps=[("現地確認・墨出し",5),("材料取付",10),("配線・接続",10),("試験確認",5)]
    total=sum(m for _,m in steps)
    nodes=[{"id":"root","text":task,"parent":None,"type":"root","meta":{"qty":qty}},
           {"id":"m_root","text":"必要材料","parent":"root","type":"category"},
           {"id":"t_root","text":"必要工具","parent":"root","type":"category"},
           {"id":"s_root","text":"施工手順・時間","parent":"root","type":"category"},
           {"id":"c_root","text":"積算・歩掛り情報","parent":"root","type":"category"}]
    for i,(name,q,u) in enumerate(mats,1):
        nodes.append({"id":f"m{i}","text":name,"parent":"m_root","type":"material","meta":{"default_qty":q,"unit":u}})
    for i,name in enumerate(tools,1):
        nodes.append({"id":f"t{i}","text":name,"parent":"t_root","type":"tool","meta":{}})
    for i,(name,mins) in enumerate(steps,1):
        nodes.append({"id":f"s{i}","text":f"{name} ({mins}分)","parent":"s_root","type":"step","meta":{"order":i,"minutes":mins}})
    nodes.append({"id":"c1","text":f"想定合計時間 {total}分/箇所","parent":"c_root","type":"cost","meta":{"minutes_per_qty":total}})
    return {"project_task":task,"root_qty":qty,"total_estimated_minutes":total*qty,"mindmap_nodes":nodes}

def sanitize_ai_json(text_value):
    s=str(text_value or '').strip()
    if '```' in s:
        s=re.sub(r'```(?:json)?', '', s).replace('```','').strip()
    m=re.search(r'\{.*\}', s, flags=re.S)
    if m: s=m.group(0)
    return json.loads(s)

class NodePropertyDialog(tk.Toplevel):
    def __init__(self, parent, node=None):
        super().__init__(parent)
        self.result=None; self.title("ノード設定"); self.geometry("420x300")
        node=node or {}
        f=ttk.Frame(self,padding=10); f.pack(fill=tk.BOTH,expand=True)
        self.v_text=tk.StringVar(value=node.get('text',''))
        self.v_type=tk.StringVar(value=node.get('type','symbol'))
        self.v_qty=tk.StringVar(value=str(node.get('qty',1)))
        self.v_unit=tk.StringVar(value=node.get('unit','個'))
        self.v_symbol=tk.StringVar(value=node.get('symbol','LEDダウンライト'))
        rows=[('表示名',self.v_text),('ノード種別',self.v_type),('数量',self.v_qty),('単位',self.v_unit),('図記号',self.v_symbol)]
        for i,(lab,var) in enumerate(rows):
            ttk.Label(f,text=lab).grid(row=i,column=0,sticky='e',pady=4)
            if lab=='ノード種別':
                ttk.Combobox(f,textvariable=var,values=['symbol','cable','panel','switch','outlet','light','process','material','tool'],width=24).grid(row=i,column=1,sticky='w')
            elif lab=='図記号':
                ttk.Combobox(f,textvariable=var,values=['LEDダウンライト','LEDベースライト','コンセント','片切スイッチ','分電盤','非常灯','誘導灯','VVF2.0-3C'],width=24).grid(row=i,column=1,sticky='w')
            else:
                ttk.Entry(f,textvariable=var,width=28).grid(row=i,column=1,sticky='w')
        btn=ttk.Frame(f); btn.grid(row=6,column=0,columnspan=2,pady=12)
        ttk.Button(btn,text='OK',command=self.ok).pack(side=tk.LEFT,padx=5)
        ttk.Button(btn,text='キャンセル',command=self.destroy).pack(side=tk.LEFT,padx=5)
        self.grab_set()
    def ok(self):
        try:
            self.result={'text':self.v_text.get().strip() or self.v_symbol.get().strip(),'type':self.v_type.get().strip(),'qty':float(self.v_qty.get() or 1),'unit':self.v_unit.get().strip(),'symbol':self.v_symbol.get().strip()}
            self.destroy()
        except Exception as e:
            messagebox.showwarning('入力エラー',str(e),parent=self)


# ---------------- Lighting rail overlap detection helpers ----------------
def is_lighting_rail_like_detection(det, image=None):
    """
    ライティングレールらしい候補かを判定。
    細長い横長/縦長の矩形を優先。画像特徴が取れる場合は暗線密度も見る。
    """
    try:
        w=max(1,int(det.w)); h=max(1,int(det.h))
        aspect=max(w/h, h/w)
        area=w*h
        if aspect >= 5.0 and area >= 120:
            return True
        if aspect >= 3.8 and area >= 250:
            return True
    except Exception:
        return False
    return False

def detection_center(det):
    return ((det.x1+det.x2)/2.0, (det.y1+det.y2)/2.0)

def detection_intersects_or_near(a, b, margin=10):
    try:
        return not (
            a.x2 + margin < b.x1 or
            b.x2 + margin < a.x1 or
            a.y2 + margin < b.y1 or
            b.y2 + margin < a.y1
        )
    except Exception:
        return False

def is_round_light_like_detection(det):
    try:
        w=max(1,int(det.w)); h=max(1,int(det.h))
        asp=w/max(1,h)
        return 0.55 <= asp <= 1.8 and 8 <= w <= 120 and 8 <= h <= 120
    except Exception:
        return False

# ---------------- Main GUI ----------------
class IntegratedApp:
    def __init__(self, root):
        self.root=root; self.root.title(APP_TITLE); self.root.geometry('1280x900')
        try: ttk.Style().theme_use('clam')
        except Exception: pass
        self.db=DatabaseManager(DB_PATH)
        self.result_rows=[]
        self.image_pages=[]; self.image_dets=[]; self.current_page=1
        self.selected_det_id=None; self.preview_tk=None; self.preview_scale=1.0
        self.source_image_file=''; self.pending_crop_path=''; self.pending_llm_answer=''
        self.annotation_busy=False
        self.last_dqn_state={}; self.last_dqn_action=""; self.last_dqn_mode=""; self.last_dqn_reward=0.0
        self.drag_start=None; self.drag_rect_id=None; self.dragging=False; self.drag_mode_enabled=True; self.manual_cable_lines=[]
        self.scale_px_per_mm=None; self.scale_line_ids=[]; self.scale_start_canvas=None
        self.torii_shins=[]          # ToriiShin リスト
        self.cad_symbols_snap=[]     # CADSymbol リスト
        self._snap_cursor_ids=[]     # スナップカーソル描画ID
        # DBから縮尺設定を復元
        try:
            saved=self.db.get_setting('scale_px_per_mm','')
            if saved: self.scale_px_per_mm=float(saved)
        except Exception: pass; self.manual_image_estimate_rows=[]; self.manual_symbol_items=[]; self.symbol_paste_mode=False
        self._suppress_tree_select=False; self._last_selected_det_id=None
        self._preview_rendering=False; self._preview_req=None
        self.chat_history=[]; self._chat_busy=False
        # ログキューはlog_q1本のみ。process_queuesは廃止 → root.after(0,fn)方式に統一
        self.log_q=queue.Queue()
        self.init_optional_ui_state(); self.build_ui(); self.refresh_all(); self.root.after(100, self._drain_log_q)

    # ------------------------------------------------------------------ #
    #  スレッドセーフUI更新ヘルパー                                        #
    #  ワーカースレッドから UI を触るための唯一の公式手段                   #
    #  → root.after(0, fn) = Tkinterメインループに安全にコールバックを渡す  #
    # ------------------------------------------------------------------ #
    def ui(self, fn, *args, **kwargs):
        """ワーカースレッドからメインスレッドのUIを安全に更新する。root.after(0,...)を使用。"""
        self.root.after(0, lambda: fn(*args, **kwargs))

    def log(self, msg):
        """ワーカースレッドからの安全なログ出力"""
        self.log_q.put(str(msg))

    def _drain_log_q(self):
        """100msごとにログキューを消化（ログだけキューを維持）"""
        try:
            while True:
                msg = self.log_q.get_nowait()
                try:
                    self.txt_log.insert(tk.END, msg+'\n'); self.txt_log.see(tk.END)
                    self.status.config(text=msg[-120:])
                except Exception: pass
        except queue.Empty: pass
        self.root.after(100, self._drain_log_q)

    def update_ai_usage_display(self, usage):
        usage = usage or {}
        msg = (
            f"AI使用量: provider={usage.get('provider','')} model={usage.get('model','')} "
            f"in={usage.get('input_tokens',0)} out={usage.get('output_tokens',0)} total={usage.get('total_tokens',0)} "
            f"概算=${usage.get('usd',0):.6f} / 約{usage.get('jpy',0):.2f}円"
        )
        self.log(msg)
        try:
            if hasattr(self, 'ai_usage_var'):
                self.ai_usage_var.set(msg)
        except Exception:
            pass
        try:
            if hasattr(self, 'img_ai_usage_var'):
                self.img_ai_usage_var.set(msg)
        except Exception:
            pass


    def snapshot_ai_settings(self):
        """メインスレッドでStringVar値をdictにコピー。ワーカースレッドに渡して使う。"""
        try:
            return {
                'provider': self.ai_provider_var.get().strip(),
                'ollama_url': self.ollama_url_var.get().strip(),
                'ollama_model': self.ollama_model_var.get().strip(),
                'openai_api_key': self.openai_api_key_var.get().strip(),
                'openai_model': self.openai_model_var.get().strip() or DEFAULT_OPENAI_MODEL,
                'anthropic_api_key': self.anthropic_api_key_var.get().strip(),
                'anthropic_model': self.anthropic_model_var.get().strip() or DEFAULT_CLAUDE_MODEL,
                'custom_openai_base_url': self.custom_openai_base_url_var.get().strip(),
                'custom_openai_api_key': self.custom_openai_api_key_var.get().strip(),
                'custom_openai_model': self.custom_openai_model_var.get().strip() or 'local-model',
                'use_litellm': '1' if getattr(self,'use_litellm_var',tk.BooleanVar(value=False)).get() else '0',
                'use_langmem': '1' if getattr(self,'use_langmem_var',tk.BooleanVar(value=True)).get() else '0',
                'memory_top_k': self.memory_top_k_var.get().strip() if hasattr(self,'memory_top_k_var') else '6',
            }
        except Exception:
            return {'provider': 'ollama',
                    'ollama_url': self.db.get_setting('ollama_url', DEFAULT_OLLAMA_URL),
                    'ollama_model': self.db.get_setting('ollama_model', DEFAULT_OLLAMA_MODEL)}

    def run_in_thread(self, worker_fn, done_fn=None, err_fn=None):
        """worker_fn()をバックグラウンドで実行し、結果をdone_fn(result)でUIに返す"""
        def _run():
            try:
                result = worker_fn()
                if done_fn: self.ui(done_fn, result)
            except Exception:
                tb = traceback.format_exc()
                if err_fn: self.ui(err_fn, tb)
                else: self.log('[ERROR] ' + tb)
        threading.Thread(target=_run, daemon=True).start()

    def apply_learning_to_detections(self, threshold=None, show_message=False):
        try:
            if threshold is None:
                try:
                    threshold=float(self.db.get_setting('learning_match_threshold','0.72'))
                except Exception:
                    threshold=0.62

            rows=self.db.get_annotation_features(limit=5000)
            dataset=self.db.get_symbol_image_dataset(limit=5000) if hasattr(self.db,'get_symbol_image_dataset') else []

            if (not rows and not dataset) or not self.image_pages:
                self.log('学習特徴量/記号画像データセットなし：通常推定のみ')
                if show_message:
                    messagebox.showinfo('学習反映','学習特徴量または記号画像データセットがありません。OK/NG登録または記号画像/PDF登録後に実行してください。')
                return 0

            applied=0
            best_overall=0.0

            for d in self.image_dets:
                try:
                    margin=min(8,max(0,min(d.w,d.h)//4))
                    crop=crop_symbol(self.image_pages[d.page-1],d,margin=margin)
                    if crop.width>120 or crop.height>120:
                        crop=crop.copy(); crop.thumbnail((96,96))
                    if crop.width<4 or crop.height<4:
                        continue  # 極小クロップはスキップ

                    label,score,row = best_learned_match_for_crop(crop,rows,threshold=threshold) if rows else (None,0.0,None)
                    label2,score2,row2 = best_symbol_dataset_match_for_crop(crop,dataset,threshold=threshold) if dataset else (None,0.0,None)

                    if score2 > score:
                        label,score,row = label2,score2,row2

                    # score2 も best_overall に反映（0.0 or 0 の落とし穴を避けるため明示的に比較）
                    best_overall=max(best_overall,
                                     float(score) if score else 0.0,
                                     float(score2) if score2 else 0.0)

                    if label:
                        # 手動で無効化された検出(memo='manual disabled')は上書きしない
                        if not d.enabled and 'manual' in str(getattr(d,'memo','')):
                            continue
                        old=d.equipment
                        d.equipment=label
                        d.enabled=True
                        d.source='learned'
                        d.memo=f'learned/dataset match {score:.2f} / old={old}'
                        d.score=max(float(getattr(d,'score',0) or 0),score)
                        applied+=1
                except Exception:
                    continue

            self.log(f'学習反映: {applied}件 / threshold={threshold:.2f} / best_sim={best_overall:.3f} / feat_rows={len(rows)}件 / dataset={len(dataset)}件')

            if applied:
                self.refresh_image_view()
                self.refresh_image_tables()

            if show_message:
                messagebox.showinfo('学習反映',f'学習反映: {applied}件\n閾値: {threshold:.2f}\n最高類似度: {best_overall:.2f}\n記号画像データセット: {len(dataset)}件')

            return applied
        except Exception:
            self.log(traceback.format_exc())
            if show_message:
                messagebox.showerror('学習反映エラー',traceback.format_exc()[-1600:])
            return 0


    def rebuild_learning_features(self):
        def worker():
            return self.db.rebuild_features_from_annotations()
        def done(res):
            added, skipped = res
            self.refresh_learning_summary()
            messagebox.showinfo('再構築完了', f'特徴量を再構築しました。\n追加: {added}件\nスキップ: {skipped}件\n\n注意: 反映には類似度閾値があります。現在の検出へ学習反映、または再度軽量解析を実行してください。')
        self.run_in_thread(worker, done)


    def local_guess_from_detection(self, d):
        """AI未接続時/ドラッグ範囲のローカル推定。学習DBと記号画像DBを最優先。"""
        if not d:
            return '未分類設備'
        try:
            if self.image_pages:
                crop = crop_symbol(self.image_pages[d.page-1], d, margin=8)


                # 0) FAISS/CLIPベクトル検索を最優先
                try:
                    label, score, top = self.vector_predict_detection(d, crop)
                    if label:
                        d.source='vector'
                        d.memo=f'vector local guess {score:.3f}'
                        d.score=max(float(getattr(d,'score',0) or 0), score)
                        return label
                except Exception:
                    pass

                # 1) OK/NG学習済み特徴量
                try:
                    rows = self.db.get_annotation_features(limit=5000)
                    label, score, row = best_learned_match_for_crop(crop, rows, threshold=0.72)
                    if label:
                        d.source='learned'
                        d.memo=f'learned local guess {score:.2f}'
                        d.score=max(float(getattr(d,'score',0) or 0), score)
                        return label
                except Exception:
                    pass

                # 2) 図面記号画像/PDFデータセット
                try:
                    if hasattr(self.db,'get_symbol_image_dataset'):
                        ds = self.db.get_symbol_image_dataset(limit=5000)
                        label, score, row = best_symbol_dataset_match_for_crop(crop, ds, threshold=0.72)
                        if label:
                            d.source='dataset'
                            d.memo=f'dataset local guess {score:.2f}'
                            d.score=max(float(getattr(d,'score',0) or 0), score)
                            return label
                except Exception:
                    pass

                # 3) テンプレート/色/形状
                guess, feat = simple_template_features(crop)
                cvfeat = opencv_shape_features(crop)
                if cvfeat and isinstance(cvfeat, dict):
                    circ = cvfeat.get('circularity',0)
                    asp = cvfeat.get('aspect',1)
                    if circ and circ > 0.55 and guess == '未分類設備':
                        guess = 'LEDダウンライト'
                    elif asp and asp > 2.2 and guess == '未分類設備':
                        guess = 'LEDベースライト'
                if guess != '未分類設備':
                    return guess
        except Exception:
            pass

        if getattr(d, 'equipment', '') and d.equipment != '未分類設備':
            return d.equipment

        cmap = {'purple':'LEDダウンライト','yellow':'LEDベースライト','green':'コンセント','cyan':'片切スイッチ','blue':'片切スイッチ','red':'分電盤'}
        return cmap.get(getattr(d, 'color_name', ''), '未分類設備')


    def memory_enabled(self):
        try:
            return bool(self.use_langmem_var.get())
        except Exception:
            return self.db.get_setting('use_langmem','1') == '1'

    def memory_context(self, query, namespace='estimation'):
        if not self.memory_enabled():
            return ''
        try:
            top_k=int(self.memory_top_k_var.get()) if hasattr(self,'memory_top_k_var') else int(self.db.get_setting('memory_top_k','6'))
        except Exception:
            top_k=6
        return LongTermMemoryManager(self.db).context(query, namespace=namespace, top_k=top_k)

    def record_memory(self, namespace, kind, text_value, tags='', source='', weight=1.0, meta=None):
        if not self.memory_enabled():
            return
        try:
            LongTermMemoryManager(self.db).add(namespace, kind, text_value, tags, source, weight, meta)
        except Exception:
            self.log(traceback.format_exc())

    def safe_save_annotation_only(self, det_id, llm_answer, final_answer, correct, crop_path='', memo=''):
        """
        OK/NG登録専用。
        学習登録時はAIへ再問い合わせせず、SQLite保存だけ行う。
        """
        d = self.get_image_det(det_id)
        if not d:
            messagebox.showwarning('未選択','図記号が選択されていません')
            return False
        final_answer = (final_answer or '').strip() or self.local_guess_from_detection(d)
        llm_answer = (llm_answer or '').strip() or final_answer
        d.equipment = final_answer
        d.enabled = True
        d.memo = memo or ('OK' if correct else 'NG correction')
        try:
            # 1) 通常のアノテーション履歴を保存
            self.db.save_annotation(self.source_image_file, d, llm_answer, final_answer, bool(correct), crop_path, d.memo)

            # 2) 再学習・再推論用の特徴量も保存
            if self.image_pages:
                crop = crop_symbol(self.image_pages[d.page-1], d, margin=8)
                if not crop_path:
                    crop_path = str(CROP_DIR / f'learn_{datetime.now().strftime("%Y%m%d_%H%M%S")}_id{d.det_id}.png')
                    crop.save(crop_path)
                feat = extract_symbol_feature(crop)
                self.db.save_annotation_feature(self.source_image_file, d, final_answer, crop_path, json.dumps(feat, ensure_ascii=False), d.memo)
                self.record_memory(
                    'symbols',
                    'annotation_correction',
                    f'図記号 correction: source={Path(self.source_image_file).name}, color={d.color_name}, bbox=({d.x1},{d.y1})-({d.x2},{d.y2}), llm={llm_answer}, final={final_answer}, correct={bool(correct)}',
                    tags=f'{final_answer},{d.color_name},annotation,symbol',
                    source=self.source_image_file,
                    weight=2.0 if not correct else 1.2,
                    meta={'det_id':d.det_id,'page':d.page,'crop_path':crop_path}
                )
        except Exception as e:
            messagebox.showerror('DB保存エラー', str(e))
            return False
        try:
            self.refresh_image_view()
            self.refresh_image_tables()
            self.refresh_learning_summary()
        except Exception:
            self.log(traceback.format_exc())
        # DQN user feedback reward
        try:
            action=getattr(self,'last_dqn_action','') or self.db.get_setting('dqn_last_action','')
            if action:
                reward=1.5 if correct else -1.5
                self.dqn_agent.update(getattr(self,'last_dqn_state',{}),action,reward,None,lr=float(self.db.get_setting('dqn_learning_rate','0.08')),gamma=float(self.db.get_setting('dqn_gamma','0.90')))
                self.db.add_dqn_reward(action,reward,f'user annotation feedback final={final_answer}')
                self.last_dqn_reward=reward
                if hasattr(self,'dqn_status_var'): self.dqn_status_var.set(f'DQN feedback: {action} reward={reward}')
        except Exception: self.log('DQN user feedback failed:\n'+traceback.format_exc())
        return True

    def init_optional_ui_state(self):
        """
        後付け拡張で使うTkinter変数・状態変数を安全に初期化する。
        Pydroid3では未初期化AttributeErrorで即終了するため、build_ui前に必ず呼ぶ。
        """
        try:
            if not hasattr(self, 'snap_enabled_var'):
                self.snap_enabled_var = tk.BooleanVar(value=True)
            if not hasattr(self, 'snap_grid_px_var'):
                self.snap_grid_px_var = tk.StringVar(value='20')
            if not hasattr(self, 'cable_draw_mode_var'):
                self.cable_draw_mode_var = tk.BooleanVar(value=False)
            if not hasattr(self, 'symbol_paste_mode_var'):
                self.symbol_paste_mode_var = tk.BooleanVar(value=False)
            if not hasattr(self, 'symbol_paste_var'):
                self.symbol_paste_var = tk.StringVar(value='LEDダウンライト')
            if not hasattr(self, 'manual_cable_lines'):
                self.manual_cable_lines = []
            if not hasattr(self, 'manual_image_estimate_rows'):
                self.manual_image_estimate_rows = []
            if not hasattr(self, 'manual_symbol_items'):
                self.manual_symbol_items = []
            if not hasattr(self, 'drag_start'):
                self.drag_start = None
            if not hasattr(self, 'drag_rect_id'):
                self.drag_rect_id = None
            if not hasattr(self, 'dragging'):
                self.dragging = False
            if not hasattr(self, 'symbol_paste_mode'):
                self.symbol_paste_mode = False
            if not hasattr(self, 'last_dqn_state'):
                self.last_dqn_state = {}
            if not hasattr(self, 'last_dqn_action'):
                self.last_dqn_action = ''
            if not hasattr(self, 'last_dqn_mode'):
                self.last_dqn_mode = ''
            if not hasattr(self, 'last_dqn_reward'):
                self.last_dqn_reward = 0.0
            if not hasattr(self, 'current_mindmap'):
                self.current_mindmap = default_mindmap_json('電気設備工事',1)
            if not hasattr(self, 'vnodes'):
                self.vnodes=[]
            if not hasattr(self, 'vedges'):
                self.vedges=[]
            if not hasattr(self, 'node_connect_mode'):
                self.node_connect_mode=False
            if not hasattr(self, 'node_selected'):
                self.node_selected=None
            if not hasattr(self, 'node_id_seq'):
                self.node_id_seq=1
        except Exception:
            # root生成前ではなくIntegratedApp生成後に呼ぶ前提なので基本ここには来ない
            pass

    def build_ui(self):
        menubar=tk.Menu(self.root); self.root.config(menu=menubar); fm=tk.Menu(menubar,tearoff=0); menubar.add_cascade(label='ファイル',menu=fm); fm.add_command(label='積算結果CSV出力',command=self.export_result_csv); fm.add_separator(); fm.add_command(label='終了',command=self.root.quit)
        tm=tk.Menu(menubar,tearoff=0); menubar.add_cascade(label='ツール',menu=tm)
        for pkg in ['requests','pypdf','PyMuPDF','ezdxf','chardet','pillow','opencv-python','pytesseract','ultralytics','litellm','langmem']: tm.add_command(label=f'{pkg} インストール',command=lambda p=pkg:self.install_pkg_async(p))
        ttk.Label(self.root,text=APP_TITLE,font=('Arial',14,'bold')).pack(pady=5)
        self.nb=ttk.Notebook(self.root); self.nb.pack(fill=tk.BOTH,expand=True,padx=6,pady=4)
        self.tab_est=ttk.Frame(self.nb); self.tab_price=ttk.Frame(self.nb); self.tab_sym=ttk.Frame(self.nb); self.tab_cad=ttk.Frame(self.nb); self.tab_img=ttk.Frame(self.nb); self.tab_tune=ttk.Frame(self.nb); self.tab_ai=ttk.Frame(self.nb); self.tab_chat=ttk.Frame(self.nb); self.tab_mindmap=ttk.Frame(self.nb); self.tab_nodecad=ttk.Frame(self.nb); self.tab_dqn=ttk.Frame(self.nb); self.tab_learn=ttk.Frame(self.nb)
        for tab,name in [(self.tab_est,'積算実行 PDF/DXF'),(self.tab_price,'単価マスター'),(self.tab_sym,'記号パターン'),(self.tab_cad,'CADライブラリ'),(self.tab_img,'画像PDF解析・アノテーション'),(self.tab_tune,'画像チューニング'),(self.tab_ai,'Ollama解析'),(self.tab_chat,'AIチャット'),(self.tab_mindmap,'ToDoマインドマップAI'),(self.tab_nodecad,'ノード図面作成'),(self.tab_dqn,'AI解析戦略(DQN)'),(self.tab_learn,'学習データ')]: self.nb.add(tab,text=name)
        self.build_est_tab(); self.build_price_tab(); self.build_sym_tab(); self.build_cad_tab(); self.build_image_tab(); self.build_tune_tab(); self.build_ai_tab(); self.build_chat_tab(); self.build_mindmap_tab(); self.build_nodecad_tab(); self.build_dqn_tab(); self.build_learning_tab()
        self.status=ttk.Label(self.root,text='準備完了',relief=tk.SUNKEN,anchor='w'); self.status.pack(fill=tk.X,side=tk.BOTTOM)



    def build_est_tab(self):
        f=self.tab_est; top=ttk.LabelFrame(f,text='ファイル選択',padding=8); top.pack(fill=tk.X,padx=6,pady=6); self.file_path_var=tk.StringVar(); row=ttk.Frame(top); row.pack(fill=tk.X); ttk.Label(row,text='PDF/DXF:').pack(side=tk.LEFT); e=ttk.Entry(row,textvariable=self.file_path_var); e.pack(side=tk.LEFT,fill=tk.X,expand=True,padx=4); add_paste_button(row,e); ttk.Button(row,text='選択',command=self.select_file).pack(side=tk.LEFT,padx=3); ttk.Button(row,text='積算実行',command=self.on_click_estimate).pack(side=tk.LEFT,padx=3); ttk.Button(row,text='AI積算実行',command=self.on_click_ai_estimate).pack(side=tk.LEFT,padx=3); ttk.Button(row,text='PDFをDXF分離保存(簡易)',command=self.on_click_pdf_split_to_dxf).pack(side=tk.LEFT,padx=3)
        self.progress=ttk.Progressbar(f,mode='indeterminate'); self.progress.pack(fill=tk.X,padx=6,pady=3); paned=ttk.PanedWindow(f,orient=tk.VERTICAL); paned.pack(fill=tk.BOTH,expand=True,padx=6,pady=5); rf=ttk.LabelFrame(paned,text='積算結果',padding=5); paned.add(rf,weight=3); cols=('カテゴリ','品名','仕様','単位','数量','単価(円)','金額(円)'); self.tree=ttk.Treeview(rf,columns=cols,show='headings',height=10)
        for c in cols: self.tree.heading(c,text=c); self.tree.column(c,width=130,anchor=tk.E if c in ('数量','単価(円)','金額(円)') else tk.W)
        self.tree.pack(side=tk.LEFT,fill=tk.BOTH,expand=True); ttk.Scrollbar(rf,orient=tk.VERTICAL,command=self.tree.yview).pack(side=tk.RIGHT,fill=tk.Y); lf=ttk.LabelFrame(paned,text='ログ',padding=5); paned.add(lf,weight=2); self.txt_log=scrolledtext.ScrolledText(lf,height=8); self.txt_log.pack(fill=tk.BOTH,expand=True); af=ttk.LabelFrame(paned,text='AI積算実行結果',padding=5); paned.add(af,weight=2); self.ai_estimate_text=scrolledtext.ScrolledText(af,height=8); self.ai_estimate_text.pack(fill=tk.BOTH,expand=True)
    def build_price_tab(self):
        f=self.tab_price; btn=ttk.Frame(f); btn.pack(fill=tk.X,padx=6,pady=6)
        for text,cmd in [('追加',self.price_add),('編集',self.price_edit),('削除',self.price_delete),('CSVインポート',self.price_import_csv),('CSVエクスポート',self.price_export_csv),('更新',self.price_refresh)]: ttk.Button(btn,text=text,command=cmd).pack(side=tk.LEFT,padx=3)
        cols=('ID','カテゴリ','品名','仕様','単位','単価(円)','キーワード'); self.price_tree=ttk.Treeview(f,columns=cols,show='headings')
        for c,w in zip(cols,[50,120,160,180,60,100,280]): self.price_tree.heading(c,text=c); self.price_tree.column(c,width=w,anchor=tk.E if c in ('ID','単価(円)') else tk.W)
        self.price_tree.pack(fill=tk.BOTH,expand=True,padx=6,pady=4); self.price_tree.bind('<Double-1>',lambda e:self.price_edit())
    def build_sym_tab(self):
        f=self.tab_sym; ttk.Label(f,text='DXF図面内の円/線/矩形パターンで設備記号を推定します。',padding=6).pack(anchor='w'); btn=ttk.Frame(f); btn.pack(fill=tk.X,padx=6,pady=4)
        for text,cmd in [('追加',self.symbol_add),('編集',self.symbol_edit),('削除',self.symbol_delete),('更新',self.symbol_refresh)]: ttk.Button(btn,text=text,command=cmd).pack(side=tk.LEFT,padx=3)
        cols=('ID','設備名','パターンタイプ','説明','プリセット'); self.sym_tree=ttk.Treeview(f,columns=cols,show='headings')
        for c,w in zip(cols,[50,180,180,360,80]): self.sym_tree.heading(c,text=c); self.sym_tree.column(c,width=w)
        self.sym_tree.pack(fill=tk.BOTH,expand=True,padx=6,pady=4); self.sym_tree.bind('<Double-1>',lambda e:self.symbol_edit())
    def build_cad_tab(self):
        f=self.tab_cad; dl=ttk.LabelFrame(f,text='CADデータダウンロード/登録',padding=8); dl.pack(fill=tk.X,padx=6,pady=6); self.cad_url_var=tk.StringVar(value='https://'); row=ttk.Frame(dl); row.pack(fill=tk.X); ttk.Entry(row,textvariable=self.cad_url_var).pack(side=tk.LEFT,fill=tk.X,expand=True); ttk.Button(row,text='URL/ZIP取込',command=self.cad_download_from_url).pack(side=tk.LEFT,padx=3); ttk.Button(row,text='ローカルDXF取込',command=self.cad_import_local).pack(side=tk.LEFT,padx=3); ttk.Button(row,text='フォルダを開く',command=self.cad_open_folder).pack(side=tk.LEFT,padx=3)
        cols=('ID','メーカー','型番','カテゴリ','機器名','ファイル名'); self.cad_tree=ttk.Treeview(f,columns=cols,show='headings')
        for c,w in zip(cols,[50,130,130,110,160,420]): self.cad_tree.heading(c,text=c); self.cad_tree.column(c,width=w)
        self.cad_tree.pack(fill=tk.BOTH,expand=True,padx=6,pady=4); b=ttk.Frame(f); b.pack(fill=tk.X,padx=6,pady=4); ttk.Button(b,text='削除',command=self.cad_delete).pack(side=tk.LEFT); ttk.Button(b,text='更新',command=self.cad_refresh).pack(side=tk.RIGHT)
    def show_image_command_menu(self, event=None):
        m=tk.Menu(self.root, tearoff=0)
        m.add_command(label='PDF/画像/DXFを選択', command=self.select_image_file)
        m.add_command(label='軽量解析', command=self.start_image_analysis)
        m.add_separator()
        m.add_command(label='図面領域最大化', command=self.maximize_drawing_area)
        m.add_command(label='右ペイン復帰', command=self.restore_drawing_area)
        m.add_separator()
        m.add_command(label='検出編集', command=self.edit_selected_detection)
        m.add_command(label='検出削除', command=self.delete_selected_detection)
        m.add_command(label='無効候補を削除', command=self.delete_disabled_detections)
        m.add_separator()
        m.add_command(label='画像積算CSV', command=self.export_image_csv)
        if hasattr(self,'import_symbol_crops_dataset'):
            m.add_command(label='symbol_crops取込', command=lambda:self.import_symbol_crops_dataset(True))
        m.add_command(label='記号画像/PDF登録', command=self.register_symbol_image_dataset)
        try:
            if event:
                m.tk_popup(event.x_root,event.y_root)
            else:
                m.tk_popup(self.root.winfo_pointerx(), self.root.winfo_pointery())
        finally:
            try: m.grab_release()
            except Exception: pass

    def build_office_ribbon(self, parent, context='image'):
        """
        Office風の小型アイコンリボン。
        ボタンを横に大量配置せず、各アイコンからプルダウンメニューを出す。
        """
        ribbon=ttk.Frame(parent)
        ribbon.pack(fill=tk.X,padx=2,pady=1)

        def rb(text, cmd, tip=''):
            b=ttk.Button(ribbon,text=text,width=4,command=cmd)
            b.pack(side=tk.LEFT,padx=1)
            return b

        rb('📂', self.show_ribbon_file_menu)
        rb('🔍', self.show_ribbon_analysis_menu)
        rb('✏', self.show_ribbon_annotation_menu)
        rb('🔌', self.show_ribbon_cable_menu)
        rb('⚡', self.show_ribbon_symbol_menu)
        rb('🤖', self.show_ribbon_ai_menu)
        rb('🧠', self.show_ribbon_learning_menu)
        rb('🗄', self.show_ribbon_db_menu)

        self.ribbon_mode_label=tk.StringVar(value='通常: ドラッグ=範囲指定')
        ttk.Label(ribbon,textvariable=self.ribbon_mode_label,foreground='blue').pack(side=tk.LEFT,padx=8)
        self.ribbon_symbol_label=tk.StringVar(value='貼付: '+self.symbol_paste_var.get())
        ttk.Label(ribbon,textvariable=self.ribbon_symbol_label,foreground='#555').pack(side=tk.LEFT,padx=4)

        # ファイルパスは薄く小さく表示
        if not hasattr(self,'image_file_var'):
            self.image_file_var=tk.StringVar()
        ttk.Entry(ribbon,textvariable=self.image_file_var,width=28).pack(side=tk.RIGHT,padx=2)
        return ribbon

    def popup_menu(self, items):
        m=tk.Menu(self.root,tearoff=0)
        for item in items:
            if item is None:
                m.add_separator()
            else:
                label,cmd=item
                m.add_command(label=label,command=cmd)
        m.tk_popup(self.root.winfo_pointerx(),self.root.winfo_pointery())

    def show_ribbon_file_menu(self):
        self.popup_menu([
            ('PDF/画像/DXFを選択', self.select_image_file),
            ('軽量解析', self.start_image_analysis),
            ('画像積算CSV出力', self.export_image_csv),
            None,
            ('図面領域最大化', self.maximize_drawing_area),
            ('右ペイン復帰', self.restore_drawing_area),
        ])

    def show_ribbon_analysis_menu(self):
        self.popup_menu([
            ('軽量解析', self.start_image_analysis),
            ('ライティングレール再検出', lambda:self.apply_lighting_rail_overlap_fix(True)),
            ('FAISS/CLIPベクトル検索を反映', lambda:self.apply_vector_search_to_detections(True) if hasattr(self,'apply_vector_search_to_detections') else None),
            ('学習を再反映', lambda:self.apply_learning_to_detections(show_message=True)),
            None,
            ('symbol_crops取込', lambda:self.import_symbol_crops_dataset(True)),
            ('FAISS/CLIP取込', lambda:self.import_symbol_crops_to_vector_db(True) if hasattr(self,'import_symbol_crops_to_vector_db') else None),
        ])

    def show_ribbon_annotation_menu(self):
        self.popup_menu([
            ('検出候補を編集', self.edit_selected_detection),
            ('検出候補を削除', self.delete_selected_detection),
            ('無効候補を削除', self.delete_disabled_detections),
            ('候補を手動追加', self.add_manual_detection),
            None,
            ('AI判定', self.annotate_selected_async),
            ('OK/NG学習登録', self.save_annotation_feedback),
        ])

    def show_ribbon_cable_menu(self):
        def cable_on():
            self.cable_draw_mode_var.set(True)
            self.symbol_paste_mode_var.set(False)
            self.ribbon_mode_label.set('ケーブル描画ON: ドラッグで配線')
        def cable_off():
            self.cable_draw_mode_var.set(False)
            self.ribbon_mode_label.set('通常: ドラッグ=範囲指定')
        def snap_on():
            self.snap_enabled_var.set(True)
        def snap_off():
            self.snap_enabled_var.set(False)
        self.popup_menu([
            ('ケーブル描画ON', cable_on),
            ('ケーブル描画OFF', cable_off),
            ('スナップON', snap_on),
            ('スナップOFF', snap_off),
            None,
            ('ケーブル編集', self.edit_selected_cable),
            ('ケーブル削除', self.delete_selected_cable),
        ])

    def show_ribbon_symbol_menu(self):
        def paste_on():
            self.symbol_paste_mode_var.set(True)
            self.cable_draw_mode_var.set(False)
            self.ribbon_mode_label.set('記号貼付ON: 図面クリックで配置')
        def paste_off():
            self.symbol_paste_mode_var.set(False)
            self.ribbon_mode_label.set('通常: ドラッグ=範囲指定')
        self.popup_menu([
            ('記号貼付ON', paste_on),
            ('記号貼付OFF', paste_off),
            None,
            ('貼付記号: LEDダウンライト', lambda:(self.symbol_paste_var.set('LEDダウンライト'), hasattr(self,'ribbon_symbol_label') and self.ribbon_symbol_label.set('貼付: LEDダウンライト'))),
            ('貼付記号: ライティングレール', lambda:(self.symbol_paste_var.set('ライティングレール'), hasattr(self,'ribbon_symbol_label') and self.ribbon_symbol_label.set('貼付: ライティングレール'))),
            ('貼付記号: コンセント', lambda:(self.symbol_paste_var.set('コンセント'), hasattr(self,'ribbon_symbol_label') and self.ribbon_symbol_label.set('貼付: コンセント'))),
            ('貼付記号: スイッチ', lambda:(self.symbol_paste_var.set('片切スイッチ'), hasattr(self,'ribbon_symbol_label') and self.ribbon_symbol_label.set('貼付: 片切スイッチ'))),
            ('貼付記号: 分電盤', lambda:(self.symbol_paste_var.set('分電盤'), hasattr(self,'ribbon_symbol_label') and self.ribbon_symbol_label.set('貼付: 分電盤'))),
            None,
            ('記号画像/PDF登録', self.register_symbol_image_dataset),
        ])

    def show_ribbon_ai_menu(self):
        self.popup_menu([
            ('選択AI接続テスト', self.test_selected_ai_async),
            ('AI判定', self.annotate_selected_async),
            ('AI積算実行', self.on_click_ai_estimate if hasattr(self,'on_click_ai_estimate') else self.start_image_analysis),
        ])

    def show_ribbon_learning_menu(self):
        self.popup_menu([
            ('学習データ更新', self.refresh_learning_summary),
            ('特徴量再構築', self.rebuild_learning_features),
            ('学習を現在候補へ反映', lambda:self.apply_learning_to_detections(show_message=True)),
            ('ベクトル索引再構築', self.rebuild_vector_index_dialog if hasattr(self,'rebuild_vector_index_dialog') else self.refresh_learning_summary),
        ])

    def show_ribbon_db_menu(self):
        self.popup_menu([
            ('rqlite接続テスト', self.test_rqlite_connection if hasattr(self,'test_rqlite_connection') else self.refresh_all),
            ('SQLite→rqlite移行', self.migrate_sqlite_to_rqlite_dialog if hasattr(self,'migrate_sqlite_to_rqlite_dialog') else self.refresh_all),
            ('DB設定保存', self.save_db_backend_settings if hasattr(self,'save_db_backend_settings') else self.refresh_all),
        ])

    def build_image_tab(self):
        f=self.tab_img; self.build_office_ribbon(f, context='image')
        main=ttk.PanedWindow(f,orient=tk.HORIZONTAL); self.image_main_pane=main; main.pack(fill=tk.BOTH,expand=True,padx=3,pady=2); left=ttk.Frame(main); main.add(left,weight=7); bar=ttk.Frame(left); bar.pack(fill=tk.X); ttk.Button(bar,text='前ページ',command=lambda:self.change_image_page(-1)).pack(side=tk.LEFT); ttk.Button(bar,text='次ページ',command=lambda:self.change_image_page(1)).pack(side=tk.LEFT,padx=3); self.page_label=tk.StringVar(value='page -/-'); ttk.Label(bar,textvariable=self.page_label).pack(side=tk.LEFT,padx=8); ttk.Label(bar,text='ドラッグ=範囲 / ケーブルON=配線 / 記号貼付ON=配置').pack(side=tk.RIGHT)
        # ──── 縮尺設定ツールバー（ページナビの直下・キャンバスの上） ────
        scalebar=ttk.LabelFrame(left,text='縮尺測定',padding=3)
        scalebar.pack(fill=tk.X,pady=2)
        self.scale_mode_var=tk.BooleanVar(value=False)
        ttk.Checkbutton(scalebar,text='縮尺測定ON',variable=self.scale_mode_var).pack(side=tk.LEFT,padx=4)
        ttk.Label(scalebar,text='← チェックしてドラッグ → mm入力 → ケーブル長を自動換算').pack(side=tk.LEFT,padx=2)
        self.scale_info_var=tk.StringVar(value='縮尺: 未設定')
        ttk.Label(scalebar,textvariable=self.scale_info_var,font=('Arial',9,'bold')).pack(side=tk.LEFT,padx=8)
        ttk.Button(scalebar,text='縮尺クリア',command=self.clear_scale_calibration).pack(side=tk.LEFT,padx=2)
        # ──── AutoCADスナップ + 通り芯ツールバー ────
        self.build_cad_snap_toolbar(left)
        # ─────────────────────────────────────────────────────────────
        preview_frame=ttk.Frame(left)
        preview_frame.pack(fill=tk.BOTH,expand=True)
        self.preview_canvas=tk.Canvas(preview_frame,bg='white',width=1120,height=780)
        self.preview_vbar=ttk.Scrollbar(preview_frame,orient=tk.VERTICAL,command=self.preview_canvas.yview)
        self.preview_hbar=ttk.Scrollbar(preview_frame,orient=tk.HORIZONTAL,command=self.preview_canvas.xview)
        self.preview_canvas.configure(yscrollcommand=self.preview_vbar.set,xscrollcommand=self.preview_hbar.set)
        self.preview_canvas.grid(row=0,column=0,sticky='nsew')
        self.preview_vbar.grid(row=0,column=1,sticky='ns')
        self.preview_hbar.grid(row=1,column=0,sticky='ew')
        preview_frame.rowconfigure(0,weight=1)
        preview_frame.columnconfigure(0,weight=1)
        self.preview_canvas.bind('<Button-1>',self.on_preview_click)
        self.preview_canvas.bind('<B1-Motion>',self.on_preview_drag)
        self.preview_canvas.bind('<ButtonRelease-1>',self.on_preview_drag_release)
        self.preview_canvas.bind('<Double-Button-1>',self.on_preview_double)
        self.preview_canvas.bind('<Button-3>',self.on_preview_right_click)
        self.preview_canvas.bind('<Button-2>',self.on_preview_right_click)
        self.preview_canvas.bind('<MouseWheel>',self.on_preview_mousewheel)
        self.preview_canvas.bind('<Shift-MouseWheel>',self.on_preview_shift_mousewheel)
        right=ttk.PanedWindow(main,orient=tk.VERTICAL); self.image_right_pane=right; main.add(right,weight=1); ann=ttk.LabelFrame(right,text='アノテーション登録・学習',padding=6); right.add(ann,weight=0); self.selected_info_var=tk.StringVar(value='未選択'); ttk.Label(ann,textvariable=self.selected_info_var).grid(row=0,column=0,columnspan=4,sticky='w'); ttk.Label(ann,text='LLM回答').grid(row=1,column=0,sticky='e'); self.llm_answer_var=tk.StringVar(); ttk.Entry(ann,textvariable=self.llm_answer_var,width=28).grid(row=1,column=1,sticky='w'); ttk.Button(ann,text='OK登録',command=self.register_llm_ok).grid(row=1,column=2,padx=3); ttk.Label(ann,text='訂正').grid(row=2,column=0,sticky='e'); self.correct_answer_var=tk.StringVar(); ttk.Combobox(ann,textvariable=self.correct_answer_var,values=['LEDダウンライト','LEDベースライト','コンセント','片切スイッチ','分電盤','換気扇','未分類設備'],width=26).grid(row=2,column=1,sticky='w'); ttk.Button(ann,text='訂正登録・学習',command=self.register_manual_correction).grid(row=2,column=2,padx=3); ttk.Button(ann,text='選択を無効化',command=self.disable_image_selected).grid(row=1,column=3,padx=3); ttk.Button(ann,text='設備名だけ変更',command=self.apply_correct_to_selected).grid(row=2,column=3,padx=3); self.drag_mode_var=tk.BooleanVar(value=True)
        ttk.Checkbutton(ann,text='左ドラッグで範囲指定',variable=self.drag_mode_var).grid(row=3,column=0,sticky='w',padx=3,pady=3)
        ttk.Button(ann,text='AI判定実行',command=self.annotate_selected_popup_async).grid(row=3,column=1,padx=3,pady=3)
        ttk.Label(ann,text='※ドラッグ範囲は即ローカル推定。必要時だけAI判定実行。').grid(row=3,column=2,columnspan=2,sticky='w')
        self.img_ai_status_var=tk.StringVar(value='AI状態: 未実行')
        ttk.Label(ann,textvariable=self.img_ai_status_var,foreground='blue').grid(row=4,column=0,columnspan=4,sticky='w',pady=2)
        self.img_ai_usage_var=tk.StringVar(value='AI使用量: 未使用')
        ttk.Label(ann,textvariable=self.img_ai_usage_var,foreground='purple').grid(row=5,column=0,columnspan=4,sticky='w',pady=2)
        detf=ttk.LabelFrame(right,text='画像検出候補',padding=4); right.add(detf,weight=3); cols=('on','id','page','equipment','color','x','y','w','h','score'); self.img_det_tree=ttk.Treeview(detf,columns=cols,show='headings',height=12)
        for c in cols: self.img_det_tree.heading(c,text=c); self.img_det_tree.column(c,width=70,anchor=tk.CENTER)
        self.img_det_tree.column('equipment',width=130); self.img_det_tree.pack(fill=tk.BOTH,expand=True); self.img_det_tree.bind('<<TreeviewSelect>>',self.on_img_tree_select); self.img_det_tree.bind('<Double-1>',lambda e:self.edit_selected_detection()); self.img_det_tree.bind('<Delete>',lambda e:self.delete_selected_detection()); self.img_det_tree.bind('<Button-3>',self.show_detection_context_menu)
        detbtn=ttk.Frame(detf); detbtn.pack(fill=tk.X,pady=3)
        ttk.Button(detbtn,text='選択候補を編集',command=self.edit_selected_detection).pack(side=tk.LEFT,padx=2)
        ttk.Button(detbtn,text='選択候補を削除',command=self.delete_selected_detection).pack(side=tk.LEFT,padx=2)
        ttk.Button(detbtn,text='無効候補を削除',command=self.delete_disabled_detections).pack(side=tk.LEFT,padx=2)
        ttk.Button(detbtn,text='候補を手動追加',command=self.add_manual_detection).pack(side=tk.LEFT,padx=2)
        ttk.Button(detbtn,text='学習を再反映',command=lambda:self.apply_learning_to_detections(show_message=True)).pack(side=tk.LEFT,padx=2)
        sumf=ttk.LabelFrame(right,text='画像積算結果',padding=4); right.add(sumf,weight=2); cols2=('カテゴリ','品名','単位','数量','単価','金額'); self.img_sum_tree=ttk.Treeview(sumf,columns=cols2,show='headings',height=7)
        for c in cols2: self.img_sum_tree.heading(c,text=c); self.img_sum_tree.column(c,width=90,anchor=tk.E if c in ('数量','単価','金額') else tk.W)
        self.img_sum_tree.pack(fill=tk.BOTH,expand=True)
        self.manual_image_estimate_rows=[]
        self.img_sum_tree.bind('<Double-1>',lambda e:self.edit_selected_image_sum_row()); self.img_sum_tree.bind('<Delete>',lambda e:self.delete_selected_image_sum_row()); self.img_sum_tree.bind('<Button-3>',self.show_sum_context_menu)
        sumbtn=ttk.Frame(sumf); sumbtn.pack(fill=tk.X,pady=3)
        ttk.Button(sumbtn,text='積算行を編集',command=self.edit_selected_image_sum_row).pack(side=tk.LEFT,padx=2)
        ttk.Button(sumbtn,text='積算行を追加',command=self.add_image_sum_row).pack(side=tk.LEFT,padx=2)
        ttk.Button(sumbtn,text='積算行を削除',command=self.delete_selected_image_sum_row).pack(side=tk.LEFT,padx=2)
        ttk.Button(sumbtn,text='検出候補から再集計',command=self.refresh_image_tables).pack(side=tk.LEFT,padx=2)
        ttk.Button(sumbtn,text='AI積算（画像解析）',command=self.run_image_ai_estimate_async).pack(side=tk.LEFT,padx=6)
        cablef=ttk.LabelFrame(right,text='手書きケーブル一覧',padding=4); right.add(cablef,weight=1)
        ccols=('No','page','種別','長さm','x1','y1','x2','y2')
        self.cable_tree=ttk.Treeview(cablef,columns=ccols,show='headings',height=5)
        for c in ccols: self.cable_tree.heading(c,text=c); self.cable_tree.column(c,width=60,anchor=tk.CENTER)
        self.cable_tree.column('種別',width=110)
        self.cable_tree.pack(fill=tk.BOTH,expand=True)
        self.cable_tree.bind('<Double-1>',lambda e:self.edit_selected_cable())
        self.cable_tree.bind('<Delete>',lambda e:self.delete_selected_cable())
        cb=ttk.Frame(cablef); cb.pack(fill=tk.X,pady=3)
        ttk.Button(cb,text='ケーブル編集',command=self.edit_selected_cable).pack(side=tk.LEFT,padx=2)
        ttk.Button(cb,text='ケーブル削除',command=self.delete_selected_cable).pack(side=tk.LEFT,padx=2)
    def maximize_drawing_area(self):
        """画像PDF解析タブの右側候補/積算ペインを隠して図面表示領域を広げる。"""
        try:
            if hasattr(self, 'image_main_pane') and hasattr(self, 'image_right_pane'):
                try:
                    self.image_main_pane.forget(self.image_right_pane)
                except Exception:
                    # forget非対応環境でも落とさない
                    pass
            self.status.config(text='図面領域を最大化しました')
        except Exception:
            messagebox.showwarning('最大化エラー', traceback.format_exc()[-1000:])

    def restore_drawing_area(self):
        """右側の候補/積算/ケーブル一覧ペインを復帰する。"""
        try:
            if hasattr(self, 'image_main_pane') and hasattr(self, 'image_right_pane'):
                try:
                    self.image_main_pane.add(self.image_right_pane, weight=1)
                except Exception:
                    pass
            self.status.config(text='右ペインを復帰しました')
        except Exception:
            messagebox.showwarning('復帰エラー', traceback.format_exc()[-1000:])

    def build_tune_tab(self):
        f=self.tab_tune; sf=ttk.LabelFrame(f,text='画像解析 軽量化/精度設定',padding=8); sf.pack(fill=tk.X,padx=6,pady=6); self.setting_vars={}; items=[('pdf_scale','PDF画像化倍率'),('analysis_max_width','解析最大幅'),('preview_max_width','プレビュー最大幅'),('scan_step','走査間隔 1高精度/4軽量'),('cluster_distance','結合距離'),('min_cluster_points','最小点数'),('min_box_w','最小幅'),('min_box_h','最小高'),('max_box_w','最大幅'),('max_box_h','最大高'),('exclude_legend','凡例除外 1/0'),('legend_right_ratio','凡例右比率'),('legend_bottom_ratio','凡例下比率'),('max_detections','最大候補数'),('analysis_engine','解析エンジン color/template/opencv/yolo/sam/ocr'),('template_enabled','テンプレート 1/0'),('opencv_enabled','OpenCV 1/0'),('ocr_enabled','OCR 1/0'),('yolo_enabled','YOLO 1/0'),('sam_enabled','SAM2 1/0'),('cnn_vit_enabled','CNN/ViT 1/0'),('learning_match_threshold','学習反映しきい値 0.50-0.90')]
        for i,(k,lab) in enumerate(items): ttk.Label(sf,text=lab).grid(row=i//2,column=(i%2)*2,sticky='e',padx=4,pady=2); v=tk.StringVar(value=self.db.get_setting(k,'')); ttk.Entry(sf,textvariable=v,width=12).grid(row=i//2,column=(i%2)*2+1,sticky='w'); self.setting_vars[k]=v
        ttk.Button(sf,text='保存',command=self.save_settings).grid(row=8,column=0,pady=8); ttk.Button(sf,text='保存して再解析',command=self.save_settings_and_reanalyze).grid(row=8,column=1,pady=8)
        rf=ttk.LabelFrame(f,text='色ルール',padding=6); rf.pack(fill=tk.BOTH,expand=True,padx=6,pady=6); cols=('color_name','equipment','r_min','r_max','g_min','g_max','b_min','b_max','enabled'); self.color_tree=ttk.Treeview(rf,columns=cols,show='headings',height=8)
        for c in cols: self.color_tree.heading(c,text=c); self.color_tree.column(c,width=90)
        self.color_tree.pack(fill=tk.BOTH,expand=True); self.color_tree.bind('<<TreeviewSelect>>',self.on_color_select); ef=ttk.LabelFrame(f,text='選択色ルール編集',padding=6); ef.pack(fill=tk.X,padx=6,pady=6); self.color_vars={}
        for i,c in enumerate(cols): ttk.Label(ef,text=c).grid(row=i//5,column=(i%5)*2,sticky='e'); v=tk.StringVar(); ent=ttk.Entry(ef,textvariable=v,width=13); ent.grid(row=i//5,column=(i%5)*2+1,sticky='w'); self.color_vars[c]=v; ent.configure(state='readonly' if c=='color_name' else 'normal')
        ttk.Button(ef,text='色ルール保存',command=self.save_color_rule).grid(row=2,column=0,columnspan=2,pady=6)
    def build_ai_tab(self):
        f=self.tab_ai
        conn=ttk.LabelFrame(f,text='AI接続設定（Ollama / OpenAI / Claude / OpenAI互換）',padding=8)
        conn.pack(fill=tk.X,padx=6,pady=6)

        self.ai_provider_var=tk.StringVar(value=self.db.get_setting('ai_provider','ollama'))
        self.ollama_url_var=tk.StringVar(value=self.db.get_setting('ollama_url',DEFAULT_OLLAMA_URL))
        self.ollama_model_var=tk.StringVar(value=self.db.get_setting('ollama_model',DEFAULT_OLLAMA_MODEL))
        self.openai_api_key_var=tk.StringVar(value=self.db.get_setting('openai_api_key',''))
        self.openai_model_var=tk.StringVar(value=self.db.get_setting('openai_model',DEFAULT_OPENAI_MODEL))
        self.anthropic_api_key_var=tk.StringVar(value=self.db.get_setting('anthropic_api_key',''))
        self.anthropic_model_var=tk.StringVar(value=self.db.get_setting('anthropic_model',DEFAULT_CLAUDE_MODEL))
        self.custom_openai_base_url_var=tk.StringVar(value=self.db.get_setting('custom_openai_base_url',DEFAULT_CUSTOM_OPENAI_BASE_URL))
        self.custom_openai_api_key_var=tk.StringVar(value=self.db.get_setting('custom_openai_api_key',''))
        self.custom_openai_model_var=tk.StringVar(value=self.db.get_setting('custom_openai_model','local-model'))

        ttk.Label(conn,text='使用AI').grid(row=0,column=0,sticky='e')
        ttk.Combobox(conn,textvariable=self.ai_provider_var,values=['ollama','openai','anthropic','custom_openai'],state='readonly',width=18).grid(row=0,column=1,sticky='w',padx=4)
        ttk.Button(conn,text='AI設定保存',command=self.save_ollama_settings).grid(row=0,column=2,padx=4)
        ttk.Button(conn,text='選択AIテスト',command=self.test_selected_ai_async).grid(row=0,column=3,padx=4)
        self.db_backend_var=tk.StringVar(value=self.db.get_setting('db_backend','sqlite'))
        self.rqlite_url_var=tk.StringVar(value=self.db.get_setting('rqlite_url','http://127.0.0.1:4001'))
        ttk.Label(conn,text='DB').grid(row=4,column=3,sticky='e')
        ttk.Combobox(conn,textvariable=self.db_backend_var,values=['sqlite','rqlite'],width=10).grid(row=4,column=4,sticky='w')
        ttk.Entry(conn,textvariable=self.rqlite_url_var,width=30).grid(row=4,column=5,sticky='w')
        ttk.Button(conn,text='DB設定保存',command=self.save_db_backend_settings).grid(row=4,column=6,padx=3)
        ttk.Button(conn,text='rqlite接続テスト',command=self.test_rqlite_connection).grid(row=4,column=7,padx=3)
        self.vector_backend_var=tk.StringVar(value=self.db.get_setting('vector_backend','auto'))
        self.clip_model_name_var=tk.StringVar(value=self.db.get_setting('clip_model_name','openai/clip-vit-base-patch32'))
        self.vector_threshold_var=tk.StringVar(value=self.db.get_setting('vector_threshold','0.72'))
        ttk.Label(conn,text='Vector').grid(row=5,column=3,sticky='e')
        ttk.Combobox(conn,textvariable=self.vector_backend_var,values=['auto','clip','fallback'],width=10).grid(row=5,column=4,sticky='w')
        ttk.Entry(conn,textvariable=self.clip_model_name_var,width=30).grid(row=5,column=5,sticky='w')
        ttk.Label(conn,text='閾値').grid(row=5,column=6,sticky='e')
        ttk.Entry(conn,textvariable=self.vector_threshold_var,width=6).grid(row=5,column=7,sticky='w')
        ttk.Button(conn,text='Vector設定保存',command=self.save_vector_settings).grid(row=5,column=8,padx=3); ttk.Button(conn,text='rqlite起動方法',command=self.show_rqlite_start_guide).grid(row=4,column=8,padx=3); self.ai_usage_var=tk.StringVar(value='使用量: -'); ttk.Label(conn,textvariable=self.ai_usage_var,foreground='blue').grid(row=0,column=4,columnspan=2,sticky='w')
        self.use_litellm_var=tk.BooleanVar(value=self.db.get_setting('use_litellm','0')=='1')
        self.use_langmem_var=tk.BooleanVar(value=self.db.get_setting('use_langmem','1')=='1')
        self.memory_top_k_var=tk.StringVar(value=self.db.get_setting('memory_top_k','6'))
        ttk.Checkbutton(conn,text='LiteLLM使用',variable=self.use_litellm_var).grid(row=1,column=3,sticky='w')
        ttk.Checkbutton(conn,text='LangMem記憶使用',variable=self.use_langmem_var).grid(row=2,column=3,sticky='w')
        ttk.Label(conn,text='記憶TopK').grid(row=3,column=3,sticky='e')
        ttk.Entry(conn,textvariable=self.memory_top_k_var,width=6).grid(row=3,column=4,sticky='w')

        ttk.Label(conn,text='Ollama URL').grid(row=1,column=0,sticky='e')
        ttk.Entry(conn,textvariable=self.ollama_url_var,width=42).grid(row=1,column=1,sticky='w')
        ttk.Label(conn,text='Ollamaモデル').grid(row=2,column=0,sticky='e')
        self.model_combo=ttk.Combobox(conn,textvariable=self.ollama_model_var,width=40)
        self.model_combo.grid(row=2,column=1,sticky='w')
        ttk.Button(conn,text='Ollamaモデル一覧取得',command=self.refresh_models_async).grid(row=1,column=2,padx=4)
        ttk.Button(conn,text='Ollama接続テスト',command=self.test_ollama_async).grid(row=2,column=2,padx=4)

        ttk.Label(conn,text='OpenAI APIキー').grid(row=3,column=0,sticky='e')
        _oai_frame=ttk.Frame(conn); _oai_frame.grid(row=3,column=1,columnspan=3,sticky='w')
        self._openai_key_entry=ttk.Entry(_oai_frame,textvariable=self.openai_api_key_var,width=36,show='*'); self._openai_key_entry.pack(side=tk.LEFT)
        ttk.Button(_oai_frame,text='貼付',width=5,command=lambda:self._paste_to_var(self.openai_api_key_var,self._openai_key_entry)).pack(side=tk.LEFT,padx=2)
        ttk.Button(_oai_frame,text='表示',width=5,command=lambda:self._toggle_show(self._openai_key_entry)).pack(side=tk.LEFT,padx=2)
        ttk.Label(conn,text='OpenAIモデル').grid(row=4,column=0,sticky='e')
        self.openai_model_combo=ttk.Combobox(conn,textvariable=self.openai_model_var,width=40); self.openai_model_combo.grid(row=4,column=1,sticky='w'); ttk.Button(conn,text='OpenAIモデル一覧',command=self.refresh_openai_models_async).grid(row=4,column=2,padx=4)

        ttk.Label(conn,text='Claude APIキー').grid(row=5,column=0,sticky='e')
        _claude_frame=ttk.Frame(conn); _claude_frame.grid(row=5,column=1,columnspan=3,sticky='w')
        self._claude_key_entry=ttk.Entry(_claude_frame,textvariable=self.anthropic_api_key_var,width=36,show='*'); self._claude_key_entry.pack(side=tk.LEFT)
        ttk.Button(_claude_frame,text='貼付',width=5,command=lambda:self._paste_to_var(self.anthropic_api_key_var,self._claude_key_entry)).pack(side=tk.LEFT,padx=2)
        ttk.Button(_claude_frame,text='表示',width=5,command=lambda:self._toggle_show(self._claude_key_entry)).pack(side=tk.LEFT,padx=2)
        ttk.Label(conn,text='Claudeモデル').grid(row=6,column=0,sticky='e')
        self.anthropic_model_combo=ttk.Combobox(conn,textvariable=self.anthropic_model_var,values=CLAUDE_MODEL_CHOICES,width=40); self.anthropic_model_combo.grid(row=6,column=1,sticky='w'); ttk.Button(conn,text='Claudeモデル一覧',command=self.refresh_anthropic_models_async).grid(row=6,column=2,padx=4)

        ttk.Label(conn,text='OpenAI互換Base URL').grid(row=7,column=0,sticky='e')
        ttk.Entry(conn,textvariable=self.custom_openai_base_url_var,width=42).grid(row=7,column=1,sticky='w')
        ttk.Label(conn,text='互換APIキー').grid(row=8,column=0,sticky='e')
        _custom_frame=ttk.Frame(conn); _custom_frame.grid(row=8,column=1,columnspan=3,sticky='w')
        self._custom_key_entry=ttk.Entry(_custom_frame,textvariable=self.custom_openai_api_key_var,width=36,show='*'); self._custom_key_entry.pack(side=tk.LEFT)
        ttk.Button(_custom_frame,text='貼付',width=5,command=lambda:self._paste_to_var(self.custom_openai_api_key_var,self._custom_key_entry)).pack(side=tk.LEFT,padx=2)
        ttk.Button(_custom_frame,text='表示',width=5,command=lambda:self._toggle_show(self._custom_key_entry)).pack(side=tk.LEFT,padx=2)
        ttk.Label(conn,text='互換モデル').grid(row=9,column=0,sticky='e')
        self.custom_openai_model_combo=ttk.Combobox(conn,textvariable=self.custom_openai_model_var,width=40); self.custom_openai_model_combo.grid(row=9,column=1,sticky='w')

        gen=ttk.LabelFrame(f,text='AI生成：図面記号 / 2Dファミリ / DB登録',padding=8)
        gen.pack(fill=tk.X,padx=6,pady=6)

        self.gen_name_var=tk.StringVar(value='LEDダウンライト')
        self.gen_category_var=tk.StringVar(value='照明器具')
        self.gen_kind_var=tk.StringVar(value='図面記号')

        ttk.Label(gen,text='名称').grid(row=0,column=0,sticky='e')
        ttk.Entry(gen,textvariable=self.gen_name_var,width=28).grid(row=0,column=1,sticky='w',padx=4)
        ttk.Label(gen,text='カテゴリ').grid(row=0,column=2,sticky='e')
        ttk.Combobox(gen,textvariable=self.gen_category_var,values=['照明器具','コンセント','スイッチ','分電盤','配線','空調','その他'],width=16).grid(row=0,column=3,sticky='w',padx=4)
        ttk.Label(gen,text='種類').grid(row=0,column=4,sticky='e')
        ttk.Combobox(gen,textvariable=self.gen_kind_var,values=['図面記号','2Dファミリ','JWCAD外部図形','DXFブロック','SVG記号'],width=16).grid(row=0,column=5,sticky='w',padx=4)

        ttk.Button(gen,text='AIで図面記号を生成',command=self.generate_symbol_async).grid(row=1,column=1,sticky='w',pady=6)
        ttk.Button(gen,text='AIで2Dファミリを生成',command=self.generate_family_async).grid(row=1,column=2,sticky='w',pady=6)
        ttk.Button(gen,text='生成結果をDB登録',command=self.register_generated_asset).grid(row=1,column=3,sticky='w',pady=6)
        ttk.Button(gen,text='生成DB一覧更新',command=self.refresh_generated_assets).grid(row=1,column=4,sticky='w',pady=6)

        self.generated_asset_text=scrolledtext.ScrolledText(f,height=8)
        self.generated_asset_text.pack(fill=tk.X,padx=6,pady=4)

        cols=('ID','日時','名称','カテゴリ','種類/説明')
        self.generated_tree=ttk.Treeview(f,columns=cols,show='headings',height=5)
        for c in cols:
            self.generated_tree.heading(c,text=c)
            self.generated_tree.column(c,width=120 if c!='種類/説明' else 360)
        self.generated_tree.pack(fill=tk.X,padx=6,pady=4)

        self.ai_prompt=scrolledtext.ScrolledText(f,height=8)
        self.ai_prompt.pack(fill=tk.X,padx=6,pady=6)
        self.ai_prompt.insert('1.0','現在の積算結果・画像検出結果について、過検出を減らす調整案と積算上の注意点を提案してください。')
        btn=ttk.Frame(f); btn.pack(fill=tk.X,padx=6)
        ttk.Button(btn,text='結果をプロンプトに追加',command=self.append_summary_prompt).pack(side=tk.LEFT,padx=3)
        ttk.Button(btn,text='選択AIで解析実行',command=self.run_ollama_async).pack(side=tk.LEFT,padx=3)
        self.ai_result=scrolledtext.ScrolledText(f)
        self.ai_result.pack(fill=tk.BOTH,expand=True,padx=6,pady=6)

    def make_symbol_generation_prompt(self, family=False):
        name=self.gen_name_var.get().strip() or '未名称'
        category=self.gen_category_var.get().strip() or 'その他'
        kind=self.gen_kind_var.get().strip() or ('2Dファミリ' if family else '図面記号')
        return f"""あなたは日本の電気設備CAD図面の記号・2Dファミリ生成補助AIです。
以下の設備について、JWCADやDXFに変換しやすい2D記号仕様を生成してください。

名称: {name}
カテゴリ: {category}
種類: {kind}

出力形式:
1. 説明
2. DXF_ENTITIES
3. SVG
4. JSON_SPEC

条件:
- 2D平面記号として単純な線・円・矩形で表現する
- 座標は原点中心、単位はmm想定
- DXF_ENTITIESはLINE/CIRCLE/LWPOLYLINE相当の簡易テキストでよい
- JSON_SPECにはname, category, kind, primitivesを含める
- 回答は日本語で簡潔に
"""

    def build_dqn_tab(self):
        f=self.tab_dqn
        top=ttk.LabelFrame(f,text='DQN解析戦略エージェント',padding=8); top.pack(fill=tk.X,padx=6,pady=6)
        self.dqn_enabled_var=tk.BooleanVar(value=self.db.get_setting('dqn_enabled','1')=='1')
        self.dqn_epsilon_var=tk.StringVar(value=self.db.get_setting('dqn_epsilon','0.15'))
        self.dqn_lr_var=tk.StringVar(value=self.db.get_setting('dqn_learning_rate','0.08'))
        self.dqn_gamma_var=tk.StringVar(value=self.db.get_setting('dqn_gamma','0.90'))
        ttk.Checkbutton(top,text='DQN戦略ON',variable=self.dqn_enabled_var).grid(row=0,column=0,sticky='w')
        ttk.Label(top,text='ε').grid(row=0,column=1); ttk.Entry(top,textvariable=self.dqn_epsilon_var,width=8).grid(row=0,column=2)
        ttk.Label(top,text='学習率').grid(row=0,column=3); ttk.Entry(top,textvariable=self.dqn_lr_var,width=8).grid(row=0,column=4)
        ttk.Label(top,text='γ').grid(row=0,column=5); ttk.Entry(top,textvariable=self.dqn_gamma_var,width=8).grid(row=0,column=6)
        ttk.Button(top,text='設定保存',command=self.save_dqn_settings).grid(row=0,column=7,padx=4)
        ttk.Button(top,text='サマリー更新',command=self.refresh_dqn_summary).grid(row=0,column=8,padx=4)
        ttk.Button(top,text='最後の戦略に+報酬',command=lambda:self.manual_dqn_reward(1.0)).grid(row=1,column=0,pady=4)
        ttk.Button(top,text='最後の戦略に-報酬',command=lambda:self.manual_dqn_reward(-1.0)).grid(row=1,column=1,pady=4)
        self.dqn_status_var=tk.StringVar(value='DQN状態: 未実行')
        ttk.Label(top,textvariable=self.dqn_status_var,foreground='blue').grid(row=2,column=0,columnspan=9,sticky='w',pady=4)
        mid=ttk.PanedWindow(f,orient=tk.HORIZONTAL); mid.pack(fill=tk.BOTH,expand=True,padx=6,pady=6)
        left=ttk.LabelFrame(mid,text='アクション別成績',padding=4); mid.add(left,weight=1)
        cols=('action','trials','reward_sum','reward_avg','last_used'); self.dqn_tree=ttk.Treeview(left,columns=cols,show='headings',height=12)
        for c in cols: self.dqn_tree.heading(c,text=c); self.dqn_tree.column(c,width=110,anchor=tk.CENTER)
        self.dqn_tree.column('action',width=150); self.dqn_tree.pack(fill=tk.BOTH,expand=True)
        right=ttk.LabelFrame(mid,text='最近の戦略イベント/状態',padding=4); mid.add(right,weight=1)
        self.dqn_text=scrolledtext.ScrolledText(right,height=18); self.dqn_text.pack(fill=tk.BOTH,expand=True)
        self.refresh_dqn_summary()

    def save_dqn_settings(self):
        self.db.set_setting('dqn_enabled','1' if self.dqn_enabled_var.get() else '0'); self.db.set_setting('dqn_epsilon',self.dqn_epsilon_var.get().strip() or '0.15'); self.db.set_setting('dqn_learning_rate',self.dqn_lr_var.get().strip() or '0.08'); self.db.set_setting('dqn_gamma',self.dqn_gamma_var.get().strip() or '0.90')
        messagebox.showinfo('保存','DQN設定を保存しました')

    def refresh_dqn_summary(self):
        try:
            if hasattr(self,'dqn_tree'):
                for i in self.dqn_tree.get_children(): self.dqn_tree.delete(i)
                for a,tr,rs,ra,last in self.db.get_dqn_summary(): self.dqn_tree.insert('',tk.END,values=(a,tr,f'{rs:.2f}',f'{ra:.3f}',last or ''))
            if hasattr(self,'dqn_text'):
                lines=[f'最後の状態: {json.dumps(getattr(self,"last_dqn_state",{}),ensure_ascii=False,indent=2)}',f'最後のアクション: {getattr(self,"last_dqn_action","")} / mode={getattr(self,"last_dqn_mode","")} / reward={getattr(self,"last_dqn_reward",0)}','', '[アクション説明]']
                for a in DQN_ACTIONS: lines.append(f'- {a}: {DQN_ACTION_DESCRIPTIONS.get(a,"")}')
                lines.append(''); lines.append('[最近のイベント]')
                for ev in self.db.get_recent_dqn_events(30): lines.append(f"{ev.get('created_at')} action={ev.get('action')} reward={ev.get('reward')} note={ev.get('note')}")
                self.dqn_text.delete('1.0',tk.END); self.dqn_text.insert('1.0','\n'.join(lines))
            if hasattr(self,'dqn_status_var'): self.dqn_status_var.set(f'DQN状態: last_action={getattr(self,"last_dqn_action","")} reward={getattr(self,"last_dqn_reward",0)}')
        except Exception: self.log(traceback.format_exc())

    def manual_dqn_reward(self,reward):
        action=getattr(self,'last_dqn_action','') or self.db.get_setting('dqn_last_action','')
        if not action: messagebox.showwarning('なし','まだDQN戦略が実行されていません'); return
        self.dqn_agent.update(getattr(self,'last_dqn_state',{}),action,reward,None,lr=float(self.db.get_setting('dqn_learning_rate','0.08')),gamma=float(self.db.get_setting('dqn_gamma','0.90')))
        self.db.add_dqn_reward(action,reward,'manual reward button'); self.last_dqn_reward=reward; self.refresh_dqn_summary(); messagebox.showinfo('報酬登録',f'{action} に reward={reward} を登録しました')

    def build_mindmap_tab(self):
        f=self.tab_mindmap
        top=ttk.Frame(f); top.pack(fill=tk.X,padx=3,pady=2)
        self.mm_task_var=tk.StringVar(value='埋込ダブルコンセント取付工事')
        self.mm_qty_var=tk.StringVar(value='1')
        ttk.Button(top,text='操作▼',command=self.show_mindmap_menu,width=7).pack(side=tk.LEFT,padx=1)
        ttk.Label(top,text='工程').pack(side=tk.LEFT)
        ttk.Entry(top,textvariable=self.mm_task_var,width=24).pack(side=tk.LEFT,fill=tk.X,expand=True,padx=2)
        ttk.Label(top,text='数量').pack(side=tk.LEFT)
        ttk.Entry(top,textvariable=self.mm_qty_var,width=5).pack(side=tk.LEFT,padx=2)
        ttk.Button(top,text='AI生成',command=self.generate_mindmap_async).pack(side=tk.LEFT,padx=2)
        ttk.Button(top,text='積算反映',command=self.mindmap_to_estimate).pack(side=tk.LEFT,padx=2)
        ttk.Button(top,text='保存',command=self.save_current_mindmap).pack(side=tk.LEFT,padx=2)
        pane=ttk.PanedWindow(f,orient=tk.HORIZONTAL); pane.pack(fill=tk.BOTH,expand=True,padx=3,pady=2)
        left=ttk.Frame(pane); pane.add(left,weight=3)
        right=ttk.Frame(pane); pane.add(right,weight=2)
        self.mm_canvas=tk.Canvas(left,bg='white',width=900,height=650,scrollregion=(0,0,1600,1200))
        v=ttk.Scrollbar(left,orient=tk.VERTICAL,command=self.mm_canvas.yview); h=ttk.Scrollbar(left,orient=tk.HORIZONTAL,command=self.mm_canvas.xview)
        self.mm_canvas.configure(yscrollcommand=v.set,xscrollcommand=h.set)
        self.mm_canvas.grid(row=0,column=0,sticky='nsew'); v.grid(row=0,column=1,sticky='ns'); h.grid(row=1,column=0,sticky='ew')
        left.rowconfigure(0,weight=1); left.columnconfigure(0,weight=1)
        self.mm_tree=ttk.Treeview(right,columns=('type','qty','unit','minutes'),show='tree headings')
        for c,t in [('#0','ノード'),('type','type'),('qty','数量'),('unit','単位'),('minutes','分')]:
            self.mm_tree.heading(c,text=t)
        self.mm_tree.pack(fill=tk.BOTH,expand=True)
        self.mm_json_text=scrolledtext.ScrolledText(right,height=8)
        self.mm_json_text.pack(fill=tk.BOTH,expand=False,pady=3)
        self.current_mindmap=default_mindmap_json(self.mm_task_var.get(),1)
        self.render_mindmap()

    def show_mindmap_menu(self):
        m=tk.Menu(self.root,tearoff=0)
        m.add_command(label='AI生成',command=self.generate_mindmap_async)
        m.add_command(label='フォールバック生成',command=lambda:self.set_mindmap(default_mindmap_json(self.mm_task_var.get(),float(self.mm_qty_var.get() or 1))))
        m.add_command(label='積算へ反映',command=self.mindmap_to_estimate)
        m.add_command(label='JSON保存',command=self.save_current_mindmap)
        m.add_command(label='JSON再描画',command=self.load_mindmap_from_json_text)
        m.tk_popup(self.root.winfo_pointerx(),self.root.winfo_pointery())

    def set_mindmap(self,data):
        self.current_mindmap=data
        self.render_mindmap()

    def generate_mindmap_async(self):
        task=self.mm_task_var.get().strip()
        try: qty=float(self.mm_qty_var.get() or 1)
        except Exception: qty=1
        snap=self.snapshot_ai_settings() if hasattr(self,'snapshot_ai_settings') else {'provider':'ollama'}
        def worker():
            prompt=('日本の電気設備工事について、材料・工具・施工手順・標準時間・積算情報をJSONだけで返してください。'
                    '必ず project_task,total_estimated_minutes,mindmap_nodes を含める。'
                    'mindmap_nodesは id,text,parent,type,meta を持つ配列。'
                    '材料 type=material meta={default_qty,unit}、手順 type=step meta={order,minutes}。'
                    f'工事名:{task} 数量:{qty}')
            ans=UnifiedAIClient(snap).generate(prompt,timeout=45)
            try: data=sanitize_ai_json(str(ans))
            except Exception:
                data=default_mindmap_json(task,qty); data['ai_parse_error']=str(ans)[:500]
            data['root_qty']=qty
            return data
        def done(data):
            self.set_mindmap(data); self.status.config(text='ToDoマインドマップAI生成完了')
        def err(tb):
            self.log(tb); self.set_mindmap(default_mindmap_json(task,qty)); messagebox.showwarning('AI生成失敗','フォールバック工程表を生成しました。')
        self.run_in_thread(worker,done,err)

    def render_mindmap(self):
        data=getattr(self,'current_mindmap',default_mindmap_json('電気設備工事',1))
        nodes=data.get('mindmap_nodes',[])
        by_parent=defaultdict(list)
        for n in nodes: by_parent[n.get('parent')].append(n)
        self.mm_canvas.delete('all')
        for i in self.mm_tree.get_children(): self.mm_tree.delete(i)
        root=next((n for n in nodes if n.get('parent') is None), nodes[0] if nodes else {'id':'root','text':'root'})
        root_x,root_y=420,80
        self.draw_mm_node(root_x,root_y,root)
        cats=by_parent.get(root.get('id'),[])
        for ci,cat in enumerate(cats):
            x=180+ci*280; y=220
            self.mm_canvas.create_line(root_x,root_y+25,x,y-25,fill='#777',width=2)
            self.draw_mm_node(x,y,cat)
            for ki,k in enumerate(by_parent.get(cat.get('id'),[])):
                ky=y+110+ki*78; kx=x
                self.mm_canvas.create_line(x,y+25,kx,ky-25,fill='#aaa')
                self.draw_mm_node(kx,ky,k,small=True)
        def add_tree(parent_iid,node):
            meta=node.get('meta') or {}; iid=node.get('id')
            self.mm_tree.insert(parent_iid,'end',iid=iid,text=node.get('text',''),values=(node.get('type',''),meta.get('default_qty',''),meta.get('unit',''),meta.get('minutes','')))
            for child in by_parent.get(node.get('id'),[]): add_tree(iid,child)
        try: add_tree('',root)
        except Exception: pass
        self.mm_json_text.delete('1.0',tk.END); self.mm_json_text.insert('1.0',json.dumps(data,ensure_ascii=False,indent=2))
        self.mm_canvas.configure(scrollregion=self.mm_canvas.bbox('all') or (0,0,1600,1200))

    def draw_mm_node(self,x,y,node,small=False):
        typ=node.get('type','')
        color={'root':'#d9edf7','category':'#eeeeee','material':'#dff0d8','tool':'#fcf8e3','step':'#f2dede','cost':'#e8daef'}.get(typ,'#ffffff')
        w=210 if not small else 190; h=54 if not small else 46
        self.mm_canvas.create_rectangle(x-w/2,y-h/2,x+w/2,y+h/2,fill=color,outline='#555',width=2)
        txt=node.get('text',''); txt=txt if len(txt)<=24 else txt[:24]+'…'
        self.mm_canvas.create_text(x,y,text=txt,font=('Arial',10 if not small else 9),width=w-10)

    def load_mindmap_from_json_text(self):
        try: self.set_mindmap(json.loads(self.mm_json_text.get('1.0',tk.END)))
        except Exception as e: messagebox.showwarning('JSONエラー',str(e))

    def save_current_mindmap(self):
        data=getattr(self,'current_mindmap',None)
        if not data: return
        self.db.save_todo_mindmap(data.get('project_task','mindmap'),float(data.get('root_qty',1) or 1),float(data.get('total_estimated_minutes',0) or 0),data)
        messagebox.showinfo('保存','ToDoマインドマップを保存しました')

    def mindmap_to_estimate(self):
        data=getattr(self,'current_mindmap',None)
        if not data: return
        try: qty=float(data.get('root_qty',self.mm_qty_var.get() or 1))
        except Exception: qty=1
        rows=[]; total_minutes=0
        for n in data.get('mindmap_nodes',[]):
            typ=n.get('type'); meta=n.get('meta') or {}
            if typ=='material':
                name=n.get('text',''); q=float(meta.get('default_qty',1) or 1)*qty; unit=meta.get('unit','個')
                cat,item,spec,u,price=self.db.find_price(name)
                rows.append((cat,item or name,unit,q,price,float(price)*q))
            elif typ=='step':
                total_minutes += float(meta.get('minutes',0) or 0)*qty
        if total_minutes:
            labor_price=float(self.db.get_setting('labor_unit_price_per_hour','4500') or 4500)
            hours=total_minutes/60.0
            rows.append(('労務','施工労務費','h',round(hours,2),labor_price,round(hours*labor_price)))
        self.result_rows=rows; self.update_result_tree(rows)
        try: self.nb.select(self.tab_est)
        except Exception: pass
        messagebox.showinfo('積算反映',f'{len(rows)} 行を積算へ反映しました')

    def build_nodecad_tab(self):
        f=self.tab_nodecad
        top=ttk.Frame(f); top.pack(fill=tk.X,padx=3,pady=2)
        ttk.Button(top,text='操作▼',command=self.show_nodecad_menu,width=7).pack(side=tk.LEFT,padx=1)
        ttk.Button(top,text='記号ノード',command=lambda:self.add_vnode('symbol')).pack(side=tk.LEFT,padx=1)
        ttk.Button(top,text='ケーブルノード',command=lambda:self.add_vnode('cable')).pack(side=tk.LEFT,padx=1)
        ttk.Button(top,text='接続モード',command=self.toggle_node_connect_mode).pack(side=tk.LEFT,padx=1)
        ttk.Button(top,text='図面生成',command=self.render_nodecad_drawing).pack(side=tk.LEFT,padx=1)
        ttk.Button(top,text='保存',command=self.save_node_graph).pack(side=tk.LEFT,padx=1)
        self.nodecad_status_var=tk.StringVar(value='ノード追加→ドラッグ配置→接続→図面生成')
        ttk.Label(top,textvariable=self.nodecad_status_var).pack(side=tk.LEFT,padx=8)
        pane=ttk.PanedWindow(f,orient=tk.HORIZONTAL); pane.pack(fill=tk.BOTH,expand=True,padx=3,pady=2)
        left=ttk.Frame(pane); pane.add(left,weight=1); right=ttk.Frame(pane); pane.add(right,weight=1)
        self.node_canvas=tk.Canvas(left,bg='#f8f8f8',width=760,height=720,scrollregion=(0,0,1800,1400)); self.node_canvas.pack(fill=tk.BOTH,expand=True)
        self.node_canvas.bind('<Button-1>',self.on_node_canvas_click); self.node_canvas.bind('<B1-Motion>',self.on_node_canvas_drag); self.node_canvas.bind('<Double-1>',self.on_node_canvas_double)
        self.node_draw_canvas=tk.Canvas(right,bg='white',width=760,height=720,scrollregion=(0,0,1800,1400)); self.node_draw_canvas.pack(fill=tk.BOTH,expand=True)
        self.vnodes=[]; self.vedges=[]; self.node_connect_mode=False; self.node_selected=None; self.node_drag_offset=(0,0); self.node_id_seq=1

    def show_nodecad_menu(self):
        m=tk.Menu(self.root,tearoff=0)
        m.add_command(label='記号ノード追加',command=lambda:self.add_vnode('symbol')); m.add_command(label='ケーブルノード追加',command=lambda:self.add_vnode('cable')); m.add_command(label='分電盤ノード追加',command=lambda:self.add_vnode('panel'))
        m.add_command(label='接続モードON/OFF',command=self.toggle_node_connect_mode); m.add_separator(); m.add_command(label='選択ノード編集',command=self.edit_selected_vnode); m.add_command(label='図面生成',command=self.render_nodecad_drawing); m.add_command(label='保存',command=self.save_node_graph)
        m.tk_popup(self.root.winfo_pointerx(),self.root.winfo_pointery())

    def add_vnode(self,typ='symbol'):
        nid=f'N{self.node_id_seq}'; self.node_id_seq+=1
        symbol={'symbol':'LEDダウンライト','cable':'VVF2.0-3C','panel':'分電盤'}.get(typ,'LEDダウンライト')
        self.vnodes.append({'id':nid,'x':120+len(self.vnodes)*40,'y':120+len(self.vnodes)*30,'type':typ,'text':symbol,'symbol':symbol,'qty':1,'unit':'個'})
        self.draw_node_graph()

    def draw_node_graph(self):
        self.node_canvas.delete('all')
        for a,b in self.vedges:
            na=next((n for n in self.vnodes if n['id']==a),None); nb=next((n for n in self.vnodes if n['id']==b),None)
            if na and nb: self.node_canvas.create_line(na['x'],na['y'],nb['x'],nb['y'],fill='#555',width=3,arrow=tk.LAST)
        for n in self.vnodes:
            x,y=n['x'],n['y']; color={'symbol':'#d9edf7','cable':'#fce5cd','panel':'#d9d2e9','process':'#eeeeee'}.get(n.get('type'),'#fff')
            self.node_canvas.create_rectangle(x-70,y-28,x+70,y+28,fill=color,outline='#333',width=2,tags=('node',n['id']))
            self.node_canvas.create_text(x,y,text=f"{n.get('text','')}\n{n.get('id')}",tags=('node',n['id']),font=('Arial',9))
        self.node_canvas.configure(scrollregion=self.node_canvas.bbox('all') or (0,0,1800,1400))

    def find_vnode_at(self,x,y):
        for n in reversed(self.vnodes):
            if abs(n['x']-x)<=75 and abs(n['y']-y)<=35: return n
        return None

    def on_node_canvas_click(self,event):
        x=self.node_canvas.canvasx(event.x); y=self.node_canvas.canvasy(event.y); n=self.find_vnode_at(x,y)
        if self.node_connect_mode and n:
            if not self.node_selected:
                self.node_selected=n['id']; self.nodecad_status_var.set(f'接続元: {n["id"]}。接続先をクリック。')
            else:
                if self.node_selected!=n['id']: self.vedges.append((self.node_selected,n['id']))
                self.node_selected=None; self.node_connect_mode=False; self.nodecad_status_var.set('接続完了'); self.draw_node_graph()
            return
        self.node_selected=n['id'] if n else None
        if n: self.node_drag_offset=(n['x']-x,n['y']-y)

    def on_node_canvas_drag(self,event):
        if not self.node_selected: return
        x=self.node_canvas.canvasx(event.x); y=self.node_canvas.canvasy(event.y)
        n=next((n for n in self.vnodes if n['id']==self.node_selected),None)
        if n:
            n['x']=x+self.node_drag_offset[0]; n['y']=y+self.node_drag_offset[1]; self.draw_node_graph()

    def on_node_canvas_double(self,event):
        x=self.node_canvas.canvasx(event.x); y=self.node_canvas.canvasy(event.y); n=self.find_vnode_at(x,y)
        if n: self.node_selected=n['id']; self.edit_selected_vnode()

    def edit_selected_vnode(self):
        n=next((n for n in self.vnodes if n['id']==self.node_selected),None)
        if not n: messagebox.showwarning('未選択','編集するノードを選択してください'); return
        dlg=NodePropertyDialog(self.root,n); self.root.wait_window(dlg)
        if dlg.result:
            n.update(dlg.result); n['text']=dlg.result.get('text') or dlg.result.get('symbol'); self.draw_node_graph()

    def toggle_node_connect_mode(self):
        self.node_connect_mode=not self.node_connect_mode; self.node_selected=None
        self.nodecad_status_var.set('接続モードON：接続元→接続先の順にクリック' if self.node_connect_mode else '接続モードOFF')

    def render_nodecad_drawing(self):
        c=self.node_draw_canvas; c.delete('all')
        for a,b in self.vedges:
            na=next((n for n in self.vnodes if n['id']==a),None); nb=next((n for n in self.vnodes if n['id']==b),None)
            if na and nb:
                c.create_line(na['x'],na['y'],nb['x'],nb['y'],fill='orange',width=3)
                c.create_text((na['x']+nb['x'])/2,(na['y']+nb['y'])/2-8,text='VVF',fill='orange')
        for n in self.vnodes:
            draw_symbol_glyph_on_canvas(c,n['x'],n['y'],n.get('symbol',n.get('text','')),scale=1.0,tag='nodecad_symbol')
        c.configure(scrollregion=c.bbox('all') or (0,0,1800,1400)); self.nodecad_status_var.set('ノード図面を生成しました')

    def save_node_graph(self):
        self.db.save_visual_node_graph('ノード図面',self.vnodes,self.vedges)
        messagebox.showinfo('保存','ノード図面グラフを保存しました')

    def build_learning_tab(self):
        f=self.tab_learn
        btn=ttk.Frame(f); btn.pack(fill=tk.X,padx=6,pady=6)
        ttk.Button(btn,text='更新',command=self.refresh_learning_summary).pack(side=tk.LEFT)
        ttk.Button(btn,text='学習CSV出力',command=self.export_learning_csv).pack(side=tk.LEFT,padx=4)
        ttk.Button(btn,text='過去アノテーションから特徴量再構築',command=self.rebuild_learning_features).pack(side=tk.LEFT,padx=4)
        ttk.Button(btn,text='現在の検出へ学習反映',command=lambda:self.apply_learning_to_detections(show_message=True)).pack(side=tk.LEFT,padx=4)
        ttk.Button(btn,text='長期記憶サマリー更新',command=self.refresh_learning_summary).pack(side=tk.LEFT,padx=4)
        self.learning_text=scrolledtext.ScrolledText(f)
        self.learning_text.pack(fill=tk.BOTH,expand=True,padx=6,pady=6)

    # behaviors
    def refresh_all(self): self.price_refresh(); self.symbol_refresh(); self.cad_refresh(); self.refresh_color_rules(); self.refresh_learning_summary()
    def select_file(self):
        p=filedialog.askopenfilename(filetypes=[('PDF/DXF','*.pdf;*.dxf'),('All','*.*')]);
        if p: self.file_path_var.set(p)
    # ----------------------------------------------------------------
    # 縮尺測定メソッド
    # ----------------------------------------------------------------
    def on_scale_measure_start(self, event):
        """縮尺測定開始: キャンバス上の開始点を記録"""
        cx=self.preview_canvas.canvasx(event.x); cy=self.preview_canvas.canvasy(event.y)
        self.scale_start_canvas=(cx,cy)
        for _id in getattr(self,'scale_line_ids',[]):
            try: self.preview_canvas.delete(_id)
            except Exception: pass
        self.scale_line_ids=[]
        self.status.config(text='縮尺測定: 既知寸法線の終点までドラッグしてください')

    def on_scale_measure_drag(self, event):
        """縮尺測定ドラッグ: 測定線をリアルタイム描画"""
        if not getattr(self,'scale_start_canvas',None): return
        cx=self.preview_canvas.canvasx(event.x); cy=self.preview_canvas.canvasy(event.y)
        x0,y0=self.scale_start_canvas
        for _id in getattr(self,'scale_line_ids',[]):
            try: self.preview_canvas.delete(_id)
            except Exception: pass
        canvas_dist=math.hypot(cx-x0,cy-y0)
        image_dist=canvas_dist/max(0.0001,self.preview_scale)
        ids=[
            self.preview_canvas.create_line(x0,y0,cx,cy,fill='red',width=2,dash=(6,3)),
            self.preview_canvas.create_oval(x0-5,y0-5,x0+5,y0+5,outline='red',width=2),
            self.preview_canvas.create_oval(cx-5,cy-5,cx+5,cy+5,outline='red',width=2),
            self.preview_canvas.create_text((x0+cx)/2,(y0+cy)/2-16,
                text=f'{image_dist:.0f}px',fill='red',font=('Arial',11,'bold'),
                tags=('scale_label',)),
        ]
        self.scale_line_ids=ids

    def on_scale_measure_release(self, event):
        """縮尺測定完了: 既知長さを入力してスケール設定"""
        if not getattr(self,'scale_start_canvas',None): return
        cx=self.preview_canvas.canvasx(event.x); cy=self.preview_canvas.canvasy(event.y)
        x0,y0=self.scale_start_canvas
        canvas_dist=math.hypot(cx-x0,cy-y0)
        image_dist=canvas_dist/max(0.0001,self.preview_scale)  # 画像ピクセル単位
        self.scale_start_canvas=None
        if image_dist<5:
            self.status.config(text='縮尺測定: 線が短すぎます。もう一度ドラッグしてください。'); return
        from tkinter.simpledialog import askfloat
        real_mm=askfloat(
            '縮尺設定',
            f'測定線: {image_dist:.1f} ピクセル\n\nこの線の実際の長さ（mm）を入力してください:\n例: 図面の「700」という寸法に合わせた場合は 700 を入力\n例: 10400mm（コンテナ幅）に合わせた場合は 10400 を入力',
        )
        if real_mm:
            self.scale_px_per_mm=image_dist/real_mm
            px_per_m=self.scale_px_per_mm*1000
            self.scale_info_var.set(f'縮尺: {image_dist:.0f}px={real_mm:.0f}mm | {px_per_m:.1f}px/m')
            try: self.db.set_setting('scale_px_per_mm',str(self.scale_px_per_mm))
            except Exception: pass
            self.status.config(text=f'縮尺設定完了: {real_mm:.0f}mm={image_dist:.0f}px → ケーブル長計算に自動適用')
            messagebox.showinfo('縮尺設定完了',
                f'縮尺を設定しました。\n\n'
                f'{image_dist:.0f}px = {real_mm:.0f}mm\n'
                f'1m = {px_per_m:.1f}px\n\n'
                f'これ以降のケーブル描画で長さが自動計算されます。')
        # 測定線を消す
        for _id in getattr(self,'scale_line_ids',[]):
            try: self.preview_canvas.delete(_id)
            except Exception: pass
        self.scale_line_ids=[]

    def clear_scale_calibration(self):
        """縮尺設定をクリア"""
        self.scale_px_per_mm=None
        self.scale_info_var.set('縮尺: 未設定')
        try: self.db.set_setting('scale_px_per_mm','')
        except Exception: pass
        for _id in getattr(self,'scale_line_ids',[]):
            try: self.preview_canvas.delete(_id)
            except Exception: pass
        self.scale_line_ids=[]
        self.status.config(text='縮尺設定をクリアしました')


    # ================================================================
    # AutoCADスナップ + 通り芯 機能
    # ================================================================

    def build_cad_snap_toolbar(self, parent):
        """AutoCADスナップモード + 通り芯機能のツールバーを構築"""
        # ── スナップ設定行 ──
        sf=ttk.LabelFrame(parent,text='AutoCADスナップ [F3=オブジェクト F9=グリッド]',padding=3)
        sf.pack(fill=tk.X,pady=1)
        self.snap_grid_var=tk.BooleanVar(value=False)
        self.snap_object_var=tk.BooleanVar(value=True)
        self.snap_grid_mm_var=tk.StringVar(value='100')
        self.snap_threshold_var=tk.StringVar(value='15')
        ttk.Checkbutton(sf,text='グリッドスナップ[F9]',variable=self.snap_grid_var).pack(side=tk.LEFT,padx=4)
        ttk.Label(sf,text='間隔:').pack(side=tk.LEFT)
        ttk.Entry(sf,textvariable=self.snap_grid_mm_var,width=6).pack(side=tk.LEFT)
        ttk.Label(sf,text='mm').pack(side=tk.LEFT)
        ttk.Checkbutton(sf,text='オブジェクトスナップ[F3]',variable=self.snap_object_var).pack(side=tk.LEFT,padx=8)
        ttk.Label(sf,text='感度:').pack(side=tk.LEFT)
        ttk.Entry(sf,textvariable=self.snap_threshold_var,width=4).pack(side=tk.LEFT)
        ttk.Label(sf,text='px').pack(side=tk.LEFT)
        self.snap_status_var=tk.StringVar(value='スナップ: 待機中')
        ttk.Label(sf,textvariable=self.snap_status_var,foreground='blue').pack(side=tk.LEFT,padx=8)
        self.root.bind('<F3>',lambda e:self.snap_object_var.set(not self.snap_object_var.get()))
        self.root.bind('<F9>',lambda e:self.snap_grid_var.set(not self.snap_grid_var.get()))

        # ── 通り芯ツール行 ──
        tf=ttk.LabelFrame(parent,text='通り芯',padding=3)
        tf.pack(fill=tk.X,pady=1)
        self.torii_mode_var=tk.StringVar(value='none')
        ttk.Radiobutton(tf,text='X通り芯(縦線)',variable=self.torii_mode_var,value='X').pack(side=tk.LEFT,padx=4)
        ttk.Radiobutton(tf,text='Y通り芯(横線)',variable=self.torii_mode_var,value='Y').pack(side=tk.LEFT,padx=4)
        ttk.Radiobutton(tf,text='通常',variable=self.torii_mode_var,value='none').pack(side=tk.LEFT,padx=4)
        ttk.Separator(tf,orient='vertical').pack(side=tk.LEFT,fill='y',padx=4)
        ttk.Label(tf,text='名称:').pack(side=tk.LEFT)
        self.torii_name_var=tk.StringVar(value='X1')
        ttk.Entry(tf,textvariable=self.torii_name_var,width=5).pack(side=tk.LEFT,padx=2)
        ttk.Separator(tf,orient='vertical').pack(side=tk.LEFT,fill='y',padx=4)
        ttk.Label(tf,text='基準:').pack(side=tk.LEFT)
        self.torii_ref_var=tk.StringVar(value='')
        self.torii_ref_combo=ttk.Combobox(tf,textvariable=self.torii_ref_var,width=6)
        self.torii_ref_combo.pack(side=tk.LEFT,padx=2)
        ttk.Label(tf,text='から').pack(side=tk.LEFT)
        self.torii_dist_var=tk.StringVar(value='1000')
        ttk.Entry(tf,textvariable=self.torii_dist_var,width=8).pack(side=tk.LEFT,padx=2)
        ttk.Label(tf,text='mm').pack(side=tk.LEFT)
        ttk.Button(tf,text='平行追加',command=self.add_torii_parallel).pack(side=tk.LEFT,padx=4)
        ttk.Button(tf,text='通り芯一覧・削除',command=self.show_torii_list).pack(side=tk.LEFT,padx=4)

        # ── スナップ記号配置行 ──
        pf=ttk.LabelFrame(parent,text='スナップ記号配置 (通り芯交点+オフセット)',padding=3)
        pf.pack(fill=tk.X,pady=1)
        self.snap_place_mode_var=tk.BooleanVar(value=False)
        ttk.Checkbutton(pf,text='スナップ配置ON',variable=self.snap_place_mode_var).pack(side=tk.LEFT,padx=4)
        ttk.Label(pf,text='基準交点:').pack(side=tk.LEFT)
        self.snap_ref_var=tk.StringVar(value='')
        self.snap_ref_combo=ttk.Combobox(pf,textvariable=self.snap_ref_var,width=10)
        self.snap_ref_combo.pack(side=tk.LEFT,padx=2)
        ttk.Label(pf,text='ΔX:').pack(side=tk.LEFT)
        self.snap_offset_x_var=tk.StringVar(value='0')
        ttk.Entry(pf,textvariable=self.snap_offset_x_var,width=6).pack(side=tk.LEFT)
        ttk.Label(pf,text='mm ΔY:').pack(side=tk.LEFT)
        self.snap_offset_y_var=tk.StringVar(value='0')
        ttk.Entry(pf,textvariable=self.snap_offset_y_var,width=6).pack(side=tk.LEFT)
        ttk.Label(pf,text='mm').pack(side=tk.LEFT)
        ttk.Button(pf,text='オフセット配置実行',command=self.place_symbol_at_snap_offset).pack(side=tk.LEFT,padx=4)
        ttk.Button(pf,text='交点一覧更新',command=self.refresh_intersection_list).pack(side=tk.LEFT,padx=4)
        ttk.Button(pf,text='CAD記号一覧・削除',command=self.show_cad_symbol_list).pack(side=tk.LEFT,padx=4)

        # ── 保存行 ──
        ef=ttk.LabelFrame(parent,text='通り芯+CAD記号込みで保存',padding=3)
        ef.pack(fill=tk.X,pady=1)
        ttk.Button(ef,text='PNG別名保存',command=lambda:self.export_modified_drawing('png')).pack(side=tk.LEFT,padx=4)
        ttk.Button(ef,text='PDF別名保存',command=lambda:self.export_modified_drawing('pdf')).pack(side=tk.LEFT,padx=4)
        ttk.Button(ef,text='通り芯をDB保存',command=self.save_torii_to_db).pack(side=tk.LEFT,padx=4)
        ttk.Button(ef,text='通り芯をDB読込',command=self.load_torii_from_db).pack(side=tk.LEFT,padx=4)

    def calculate_snap_position(self, canvas_x, canvas_y):
        """スナップ適用後の画像座標を返す: (img_x, img_y, snap_type, snap_label)"""
        ix,iy=self.canvas_xy_to_img(canvas_x,canvas_y)
        bx,by=float(ix),float(iy)
        snap_type='free'; snap_label=''; best_dist=float('inf')
        try: thr=float(self.snap_threshold_var.get())/max(0.001,self.preview_scale)
        except Exception: thr=15.0/max(0.001,self.preview_scale)

        if getattr(self,'snap_object_var',None) and self.snap_object_var.get():
            x_ts=[ts for ts in self.torii_shins if ts.axis=='X' and ts.page==self.current_page and ts.enabled]
            y_ts=[ts for ts in self.torii_shins if ts.axis=='Y' and ts.page==self.current_page and ts.enabled]
            # 交点スナップ（最優先）
            for xt in x_ts:
                for yt in y_ts:
                    d=math.hypot(ix-xt.img_pos,iy-yt.img_pos)
                    if d<thr*1.5 and d<best_dist:
                        bx,by=xt.img_pos,yt.img_pos; best_dist=d
                        snap_type='intersection'; snap_label=f'{xt.name}×{yt.name}'
            # 通り芯スナップ
            if snap_type=='free':
                for ts in (x_ts+y_ts):
                    if ts.axis=='X':
                        d=abs(ix-ts.img_pos)
                        if d<thr and d<best_dist: bx=ts.img_pos; best_dist=d; snap_type='torii'; snap_label=ts.name
                    else:
                        d=abs(iy-ts.img_pos)
                        if d<thr and d<best_dist: by=ts.img_pos; best_dist=d; snap_type='torii'; snap_label=ts.name

        if getattr(self,'snap_grid_var',None) and self.snap_grid_var.get() and self.scale_px_per_mm:
            try:
                gpx=float(self.snap_grid_mm_var.get())*self.scale_px_per_mm
                if gpx>0:
                    sx=round(ix/gpx)*gpx; sy=round(iy/gpx)*gpx
                    d=math.hypot(ix-sx,iy-sy)
                    if d<thr and d<best_dist:
                        bx,by=sx,sy; snap_type='grid'
                        snap_label=f'{self.snap_grid_mm_var.get()}mmグリッド'
            except Exception: pass

        return int(bx),int(by),snap_type,snap_label

    def show_snap_cursor(self, canvas_x, canvas_y):
        """スナップカーソルをキャンバスに表示"""
        for _id in self._snap_cursor_ids:
            try: self.preview_canvas.delete(_id)
            except Exception: pass
        self._snap_cursor_ids=[]
        if not self.image_pages: return
        ix,iy,stype,slabel=self.calculate_snap_position(canvas_x,canvas_y)
        cx=ix*self.preview_scale; cy=iy*self.preview_scale
        col={'intersection':'#FF6600','torii':'#0066FF','grid':'#009900','free':'#999999'}.get(stype,'#999999')
        r=8
        ids=[
            self.preview_canvas.create_rectangle(cx-r,cy-r,cx+r,cy+r,outline=col,width=2,tags='snap_cursor'),
            self.preview_canvas.create_line(cx-r*2,cy,cx+r*2,cy,fill=col,width=1,tags='snap_cursor'),
            self.preview_canvas.create_line(cx,cy-r*2,cx,cy+r*2,fill=col,width=1,tags='snap_cursor'),
        ]
        if slabel:
            ids.append(self.preview_canvas.create_text(cx+12,cy-12,text=slabel,fill=col,font=('Arial',9,'bold'),anchor='w',tags='snap_cursor'))
        self._snap_cursor_ids=ids
        try: self.snap_status_var.set(f'スナップ: {stype} [{slabel}] ({ix},{iy}px)')
        except Exception: pass

    def draw_torii_on_canvas(self):
        """通り芯とCAD記号をキャンバスにオーバーレイ描画"""
        self.preview_canvas.delete('torii_layer')
        if not self.image_pages or not self.torii_shins: return
        img=self.image_pages[self.current_page-1]
        sc=self.preview_scale
        W=img.width*sc; H=img.height*sc
        for ts in self.torii_shins:
            if ts.page!=self.current_page or not ts.enabled: continue
            pos=ts.img_pos*sc
            col=ts.color
            if ts.axis=='X':  # 縦線
                self.preview_canvas.create_line(pos,0,pos,H,fill=col,width=2,dash=(10,5),tags='torii_layer')
                self.preview_canvas.create_text(pos,16,text=ts.name,fill=col,font=('Arial',10,'bold'),anchor='n',tags='torii_layer')
                self.preview_canvas.create_oval(pos-5,0,pos+5,10,fill=col,outline=col,tags='torii_layer')
            else:  # 横線
                self.preview_canvas.create_line(0,pos,W,pos,fill=col,width=2,dash=(10,5),tags='torii_layer')
                self.preview_canvas.create_text(16,pos,text=ts.name,fill=col,font=('Arial',10,'bold'),anchor='w',tags='torii_layer')
                self.preview_canvas.create_oval(0,pos-5,10,pos+5,fill=col,outline=col,tags='torii_layer')
        # CAD記号
        for cs in self.cad_symbols_snap:
            if cs.page!=self.current_page: continue
            cx,cy=cs.img_x*sc,cs.img_y*sc
            draw_symbol_glyph_on_canvas(self.preview_canvas,cx,cy,cs.name,scale=sc,tag='torii_layer')
            if cs.ref_torii:
                self.preview_canvas.create_text(cx,cy+14,text=f'{cs.ref_torii}+({cs.offset_x_mm:.0f},{cs.offset_y_mm:.0f})mm',fill='#003366',font=('Arial',7),anchor='n',tags='torii_layer')

    def add_torii_shin_at_canvas(self, event):
        """通り芯描画モード: キャンバスクリックで通り芯を追加"""
        ix,iy,stype,slabel=self.calculate_snap_position(event.x,event.y)
        axis=getattr(self,'torii_mode_var',tk.StringVar()).get()
        if axis not in ('X','Y'): return
        name=getattr(self,'torii_name_var',tk.StringVar()).get().strip() or f'{axis}{len(self.torii_shins)+1}'
        img_pos=float(ix if axis=='X' else iy)
        ts=ToriiShin(id=len(self.torii_shins)+1,name=name,axis=axis,img_pos=img_pos,page=self.current_page)
        self.torii_shins.append(ts)
        # 次の通り芯名を自動インクリメント
        try:
            prefix=''.join(c for c in name if not c.isdigit())
            num=int(''.join(c for c in name if c.isdigit()) or '1')
            self.torii_name_var.set(f'{prefix}{num+1}')
        except Exception: pass
        self.refresh_intersection_list()
        self.draw_torii_on_canvas()
        mm_pos=img_pos/self.scale_px_per_mm if self.scale_px_per_mm else img_pos
        self.status.config(text=f'通り芯追加: {name} ({axis}軸 {mm_pos:.1f}{"mm" if self.scale_px_per_mm else "px"} スナップ={stype})')

    def add_torii_parallel(self):
        """既存の通り芯から指定距離だけ平行な通り芯を追加"""
        ref_name=getattr(self,'torii_ref_var',tk.StringVar()).get().strip()
        new_name=getattr(self,'torii_name_var',tk.StringVar()).get().strip()
        if not self.scale_px_per_mm:
            messagebox.showwarning('縮尺未設定','先に縮尺測定で縮尺を設定してください'); return
        if not ref_name:
            messagebox.showwarning('基準未選択','基準の通り芯名を選択してください'); return
        ref_ts=next((ts for ts in self.torii_shins if ts.name==ref_name),None)
        if not ref_ts:
            messagebox.showwarning('通り芯なし',f'通り芯 {ref_name} が見つかりません'); return
        try: dist_mm=float(self.torii_dist_var.get())
        except Exception: messagebox.showwarning('入力エラー','距離を数値で入力してください'); return
        if not new_name: new_name=ref_name[0]+str(len(self.torii_shins)+1)
        offset_px=dist_mm*self.scale_px_per_mm
        new_pos=ref_ts.img_pos+offset_px
        ts=ToriiShin(id=len(self.torii_shins)+1,name=new_name,axis=ref_ts.axis,img_pos=new_pos,page=ref_ts.page)
        self.torii_shins.append(ts)
        self.refresh_intersection_list()
        self.draw_torii_on_canvas()
        self.status.config(text=f'平行通り芯追加: {new_name} ({ref_name}から{dist_mm:.0f}mm 平行)')
        messagebox.showinfo('追加完了',f'{ref_name}から{dist_mm:.0f}mm平行に {new_name} を追加しました')

    def show_torii_list(self):
        """通り芯一覧ダイアログ（編集・削除）"""
        win=tk.Toplevel(self.root); win.title('通り芯一覧'); win.geometry('620x400')
        cols=('id','name','axis','pos_px','pos_mm','page','enabled')
        tree=ttk.Treeview(win,columns=cols,show='headings')
        for c,w in zip(cols,[40,60,50,80,80,50,60]):
            tree.heading(c,text={'id':'ID','name':'名称','axis':'軸','pos_px':'位置(px)','pos_mm':'位置(mm)','page':'ページ','enabled':'有効'}[c]); tree.column(c,width=w)
        for ts in self.torii_shins:
            mm=f'{ts.img_pos/self.scale_px_per_mm:.1f}' if self.scale_px_per_mm else '-'
            tree.insert('',tk.END,iid=str(ts.id),values=(ts.id,ts.name,ts.axis,f'{ts.img_pos:.1f}',mm,ts.page,'○'if ts.enabled else'×'))
        tree.pack(fill=tk.BOTH,expand=True,padx=8,pady=8)
        def do_del():
            sel=tree.selection()
            if not sel: return
            ids={int(i) for i in sel}
            self.torii_shins=[ts for ts in self.torii_shins if ts.id not in ids]
            for i in sel: tree.delete(i)
            self.draw_torii_on_canvas(); self.refresh_intersection_list()
        def do_toggle():
            sel=tree.selection()
            if not sel: return
            ids={int(i) for i in sel}
            for ts in self.torii_shins:
                if ts.id in ids: ts.enabled=not ts.enabled
            win.destroy(); self.show_torii_list()
            self.draw_torii_on_canvas()
        def do_edit():
            sel=tree.selection()
            if not sel: return
            ts_id=int(sel[0])
            ts=next((t for t in self.torii_shins if t.id==ts_id),None)
            if not ts: return
            from tkinter.simpledialog import askstring
            new_name=askstring('名称変更',f'新しい名称 (現在: {ts.name})',initialvalue=ts.name,parent=win)
            if new_name: ts.name=new_name.strip()
            self.draw_torii_on_canvas(); self.refresh_intersection_list()
            win.destroy(); self.show_torii_list()
        bf=ttk.Frame(win); bf.pack(fill=tk.X,pady=4)
        ttk.Button(bf,text='選択を削除',command=do_del).pack(side=tk.LEFT,padx=6)
        ttk.Button(bf,text='有効/無効切替',command=do_toggle).pack(side=tk.LEFT,padx=6)
        ttk.Button(bf,text='名称変更',command=do_edit).pack(side=tk.LEFT,padx=6)
        ttk.Button(bf,text='閉じる',command=win.destroy).pack(side=tk.RIGHT,padx=6)

    def refresh_intersection_list(self):
        """通り芯交点の一覧を更新してコンボボックスに設定"""
        x_ts=[ts for ts in self.torii_shins if ts.axis=='X' and ts.enabled]
        y_ts=[ts for ts in self.torii_shins if ts.axis=='Y' and ts.enabled]
        intersections=[f'{xt.name}×{yt.name}' for xt in x_ts for yt in y_ts]
        ref_names=[ts.name for ts in self.torii_shins if ts.enabled]
        try:
            if hasattr(self,'torii_ref_combo'): self.torii_ref_combo.config(values=ref_names)
            if hasattr(self,'snap_ref_combo'): self.snap_ref_combo.config(values=intersections)
        except Exception: pass

    def get_intersection_pos(self, name):
        """交点名 'X1×Y1' から画像ピクセル座標を返す: (ix, iy) or None"""
        parts=name.split('×')
        if len(parts)!=2: return None
        xn,yn=parts[0].strip(),parts[1].strip()
        xt=next((ts for ts in self.torii_shins if ts.name==xn and ts.axis=='X'),None)
        yt=next((ts for ts in self.torii_shins if ts.name==yn and ts.axis=='Y'),None)
        if xt is None or yt is None: return None
        return (xt.img_pos, yt.img_pos)

    def place_symbol_at_snap_offset(self):
        """通り芯交点+オフセットで図面記号を配置"""
        if not self.scale_px_per_mm:
            messagebox.showwarning('縮尺未設定','先に縮尺測定で縮尺を設定してください'); return
        ref=getattr(self,'snap_ref_var',tk.StringVar()).get().strip()
        if not ref:
            messagebox.showwarning('未選択','基準交点を選択してください'); return
        pos=self.get_intersection_pos(ref)
        if pos is None:
            messagebox.showwarning('交点なし',f'交点 {ref} が見つかりません'); return
        ix0,iy0=pos
        try:
            dx_mm=float(self.snap_offset_x_var.get())
            dy_mm=float(self.snap_offset_y_var.get())
        except Exception:
            messagebox.showwarning('入力エラー','ΔX/ΔYを数値で入力してください'); return
        ix=ix0+dx_mm*self.scale_px_per_mm
        iy=iy0+dy_mm*self.scale_px_per_mm
        name=getattr(self,'symbol_paste_var',tk.StringVar(value='LEDダウンライト')).get().strip()
        cs=CADSymbol(id=len(self.cad_symbols_snap)+1,name=name,page=self.current_page,
                     img_x=ix,img_y=iy,ref_torii=ref,offset_x_mm=dx_mm,offset_y_mm=dy_mm)
        self.cad_symbols_snap.append(cs)
        # image_detsにも追加して検出候補・積算に反映
        size=35
        img=self.image_pages[self.current_page-1]
        det=ImageDetection(len(self.image_dets)+1,self.current_page,
            max(0,int(ix)-size),max(0,int(iy)-size),
            min(img.width-1,int(ix)+size),min(img.height-1,int(iy)+size),
            'symbol',name,1.0,0,True,'cad_snap',
            f'CAD snap: {ref}+({dx_mm},{dy_mm})mm')
        self.image_dets.append(det)
        self.draw_torii_on_canvas()
        self.refresh_image_tables()
        self.status.config(text=f'スナップ配置: {name} @ {ref}+({dx_mm:.0f},{dy_mm:.0f})mm')

    def place_cad_symbol_at_click(self, event):
        """スナップ配置モード: クリック位置にスナップして記号を配置"""
        ix,iy,stype,slabel=self.calculate_snap_position(event.x,event.y)
        name=getattr(self,'symbol_paste_var',tk.StringVar(value='LEDダウンライト')).get().strip()
        ref_label=slabel if stype=='intersection' else ''
        cs=CADSymbol(id=len(self.cad_symbols_snap)+1,name=name,page=self.current_page,
                     img_x=float(ix),img_y=float(iy),ref_torii=ref_label)
        self.cad_symbols_snap.append(cs)
        size=35; img=self.image_pages[self.current_page-1]
        det=ImageDetection(len(self.image_dets)+1,self.current_page,
            max(0,ix-size),max(0,iy-size),min(img.width-1,ix+size),min(img.height-1,iy+size),
            'symbol',name,1.0,0,True,'cad_snap',f'cad snap click {stype}:{slabel}')
        self.image_dets.append(det)
        self.draw_torii_on_canvas(); self.refresh_image_tables()
        self.status.config(text=f'スナップ配置: {name} [{stype}:{slabel}]')

    def show_cad_symbol_list(self):
        """CAD配置記号の一覧・削除ダイアログ"""
        win=tk.Toplevel(self.root); win.title('CAD配置記号一覧'); win.geometry('700x380')
        cols=('id','name','ref','dx','dy','page')
        tree=ttk.Treeview(win,columns=cols,show='headings')
        for c,w in zip(cols,[40,120,100,70,70,50]):
            tree.heading(c,text={'id':'ID','name':'記号名','ref':'基準交点','dx':'ΔX(mm)','dy':'ΔY(mm)','page':'ページ'}[c]); tree.column(c,width=w)
        for cs in self.cad_symbols_snap:
            tree.insert('',tk.END,iid=str(cs.id),values=(cs.id,cs.name,cs.ref_torii,f'{cs.offset_x_mm:.1f}',f'{cs.offset_y_mm:.1f}',cs.page))
        tree.pack(fill=tk.BOTH,expand=True,padx=8,pady=8)
        def do_del():
            sel=tree.selection()
            if not sel: return
            ids={int(i) for i in sel}
            self.cad_symbols_snap=[cs for cs in self.cad_symbols_snap if cs.id not in ids]
            self.image_dets=[d for d in self.image_dets if not (d.source=='cad_snap' and d.det_id in ids)]
            for i in sel: tree.delete(i)
            self.draw_torii_on_canvas(); self.refresh_image_tables()
        bf=ttk.Frame(win); bf.pack(fill=tk.X,pady=4)
        ttk.Button(bf,text='選択を削除',command=do_del).pack(side=tk.LEFT,padx=6)
        ttk.Button(bf,text='閉じる',command=win.destroy).pack(side=tk.RIGHT,padx=6)

    def export_modified_drawing(self, fmt='png'):
        """通り芯・CAD記号を画像に焼き込んで別名保存"""
        if not self.image_pages:
            messagebox.showwarning('なし','先に画像/PDFを読み込んでください'); return
        from PIL import ImageDraw, ImageFont
        from tkinter.filedialog import asksaveasfilename
        ext={'png':'PNG','pdf':'PDF'}.get(fmt,'PNG')
        path=asksaveasfilename(title='別名保存',defaultextension=f'.{fmt}',
            filetypes=[(f'{ext}ファイル',f'*.{fmt}'),('全ファイル','*.*')],
            initialfile=f'modified_drawing.{fmt}')
        if not path: return
        try:
            pages_out=[]
            for pg_idx,img in enumerate(self.image_pages,1):
                out=img.copy().convert('RGB')
                draw=ImageDraw.Draw(out)
                W,H=out.size
                try: font_big=ImageFont.truetype('arial.ttf',16)
                except Exception:
                    try: font_big=ImageFont.truetype('/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf',16)
                    except Exception: font_big=ImageFont.load_default()
                # 通り芯を描画
                for ts in self.torii_shins:
                    if ts.page!=pg_idx or not ts.enabled: continue
                    col=ts.color if ts.color.startswith('#') else '#CC0000'
                    # PILは'#RRGGBB'形式を受け付ける
                    if ts.axis=='X':  # 縦線
                        for dash_y in range(0,H,20):
                            draw.line([(ts.img_pos,dash_y),(ts.img_pos,min(H,dash_y+12))],fill=col,width=2)
                        draw.text((ts.img_pos+4,4),ts.name,fill=col,font=font_big)
                    else:  # 横線
                        for dash_x in range(0,W,20):
                            draw.line([(dash_x,ts.img_pos),(min(W,dash_x+12),ts.img_pos)],fill=col,width=2)
                        draw.text((4,ts.img_pos+2),ts.name,fill=col,font=font_big)
                # CAD記号を描画（シンプルな丸+名称）
                for cs in self.cad_symbols_snap:
                    if cs.page!=pg_idx: continue
                    r=16; x,y=int(cs.img_x),int(cs.img_y)
                    draw.ellipse([(x-r,y-r),(x+r,y+r)],outline='#003366',width=2)
                    draw.line([(x-r//2,y),(x+r//2,y)],fill='#003366',width=2)
                    draw.line([(x,y-r//2),(x,y+r//2)],fill='#003366',width=2)
                    draw.text((x+r+2,y-8),cs.name,fill='#003366',font=font_big)
                pages_out.append(out)
            if fmt=='pdf' and len(pages_out)>0:
                pages_out[0].save(path,save_all=True,append_images=pages_out[1:] if len(pages_out)>1 else [])
            else:
                if len(pages_out)==1: pages_out[0].save(path)
                else:
                    for i,p in enumerate(pages_out,1):
                        p.save(path.replace(f'.{fmt}',f'_p{i}.{fmt}'))
            messagebox.showinfo('保存完了',f'保存しました: {path}')
        except Exception:
            messagebox.showerror('保存エラー',traceback.format_exc()[-1200:])

    def save_torii_to_db(self):
        """通り芯データをDBに保存"""
        try:
            data=json.dumps([{'id':ts.id,'name':ts.name,'axis':ts.axis,'img_pos':ts.img_pos,
                              'page':ts.page,'color':ts.color,'enabled':ts.enabled}
                             for ts in self.torii_shins],ensure_ascii=False)
            self.db.set_setting('torii_shins',data)
            cad_data=json.dumps([{'id':cs.id,'name':cs.name,'page':cs.page,'img_x':cs.img_x,'img_y':cs.img_y,
                                   'ref_torii':cs.ref_torii,'offset_x_mm':cs.offset_x_mm,'offset_y_mm':cs.offset_y_mm}
                                  for cs in self.cad_symbols_snap],ensure_ascii=False)
            self.db.set_setting('cad_symbols_snap',cad_data)
            self.status.config(text=f'通り芯{len(self.torii_shins)}件をDB保存しました')
        except Exception:
            messagebox.showerror('保存エラー',traceback.format_exc()[-800:])

    def load_torii_from_db(self):
        """通り芯データをDBから読込"""
        try:
            data=self.db.get_setting('torii_shins','[]')
            rows=json.loads(data)
            self.torii_shins=[ToriiShin(**r) for r in rows]
            cad_data=self.db.get_setting('cad_symbols_snap','[]')
            cad_rows=json.loads(cad_data)
            self.cad_symbols_snap=[CADSymbol(**r) for r in cad_rows]
            self.refresh_intersection_list()
            self.draw_torii_on_canvas()
            self.status.config(text=f'通り芯{len(self.torii_shins)}件をDB読込しました')
        except Exception:
            messagebox.showerror('読込エラー',traceback.format_exc()[-800:])

    def run_image_ai_estimate_async(self):
        """画像PDF解析タブのAI積算: 画像検出結果をコンテキストにAIへ積算依頼"""
        rows=self.image_estimate_rows()
        if not rows:
            messagebox.showwarning('なし','画像積算結果がありません。\n軽量解析を実行してから使用してください。'); return
        snap=self.snapshot_ai_settings(); p=self.image_file_var.get().strip()
        enabled_dets=[d for d in self.image_dets if d.enabled and d.equipment!='未分類設備']
        def worker():
            counts={}
            for r in rows: counts[r[1]]=counts.get(r[1],0)+r[3]
            prompt=(
                'あなたは日本の電気設備積算の専門家AIです。\n'
                '画像PDF図面を解析した結果、以下の設備が検出されました。\n'
                '過検出・漏れ・積算上の注意点・見積留意事項を日本語で指摘してください。\n'
                '（補助ツールです。最終積算は専門家が確認してください）\n\n'
                f'対象ファイル: {Path(p).name if p else "不明"}\n'
                f'検出設備数量: {json.dumps(counts,ensure_ascii=False)}\n'
                f'積算明細: {json.dumps(rows,ensure_ascii=False)}\n'
                f'有効検出候補数: {len(enabled_dets)}件\n'
            )
            return str(UnifiedAIClient(snap).generate(prompt,timeout=90))
        def done(ans):
            self.img_ai_status_var.set(f'AI積算完了')
            try:
                win=tk.Toplevel(self.root); win.title('AI積算結果（画像解析）'); win.geometry('680x480')
                txt=scrolledtext.ScrolledText(win,wrap=tk.WORD,font=('Arial',10)); txt.pack(fill=tk.BOTH,expand=True,padx=8,pady=8)
                txt.insert('1.0',ans); txt.configure(state='disabled')
                ttk.Button(win,text='閉じる',command=win.destroy).pack(pady=4)
            except Exception: messagebox.showinfo('AI積算結果',ans[:2000])
        def err(tb):
            self.log(tb); messagebox.showerror('AI積算エラー',tb[-1200:])
        self.run_in_thread(worker,done,err)
        self.img_ai_status_var.set('AI積算中...')

    def on_click_ai_estimate(self):
        p=self.file_path_var.get().strip()
        if not p:
            messagebox.showwarning('未選択','PDFまたはDXFを選択してください')
            return
        snap=self.snapshot_ai_settings()

        def worker():
            ext=Path(p).suffix.lower()
            counts={}
            details=''
            if ext=='.pdf':
                text_pdf,method=extract_pdf_text(p)
                counts=extract_counts_from_text(text_pdf,self.db)
                details=f'PDF text method={method}, chars={len(text_pdf)}'
            elif ext=='.dxf':
                parser=AdvancedDXFParser()
                parser.parse_dxf(p)
                rules=[]
                for r in self.db.get_all_symbol_patterns():
                    try:
                        rules.append({'name':r[1],'pattern':json.loads(r[3])})
                    except Exception:
                        pass
                counts=parser.count_by_patterns(rules)
                tcounts=extract_counts_from_text('\n'.join(parser.texts),self.db)
                for k,v in tcounts.items():
                    counts[k]=counts.get(k,0)+v
                details=json.dumps(parser.signature(),ensure_ascii=False)
            else:
                details='未対応形式'
            rows=self.make_estimate_rows(counts)
            mem_ctx=''
            try:
                mem_ctx=LongTermMemoryManager(self.db).context(
                    f'積算 {Path(p).name} {json.dumps(counts,ensure_ascii=False)} {details}',
                    namespace='estimation',
                    top_k=int(snap.get('memory_top_k','6')) if snap.get('use_langmem','1')=='1' else 0
                )
            except Exception:
                mem_ctx=''
            prompt=(
                'あなたは日本の電気設備積算の補助AIです。\n'
                '以下のPDF/DXF解析結果をもとに、過検出・不足・見積上の注意点を簡潔に指摘してください。\n'
                'なお、これは法的・最終見積ではなく、積算補助です。\n\n'
                f'対象ファイル: {p}\n'
                f'解析詳細: {details}\n'
                f'数量候補: {json.dumps(counts,ensure_ascii=False)}\n'
                f'積算行: {json.dumps(rows,ensure_ascii=False)}\n'
                f'{mem_ctx}\n'
            )
            client=UnifiedAIClient(snap)
            ai_ans=client.generate(prompt,timeout=180)
            usage=getattr(client,'last_usage',{})
            return rows, ai_ans, usage

        def done(res):
            rows, ai_ans, usage = res
            self.result_rows=rows
            self.update_result_tree(rows)
            try:
                self.ai_result.insert(tk.END,'\n--- AI積算結果 ---\n'+ai_ans+'\n')
                self.ai_result.see(tk.END)
            except Exception:
                pass
            try:
                if usage and hasattr(self,'update_ai_usage_display'):
                    self.update_ai_usage_display(usage)
            except Exception:
                pass
            self.status.config(text='AI積算完了')
            try:
                self.record_memory('estimation','ai_estimate_review',f'AI積算レビュー: file={Path(p).name}\n{str(ai_ans)[:1000]}',tags='ai_estimate,review',source=p,weight=1.0)
            except Exception:
                pass

        def err(tb):
            self.log(tb)
            messagebox.showerror('AI積算エラー', tb[-1600:])

        self.run_in_thread(worker, done, err)


    def on_click_estimate(self):
        p=self.file_path_var.get().strip();
        if p: threading.Thread(target=self.estimate_worker,args=(p,),daemon=True).start()
    def estimate_worker(self,p):
        try:
            self.ui(lambda:self.progress.start(10)); self.log(f'解析開始: {p}'); ext=Path(p).suffix.lower(); counts={}
            if ext=='.pdf':
                text,method=extract_pdf_text(p); self.log(f'PDF text method={method}, chars={len(text)}')
                if text.strip(): counts=extract_counts_from_text(text,self.db)
                else: self.log('PDFテキスト抽出不可。画像PDF解析タブで解析してください。')
            elif ext=='.dxf':
                parser=AdvancedDXFParser(); parser.parse_dxf(p); self.log('DXF解析: '+json.dumps(parser.signature(),ensure_ascii=False)[:1000]); rules=[]
                for r in self.db.get_all_symbol_patterns():
                    try: rules.append({'name':r[1],'pattern':json.loads(r[3])})
                    except Exception: pass
                counts=parser.count_by_patterns(rules); tcounts=extract_counts_from_text('\\n'.join(parser.texts),self.db)
                for k,v in tcounts.items(): counts[k]=counts.get(k,0)+v
            rows=self.make_estimate_rows(counts); self.result_rows=rows; self.ui(self.update_result_tree,rows); self.log('積算完了')
        except Exception: self.log(traceback.format_exc())
        finally: self.ui(lambda:self.progress.stop())
    def make_estimate_rows(self,counts):
        rows=[]
        for name,qty in sorted(counts.items()):
            cat,item,spec,unit,price=self.db.find_price(name); rows.append((cat,item,spec,unit,qty,int(price),int(float(price)*qty)))
        return rows
    def update_result_tree(self,rows):
        for i in self.tree.get_children(): self.tree.delete(i)
        for r in rows: self.tree.insert('',tk.END,values=r)
    def export_result_csv(self):
        rows=self.result_rows or [self.tree.item(i,'values') for i in self.tree.get_children()]
        if not rows: messagebox.showwarning('なし','積算結果がありません'); return
        p=filedialog.asksaveasfilename(defaultextension='.csv',filetypes=[('CSV','*.csv')])
        if not p: return
        with open(p,'w',encoding='utf-8-sig',newline='') as f: w=csv.writer(f); w.writerow(['カテゴリ','品名','仕様','単位','数量','単価','金額']); w.writerows(rows)
        messagebox.showinfo('保存',p)
    def on_click_pdf_split_to_dxf(self):
        p=self.file_path_var.get().strip()
        if not p or Path(p).suffix.lower()!='.pdf': messagebox.showwarning('PDF選択','PDFを選択してください'); return
        out=Path(p).with_suffix('.pseudo_layers.dxf'); out.write_text('0\nSECTION\n2\nENTITIES\n0\nENDSEC\n0\nEOF\n',encoding='ascii'); messagebox.showinfo('保存',f'簡易DXFを保存しました:\n{out}')
    def price_refresh(self):
        for i in self.price_tree.get_children(): self.price_tree.delete(i)
        for r in self.db.get_all_unit_prices(): self.price_tree.insert('',tk.END,values=r)
    def selected_price_record(self):
        sel=self.price_tree.selection(); return self.price_tree.item(sel[0],'values') if sel else None
    def price_add(self): d=UnitPriceDialog(self.root,self.db); self.root.wait_window(d); self.price_refresh()
    def price_edit(self):
        rec=self.selected_price_record()
        if rec: rec=list(rec); rec[0]=int(rec[0]); rec[5]=float(rec[5]); d=UnitPriceDialog(self.root,self.db,rec); self.root.wait_window(d); self.price_refresh()
    def price_delete(self):
        rec=self.selected_price_record()
        if rec and messagebox.askyesno('確認','削除しますか？'): self.db.delete_unit_price(int(rec[0])); self.price_refresh()
    def price_import_csv(self):
        p=filedialog.askopenfilename(filetypes=[('CSV','*.csv')])
        if p: a,b=self.db.import_unit_csv(p); self.price_refresh(); messagebox.showinfo('完了',f'取込 {a} / スキップ {b}')
    def price_export_csv(self):
        p=filedialog.asksaveasfilename(defaultextension='.csv',filetypes=[('CSV','*.csv')])
        if p: n=self.db.export_unit_csv(p); messagebox.showinfo('完了',f'{n}件保存')
    def symbol_refresh(self):
        for i in self.sym_tree.get_children(): self.sym_tree.delete(i)
        for r in self.db.get_all_symbol_patterns(): self.sym_tree.insert('',tk.END,values=r)
    def selected_symbol_record(self):
        sel=self.sym_tree.selection(); return self.sym_tree.item(sel[0],'values') if sel else None
    def symbol_add(self): d=SymbolPatternDialog(self.root,self.db); self.root.wait_window(d); self.symbol_refresh()
    def symbol_edit(self):
        rec=self.selected_symbol_record()
        if rec: rec=list(rec); rec[0]=int(rec[0]); rec[5]=int(rec[5]); d=SymbolPatternDialog(self.root,self.db,rec); self.root.wait_window(d); self.symbol_refresh()
    def symbol_delete(self):
        rec=self.selected_symbol_record()
        if rec and int(rec[4])==0 and messagebox.askyesno('確認','削除しますか？'): self.db.delete_symbol_pattern(int(rec[0])); self.symbol_refresh()
    def cad_refresh(self):
        for i in self.cad_tree.get_children(): self.cad_tree.delete(i)
        for r in self.db.get_cad_library(): self.cad_tree.insert('',tk.END,values=(r[0],r[1],r[2],r[3],r[4],os.path.basename(r[5] or '')))
    def cad_import_local(self):
        p=filedialog.askopenfilename(filetypes=[('DXF','*.dxf'),('All','*.*')])
        if not p: return
        dst=str(Path(CAD_LIBRARY_PATH)/Path(p).name); shutil.copy2(p,dst); d=CADRegisterDialog(self.root,self.db,dst); self.root.wait_window(d); self.cad_refresh()
    def cad_download_from_url(self):
        url=self.cad_url_var.get().strip()
        if url and url!='https://': threading.Thread(target=self.cad_download_worker,args=(url,),daemon=True).start()
    def cad_download_worker(self,url):
        try:
            requests=try_import_requests();
            if not requests: raise RuntimeError('requestsが必要です')
            self.log(f'CADダウンロード: {url}'); r=requests.get(url,headers={'User-Agent':'Mozilla/5.0'},timeout=120); r.raise_for_status(); files=[]
            try:
                import io; z=zipfile.ZipFile(io.BytesIO(r.content))
                for m in z.namelist():
                    if m.lower().endswith(('.dxf','.dwg')):
                        name=re.sub(r'[<>:"/\\|?*]','_',os.path.basename(m)) or f'cad_{len(files)}.dxf'; out=str(Path(CAD_LIBRARY_PATH)/name)
                        with z.open(m) as src, open(out,'wb') as dst: shutil.copyfileobj(src,dst)
                        files.append(out)
            except zipfile.BadZipFile:
                name=os.path.basename(url.split('?')[0]) or 'downloaded.dxf'; out=str(Path(CAD_LIBRARY_PATH)/name); Path(out).write_bytes(r.content); files.append(out)
            self.log(f'取得 {len(files)}件'); self.ui(self.cad_refresh)
        except Exception: self.log(traceback.format_exc())
    def cad_delete(self):
        sel=self.cad_tree.selection()
        if sel and messagebox.askyesno('確認','削除しますか？'): self.db.delete_cad_file(int(self.cad_tree.item(sel[0],'values')[0])); self.cad_refresh()
    def cad_open_folder(self): os.startfile(CAD_LIBRARY_PATH) if os.name=='nt' else messagebox.showinfo('フォルダ',CAD_LIBRARY_PATH)

    # image functions
    def select_image_file(self):
        p=filedialog.askopenfilename(filetypes=[('PDF/画像/DXF','*.pdf;*.png;*.jpg;*.jpeg;*.bmp;*.webp;*.dxf'),('DXF','*.dxf'),('All','*.*')])
        if p: self.image_file_var.set(p); self.source_image_file=p
    def start_image_analysis(self):
        p=self.image_file_var.get().strip()
        try:
            self.import_symbol_crops_dataset(False)
        except Exception:
            pass
        if p: self.source_image_file=p; threading.Thread(target=self.image_analysis_worker,args=(p,),daemon=True).start()
    def edit_selected_detection(self):
        d=self.get_image_det(self.selected_det_id)
        if not d:
            messagebox.showwarning('未選択','編集する検出候補を選択してください')
            return
        dlg=ImageDetectionEditDialog(self.root,d)
        self.root.wait_window(dlg)
        if not dlg.result: return
        r=dlg.result
        d.enabled=bool(r['enabled']); d.equipment=r['equipment']; d.color_name=r['color_name']
        d.x1=max(0,r['x1']); d.y1=max(0,r['y1']); d.x2=max(d.x1,r['x2']); d.y2=max(d.y1,r['y2'])
        d.score=r['score']; d.memo=r['memo'] or 'manual table edit'; d.source='manual_edit'
        self.refresh_image_view(); self.refresh_image_tables()
        self.status.config(text=f'検出候補を編集: ID={d.det_id} {d.equipment}')

    def delete_selected_detection(self):
        d=self.get_image_det(self.selected_det_id)
        if not d:
            messagebox.showwarning('未選択','削除する検出候補を選択してください')
            return
        if not messagebox.askyesno('確認',f'検出候補 ID={d.det_id} を削除しますか？'):
            return
        self.image_dets=[x for x in self.image_dets if x.det_id!=d.det_id]
        for i,x in enumerate(self.image_dets,1): x.det_id=i
        self.selected_det_id=None
        self.refresh_image_view(); self.refresh_image_tables()
        self.status.config(text='検出候補を削除しました')

    def delete_disabled_detections(self):
        n0=len(self.image_dets)
        self.image_dets=[d for d in self.image_dets if d.enabled]
        for i,d in enumerate(self.image_dets,1): d.det_id=i
        self.selected_det_id=None
        self.refresh_image_view(); self.refresh_image_tables()
        messagebox.showinfo('削除完了',f'無効候補を {n0-len(self.image_dets)} 件削除しました')

    def add_manual_detection(self):
        if not self.image_pages:
            messagebox.showwarning('未読込','先にPDF/画像/DXFを軽量解析してください')
            return
        img=self.image_pages[self.current_page-1]; cx,cy=img.width//2,img.height//2
        det=ImageDetection(len(self.image_dets)+1,self.current_page,cx-25,cy-25,cx+25,cy+25,'manual','未分類設備',0.5,0,True,'manual_add','manual added from table')
        self.image_dets.append(det); self.select_image_det(det.det_id)
        self.refresh_image_view(); self.refresh_image_tables(); self.edit_selected_detection()

    def edit_selected_image_sum_row(self):
        sel=self.img_sum_tree.selection()
        if not sel:
            messagebox.showwarning('未選択','編集する積算行を選択してください')
            return
        vals=list(self.img_sum_tree.item(sel[0],'values'))
        if vals and vals[0]=='合計':
            messagebox.showinfo('対象外','合計行は編集できません')
            return
        dlg=EstimateRowEditDialog(self.root, vals)
        self.root.wait_window(dlg)
        if not dlg.result: return
        if messagebox.askyesno('反映方法','この編集を手動積算行として固定しますか？\n「はい」=手動行として追加\n「いいえ」=表示行だけ更新'):
            self.manual_image_estimate_rows.append(dlg.result); self.refresh_image_tables()
        else:
            self.img_sum_tree.item(sel[0],values=dlg.result)

    def add_image_sum_row(self):
        dlg=EstimateRowEditDialog(self.root, ('その他','手動追加','個',1,0,0))
        self.root.wait_window(dlg)
        if dlg.result:
            self.manual_image_estimate_rows.append(dlg.result); self.refresh_image_tables()

    def delete_selected_image_sum_row(self):
        sel=self.img_sum_tree.selection()
        if not sel:
            messagebox.showwarning('未選択','削除する積算行を選択してください')
            return
        vals=list(self.img_sum_tree.item(sel[0],'values'))
        if vals and vals[0]=='合計': return
        try:
            for i,r in enumerate(getattr(self,'manual_image_estimate_rows',[])):
                if str(r[1])==str(vals[1]) and str(r[3])==str(vals[3]):
                    del self.manual_image_estimate_rows[i]; break
        except Exception: pass
        self.img_sum_tree.delete(sel[0]); self.status.config(text='積算行を削除しました')


    def import_symbol_crops_to_vector_db(self, show_message=True):
        """
        symbol_crops内のラベル付き画像をCLIP/FAISSベクトルDBへ登録。
        ファイル名からラベル推定できない画像は誤学習防止のためスキップ。
        """
        added=0; skipped=0; unlabeled=0; backend_count={}
        try:
            self.vector_engine=SymbolVectorEngine(self.db, self.db.get_setting('clip_model_name',CLIP_MODEL_NAME_DEFAULT), self.db.get_setting('vector_backend','auto'), self.log)
            for cp in sorted(CROP_DIR.glob("*")):
                if cp.suffix.lower() not in ('.png','.jpg','.jpeg','.bmp'):
                    continue
                label=infer_label_from_crop_filename(cp)
                if not label:
                    unlabeled+=1; continue
                if hasattr(self.db,'vector_crop_exists') and self.db.vector_crop_exists(str(cp), label):
                    skipped+=1; continue
                try:
                    backend,dim=self.vector_engine.add_crop(str(cp),label,source_file=str(cp),meta={'source':'symbol_crops','dim':dim if 'dim' in locals() else None})
                    backend_count[backend]=backend_count.get(backend,0)+1
                    added+=1
                except Exception:
                    skipped+=1
            n,idx_backend=self.vector_engine.rebuild_from_db()
            self.vector_engine.save_faiss_index()
            self.refresh_learning_summary()
            msg=f'ベクトルDB取込: 追加{added}件 / 既存・失敗{skipped}件 / ラベル不明{unlabeled}件 / index={n}件({idx_backend})'
            self.status.config(text=msg)
            self.log(msg)
            if show_message:
                messagebox.showinfo('FAISS/CLIP取込', msg + '\n\nbackend内訳: ' + json.dumps(backend_count,ensure_ascii=False))
            return added
        except Exception:
            tb=traceback.format_exc()
            self.log(tb)
            if show_message:
                messagebox.showerror('FAISS/CLIP取込エラー',tb[-2000:])
            return 0

    def rebuild_vector_index_dialog(self):
        def worker():
            self.vector_engine=SymbolVectorEngine(self.db, self.db.get_setting('clip_model_name',CLIP_MODEL_NAME_DEFAULT), self.db.get_setting('vector_backend','auto'), self.log)
            n,backend=self.vector_engine.rebuild_from_db()
            saved=self.vector_engine.save_faiss_index()
            return n,backend,saved
        def done(res):
            n,backend,saved=res
            self.refresh_learning_summary()
            messagebox.showinfo('ベクトル索引再構築',f'件数: {n}\nbackend: {backend}\nFAISS保存: {saved}')
        def err(tb):
            messagebox.showerror('ベクトル索引エラー',tb[-2000:])
        self.run_in_thread(worker,done,err)

    def vector_predict_detection(self, d, crop=None):
        """
        1候補に対してベクトル検索でラベル推定。
        threshold未満ならNone。
        """
        try:
            if self.db.get_setting('use_vector_search','1')!='1':
                return None,0,[]
            if crop is None:
                crop=crop_symbol(self.image_pages[d.page-1],d,margin=8)
            top_k=int(self.db.get_setting('vector_top_k','5'))
            threshold=float(self.db.get_setting('vector_threshold','0.72'))
            results,backend=self.vector_engine.search_image(crop, top_k=top_k)
            if not results:
                return None,0,[]
            best=results[0]
            label=best.get('label')
            score=float(best.get('score') or 0)
            bbox=f'{d.x1},{d.y1},{d.x2},{d.y2}'
            self.db.save_vector_search_event(self.source_image_file,bbox,label,score,results,backend)
            if label and score>=threshold:
                return label,score,results
            return None,score,results
        except Exception:
            self.log('vector_predict_detection failed:\n'+traceback.format_exc())
            return None,0,[]

    def apply_vector_search_to_detections(self, show_message=False):
        applied=0; best_score=0.0
        try:
            self.vector_engine=SymbolVectorEngine(self.db, self.db.get_setting('clip_model_name',CLIP_MODEL_NAME_DEFAULT), self.db.get_setting('vector_backend','auto'), self.log)
            self.vector_engine.load_faiss_index()
            self.vector_engine.rebuild_from_db()
            for d in self.image_dets:
                if not getattr(d,'enabled',True): continue
                crop=crop_symbol(self.image_pages[d.page-1],d,margin=8)
                label,score,top=self.vector_predict_detection(d,crop)
                best_score=max(best_score,float(score or 0))
                if label:
                    old=d.equipment
                    d.equipment=label
                    d.source='vector'
                    d.memo=f'FAISS/CLIP vector match {score:.3f} old={old}'
                    d.score=max(float(getattr(d,'score',0) or 0),float(score))
                    applied+=1
            if applied:
                self.refresh_image_view(); self.refresh_image_tables()
            msg=f'ベクトル検索反映: {applied}件 / best={best_score:.3f}'
            self.status.config(text=msg); self.log(msg)
            if show_message:
                messagebox.showinfo('ベクトル検索反映',msg)
            return applied
        except Exception:
            tb=traceback.format_exc()
            self.log(tb)
            if show_message: messagebox.showerror('ベクトル検索エラー',tb[-2000:])
            return 0

    def import_symbol_crops_dataset(self, show_message=True):
        """
        symbol_cropsを特徴量DBへ取り込む。
        ラベル解決の優先順位:
        1. annotation_features DBに同じcrop_pathで登録済みのfinal_answer
        2. ファイル名からの推定（downlight / dl / コンセント 等）
        3. どちらでも不明 → 誤学習防止のためスキップ
        """
        added=0; skipped=0; unlabeled=0
        try:
            # DBに記録済みの crop_path→label マッピングを先に構築
            db_labels = {}
            try:
                for row in self.db.get_annotation_features(limit=100000):
                    cp_key = str(row.get('crop_path',''))
                    la = str(row.get('final_answer','') or '').strip()
                    if cp_key and la and la not in ('未分類設備',''):
                        db_labels[cp_key] = la
            except Exception:
                pass

            for cp in sorted(CROP_DIR.glob("*")):
                if cp.suffix.lower() not in ('.png','.jpg','.jpeg','.bmp'):
                    continue
                try:
                    str_cp = str(cp)
                    # ① DBに既存のラベルがある場合はそちらを使い特徴量の重複登録のみ避ける
                    if str_cp in db_labels and hasattr(self.db,'crop_path_exists_in_features') and self.db.crop_path_exists_in_features(str_cp):
                        skipped+=1; continue
                    # ② ラベル解決: DB優先 → ファイル名推定
                    label = db_labels.get(str_cp) or infer_label_from_crop_filename(cp)
                    if not label:
                        unlabeled+=1; continue
                    # ③ 特徴量抽出してDB登録
                    img=Image.open(cp).convert('RGB')
                    feat=extract_symbol_feature(img)
                    dummy=ImageDetection(0,1,0,0,img.width-1,img.height-1,'crop',label,1.0,0,True,'crop_import',f'label={label}')
                    self.db.save_annotation_feature(str_cp,dummy,label,str_cp,json.dumps(feat,ensure_ascii=False),f'imported label={label}')
                    if hasattr(self.db,'save_symbol_image_dataset'):
                        self.db.save_symbol_image_dataset(label,str_cp,str_cp,json.dumps(feat,ensure_ascii=False),f'imported label={label}')
                    added+=1
                except Exception:
                    skipped+=1

            self.refresh_learning_summary()
            self.status.config(text=f'symbol_crops取込: DB参照+ファイル名推定 追加{added}件 / 既存{skipped}件 / ラベル不明スキップ{unlabeled}件')
            if show_message:
                messagebox.showinfo('symbol_crops取込',
                    f'追加: {added}件\n既存・失敗: {skipped}件\nラベル不明スキップ: {unlabeled}件\n\n'
                    f'取込方法:\n'
                    f'1. OK/NG登録済みクロップ: DBから正解ラベルを取得\n'
                    f'2. ファイル名に設備名を含むもの: ファイル名から推定\n'
                    f'3. それ以外: 誤学習防止のためスキップ')
            return added
        except Exception:
            tb=traceback.format_exc()
            self.log(tb)
            if show_message:
                messagebox.showerror('symbol_crops取込エラー',tb[-1600:])
            return 0


    def register_symbol_image_dataset(self):
        dlg=SymbolImageDatasetDialog(self.root)
        self.root.wait_window(dlg)
        if not dlg.result: return
        try:
            pages=load_pdf_or_image(dlg.result['file'],1.0,self.log)
            img=pages[0]
            crop_path=str(CROP_DIR/f"symbol_dataset_{datetime.now().strftime('%Y%m%d_%H%M%S')}_{dlg.result['name']}.png")
            img.save(crop_path)
            feat=extract_symbol_feature(img)
            self.db.save_symbol_image_dataset(dlg.result['name'],dlg.result['file'],crop_path,json.dumps(feat,ensure_ascii=False),dlg.result.get('memo',''))
            dummy=ImageDetection(0,1,0,0,img.width-1,img.height-1,'dataset',dlg.result['name'],1.0,0,True,'dataset','symbol image dataset')
            self.db.save_annotation_feature(dlg.result['file'],dummy,dlg.result['name'],crop_path,json.dumps(feat,ensure_ascii=False),'symbol image dataset')
            try:
                self.vector_engine=SymbolVectorEngine(self.db,self.db.get_setting('clip_model_name',CLIP_MODEL_NAME_DEFAULT),self.db.get_setting('vector_backend','auto'),self.log)
                self.vector_engine.add_crop(crop_path,dlg.result['name'],source_file=dlg.result['file'],meta={'source':'manual_symbol_dataset'})
                self.vector_engine.rebuild_from_db()
                self.vector_engine.save_faiss_index()
            except Exception:
                self.log('vector add failed:\n'+traceback.format_exc())
            self.refresh_learning_summary()
            messagebox.showinfo('登録完了',f"記号画像データセットへ登録しました: {dlg.result['name']}")
        except Exception:
            tb=traceback.format_exc(); self.log(tb); messagebox.showerror('登録エラー',tb[-1600:])

    def snap_canvas_point(self, cx, cy):
        try:
            if hasattr(self,'snap_enabled_var') and not self.snap_enabled_var.get():
                return cx,cy
        except Exception:
            pass
        radius=18
        best=None; bestd=10**9
        for d in getattr(self,'image_dets',[]):
            if getattr(d,'page',self.current_page)!=self.current_page or not getattr(d,'enabled',True):
                continue
            sx=((d.x1+d.x2)/2.0)*self.preview_scale
            sy=((d.y1+d.y2)/2.0)*self.preview_scale
            dist=((sx-cx)**2+(sy-cy)**2)**0.5
            if dist<bestd and dist<=radius:
                best=(sx,sy); bestd=dist
        for it in getattr(self,'manual_symbol_items',[]):
            if int(it.get('page',0))!=self.current_page:
                continue
            sx=float(it.get('x',0))*self.preview_scale
            sy=float(it.get('y',0))*self.preview_scale
            dist=((sx-cx)**2+(sy-cy)**2)**0.5
            if dist<bestd and dist<=radius:
                best=(sx,sy); bestd=dist
        if best:
            return best
        try:
            g=float(self.snap_grid_px_var.get() or 20)
            if g>1:
                return round(cx/g)*g, round(cy/g)*g
        except Exception:
            pass
        return cx,cy

    def start_cable_draw(self, x, y):
        if not self.image_pages:
            return
        cx,cy=self.canvas_event_to_canvas_xy(type('E',(),{'x':x,'y':y})())
        cx,cy=self.snap_canvas_point(cx,cy)
        self.cable_points=[(cx,cy)]
        self.cable_temp_id=None
        self.status.config(text='ケーブル描画開始：ドラッグして終点で離してください')

    def update_cable_draw(self, x, y):
        if not hasattr(self,'cable_points') or not self.cable_points:
            return
        cx,cy=self.canvas_event_to_canvas_xy(type('E',(),{'x':x,'y':y})())
        cx,cy=self.snap_canvas_point(cx,cy)
        self.cable_points.append((cx,cy))
        try:
            if getattr(self,'cable_temp_id',None):
                self.preview_canvas.delete(self.cable_temp_id)

            # 手書きのヨレヨレ線を表示しつつ、補正結果は始点終点の直線
            flat=[]
            for px,py in self.cable_points:
                flat.extend([px,py])
            self.cable_temp_id=self.preview_canvas.create_line(*flat,fill='orange',width=3)
        except Exception:
            pass

    def finish_cable_draw(self, x, y):
        pts=getattr(self,'cable_points',[])
        if len(pts)<2:
            return
        cx,cy=self.canvas_event_to_canvas_xy(type('E',(),{'x':x,'y':y})())
        cx,cy=self.snap_canvas_point(cx,cy)
        pts.append((cx,cy))
        x0,y0=pts[0]
        x1,y1=pts[-1]
        length_px=math.hypot(x1-x0,y1-y0)
        try:
            if getattr(self,'cable_temp_id',None):
                self.preview_canvas.delete(self.cable_temp_id)
        except Exception:
            pass
        try:
            self.cable_temp_id=self.preview_canvas.create_line(x0,y0,x1,y1,fill='orange',width=4)
        except Exception:
            pass
        dlg=CableInputDialog(self.root,length_px=length_px)
        self.root.wait_window(dlg)
        self.cable_points=[]
        try:
            if getattr(self,'cable_temp_id',None):
                self.preview_canvas.delete(self.cable_temp_id)
                self.cable_temp_id=None
        except Exception:
            pass
        if not dlg.result:
            return
        ix0,iy0=self.canvas_xy_to_image_xy(x0,y0)
        ix1,iy1=self.canvas_xy_to_image_xy(x1,y1)
        cable=dlg.result['type']
        px_per_m=float(dlg.result.get('scale',100) or 100)
        length_m=float(length_px)/max(0.0001,px_per_m)
        self.db.save_manual_cable(self.source_image_file,self.current_page,cable,ix0,iy0,ix1,iy1,length_px,length_m,f"px_per_m={px_per_m}; "+dlg.result.get('memo',''))
        if not hasattr(self,'manual_cable_lines'):
            self.manual_cable_lines=[]
        self.manual_cable_lines.append((self.current_page,ix0,iy0,ix1,iy1,cable,length_m,px_per_m))
        self.rebuild_manual_cable_estimate_rows()
        self.refresh_image_view()
        self.refresh_image_tables()
        messagebox.showinfo('ケーブル登録',f'{cable}: {length_px:.1f}px / {px_per_m:.1f}px/m = {length_m:.2f}m を積算へ反映しました')


    def update_cable_draw(self, x, y):
        if not hasattr(self,'cable_points') or not self.cable_points: return
        cx=self.preview_canvas.canvasx(x); cy=self.preview_canvas.canvasy(y)
        if len(self.cable_points)==1 or abs(cx-self.cable_points[-1][0])+abs(cy-self.cable_points[-1][1])>4:
            self.cable_points.append((cx,cy))
        try:
            if getattr(self,'cable_temp_id',None): self.preview_canvas.delete(self.cable_temp_id)
            flat=[]
            for p in self.cable_points: flat.extend(p)
            self.cable_temp_id=self.preview_canvas.create_line(*flat,fill='orange',width=3)
        except Exception: pass

    def finish_cable_draw(self, x, y):
        pts=getattr(self,'cable_points',[])
        if len(pts)<2: return
        x0,y0=pts[0]; x1,y1=pts[-1]
        length_px=math.hypot(x1-x0,y1-y0)
        # 縮尺設定がある場合は px/m と長さ(m) を自動計算して渡す
        scale_init=None; auto_length_m=None
        if getattr(self,'scale_px_per_mm',None):
            image_px=length_px/max(0.0001,self.preview_scale)
            auto_length_m=image_px/(self.scale_px_per_mm*1000)
            scale_init=self.scale_px_per_mm*1000  # px/m に換算
        dlg=CableInputDialog(self.root,length_px=length_px,scale_init=scale_init,auto_length_m=auto_length_m)
        self.root.wait_window(dlg)
        try:
            if getattr(self,'cable_temp_id',None): self.preview_canvas.delete(self.cable_temp_id)
        except Exception: pass
        self.cable_points=[]
        if not dlg.result: return
        ix0,iy0=int(x0/max(0.0001,self.preview_scale)),int(y0/max(0.0001,self.preview_scale))
        ix1,iy1=int(x1/max(0.0001,self.preview_scale)),int(y1/max(0.0001,self.preview_scale))
        cable=dlg.result['type']; length_m=dlg.result['length_m']
        self.db.save_manual_cable(self.source_image_file,self.current_page,cable,ix0,iy0,ix1,iy1,length_px,length_m,dlg.result.get('memo',''))
        if not hasattr(self,'manual_image_estimate_rows'): self.manual_image_estimate_rows=[]
        cat,item,spec,unit,price=self.db.find_price(cable)
        self.manual_image_estimate_rows.append((cat,item or cable,'m',round(length_m,2),price,round(float(price)*length_m),'cable_manual'))
        if not hasattr(self,'manual_cable_lines'): self.manual_cable_lines=[]
        self.manual_cable_lines.append((self.current_page,ix0,iy0,ix1,iy1,cable,length_m))
        self.refresh_image_view(); self.refresh_image_tables()
        messagebox.showinfo('ケーブル登録',f'{cable}: {length_m:.2f}m を積算へ反映しました')

    def show_detection_context_menu(self,event):
        try:
            iid=self.img_det_tree.identify_row(event.y)
            if iid:
                self.img_det_tree.selection_set(iid)
                self.selected_det_id=int(iid)
            m=tk.Menu(self.root,tearoff=0)
            m.add_command(label='編集',command=self.edit_selected_detection)
            m.add_command(label='削除',command=self.delete_selected_detection)
            m.add_command(label='無効候補を一括削除',command=self.delete_disabled_detections)
            m.tk_popup(event.x_root,event.y_root)
        finally:
            try: m.grab_release()
            except Exception: pass

    def show_sum_context_menu(self,event):
        try:
            iid=self.img_sum_tree.identify_row(event.y)
            if iid:
                self.img_sum_tree.selection_set(iid)
            m=tk.Menu(self.root,tearoff=0)
            m.add_command(label='積算行を編集',command=self.edit_selected_image_sum_row)
            m.add_command(label='積算行を追加',command=self.add_image_sum_row)
            m.add_command(label='積算行を削除',command=self.delete_selected_image_sum_row)
            m.tk_popup(event.x_root,event.y_root)
        finally:
            try: m.grab_release()
            except Exception: pass

    def add_symbol_at_canvas(self, event):
        """記号貼付ON時：クリック位置に簡易シンボルを貼る。枠だけにしない。"""
        if not self.image_pages:
            return
        cx,cy=self.canvas_event_to_canvas_xy(event)
        ix,iy=self.canvas_xy_to_image_xy(cx,cy)
        name=self.symbol_paste_var.get().strip() if hasattr(self,'symbol_paste_var') else 'LEDダウンライト'
        img=self.image_pages[self.current_page-1]
        size=40  # 選択・削除しやすいよう検出範囲を拡大
        det=ImageDetection(len(self.image_dets)+1,self.current_page,max(0,ix-size),max(0,iy-size),min(img.width-1,ix+size),min(img.height-1,iy+size),'symbol',name,1.0,0,True,'symbol_paste','manual symbol paste')
        self.image_dets.append(det)
        if not hasattr(self,'manual_symbol_items'):
            self.manual_symbol_items=[]
        self.manual_symbol_items.append({'page':self.current_page,'x':ix,'y':iy,'name':name,'det_id':det.det_id})
        self.select_image_det(det.det_id)
        self.refresh_image_view()
        self.refresh_image_tables()
        self.status.config(text=f'図面記号を貼り付け: {name}')


    def edit_selected_cable(self):
        if not hasattr(self,'manual_cable_lines') or not self.manual_cable_lines:
            messagebox.showwarning('未選択','編集するケーブルがありません')
            return
        idx=None
        if hasattr(self,'cable_tree'):
            sel=self.cable_tree.selection()
            if sel:
                try: idx=int(self.cable_tree.item(sel[0],'values')[0])-1
                except Exception: idx=None
        if idx is None:
            idx=len(self.manual_cable_lines)-1
        if idx<0 or idx>=len(self.manual_cable_lines):
            return
        row=self.manual_cable_lines[idx]
        if len(row)>=8:
            pg,x1,y1,x2,y2,cable,length_m,px_per_m=row[:8]
        else:
            pg,x1,y1,x2,y2,cable,length_m=row[:7]
            px_per_m=100.0
        dlg=ManualCableEditDialog(self.root,{'type':cable,'length_m':length_m,'x1':x1,'y1':y1,'x2':x2,'y2':y2})
        self.root.wait_window(dlg)
        if not dlg.result:
            return
        r=dlg.result
        px_len=math.hypot(r['x2']-r['x1'],r['y2']-r['y1'])
        new_len=max(0.0,float(r['length_m']))
        new_px_per_m=px_len/max(0.0001,new_len) if new_len>0 else px_per_m
        self.manual_cable_lines[idx]=(pg,r['x1'],r['y1'],r['x2'],r['y2'],r['type'],new_len,new_px_per_m)
        self.rebuild_manual_cable_estimate_rows()
        self.refresh_image_view()
        self.refresh_image_tables()
        self.status.config(text=f'ケーブル編集: {r["type"]} {new_len:.2f}m')


    def delete_selected_cable(self):
        if not hasattr(self,'manual_cable_lines') or not self.manual_cable_lines:
            messagebox.showwarning('未選択','削除するケーブルがありません')
            return
        idx=None
        if hasattr(self,'cable_tree'):
            sel=self.cable_tree.selection()
            if sel:
                try: idx=int(self.cable_tree.item(sel[0],'values')[0])-1
                except Exception: idx=None
        if idx is None:
            idx=len(self.manual_cable_lines)-1
        if idx<0 or idx>=len(self.manual_cable_lines):
            return
        if not messagebox.askyesno('確認','選択した手書きケーブルを削除しますか？'):
            return
        del self.manual_cable_lines[idx]
        self.rebuild_manual_cable_estimate_rows()
        self.refresh_image_view()
        self.refresh_image_tables()
        self.status.config(text='手書きケーブルを削除しました')

    def rebuild_manual_cable_estimate_rows(self):
        preserved=[]
        for r in getattr(self,'manual_image_estimate_rows',[]):
            if not (len(r)>=7 and r[6]=='cable_manual'):
                preserved.append(r)
        self.manual_image_estimate_rows=preserved
        for row in getattr(self,'manual_cable_lines',[]):
            if len(row)>=8:
                pg,x1,y1,x2,y2,cable,length_m,px_per_m=row[:8]
            else:
                pg,x1,y1,x2,y2,cable,length_m=row[:7]
            cat,item,spec,unit,price=self.db.find_price(cable)
            if not item or item=='未分類設備':
                item=cable
            self.manual_image_estimate_rows.append((cat,item,'m',round(float(length_m),2),price,round(float(price)*float(length_m)),'cable_manual'))


    def apply_lighting_rail_overlap_fix(self, show_message=False):
        """
        ライティングレールとダウンライトが重なっている図面向けの補正。
        細長い候補をライティングレールとして別カウントし、
        その近傍/重なりにある丸形候補はダウンライトとして残す。
        """
        try:
            if not getattr(self,'image_dets',None):
                if show_message:
                    messagebox.showinfo('ライティングレール補正','検出候補がありません')
                return 0

            rails=[]
            lights=[]
            changed=0

            for d in self.image_dets:
                if not getattr(d,'enabled',True):
                    continue
                if is_lighting_rail_like_detection(d):
                    rails.append(d)
                elif is_round_light_like_detection(d):
                    lights.append(d)

            for r in rails:
                old=r.equipment
                r.equipment='ライティングレール'
                r.color_name='rail'
                r.source='rail_shape'
                r.memo=(getattr(r,'memo','') + ' / lighting rail shape detected').strip()
                r.enabled=True
                if old!='ライティングレール':
                    changed+=1

            # レール付近の丸形記号はダウンライトとして明示
            for l in lights:
                near_any=False
                for r in rails:
                    if detection_intersects_or_near(l,r,margin=18):
                        near_any=True
                        break
                if near_any:
                    old=l.equipment
                    if old in ('未分類設備','ライティングレール','LEDダウンライト',''):
                        l.equipment='LEDダウンライト'
                        l.source='rail_mounted_light'
                        l.memo=(getattr(l,'memo','') + ' / mounted on lighting rail').strip()
                        l.enabled=True
                        if old!='LEDダウンライト':
                            changed+=1

            # もしレールが検出候補化されておらず、横長線だけが大きすぎて除外された場合に備え、
            # 既存のダウンライト群が水平に並んでいる場合は仮レールを追加する。
            if not rails and len(lights) >= 2:
                sorted_l=sorted(lights,key=lambda d:d.x1)
                for i in range(len(sorted_l)-1):
                    a,b=sorted_l[i],sorted_l[i+1]
                    ay=(a.y1+a.y2)/2; by=(b.y1+b.y2)/2
                    if abs(ay-by) < 20 and abs(b.x1-a.x2) < 220:
                        x1=min(a.x1,b.x1)-20; y1=int(min(ay,by))-6
                        x2=max(a.x2,b.x2)+20; y2=int(max(ay,by))+6
                        det=ImageDetection(len(self.image_dets)+1,a.page,x1,y1,x2,y2,'rail','ライティングレール',0.75,0,True,'rail_inferred','inferred from aligned downlights')
                        self.image_dets.append(det)
                        changed+=1
                        break

            if changed:
                for i,d in enumerate(self.image_dets,1):
                    d.det_id=i
                self.refresh_image_view()
                self.refresh_image_tables()

            if show_message:
                messagebox.showinfo('ライティングレール補正',f'補正件数: {changed}\nレール候補: {len(rails)}\n丸形候補: {len(lights)}')
            return changed
        except Exception:
            tb=traceback.format_exc()
            self.log(tb)
            if show_message:
                messagebox.showerror('ライティングレール補正エラー',tb[-1600:])
            return 0

    def image_analysis_worker(self,p):
        try:
            self.log(f'画像/DXF解析開始: {p}'); s=self.db.get_settings(); all_d=[]
            try:
                if self.db.get_setting('dqn_enabled','1')=='1':
                    dqn_state={'image_area_norm':1.0,'aspect':1.0,'color_rule_count':len(self.db.get_image_rules())/10.0,'feature_count_norm':min((self.db.count_annotation_features() if hasattr(self.db,'count_annotation_features') else 0)/500.0,5.0),'dataset_count_norm':min((len(self.db.get_symbol_image_dataset(limit=5000)) if hasattr(self.db,'get_symbol_image_dataset') else 0)/200.0,5.0),'annotation_count_norm':0,'last_error_rate':0,'manual_ratio':0,'huge_box_penalty':0,'unknown_ratio':0}
                    dqn_action,dqn_mode=self.dqn_agent.select_action(dqn_state,epsilon=float(self.db.get_setting('dqn_epsilon','0.15')))
                    s=self.dqn_agent.apply_action_to_settings(s,dqn_action)
                    self.last_dqn_state=dqn_state; self.last_dqn_action=dqn_action; self.last_dqn_mode=dqn_mode; self.db.set_setting('dqn_last_action',dqn_action)
                    self.log(f'DQN解析戦略: action={dqn_action} mode={dqn_mode} desc={DQN_ACTION_DESCRIPTIONS.get(dqn_action,"")}')
            except Exception:
                self.log('DQN strategy selection failed:\n'+traceback.format_exc())
            self.manual_image_estimate_rows=[]
            if Path(p).suffix.lower()=='.dxf':
                pages, all_d = dxf_to_preview_image_and_detections(p, self.db, max_size=int(s.get('preview_max_width','850')), log=self.log)
            else:
                pages=load_pdf_or_image(p,float(s.get('pdf_scale','1.25')),log=self.log); rules=self.db.get_image_rules()
                for idx,img in enumerate(pages,1): all_d += analyze_image_page_light(img,idx,s,rules,log=self.log)
            s_local=self.db.get_settings()
            exp_w=int(s_local.get('max_box_w','55')); exp_h=int(s_local.get('max_box_h','55'))
            split_rules=rules if 'rules' in dir() else []  # DXFパスでは rules が未定義
            split_dets=[]
            for d in all_d:
                try:
                    img=pages[d.page-1]
                    area_ratio=d.w*d.h/max(1,img.width*img.height)
                    if area_ratio>0.008 and d.source not in ('drag','dxf') and split_rules:
                        # 実際の記号位置でスマート分割
                        cells=split_large_detection_smart(d,img,split_rules,s_local,exp_w,exp_h)
                        split_dets.extend(cells)
                    elif area_ratio>0.008 and d.source not in ('drag','dxf'):
                        d.enabled=False; d.memo=(d.memo+' / large box disabled').strip()
                        split_dets.append(d)
                    else:
                        split_dets.append(d)
                except Exception:
                    split_dets.append(d)
            for i,d in enumerate(split_dets,1): d.det_id=i
            self.ui(self.finish_image_analysis,pages,split_dets)
        except Exception: self.log(traceback.format_exc())
    def finish_image_analysis(self,pages,dets):
        self.image_pages=pages
        self.image_dets=dets
        self.current_page=1
        self.selected_det_id=None
        learned_count = self.apply_learning_to_detections(threshold=None)
        vector_count = self.apply_vector_search_to_detections(show_message=False)
        rail_count = self.apply_lighting_rail_overlap_fix(show_message=False)
        self.refresh_image_view()
        self.refresh_image_tables()
        self.status.config(text=f'画像解析完了: {len(dets)}件 / 学習反映 {learned_count}件 / ベクトル反映 {vector_count}件 / レール補正 {rail_count}件')
        # DQN reward update after analysis
        try:
            if self.db.get_setting('dqn_enabled','1')=='1' and getattr(self,'last_dqn_action',''):
                next_state=self.dqn_agent.state_from_image(self.image_pages[0] if self.image_pages else None,self.db,self.image_dets)
                reward=self.dqn_agent.reward_from_detections(self.image_dets)
                td=self.dqn_agent.update(getattr(self,'last_dqn_state',{}),self.last_dqn_action,reward,next_state,lr=float(self.db.get_setting('dqn_learning_rate','0.08')),gamma=float(self.db.get_setting('dqn_gamma','0.90')))
                self.db.save_dqn_event(self.source_image_file,self.current_page,getattr(self,'last_dqn_state',{}),self.last_dqn_action,reward,next_state,f'auto after analysis td={td:.3f}')
                self.last_dqn_reward=reward
                if hasattr(self,'dqn_status_var'): self.dqn_status_var.set(f'DQN: {self.last_dqn_action} reward={reward:.2f}')
                self.refresh_dqn_summary()
        except Exception:
            self.log('DQN reward update failed:\n'+traceback.format_exc())
    def refresh_image_view(self):
        """draw_preview()はPIL処理が重いためバックグラウンドで実行。連続呼び出しは最新だけ処理。"""
        if not self.image_pages:
            self.preview_canvas.delete('all'); return
        self._preview_req=(self.current_page, self.selected_det_id, list(self.image_dets))
        if not getattr(self,'_preview_rendering',False):
            self._preview_rendering=True
            threading.Thread(target=self._preview_render_worker,daemon=True).start()
    def _preview_render_worker(self):
        try:
            while True:
                req=self._preview_req
                if req is None: break
                page,sel_id,dets=req; self._preview_req=None
                if page<1 or page>len(self.image_pages): break
                img=self.image_pages[page-1]
                page_d=[d for d in dets if d.page==page]
                try: max_w=int(self.db.get_setting('preview_max_width','850'))
                except Exception: max_w=850
                preview=draw_preview(img,page_d,sel_id,max_w)
                scale=preview.width/max(1,img.width); total=len(self.image_pages)
                def _up(preview=preview,scale=scale,page=page,total=total):
                    try:
                        if not self.root.winfo_exists():
                            return
                        self.preview_tk=ImageTk.PhotoImage(preview); self.preview_scale=scale
                        self.preview_canvas.delete('all')
                        self.preview_canvas.create_image(0,0,image=self.preview_tk,anchor='nw')
                        try:
                            for it in getattr(self,'manual_symbol_items',[]):
                                if int(it.get('page',0))==page:
                                    sx=float(it.get('x',0))*self.preview_scale
                                    sy=float(it.get('y',0))*self.preview_scale
                                    draw_symbol_glyph_on_canvas(self.preview_canvas,sx,sy,it.get('name',''),scale=self.preview_scale,tag='manual_symbol')
                        except Exception:
                            pass
                        try:
                            for row in getattr(self,'manual_cable_lines',[]):
                                if len(row)>=8:
                                    pg,x0,y0,x1,y1,cable,length_m,px_per_m=row[:8]
                                else:
                                    pg,x0,y0,x1,y1,cable,length_m=row[:7]
                                if pg==self.current_page:
                                    sx0,sy0=x0*self.preview_scale,y0*self.preview_scale
                                    sx1,sy1=x1*self.preview_scale,y1*self.preview_scale
                                    self.preview_canvas.create_line(sx0,sy0,sx1,sy1,fill='orange',width=3)
                                    self.preview_canvas.create_text((sx0+sx1)/2,(sy0+sy1)/2,text=f'{cable} {length_m:.1f}m',fill='orange',anchor='s')
                        except Exception:
                            pass
                        # 通り芯・CADシンボルをキャンバスにオーバーレイ描画
                        try: self.draw_torii_on_canvas()
                        except Exception: pass
                        self.preview_canvas.config(scrollregion=(0,0,preview.width,preview.height))
                        self.page_label.set(f'page {page}/{total}')
                    except Exception:
                        self.log(traceback.format_exc())
                self.ui(_up)
                if self._preview_req is None: break
        except Exception: traceback.print_exc()
        finally: self._preview_rendering=False
    def refresh_image_tables(self):
        """
        v12.12:
        画像検出候補リストと画像積算結果リストを完全同期。
        画像積算結果 = 有効な画像検出候補の equipment 集計 + 手動ケーブル等。
        """
        try:
            for i in self.img_det_tree.get_children():
                self.img_det_tree.delete(i)
            for d in self.image_dets:
                self.img_det_tree.insert('',tk.END,iid=str(d.det_id),values=(
                    '1' if d.enabled else '0',
                    d.det_id,d.page,d.equipment,d.color_name,d.x1,d.y1,d.w,d.h,f'{d.score:.2f}'
                ))

            for i in self.img_sum_tree.get_children():
                self.img_sum_tree.delete(i)

            counts=defaultdict(float)
            for d in self.image_dets:
                if not d.enabled:
                    continue
                name=(d.equipment or '').strip()
                if not name or name=='未分類設備':
                    continue
                counts[name]+=1

            rows=[]; total=0.0
            for name,qty in sorted(counts.items()):
                cat,item,spec,unit,price=self.db.find_price(name)
                amount=float(price)*float(qty)
                total+=amount
                rows.append((cat,item,unit,qty,price,amount,'from_detection'))

            for r in getattr(self,'manual_image_estimate_rows',[]):
                rows.append(r)
                try:
                    total+=float(r[5])
                except Exception:
                    pass

            for r in rows:
                try: amount=int(float(r[5]))
                except Exception: amount=r[5]
                self.img_sum_tree.insert('',tk.END,values=(r[0],r[1],r[2],r[3],r[4],amount))
            self.img_sum_tree.insert('',tk.END,values=('合計','','','','',int(total)))

            if hasattr(self,'cable_tree'):
                for i in self.cable_tree.get_children():
                    self.cable_tree.delete(i)
                for idx,row in enumerate(getattr(self,'manual_cable_lines',[]),1):
                    if len(row)>=8:
                        pg,x1,y1,x2,y2,cable,length_m,px_per_m=row[:8]
                    else:
                        pg,x1,y1,x2,y2,cable,length_m=row[:7]
                    self.cable_tree.insert('',tk.END,values=(idx,pg,cable,round(float(length_m),2),x1,y1,x2,y2))
        except Exception:
            self.log(traceback.format_exc())


    def image_estimate_rows(self):
        counts=defaultdict(int)
        for d in self.image_dets:
            name=(d.equipment or '').strip()
            if d.enabled and name and name!='未分類設備':
                counts[name]+=1
        rows=[]
        for name,qty in sorted(counts.items()):
            cat,item,spec,unit,price=self.db.find_price(name)
            rows.append((cat,item,unit,qty,int(price),int(float(price)*qty)))
        return rows
    def change_image_page(self,delta):
        if not self.image_pages: return
        self.current_page=max(1,min(len(self.image_pages),self.current_page+delta)); self.selected_det_id=None; self.refresh_image_view()
    def canvas_xy_to_img(self,x,y): return int(self.preview_canvas.canvasx(x)/max(0.0001,self.preview_scale)), int(self.preview_canvas.canvasy(y)/max(0.0001,self.preview_scale))
    def find_det_at_canvas(self,x,y):
        ix,iy=self.canvas_xy_to_img(x,y); cand=[d for d in self.image_dets if d.page==self.current_page and d.x1<=ix<=d.x2 and d.y1<=iy<=d.y2]
        return sorted(cand,key=lambda d:d.w*d.h)[0] if cand else None
    def canvas_event_to_canvas_xy(self, event):
        return self.preview_canvas.canvasx(event.x), self.preview_canvas.canvasy(event.y)

    def canvas_xy_to_image_xy(self, cx, cy):
        return int(cx/max(0.0001,self.preview_scale)), int(cy/max(0.0001,self.preview_scale))

    def on_preview_mousewheel(self,event):
        try:
            self.preview_canvas.yview_scroll(int(-1*(event.delta/120)), 'units')
        except Exception:
            pass

    def on_preview_shift_mousewheel(self,event):
        try:
            self.preview_canvas.xview_scroll(int(-1*(event.delta/120)), 'units')
        except Exception:
            pass
    def on_preview_click(self,event):
        # 縮尺測定モードが最優先
        if getattr(self,'scale_mode_var',None) and self.scale_mode_var.get():
            self.on_scale_measure_start(event); return
        if not self.image_pages:
            return

        # 通り芯描画モードが優先
        if getattr(self,'torii_mode_var',None) and self.torii_mode_var.get() in ('X','Y'):
            self.add_torii_shin_at_canvas(event); return
        # スナップ記号配置モード
        if getattr(self,'snap_place_mode_var',None) and self.snap_place_mode_var.get():
            self.place_cad_symbol_at_click(event); return
        # 記号貼付ONならクリック位置に記号候補を貼り付け
        if hasattr(self,'symbol_paste_mode_var') and self.symbol_paste_mode_var.get():
            self.add_symbol_at_canvas(event)
            return

        # ケーブル描画ONなら左ドラッグはケーブル描画専用
        if hasattr(self,'cable_draw_mode_var') and self.cable_draw_mode_var.get():
            self.start_cable_draw(event.x,event.y)
            return

        cx,cy=self.canvas_event_to_canvas_xy(event)
        self.drag_start=(cx,cy)
        self.dragging=False

        d=self.find_det_at_canvas(event.x,event.y)
        if d:
            self.select_image_det(d.det_id)

    def on_preview_drag(self,event):
        # 縮尺測定モードが最優先
        if getattr(self,'scale_mode_var',None) and self.scale_mode_var.get():
            self.on_scale_measure_drag(event); return
        if not self.image_pages:
            return
        # スナップカーソル表示（通り芯モード / スナップ配置モード時）
        if getattr(self,'snap_object_var',None) and self.snap_object_var.get():
            if getattr(self,'torii_mode_var',tk.StringVar()).get() in ('X','Y') or                getattr(self,'snap_place_mode_var',tk.BooleanVar()).get():
                self.show_snap_cursor(event.x, event.y)
        if hasattr(self,'symbol_paste_mode_var') and self.symbol_paste_mode_var.get():
            return
        if hasattr(self,'cable_draw_mode_var') and self.cable_draw_mode_var.get():
            self.update_cable_draw(event.x,event.y)
            return

        if self.drag_start is None:
            cx,cy=self.canvas_event_to_canvas_xy(event)
            self.drag_start=(cx,cy)
            return
        cx,cy=self.canvas_event_to_canvas_xy(event)
        x0,y0=self.drag_start
        if abs(cx-x0)+abs(cy-y0) < 8:
            return
        self.dragging=True
        try:
            if self.drag_rect_id:
                self.preview_canvas.delete(self.drag_rect_id)
            self.drag_rect_id=self.preview_canvas.create_rectangle(x0,y0,cx,cy,outline='red',width=2,dash=(4,2))
        except Exception:
            pass

    def on_preview_drag_release(self,event):
        # 縮尺測定モードが最優先
        if getattr(self,'scale_mode_var',None) and self.scale_mode_var.get():
            self.on_scale_measure_release(event); return
        if not self.image_pages:
            return
        if hasattr(self,'symbol_paste_mode_var') and self.symbol_paste_mode_var.get():
            return
        if hasattr(self,'cable_draw_mode_var') and self.cable_draw_mode_var.get():
            self.finish_cable_draw(event.x,event.y)
            return

        if not getattr(self,'dragging',False) or self.drag_start is None:
            self.drag_start=None
            self.dragging=False
            return
        try:
            if self.drag_rect_id:
                self.preview_canvas.delete(self.drag_rect_id)
                self.drag_rect_id=None
        except Exception:
            pass
        x0,y0=self.drag_start
        cx,cy=self.canvas_event_to_canvas_xy(event)
        self.drag_start=None
        self.dragging=False
        ix0,iy0=self.canvas_xy_to_image_xy(x0,y0)
        ix1,iy1=self.canvas_xy_to_image_xy(cx,cy)
        xa,xb=sorted([ix0,ix1])
        ya,yb=sorted([iy0,iy1])
        if xb-xa < 8 or yb-ya < 8:
            return
        img=self.image_pages[self.current_page-1]
        det=ImageDetection(len(self.image_dets)+1,self.current_page,max(0,xa),max(0,ya),min(img.width-1,xb),min(img.height-1,yb),'manual','未分類設備',0.5,0,True,'drag','manual drag annotation range')
        self.image_dets.append(det)
        guess=self.local_guess_from_detection(det)
        det.equipment=guess
        det.memo=f'drag local/learned guess: {guess}'
        self.select_image_det(det.det_id)
        self.llm_answer_var.set(guess)
        self.correct_answer_var.set(guess)
        if hasattr(self,'img_ai_status_var'):
            self.img_ai_status_var.set(f'AI状態: ドラッグ範囲を推定 → {guess}')
        if hasattr(self,'img_ai_usage_var'):
            self.img_ai_usage_var.set('AI使用量: ローカル/学習推定のため0トークン/0円')
        self.refresh_image_view()
        self.refresh_image_tables()
        self.status.config(text=f'ドラッグ範囲をアノテーション: {guess}')

    def on_preview_double(self,event):
        d=self.find_det_at_canvas(event.x,event.y)
        if not d and self.image_pages:
            ix,iy=self.canvas_xy_to_img(event.x,event.y); img=self.image_pages[self.current_page-1]; box=30; d=ImageDetection(len(self.image_dets)+1,self.current_page,max(0,ix-box),max(0,iy-box),min(img.width-1,ix+box),min(img.height-1,iy+box),'manual','未分類設備',0.5,0,True,'manual','manual'); self.image_dets.append(d); self.refresh_image_tables()
        if d: self.select_image_det(d.det_id); self.annotate_selected_popup_async()
    def on_preview_right_click(self,event):
        """
        右クリック:
        - 既存検出枠上: コンテキストメニュー（削除/編集/AI判定）
        - 貼付記号上: 削除/設備名変更/AI判定
        - 空白: 手動枠を作成してAI判定
        """
        d=self.find_det_at_canvas(event.x,event.y)
        if d:
            self.select_image_det(d.det_id)
            src=getattr(d,'source','')
            m=tk.Menu(self.root,tearoff=0)
            lbl=f'ID={d.det_id}: {d.equipment}'
            if src=='symbol_paste': lbl='[貼付] '+lbl
            elif src in ('smart_split','auto_split','learned','dataset'): lbl='[自動] '+lbl
            m.add_command(label=lbl,state='disabled')
            m.add_separator()
            m.add_command(label='削除',command=self.delete_selected_detection)
            m.add_command(label='編集（設備名・有効/無効）',command=self.edit_selected_detection)
            m.add_separator()
            m.add_command(label='AI判定実行',command=lambda:self.annotate_selected_popup_async(event.x_root,event.y_root))
            m.add_command(label='設備名だけ変更',command=self.apply_correct_to_selected)
            m.add_command(label='無効化（積算から除外）',command=self.disable_image_selected)
            try: m.tk_popup(event.x_root,event.y_root)
            finally:
                try: m.grab_release()
                except Exception: pass
        elif self.image_pages:
            # 空白クリック: 小さい手動枠を作成してAI判定
            ix,iy=self.canvas_xy_to_img(event.x,event.y)
            img=self.image_pages[self.current_page-1]; box=30
            d=ImageDetection(len(self.image_dets)+1,self.current_page,
                max(0,ix-box),max(0,iy-box),min(img.width-1,ix+box),min(img.height-1,iy+box),
                'manual','未分類設備',0.5,0,True,'manual','right click manual')
            self.image_dets.append(d); self.refresh_image_tables()
            self.select_image_det(d.det_id)
            self.annotate_selected_popup_async(event.x_root,event.y_root)

    def annotate_selected_popup_async(self,screen_x=None,screen_y=None):
        d=self.get_image_det(self.selected_det_id)
        if not d:
            messagebox.showwarning('未選択','図記号を選択してください')
            return

        # 学習/ドラッグ/手動貼付の判定結果が既にある場合は、LLMに投げ直してLED固定化させない。
        preferred = ''
        try:
            preferred = (self.correct_answer_var.get() or self.llm_answer_var.get() or '').strip()
        except Exception:
            preferred = ''
        learned_guess = self.local_guess_from_detection(d)
        if preferred and preferred not in ('AI判定中...','未分類設備'):
            d.equipment = preferred
            self.show_annotation_popup(d.det_id, preferred, preferred + '\n既存の学習/訂正候補を優先しました。', getattr(self,'pending_crop_path',''), screen_x, screen_y)
            return
        if learned_guess and learned_guess not in ('未分類設備','LEDダウンライト') and getattr(d,'source','') in ('learned','dataset','drag','symbol_paste','manual_edit'):
            d.equipment = learned_guess
            self.llm_answer_var.set(learned_guess)
            self.correct_answer_var.set(learned_guess)
            self.show_annotation_popup(d.det_id, learned_guess, learned_guess + '\n学習DB/記号DB/手動範囲の推定を優先しました。', getattr(self,'pending_crop_path',''), screen_x, screen_y)
            return

        if getattr(self,'annotation_busy',False):
            messagebox.showinfo('処理中','現在AI判定中です。12秒以内に自動復帰します。')
            return

        self.annotation_busy=True
        self.annotation_token=datetime.now().strftime('%Y%m%d%H%M%S%f')
        token=self.annotation_token

        self.llm_answer_var.set('AI判定中...')
        self.status.config(text='AI判定中...')
        wait_win=self.show_annotation_wait_popup(d,screen_x,screen_y)

        provider='ollama'
        try:
            provider=self.ai_provider_var.get().strip()
        except Exception:
            pass

        # Ollama未起動なら問い合わせせず即ローカル推定
        if provider=='ollama':
            try:
                url=self.ollama_url_var.get().strip()
            except Exception:
                url=DEFAULT_OLLAMA_URL
            if not quick_ollama_ping(url,timeout=2):
                fallback=self.local_guess_from_detection(d)
                def offline():
                    self.annotation_busy=False
                    self.annotation_token=None
                    try:
                        if wait_win and wait_win.winfo_exists():
                            wait_win.destroy()
                    except Exception:
                        pass
                    self.llm_answer_var.set(fallback)
                    self.correct_answer_var.set(fallback)
                    d.equipment=fallback
                    d.memo='ollama offline local fallback'
                    self.refresh_image_view()
                    self.refresh_image_tables()
                    self.show_annotation_popup(
                        d.det_id,
                        fallback,
                        fallback+'\nOllama未接続のためローカル推定です。\nOK/NGで学習できます。',
                        getattr(self,'pending_crop_path',''),
                        None,
                        None
                    )
                    self.status.config(text=f'Ollama未接続。ローカル推定: {fallback}')
                self.root.after(0,offline)
                return

        ai_snap=self.snapshot_ai_settings()
        timeout_sec=90 if ai_snap.get('provider','ollama') in ('anthropic','openai','custom_openai') else 13
        threading.Thread(
            target=self.annotation_popup_worker_safe,
            args=(d.det_id,wait_win,token,ai_snap),
            daemon=True
        ).start()

        watchdog_ms=max(15000,int(timeout_sec*1000+3000))
        def ai_watchdog(det_id=d.det_id,win=wait_win,tk_token=token):
            if getattr(self,'annotation_busy',False) and getattr(self,'annotation_token',None)==tk_token:
                dd=self.get_image_det(det_id)
                fallback=self.local_guess_from_detection(dd)
                self.annotation_busy=False
                self.annotation_token=None
                try:
                    if win and win.winfo_exists():
                        win.destroy()
                except Exception:
                    pass
                self.llm_answer_var.set(fallback)
                self.correct_answer_var.set(fallback)
                if dd:
                    dd.equipment=fallback
                    dd.memo='AI watchdog 12s local fallback'
                self.refresh_image_view()
                self.refresh_image_tables()
                self.show_annotation_popup(
                    det_id,
                    fallback,
                    fallback+'\nAI応答が12秒以内に返らなかったためローカル推定に切替しました。\nOK/NGで学習できます。',
                    getattr(self,'pending_crop_path',''),
                    None,
                    None
                )
                self.status.config(text=f'AI応答なし。ローカル推定へ切替: {fallback}')
        self.root.after(watchdog_ms if ai_snap.get('provider','ollama') in ('anthropic','openai','custom_openai') else 12000,ai_watchdog)

    def show_annotation_wait_popup(self,d,screen_x=None,screen_y=None):
        win=tk.Toplevel(self.root)
        win.title('AI図記号アノテーション')
        win.geometry('500x260')
        if screen_x is not None and screen_y is not None:
            try:
                win.geometry(f'+{int(screen_x)+20}+{int(screen_y)+20}')
            except Exception:
                pass
        win.transient(self.root)

        frame=ttk.Frame(win,padding=14)
        frame.pack(fill=tk.BOTH,expand=True)

        ttk.Label(frame,text='AIが図記号を判定中です...',font=('Arial',12,'bold')).pack(anchor='w')
        ttk.Label(frame,text=f'ID={d.det_id} / page={d.page} / 色={d.color_name} / 自動推定={d.equipment}',foreground='gray').pack(anchor='w',pady=6)
        ttk.Label(frame,text='12秒以内に応答が無い場合はローカル推定へ自動切替します。',wraplength=450).pack(anchor='w',pady=6)

        pb=ttk.Progressbar(frame,mode='indeterminate')
        pb.pack(fill=tk.X,pady=10)
        pb.start(10)

        status=tk.StringVar(value='Ollamaへ問い合わせ中...')
        ttk.Label(frame,textvariable=status).pack(anchor='w')

        def close_only():
            try:
                win.destroy()
            except Exception:
                pass

        ttk.Button(frame,text='閉じる（処理は継続）',command=close_only).pack(anchor='e',pady=10)
        win._status_var=status
        return win

    def annotation_popup_worker_safe(self,det_id,wait_win,token=None,ai_snap=None):
        if ai_snap is None: ai_snap=self.snapshot_ai_settings()
        try:
            self.annotation_popup_worker_core(det_id,wait_win,ai_snap,token)
        except Exception as e:
            d=self.get_image_det(det_id)
            fallback=self.local_guess_from_detection(d)
            crop_path=getattr(self,'pending_crop_path','')
            def err_done():
                if token is not None and getattr(self,'annotation_token',None)!=token:
                    return
                self.annotation_busy=False
                self.annotation_token=None
                try:
                    if wait_win and wait_win.winfo_exists():
                        wait_win.destroy()
                except Exception:
                    pass
                self.llm_answer_var.set(fallback)
                self.correct_answer_var.set(fallback)
                if d:
                    d.equipment=fallback
                    d.memo='AI error local fallback'
                self.refresh_image_view()
                self.refresh_image_tables()
                self.status.config(text=f'AIエラー：ローカル推定 {fallback}')
                self.log('AIエラーのためローカル推定へ切替: '+str(e))
                self.show_annotation_popup(
                    det_id,
                    fallback,
                    fallback+'\nAIエラーのためローカル推定です。\nOK/NGで学習できます。',
                    crop_path,
                    None,
                    None
                )
            self.ui(err_done)


    def annotation_popup_worker_core(self,det_id,wait_win,ai_snap,token=None):
        d=self.get_image_det(det_id)
        if not d: self.ui(lambda:setattr(self,'annotation_busy',False)); return
        img=self.image_pages[d.page-1]; crop=crop_symbol(img,d)
        crop_path=str(CROP_DIR/f"rightclick_crop_{datetime.now().strftime('%Y%m%d_%H%M%S')}_id{d.det_id}.png")
        crop.save(crop_path)
        mem_ctx=''
        try:
            mem_ctx=LongTermMemoryManager(self.db).context(
                f'図記号 {d.equipment} {d.color_name} bbox {d.w}x{d.h}',
                namespace='symbols',
                top_k=int(ai_snap.get('memory_top_k','6')) if ai_snap.get('use_langmem','1')=='1' else 0
            )
        except Exception:
            mem_ctx=''
        prompt=(
            f"あなたは日本の電気工事図面の図記号判定補助AIです。\n"
            f"学習/ローカル推定: {self.local_guess_from_detection(d)}\n候補: {self.local_guess_from_detection(d)}, LEDダウンライト, LEDベースライト, コンセント, 片切スイッチ, 分電盤, 換気扇, 非常灯, 誘導灯, 未分類設備\n"
            f"特徴: 色={d.color_name}, 自動推定={d.equipment}, page={d.page}, "
            f"bbox=({d.x1},{d.y1})-({d.x2},{d.y2}), 幅={d.w}, 高さ={d.h}\n"
            f"{mem_ctx}\n"
            f"1行目は設備名だけ。2行目以降は理由20文字以内。不明なら「未分類設備」。"
        )
        provider=ai_snap.get('provider','ollama')
        timeout_sec=90 if provider in ('anthropic','openai','custom_openai') else 13
        def _upd():
            try:
                if wait_win and wait_win.winfo_exists(): wait_win._status_var.set(f'AI({provider})へ問い合わせ中...')
            except Exception: pass
        self.ui(_upd)
        answer=UnifiedAIClient(ai_snap).generate(prompt,timeout=timeout_sec)
        short=str(answer).strip().splitlines()[0].strip() if str(answer).strip() else '未分類設備'
        usage=answer.usage_summary() if hasattr(answer,'usage_summary') else ''
        def done():
            if token is not None and getattr(self,'annotation_token',None)!=token: return
            self.annotation_busy=False; self.annotation_token=None
            try:
                if wait_win and wait_win.winfo_exists(): wait_win.destroy()
            except Exception: pass
            try:
                if hasattr(self,'img_ai_usage_var') and usage: self.img_ai_usage_var.set('AI使用量: '+usage)
            except Exception: pass
            self.show_annotation_popup(det_id,short,str(answer),crop_path,None,None)
        self.ui(done)
    def show_annotation_popup(self,det_id,short_answer,full_answer,crop_path,screen_x=None,screen_y=None):
        d=self.get_image_det(det_id)
        if not d:
            self.annotation_busy=False
            return
        self.pending_crop_path=crop_path
        self.pending_llm_answer=short_answer
        self.llm_answer_var.set(short_answer)
        self.correct_answer_var.set(short_answer)
        self.ai_result.insert(tk.END,'\n--- 右クリック図記号AI判定 ---\n'+full_answer+f'\n切り出し: {crop_path}\n')
        self.ai_result.see(tk.END)
        self.status.config(text=f'AI判定完了: {short_answer}')
        try:
            if hasattr(self,'img_ai_status_var'):
                self.img_ai_status_var.set(f'AI状態: 判定完了 → {short_answer}')
        except Exception:
            pass

        win=tk.Toplevel(self.root)
        win.title('AI図記号アノテーション')
        win.geometry('540x380')
        if screen_x is not None and screen_y is not None:
            try:
                win.geometry(f'+{int(screen_x)+20}+{int(screen_y)+20}')
            except Exception:
                pass
        win.transient(self.root)

        outer=ttk.Frame(win,padding=12)
        outer.pack(fill=tk.BOTH,expand=True)

        ttk.Label(outer,text=f'選択図記号 ID={d.det_id} / page={d.page}',font=('Arial',10,'bold')).pack(anchor='w')
        ttk.Label(outer,text=f'bbox=({d.x1},{d.y1})-({d.x2},{d.y2}) / 色={d.color_name}',foreground='gray').pack(anchor='w',pady=(0,8))

        ans_frame=ttk.LabelFrame(outer,text='AI回答',padding=8)
        ans_frame.pack(fill=tk.BOTH,expand=True)

        ttk.Label(ans_frame,text=short_answer,font=('Arial',14,'bold'),foreground='blue').pack(anchor='w')

        txt=scrolledtext.ScrolledText(ans_frame,height=7,wrap='word')
        txt.pack(fill=tk.BOTH,expand=True,pady=6)
        txt.insert('1.0',full_answer)
        txt.configure(state='disabled')

        btns=ttk.Frame(outer)
        btns.pack(fill=tk.X,pady=8)

        ttk.Button(btns,text='OK：正しいので反映・学習',command=lambda:self.popup_ok_register(win,det_id,short_answer,crop_path)).pack(side=tk.LEFT,padx=4)
        ttk.Button(btns,text='NG：訂正して学習',command=lambda:self.popup_ng_input(win,det_id,short_answer,crop_path)).pack(side=tk.LEFT,padx=4)
        ttk.Button(btns,text='閉じる',command=win.destroy).pack(side=tk.RIGHT,padx=4)

    def popup_ok_register(self,win,det_id,answer,crop_path):
        # OK登録ではAIへ再問い合わせしない。人間のOKを正解として即保存する。
        ok = self.safe_save_annotation_only(
            det_id=det_id,
            llm_answer=answer,
            final_answer=answer,
            correct=True,
            crop_path=crop_path,
            memo='popup OK no requery'
        )
        try:
            win.destroy()
        except Exception:
            pass
        if ok:
            messagebox.showinfo('登録完了',f'OK登録しました: {answer}')

    def popup_ng_input(self,win,det_id,llm_answer,crop_path):
        try:
            win.destroy()
        except Exception:
            pass
        d=self.get_image_det(det_id)
        if not d:
            return

        ng=tk.Toplevel(self.root)
        ng.title('正しい図記号名を入力')
        ng.geometry('460x220')
        ng.transient(self.root)

        frame=ttk.Frame(ng,padding=14)
        frame.pack(fill=tk.BOTH,expand=True)

        ttk.Label(frame,text='AI回答が間違っている場合、正しい設備名を入力してください。').pack(anchor='w')
        ttk.Label(frame,text=f'AI回答: {llm_answer}',foreground='gray').pack(anchor='w',pady=(4,8))

        final_var=tk.StringVar(value=d.equipment if d.equipment!='未分類設備' else '')
        combo=ttk.Combobox(frame,textvariable=final_var,values=['LEDダウンライト','LEDベースライト','コンセント','片切スイッチ','分電盤','換気扇','未分類設備'],width=34)
        combo.pack(anchor='w',pady=4)
        combo.focus_set()

        def save():
            final=final_var.get().strip()
            if not final:
                messagebox.showwarning('未入力','正しい設備名を入力してください',parent=ng)
                return
            d.equipment=final
            d.enabled=True
            d.memo='right click popup NG correction'
            self.db.save_annotation(self.source_image_file,d,llm_answer,final,False,crop_path,d.memo)
            self.correct_answer_var.set(final)
            self.llm_answer_var.set(llm_answer)
            self.refresh_image_view()
            self.refresh_image_tables()
            self.refresh_learning_summary()
            ng.destroy()
            messagebox.showinfo('登録完了',f'訂正登録しました: {final}')

        btn=ttk.Frame(frame)
        btn.pack(fill=tk.X,pady=12)
        ttk.Button(btn,text='訂正登録・学習',command=save).pack(side=tk.LEFT,padx=4)
        ttk.Button(btn,text='キャンセル',command=ng.destroy).pack(side=tk.RIGHT,padx=4)

    def select_image_det(self,det_id, update_tree=True, redraw=True):
        # 左クリック/Treeview選択の相互再帰を防ぐ安全版。
        # 旧版は canvas -> selection_set -> <<TreeviewSelect>> -> select_image_det が再帰し、
        # プレビュー再描画スレッドが大量発生してフリーズ/クラッシュすることがあった。
        try:
            det_id=int(det_id)
        except Exception:
            return
        d=self.get_image_det(det_id)
        if not d:
            return

        same = (self.selected_det_id == det_id)
        self.selected_det_id=det_id
        self.current_page=d.page
        self.selected_info_var.set(
            f'選択 ID={d.det_id} page={d.page} {d.equipment} {d.color_name} '
            f'bbox=({d.x1},{d.y1})-({d.x2},{d.y2})'
        )
        self.correct_answer_var.set(d.equipment)

        if update_tree and hasattr(self, 'img_det_tree'):
            try:
                self._suppress_tree_select=True
                if str(det_id) in self.img_det_tree.get_children(''):
                    self.img_det_tree.selection_set(str(det_id))
                    self.img_det_tree.see(str(det_id))
            except Exception:
                pass
            finally:
                self._suppress_tree_select=False

        # 同じ図記号を連続クリックした場合は再描画しない。
        if redraw and not same:
            self.refresh_image_view()

    def get_image_det(self,det_id):
        try:
            det_id=int(det_id)
        except Exception:
            return None
        for d in self.image_dets:
            if d.det_id==det_id: return d
        return None

    def on_img_tree_select(self,event=None):
        if getattr(self,'_suppress_tree_select',False):
            return 'break'
        try:
            sel=self.img_det_tree.selection()
            if sel:
                self.select_image_det(int(sel[0]), update_tree=False, redraw=True)
        except Exception:
            self.log(traceback.format_exc())
        return 'break'
    def disable_image_selected(self):
        d=self.get_image_det(self.selected_det_id)
        if d: d.enabled=False; d.memo='manual disabled'; self.refresh_image_view(); self.refresh_image_tables()
    def apply_correct_to_selected(self):
        d=self.get_image_det(self.selected_det_id)
        if d: d.equipment=self.correct_answer_var.get().strip() or d.equipment; d.enabled=True; d.memo='manual equipment changed'; self.refresh_image_view(); self.refresh_image_tables()
    def annotate_selected_async(self):
        d=self.get_image_det(self.selected_det_id)
        if not d: messagebox.showwarning('未選択','図記号を選択してください'); return
        ai_snap=self.snapshot_ai_settings()
        self.llm_answer_var.set('AI問い合わせ中...')
        threading.Thread(target=self.annotation_worker,args=(d.det_id,ai_snap),daemon=True).start()
    def annotation_worker(self,det_id,ai_snap):
        try:
            d=self.get_image_det(det_id); img=self.image_pages[d.page-1]; crop=crop_symbol(img,d)
            crop_path=str(CROP_DIR/f"crop_{datetime.now().strftime('%Y%m%d_%H%M%S')}_id{d.det_id}.png"); crop.save(crop_path)
            prompt=f"""あなたは日本の電気工事図面の図記号判定補助AIです。
この切り出し画像または特徴情報から設備名を1行で答えてください。
候補: LEDダウンライト, LEDベースライト, コンセント, 片切スイッチ, 分電盤, 換気扇, 未分類設備
特徴: 色={d.color_name}, 自動推定={d.equipment}, bbox=({d.x1},{d.y1})-({d.x2},{d.y2}), 幅={d.w}, 高さ={d.h}
第一行は設備名だけにしてください。"""
            ans=UnifiedAIClient(ai_snap).generate(prompt,timeout=60)
            short=ans.strip().splitlines()[0].strip() if ans.strip() else '未分類設備'
            self.ui(self.finish_annotation,short,ans,crop_path)
        except Exception:
            tb=traceback.format_exc()
            self.ui(lambda tb=tb:self.ai_result.insert(tk.END,tb+'\n'))
            self.ui(lambda:self.llm_answer_var.set('推定失敗'))
    def finish_annotation(self,short,ans,crop_path):
        self.pending_crop_path=crop_path; self.pending_llm_answer=short; self.llm_answer_var.set(short); self.correct_answer_var.set(short); self.ai_result.insert(tk.END,'\n--- 図記号推定 ---\n'+ans+f'\n切り出し: {crop_path}\n'); self.ai_result.see(tk.END); self.status.config(text=f'LLM推定完了: {short}')
    def register_llm_ok(self):
        d = self.get_image_det(self.selected_det_id)
        ans = self.llm_answer_var.get().strip() if hasattr(self,'llm_answer_var') else ''
        if not d:
            messagebox.showwarning('未登録','図記号が選択されていません')
            return
        if not ans or ans in ('LLM問い合わせ中...','AI判定中...','推定失敗'):
            ans = self.local_guess_from_detection(d)
        ok = self.safe_save_annotation_only(
            det_id=d.det_id,
            llm_answer=ans,
            final_answer=ans,
            correct=True,
            crop_path=getattr(self,'pending_crop_path',''),
            memo='OK button no requery'
        )
        if ok:
            messagebox.showinfo('登録完了',f'OK登録しました: {ans}')

    def register_manual_correction(self):
        d = self.get_image_det(self.selected_det_id)
        final = self.correct_answer_var.get().strip() if hasattr(self,'correct_answer_var') else ''
        llm_ans = self.llm_answer_var.get().strip() if hasattr(self,'llm_answer_var') else ''
        if not d:
            messagebox.showwarning('未登録','図記号が選択されていません')
            return
        if not final:
            messagebox.showwarning('未入力','正しい設備名を入力してください')
            return
        ok = self.safe_save_annotation_only(
            det_id=d.det_id,
            llm_answer=llm_ans or self.local_guess_from_detection(d),
            final_answer=final,
            correct=False,
            crop_path=getattr(self,'pending_crop_path',''),
            memo='manual correction no requery'
        )
        if ok:
            messagebox.showinfo('登録完了',f'訂正登録しました: {final}')


    def export_image_csv(self):
        rows=self.image_estimate_rows()
        if not rows: messagebox.showwarning('なし','画像積算結果がありません'); return
        p=filedialog.asksaveasfilename(defaultextension='.csv',filetypes=[('CSV','*.csv')])
        if not p: return
        with open(p,'w',encoding='utf-8-sig',newline='') as f:
            w=csv.writer(f); w.writerow(['カテゴリ','品名','単位','数量','単価','金額']); w.writerows(rows); w.writerow([]); w.writerow(['検出明細']); w.writerow(['有効','ID','頁','設備名','色','x1','y1','x2','y2','score','memo'])
            for d in self.image_dets: w.writerow([d.enabled,d.det_id,d.page,d.equipment,d.color_name,d.x1,d.y1,d.x2,d.y2,d.score,d.memo])
        messagebox.showinfo('保存',p)

    def save_settings(self):
        for k,v in self.setting_vars.items(): self.db.set_setting(k,v.get())
        messagebox.showinfo('保存','設定を保存しました')
    def save_settings_and_reanalyze(self): self.save_settings(); self.start_image_analysis() if self.image_file_var.get().strip() else None
    def refresh_color_rules(self):
        for i in self.color_tree.get_children(): self.color_tree.delete(i)
        for r in self.db.get_image_rules(): self.color_tree.insert('',tk.END,iid=r['color_name'],values=[r.get(c,'') for c in self.color_tree['columns']])
    def on_color_select(self,event=None):
        sel=self.color_tree.selection()
        if not sel: return
        vals=self.color_tree.item(sel[0],'values')
        for c,v in zip(self.color_tree['columns'],vals): self.color_vars[c].set(str(v))
    def save_color_rule(self):
        try:
            r={k:v.get() for k,v in self.color_vars.items()}
            if not r.get('color_name'): return
            self.db.save_image_rule(r); self.refresh_color_rules(); messagebox.showinfo('保存','色ルールを保存しました')
        except Exception as e: messagebox.showerror('保存エラー',str(e))
    def _paste_to_var(self, var, entry):
        """クリップボードからStringVarへ貼り付け（Android対応）"""
        try:
            text = self.root.clipboard_get()
            if text:
                var.set(text.strip())
                # 入力確認のため先頭8文字+****を一時表示
                preview = text.strip()[:8] + ('****' if len(text.strip())>8 else '')
                self.status.config(text=f'貼付完了: {preview}... ({len(text.strip())}文字)')
        except Exception as e:
            messagebox.showwarning('貼付エラー', f'クリップボードから取得できません。直接入力してください。\n{e}')

    def _toggle_show(self, entry):
        """APIキーフィールドの表示/非表示切替"""
        try:
            if entry.cget('show') == '*':
                entry.config(show='')
                self.root.after(3000, lambda: self._hide_entry(entry))  # 3秒後に自動非表示
                self.status.config(text='APIキーを表示中（3秒後に自動非表示）')
            else:
                entry.config(show='*')
                self.status.config(text='APIキーを非表示にしました')
        except Exception: pass

    def _hide_entry(self, entry):
        try: entry.config(show='*')
        except Exception: pass

    def save_ollama_settings_silent(self):
        """メッセージダイアログなしで設定保存（テスト前自動保存用）"""
        try:
            self.db.set_setting('ai_provider',self.ai_provider_var.get().strip())
            self.db.set_setting('ollama_url',self.ollama_url_var.get().strip())
            self.db.set_setting('ollama_model',self.ollama_model_var.get().strip())
            self.db.set_setting('openai_api_key',self.openai_api_key_var.get().strip())
            self.db.set_setting('openai_model',self.openai_model_var.get().strip())
            self.db.set_setting('anthropic_api_key',self.anthropic_api_key_var.get().strip())
            self.db.set_setting('anthropic_model',self.anthropic_model_var.get().strip())
            self.db.set_setting('custom_openai_base_url',self.custom_openai_base_url_var.get().strip())
            self.db.set_setting('custom_openai_api_key',self.custom_openai_api_key_var.get().strip())
            self.db.set_setting('custom_openai_model',self.custom_openai_model_var.get().strip())
            self.db.set_setting('use_litellm','1' if self.use_litellm_var.get() else '0')
            self.db.set_setting('use_langmem','1' if self.use_langmem_var.get() else '0')
            self.db.set_setting('memory_top_k',self.memory_top_k_var.get().strip() or '6')
        except Exception: pass

    def save_ollama_settings(self):
        self.db.set_setting('ai_provider',self.ai_provider_var.get().strip())
        self.db.set_setting('ollama_url',self.ollama_url_var.get().strip())
        self.db.set_setting('ollama_model',self.ollama_model_var.get().strip())
        self.db.set_setting('openai_api_key',self.openai_api_key_var.get().strip())
        self.db.set_setting('openai_model',self.openai_model_var.get().strip())
        self.db.set_setting('anthropic_api_key',self.anthropic_api_key_var.get().strip())
        self.db.set_setting('anthropic_model',self.anthropic_model_var.get().strip())
        self.db.set_setting('custom_openai_base_url',self.custom_openai_base_url_var.get().strip())
        self.db.set_setting('custom_openai_api_key',self.custom_openai_api_key_var.get().strip())
        self.db.set_setting('custom_openai_model',self.custom_openai_model_var.get().strip())
        self.db.set_setting('use_litellm','1' if self.use_litellm_var.get() else '0')
        self.db.set_setting('use_langmem','1' if self.use_langmem_var.get() else '0')
        self.db.set_setting('memory_top_k',self.memory_top_k_var.get().strip() or '6')
        messagebox.showinfo('保存','AI設定を保存しました')

    def test_selected_ai_async(self):
        self.save_ollama_settings_silent()  # テスト前に自動保存
        snap = self.snapshot_ai_settings()
        threading.Thread(target=self.test_selected_ai_worker, args=(snap,), daemon=True).start()

    def test_selected_ai_worker(self, snap=None):
        snap = snap or self.snapshot_ai_settings()
        try:
            provider=snap.get('provider','ollama')
            model=snap.get('ollama_model') if provider=='ollama' else snap.get('anthropic_model') if provider=='anthropic' else snap.get('openai_model') if provider=='openai' else snap.get('custom_openai_model')
            # Ollama大型モデルは初回ロードに60秒以上かかる場合がある
            # show_model()は不要なオーバーヘッドのため除去
            # タイムアウト: Ollama=300秒(大型モデル対応), API系=30秒
            test_timeout = 300 if provider == 'ollama' else 30
            self.ui(lambda m=model,p=provider: self.status.config(
                text=f'AIテスト中: {p}/{m} (最大{300 if p=="ollama" else 30}秒待機中...)'
            ))
            ans = UnifiedAIClient(snap).generate(
                f'接続テストです。「接続OK」とだけ答えてください。',
                timeout=test_timeout
            )
            def ok():
                try:
                    if hasattr(self, 'img_ai_status_var'):
                        self.img_ai_status_var.set(f"AI状態: {provider} / {model} 接続OK")
                except Exception:
                    pass
                messagebox.showinfo('AI接続OK', f'provider={provider}\nmodel={model}\n\n{str(ans)[:500]}')
            self.ui(ok)
        except Exception as e:
            msg = str(e)
            def ng():
                try:
                    if hasattr(self, 'img_ai_status_var'):
                        self.img_ai_status_var.set('AI状態: 接続エラー')
                except Exception:
                    pass
                messagebox.showwarning(
                    'AI接続エラー',
                    f'provider={snap.get("provider","")}\n'
                    f'model={snap.get("ollama_model","") or snap.get("anthropic_model","") or snap.get("openai_model","")}\n\n'
                    f'原因候補:\n'
                    f'1. そのモデルがまだOllamaにpullされていない\n'
                    f'2. 初回ロードが重く60秒を超えている\n'
                    f'3. VRAM/RAM不足\n'
                    f'4. Ollama側で該当モデルが破損\n\n'
                    f'error={msg[:1200]}'
                )
            self.ui(ng)


    def update_usage_display(self, resp):
        try:
            if hasattr(resp,'usage_summary'):
                self.ai_usage_var.set(resp.usage_summary())
            else:
                self.ai_usage_var.set('使用量: 取得不可')
        except Exception:
            pass
    def refresh_openai_models_async(self):
        key=self.openai_api_key_var.get().strip()
        def worker(): return list_openai_models(key)
        def done(models):
            self.openai_model_combo.config(values=models)
            if models and self.openai_model_var.get() not in models: self.openai_model_var.set(models[0])
            self.log('OpenAI models: '+', '.join(models[:80]))
        self.run_in_thread(worker,done)
    def refresh_anthropic_models_async(self):
        key=self.anthropic_api_key_var.get().strip()
        def worker(): return list_anthropic_models(key)
        def done(models):
            vals=models or CLAUDE_MODEL_CHOICES
            self.anthropic_model_combo.config(values=vals)
            if vals and self.anthropic_model_var.get() not in vals: self.anthropic_model_var.set(vals[0])
            self.log('Claude models: '+', '.join(vals[:80]))
        self.run_in_thread(worker,done)

    def migrate_sqlite_to_rqlite_dialog(self):
        if not messagebox.askyesno(
            'SQLite→rqlite移行',
            '現在のSQLiteデータをrqliteへコピーします。\n\n'
            '既存のrqliteテーブルに同じデータがある場合、重複する可能性があります。\n'
            '最初は空のrqliteで実行することを推奨します。\n\n続行しますか？'
        ):
            return
        sqlite_path = DB_PATH
        rqlite_url = normalize_rqlite_url(self.rqlite_url_var.get() if hasattr(self,'rqlite_url_var') else self.db.get_setting('rqlite_url','http://127.0.0.1:4001'))
        def worker():
            return self.migrate_sqlite_to_rqlite(sqlite_path, rqlite_url)
        def done(result):
            msg='\n'.join([f'{k}: {v}' for k,v in result.items()])
            messagebox.showinfo('移行完了', msg)
            self.log('SQLite→rqlite移行完了\n'+msg)
        def err(tb):
            messagebox.showerror('移行エラー', tb[-2200:])
            self.log(tb)
        self.run_in_thread(worker, done, err)

    def migrate_sqlite_to_rqlite(self, sqlite_path, rqlite_url):
        requests=try_import_requests()
        if not requests:
            raise RuntimeError('requests が必要です')
        rqlite_url=normalize_rqlite_url(rqlite_url)
        try:
            st=requests.get(rqlite_url+'/status',timeout=8)
            st.raise_for_status()
        except Exception as e:
            raise RuntimeError(f'rqliteに接続できません: {rqlite_url}\n先に rqlited.exe を起動してください。\n{e}')
        src_con=sqlite3.connect(sqlite_path)
        src_con.row_factory=sqlite3.Row
        cur=src_con.cursor()
        tables=[]
        for row in cur.execute("SELECT name, sql FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%' ORDER BY name").fetchall():
            if row['sql']:
                tables.append((row['name'], row['sql']))
        def post_execute(stmts,timeout=60):
            r=requests.post(rqlite_url+'/db/execute',json=stmts,timeout=timeout)
            r.raise_for_status()
            data=r.json()
            if isinstance(data,dict) and data.get('error'):
                raise RuntimeError(data.get('error'))
            if isinstance(data,dict):
                for res in data.get('results',[]) or []:
                    if isinstance(res,dict) and res.get('error'):
                        raise RuntimeError(res.get('error'))
            return data
        result={'tables':len(tables),'inserted_rows':0,'skipped_tables':0}
        for name,create_sql in tables:
            safe_sql=re.sub(r'CREATE TABLE\s+', 'CREATE TABLE IF NOT EXISTS ', create_sql, count=1, flags=re.I)
            try:
                post_execute([[safe_sql]])
            except Exception as e:
                result[f'create_error_{name}']=str(e)[:220]
                result['skipped_tables']+=1
        for name,create_sql in tables:
            try:
                rows=cur.execute(f'SELECT * FROM "{name}"').fetchall()
                if not rows:
                    result[name]=0
                    continue
                cols=rows[0].keys()
                placeholders=','.join(['?']*len(cols))
                col_sql=','.join([f'"{c}"' for c in cols])
                insert_sql=f'INSERT INTO "{name}" ({col_sql}) VALUES ({placeholders})'
                batch=[]; inserted=0
                for r in rows:
                    vals=[sqlite_value_to_rqlite(r[c]) for c in cols]
                    batch.append([insert_sql]+vals)
                    if len(batch)>=100:
                        post_execute(batch,timeout=90); inserted+=len(batch); batch=[]
                if batch:
                    post_execute(batch,timeout=90); inserted+=len(batch)
                result[name]=inserted
                result['inserted_rows']+=inserted
            except Exception as e:
                result[f'insert_error_{name}']=str(e)[:220]
        src_con.close()
        return result

    def show_rqlite_startup_help(self):
        messagebox.showinfo(
            'rqlite起動方法',
            'PowerShell例:\n\n'
            'mkdir E:\\rqlite\\data\n'
            'E:\\rqlite\\rqlited.exe E:\\rqlite\\data\n\n'
            '別PowerShellで確認:\n'
            'curl http://127.0.0.1:4001/status\n\n'
            'rqlite接続で起動する場合:\n'
            '$env:ESTIMATION_DB_BACKEND="rqlite"\n'
            '$env:RQLITE_URL="http://127.0.0.1:4001"\n'
            'py electrical_estimation_v13_3_rqlite_migration_compact_snap.py'
        )

    def save_vector_settings(self):
        self.db.set_setting('vector_backend', self.vector_backend_var.get().strip() or 'auto')
        self.db.set_setting('clip_model_name', self.clip_model_name_var.get().strip() or CLIP_MODEL_NAME_DEFAULT)
        self.db.set_setting('vector_threshold', self.vector_threshold_var.get().strip() or '0.72')
        self.db.set_setting('use_vector_search','1')
        self.vector_engine=SymbolVectorEngine(self.db,self.db.get_setting('clip_model_name',CLIP_MODEL_NAME_DEFAULT),self.db.get_setting('vector_backend','auto'),self.log)
        messagebox.showinfo('Vector設定保存','FAISS/CLIP設定を保存しました')

    def save_db_backend_settings(self):
        self.db.set_setting('db_backend', self.db_backend_var.get().strip() or 'sqlite')
        self.db.set_setting('rqlite_url', RqliteConnection.normalize_url(self.rqlite_url_var.get().strip() or 'http://127.0.0.1:4001'))
        messagebox.showinfo(
            'DB設定保存',
            'DB設定を保存しました。\n\n'
            '安全のため、実際のrqlite切替は次回起動時の環境変数で行います。\n\n'
            'PowerShell例:\n'
            '$env:ESTIMATION_DB_BACKEND="rqlite"\n'
            '$env:RQLITE_URL="http://127.0.0.1:4001"\n'
            'py electrical_estimation_v13_1_ui_rqlite.py'
        )

    def test_rqlite_connection(self):
        def worker():
            url=(self.rqlite_url_var.get().strip() if hasattr(self,'rqlite_url_var') else '') or 'http://127.0.0.1:4001'
            con=RqliteConnection(url, timeout=5)
            status=con.status()
            rows=con.execute('SELECT 1')
            return con.base_url, status, rows
        def done(res):
            url,status,rows=res
            try:
                version=status.get('version','unknown') if isinstance(status,dict) else 'unknown'
                store=status.get('store',{}) if isinstance(status,dict) else {}
                addr=store.get('addr','') if isinstance(store,dict) else ''
            except Exception:
                version='unknown'; addr=''
            messagebox.showinfo('rqlite接続OK', f'rqliteに接続できました。\nURL: {url}\nversion: {version}\naddr: {addr}\nSQL結果: {rows}')
        def err(tb):
            # tracebackをそのまま出すと読みにくいため、最後のRuntimeError本文を優先して表示する。
            msg=tb
            if 'RuntimeError:' in tb:
                msg=tb.split('RuntimeError:',1)[1].strip()
            messagebox.showerror('rqlite接続エラー', msg[-2200:])
        self.run_in_thread(worker, done, err)

    def show_rqlite_start_guide(self):
        messagebox.showinfo(
            'rqlite起動方法',
            'rqliteはSQLiteファイルではなく、別途起動するDBサーバです。\n\n'
            '1. rqliteをダウンロードして展開\n'
            '2. rqlited.exe を起動\n\n'
            'PowerShell例:\n'
            '  mkdir E:\\rqlite\\node1\n'
            '  E:\\rqlite\\rqlited.exe E:\\rqlite\\node1\n\n'
            '起動できたらブラウザで確認:\n'
            '  http://127.0.0.1:4001/status\n\n'
            'その後、このアプリの rqlite接続テストを押してください。'
        )

    def refresh_models_async(self):
        snap=self.snapshot_ai_settings()
        threading.Thread(target=self.refresh_models_worker,args=(snap,),daemon=True).start()

    def refresh_models_worker(self, snap=None):
        snap = snap or self.snapshot_ai_settings()
        full = []
        err_text = ""
        try:
            url = snap.get('ollama_url', DEFAULT_OLLAMA_URL)
            models = OllamaClient(url).models()
            for mm in models:
                if mm and mm not in full:
                    full.append(mm)
        except Exception as e:
            err_text = str(e)

        # /api/tagsで取れない場合は ollama list でも取得
        if not full:
            fallback = list_ollama_models_fallback()
            for mm in fallback:
                if mm and mm not in full:
                    full.append(mm)

        def done():
            try:
                self.model_combo.config(values=full)
                if full:
                    current = self.ollama_model_var.get()
                    # 既存選択が一覧にあるなら保持。なければ先頭へ。
                    if (not current) and full:
                        self.ollama_model_var.set(full[0])
                    self.status.config(text=f'Ollamaモデル取得: {len(full)}件')
                    self.log('Ollama models: ' + ', '.join(full))
                else:
                    self.status.config(text='Ollamaモデル取得失敗')
                    self.log('Ollamaモデル取得失敗: ' + err_text)
            except Exception:
                self.log(traceback.format_exc())
        self.ui(done)

    def test_ollama_async(self):
        snap=self.snapshot_ai_settings()
        threading.Thread(target=self.test_ollama_worker,args=(snap,),daemon=True).start()

    def test_ollama_worker(self, snap=None):
        snap = snap or self.snapshot_ai_settings()
        try:
            models = []
            try:
                models = OllamaClient(snap.get('ollama_url', DEFAULT_OLLAMA_URL)).models()
            except Exception:
                models = list_ollama_models_fallback()
            msg = '\n'.join(models[:50]) if models else 'モデルを取得できませんでした。\nollama serve または ollama list を確認してください。'
            self.ui(lambda msg=msg: messagebox.showinfo('Ollamaモデル一覧', msg))
        except Exception as e:
            self.log('Ollama接続不可: ' + str(e))
            self.ui(lambda: messagebox.showwarning(
                'Ollama未接続',
                'Ollamaサーバーに接続できません。\n\n'
                'PowerShellで次を実行してください:\n'
                'E:\\Ollama\\ollama.exe serve\n\n'
                'または別PowerShellで:\n'
                'E:\\Ollama\\ollama.exe list'
            ))

    def summary_json(self): return json.dumps({'pdf_dxf_estimate':[self.tree.item(i,'values') for i in self.tree.get_children()],'image_estimate':self.image_estimate_rows(),'image_detections':[{'id':d.det_id,'page':d.page,'equipment':d.equipment,'color':d.color_name,'enabled':d.enabled,'bbox':[d.x1,d.y1,d.x2,d.y2]} for d in self.image_dets[:80]]},ensure_ascii=False,indent=2)
    def append_summary_prompt(self): self.ai_prompt.insert(tk.END,'\n\n--- 現在の結果 ---\n'+self.summary_json())
    def run_ollama_async(self):
        snap=self.snapshot_ai_settings(); prompt_text=self.ai_prompt.get('1.0',tk.END).strip()+'\n\n'+self.summary_json()
        threading.Thread(target=self.run_ollama_worker,args=(snap,prompt_text),daemon=True).start()
    def run_ollama_worker(self,snap,prompt_text):
        try:
            ans=UnifiedAIClient(snap).generate(prompt_text,timeout=120)
            self.ui(lambda ans=ans:(self.ai_result.insert(tk.END,'\n--- AI解析 ---\n'+str(ans)+'\n\n--- 使用量 ---\n'+(ans.usage_summary() if hasattr(ans,'usage_summary') else '')+'\n'), self.update_usage_display(ans)))
        except Exception:
            tb=traceback.format_exc(); self.ui(lambda tb=tb:self.ai_result.insert(tk.END,tb+'\n'))
    def build_chat_tab(self):
        f=self.tab_chat; self.chat_history=[]; self._chat_busy=False
        hdr=ttk.Frame(f); hdr.pack(fill=tk.X,padx=6,pady=4)
        ttk.Label(hdr,text='AIチャット（多ターン会話）',font=('Arial',11,'bold')).pack(side=tk.LEFT)
        ttk.Button(hdr,text='会話クリア',command=self.chat_clear).pack(side=tk.RIGHT,padx=4)
        ttk.Button(hdr,text='積算結果をコンテキスト追加',command=self.chat_add_context).pack(side=tk.RIGHT,padx=4)
        spf=ttk.LabelFrame(f,text='システムプロンプト',padding=4); spf.pack(fill=tk.X,padx=6,pady=2)
        self.chat_system_var=tk.StringVar(value='あなたは日本の電気設備図面と積算の専門家AIです。日本語で丁寧に回答してください。')
        ttk.Entry(spf,textvariable=self.chat_system_var).pack(fill=tk.X)
        hf=ttk.LabelFrame(f,text='会話履歴',padding=4); hf.pack(fill=tk.BOTH,expand=True,padx=6,pady=4)
        self.chat_display=scrolledtext.ScrolledText(hf,state='disabled',wrap='word',font=('Arial',10)); self.chat_display.pack(fill=tk.BOTH,expand=True)
        self.chat_display.tag_config('user',foreground='#005580',font=('Arial',10,'bold'))
        self.chat_display.tag_config('assistant',foreground='#1a6600')
        self.chat_display.tag_config('system',foreground='#888888',font=('Arial',9,'italic'))
        inf=ttk.Frame(f,padding=4); inf.pack(fill=tk.X,padx=6,pady=4)
        self.chat_input=scrolledtext.ScrolledText(inf,height=4,wrap='word',font=('Arial',10)); self.chat_input.pack(side=tk.LEFT,fill=tk.X,expand=True,padx=(0,4))
        self.chat_input.bind('<Control-Return>',lambda e:self.send_chat_async())
        vbtn=ttk.Frame(inf); vbtn.pack(side=tk.RIGHT,fill=tk.Y)
        self.chat_send_btn=ttk.Button(vbtn,text='送信\n(Ctrl+Enter)',width=12,command=self.send_chat_async); self.chat_send_btn.pack(pady=2)
        ttk.Label(vbtn,text='接続設定は\n「Ollama解析」タブ',font=('Arial',8),foreground='gray').pack(pady=2)
        self._chat_append('[システム] AIチャットへようこそ。「Ollama解析」タブでAIを設定してから使用してください。\n','system')
    def _chat_append(self,text,tag=''):
        self.chat_display.configure(state='normal')
        self.chat_display.insert(tk.END,text,tag) if tag else self.chat_display.insert(tk.END,text)
        self.chat_display.see(tk.END); self.chat_display.configure(state='disabled')
    def send_chat_async(self):
        msg=self.chat_input.get('1.0',tk.END).strip()
        if not msg: return
        if getattr(self,'_chat_busy',False): messagebox.showinfo('処理中','前のメッセージを処理中です。'); return
        self._chat_busy=True; self.chat_send_btn.config(state='disabled')
        self.chat_input.delete('1.0',tk.END)
        self.chat_history.append({'role':'user','content':msg})
        self._chat_append(f'\n【あなた】\n{msg}\n','user')
        snap=self.snapshot_ai_settings(); system_prompt=self.chat_system_var.get().strip() or 'あなたは電気設備積算の専門家AIです。'
        history_copy=list(self.chat_history)
        threading.Thread(target=self._send_chat_worker,args=(snap,system_prompt,history_copy),daemon=True).start()
    def _send_chat_worker(self,snap,system_prompt,history):
        try:
            requests=try_import_requests()
            if not requests: raise RuntimeError('requestsライブラリが必要です')
            provider=snap.get('provider','ollama')
            if provider=='ollama':
                url=snap.get('ollama_url',DEFAULT_OLLAMA_URL).rstrip('/'); model=snap.get('ollama_model',DEFAULT_OLLAMA_MODEL)
                messages=[{'role':'system','content':system_prompt}]+history
                r=requests.post(url+'/api/chat',json={'model':model,'messages':messages,'stream':False,'options':{'num_ctx':4096,'temperature':0.2,'num_predict':1024}},timeout=120); r.raise_for_status()
                ans=r.json().get('message',{}).get('content','（応答なし）')
            elif provider=='openai':
                key=snap.get('openai_api_key',''); model=snap.get('openai_model',DEFAULT_OPENAI_MODEL)
                if not key: raise RuntimeError('OpenAI APIキーが未設定です')
                r=requests.post('https://api.openai.com/v1/chat/completions',headers={'Authorization':f'Bearer {key}','Content-Type':'application/json'},json={'model':model,'messages':[{'role':'system','content':system_prompt}]+history,'temperature':0.2,'max_tokens':1024},timeout=120); r.raise_for_status()
                ans=r.json()['choices'][0]['message']['content']
            elif provider=='anthropic':
                key=snap.get('anthropic_api_key',''); model=snap.get('anthropic_model',DEFAULT_CLAUDE_MODEL)
                if not key: raise RuntimeError('Anthropic APIキーが未設定です')
                r=requests.post('https://api.anthropic.com/v1/messages',headers={'x-api-key':key,'anthropic-version':'2023-06-01','content-type':'application/json'},json={'model':model,'max_tokens':1024,'system':system_prompt,'messages':history},timeout=120); r.raise_for_status()
                ans='\n'.join(p.get('text','') for p in r.json().get('content',[]) if p.get('type')=='text')
            else:
                base=snap.get('custom_openai_base_url',DEFAULT_CUSTOM_OPENAI_BASE_URL).rstrip('/'); key=snap.get('custom_openai_api_key',''); model=snap.get('custom_openai_model','local-model')
                hdrs={'Content-Type':'application/json'}
                if key: hdrs['Authorization']=f'Bearer {key}'
                r=requests.post(base+'/chat/completions',headers=hdrs,json={'model':model,'messages':[{'role':'system','content':system_prompt}]+history,'temperature':0.2,'max_tokens':1024},timeout=120); r.raise_for_status()
                ans=r.json()['choices'][0]['message']['content']
            self.chat_history.append({'role':'assistant','content':ans})
            def done(ans=ans,provider=provider):
                self._chat_busy=False; self.chat_send_btn.config(state='normal')
                self._chat_append(f'\n【AI ({provider})】\n{ans}\n','assistant')
            self.ui(done)
        except Exception:
            tb=traceback.format_exc()
            def err(tb=tb):
                self._chat_busy=False; self.chat_send_btn.config(state='normal')
                self._chat_append(f'\n[エラー] {tb}\n','system')
                if self.chat_history and self.chat_history[-1]['role']=='user': self.chat_history.pop()
            self.ui(err)
    def chat_clear(self):
        self.chat_history=[]
        self.chat_display.configure(state='normal'); self.chat_display.delete('1.0',tk.END); self.chat_display.configure(state='disabled')
        self._chat_append('[システム] 会話履歴をクリアしました。\n','system')
    def chat_add_context(self):
        ctx=self.summary_json(); self.chat_input.insert(tk.END,f'\n\n【現在の積算結果】\n{ctx}')
        self._chat_append('[システム] 積算結果をメッセージ入力欄に追加しました。内容を確認して送信してください。\n','system')
    def generate_symbol_async(self):
        snap=self.snapshot_ai_settings(); name=self.gen_name_var.get().strip() or 'LED照明'; cat=self.gen_category_var.get().strip(); kind=self.gen_kind_var.get().strip()
        threading.Thread(target=self._gen_worker,args=(snap,name,cat,kind,False),daemon=True).start()
    def generate_family_async(self):
        snap=self.snapshot_ai_settings(); name=self.gen_name_var.get().strip() or 'LED照明'; cat=self.gen_category_var.get().strip(); kind=self.gen_kind_var.get().strip()
        threading.Thread(target=self._gen_worker,args=(snap,name,cat,kind,True),daemon=True).start()
    def _gen_worker(self,snap,name,cat,kind,family):
        try:
            prompt=self.make_symbol_generation_prompt(family=family)
            body=UnifiedAIClient(snap).generate(prompt,timeout=120)
            def done(body=body,name=name,cat=cat,kind=kind):
                self.generated_asset_text.delete('1.0',tk.END); self.generated_asset_text.insert('1.0',body)
                self._last_gen={'name':name,'category':cat,'kind':kind,'body':body,'family':family}
                self.status.config(text=f'AI生成完了: {name}')
            self.ui(done)
        except Exception:
            tb=traceback.format_exc(); self.ui(lambda tb=tb:messagebox.showerror('AI生成エラー',tb[-800:]))
    def register_generated_asset(self):
        g=getattr(self,'_last_gen',None)
        if not g: messagebox.showwarning('なし','先にAI生成を実行してください'); return
        body=g['body']; name=g['name']; cat=g['category']; kind=g['kind']
        dxf=svg=json_spec=''
        for tag,attr in [('DXF_ENTITIES','dxf'),('SVG','svg'),('JSON','json_spec')]:
            m=re.search(rf'{tag}\s*\n(.*?)(?=\n[A-Z_]+\n|\Z)',body,re.S)
            if m: exec(f"{attr}=m.group(1).strip()")
        if g.get('family'): self.db.save_generated_family(name,cat,kind[:80],dxf,svg,json_spec,body)
        else: self.db.save_generated_symbol(name,cat,kind,kind[:80],dxf,'',svg,json_spec,body)
        self.refresh_generated_assets(); messagebox.showinfo('登録完了',f'{name} をDBに登録しました')
    def refresh_generated_assets(self):
        for i in self.generated_tree.get_children(): self.generated_tree.delete(i)
        for r in self.db.get_generated_symbols():
            self.generated_tree.insert('',tk.END,values=(r[0],r[1],r[2],r[3],f'{r[4]} - {r[5]}'))
        for r in self.db.get_generated_families():
            self.generated_tree.insert('',tk.END,values=(r[0],r[1],r[2],r[3],'2Dファミリ - '+r[4]))
    def refresh_learning_summary(self):
        rows=self.db.annotation_summary()
        feat_rows=self.db.annotation_feature_summary()
        lines=['図記号アノテーション学習サマリー','='*50]
        try:
            lines.append(f'symbol_cropsフォルダ画像数: {len([p for p in CROP_DIR.glob("*") if p.suffix.lower() in (".png",".jpg",".jpeg",".bmp")])}')
            lines.append(f'特徴量DB件数: {self.db.count_annotation_features() if hasattr(self.db,"count_annotation_features") else "?"}')
        except Exception:
            pass
        if not rows:
            lines.append('まだ学習データがありません。')
        else:
            lines.append('[アノテーション履歴]')
            for name,cnt,ok in rows:
                lines.append(f'{name}: {cnt}件 / LLM OK {ok or 0}件')
        lines.append('')
        lines.append('[再学習・再推論用 特徴量]')
        if not feat_rows:
            lines.append('特徴量データなし。過去アノテーションから特徴量再構築を実行してください。')
        else:
            for name,cnt in feat_rows:
                lines.append(f'{name}: {cnt}件')
        mem_rows=[]
        try:
            mem_rows=self.db.memory_summary()
        except Exception:
            mem_rows=[]
        lines.append('')
        lines.append('[LangMem互換 長期記憶]')
        if not mem_rows:
            lines.append('長期記憶データなし。OK/NG登録やAI積算で記録されます。')
        else:
            for ns,kind,cnt in mem_rows:
                lines.append(f'{ns}/{kind}: {cnt}件')
        lines += [
            '',
            '使い方:',
            '1. 図記号をOK/NG登録すると、crop画像の特徴量も保存されます。',
            '2. 次回PDF/DXF/画像を解析すると、保存済み特徴量と類似検索してラベルを自動反映します。',
            '3. 既存の古い学習データは「過去アノテーションから特徴量再構築」で反映できます。',
            f'4. 現在の学習反映しきい値: {self.db.get_setting("learning_match_threshold","0.68")}',
            '',
            f'DB: {DB_PATH}',
            f'切り出し画像: {CROP_DIR}'
        ]
        self.learning_text.delete('1.0',tk.END)
        self.learning_text.insert('1.0','\\n'.join(lines))
    def export_learning_csv(self):
        p=filedialog.asksaveasfilename(defaultextension='.csv',filetypes=[('CSV','*.csv')])
        if not p: return
        con=self.db.connect(); con.row_factory=sqlite3.Row; rows=con.execute('SELECT * FROM annotation_samples ORDER BY id').fetchall(); con.close()
        with open(p,'w',encoding='utf-8-sig',newline='') as f:
            if rows:
                w=csv.writer(f); w.writerow(rows[0].keys())
                for r in rows: w.writerow([r[k] for k in r.keys()])
            else: f.write('no data\n')
        messagebox.showinfo('保存',p)
    def install_pkg_async(self,pkg): threading.Thread(target=self.install_worker,args=(pkg,),daemon=True).start()
    def install_worker(self,pkg):
        self.log(f'pip install {pkg}'); ok,lines=run_pip_install(pkg)
        for ln in lines: self.log(ln)
        self.log('完了' if ok else '失敗')

def main():
    try:
        diagnose_pydroid_environment()
        root=tk.Tk()
        try:
            root.title(APP_TITLE)
            if is_android_pydroid():
                # スマホ画面では初期サイズを抑える
                root.geometry("1000x700")
        except Exception:
            pass
        IntegratedApp(root)
        root.mainloop()
    except Exception:
        tb = traceback.format_exc()
        write_crash_log(tb)
        print(tb)
        try:
            # tkinterが生きている場合のみ表示
            messagebox.showerror(
                "起動エラー",
                "Pydroid3で起動に失敗しました。\n"
                f"ログを確認してください:\n{get_safe_base_dir() / 'pydroid_crash_log.txt'}\n\n"
                + tb[-1200:]
            )
        except Exception:
            pass

if __name__ == '__main__': main()
