#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PDF/DXF 完全自動積算システム（Pydroid安定版 / バイナリDXF対応 / 日本語入力回避付き）

【できること】
- PDF（テキスト抽出）→ 設備名 x 数量 を抽出して積算
- DXF（ASCII/バイナリ両対応 ※ezdxf使用）→ 図形パターン + テキスト抽出で積算
- 単価マスター編集（DB）
- 記号パターン編集（DB）
- CADライブラリ（DXF登録・URL/ZIPダウンロード）

【重要な注意（ユーザー要望2について）】
- 「1レイヤーPDFを電気/建築でレイヤー分割」は、PDFが “ベクター図形” の場合のみ
  ある程度可能です（スキャン画像PDFは不可）。
- 本実装ではオプションとして「PDFをARCH/ELECに擬似分類してDXF出力」を追加。
  ※精度は図面の作り方に依存し、案件ごとの調整が必要です。

【Pydroidで日本語入力できない問題の回避】
- 入力欄の横に「貼り付け」ボタンを付けています（他アプリで入力→コピー→貼り付け運用）
"""

import sys
import os
import re
import csv
import sqlite3
import subprocess
import threading
import queue
import time
import json
import math
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter import scrolledtext
from collections import defaultdict
import unicodedata
import zipfile
import io
import shutil
import platform

APP_TITLE = "PDF/DXF自動積算システム（Pydroid安定 / BinaryDXF対応）"
DB_PATH = "estimation_master.db"
CAD_LIBRARY_PATH = "cad_library"

# -----------------------------
#  便利：貼り付けボタン（Pydroid日本語入力回避）
# -----------------------------
def add_paste_button(parent, widget, width=8):
    def paste():
        try:
            text = parent.clipboard_get()
        except Exception:
            text = ""
        if not text:
            return
        try:
            # Entry
            widget.delete(0, tk.END)
            widget.insert(0, text)
        except Exception:
            # Text
            try:
                widget.delete("1.0", tk.END)
                widget.insert("1.0", text)
            except Exception:
                pass

    ttk.Button(parent, text="貼り付け", command=paste, width=width).pack(side=tk.LEFT, padx=3)


# -----------------------------
#  文字コード判定（ASCII DXF対策のフォールバック用）
# -----------------------------
def try_import_chardet():
    try:
        import chardet
        return chardet
    except Exception:
        return None

def detect_encoding(file_path):
    ch = try_import_chardet()
    if not ch:
        return "shift-jis"
    try:
        with open(file_path, "rb") as f:
            raw = f.read(50000)
        r = ch.detect(raw)
        enc = r.get("encoding", "shift-jis")
        conf = r.get("confidence", 0)
        if enc and ("shift" in enc.lower() or "cp932" in enc.lower()):
            return "cp932"
        if enc and conf > 0.8:
            return enc
        return "shift-jis"
    except Exception:
        return "shift-jis"

def read_text_file_safe(file_path):
    encs = ["cp932", "shift-jis", "utf-8", "euc-jp", "iso-2022-jp", "latin1"]
    det = detect_encoding(file_path)
    if det in encs:
        encs.remove(det)
    encs.insert(0, det)

    for enc in encs:
        try:
            with open(file_path, "r", encoding=enc, errors="ignore") as f:
                s = f.read()
            if s.strip() and len(s) > 50:
                return s, enc
        except Exception:
            continue

    try:
        with open(file_path, "rb") as f:
            raw = f.read()
        return raw.decode("cp932", errors="ignore"), "cp932"
    except Exception:
        return "", "unknown"


# -----------------------------
#  ライブラリ検出
# -----------------------------
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


# -----------------------------
#  pip インストール（ログ付き / タイムアウト）
# -----------------------------
def run_pip_install(package, extra_args=None, timeout_sec=180):
    if extra_args is None:
        extra_args = []
    cmd = [sys.executable, "-m", "pip", "install", package] + list(extra_args)
    proc = subprocess.Popen(
        cmd,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        bufsize=1,
        universal_newlines=True,
    )
    start = time.time()
    lines = []
    try:
        while True:
            if time.time() - start > timeout_sec:
                try:
                    proc.kill()
                except Exception:
                    pass
                return False, lines + [f"[TIMEOUT] {timeout_sec}s"]
            line = proc.stdout.readline() if proc.stdout else ""
            if line:
                lines.append(line.rstrip("\n"))
            else:
                if proc.poll() is not None:
                    break
                time.sleep(0.05)
        ok = (proc.wait() == 0)
        return ok, lines
    except Exception as e:
        try:
            proc.kill()
        except Exception:
            pass
        return False, lines + [f"[EXCEPTION] {e}"]


# -----------------------------
#  DXF解析（バイナリ対応：ezdxf優先）
# -----------------------------
class AdvancedDXFParser:
    """
    - ezdxf が使える場合：ASCII/Binary DXFを安定して解析
    - 使えない場合：ASCII（テキスト）DXFだけ簡易解析
    """

    def __init__(self):
        self.entities = []
        self.circles = []
        self.lines = []
        self.hatches = []
        self.texts = []
        self.blocks = []
        self.encoding_used = None

    def parse_dxf(self, dxf_path):
        ezdxf = try_import_ezdxf()
        if ezdxf is not None:
            try:
                doc = ezdxf.readfile(dxf_path)
                msp = doc.modelspace()
                self.encoding_used = "ezdxf"

                self.entities, self.circles, self.lines, self.hatches, self.texts, self.blocks = [], [], [], [], [], []

                for e in msp:
                    et = e.dxftype()
                    layer = getattr(e.dxf, "layer", "0") if hasattr(e, "dxf") else "0"

                    if et == "CIRCLE":
                        c = {
                            "type": "CIRCLE",
                            "layer": layer,
                            "data": {
                                "x": float(e.dxf.center.x),
                                "y": float(e.dxf.center.y),
                                "radius": float(e.dxf.radius),
                            },
                        }
                        self.entities.append(c)
                        self.circles.append(c)

                    elif et == "LINE":
                        l = {
                            "type": "LINE",
                            "layer": layer,
                            "data": {
                                "x": float(e.dxf.start.x),
                                "y": float(e.dxf.start.y),
                                "x2": float(e.dxf.end.x),
                                "y2": float(e.dxf.end.y),
                            },
                        }
                        self.entities.append(l)
                        self.lines.append(l)

                    elif et in ("LWPOLYLINE", "POLYLINE"):
                        pts = []
                        try:
                            pts = [(float(p[0]), float(p[1])) for p in e.get_points()]  # lwpolyline
                        except Exception:
                            try:
                                pts = [(float(v.dxf.location.x), float(v.dxf.location.y)) for v in e.vertices()]
                            except Exception:
                                pts = []
                        if pts:
                            x, y = pts[0]
                            x2, y2 = pts[-1]
                        else:
                            x = y = x2 = y2 = 0.0
                        pl = {"type": et, "layer": layer, "data": {"x": x, "y": y, "x2": x2, "y2": y2}}
                        self.entities.append(pl)
                        self.lines.append(pl)

                    elif et == "HATCH":
                        h = {"type": "HATCH", "layer": layer, "data": {}}
                        self.entities.append(h)
                        self.hatches.append(h)

                    elif et in ("TEXT", "MTEXT"):
                        txt = ""
                        try:
                            txt = e.dxf.text if et == "TEXT" else e.text
                        except Exception:
                            txt = ""
                        txt = txt or ""
                        self.texts.append(txt)
                        t = {"type": et, "layer": layer, "data": {"text": txt}}
                        self.entities.append(t)

                    elif et == "INSERT":
                        bname = getattr(e.dxf, "name", "")
                        ins = {"type": "INSERT", "layer": layer, "data": {"block_name": bname}}
                        self.entities.append(ins)
                        self.blocks.append(ins)

                return  # ezdxfで完了
            except Exception:
                # ezdxfが失敗したらASCII簡易へ
                pass

        # ---- ASCII簡易パース（フォールバック） ----
        content, enc = read_text_file_safe(dxf_path)
        self.encoding_used = enc
        if not content:
            raise RuntimeError(f"DXFファイルを読み込めませんでした: {dxf_path}")

        lines = content.split("\n")
        in_entities = False
        in_blocks = False
        current_entity = None
        current_code = None

        self.entities, self.circles, self.lines, self.hatches, self.texts, self.blocks = [], [], [], [], [], []

        for line in lines:
            s = line.strip()
            if s == "ENTITIES":
                in_entities, in_blocks = True, False
                continue
            if s == "BLOCKS":
                in_blocks, in_entities = True, False
                continue
            if s == "ENDSEC":
                if current_entity:
                    self._classify_entity(current_entity, in_blocks)
                in_entities, in_blocks = False, False
                continue
            if not (in_entities or in_blocks):
                continue

            if s.lstrip("-").isdigit():
                current_code = int(s)
                continue

            if current_code is None:
                continue

            if current_code == 0:
                if current_entity:
                    self._classify_entity(current_entity, in_blocks)
                current_entity = {"type": s, "layer": "0", "data": {}}
            elif current_code == 8:
                if current_entity:
                    current_entity["layer"] = s
            elif current_code == 1:
                if current_entity:
                    current_entity["data"]["text"] = s
                    self.texts.append(s)
            elif current_code == 2:
                if current_entity:
                    current_entity["data"]["block_name"] = s
            elif current_code in (10, 11):
                if current_entity:
                    key = "x" if current_code == 10 else "x2"
                    try:
                        current_entity["data"][key] = float(s)
                    except Exception:
                        pass
            elif current_code in (20, 21):
                if current_entity:
                    key = "y" if current_code == 20 else "y2"
                    try:
                        current_entity["data"][key] = float(s)
                    except Exception:
                        pass
            elif current_code == 40:
                if current_entity:
                    try:
                        current_entity["data"]["radius"] = float(s)
                    except Exception:
                        pass

            current_code = None

        if current_entity:
            self._classify_entity(current_entity, in_blocks)

    def _classify_entity(self, entity, is_block):
        self.entities.append(entity)
        if is_block:
            self.blocks.append(entity)
        if entity["type"] == "CIRCLE":
            self.circles.append(entity)
        elif entity["type"] in ("LINE", "LWPOLYLINE", "POLYLINE"):
            self.lines.append(entity)
        elif entity["type"] == "HATCH":
            self.hatches.append(entity)
        elif entity["type"] == "INSERT":
            self.blocks.append(entity)

    def extract_pattern_signature(self):
        sig = {
            "circle_count": len(self.circles),
            "line_count": len(self.lines),
            "hatch_count": len(self.hatches),
            "block_count": len(self.blocks),
            "text_count": len(self.texts),
            "circle_radii": [],
            "bounding_box": None,
            "texts": self.texts[:20],
            "encoding": self.encoding_used,
        }

        if self.circles:
            sig["circle_radii"] = sorted([c["data"].get("radius", 0) for c in self.circles])[:10]

        all_x, all_y = [], []
        for e in self.entities:
            d = e.get("data", {})
            if "x" in d:
                all_x.append(d["x"])
            if "x2" in d:
                all_x.append(d["x2"])
            if "y" in d:
                all_y.append(d["y"])
            if "y2" in d:
                all_y.append(d["y2"])
        if all_x and all_y:
            sig["bounding_box"] = {"width": max(all_x) - min(all_x), "height": max(all_y) - min(all_y)}
        return sig

    def get_combined_text(self):
        return "\n".join([t for t in self.texts if t])

    def _check_line_near_circle(self, cx, cy, radius, direction, count):
        tolerance = radius * 1.5 if radius > 0 else 10.0
        found = 0
        for line in self.lines:
            d = line.get("data", {})
            x1 = d.get("x", 0)
            y1 = d.get("y", 0)
            x2 = d.get("x2", x1)
            y2 = d.get("y2", y1)
            mx = (x1 + x2) / 2.0
            my = (y1 + y2) / 2.0
            dist = math.sqrt((mx - cx) ** 2 + (my - cy) ** 2)
            if dist > tolerance:
                continue
            if direction == "vertical":
                if abs(x2 - x1) < max(radius * 0.3, 1.0):
                    found += 1
            elif direction == "horizontal":
                if abs(y2 - y1) < max(radius * 0.3, 1.0):
                    found += 1
            else:
                found += 1
        return found >= count

    def _count_radial_lines(self, cx, cy, radius):
        tol = radius * 0.5 if radius > 0 else 5.0
        cnt = 0
        for line in self.lines:
            d = line.get("data", {})
            x1 = d.get("x", 0)
            y1 = d.get("y", 0)
            x2 = d.get("x2", x1)
            y2 = d.get("y2", y1)
            dist1 = math.sqrt((x1 - cx) ** 2 + (y1 - cy) ** 2)
            dist2 = math.sqrt((x2 - cx) ** 2 + (y2 - cy) ** 2)
            if dist1 < tol or dist2 < tol:
                cnt += 1
        return cnt

    def count_equipment_by_patterns(self, pattern_rules):
        equipment_counts = defaultdict(int)

        for rule in pattern_rules:
            pattern = rule.get("pattern", {})
            ptype = pattern.get("type", "simple_circle")

            if ptype == "simple_circle":
                rmin = pattern.get("radius_min", 0)
                rmax = pattern.get("radius_max", 999999)
                for c in self.circles:
                    r = c["data"].get("radius", 0)
                    if rmin <= r <= rmax:
                        equipment_counts[rule["name"]] += 1

            elif ptype == "circle_with_line":
                rmin = pattern.get("radius_min", 0)
                rmax = pattern.get("radius_max", 999999)
                direction = pattern.get("line_direction", "any")
                need = int(pattern.get("line_count", 1))
                for c in self.circles:
                    r = c["data"].get("radius", 0)
                    if not (rmin <= r <= rmax):
                        continue
                    cx = c["data"].get("x", 0)
                    cy = c["data"].get("y", 0)
                    if self._check_line_near_circle(cx, cy, r, direction, need):
                        equipment_counts[rule["name"]] += 1

            elif ptype == "rectangle_with_cross":
                # ここは簡易：ポリライン/線の密度から推定（案件依存なので保守的）
                approx = len([e for e in self.entities if e["type"] in ("LWPOLYLINE", "POLYLINE")])
                if approx > 0:
                    equipment_counts[rule["name"]] += max(0, min(approx // 5, 50))

            elif ptype == "circle_with_radial_lines":
                rmin = pattern.get("radius_min", 0)
                rmax = pattern.get("radius_max", 999999)
                need = int(pattern.get("min_radial_lines", 3))
                for c in self.circles:
                    r = c["data"].get("radius", 0)
                    if not (rmin <= r <= rmax):
                        continue
                    cx = c["data"].get("x", 0)
                    cy = c["data"].get("y", 0)
                    rc = self._count_radial_lines(cx, cy, r)
                    if rc >= need:
                        equipment_counts[rule["name"]] += 1

        return dict(equipment_counts)


# -----------------------------
#  CADライブラリ（ZIPダウンロード/展開）
# -----------------------------
class CADLibraryDownloader:
    def __init__(self, library_path):
        self.library_path = library_path
        os.makedirs(library_path, exist_ok=True)
        self.requests = try_import_requests()

    def download_zip(self, url, callback=None):
        if not self.requests:
            raise RuntimeError("requestsが必要です（メニュー → ツール → requestsインストール）")

        if callback:
            callback(f"ダウンロード開始: {url}")

        headers = {"User-Agent": "Mozilla/5.0"}
        r = self.requests.get(url, headers=headers, timeout=120, allow_redirects=True)
        r.raise_for_status()

        if callback:
            callback(f"ダウンロード完了: {len(r.content)} bytes")

        extracted_files = []
        try:
            with zipfile.ZipFile(io.BytesIO(r.content)) as zf:
                if callback:
                    callback(f"ZIP内ファイル数: {len(zf.namelist())}")
                for member in zf.namelist():
                    if not (member.lower().endswith(".dxf") or member.lower().endswith(".dwg")):
                        continue

                    filename = member
                    # 日本語ZIP名対策（よくあるCP437→ShiftJIS）
                    try:
                        filename = member.encode("cp437").decode("cp932")
                    except Exception:
                        filename = os.path.basename(member)

                    safe = os.path.basename(filename)
                    safe = re.sub(r'[<>:"/\\|?*]', "_", safe)
                    if not safe:
                        safe = f"cad_{len(extracted_files)}.dxf"

                    out = os.path.join(self.library_path, safe)
                    with zf.open(member) as src, open(out, "wb") as dst:
                        shutil.copyfileobj(src, dst)

                    extracted_files.append(out)
                    if callback:
                        callback(f"展開: {safe}")
        except zipfile.BadZipFile:
            # ZIPでない → そのまま保存
            fn = os.path.basename(url.split("?")[0])
            if not fn.lower().endswith((".dxf", ".dwg")):
                fn = "downloaded_file.dxf"
            out = os.path.join(self.library_path, fn)
            with open(out, "wb") as f:
                f.write(r.content)
            extracted_files.append(out)
            if callback:
                callback(f"保存: {fn}")

        return extracted_files

    def import_local_file(self, source_path):
        fn = os.path.basename(source_path)
        dst = os.path.join(self.library_path, fn)
        shutil.copy2(source_path, dst)
        return dst

    def scan_library(self):
        out = []
        for root, _, files in os.walk(self.library_path):
            for f in files:
                if f.lower().endswith(".dxf"):
                    out.append(os.path.join(root, f))
        return out


# -----------------------------
#  DB
# -----------------------------
class DatabaseManager:
    def __init__(self, db_path):
        self.db_path = db_path
        self.init_db()

    def init_db(self):
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()

        c.execute('''CREATE TABLE IF NOT EXISTS unit_prices
                     (id INTEGER PRIMARY KEY AUTOINCREMENT,
                      category TEXT,
                      item_name TEXT,
                      spec TEXT,
                      unit TEXT,
                      unit_price REAL,
                      keywords TEXT)''')

        c.execute('''CREATE TABLE IF NOT EXISTS symbol_patterns
                     (id INTEGER PRIMARY KEY AUTOINCREMENT,
                      name TEXT,
                      pattern_type TEXT,
                      pattern_json TEXT,
                      description TEXT,
                      preset INTEGER DEFAULT 0,
                      added_at TEXT DEFAULT CURRENT_TIMESTAMP)''')

        c.execute('''CREATE TABLE IF NOT EXISTS cad_library
                     (id INTEGER PRIMARY KEY AUTOINCREMENT,
                      manufacturer TEXT,
                      model_number TEXT,
                      category TEXT,
                      item_name TEXT,
                      file_path TEXT,
                      pattern_signature TEXT,
                      spec_json TEXT,
                      url TEXT,
                      added_at TEXT DEFAULT CURRENT_TIMESTAMP)''')

        c.execute("SELECT COUNT(*) FROM unit_prices")
        if c.fetchone()[0] == 0:
            default_prices = [
                ('照明器具', 'LEDダウンライト',  '埋込型 φ150',    '台', 12000, 'ダウンライト,DL,ＤＬ'),
                ('照明器具', 'LEDベースライト',  '埋込型 40形',    '台', 15000, 'ベースライト,BL,ＢＬ'),
                ('照明器具', 'LED防湿型照明',    '壁付型 20形',    '台', 18000, '防湿,防湿型'),
                ('照明器具', '非常灯',          'LED 20形',      '台', 28000, '非常灯,emergency'),
                ('照明器具', '誘導灯',          'LED 避難口',     '台', 25000, '誘導灯,exit'),
                ('コンセント','コンセント',      '2P 15A 接地極付','個', 3500,  'コンセント,CO,ＣＯ'),
                ('コンセント','防水コンセント',  '2P 15A',        '個', 6500,  '防水コンセント'),
                ('コンセント','フロアコンセント','2P 15A',        '個', 8500,  'フロアコンセント,FC,ＦＣ'),
                ('スイッチ', '片切スイッチ',     '15A 埋込',      '個', 2200,  'スイッチ,SW,ＳＷ'),
                ('スイッチ', '3路スイッチ',      '15A 埋込',      '個', 2800,  '3路スイッチ'),
                ('分電盤',   '分電盤',           '20回路 壁埋込',  '面', 85000, '分電盤,DB,ＤＢ'),
                ('分電盤',   '主配電盤',         '750A',          '面', 450000,'主配電盤,MDB,ＭＤＢ'),
                ('配線',     'VVF',              '2.0-2C',        'm',  250,   'VVF,2.0-2C'),
                ('配線',     'VVF',              '2.0-3C',        'm',  320,   'VVF,2.0-3C'),
                ('配線',     'CV',               '60sq-3C',       'm',  2400,  'CV,60sq'),
                ('空調',     '換気扇',           'φ100',          '台', 18000, '換気扇,VF,ＶＦ'),
                ('空調',     '換気扇',           'φ150',          '台', 25000, '換気扇,VF,ＶＦ'),
                ('空調',     'エアコン',         '5.0kW',         '台', 180000,'エアコン,AC,ＡＣ'),
            ]
            c.executemany("""INSERT INTO unit_prices
                             (category, item_name, spec, unit, unit_price, keywords)
                             VALUES (?, ?, ?, ?, ?, ?)""", default_prices)

        c.execute("SELECT COUNT(*) FROM symbol_patterns")
        if c.fetchone()[0] == 0:
            presets = [
                ('ダウンライト', 'simple_circle',
                 json.dumps({'type': 'simple_circle', 'radius_min': 40, 'radius_max': 120}, ensure_ascii=False),
                 '単純な円形記号 半径40-120', 1),
                ('ベースライト', 'simple_circle',
                 json.dumps({'type': 'simple_circle', 'radius_min': 120, 'radius_max': 220}, ensure_ascii=False),
                 '大きな円形記号 半径120-220', 1),
                ('コンセント', 'circle_with_line',
                 json.dumps({'type': 'circle_with_line', 'radius_min': 30, 'radius_max': 90,
                            'line_direction': 'vertical', 'line_count': 1}, ensure_ascii=False),
                 '円+縦線', 1),
                ('スイッチ', 'circle_with_line',
                 json.dumps({'type': 'circle_with_line', 'radius_min': 30, 'radius_max': 90,
                            'line_direction': 'horizontal', 'line_count': 1}, ensure_ascii=False),
                 '円+横線', 1),
                ('分電盤', 'rectangle_with_cross',
                 json.dumps({'type': 'rectangle_with_cross', 'width_min': 80, 'width_max': 220}, ensure_ascii=False),
                 '矩形+×印（簡易推定）', 1),
                ('照明（放射線）', 'circle_with_radial_lines',
                 json.dumps({'type': 'circle_with_radial_lines', 'radius_min': 40, 'radius_max': 120,
                            'min_radial_lines': 4}, ensure_ascii=False),
                 '円+放射線4本以上', 1),
            ]
            c.executemany("""INSERT INTO symbol_patterns
                             (name, pattern_type, pattern_json, description, preset)
                             VALUES (?, ?, ?, ?, ?)""", presets)

        conn.commit()
        conn.close()

    def get_all_unit_prices(self):
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()
        c.execute("SELECT id, category, item_name, spec, unit, unit_price, keywords FROM unit_prices ORDER BY category, item_name")
        rows = c.fetchall()
        conn.close()
        return rows

    def upsert_unit_price(self, data, record_id=None):
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()
        if record_id:
            c.execute("""UPDATE unit_prices SET category=?, item_name=?, spec=?, unit=?, unit_price=?, keywords=? WHERE id=?""",
                      (data['category'], data['item_name'], data['spec'], data['unit'], data['unit_price'], data['keywords'], record_id))
        else:
            c.execute("""INSERT INTO unit_prices (category, item_name, spec, unit, unit_price, keywords)
                         VALUES (?, ?, ?, ?, ?, ?)""",
                      (data['category'], data['item_name'], data['spec'], data['unit'], data['unit_price'], data['keywords']))
        conn.commit()
        conn.close()

    def delete_unit_price(self, record_id):
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()
        c.execute("DELETE FROM unit_prices WHERE id=?", (record_id,))
        conn.commit()
        conn.close()

    def import_csv(self, csv_path):
        imported = 0
        skipped = 0
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()
        with open(csv_path, "r", encoding="utf-8-sig") as f:
            reader = csv.DictReader(f)
            for row in reader:
                try:
                    unit_price = float(str(row.get("単価", "0")).replace(",", "").strip() or 0)
                    c.execute("""INSERT INTO unit_prices (category, item_name, spec, unit, unit_price, keywords)
                                 VALUES (?, ?, ?, ?, ?, ?)""",
                              (row.get("カテゴリ", ""), row.get("品名", ""), row.get("仕様", ""),
                               row.get("単位", ""), unit_price, row.get("キーワード", "")))
                    imported += 1
                except Exception:
                    skipped += 1
        conn.commit()
        conn.close()
        return imported, skipped

    def export_csv(self, csv_path):
        rows = self.get_all_unit_prices()
        with open(csv_path, "w", encoding="utf-8-sig", newline="") as f:
            writer = csv.writer(f)
            writer.writerow(["カテゴリ", "品名", "仕様", "単位", "単価", "キーワード"])
            for r in rows:
                writer.writerow([r[1], r[2], r[3], r[4], r[5], r[6]])
        return len(rows)

    def find_price_by_keyword(self, keyword):
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()
        c.execute("SELECT category, item_name, spec, unit, unit_price FROM unit_prices WHERE keywords LIKE ?",
                  (f"%{keyword}%",))
        row = c.fetchone()
        conn.close()
        return row

    def get_all_symbol_patterns(self):
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()
        c.execute("SELECT id, name, pattern_type, pattern_json, description, preset FROM symbol_patterns ORDER BY preset DESC, name")
        rows = c.fetchall()
        conn.close()
        return rows

    def upsert_symbol_pattern(self, data, record_id=None):
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()
        if record_id:
            c.execute("""UPDATE symbol_patterns SET name=?, pattern_type=?, pattern_json=?, description=? WHERE id=?""",
                      (data["name"], data["pattern_type"], data["pattern_json"], data["description"], record_id))
        else:
            c.execute("""INSERT INTO symbol_patterns (name, pattern_type, pattern_json, description, preset)
                         VALUES (?, ?, ?, ?, 0)""",
                      (data["name"], data["pattern_type"], data["pattern_json"], data["description"]))
        conn.commit()
        conn.close()

    def delete_symbol_pattern(self, record_id):
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()
        c.execute("DELETE FROM symbol_patterns WHERE id=? AND preset=0", (record_id,))
        conn.commit()
        conn.close()

    def get_all_cad_library(self):
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()
        c.execute("""SELECT id, manufacturer, model_number, category, item_name, file_path, url
                     FROM cad_library ORDER BY manufacturer, category, item_name""")
        rows = c.fetchall()
        conn.close()
        return rows

    def register_cad_file(self, data):
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()
        c.execute("""INSERT INTO cad_library
                     (manufacturer, model_number, category, item_name, file_path, pattern_signature, spec_json, url)
                     VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
                  (data["manufacturer"], data["model_number"], data["category"], data["item_name"],
                   data["file_path"], data["pattern_signature"], data["spec_json"], data.get("url", "")))
        conn.commit()
        conn.close()

    def delete_cad_file(self, record_id):
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()
        c.execute("DELETE FROM cad_library WHERE id=?", (record_id,))
        conn.commit()
        conn.close()

    def search_cad_by_pattern(self, pattern_signature):
        # ※本格的な類似検索は別途（シグネチャ比較等）必要。ここは一覧返すだけ。
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()
        c.execute("SELECT manufacturer, model_number, spec_json, item_name FROM cad_library LIMIT 10")
        rows = c.fetchall()
        conn.close()
        return rows


# -----------------------------
#  ダイアログ：単価
# -----------------------------
class UnitPriceDialog(tk.Toplevel):
    def __init__(self, parent, db, record=None):
        super().__init__(parent)
        self.db = db
        self.record = record
        self.result = None
        self.title("単価マスター編集" if record else "単価マスター追加")
        self.geometry("560x420")
        self.resizable(True, True)
        self._build()
        if record:
            self._fill(record)
        self.grab_set()

    def _build(self):
        f = ttk.Frame(self, padding=15)
        f.pack(fill=tk.BOTH, expand=True)

        labels = ["カテゴリ", "品名", "仕様", "単位", "単価(円)", "キーワード(カンマ区切り)"]
        self.vars = {}

        for i, label in enumerate(labels):
            ttk.Label(f, text=label).grid(row=i, column=0, sticky="e", pady=4)
            row = ttk.Frame(f)
            row.grid(row=i, column=1, sticky="ew", padx=8, pady=4)
            var = tk.StringVar()
            ent = ttk.Entry(row, textvariable=var)
            ent.pack(side=tk.LEFT, fill=tk.X, expand=True)
            add_paste_button(row, ent)
            self.vars[label] = var

        f.columnconfigure(1, weight=1)

        btn = ttk.Frame(f)
        btn.grid(row=len(labels), column=0, columnspan=2, pady=12)
        ttk.Button(btn, text="保存", command=self._save, width=15).pack(side=tk.LEFT, padx=10)
        ttk.Button(btn, text="キャンセル", command=self.destroy, width=15).pack(side=tk.LEFT, padx=10)

    def _fill(self, record):
        keys = ["カテゴリ", "品名", "仕様", "単位", "単価(円)", "キーワード(カンマ区切り)"]
        vals = [record[1], record[2], record[3], record[4], str(record[5]), record[6]]
        for k, v in zip(keys, vals):
            self.vars[k].set(v)

    def _save(self):
        try:
            data = {
                "category": self.vars["カテゴリ"].get().strip(),
                "item_name": self.vars["品名"].get().strip(),
                "spec": self.vars["仕様"].get().strip(),
                "unit": self.vars["単位"].get().strip(),
                "unit_price": float(self.vars["単価(円)"].get().strip().replace(",", "")),
                "keywords": self.vars["キーワード(カンマ区切り)"].get().strip(),
            }
            if not data["item_name"]:
                messagebox.showwarning("入力エラー", "品名を入力してください", parent=self)
                return
            rid = self.record[0] if self.record else None
            self.db.upsert_unit_price(data, rid)
            self.result = True
            self.destroy()
        except Exception:
            messagebox.showwarning("入力エラー", "単価は数値で入力してください", parent=self)


# -----------------------------
#  ダイアログ：記号パターン
# -----------------------------
class SymbolPatternDialog(tk.Toplevel):
    def __init__(self, parent, db, record=None):
        super().__init__(parent)
        self.db = db
        self.record = record
        self.result = None
        self.title("記号パターン編集" if record else "記号パターン追加")
        self.geometry("680x760")
        self.resizable(True, True)
        self._build()
        if record:
            self._fill(record)
        self.grab_set()

    def _build(self):
        canvas = tk.Canvas(self)
        sb = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        frm = ttk.Frame(canvas, padding=15)

        frm.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=frm, anchor="nw")
        canvas.configure(yscrollcommand=sb.set)

        # 設備名
        ttk.Label(frm, text="設備名", font=("", 10, "bold")).grid(row=0, column=0, sticky="e", pady=4)
        row0 = ttk.Frame(frm)
        row0.grid(row=0, column=1, sticky="ew", padx=8, pady=4)
        self.v_name = tk.StringVar()
        e0 = ttk.Entry(row0, textvariable=self.v_name)
        e0.pack(side=tk.LEFT, fill=tk.X, expand=True)
        add_paste_button(row0, e0)

        # タイプ
        ttk.Label(frm, text="パターンタイプ", font=("", 10, "bold")).grid(row=1, column=0, sticky="e", pady=4)
        self.v_type = tk.StringVar(value="simple_circle")
        types = ["simple_circle", "circle_with_line", "rectangle_with_cross", "circle_with_radial_lines"]
        ttk.Combobox(frm, textvariable=self.v_type, values=types, state="readonly", width=30)\
            .grid(row=1, column=1, sticky="w", padx=8, pady=4)

        params = ttk.LabelFrame(frm, text="パラメータ", padding=10)
        params.grid(row=2, column=0, columnspan=2, sticky="ew", pady=10)

        # 円
        ttk.Label(params, text="円 半径最小").grid(row=0, column=0, sticky="e", pady=2)
        self.v_rmin = tk.StringVar(value="40")
        ttk.Entry(params, textvariable=self.v_rmin, width=12).grid(row=0, column=1, sticky="w", pady=2)

        ttk.Label(params, text="円 半径最大").grid(row=1, column=0, sticky="e", pady=2)
        self.v_rmax = tk.StringVar(value="120")
        ttk.Entry(params, textvariable=self.v_rmax, width=12).grid(row=1, column=1, sticky="w", pady=2)

        # 線
        ttk.Label(params, text="線 方向").grid(row=2, column=0, sticky="e", pady=2)
        self.v_line_dir = tk.StringVar(value="any")
        ttk.Combobox(params, textvariable=self.v_line_dir, values=["any", "vertical", "horizontal"],
                     state="readonly", width=10).grid(row=2, column=1, sticky="w", pady=2)

        ttk.Label(params, text="線 本数").grid(row=3, column=0, sticky="e", pady=2)
        self.v_line_count = tk.StringVar(value="1")
        ttk.Entry(params, textvariable=self.v_line_count, width=12).grid(row=3, column=1, sticky="w", pady=2)

        # 矩形
        ttk.Label(params, text="矩形 幅最小").grid(row=4, column=0, sticky="e", pady=2)
        self.v_wmin = tk.StringVar(value="80")
        ttk.Entry(params, textvariable=self.v_wmin, width=12).grid(row=4, column=1, sticky="w", pady=2)

        ttk.Label(params, text="矩形 幅最大").grid(row=5, column=0, sticky="e", pady=2)
        self.v_wmax = tk.StringVar(value="220")
        ttk.Entry(params, textvariable=self.v_wmax, width=12).grid(row=5, column=1, sticky="w", pady=2)

        # 放射線
        ttk.Label(params, text="放射線 最小本数").grid(row=6, column=0, sticky="e", pady=2)
        self.v_radial = tk.StringVar(value="4")
        ttk.Entry(params, textvariable=self.v_radial, width=12).grid(row=6, column=1, sticky="w", pady=2)

        # 説明
        ttk.Label(frm, text="説明", font=("", 10, "bold")).grid(row=3, column=0, sticky="e", pady=4)
        row3 = ttk.Frame(frm)
        row3.grid(row=3, column=1, sticky="ew", padx=8, pady=4)
        self.v_desc = tk.StringVar()
        e3 = ttk.Entry(row3, textvariable=self.v_desc)
        e3.pack(side=tk.LEFT, fill=tk.X, expand=True)
        add_paste_button(row3, e3)

        btn = ttk.Frame(frm)
        btn.grid(row=4, column=0, columnspan=2, pady=15)
        ttk.Button(btn, text="保存", command=self._save, width=15).pack(side=tk.LEFT, padx=10)
        ttk.Button(btn, text="キャンセル", command=self.destroy, width=15).pack(side=tk.LEFT, padx=10)

        frm.columnconfigure(1, weight=1)
        canvas.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

    def _fill(self, record):
        # record: id, name, pattern_type, pattern_json, description, preset
        self.v_name.set(record[1])
        self.v_type.set(record[2])
        self.v_desc.set(record[4] or "")
        try:
            p = json.loads(record[3])
            self.v_rmin.set(str(p.get("radius_min", 40)))
            self.v_rmax.set(str(p.get("radius_max", 120)))
            self.v_line_dir.set(p.get("line_direction", "any"))
            self.v_line_count.set(str(p.get("line_count", 1)))
            self.v_wmin.set(str(p.get("width_min", 80)))
            self.v_wmax.set(str(p.get("width_max", 220)))
            self.v_radial.set(str(p.get("min_radial_lines", 4)))
        except Exception:
            pass

    def _save(self):
        try:
            ptype = self.v_type.get().strip()
            pattern = {"type": ptype}

            if ptype in ("simple_circle", "circle_with_line", "circle_with_radial_lines"):
                pattern["radius_min"] = float(self.v_rmin.get())
                pattern["radius_max"] = float(self.v_rmax.get())

            if ptype == "circle_with_line":
                pattern["line_direction"] = self.v_line_dir.get().strip()
                pattern["line_count"] = int(self.v_line_count.get())

            if ptype == "rectangle_with_cross":
                pattern["width_min"] = float(self.v_wmin.get())
                pattern["width_max"] = float(self.v_wmax.get())

            if ptype == "circle_with_radial_lines":
                pattern["min_radial_lines"] = int(self.v_radial.get())

            data = {
                "name": self.v_name.get().strip(),
                "pattern_type": ptype,
                "pattern_json": json.dumps(pattern, ensure_ascii=False),
                "description": self.v_desc.get().strip(),
            }

            if not data["name"]:
                messagebox.showwarning("入力エラー", "設備名を入力してください", parent=self)
                return

            rid = self.record[0] if self.record else None
            self.db.upsert_symbol_pattern(data, rid)
            self.result = True
            self.destroy()

        except Exception as e:
            messagebox.showwarning("入力エラー", f"入力値を確認してください:\n{e}", parent=self)


# -----------------------------
#  ダイアログ：CAD登録
# -----------------------------
class CADRegisterDialog(tk.Toplevel):
    def __init__(self, parent, db, file_path):
        super().__init__(parent)
        self.db = db
        self.file_path = file_path
        self.result = None
        self.signature = {}
        self.title("CADデータ登録")
        self.geometry("640x590")
        self.resizable(True, True)
        self._build()
        self._analyze_file()
        self.grab_set()

    def _build(self):
        f = ttk.Frame(self, padding=15)
        f.pack(fill=tk.BOTH, expand=True)

        ttk.Label(f, text="ファイル", font=("", 10, "bold")).grid(row=0, column=0, sticky="e", pady=4)
        ttk.Label(f, text=os.path.basename(self.file_path), foreground="blue", wraplength=420)\
            .grid(row=0, column=1, sticky="w", padx=8, pady=4)

        # メーカー
        ttk.Label(f, text="メーカー").grid(row=1, column=0, sticky="e", pady=4)
        row1 = ttk.Frame(f)
        row1.grid(row=1, column=1, sticky="ew", padx=8, pady=4)
        self.v_mfr = tk.StringVar()
        e1 = ttk.Entry(row1, textvariable=self.v_mfr)
        e1.pack(side=tk.LEFT, fill=tk.X, expand=True)
        add_paste_button(row1, e1)

        # 型番
        ttk.Label(f, text="型番").grid(row=2, column=0, sticky="e", pady=4)
        row2 = ttk.Frame(f)
        row2.grid(row=2, column=1, sticky="ew", padx=8, pady=4)
        self.v_model = tk.StringVar()
        e2 = ttk.Entry(row2, textvariable=self.v_model)
        e2.pack(side=tk.LEFT, fill=tk.X, expand=True)
        add_paste_button(row2, e2)

        # カテゴリ
        ttk.Label(f, text="カテゴリ").grid(row=3, column=0, sticky="e", pady=4)
        self.v_cat = tk.StringVar()
        ttk.Combobox(f, textvariable=self.v_cat,
                     values=["照明器具", "コンセント", "スイッチ", "分電盤", "配線", "空調", "その他"],
                     width=28, state="readonly").grid(row=3, column=1, sticky="w", padx=8, pady=4)

        # 機器名
        ttk.Label(f, text="機器名").grid(row=4, column=0, sticky="e", pady=4)
        row4 = ttk.Frame(f)
        row4.grid(row=4, column=1, sticky="ew", padx=8, pady=4)
        self.v_name = tk.StringVar()
        e4 = ttk.Entry(row4, textvariable=self.v_name)
        e4.pack(side=tk.LEFT, fill=tk.X, expand=True)
        add_paste_button(row4, e4)

        info_f = ttk.LabelFrame(f, text="解析情報", padding=10)
        info_f.grid(row=5, column=0, columnspan=2, sticky="nsew", pady=10)
        self.txt_info = scrolledtext.ScrolledText(info_f, height=12, wrap=tk.WORD)
        self.txt_info.pack(fill=tk.BOTH, expand=True)

        f.columnconfigure(1, weight=1)
        f.rowconfigure(5, weight=1)

        btn = ttk.Frame(f)
        btn.grid(row=6, column=0, columnspan=2, pady=12)
        ttk.Button(btn, text="登録", command=self._save, width=15).pack(side=tk.LEFT, padx=10)
        ttk.Button(btn, text="キャンセル", command=self.destroy, width=15).pack(side=tk.LEFT, padx=10)

    def _analyze_file(self):
        try:
            parser = AdvancedDXFParser()
            parser.parse_dxf(self.file_path)
            sig = parser.extract_pattern_signature()
            self.signature = sig

            info = []
            info.append("━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
            info.append(f"解析結果（方式: {sig.get('encoding','unknown')}）")
            info.append("━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
            info.append(f"円: {sig['circle_count']} / 線: {sig['line_count']} / HATCH: {sig['hatch_count']}")
            info.append(f"ブロック: {sig['block_count']} / テキスト: {sig['text_count']}")
            if sig.get("bounding_box"):
                info.append(f"BBox 幅: {sig['bounding_box']['width']:.1f} / 高さ: {sig['bounding_box']['height']:.1f}")
            if sig.get("circle_radii"):
                info.append("円半径（上位10）: " + ", ".join([f"{r:.1f}" for r in sig["circle_radii"]]))
            info.append("")
            info.append("抽出テキスト（先頭15）:")
            for t in sig.get("texts", [])[:15]:
                info.append("  " + str(t))

            self.txt_info.insert("1.0", "\n".join(info))

            # 推測（ざっくり）
            filename = os.path.basename(self.file_path).lower()
            combined = (" ".join(sig.get("texts", []))).lower()

            if any(k in filename or k in combined for k in ["panasonic", "pana", "パナ", "ﾊﾟﾅ"]):
                self.v_mfr.set("Panasonic")
            elif any(k in filename or k in combined for k in ["mitsubishi", "三菱", "ﾐﾂﾋﾞｼ"]):
                self.v_mfr.set("三菱電機")
            elif any(k in filename or k in combined for k in ["toshiba", "東芝", "ﾄｳｼﾊﾞ"]):
                self.v_mfr.set("東芝")
            elif any(k in filename or k in combined for k in ["hitachi", "日立", "ﾋﾀﾁ"]):
                self.v_mfr.set("日立")

            if any(k in filename or k in combined for k in ["down", "dl", "ダウン", "ﾀﾞｳﾝ", "照明", "light"]):
                self.v_cat.set("照明器具")
                self.v_name.set("ダウンライト" if ("down" in filename or "ダウン" in combined or "ﾀﾞｳﾝ" in combined) else "照明器具")
            elif any(k in filename or k in combined for k in ["outlet", "co", "コンセント", "ｺﾝｾﾝﾄ"]):
                self.v_cat.set("コンセント")
                self.v_name.set("コンセント")
            elif any(k in filename or k in combined for k in ["switch", "sw", "スイッチ", "ｽｲｯﾁ"]):
                self.v_cat.set("スイッチ")
                self.v_name.set("スイッチ")
            elif any(k in filename or k in combined for k in ["panel", "db", "分電盤", "配電盤", "ﾌﾞﾝﾃﾞﾝ"]):
                self.v_cat.set("分電盤")
                self.v_name.set("分電盤")
            elif any(k in filename or k in combined for k in ["fan", "vf", "換気扇", "ｶﾝｷ"]):
                self.v_cat.set("空調")
                self.v_name.set("換気扇")

            m = re.search(r"[A-Z]{2,4}[-_]?\d{3,6}[A-Z]?", (filename + " " + combined).upper())
            if m:
                self.v_model.set(m.group(0))

        except Exception as e:
            self.txt_info.insert("1.0", f"解析エラー:\n{e}\n\n対応外のDXF/破損の可能性があります。")
            self.signature = {}

    def _save(self):
        try:
            data = {
                "manufacturer": self.v_mfr.get().strip(),
                "model_number": self.v_model.get().strip(),
                "category": self.v_cat.get().strip(),
                "item_name": self.v_name.get().strip(),
                "file_path": self.file_path,
                "pattern_signature": json.dumps(self.signature, ensure_ascii=False),
                "spec_json": json.dumps({}, ensure_ascii=False),
                "url": "",
            }
            if not data["item_name"]:
                messagebox.showwarning("入力エラー", "機器名を入力してください", parent=self)
                return
            self.db.register_cad_file(data)
            self.result = True
            self.destroy()
        except Exception as e:
            messagebox.showerror("エラー", f"登録に失敗しました:\n{e}", parent=self)


# -----------------------------
#  メインアプリ
# -----------------------------
class AutoEstimationSystem:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_TITLE)

        # 画面
        sw = root.winfo_screenwidth()
        sh = root.winfo_screenheight()
        w = min(1250, max(450, int(sw * 0.98)))
        h = min(950,  max(650, int(sh * 0.95)))
        root.geometry(f"{w}x{h}")

        try:
            ttk.Style().theme_use("clam")
        except Exception:
            pass

        self.db = DatabaseManager(DB_PATH)
        self.cad_downloader = CADLibraryDownloader(CAD_LIBRARY_PATH)

        self.fitz = try_import_fitz()
        self.PdfReader = try_import_pypdf()
        self.requests = try_import_requests()
        self.ezdxf = try_import_ezdxf()

        self.log_q = queue.Queue()

        self.create_widgets()
        self.refresh_library_status()
        self.root.after(100, self._poll_log_queue)

    # ---- UI基盤 ----
    def create_widgets(self):
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)

        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="ファイル", menu=file_menu)
        file_menu.add_command(label="積算結果CSV出力", command=self.export_result_csv)
        file_menu.add_separator()
        file_menu.add_command(label="終了", command=self.root.quit)

        tool_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="ツール", menu=tool_menu)
        tool_menu.add_command(label="requestsインストール", command=lambda: self._install_worker_async("requests", "requests"))
        tool_menu.add_command(label="pypdfインストール（PDF推奨）", command=lambda: self._install_worker_async("pypdf", "pypdf"))
        tool_menu.add_command(label="PyMuPDFインストール（任意）", command=lambda: self._install_worker_async("PyMuPDF", "PyMuPDF"))
        tool_menu.add_command(label="ezdxfインストール（DXF必須/推奨）", command=lambda: self._install_worker_async("ezdxf", "ezdxf"))
        tool_menu.add_command(label="chardetインストール（任意）", command=lambda: self._install_worker_async("chardet", "chardet"))

        ttk.Label(self.root, text=APP_TITLE, font=("Arial", 13, "bold")).pack(pady=6)

        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=8, pady=4)

        tab_est = ttk.Frame(self.notebook, padding=6)
        self.notebook.add(tab_est, text="  積算実行  ")
        self._build_estimation_tab(tab_est)

        tab_price = ttk.Frame(self.notebook, padding=6)
        self.notebook.add(tab_price, text="  単価マスター  ")
        self._build_price_tab(tab_price)

        tab_sym = ttk.Frame(self.notebook, padding=6)
        self.notebook.add(tab_sym, text="  記号パターン  ")
        self._build_symbol_tab(tab_sym)

        tab_cad = ttk.Frame(self.notebook, padding=6)
        self.notebook.add(tab_cad, text="  CADライブラリ  ")
        self._build_cad_tab(tab_cad)

        self.status = ttk.Label(self.root, text="準備完了", relief=tk.SUNKEN, anchor="w")
        self.status.pack(side=tk.BOTTOM, fill=tk.X)

    def _build_estimation_tab(self, parent):
        top = ttk.LabelFrame(parent, text="ファイル選択", padding=10)
        top.pack(fill=tk.X, pady=4)

        row = ttk.Frame(top)
        row.pack(fill=tk.X, pady=4)
        ttk.Label(row, text="PDF/DXF:").pack(side=tk.LEFT, padx=5)

        self.file_path_var = tk.StringVar()
        ent = ttk.Entry(row, textvariable=self.file_path_var)
        ent.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        add_paste_button(row, ent)

        btns = ttk.Frame(top)
        btns.pack(pady=6)
        ttk.Button(btns, text="選択", command=self.select_file, width=12).pack(side=tk.LEFT, padx=5)
        ttk.Button(btns, text="積算実行", command=self.on_click_estimate, width=12).pack(side=tk.LEFT, padx=5)

        # PDF→DXF分離（オプション）
        ttk.Button(btns, text="PDFをDXF分離保存(ARCH/ELEC)", command=self.on_click_pdf_split_to_dxf, width=28)\
            .pack(side=tk.LEFT, padx=8)

        self.progress = ttk.Progressbar(parent, mode="indeterminate")
        self.progress.pack(fill=tk.X, pady=3)

        paned = ttk.PanedWindow(parent, orient=tk.VERTICAL)
        paned.pack(fill=tk.BOTH, expand=True, pady=4)

        result_f = ttk.LabelFrame(paned, text="積算結果", padding=6)
        paned.add(result_f, weight=3)

        columns = ("カテゴリ", "品名", "仕様", "単位", "数量", "単価(円)", "金額(円)")
        self.tree = ttk.Treeview(result_f, columns=columns, show="headings", height=10)
        for col in columns:
            self.tree.heading(col, text=col)
            if col == "数量":
                self.tree.column(col, width=70, anchor=tk.E)
            elif "円" in col:
                self.tree.column(col, width=120, anchor=tk.E)
            else:
                self.tree.column(col, width=140, anchor=tk.W)
        sb = ttk.Scrollbar(result_f, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)

        log_f = ttk.LabelFrame(paned, text="ログ", padding=6)
        paned.add(log_f, weight=2)
        self.txt_log = tk.Text(log_f, height=7, wrap="word")
        sb2 = ttk.Scrollbar(log_f, orient=tk.VERTICAL, command=self.txt_log.yview)
        self.txt_log.configure(yscrollcommand=sb2.set)
        self.txt_log.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb2.pack(side=tk.RIGHT, fill=tk.Y)

    def _build_price_tab(self, parent):
        btn_row = ttk.Frame(parent)
        btn_row.pack(fill=tk.X, pady=6)
        ttk.Button(btn_row, text="追加", command=self.price_add).pack(side=tk.LEFT, padx=4)
        ttk.Button(btn_row, text="編集", command=self.price_edit).pack(side=tk.LEFT, padx=4)
        ttk.Button(btn_row, text="削除", command=self.price_delete).pack(side=tk.LEFT, padx=4)
        ttk.Separator(btn_row, orient=tk.VERTICAL).pack(side=tk.LEFT, padx=8, fill=tk.Y)
        ttk.Button(btn_row, text="CSVインポート", command=self.price_import_csv).pack(side=tk.LEFT, padx=4)
        ttk.Button(btn_row, text="CSVエクスポート", command=self.price_export_csv).pack(side=tk.LEFT, padx=4)
        ttk.Button(btn_row, text="更新", command=self.price_refresh).pack(side=tk.RIGHT, padx=4)

        cols = ("ID", "カテゴリ", "品名", "仕様", "単位", "単価(円)", "キーワード")
        self.price_tree = ttk.Treeview(parent, columns=cols, show="headings", height=20)
        widths = [50, 110, 140, 160, 60, 110, 240]
        for col, w in zip(cols, widths):
            self.price_tree.heading(col, text=col)
            self.price_tree.column(col, width=w, anchor=tk.E if col in ("ID", "単価(円)") else tk.W)
        sb = ttk.Scrollbar(parent, orient=tk.VERTICAL, command=self.price_tree.yview)
        self.price_tree.configure(yscrollcommand=sb.set)
        self.price_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        self.price_tree.bind("<Double-1>", lambda e: self.price_edit())
        self.price_refresh()

    def _build_symbol_tab(self, parent):
        help_f = ttk.LabelFrame(parent, text="記号パターンとは", padding=8)
        help_f.pack(fill=tk.X, pady=4)
        ttk.Label(help_f, text=(
            "DXF図面内の電気設備記号を「円/線/矩形」などの形状で推定します。\n"
            "案件ごとに記号の大きさが違うので、必要に応じて追加・調整してください。"
        ), justify=tk.LEFT).pack(anchor="w")

        btn_row = ttk.Frame(parent)
        btn_row.pack(fill=tk.X, pady=6)
        ttk.Button(btn_row, text="追加", command=self.symbol_add).pack(side=tk.LEFT, padx=4)
        ttk.Button(btn_row, text="編集", command=self.symbol_edit).pack(side=tk.LEFT, padx=4)
        ttk.Button(btn_row, text="削除", command=self.symbol_delete).pack(side=tk.LEFT, padx=4)
        ttk.Button(btn_row, text="更新", command=self.symbol_refresh).pack(side=tk.RIGHT, padx=4)

        cols = ("ID", "設備名", "パターンタイプ", "説明", "プリセット")
        self.sym_tree = ttk.Treeview(parent, columns=cols, show="headings", height=20)
        widths = [50, 180, 180, 320, 80]
        for col, w in zip(cols, widths):
            self.sym_tree.heading(col, text=col)
            self.sym_tree.column(col, width=w, anchor=tk.E if col == "ID" else tk.W)
        sb = ttk.Scrollbar(parent, orient=tk.VERTICAL, command=self.sym_tree.yview)
        self.sym_tree.configure(yscrollcommand=sb.set)
        self.sym_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        self.sym_tree.bind("<Double-1>", lambda e: self.symbol_edit())
        self.symbol_refresh()

    def _build_cad_tab(self, parent):
        help_f = ttk.LabelFrame(parent, text="CADライブラリとは", padding=8)
        help_f.pack(fill=tk.X, pady=4)
        ttk.Label(help_f, text=(
            "メーカー提供のCADデータ（DXF）を登録します。\n"
            "登録したCADデータの解析情報（円/線/テキスト）を確認できます。"
        ), justify=tk.LEFT).pack(anchor="w")

        dl_f = ttk.LabelFrame(parent, text="CADデータダウンロード（ZIP/DXF）", padding=10)
        dl_f.pack(fill=tk.X, pady=6)

        ttk.Label(dl_f, text="ZIP/DXF URL:").pack(anchor="w", padx=4)
        url_row = ttk.Frame(dl_f)
        url_row.pack(fill=tk.X, pady=6)
        self.cad_url_var = tk.StringVar(value="https://...（メーカーZIP/DXFの直接URL）")
        ent = ttk.Entry(url_row, textvariable=self.cad_url_var)
        ent.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        add_paste_button(url_row, ent)
        ttk.Button(url_row, text="ダウンロード開始", command=self.cad_download_from_url, width=16)\
            .pack(side=tk.LEFT, padx=5)

        ttk.Label(dl_f, text="または").pack(pady=4)
        ttk.Button(dl_f, text="ローカルDXFをインポート", command=self.cad_import_local, width=26).pack(pady=4)

        btn_row = ttk.Frame(parent)
        btn_row.pack(fill=tk.X, pady=6)
        ttk.Button(btn_row, text="削除", command=self.cad_delete).pack(side=tk.LEFT, padx=4)
        ttk.Button(btn_row, text="フォルダを開く", command=self.cad_open_folder).pack(side=tk.LEFT, padx=4)
        ttk.Button(btn_row, text="更新", command=self.cad_refresh).pack(side=tk.RIGHT, padx=4)

        cols = ("ID", "メーカー", "型番", "カテゴリ", "機器名", "ファイル名")
        self.cad_tree = ttk.Treeview(parent, columns=cols, show="headings", height=15)
        widths = [50, 130, 140, 110, 160, 320]
        for col, w in zip(cols, widths):
            self.cad_tree.heading(col, text=col)
            self.cad_tree.column(col, width=w, anchor=tk.E if col == "ID" else tk.W)
        sb = ttk.Scrollbar(parent, orient=tk.VERTICAL, command=self.cad_tree.yview)
        self.cad_tree.configure(yscrollcommand=sb.set)
        self.cad_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        self.cad_refresh()

    # ---- ログ・状態 ----
    def log(self, msg):
        self.log_q.put(str(msg))

    def _poll_log_queue(self):
        try:
            while True:
                msg = self.log_q.get_nowait()
                self.txt_log.insert(tk.END, msg + "\n")
                self.txt_log.see(tk.END)
        except queue.Empty:
            pass
        self.root.after(100, self._poll_log_queue)

    def set_busy(self, busy, text):
        if busy:
            self.progress.start(10)
        else:
            self.progress.stop()
        self.status.config(text=text)

    def refresh_library_status(self):
        self.fitz = try_import_fitz()
        self.PdfReader = try_import_pypdf()
        self.requests = try_import_requests()
        self.ezdxf = try_import_ezdxf()
        libs = [
            "pypdf:" + ("OK" if self.PdfReader else "NG"),
            "PyMuPDF:" + ("OK" if self.fitz else "NG"),
            "ezdxf:" + ("OK" if self.ezdxf else "NG"),
            "requests:" + ("OK" if self.requests else "NG"),
        ]
        self.status.config(text=" / ".join(libs))

    # ---- インストール（非同期）----
    def _install_worker_async(self, label, package):
        threading.Thread(target=self._install_worker, args=(label, package), daemon=True).start()

    def _install_worker(self, label, package):
        self.root.after(0, lambda: self.set_busy(True, f"{label}インストール中..."))
        self.log(f"== pip install {package} ==")

        ok, lines = run_pip_install(package, timeout_sec=220)
        for ln in lines:
            self.log(ln)

        # PyMuPDFだけ --break-system-packages を追加で試す（Pydroid環境向け）
        if (not ok) and package.lower() == "pymupdf":
            self.log("---- 再試行: --break-system-packages ----")
            ok2, lines2 = run_pip_install(package, extra_args=["--break-system-packages"], timeout_sec=220)
            for ln in lines2:
                self.log(ln)
            ok = ok2

        def finish():
            self.set_busy(False, "準備完了" if ok else "インストール失敗")
            self.refresh_library_status()
            if ok:
                messagebox.showinfo("完了", f"{label}のインストールが完了しました。")
            else:
                messagebox.showerror("失敗", f"{label}のインストールに失敗しました。\nログを確認してください。")
        self.root.after(0, finish)

    # ---- 単価マスター ----
    def price_refresh(self):
        for row in self.price_tree.get_children():
            self.price_tree.delete(row)
        for r in self.db.get_all_unit_prices():
            self.price_tree.insert("", tk.END, values=(r[0], r[1], r[2], r[3], r[4], f"{r[5]:,.0f}", r[6]))

    def price_add(self):
        dlg = UnitPriceDialog(self.root, self.db)
        self.root.wait_window(dlg)
        if dlg.result:
            self.price_refresh()

    def price_edit(self):
        sel = self.price_tree.selection()
        if not sel:
            messagebox.showwarning("注意", "編集する行を選択してください。")
            return
        item = self.price_tree.item(sel[0])["values"]
        rid = item[0]
        record = next((r for r in self.db.get_all_unit_prices() if r[0] == rid), None)
        if not record:
            return
        dlg = UnitPriceDialog(self.root, self.db, record)
        self.root.wait_window(dlg)
        if dlg.result:
            self.price_refresh()

    def price_delete(self):
        sel = self.price_tree.selection()
        if not sel:
            messagebox.showwarning("注意", "削除する行を選択してください。")
            return
        item = self.price_tree.item(sel[0])["values"]
        if messagebox.askyesno("確認", f"「{item[2]}」を削除しますか?"):
            self.db.delete_unit_price(item[0])
            self.price_refresh()

    def price_import_csv(self):
        fp = filedialog.askopenfilename(title="CSVファイルを選択", filetypes=[("CSV files", "*.csv"), ("All files", "*.*")])
        if not fp:
            return
        try:
            imported, skipped = self.db.import_csv(fp)
            self.price_refresh()
            messagebox.showinfo("インポート完了", f"インポート: {imported}件\nスキップ: {skipped}件")
        except Exception as e:
            messagebox.showerror("エラー", f"CSVインポートに失敗:\n{e}")

    def price_export_csv(self):
        fp = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")], initialfile="単価マスター.csv")
        if not fp:
            return
        try:
            count = self.db.export_csv(fp)
            messagebox.showinfo("エクスポート完了", f"{count}件をエクスポート:\n{fp}")
        except Exception as e:
            messagebox.showerror("エラー", f"CSVエクスポートに失敗:\n{e}")

    # ---- 記号パターン ----
    def symbol_refresh(self):
        for row in self.sym_tree.get_children():
            self.sym_tree.delete(row)
        for r in self.db.get_all_symbol_patterns():
            preset = "✓" if r[5] else ""
            self.sym_tree.insert("", tk.END, values=(r[0], r[1], r[2], r[4], preset))

    def symbol_add(self):
        dlg = SymbolPatternDialog(self.root, self.db)
        self.root.wait_window(dlg)
        if dlg.result:
            self.symbol_refresh()

    def symbol_edit(self):
        sel = self.sym_tree.selection()
        if not sel:
            messagebox.showwarning("注意", "編集する行を選択してください。")
            return
        item = self.sym_tree.item(sel[0])["values"]
        rid = item[0]
        record = next((r for r in self.db.get_all_symbol_patterns() if r[0] == rid), None)
        if not record:
            return
        if record[5]:
            if not messagebox.askyesno("確認", "プリセット記号です。編集しますか？"):
                return
        dlg = SymbolPatternDialog(self.root, self.db, record)
        self.root.wait_window(dlg)
        if dlg.result:
            self.symbol_refresh()

    def symbol_delete(self):
        sel = self.sym_tree.selection()
        if not sel:
            messagebox.showwarning("注意", "削除する行を選択してください。")
            return
        item = self.sym_tree.item(sel[0])["values"]
        if item[4] == "✓":
            messagebox.showwarning("注意", "プリセット記号は削除できません。")
            return
        if messagebox.askyesno("確認", f"「{item[1]}」のパターンを削除しますか?"):
            self.db.delete_symbol_pattern(item[0])
            self.symbol_refresh()

    # ---- CADライブラリ ----
    def cad_refresh(self):
        for row in self.cad_tree.get_children():
            self.cad_tree.delete(row)
        for r in self.db.get_all_cad_library():
            filename = os.path.basename(r[5]) if r[5] else ""
            self.cad_tree.insert("", tk.END, values=(r[0], r[1], r[2], r[3], r[4], filename))

    def cad_download_from_url(self):
        url = self.cad_url_var.get().strip()
        if not url or url.startswith("https://..."):
            messagebox.showwarning("注意", "有効なURLを入力してください。")
            return
        if not self.requests:
            messagebox.showwarning("注意", "requestsが必要です（ツール→requestsインストール）")
            return

        def worker():
            try:
                self.root.after(0, lambda: self.set_busy(True, "ダウンロード中..."))
                files = self.cad_downloader.download_zip(url, callback=lambda m: self.log(m))
                self.log(f"ダウンロード完了: {len(files)}ファイル")
                self.root.after(0, lambda: self.set_busy(False, "ダウンロード完了"))
                self.root.after(0, lambda: messagebox.showinfo("完了", f"{len(files)}個保存しました。\n登録ダイアログを順に開きます。"))

                # 最大10件まで順に登録
                for fp in files[:10]:
                    self.root.after(0, lambda p=fp: self._open_cad_register_dialog(p))

            except Exception as e:
                self.log(f"[ERROR] {e}")
                self.root.after(0, lambda: self.set_busy(False, "ダウンロード失敗"))
                self.root.after(0, lambda: messagebox.showerror("エラー", f"ダウンロードに失敗:\n{e}"))

        threading.Thread(target=worker, daemon=True).start()

    def cad_import_local(self):
        files = filedialog.askopenfilenames(
            title="DXFファイルを選択",
            filetypes=[("DXF files", "*.dxf"), ("All files", "*.*")]
        )
        if not files:
            return
        for fp in files:
            try:
                dst = self.cad_downloader.import_local_file(fp)
                self.log(f"インポート: {os.path.basename(fp)}")
                self._open_cad_register_dialog(dst)
            except Exception as e:
                messagebox.showerror("エラー", f"インポート失敗:\n{fp}\n\n{e}")

    def _open_cad_register_dialog(self, file_path):
        dlg = CADRegisterDialog(self.root, self.db, file_path)
        self.root.wait_window(dlg)
        if dlg.result:
            self.cad_refresh()
            self.log(f"登録完了: {os.path.basename(file_path)}")

    def cad_delete(self):
        sel = self.cad_tree.selection()
        if not sel:
            messagebox.showwarning("注意", "削除する行を選択してください。")
            return
        item = self.cad_tree.item(sel[0])["values"]
        if messagebox.askyesno("確認", f"「{item[4]}」を削除しますか?"):
            self.db.delete_cad_file(item[0])
            self.cad_refresh()

    def cad_open_folder(self):
        try:
            os.makedirs(CAD_LIBRARY_PATH, exist_ok=True)
            p = os.path.abspath(CAD_LIBRARY_PATH)
            if platform.system() == "Windows":
                os.startfile(p)
            elif platform.system() == "Darwin":
                subprocess.Popen(["open", p])
            else:
                subprocess.Popen(["xdg-open", p])
        except Exception:
            messagebox.showinfo("フォルダ", os.path.abspath(CAD_LIBRARY_PATH))

    # ---- 積算実行 ----
    def select_file(self):
        fp = filedialog.askopenfilename(
            title="PDF/DXFファイルを選択",
            filetypes=[("対応ファイル", "*.pdf *.dxf"), ("PDF files", "*.pdf"), ("DXF files", "*.dxf"), ("All files", "*.*")]
        )
        if fp:
            self.file_path_var.set(fp)
            self.status.config(text=f"選択: {os.path.basename(fp)}")

    def on_click_estimate(self):
        file_path = self.file_path_var.get().strip()
        if not file_path:
            messagebox.showwarning("注意", "ファイルを選択してください。")
            return
        if not os.path.exists(file_path):
            messagebox.showerror("エラー", f"ファイルが見つかりません:\n{file_path}")
            return
        threading.Thread(target=self._estimate_worker, args=(file_path,), daemon=True).start()

    def _estimate_worker(self, file_path):
        self.root.after(0, lambda: self.set_busy(True, "解析中..."))
        self.log("== ファイル解析開始 ==")
        self.log(f"ファイル: {os.path.basename(file_path)}")

        try:
            equipment_counts = {}

            if file_path.lower().endswith(".dxf"):
                self.log("形式: DXF")
                if not try_import_ezdxf():
                    self.log("警告: ezdxfが未導入でもASCII DXFは簡易解析できますが、バイナリDXFは読めません。")
                    self.log("→ ツール → ezdxfインストール を推奨")
                parser = AdvancedDXFParser()
                parser.parse_dxf(file_path)
                self.log(f"方式: {parser.encoding_used} / 円:{len(parser.circles)} 線:{len(parser.lines)} ブロック:{len(parser.blocks)}")

                # DBからパターン
                raw_patterns = self.db.get_all_symbol_patterns()
                pattern_rules = []
                for r in raw_patterns:
                    try:
                        pattern_rules.append({"name": r[1], "pattern": json.loads(r[3])})
                    except Exception:
                        pass

                shape_counts = parser.count_equipment_by_patterns(pattern_rules)
                self.log("== 図形から検出 ==")
                for k, v in sorted(shape_counts.items()):
                    self.log(f"  {k}: {v}")

                # テキストからも検出
                text = parser.get_combined_text()
                if text.strip():
                    text_counts = self._extract_equipment(text)
                    self.log("== テキストから検出 ==")
                    for k, v in sorted(text_counts.items()):
                        self.log(f"  {k}: {v}")
                    # マージ（多い方を採用）
                    equipment_counts = shape_counts.copy()
                    for eq, cnt in text_counts.items():
                        if cnt > equipment_counts.get(eq, 0):
                            equipment_counts[eq] = cnt
                else:
                    equipment_counts = shape_counts

            elif file_path.lower().endswith(".pdf"):
                self.log("形式: PDF")
                if not self.PdfReader and not self.fitz:
                    raise RuntimeError("PDF読込ライブラリが必要です（ツール→pypdfインストール推奨）")
                text = self._extract_pdf_text(file_path)
                if not text.strip():
                    raise RuntimeError("PDFからテキストを抽出できませんでした（画像スキャンPDFの可能性）。")
                self.log(f"抽出文字数: {len(text)}")
                equipment_counts = self._extract_equipment(text)

            else:
                raise RuntimeError("非対応形式です（.pdf / .dxf のみ）")

            self.log("== 最終検出結果 ==")
            for k, v in sorted(equipment_counts.items()):
                self.log(f"  {k}: {v}")

            if not equipment_counts or sum(equipment_counts.values()) == 0:
                self.root.after(0, lambda: self.set_busy(False, "設備検出なし"))
                self.root.after(0, lambda: messagebox.showwarning(
                    "警告",
                    "設備情報を検出できませんでした。\n\n"
                    "PDF: 「ダウンライト x 8」等のテキストが必要\n"
                    "DXF: 記号パターンの調整/追加が必要な場合があります"
                ))
                return

            self.root.after(0, lambda: self._show_estimation(equipment_counts))

        except Exception as e:
            self.log(f"[ERROR] {e}")
            self.root.after(0, lambda: self.set_busy(False, "エラー発生"))
            self.root.after(0, lambda: messagebox.showerror("エラー", str(e)))

    def _extract_pdf_text(self, pdf_path):
        combined = ""
        if self.fitz is not None:
            try:
                self.log("エンジン: PyMuPDF")
                doc = self.fitz.open(pdf_path)
                texts = []
                for p in doc:
                    texts.append(p.get_text() or "")
                doc.close()
                combined = "\n".join(texts).strip()
            except Exception as e:
                self.log(f"[PyMuPDF失敗] {e}")
                combined = ""

        if not combined and self.PdfReader is not None:
            try:
                self.log("エンジン: pypdf")
                reader = self.PdfReader(pdf_path)
                texts = []
                for p in reader.pages:
                    texts.append(p.extract_text() or "")
                combined = "\n".join(texts).strip()
            except Exception as e:
                self.log(f"[pypdf失敗] {e}")
                combined = ""
        return combined

    def _extract_equipment(self, text):
        t = unicodedata.normalize("NFKC", text)
        t = t.replace("×", "x").replace("＊", "x").replace("*", "x")
        t = re.sub(r"[ \t]+", " ", t)

        counts = defaultdict(int)
        patterns = {
            "ダウンライト": [
                r"(ダウンライト|DL|ＤＬ)\s*[xX]\s*(\d+)",
                r"(ダウンライト|DL|ＤＬ).*?(\d+)\s*台",
            ],
            "ベースライト": [
                r"(ベースライト|BL|ＢＬ)\s*[xX]\s*(\d+)",
                r"(ベースライト|BL|ＢＬ).*?(\d+)\s*台",
            ],
            "コンセント": [
                r"(コンセント|CO|ＣＯ)\s*[xX]\s*(\d+)",
                r"(コンセント|CO|ＣＯ).*?(\d+)\s*個",
            ],
            "スイッチ": [
                r"(スイッチ|SW|ＳＷ)\s*[xX]\s*(\d+)",
                r"(スイッチ|SW|ＳＷ).*?(\d+)\s*個",
            ],
            "分電盤": [
                r"(分電盤|DB|ＤＢ)\s*[xX]\s*(\d+)",
                r"(分電盤|DB|ＤＢ).*?(\d+)\s*面",
                r"DB-(\d+)",
            ],
            "換気扇": [
                r"(換気扇|VF|ＶＦ)\s*[xX]\s*(\d+)",
                r"(換気扇|VF|ＶＦ).*?(\d+)\s*台",
            ],
        }

        for eq, plist in patterns.items():
            for pat in plist:
                for m in re.finditer(pat, t, flags=re.IGNORECASE | re.DOTALL):
                    try:
                        counts[eq] += int(m.group(m.lastindex))
                    except Exception:
                        pass
        return dict(counts)

    def _show_estimation(self, equipment_counts):
        for item in self.tree.get_children():
            self.tree.delete(item)

        total = 0
        found = 0
        for equipment, quantity in sorted(equipment_counts.items()):
            if quantity <= 0:
                continue
            row = self.db.find_price_by_keyword(equipment)
            if row:
                category, item_name, spec, unit, unit_price = row
                amount = quantity * unit_price
                total += amount
                found += 1
                self.tree.insert("", tk.END, values=(
                    category, item_name, spec, unit,
                    f"{quantity:.1f}", f"{unit_price:,.0f}", f"{amount:,.0f}"
                ))
            else:
                self.log(f"[単価未登録] {equipment} → 単価マスターで登録してください")

        if total > 0:
            self.tree.insert("", tk.END, values=("", "", "", "", "", "合計", f"{total:,.0f}"), tags=("total",))
            self.tree.tag_configure("total", background="#ffffcc")

        self.set_busy(False, f"積算完了: ¥{total:,.0f}")
        messagebox.showinfo("完了",
                            f"積算完了\n\n検出設備: {sum(equipment_counts.values())}個\n"
                            f"積算項目: {found}項目\n合計金額: ¥{total:,.0f}")

    # ---- 結果CSV ----
    def export_result_csv(self):
        if not self.tree.get_children():
            messagebox.showwarning("警告", "積算データがありません")
            return
        fp = filedialog.asksaveasfilename(defaultextension=".csv", initialfile="積算結果.csv",
                                          filetypes=[("CSV files", "*.csv")])
        if not fp:
            return
        with open(fp, "w", encoding="utf-8-sig", newline="") as f:
            writer = csv.writer(f)
            writer.writerow(["カテゴリ", "品名", "仕様", "単位", "数量", "単価(円)", "金額(円)"])
            for item in self.tree.get_children():
                writer.writerow(self.tree.item(item)["values"])
        messagebox.showinfo("成功", f"保存完了:\n{fp}")

    # -----------------------------
    #  追加要望2：PDF → ARCH/ELEC 擬似分離 DXF出力（ベクターPDFのみ）
    # -----------------------------
    def on_click_pdf_split_to_dxf(self):
        pdf_path = self.file_path_var.get().strip()
        if not pdf_path.lower().endswith(".pdf"):
            messagebox.showwarning("注意", "PDFを選択してください。")
            return
        if not self.fitz:
            messagebox.showwarning("注意", "PyMuPDFが必要です（ツール→PyMuPDFインストール）")
            return
        if not try_import_ezdxf():
            messagebox.showwarning("注意", "ezdxfが必要です（ツール→ezdxfインストール）")
            return

        out = filedialog.asksaveasfilename(defaultextension=".dxf", initialfile="split_ARCH_ELEC.dxf",
                                           filetypes=[("DXF files", "*.dxf")])
        if not out:
            return

        threading.Thread(target=self._pdf_split_to_dxf_worker, args=(pdf_path, out), daemon=True).start()

    def _pdf_split_to_dxf_worker(self, pdf_path, out_dxf):
        self.root.after(0, lambda: self.set_busy(True, "PDF→DXF分離中..."))
        self.log("== PDF→DXF分離開始 ==")

        fitz = self.fitz
        ezdxf = try_import_ezdxf()
        try:
            doc = fitz.open(pdf_path)

            # DXF作成
            dxf_doc = ezdxf.new("R2010")
            msp = dxf_doc.modelspace()
            # レイヤー
            if "ARCH" not in dxf_doc.layers:
                dxf_doc.layers.new(name="ARCH")
            if "ELEC" not in dxf_doc.layers:
                dxf_doc.layers.new(name="ELEC")

            # 簡易ヒューリスティック
            # - bboxが小さい要素 = ELEC（記号/注記っぽい）
            # - bboxが大きい要素 = ARCH（壁/外形っぽい）
            # ※単位はPDFポイントのまま
            SMALL_BBOX = 40.0  # この値は図面により調整が必要

            for pi, page in enumerate(doc, start=1):
                drawings = page.get_drawings()
                self.log(f"page {pi}: drawings={len(drawings)}")

                for d in drawings:
                    rect = d.get("rect", None)
                    if rect:
                        bw = float(rect.width)
                        bh = float(rect.height)
                    else:
                        bw = bh = 0.0

                    layer = "ELEC" if (bw <= SMALL_BBOX and bh <= SMALL_BBOX) else "ARCH"

                    # items: ('l', p1, p2), ('re', rect, ...), ('c', ... curve ...)
                    for it in d.get("items", []):
                        t = it[0]
                        if t == "l":
                            p1 = it[1]; p2 = it[2]
                            msp.add_line((float(p1.x), float(p1.y)), (float(p2.x), float(p2.y)), dxfattribs={"layer": layer})
                        elif t == "re":
                            r = it[1]
                            x1, y1, x2, y2 = float(r.x0), float(r.y0), float(r.x1), float(r.y1)
                            msp.add_lwpolyline([(x1, y1), (x2, y1), (x2, y2), (x1, y2), (x1, y1)], dxfattribs={"layer": layer})
                        else:
                            # 曲線等はここでは未対応（必要なら後で拡張）
                            pass

            doc.close()
            dxf_doc.saveas(out_dxf)

            self.log(f"保存: {out_dxf}")
            self.root.after(0, lambda: self.set_busy(False, "分離DXF出力完了"))
            self.root.after(0, lambda: messagebox.showinfo(
                "完了",
                "PDF→DXF分離保存が完了しました。\n\n"
                "ARCH/ELECの分類は簡易推定です。\n"
                "精度が悪い場合は SMALL_BBOX の調整が必要です。"
            ))
        except Exception as e:
            self.log(f"[ERROR] {e}")
            self.root.after(0, lambda: self.set_busy(False, "分離失敗"))
            self.root.after(0, lambda: messagebox.showerror("エラー", f"分離に失敗しました:\n{e}"))


if __name__ == "__main__":
    root = tk.Tk()
    AutoEstimationSystem(root)
    root.mainloop()