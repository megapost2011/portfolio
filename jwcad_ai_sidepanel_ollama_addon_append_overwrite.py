#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Jw_cad AI SidePanel Addon
Ollama + Gemma3:4b + Jw_cad R12 DXF generator

【目的】
- Jw_cadの横に常駐するAIチャットGUIを表示
- Ollama app.exe / ollama.exe を起動
- Ollamaサーバー接続確認
- モデル一覧取得、gemma3:4b選択
- 自然文 → 電気設備作図コマンドJSON → Jw_cad互換R12 DXF生成
- Jw_cadで自動オープン
- Jw_cad外部変形用BAT生成
- 常に前面表示ON/OFF

【必要ライブラリ】
    pip install requests

【推奨フォルダ例】
    E:\\JWCAD_AI_ADDON\\jwcad_ai_sidepanel_ollama_addon.py

【注意】
- Jw_cadの完全な内部ドッキングではなく、横に固定表示するサイドパネル方式です。
- DXFはJw_cad互換性重視で R12 ASCII DXF を手書き生成します。
- AIのJSONが崩れた場合でも、簡易ルールでフォールバックします。
"""

import os
import re
import sys
import json
import time
import math
import shutil
import socket
import datetime
import subprocess
import threading
import traceback
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, scrolledtext

try:
    import requests
except Exception:
    requests = None

APP_NAME = "Jw_cad AI SidePanel Addon"
APP_VERSION = "1.1.0-append"

DEFAULT_OLLAMA_URL = "http://127.0.0.1:11434"
DEFAULT_MODEL = "gemma3:4b"
DEFAULT_LAYER = "AI_DRAW"

BASE_DIR = os.path.abspath(os.path.dirname(__file__))
OUT_DIR = os.path.join(BASE_DIR, "out")
LOG_DIR = os.path.join(BASE_DIR, "logs")
CONFIG_PATH = os.path.join(BASE_DIR, "jwcad_ai_config.json")
BAT_PATH = os.path.join(BASE_DIR, "run_jwcad_ai_sidepanel.bat")
os.makedirs(OUT_DIR, exist_ok=True)
os.makedirs(LOG_DIR, exist_ok=True)


# ============================================================
# Utility
# ============================================================

def now_stamp():
    return datetime.datetime.now().strftime("%Y%m%d_%H%M%S")


def safe_float(v, default=0.0):
    try:
        return float(v)
    except Exception:
        return default


def safe_int(v, default=0):
    try:
        return int(float(v))
    except Exception:
        return default


def open_with_windows(path):
    try:
        os.startfile(path)  # type: ignore[attr-defined]
        return True
    except Exception:
        return False


def find_existing_path(candidates):
    for p in candidates:
        if p and os.path.exists(p):
            return p
    return ""


def load_config():
    default = {
        "ollama_url": DEFAULT_OLLAMA_URL,
        "model": DEFAULT_MODEL,
        "ollama_app_path": r"E:\Ollama\ollama app.exe",
        "ollama_exe_path": r"E:\Ollama\ollama.exe",
        "jwcad_exe_path": r"E:\JWW\Jw_win.exe",
        "always_on_top": True,
        "auto_open_jwcad": True,
        "panel_width": 460,
        "panel_height": 900,
        "panel_side": "right",
        "working_dxf_path": os.path.join(OUT_DIR, "jwcad_ai_working.dxf"),
        "append_to_working_file": True,
    }
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
            default.update(data)
        except Exception:
            pass
    # 自動補正候補
    if not os.path.exists(default.get("ollama_app_path", "")):
        p = find_existing_path([
            r"E:\Ollama\ollama app.exe",
            r"D:\Ollama\ollama app.exe",
            r"C:\Users\%USERNAME%\AppData\Local\Programs\Ollama\ollama app.exe".replace("%USERNAME%", os.environ.get("USERNAME", "")),
        ])
        if p:
            default["ollama_app_path"] = p
    if not os.path.exists(default.get("ollama_exe_path", "")):
        p = find_existing_path([
            r"E:\Ollama\ollama.exe",
            r"D:\Ollama\ollama.exe",
            r"C:\Users\%USERNAME%\AppData\Local\Programs\Ollama\ollama.exe".replace("%USERNAME%", os.environ.get("USERNAME", "")),
        ])
        if p:
            default["ollama_exe_path"] = p
    if not os.path.exists(default.get("jwcad_exe_path", "")):
        p = find_existing_path([
            r"E:\JWW\Jw_win.exe",
            r"C:\JWW\Jw_win.exe",
            r"C:\jww\Jw_win.exe",
            r"D:\JWW\Jw_win.exe",
        ])
        if p:
            default["jwcad_exe_path"] = p
    return default


def save_config(cfg):
    try:
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


# ============================================================
# R12 ASCII DXF Writer for Jw_cad compatibility
# ============================================================

class R12DXFWriter:
    def __init__(self):
        self.entities = []
        self.layers = set(["0", DEFAULT_LAYER, "AI_ARCH", "AI_ELEC", "AI_TEXT", "AI_DIM", "AI_WIRE"])
        self.points = []

    def add_layer(self, layer):
        if layer:
            self.layers.add(str(layer))

    def _pt(self, x, y):
        x = safe_float(x)
        y = safe_float(y)
        self.points.append((x, y))
        return x, y

    def add_line(self, x1, y1, x2, y2, layer=DEFAULT_LAYER):
        self.add_layer(layer)
        x1, y1 = self._pt(x1, y1)
        x2, y2 = self._pt(x2, y2)
        self.entities.append(("LINE", layer, x1, y1, x2, y2))

    def add_circle(self, x, y, r, layer=DEFAULT_LAYER):
        self.add_layer(layer)
        x, y = self._pt(x, y)
        r = abs(safe_float(r, 10.0))
        self.points += [(x-r, y-r), (x+r, y+r)]
        self.entities.append(("CIRCLE", layer, x, y, r))

    def add_text(self, x, y, text, height=80, layer="AI_TEXT", rotation=0):
        self.add_layer(layer)
        x, y = self._pt(x, y)
        text = str(text).replace("\n", " ")
        height = safe_float(height, 80)
        rotation = safe_float(rotation, 0)
        self.entities.append(("TEXT", layer, x, y, height, rotation, text))

    def add_rect(self, x, y, w, h, layer="AI_ARCH"):
        x = safe_float(x); y = safe_float(y); w = safe_float(w); h = safe_float(h)
        self.add_line(x, y, x+w, y, layer)
        self.add_line(x+w, y, x+w, y+h, layer)
        self.add_line(x+w, y+h, x, y+h, layer)
        self.add_line(x, y+h, x, y, layer)

    def add_polyline(self, pts, closed=False, layer=DEFAULT_LAYER):
        if not pts or len(pts) < 2:
            return
        for i in range(len(pts)-1):
            self.add_line(pts[i][0], pts[i][1], pts[i+1][0], pts[i+1][1], layer)
        if closed:
            self.add_line(pts[-1][0], pts[-1][1], pts[0][0], pts[0][1], layer)

    def add_dim_like(self, x1, y1, x2, y2, text=None, offset=150, layer="AI_DIM"):
        x1=safe_float(x1); y1=safe_float(y1); x2=safe_float(x2); y2=safe_float(y2); offset=safe_float(offset,150)
        dx = x2 - x1; dy = y2 - y1
        length = math.hypot(dx, dy)
        if length <= 0:
            return
        nx, ny = -dy / length, dx / length
        ax1, ay1 = x1 + nx*offset, y1 + ny*offset
        ax2, ay2 = x2 + nx*offset, y2 + ny*offset
        self.add_line(x1, y1, ax1, ay1, layer)
        self.add_line(x2, y2, ax2, ay2, layer)
        self.add_line(ax1, ay1, ax2, ay2, layer)
        # arrow ticks
        tick = 45
        self.add_line(ax1-tick, ay1-tick, ax1+tick, ay1+tick, layer)
        self.add_line(ax2-tick, ay2-tick, ax2+tick, ay2+tick, layer)
        if text is None:
            text = str(int(round(length)))
        self.add_text((ax1+ax2)/2, (ay1+ay2)/2 + 30, text, 90, layer)

    def add_downlight(self, x, y, r=80, label="DL", layer="AI_ELEC"):
        self.add_circle(x, y, r, layer)
        self.add_line(x-r*0.7, y, x+r*0.7, y, layer)
        self.add_line(x, y-r*0.7, x, y+r*0.7, layer)
        self.add_text(x+r+20, y-r/2, label, 70, layer)

    def add_light(self, x, y, r=90, label="L", layer="AI_ELEC"):
        self.add_downlight(x, y, r, label, layer)

    def add_outlet(self, x, y, r=55, label="CO", layer="AI_ELEC"):
        self.add_circle(x, y, r, layer)
        self.add_line(x, y-r*0.9, x, y+r*0.9, layer)
        self.add_text(x+r+20, y-r/2, label, 65, layer)

    def add_switch(self, x, y, r=50, label="SW", layer="AI_ELEC"):
        self.add_circle(x, y, r, layer)
        self.add_line(x-r*0.9, y, x+r*0.9, y, layer)
        self.add_text(x+r+20, y-r/2, label, 65, layer)

    def add_legend(self, x, y, items=None, layer="AI_TEXT"):
        if items is None:
            items = [("DL", "ダウンライト"), ("CO", "コンセント"), ("SW", "スイッチ"), ("W", "配線")]
        self.add_text(x, y, "凡例", 120, layer)
        yy = y - 180
        for code, name in items:
            self.add_text(x, yy, f"{code}: {name}", 90, layer)
            yy -= 130

    def bounds(self):
        if not self.points:
            return (-1000, -1000, 1000, 1000)
        xs = [p[0] for p in self.points]
        ys = [p[1] for p in self.points]
        pad = 500
        return (min(xs)-pad, min(ys)-pad, max(xs)+pad, max(ys)+pad)

    def _header(self):
        xmin, ymin, xmax, ymax = self.bounds()
        return [
            "0", "SECTION", "2", "HEADER",
            "9", "$ACADVER", "1", "AC1009",
            "9", "$EXTMIN", "10", f"{xmin:.6f}", "20", f"{ymin:.6f}", "30", "0.0",
            "9", "$EXTMAX", "10", f"{xmax:.6f}", "20", f"{ymax:.6f}", "30", "0.0",
            "0", "ENDSEC",
        ]

    def _tables(self):
        lines = ["0", "SECTION", "2", "TABLES", "0", "TABLE", "2", "LAYER", "70", str(len(self.layers))]
        for layer in sorted(self.layers):
            lines += ["0", "LAYER", "2", layer, "70", "0", "62", "7", "6", "CONTINUOUS"]
        lines += ["0", "ENDTAB", "0", "ENDSEC"]
        return lines

    def _entities(self):
        out = ["0", "SECTION", "2", "ENTITIES"]
        for ent in self.entities:
            typ = ent[0]
            if typ == "LINE":
                _, layer, x1, y1, x2, y2 = ent
                out += ["0", "LINE", "8", layer, "10", f"{x1:.6f}", "20", f"{y1:.6f}", "30", "0.0", "11", f"{x2:.6f}", "21", f"{y2:.6f}", "31", "0.0"]
            elif typ == "CIRCLE":
                _, layer, x, y, r = ent
                out += ["0", "CIRCLE", "8", layer, "10", f"{x:.6f}", "20", f"{y:.6f}", "30", "0.0", "40", f"{r:.6f}"]
            elif typ == "TEXT":
                _, layer, x, y, height, rot, text = ent
                out += ["0", "TEXT", "8", layer, "10", f"{x:.6f}", "20", f"{y:.6f}", "30", "0.0", "40", f"{height:.6f}", "1", text, "50", f"{rot:.6f}"]
        out += ["0", "ENDSEC", "0", "EOF"]
        return out

    def save(self, path):
        lines = []
        lines += self._header()
        lines += self._tables()
        lines += self._entities()
        text = "\r\n".join(lines) + "\r\n"
        # Jw_cad互換を優先してcp932。ただしASCII中心。
        with open(path, "w", encoding="cp932", errors="replace", newline="") as f:
            f.write(text)
        return path



# ============================================================
# DXF append/merge helpers
# ============================================================

def _read_dxf_lines(path):
    for enc in ("cp932", "utf-8", "shift_jis", "latin1"):
        try:
            with open(path, "r", encoding=enc, errors="ignore") as f:
                return f.read().splitlines()
        except Exception:
            continue
    return []


def _write_dxf_lines(path, lines):
    with open(path, "w", encoding="cp932", errors="replace", newline="") as f:
        f.write("\r\n".join(lines).rstrip("\r\n") + "\r\n")


def _new_entity_lines(writer):
    lines = writer._entities()
    # remove: 0 SECTION 2 ENTITIES ... 0 ENDSEC 0 EOF
    if len(lines) >= 8 and lines[:4] == ["0", "SECTION", "2", "ENTITIES"]:
        lines = lines[4:]
    if len(lines) >= 4 and lines[-4:] == ["0", "ENDSEC", "0", "EOF"]:
        lines = lines[:-4]
    return lines


def _parse_existing_bounds(lines):
    pts = []
    i = 0
    cur = {}
    etype = None
    while i < len(lines) - 1:
        code = lines[i].strip()
        val = lines[i + 1].strip()
        if code == "0":
            # flush previous circle with radius
            if etype == "CIRCLE" and "10" in cur and "20" in cur and "40" in cur:
                x = safe_float(cur.get("10")); y = safe_float(cur.get("20")); r = abs(safe_float(cur.get("40")))
                pts.extend([(x-r, y-r), (x+r, y+r)])
            etype = val
            cur = {}
        elif code in ("10", "11"):
            x = safe_float(val)
            # y should be next matching 20/21; handled when reading y also
            cur[code] = val
        elif code in ("20", "21"):
            cur[code] = val
            xcode = "10" if code == "20" else "11"
            if xcode in cur:
                pts.append((safe_float(cur[xcode]), safe_float(val)))
        elif code == "40":
            cur[code] = val
        i += 2
    if etype == "CIRCLE" and "10" in cur and "20" in cur and "40" in cur:
        x = safe_float(cur.get("10")); y = safe_float(cur.get("20")); r = abs(safe_float(cur.get("40")))
        pts.extend([(x-r, y-r), (x+r, y+r)])
    if not pts:
        return None
    xs = [p[0] for p in pts]; ys = [p[1] for p in pts]
    return (min(xs), min(ys), max(xs), max(ys))


def _replace_header_extents(lines, xmin, ymin, xmax, ymax):
    def set_pair_after(var_name, x, y):
        try:
            idx = lines.index(var_name)
        except ValueError:
            return
        # pattern: 9, $EXTMIN, 10, x, 20, y, 30, 0
        j = idx + 1
        while j < min(idx + 12, len(lines) - 1):
            if lines[j].strip() == "10" and j + 1 < len(lines):
                lines[j + 1] = f"{x:.6f}"
            if lines[j].strip() == "20" and j + 1 < len(lines):
                lines[j + 1] = f"{y:.6f}"
            j += 1
    set_pair_after("$EXTMIN", xmin, ymin)
    set_pair_after("$EXTMAX", xmax, ymax)


def append_writer_to_dxf(existing_path, writer, out_path=None):
    """Append new writer entities into an existing R12/ASCII DXF and overwrite it.
    If the existing file is missing or not usable, create a fresh DXF.
    """
    out_path = out_path or existing_path
    new_entities = _new_entity_lines(writer)
    if not existing_path or (not os.path.exists(existing_path)):
        writer.save(out_path)
        return out_path

    lines = _read_dxf_lines(existing_path)
    if not lines or "ENTITIES" not in [x.strip() for x in lines]:
        writer.save(out_path)
        return out_path

    # Find ENTITIES section's ENDSEC. Prefer the ENDSEC before EOF.
    insert_at = None
    for i in range(len(lines) - 1, -1, -1):
        if lines[i].strip() == "ENDSEC":
            # Usually the last ENDSEC is ENTITIES section in our R12 files.
            insert_at = i - 1 if i > 0 and lines[i-1].strip() == "0" else i
            break
    if insert_at is None:
        writer.save(out_path)
        return out_path

    merged = lines[:insert_at] + new_entities + lines[insert_at:]

    # Update EXTMIN/EXTMAX so Jw_cad zoom/open has reasonable bounds.
    b1 = _parse_existing_bounds(lines)
    b2 = writer.bounds()
    if b1:
        xmin = min(b1[0], b2[0]); ymin = min(b1[1], b2[1]); xmax = max(b1[2], b2[2]); ymax = max(b1[3], b2[3])
    else:
        xmin, ymin, xmax, ymax = b2
    _replace_header_extents(merged, xmin, ymin, xmax, ymax)
    _write_dxf_lines(out_path, merged)
    return out_path

# ============================================================
# Command execution
# ============================================================

class CommandExecutor:
    def __init__(self):
        self.writer = R12DXFWriter()
        self.context = {}

    def execute_all(self, commands):
        if isinstance(commands, dict):
            commands = [commands]
        if not isinstance(commands, list):
            commands = []
        for cmd in commands:
            if isinstance(cmd, dict):
                self.execute(cmd)
        return self.writer

    def execute(self, cmd):
        name = str(cmd.get("cmd") or cmd.get("command") or "").strip().lower()
        layer = cmd.get("layer") or DEFAULT_LAYER

        if name in ("draw_circle", "circle"):
            center = cmd.get("center", [cmd.get("x", 0), cmd.get("y", 0)])
            x, y = center[0], center[1]
            r = cmd.get("radius", cmd.get("r", 100))
            self.writer.add_circle(x, y, r, layer)

        elif name in ("draw_line", "line"):
            if "start" in cmd and "end" in cmd:
                x1, y1 = cmd["start"]
                x2, y2 = cmd["end"]
            else:
                x1, y1, x2, y2 = cmd.get("x1", 0), cmd.get("y1", 0), cmd.get("x2", 100), cmd.get("y2", 0)
            self.writer.add_line(x1, y1, x2, y2, layer)

        elif name in ("draw_text", "text"):
            pos = cmd.get("pos", [cmd.get("x", 0), cmd.get("y", 0)])
            self.writer.add_text(pos[0], pos[1], cmd.get("text", "TEXT"), cmd.get("height", 100), layer)

        elif name in ("room_rect", "rect", "rectangle", "draw_rect"):
            x = cmd.get("x", 0); y = cmd.get("y", 0)
            w = cmd.get("w", cmd.get("width", 3000))
            h = cmd.get("h", cmd.get("height", 2000))
            self.writer.add_rect(x, y, w, h, cmd.get("layer", "AI_ARCH"))
            self.context["last_room"] = {"x": safe_float(x), "y": safe_float(y), "w": safe_float(w), "h": safe_float(h)}
            if cmd.get("dimension", True):
                self.writer.add_dim_like(x, y, safe_float(x)+safe_float(w), y, f"{int(safe_float(w))}", -180, "AI_DIM")
                self.writer.add_dim_like(x, y, x, safe_float(y)+safe_float(h), f"{int(safe_float(h))}", 180, "AI_DIM")

        elif name in ("light", "downlight", "add_downlight"):
            pos = cmd.get("pos", [cmd.get("x", 0), cmd.get("y", 0)])
            self.writer.add_downlight(pos[0], pos[1], cmd.get("r", cmd.get("radius", 80)), cmd.get("label", "DL"), cmd.get("layer", "AI_ELEC"))

        elif name in ("outlet", "add_outlet"):
            pos = cmd.get("pos", [cmd.get("x", 0), cmd.get("y", 0)])
            self.writer.add_outlet(pos[0], pos[1], cmd.get("r", 55), cmd.get("label", "CO"), cmd.get("layer", "AI_ELEC"))

        elif name in ("switch", "add_switch"):
            pos = cmd.get("pos", [cmd.get("x", 0), cmd.get("y", 0)])
            self.writer.add_switch(pos[0], pos[1], cmd.get("r", 50), cmd.get("label", "SW"), cmd.get("layer", "AI_ELEC"))

        elif name in ("light_grid", "downlight_grid"):
            room = cmd.get("room") or self.context.get("last_room") or {"x":0,"y":0,"w":3000,"h":2000}
            x = safe_float(room.get("x", 0)); y = safe_float(room.get("y", 0))
            w = safe_float(room.get("w", room.get("width", 3000)))
            h = safe_float(room.get("h", room.get("height", 2000)))
            count_x = max(1, safe_int(cmd.get("count_x", 2), 2))
            count_y = max(1, safe_int(cmd.get("count_y", 2), 2))
            mx = safe_float(cmd.get("margin_x", w/(count_x+1)))
            my = safe_float(cmd.get("margin_y", h/(count_y+1)))
            if count_x == 1:
                xs = [x + w/2]
            else:
                xs = [x + mx + i*((w-2*mx)/(count_x-1)) for i in range(count_x)]
            if count_y == 1:
                ys = [y + h/2]
            else:
                ys = [y + my + j*((h-2*my)/(count_y-1)) for j in range(count_y)]
            idx = 1
            for yy in ys:
                for xx in xs:
                    self.writer.add_downlight(xx, yy, cmd.get("r", 80), f"{cmd.get('label','DL')}{idx}", cmd.get("layer", "AI_ELEC"))
                    idx += 1

        elif name in ("outlet_wall", "switch_wall"):
            room = cmd.get("room") or self.context.get("last_room") or {"x":0,"y":0,"w":3000,"h":2000}
            x = safe_float(room.get("x", 0)); y = safe_float(room.get("y", 0))
            w = safe_float(room.get("w", room.get("width", 3000)))
            h = safe_float(room.get("h", room.get("height", 2000)))
            wall = str(cmd.get("wall", "bottom")).lower()
            count = max(1, safe_int(cmd.get("count", 1), 1))
            offset = safe_float(cmd.get("offset", 250))
            margin = safe_float(cmd.get("margin", 500))
            positions = []
            if wall in ("bottom", "下", "下壁"):
                xs = [x+w/2] if count == 1 else [x+margin+i*((w-2*margin)/(count-1)) for i in range(count)]
                positions = [(xx, y+offset) for xx in xs]
            elif wall in ("top", "上", "上壁"):
                xs = [x+w/2] if count == 1 else [x+margin+i*((w-2*margin)/(count-1)) for i in range(count)]
                positions = [(xx, y+h-offset) for xx in xs]
            elif wall in ("left", "左", "左壁"):
                ys = [y+h/2] if count == 1 else [y+margin+i*((h-2*margin)/(count-1)) for i in range(count)]
                positions = [(x+offset, yy) for yy in ys]
            else:
                ys = [y+h/2] if count == 1 else [y+margin+i*((h-2*margin)/(count-1)) for i in range(count)]
                positions = [(x+w-offset, yy) for yy in ys]
            for i, (xx, yy) in enumerate(positions, start=1):
                if name == "outlet_wall":
                    self.writer.add_outlet(xx, yy, cmd.get("r", 55), f"{cmd.get('label','CO')}{i}", cmd.get("layer", "AI_ELEC"))
                else:
                    self.writer.add_switch(xx, yy, cmd.get("r", 50), f"{cmd.get('label','SW')}{i}", cmd.get("layer", "AI_ELEC"))

        elif name in ("wire", "polyline", "wiring"):
            pts = cmd.get("points") or []
            if len(pts) >= 2:
                self.writer.add_polyline(pts, False, cmd.get("layer", "AI_WIRE"))

        elif name in ("dimension", "dim"):
            if "start" in cmd and "end" in cmd:
                x1, y1 = cmd["start"]
                x2, y2 = cmd["end"]
                self.writer.add_dim_like(x1, y1, x2, y2, cmd.get("text"), cmd.get("offset", 150), cmd.get("layer", "AI_DIM"))

        elif name in ("legend", "add_legend"):
            pos = cmd.get("pos", [cmd.get("x", 0), cmd.get("y", -600)])
            self.writer.add_legend(pos[0], pos[1], cmd.get("items"), cmd.get("layer", "AI_TEXT"))

        # unknown commands are ignored deliberately


# ============================================================
# Ollama client and prompt/fallback parser
# ============================================================

SYSTEM_PROMPT = r"""
あなたはJw_cad電気設備作図コマンド変換AIです。
ユーザーの日本語指示を、必ずJSON配列だけで返してください。説明文は禁止です。
DXFやJWWの生データは書かず、以下のコマンドだけを返してください。

使用可能コマンド:
1) 部屋矩形
{"cmd":"room_rect","x":0,"y":0,"w":3000,"h":2000,"dimension":true}

2) 円
{"cmd":"circle","center":[0,0],"radius":100,"layer":"AI_DRAW"}

3) 線
{"cmd":"line","start":[0,0],"end":[1000,0],"layer":"AI_DRAW"}

4) 文字
{"cmd":"text","pos":[0,0],"text":"文字","height":100,"layer":"AI_TEXT"}

5) ダウンライト単体
{"cmd":"downlight","pos":[1500,1000],"label":"DL1","r":80}

6) ダウンライト格子配置
{"cmd":"light_grid","count_x":2,"count_y":2,"margin_x":700,"margin_y":500,"label":"DL"}
※直前にroom_rectがある場合、その部屋内に配置する。

7) コンセント壁配置
{"cmd":"outlet_wall","wall":"bottom","count":3,"offset":300,"margin":500,"label":"CO"}

8) スイッチ壁配置
{"cmd":"switch_wall","wall":"left","count":1,"offset":300,"margin":500,"label":"SW"}

9) 配線
{"cmd":"wire","points":[[300,300],[1500,1000],[2700,1000]]}

10) 寸法線風
{"cmd":"dimension","start":[0,0],"end":[3000,0],"text":"3000","offset":-200}

11) 凡例
{"cmd":"legend","pos":[0,-800]}

例:
ユーザー: 3000×2000の部屋を描いて中央に照明4台、下壁にコンセント3個
返答:
[
 {"cmd":"room_rect","x":0,"y":0,"w":3000,"h":2000,"dimension":true},
 {"cmd":"light_grid","count_x":2,"count_y":2,"margin_x":700,"margin_y":500,"label":"DL"},
 {"cmd":"outlet_wall","wall":"bottom","count":3,"offset":300,"margin":500,"label":"CO"},
 {"cmd":"legend","pos":[0,-800]}
]

注意:
- JSON以外を絶対に返さない。
- 数値はmm単位として扱う。
- 日本語キーは禁止。必ず上記の英語キーを使う。
"""


def extract_json(text):
    text = (text or "").strip()
    # remove markdown fences
    text = re.sub(r"^```(?:json)?\s*", "", text, flags=re.I)
    text = re.sub(r"\s*```$", "", text)
    try:
        return json.loads(text)
    except Exception:
        pass
    # find first JSON array or object
    m = re.search(r"(\[.*\]|\{.*\})", text, flags=re.S)
    if m:
        try:
            return json.loads(m.group(1))
        except Exception:
            pass
    raise ValueError("AI応答からJSONを抽出できませんでした。")


def fallback_parse_japanese(prompt):
    """AI失敗時の最低限ルールベース変換。"""
    p = prompt.replace("，", ",").replace("×", "x").replace("＊", "x")
    cmds = []

    # room 3000x2000
    m = re.search(r"(\d+(?:\.\d+)?)\s*[xX]\s*(\d+(?:\.\d+)?)", p)
    if m and ("部屋" in p or "室" in p or "矩形" in p or "四角" in p):
        w = float(m.group(1)); h = float(m.group(2))
        cmds.append({"cmd":"room_rect","x":0,"y":0,"w":w,"h":h,"dimension":True})

    # circle radius
    m = re.search(r"半径\s*(\d+(?:\.\d+)?)", p)
    if m and ("円" in p or "丸" in p):
        r = float(m.group(1))
        cmds.append({"cmd":"circle","center":[0,0],"radius":r,"layer":"AI_DRAW"})

    # line length to right
    m = re.search(r"右に\s*(\d+(?:\.\d+)?)", p)
    if m and "線" in p:
        l = float(m.group(1))
        cmds.append({"cmd":"line","start":[0,0],"end":[l,0],"layer":"AI_DRAW"})

    # lights count
    light_count = None
    m = re.search(r"(照明|ダウンライト|DL).*?(\d+)\s*(台|個)?", p, flags=re.I)
    if m:
        light_count = int(m.group(2))
    elif "照明4" in p or "4台" in p:
        light_count = 4
    if light_count:
        if light_count <= 1:
            cmds.append({"cmd":"downlight","pos":[1500,1000],"label":"DL1"})
        elif light_count == 2:
            cmds.append({"cmd":"light_grid","count_x":2,"count_y":1,"margin_x":700,"margin_y":500,"label":"DL"})
        elif light_count <= 4:
            cmds.append({"cmd":"light_grid","count_x":2,"count_y":2,"margin_x":700,"margin_y":500,"label":"DL"})
        else:
            cx = math.ceil(math.sqrt(light_count))
            cy = math.ceil(light_count / cx)
            cmds.append({"cmd":"light_grid","count_x":cx,"count_y":cy,"margin_x":700,"margin_y":500,"label":"DL"})

    # outlets
    m = re.search(r"(コンセント|CO).*?(\d+)\s*(個|台)?", p, flags=re.I)
    if m:
        count = int(m.group(2))
        wall = "bottom"
        if "上" in p: wall = "top"
        elif "左" in p: wall = "left"
        elif "右" in p: wall = "right"
        cmds.append({"cmd":"outlet_wall","wall":wall,"count":count,"offset":300,"margin":500,"label":"CO"})

    # switch
    if "スイッチ" in p or "SW" in p.upper():
        wall = "left"
        if "上" in p: wall = "top"
        elif "下" in p: wall = "bottom"
        elif "右" in p: wall = "right"
        cmds.append({"cmd":"switch_wall","wall":wall,"count":1,"offset":300,"margin":500,"label":"SW"})

    if "凡例" in p:
        cmds.append({"cmd":"legend","pos":[0,-800]})

    if not cmds:
        cmds = [{"cmd":"text","pos":[0,0],"text":"AI command parse failed","height":100}]
    return cmds


class OllamaClient:
    def __init__(self, base_url):
        self.base_url = base_url.rstrip("/")

    def is_alive(self, timeout=3):
        if requests is None:
            return False, "requests未導入"
        try:
            r = requests.get(self.base_url + "/api/tags", timeout=timeout)
            if r.status_code == 200:
                return True, "OK"
            return False, f"HTTP {r.status_code}"
        except Exception as e:
            return False, str(e)

    def list_models(self):
        if requests is None:
            raise RuntimeError("requestsが未導入です。pip install requests を実行してください。")
        r = requests.get(self.base_url + "/api/tags", timeout=10)
        r.raise_for_status()
        data = r.json()
        models = []
        for m in data.get("models", []):
            name = m.get("name")
            if name:
                models.append(name)
        return models

    def chat_to_commands(self, model, user_prompt):
        if requests is None:
            raise RuntimeError("requestsが未導入です。pip install requests を実行してください。")
        payload = {
            "model": model,
            "messages": [
                {"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user", "content": user_prompt},
            ],
            "stream": False,
            "options": {"temperature": 0.1}
        }
        r = requests.post(self.base_url + "/api/chat", json=payload, timeout=180)
        r.raise_for_status()
        data = r.json()
        text = data.get("message", {}).get("content", "")
        return extract_json(text), text


# ============================================================
# GUI
# ============================================================

class SidePanelApp:
    def __init__(self, root):
        self.root = root
        self.cfg = load_config()
        self.client = OllamaClient(self.cfg.get("ollama_url", DEFAULT_OLLAMA_URL))
        self.generated_dxf_path = ""
        self.last_commands = []
        self._build_ui()
        self.apply_geometry()
        self.root.after(300, self.check_server)

    def _build_ui(self):
        self.root.title(f"{APP_NAME} v{APP_VERSION}")
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

        main = ttk.Frame(self.root, padding=8)
        main.pack(fill=tk.BOTH, expand=True)

        title_row = ttk.Frame(main)
        title_row.pack(fill=tk.X)
        ttk.Label(title_row, text="Jw_cad AI Assistant", font=("Meiryo", 13, "bold")).pack(side=tk.LEFT)
        self.status_var = tk.StringVar(value="起動中")
        ttk.Label(title_row, textvariable=self.status_var, foreground="blue").pack(side=tk.RIGHT)

        # settings
        settings = ttk.LabelFrame(main, text="接続・設定", padding=6)
        settings.pack(fill=tk.X, pady=6)

        row0 = ttk.Frame(settings); row0.pack(fill=tk.X, pady=2)
        ttk.Label(row0, text="Ollama URL", width=12).pack(side=tk.LEFT)
        self.ollama_url_var = tk.StringVar(value=self.cfg.get("ollama_url", DEFAULT_OLLAMA_URL))
        ttk.Entry(row0, textvariable=self.ollama_url_var).pack(side=tk.LEFT, fill=tk.X, expand=True)

        row1 = ttk.Frame(settings); row1.pack(fill=tk.X, pady=2)
        ttk.Label(row1, text="モデル", width=12).pack(side=tk.LEFT)
        self.model_var = tk.StringVar(value=self.cfg.get("model", DEFAULT_MODEL))
        self.model_combo = ttk.Combobox(row1, textvariable=self.model_var, values=[self.model_var.get()], state="normal")
        self.model_combo.pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(row1, text="一覧取得", command=self.fetch_models, width=9).pack(side=tk.LEFT, padx=3)

        row2 = ttk.Frame(settings); row2.pack(fill=tk.X, pady=2)
        ttk.Button(row2, text="Ollama app.exe起動", command=self.start_ollama_app).pack(side=tk.LEFT, padx=2)
        ttk.Button(row2, text="ollama serve起動", command=self.start_ollama_serve).pack(side=tk.LEFT, padx=2)
        ttk.Button(row2, text="接続確認", command=self.check_server).pack(side=tk.LEFT, padx=2)

        row3 = ttk.Frame(settings); row3.pack(fill=tk.X, pady=2)
        ttk.Button(row3, text="Jw_cadパス", command=self.select_jwcad).pack(side=tk.LEFT, padx=2)
        ttk.Button(row3, text="Ollamaパス", command=self.select_ollama_paths).pack(side=tk.LEFT, padx=2)
        ttk.Button(row3, text="外部変形BAT生成", command=self.generate_bat).pack(side=tk.LEFT, padx=2)

        row4 = ttk.Frame(settings); row4.pack(fill=tk.X, pady=2)
        self.topmost_var = tk.BooleanVar(value=bool(self.cfg.get("always_on_top", True)))
        ttk.Checkbutton(row4, text="常に前面", variable=self.topmost_var, command=self.toggle_topmost).pack(side=tk.LEFT)
        self.auto_open_var = tk.BooleanVar(value=bool(self.cfg.get("auto_open_jwcad", True)))
        ttk.Checkbutton(row4, text="生成後Jw_cadで自動オープン", variable=self.auto_open_var, command=self.save_current_config).pack(side=tk.LEFT, padx=10)
        ttk.Button(row4, text="横に再配置", command=self.apply_geometry).pack(side=tk.RIGHT)

        row5 = ttk.Frame(settings); row5.pack(fill=tk.X, pady=2)
        ttk.Label(row5, text="作業DXF", width=12).pack(side=tk.LEFT)
        self.working_dxf_var = tk.StringVar(value=self.cfg.get("working_dxf_path", os.path.join(OUT_DIR, "jwcad_ai_working.dxf")))
        ttk.Entry(row5, textvariable=self.working_dxf_var).pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(row5, text="新規", command=self.new_working_dxf, width=6).pack(side=tk.LEFT, padx=2)
        ttk.Button(row5, text="既存選択", command=self.select_working_dxf, width=8).pack(side=tk.LEFT, padx=2)

        row6 = ttk.Frame(settings); row6.pack(fill=tk.X, pady=2)
        self.append_mode_var = tk.BooleanVar(value=bool(self.cfg.get("append_to_working_file", True)))
        ttk.Checkbutton(row6, text="同じDXFへ追記して上書き", variable=self.append_mode_var, command=self.save_current_config).pack(side=tk.LEFT)
        ttk.Button(row6, text="作業DXFを空にする", command=self.clear_working_dxf).pack(side=tk.LEFT, padx=10)

        # chat area
        chat_frame = ttk.LabelFrame(main, text="チャット", padding=6)
        chat_frame.pack(fill=tk.BOTH, expand=True, pady=6)
        self.chat_log = scrolledtext.ScrolledText(chat_frame, height=16, wrap=tk.WORD, font=("Meiryo", 10))
        self.chat_log.pack(fill=tk.BOTH, expand=True)
        self.chat_log.tag_configure("user", foreground="#0055aa", font=("Meiryo", 10, "bold"))
        self.chat_log.tag_configure("ai", foreground="#222222")
        self.chat_log.tag_configure("sys", foreground="#666666")
        self.chat_log.tag_configure("err", foreground="#bb0000")

        input_frame = ttk.Frame(main)
        input_frame.pack(fill=tk.X, pady=4)
        self.prompt_var = tk.StringVar()
        self.prompt_entry = ttk.Entry(input_frame, textvariable=self.prompt_var, font=("Meiryo", 11))
        self.prompt_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.prompt_entry.bind("<Return>", lambda e: self.send_prompt())
        ttk.Button(input_frame, text="送信・作図", command=self.send_prompt, width=10).pack(side=tk.LEFT, padx=4)

        # buttons
        act = ttk.Frame(main)
        act.pack(fill=tk.X, pady=3)
        ttk.Button(act, text="サンプル", command=self.insert_sample).pack(side=tk.LEFT, padx=2)
        ttk.Button(act, text="JSON表示", command=self.show_json).pack(side=tk.LEFT, padx=2)
        ttk.Button(act, text="DXFを開く", command=self.open_dxf).pack(side=tk.LEFT, padx=2)
        ttk.Button(act, text="Jw_cadで開く", command=self.open_in_jwcad).pack(side=tk.LEFT, padx=2)
        ttk.Button(act, text="出力フォルダ", command=lambda: open_with_windows(OUT_DIR)).pack(side=tk.RIGHT, padx=2)

        # command JSON preview
        json_frame = ttk.LabelFrame(main, text="作図コマンドJSON / ログ", padding=6)
        json_frame.pack(fill=tk.BOTH, expand=True, pady=6)
        self.json_text = scrolledtext.ScrolledText(json_frame, height=10, wrap=tk.NONE, font=("Consolas", 9))
        self.json_text.pack(fill=tk.BOTH, expand=True)

        self.append_chat("sys", "準備完了。例: 3000×2000の部屋を描いて中央に照明4台、下壁にコンセント3個")
        self.toggle_topmost()

    def append_chat(self, tag, msg):
        prefix = {"user": "あなた: ", "ai": "AI: ", "sys": "システム: ", "err": "エラー: "}.get(tag, "")
        self.chat_log.insert(tk.END, prefix, tag)
        self.chat_log.insert(tk.END, str(msg) + "\n\n", tag)
        self.chat_log.see(tk.END)

    def append_json(self, msg):
        self.json_text.insert(tk.END, str(msg) + "\n")
        self.json_text.see(tk.END)

    def apply_geometry(self):
        w = safe_int(self.cfg.get("panel_width", 460), 460)
        h = safe_int(self.cfg.get("panel_height", 900), 900)
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        side = self.cfg.get("panel_side", "right")
        if side == "left":
            x = 0
        else:
            x = max(0, sw - w - 10)
        y = 20
        h = min(h, sh - 80)
        self.root.geometry(f"{w}x{h}+{x}+{y}")

    def toggle_topmost(self):
        self.root.attributes("-topmost", bool(self.topmost_var.get()))
        self.save_current_config()

    def save_current_config(self):
        self.cfg["ollama_url"] = self.ollama_url_var.get().strip() or DEFAULT_OLLAMA_URL
        self.cfg["model"] = self.model_var.get().strip() or DEFAULT_MODEL
        self.cfg["always_on_top"] = bool(self.topmost_var.get())
        self.cfg["auto_open_jwcad"] = bool(self.auto_open_var.get())
        if hasattr(self, "working_dxf_var"):
            self.cfg["working_dxf_path"] = self.working_dxf_var.get().strip() or os.path.join(OUT_DIR, "jwcad_ai_working.dxf")
        if hasattr(self, "append_mode_var"):
            self.cfg["append_to_working_file"] = bool(self.append_mode_var.get())
        save_config(self.cfg)
        self.client = OllamaClient(self.cfg["ollama_url"])

    def on_close(self):
        self.save_current_config()
        self.root.destroy()

    def start_ollama_app(self):
        self.save_current_config()
        path = self.cfg.get("ollama_app_path", "")
        if not path or not os.path.exists(path):
            messagebox.showwarning("未設定", "ollama app.exe が見つかりません。Ollamaパスを設定してください。")
            return
        try:
            subprocess.Popen([path], cwd=os.path.dirname(path), shell=False)
            self.append_chat("sys", f"Ollama app.exeを起動しました: {path}")
            self.root.after(2500, self.check_server)
        except Exception as e:
            self.append_chat("err", e)

    def start_ollama_serve(self):
        self.save_current_config()
        path = self.cfg.get("ollama_exe_path", "")
        if not path or not os.path.exists(path):
            messagebox.showwarning("未設定", "ollama.exe が見つかりません。Ollamaパスを設定してください。")
            return
        try:
            creationflags = 0
            if hasattr(subprocess, "CREATE_NEW_CONSOLE"):
                creationflags = subprocess.CREATE_NEW_CONSOLE
            env = os.environ.copy()
            env.setdefault("OLLAMA_MODELS", r"E:\Ollama\models")
            env.setdefault("OLLAMA_LOAD_TIMEOUT", "15m")
            subprocess.Popen([path, "serve"], cwd=os.path.dirname(path), env=env, creationflags=creationflags)
            self.append_chat("sys", f"ollama serveを起動しました: {path}")
            self.root.after(2500, self.check_server)
        except Exception as e:
            self.append_chat("err", e)

    def check_server(self):
        self.save_current_config()
        ok, msg = self.client.is_alive()
        if ok:
            self.status_var.set("Ollama接続OK")
            self.append_json(f"[{datetime.datetime.now().strftime('%H:%M:%S')}] Ollama接続OK")
        else:
            self.status_var.set("Ollama未接続")
            self.append_json(f"[{datetime.datetime.now().strftime('%H:%M:%S')}] Ollama未接続: {msg}")

    def fetch_models(self):
        self.save_current_config()
        def worker():
            try:
                models = self.client.list_models()
                def done():
                    self.model_combo["values"] = models or [DEFAULT_MODEL]
                    if DEFAULT_MODEL in models:
                        self.model_var.set(DEFAULT_MODEL)
                    elif models and not self.model_var.get():
                        self.model_var.set(models[0])
                    self.append_chat("sys", "モデル一覧を取得しました: " + ", ".join(models))
                self.root.after(0, done)
            except Exception as e:
                self.root.after(0, lambda: self.append_chat("err", f"モデル一覧取得失敗: {e}"))
        threading.Thread(target=worker, daemon=True).start()

    def select_jwcad(self):
        fp = filedialog.askopenfilename(title="Jw_win.exeを選択", filetypes=[("Jw_cad", "Jw_win.exe"), ("exe", "*.exe"), ("all", "*.*")])
        if fp:
            self.cfg["jwcad_exe_path"] = fp
            save_config(self.cfg)
            self.append_chat("sys", f"Jw_cadパス設定: {fp}")

    def select_ollama_paths(self):
        fp = filedialog.askopenfilename(title="ollama app.exe または ollama.exeを選択", filetypes=[("exe", "*.exe"), ("all", "*.*")])
        if fp:
            name = os.path.basename(fp).lower()
            if "app" in name:
                self.cfg["ollama_app_path"] = fp
            else:
                self.cfg["ollama_exe_path"] = fp
            save_config(self.cfg)
            self.append_chat("sys", f"Ollamaパス設定: {fp}")

    def select_working_dxf(self):
        fp = filedialog.askopenfilename(
            title="追記先DXFを選択",
            filetypes=[("DXF files", "*.dxf"), ("All files", "*.*")]
        )
        if fp:
            self.working_dxf_var.set(fp)
            self.generated_dxf_path = fp
            self.save_current_config()
            self.append_chat("sys", f"追記先DXFを設定しました:\n{fp}")

    def new_working_dxf(self):
        fp = filedialog.asksaveasfilename(
            title="新規作業DXFを作成",
            defaultextension=".dxf",
            initialfile="jwcad_ai_working.dxf",
            initialdir=OUT_DIR,
            filetypes=[("DXF files", "*.dxf"), ("All files", "*.*")]
        )
        if fp:
            self.working_dxf_var.set(fp)
            self.generated_dxf_path = fp
            self.save_current_config()
            # 空DXFを作成
            R12DXFWriter().save(fp)
            self.append_chat("sys", f"新規作業DXFを作成しました:\n{fp}")
            if self.auto_open_var.get():
                self.open_in_jwcad()

    def clear_working_dxf(self):
        fp = self.working_dxf_var.get().strip()
        if not fp:
            return
        if messagebox.askyesno("確認", "作業DXFを空の図面で上書きしますか？\n既存図形は消えます。"):
            os.makedirs(os.path.dirname(fp) or OUT_DIR, exist_ok=True)
            R12DXFWriter().save(fp)
            self.generated_dxf_path = fp
            self.append_chat("sys", f"作業DXFを空にしました:\n{fp}")

    def save_writer_to_target(self, writer):
        self.save_current_config()
        work = self.working_dxf_var.get().strip() or os.path.join(OUT_DIR, "jwcad_ai_working.dxf")
        os.makedirs(os.path.dirname(work) or OUT_DIR, exist_ok=True)
        if self.append_mode_var.get():
            append_writer_to_dxf(work, writer, work)
            return work
        else:
            out_path = os.path.join(OUT_DIR, f"ai_jwcad_draw_{now_stamp()}.dxf")
            writer.save(out_path)
            return out_path

    def insert_sample(self):
        self.prompt_var.set("3000×2000の部屋を描いて、中央に照明4台、下壁にコンセント3個、左壁にスイッチ1個、凡例も描いて")
        self.prompt_entry.focus_set()

    def send_prompt(self):
        prompt = self.prompt_var.get().strip()
        if not prompt:
            return
        self.prompt_var.set("")
        self.append_chat("user", prompt)
        self.status_var.set("AI作図中...")
        self.save_current_config()

        def worker():
            raw = ""
            try:
                try:
                    commands, raw = self.client.chat_to_commands(self.model_var.get().strip() or DEFAULT_MODEL, prompt)
                except Exception as ai_err:
                    commands = fallback_parse_japanese(prompt)
                    raw = f"AI変換失敗のためフォールバック使用: {ai_err}"

                executor = CommandExecutor()
                writer = executor.execute_all(commands)
                out_path = self.save_writer_to_target(writer)
                self.generated_dxf_path = out_path
                self.last_commands = commands if isinstance(commands, list) else [commands]

                def done():
                    self.status_var.set("DXF生成完了")
                    self.append_chat("ai", f"DXFを生成しました。\n{out_path}")
                    self.json_text.delete("1.0", tk.END)
                    self.append_json("--- AI raw response / fallback info ---")
                    self.append_json(raw)
                    self.append_json("\n--- commands ---")
                    self.append_json(json.dumps(self.last_commands, ensure_ascii=False, indent=2))
                    if self.auto_open_var.get():
                        self.open_in_jwcad()
                self.root.after(0, done)
            except Exception as e:
                tb = traceback.format_exc()
                self.root.after(0, lambda: self.status_var.set("エラー"))
                self.root.after(0, lambda: self.append_chat("err", f"{e}\n{tb}"))
        threading.Thread(target=worker, daemon=True).start()

    def show_json(self):
        self.json_text.delete("1.0", tk.END)
        self.append_json(json.dumps(self.last_commands, ensure_ascii=False, indent=2))

    def open_dxf(self):
        if self.generated_dxf_path and os.path.exists(self.generated_dxf_path):
            open_with_windows(self.generated_dxf_path)
        else:
            messagebox.showinfo("未生成", "まだDXFが生成されていません。")

    def open_in_jwcad(self):
        path = self.generated_dxf_path
        if not path or not os.path.exists(path):
            self.append_chat("err", "DXFファイルがありません。")
            return
        jw = self.cfg.get("jwcad_exe_path", "")
        try:
            if jw and os.path.exists(jw):
                subprocess.Popen([jw, path], cwd=os.path.dirname(jw), shell=False)
                self.append_chat("sys", f"Jw_cadで開きました: {path}")
            else:
                ok = open_with_windows(path)
                if ok:
                    self.append_chat("sys", f"既定アプリでDXFを開きました: {path}")
                else:
                    self.append_chat("err", "Jw_cadパス未設定、かつ既定アプリで開けませんでした。")
        except Exception as e:
            self.append_chat("err", e)

    def generate_bat(self):
        py = sys.executable
        script = os.path.abspath(__file__)
        content = f'''@echo off
REM Jw_cad AI SidePanel Addon launcher
REM Jw_cad 外部変形メニュー等からこのBATを起動してください。
cd /d "{BASE_DIR}"
"{py}" "{script}"
'''
        try:
            with open(BAT_PATH, "w", encoding="cp932", errors="replace") as f:
                f.write(content)
            self.append_chat("sys", f"外部変形/起動用BATを生成しました:\n{BAT_PATH}")
            messagebox.showinfo("BAT生成完了", BAT_PATH)
        except Exception as e:
            self.append_chat("err", e)


def main():
    root = tk.Tk()
    try:
        style = ttk.Style()
        style.theme_use("clam")
    except Exception:
        pass
    app = SidePanelApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
