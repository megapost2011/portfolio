#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Jw_cad 電気設備作図コマンド強化版
Ollama + Gemma3:4b 連携 / Jw_cad互換 R12 ASCII DXF 出力 GUI

目的:
- Ollama app.exe / ollama.exe のローカルAPIを使う
- 日本語チャットを電気設備作図コマンドJSONに変換
- Python側の固定ロジックでJw_cad互換DXFを生成
- Jw_cadで自動オープン可能

対応例:
- 原点を中心とした半径100の円を描いて
- 3000×2000の部屋を描いて、中央にダウンライト4台を2×2で配置して
- 部屋の下壁にコンセントを3個、壁から300離して配置して
- 左壁にスイッチを1個配置して、照明へ配線して
- L1から右に1200、上に800の位置にDLを追加して

必要ライブラリ:
    pip install requests

注意:
- DXFは ezdxf を使わず、自前で R12 ASCII DXF を出力します。
- Jw_cad互換性優先のため、CIRCLE / LINE / TEXT / POLYLINE中心です。
"""

import os
import re
import sys
import json
import math
import time
import queue
import shutil
import threading
import subprocess
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from datetime import datetime

try:
    import requests
except Exception:
    requests = None

APP_TITLE = "Jw_cad 電気設備作図AI 強化版 - Ollama Gemma3 MCP風GUI"
DEFAULT_OLLAMA_EXE = r"E:\Ollama\ollama.exe"
DEFAULT_OLLAMA_APP_EXE = r"E:\Ollama\ollama app.exe"
DEFAULT_JWCAD_EXE = r"E:\JWW\Jw_win.exe"
DEFAULT_MODEL = "gemma3:4b"
DEFAULT_OUT_DIR = r"E:\JWCAD_AI_MCP\out"
OLLAMA_API_URL = "http://127.0.0.1:11434/api/chat"

# ============================================================
#  汎用
# ============================================================

def now_stamp():
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def ensure_dir(path):
    os.makedirs(path, exist_ok=True)


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


def extract_json_from_text(text):
    """LLM応答からJSON配列/オブジェクトだけを抽出する。"""
    if not text:
        raise ValueError("LLM応答が空です")
    s = text.strip()

    # ```json ... ``` 対策
    s = re.sub(r"^```(?:json)?", "", s, flags=re.I).strip()
    s = re.sub(r"```$", "", s).strip()

    # そのまま試す
    try:
        return json.loads(s)
    except Exception:
        pass

    # 配列を探す
    start = s.find("[")
    end = s.rfind("]")
    if start >= 0 and end > start:
        try:
            return json.loads(s[start:end+1])
        except Exception:
            pass

    # オブジェクトを探す
    start = s.find("{")
    end = s.rfind("}")
    if start >= 0 and end > start:
        try:
            return json.loads(s[start:end+1])
        except Exception:
            pass

    raise ValueError("JSONを抽出できませんでした:\n" + text[:1000])


# ============================================================
#  R12 ASCII DXF Writer - Jw_cad互換重視
# ============================================================

class R12DXFWriter:
    def __init__(self):
        self.entities = []
        self.layers = set(["0", "AI_DRAW", "ROOM", "LIGHT", "POWER", "SWITCH", "WIRE", "TEXT", "DIM"])
        self.points = []

    def add_layer(self, layer):
        if layer:
            self.layers.add(str(layer))

    def _pt(self, x, y):
        self.points.append((float(x), float(y)))

    def line(self, x1, y1, x2, y2, layer="AI_DRAW"):
        self.add_layer(layer)
        self._pt(x1, y1); self._pt(x2, y2)
        self.entities.append([
            "0", "LINE", "8", layer,
            "10", f"{float(x1):.6f}", "20", f"{float(y1):.6f}", "30", "0.0",
            "11", f"{float(x2):.6f}", "21", f"{float(y2):.6f}", "31", "0.0",
        ])

    def circle(self, x, y, r, layer="AI_DRAW"):
        self.add_layer(layer)
        r = abs(float(r))
        self._pt(x-r, y-r); self._pt(x+r, y+r)
        self.entities.append([
            "0", "CIRCLE", "8", layer,
            "10", f"{float(x):.6f}", "20", f"{float(y):.6f}", "30", "0.0",
            "40", f"{r:.6f}",
        ])

    def arc(self, x, y, r, start_angle, end_angle, layer="AI_DRAW"):
        self.add_layer(layer)
        r = abs(float(r))
        self._pt(x-r, y-r); self._pt(x+r, y+r)
        self.entities.append([
            "0", "ARC", "8", layer,
            "10", f"{float(x):.6f}", "20", f"{float(y):.6f}", "30", "0.0",
            "40", f"{r:.6f}",
            "50", f"{float(start_angle):.6f}",
            "51", f"{float(end_angle):.6f}",
        ])

    def text(self, x, y, text, height=120, layer="TEXT", rotation=0):
        self.add_layer(layer)
        self._pt(x, y)
        text = str(text).replace("\n", " ")
        self.entities.append([
            "0", "TEXT", "8", layer,
            "10", f"{float(x):.6f}", "20", f"{float(y):.6f}", "30", "0.0",
            "40", f"{float(height):.6f}",
            "1", text,
            "50", f"{float(rotation):.6f}",
        ])

    def polyline(self, pts, closed=False, layer="AI_DRAW"):
        if not pts:
            return
        self.add_layer(layer)
        flags = "1" if closed else "0"
        ent = ["0", "POLYLINE", "8", layer, "66", "1", "70", flags]
        for x, y in pts:
            self._pt(x, y)
            ent += ["0", "VERTEX", "8", layer, "10", f"{float(x):.6f}", "20", f"{float(y):.6f}", "30", "0.0"]
        ent += ["0", "SEQEND"]
        self.entities.append(ent)

    def rect(self, x, y, w, h, layer="AI_DRAW"):
        pts = [(x, y), (x+w, y), (x+w, y+h), (x, y+h)]
        self.polyline(pts, closed=True, layer=layer)

    def dimension_line(self, x1, y1, x2, y2, text=None, offset=150, layer="DIM"):
        # 簡易寸法線: 寸法補助線 + 寸法線 + 矢印風 + 文字
        dx, dy = x2-x1, y2-y1
        length = math.hypot(dx, dy)
        if length <= 0:
            return
        nx, ny = -dy/length, dx/length
        ox, oy = nx*offset, ny*offset
        ax1, ay1 = x1+ox, y1+oy
        ax2, ay2 = x2+ox, y2+oy
        self.line(x1, y1, ax1, ay1, layer)
        self.line(x2, y2, ax2, ay2, layer)
        self.line(ax1, ay1, ax2, ay2, layer)
        # 矢印風
        ah = min(80, max(30, length*0.03))
        ux, uy = dx/length, dy/length
        self.line(ax1, ay1, ax1 + ux*ah + nx*ah*0.4, ay1 + uy*ah + ny*ah*0.4, layer)
        self.line(ax1, ay1, ax1 + ux*ah - nx*ah*0.4, ay1 + uy*ah - ny*ah*0.4, layer)
        self.line(ax2, ay2, ax2 - ux*ah + nx*ah*0.4, ay2 - uy*ah + ny*ah*0.4, layer)
        self.line(ax2, ay2, ax2 - ux*ah - nx*ah*0.4, ay2 - uy*ah - ny*ah*0.4, layer)
        if text is None:
            text = f"{length:.0f}"
        self.text((ax1+ax2)/2, (ay1+ay2)/2 + 40, text, height=100, layer=layer)

    def save(self, path):
        if self.points:
            xs = [p[0] for p in self.points]; ys = [p[1] for p in self.points]
            minx, maxx = min(xs)-500, max(xs)+500
            miny, maxy = min(ys)-500, max(ys)+500
        else:
            minx = miny = -1000; maxx = maxy = 1000

        lines = []
        # HEADER
        lines += [
            "0", "SECTION", "2", "HEADER",
            "9", "$ACADVER", "1", "AC1009",
            "9", "$EXTMIN", "10", f"{minx:.6f}", "20", f"{miny:.6f}", "30", "0.0",
            "9", "$EXTMAX", "10", f"{maxx:.6f}", "20", f"{maxy:.6f}", "30", "0.0",
            "0", "ENDSEC",
        ]
        # TABLES/LAYERS
        lines += ["0", "SECTION", "2", "TABLES", "0", "TABLE", "2", "LAYER", "70", str(len(self.layers))]
        for layer in sorted(self.layers):
            lines += ["0", "LAYER", "2", layer, "70", "0", "62", "7", "6", "CONTINUOUS"]
        lines += ["0", "ENDTAB", "0", "ENDSEC"]
        # ENTITIES
        lines += ["0", "SECTION", "2", "ENTITIES"]
        for ent in self.entities:
            lines += ent
        lines += ["0", "ENDSEC", "0", "EOF"]
        # Jw_cad向けにcp932。英数字だけならASCII同等。
        with open(path, "w", encoding="cp932", errors="replace", newline="\r\n") as f:
            f.write("\n".join(lines))


# ============================================================
#  電気設備コマンド描画エンジン
# ============================================================

class ElectricalCommandRenderer:
    def __init__(self):
        self.dxf = R12DXFWriter()
        self.named_points = {}
        self.last_room = None  # {x,y,w,h}
        self.symbol_index = 1

    def render(self, commands):
        if isinstance(commands, dict):
            commands = [commands]
        if not isinstance(commands, list):
            raise ValueError("commands must be list or dict")
        for cmd in commands:
            if not isinstance(cmd, dict):
                continue
            self.render_one(cmd)
        return self.dxf

    def render_one(self, cmd):
        name = str(cmd.get("cmd") or cmd.get("command") or "").strip()
        if not name:
            return
        name = name.lower()

        if name in ("draw_circle", "circle"):
            x, y = self.get_xy(cmd, default=(0,0))
            r = safe_float(cmd.get("radius", cmd.get("r", 100)), 100)
            self.dxf.circle(x, y, r, cmd.get("layer", "AI_DRAW"))

        elif name in ("draw_line", "line"):
            x1, y1 = self.get_xy(cmd, keys=("start", "p1"), default=(0,0))
            x2, y2 = self.get_xy(cmd, keys=("end", "p2"), default=(100,0))
            self.dxf.line(x1, y1, x2, y2, cmd.get("layer", "AI_DRAW"))

        elif name in ("draw_rect", "rect", "rectangle", "room_rect", "room"):
            x = safe_float(cmd.get("x", 0)); y = safe_float(cmd.get("y", 0))
            w = safe_float(cmd.get("w", cmd.get("width", 3000)), 3000)
            h = safe_float(cmd.get("h", cmd.get("height", 2000)), 2000)
            self.dxf.rect(x, y, w, h, cmd.get("layer", "ROOM"))
            self.last_room = {"x": x, "y": y, "w": w, "h": h}
            if cmd.get("label", True):
                self.dxf.text(x+100, y+h+120, cmd.get("name", "ROOM"), height=120, layer="TEXT")
            if cmd.get("dimension", True):
                self.dxf.dimension_line(x, y, x+w, y, text=f"{w:.0f}", offset=-180, layer="DIM")
                self.dxf.dimension_line(x, y, x, y+h, text=f"{h:.0f}", offset=180, layer="DIM")

        elif name in ("text", "draw_text"):
            x, y = self.get_xy(cmd, default=(0,0))
            self.dxf.text(x, y, cmd.get("text", "TEXT"), height=safe_float(cmd.get("height", 120), 120), layer=cmd.get("layer", "TEXT"))

        elif name in ("dimension", "dim"):
            x1, y1 = self.get_xy(cmd, keys=("start", "p1"), default=(0,0))
            x2, y2 = self.get_xy(cmd, keys=("end", "p2"), default=(1000,0))
            self.dxf.dimension_line(x1, y1, x2, y2, text=cmd.get("text"), offset=safe_float(cmd.get("offset", 150), 150), layer=cmd.get("layer", "DIM"))

        elif name in ("light", "place_light", "downlight", "dl"):
            x, y = self.get_xy(cmd, default=(0,0))
            label = cmd.get("label") or cmd.get("symbol") or f"L{self.symbol_index}"
            self.draw_light(x, y, label=label, radius=safe_float(cmd.get("radius", 90), 90))
            self.named_points[str(label)] = (x, y)
            self.symbol_index += 1

        elif name in ("light_grid", "dl_grid"):
            self.draw_light_grid(cmd)

        elif name in ("outlet", "receptacle", "socket"):
            x, y = self.get_xy(cmd, default=(0,0))
            label = cmd.get("label", "CO")
            self.draw_outlet(x, y, label=label)
            self.named_points[str(label)] = (x, y)

        elif name in ("outlet_wall", "wall_outlets"):
            self.draw_outlet_wall(cmd)

        elif name in ("switch", "sw"):
            x, y = self.get_xy(cmd, default=(0,0))
            label = cmd.get("label", "SW")
            self.draw_switch(x, y, label=label)
            self.named_points[str(label)] = (x, y)

        elif name in ("switch_wall", "wall_switch"):
            self.draw_switch_wall(cmd)

        elif name in ("wire", "cable", "connect"):
            pts = self.resolve_points(cmd)
            self.draw_wire(pts, layer=cmd.get("layer", "WIRE"))

        elif name in ("place_from", "relative_place"):
            self.draw_relative(cmd)

        elif name in ("legend", "symbol_legend"):
            self.draw_legend(cmd)

        elif name in ("note", "notes"):
            x = safe_float(cmd.get("x", 0)); y = safe_float(cmd.get("y", -500))
            text = cmd.get("text", "")
            self.dxf.text(x, y, text, height=safe_float(cmd.get("height", 120), 120), layer="TEXT")

        else:
            # 未対応コマンドは図面に注記として残す
            self.dxf.text(0, -800 - 160*self.symbol_index, f"未対応cmd: {name}", height=100, layer="TEXT")
            self.symbol_index += 1

    def get_xy(self, cmd, keys=("point", "center", "pos"), default=(0,0)):
        for k in keys:
            if k in cmd and isinstance(cmd[k], (list, tuple)) and len(cmd[k]) >= 2:
                return safe_float(cmd[k][0]), safe_float(cmd[k][1])
        if "x" in cmd or "y" in cmd:
            return safe_float(cmd.get("x", default[0])), safe_float(cmd.get("y", default[1]))
        return default

    def room_point(self, wall, index=1, count=1, offset=300):
        room = self.last_room or {"x":0,"y":0,"w":3000,"h":2000}
        x, y, w, h = room["x"], room["y"], room["w"], room["h"]
        wall = str(wall).lower()
        t = index / (count + 1)
        if wall in ("bottom", "下", "south"):
            return x + w*t, y + offset
        if wall in ("top", "上", "north"):
            return x + w*t, y + h - offset
        if wall in ("left", "左", "west"):
            return x + offset, y + h*t
        if wall in ("right", "右", "east"):
            return x + w - offset, y + h*t
        return x + w*t, y + offset

    def draw_light(self, x, y, label="L", radius=90):
        self.dxf.circle(x, y, radius, "LIGHT")
        # ダウンライト風: 十字または中心点
        self.dxf.line(x-radius*0.6, y, x+radius*0.6, y, "LIGHT")
        self.dxf.line(x, y-radius*0.6, x, y+radius*0.6, "LIGHT")
        self.dxf.text(x+radius+35, y-radius/2, label, height=90, layer="TEXT")

    def draw_light_grid(self, cmd):
        room = self.last_room or {"x":0,"y":0,"w":3000,"h":2000}
        count_x = safe_int(cmd.get("count_x", cmd.get("cols", 2)), 2)
        count_y = safe_int(cmd.get("count_y", cmd.get("rows", 2)), 2)
        margin_x = safe_float(cmd.get("margin_x", cmd.get("margin", 500)), 500)
        margin_y = safe_float(cmd.get("margin_y", cmd.get("margin", 500)), 500)
        symbol = cmd.get("symbol", "L")
        radius = safe_float(cmd.get("radius", 90), 90)
        x0 = safe_float(cmd.get("x", room["x"]), room["x"])
        y0 = safe_float(cmd.get("y", room["y"]), room["y"])
        w = safe_float(cmd.get("w", room["w"]), room["w"])
        h = safe_float(cmd.get("h", room["h"]), room["h"])
        usable_w = max(0, w - 2*margin_x)
        usable_h = max(0, h - 2*margin_y)
        for iy in range(count_y):
            yy = y0 + margin_y + (usable_h * iy / max(1, count_y-1) if count_y > 1 else usable_h/2)
            for ix in range(count_x):
                xx = x0 + margin_x + (usable_w * ix / max(1, count_x-1) if count_x > 1 else usable_w/2)
                label = f"{symbol}{self.symbol_index}"
                self.draw_light(xx, yy, label, radius)
                self.named_points[label] = (xx, yy)
                self.symbol_index += 1

    def draw_outlet(self, x, y, label="CO"):
        r = 70
        self.dxf.circle(x, y, r, "POWER")
        self.dxf.line(x, y-r, x, y+r, "POWER")
        self.dxf.text(x+r+30, y-r/2, label, height=80, layer="TEXT")

    def draw_outlet_wall(self, cmd):
        wall = cmd.get("wall", "bottom")
        count = safe_int(cmd.get("count", 1), 1)
        offset = safe_float(cmd.get("offset", 300), 300)
        label = cmd.get("label", "CO")
        for i in range(1, count+1):
            x, y = self.room_point(wall, i, count, offset)
            name = f"{label}{i}" if count > 1 else label
            self.draw_outlet(x, y, name)
            self.named_points[name] = (x, y)

    def draw_switch(self, x, y, label="SW"):
        r = 65
        self.dxf.circle(x, y, r, "SWITCH")
        self.dxf.line(x-r, y, x+r, y, "SWITCH")
        self.dxf.text(x+r+30, y-r/2, label, height=80, layer="TEXT")

    def draw_switch_wall(self, cmd):
        wall = cmd.get("wall", "left")
        count = safe_int(cmd.get("count", 1), 1)
        offset = safe_float(cmd.get("offset", 300), 300)
        label = cmd.get("label", "SW")
        for i in range(1, count+1):
            x, y = self.room_point(wall, i, count, offset)
            name = f"{label}{i}" if count > 1 else label
            self.draw_switch(x, y, name)
            self.named_points[name] = (x, y)

    def resolve_points(self, cmd):
        pts = []
        raw = cmd.get("points")
        if isinstance(raw, list):
            for p in raw:
                if isinstance(p, str) and p in self.named_points:
                    pts.append(self.named_points[p])
                elif isinstance(p, (list, tuple)) and len(p) >= 2:
                    pts.append((safe_float(p[0]), safe_float(p[1])))
        if not pts:
            for key in ("from", "to"):
                p = cmd.get(key)
                if isinstance(p, str) and p in self.named_points:
                    pts.append(self.named_points[p])
                elif isinstance(p, (list, tuple)) and len(p) >= 2:
                    pts.append((safe_float(p[0]), safe_float(p[1])))
        return pts

    def draw_wire(self, pts, layer="WIRE"):
        if len(pts) < 2:
            return
        # 折れ線
        for (x1,y1), (x2,y2) in zip(pts[:-1], pts[1:]):
            self.dxf.line(x1, y1, x2, y2, layer)

    def draw_relative(self, cmd):
        base = cmd.get("base", cmd.get("from", ""))
        if base not in self.named_points:
            self.dxf.text(0, -1000, f"基準点がありません: {base}", height=100, layer="TEXT")
            return
        bx, by = self.named_points[base]
        dx = safe_float(cmd.get("dx", 0)); dy = safe_float(cmd.get("dy", 0))
        x, y = bx+dx, by+dy
        symbol = str(cmd.get("symbol", "DL")).upper()
        label = cmd.get("label", f"{symbol}{self.symbol_index}")
        if symbol in ("DL", "L", "LIGHT"):
            self.draw_light(x, y, label)
        elif symbol in ("CO", "OUTLET", "SOCKET"):
            self.draw_outlet(x, y, label)
        elif symbol in ("SW", "SWITCH"):
            self.draw_switch(x, y, label)
        else:
            self.dxf.circle(x, y, 70, "AI_DRAW")
            self.dxf.text(x+100, y, label, height=80, layer="TEXT")
        self.named_points[str(label)] = (x, y)
        self.symbol_index += 1

    def draw_legend(self, cmd):
        x = safe_float(cmd.get("x", 0)); y = safe_float(cmd.get("y", -900))
        self.dxf.text(x, y, "凡例", height=130, layer="TEXT")
        self.draw_light(x+120, y-220, "DL")
        self.dxf.text(x+300, y-250, "DL: ダウンライト", height=90, layer="TEXT")
        self.draw_outlet(x+120, y-420, "CO")
        self.dxf.text(x+300, y-450, "CO: コンセント", height=90, layer="TEXT")
        self.draw_switch(x+120, y-620, "SW")
        self.dxf.text(x+300, y-650, "SW: スイッチ", height=90, layer="TEXT")


# ============================================================
#  Ollama連携
# ============================================================

SYSTEM_PROMPT = r"""
あなたはJw_cad用の電気設備作図コマンド変換AIです。
ユーザーの日本語指示を、必ずJSON配列だけで返してください。説明文は禁止です。

目的:
- Jw_cadに読み込ませるDXFをPythonが生成するためのコマンドJSONを作る。
- DXFそのものは書かない。
- 単位はmm。
- 座標は [x, y]。

対応cmd:
1. room_rect
{"cmd":"room_rect","x":0,"y":0,"w":3000,"h":2000,"name":"部屋","dimension":true}

2. light_grid
{"cmd":"light_grid","count_x":2,"count_y":2,"margin_x":500,"margin_y":500,"symbol":"L"}
直前のroom_rect内に照明を格子配置する。

3. light
{"cmd":"light","x":1500,"y":1000,"label":"L1"}

4. outlet_wall
{"cmd":"outlet_wall","wall":"bottom","count":3,"offset":300,"label":"CO"}
wallは bottom/top/left/right。

5. outlet
{"cmd":"outlet","x":500,"y":300,"label":"CO1"}

6. switch_wall
{"cmd":"switch_wall","wall":"left","count":1,"offset":300,"label":"SW"}

7. switch
{"cmd":"switch","x":300,"y":1000,"label":"SW1"}

8. wire
{"cmd":"wire","points":["SW","L1","L2"]}
または {"cmd":"wire","points":[[0,0],[1000,0],[1000,500]]}

9. place_from
{"cmd":"place_from","base":"L1","dx":1200,"dy":800,"symbol":"DL","label":"L2"}

10. draw_circle
{"cmd":"draw_circle","center":[0,0],"radius":100}

11. draw_line
{"cmd":"draw_line","start":[0,0],"end":[1000,0]}

12. text
{"cmd":"text","x":0,"y":-500,"text":"注記","height":120}

13. dimension
{"cmd":"dimension","start":[0,0],"end":[3000,0],"text":"3000","offset":-200}

14. legend
{"cmd":"legend","x":0,"y":-800}

重要ルール:
- JSON配列だけ返す。
- コメントや説明文は禁止。
- 不明な場合も、もっとも合理的な作図コマンドに変換する。
- 「部屋」「矩形」「四角」は room_rect にする。
- 「ダウンライト」「照明」「DL」は light または light_grid にする。
- 「コンセント」「CO」は outlet または outlet_wall にする。
- 「スイッチ」「SW」は switch または switch_wall にする。
- 「原点中心 半径100 円」は draw_circle。
- 「中央に4台」は light_grid count_x=2 count_y=2 と解釈する。
- 「四隅から500内側に照明」は light_grid count_x=2 count_y=2 margin_x=500 margin_y=500。

例1:
ユーザー: 原点を中心とした半径100の円を描いて
出力:
[{"cmd":"draw_circle","center":[0,0],"radius":100}]

例2:
ユーザー: 3000×2000の部屋を描いて、中央にダウンライト4台を2×2で配置して、凡例も付けて
出力:
[
 {"cmd":"room_rect","x":0,"y":0,"w":3000,"h":2000,"name":"部屋","dimension":true},
 {"cmd":"light_grid","count_x":2,"count_y":2,"margin_x":750,"margin_y":500,"symbol":"L"},
 {"cmd":"legend","x":0,"y":-800}
]

例3:
ユーザー: 4000×3000の部屋、下壁にコンセント3個、左壁にスイッチ1個、照明4台、スイッチから照明へ配線
出力:
[
 {"cmd":"room_rect","x":0,"y":0,"w":4000,"h":3000,"name":"部屋","dimension":true},
 {"cmd":"light_grid","count_x":2,"count_y":2,"margin_x":800,"margin_y":700,"symbol":"L"},
 {"cmd":"outlet_wall","wall":"bottom","count":3,"offset":300,"label":"CO"},
 {"cmd":"switch_wall","wall":"left","count":1,"offset":300,"label":"SW"},
 {"cmd":"wire","points":["SW","L1","L2","L4","L3"]},
 {"cmd":"legend","x":0,"y":-900}
]
"""


def call_ollama(prompt, model=DEFAULT_MODEL, timeout=180):
    if requests is None:
        raise RuntimeError("requests が未インストールです。pip install requests を実行してください。")
    payload = {
        "model": model,
        "messages": [
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": prompt},
        ],
        "stream": False,
        "options": {
            "temperature": 0.1,
            "num_ctx": 4096,
        }
    }
    r = requests.post(OLLAMA_API_URL, json=payload, timeout=timeout)
    r.raise_for_status()
    data = r.json()
    text = data.get("message", {}).get("content", "")
    return text


# ============================================================
#  フォールバック: LLMなし簡易ルール変換
# ============================================================

def rule_based_convert(prompt):
    p = prompt.strip()
    cmds = []
    # 円
    if "円" in p:
        r = 100
        m = re.search(r"半径\s*(\d+(?:\.\d+)?)", p)
        if m:
            r = float(m.group(1))
        cmds.append({"cmd":"draw_circle", "center":[0,0], "radius":r})
        return cmds
    # 部屋サイズ
    m = re.search(r"(\d+(?:\.\d+)?)\s*[×xX*]\s*(\d+(?:\.\d+)?)", p)
    if m:
        w, h = float(m.group(1)), float(m.group(2))
        cmds.append({"cmd":"room_rect", "x":0, "y":0, "w":w, "h":h, "name":"部屋", "dimension":True})
    if any(k in p for k in ["照明", "ダウンライト", "DL", "ライト"]):
        if "4" in p or "四" in p or "2×2" in p or "2x2" in p:
            cmds.append({"cmd":"light_grid", "count_x":2, "count_y":2, "margin_x":500, "margin_y":500, "symbol":"L"})
        else:
            cmds.append({"cmd":"light", "x":1500, "y":1000, "label":"L1"})
    if "コンセント" in p or "CO" in p:
        cnt = 1
        m2 = re.search(r"コンセント\D*(\d+)", p)
        if m2: cnt = int(m2.group(1))
        cmds.append({"cmd":"outlet_wall", "wall":"bottom", "count":cnt, "offset":300, "label":"CO"})
    if "スイッチ" in p or "SW" in p:
        cmds.append({"cmd":"switch_wall", "wall":"left", "count":1, "offset":300, "label":"SW"})
    if "配線" in p:
        cmds.append({"cmd":"wire", "points":["SW", "L1", "L2", "L4", "L3"]})
    if "凡例" in p:
        cmds.append({"cmd":"legend", "x":0, "y":-900})
    if not cmds:
        cmds.append({"cmd":"text", "x":0, "y":0, "text":"作図指示を解釈できませんでした", "height":120})
    return cmds


# ============================================================
#  GUI
# ============================================================

class App:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("1120x820")
        self.q = queue.Queue()
        self.last_commands = []
        self.last_dxf = ""
        self.build_ui()
        self.root.after(100, self.poll_queue)

    def build_ui(self):
        top = ttk.Frame(self.root, padding=8)
        top.pack(fill=tk.X)
        ttk.Label(top, text=APP_TITLE, font=("Meiryo", 13, "bold")).pack(side=tk.LEFT)

        cfg = ttk.LabelFrame(self.root, text="設定", padding=8)
        cfg.pack(fill=tk.X, padx=8, pady=4)

        row1 = ttk.Frame(cfg); row1.pack(fill=tk.X, pady=2)
        ttk.Label(row1, text="Ollama exe:", width=14).pack(side=tk.LEFT)
        self.ollama_exe = tk.StringVar(value=DEFAULT_OLLAMA_EXE)
        ttk.Entry(row1, textvariable=self.ollama_exe).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=4)
        ttk.Button(row1, text="参照", command=lambda: self.browse_file(self.ollama_exe)).pack(side=tk.LEFT)
        ttk.Button(row1, text="serve起動", command=self.start_ollama_serve).pack(side=tk.LEFT, padx=4)

        row1b = ttk.Frame(cfg); row1b.pack(fill=tk.X, pady=2)
        ttk.Label(row1b, text="Ollama app:", width=14).pack(side=tk.LEFT)
        self.ollama_app_exe = tk.StringVar(value=DEFAULT_OLLAMA_APP_EXE)
        ttk.Entry(row1b, textvariable=self.ollama_app_exe).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=4)
        ttk.Button(row1b, text="参照", command=lambda: self.browse_file(self.ollama_app_exe)).pack(side=tk.LEFT)
        ttk.Button(row1b, text="app起動", command=self.start_ollama_app).pack(side=tk.LEFT, padx=4)

        row2 = ttk.Frame(cfg); row2.pack(fill=tk.X, pady=2)
        ttk.Label(row2, text="Jw_cad exe:", width=14).pack(side=tk.LEFT)
        self.jwcad_exe = tk.StringVar(value=DEFAULT_JWCAD_EXE)
        ttk.Entry(row2, textvariable=self.jwcad_exe).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=4)
        ttk.Button(row2, text="参照", command=lambda: self.browse_file(self.jwcad_exe)).pack(side=tk.LEFT)

        row3 = ttk.Frame(cfg); row3.pack(fill=tk.X, pady=2)
        ttk.Label(row3, text="モデル:", width=14).pack(side=tk.LEFT)
        self.model = tk.StringVar(value=DEFAULT_MODEL)
        ttk.Entry(row3, textvariable=self.model, width=24).pack(side=tk.LEFT, padx=4)
        ttk.Label(row3, text="出力先:").pack(side=tk.LEFT, padx=(20,4))
        self.out_dir = tk.StringVar(value=DEFAULT_OUT_DIR)
        ttk.Entry(row3, textvariable=self.out_dir).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=4)
        ttk.Button(row3, text="参照", command=self.browse_out_dir).pack(side=tk.LEFT)

        main = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        main.pack(fill=tk.BOTH, expand=True, padx=8, pady=6)

        left = ttk.Frame(main)
        right = ttk.Frame(main)
        main.add(left, weight=3)
        main.add(right, weight=2)

        prompt_frame = ttk.LabelFrame(left, text="チャット指示", padding=8)
        prompt_frame.pack(fill=tk.BOTH, expand=True)
        self.prompt = scrolledtext.ScrolledText(prompt_frame, height=10, font=("Meiryo", 11), wrap=tk.WORD)
        self.prompt.pack(fill=tk.BOTH, expand=True)
        self.prompt.insert("1.0", "4000×3000の部屋を描いて、中央にダウンライト4台を2×2で配置して、下壁にコンセント3個、左壁にスイッチ1個、スイッチから照明へ配線、凡例も付けて")

        btns = ttk.Frame(left)
        btns.pack(fill=tk.X, pady=6)
        ttk.Button(btns, text="Ollamaで作図", command=self.run_with_ollama).pack(side=tk.LEFT, padx=4)
        ttk.Button(btns, text="LLMなし簡易変換", command=self.run_rule_based).pack(side=tk.LEFT, padx=4)
        ttk.Button(btns, text="JSONから直接DXF生成", command=self.generate_from_json_box).pack(side=tk.LEFT, padx=4)
        ttk.Button(btns, text="Jw_cadで開く", command=self.open_last_dxf).pack(side=tk.RIGHT, padx=4)

        log_frame = ttk.LabelFrame(left, text="ログ", padding=8)
        log_frame.pack(fill=tk.BOTH, expand=True)
        self.log = scrolledtext.ScrolledText(log_frame, height=12, font=("Consolas", 10), wrap=tk.WORD)
        self.log.pack(fill=tk.BOTH, expand=True)

        json_frame = ttk.LabelFrame(right, text="作図コマンドJSON / 編集可", padding=8)
        json_frame.pack(fill=tk.BOTH, expand=True)
        self.json_box = scrolledtext.ScrolledText(json_frame, height=25, font=("Consolas", 10), wrap=tk.NONE)
        self.json_box.pack(fill=tk.BOTH, expand=True)

        sample_frame = ttk.LabelFrame(right, text="サンプル", padding=8)
        sample_frame.pack(fill=tk.X, pady=6)
        samples = [
            "原点を中心とした半径100の円を描いて",
            "3000×2000の部屋を描いて、中央にダウンライト4台を2×2で配置して、凡例も付けて",
            "4000×3000の部屋、下壁にコンセント3個、左壁にスイッチ1個、照明4台、スイッチから照明へ配線",
            "L1から右に1200、上に800の位置にDLを追加して",
        ]
        self.sample_var = tk.StringVar(value=samples[1])
        ttk.Combobox(sample_frame, textvariable=self.sample_var, values=samples, state="readonly").pack(fill=tk.X, side=tk.LEFT, expand=True)
        ttk.Button(sample_frame, text="入力へ", command=self.load_sample).pack(side=tk.LEFT, padx=4)

        self.status = ttk.Label(self.root, text="準備完了", relief=tk.SUNKEN, anchor="w")
        self.status.pack(fill=tk.X, side=tk.BOTTOM)

    def browse_file(self, var):
        fp = filedialog.askopenfilename(filetypes=[("EXE", "*.exe"), ("All", "*.*")])
        if fp:
            var.set(fp)

    def browse_out_dir(self):
        d = filedialog.askdirectory()
        if d:
            self.out_dir.set(d)

    def load_sample(self):
        self.prompt.delete("1.0", tk.END)
        self.prompt.insert("1.0", self.sample_var.get())

    def write_log(self, msg):
        self.log.insert(tk.END, str(msg) + "\n")
        self.log.see(tk.END)

    def start_ollama_serve(self):
        exe = self.ollama_exe.get().strip()
        if not os.path.exists(exe):
            messagebox.showerror("エラー", f"ollama.exe が見つかりません:\n{exe}")
            return
        try:
            subprocess.Popen([exe, "serve"], creationflags=subprocess.CREATE_NEW_CONSOLE if os.name == "nt" else 0)
            self.write_log("Ollama serve を起動しました。既に起動中ならエラー表示される場合がありますが問題ないことがあります。")
        except Exception as e:
            messagebox.showerror("エラー", str(e))

    def start_ollama_app(self):
        exe = self.ollama_app_exe.get().strip()
        if not os.path.exists(exe):
            messagebox.showerror("エラー", f"ollama app.exe が見つかりません:\n{exe}")
            return
        try:
            subprocess.Popen([exe], creationflags=subprocess.CREATE_NEW_CONSOLE if os.name == "nt" else 0)
            self.write_log("Ollama app.exe を起動しました。")
        except Exception as e:
            messagebox.showerror("エラー", str(e))

    def run_with_ollama(self):
        prompt = self.prompt.get("1.0", tk.END).strip()
        if not prompt:
            return
        threading.Thread(target=self._worker_ollama, args=(prompt,), daemon=True).start()

    def _worker_ollama(self, prompt):
        self.q.put(("status", "Ollama問い合わせ中..."))
        self.q.put(("log", "== Ollama request =="))
        try:
            raw = call_ollama(prompt, model=self.model.get().strip() or DEFAULT_MODEL)
            self.q.put(("log", "== Ollama raw response ==\n" + raw))
            commands = extract_json_from_text(raw)
            self.q.put(("commands", commands))
            self.generate_dxf(commands)
        except Exception as e:
            self.q.put(("log", "[Ollama失敗] " + str(e)))
            self.q.put(("log", "LLMなし簡易変換にフォールバックします。"))
            commands = rule_based_convert(prompt)
            self.q.put(("commands", commands))
            self.generate_dxf(commands)

    def run_rule_based(self):
        prompt = self.prompt.get("1.0", tk.END).strip()
        commands = rule_based_convert(prompt)
        self.show_commands(commands)
        self.generate_dxf(commands)

    def show_commands(self, commands):
        self.last_commands = commands
        self.json_box.delete("1.0", tk.END)
        self.json_box.insert("1.0", json.dumps(commands, ensure_ascii=False, indent=2))

    def generate_from_json_box(self):
        try:
            commands = json.loads(self.json_box.get("1.0", tk.END).strip())
            self.generate_dxf(commands)
        except Exception as e:
            messagebox.showerror("JSONエラー", str(e))

    def generate_dxf(self, commands):
        try:
            ensure_dir(self.out_dir.get())
            renderer = ElectricalCommandRenderer()
            dxf = renderer.render(commands)
            out_path = os.path.join(self.out_dir.get(), f"jwcad_ai_electrical_{now_stamp()}.dxf")
            dxf.save(out_path)
            self.last_dxf = out_path
            self.q.put(("log", f"DXF生成完了: {out_path}"))
            self.q.put(("status", "DXF生成完了"))
            # 自動オープン
            self.open_dxf(out_path, silent=True)
        except Exception as e:
            self.q.put(("log", "[DXF生成エラー] " + str(e)))
            self.q.put(("status", "エラー"))

    def open_last_dxf(self):
        if not self.last_dxf or not os.path.exists(self.last_dxf):
            messagebox.showwarning("注意", "まだDXFが生成されていません。")
            return
        self.open_dxf(self.last_dxf, silent=False)

    def open_dxf(self, path, silent=False):
        exe = self.jwcad_exe.get().strip()
        try:
            if os.path.exists(exe):
                subprocess.Popen([exe, path])
                self.q.put(("log", "Jw_cadで開きました: " + path))
            else:
                if os.name == "nt":
                    os.startfile(path)
                else:
                    subprocess.Popen(["xdg-open", path])
                self.q.put(("log", "関連付けで開きました: " + path))
        except Exception as e:
            if not silent:
                messagebox.showerror("Jw_cad起動エラー", str(e))
            self.q.put(("log", "[Jw_cad起動エラー] " + str(e)))

    def poll_queue(self):
        try:
            while True:
                kind, payload = self.q.get_nowait()
                if kind == "log":
                    self.write_log(payload)
                elif kind == "status":
                    self.status.config(text=payload)
                elif kind == "commands":
                    self.show_commands(payload)
        except queue.Empty:
            pass
        self.root.after(100, self.poll_queue)


if __name__ == "__main__":
    root = tk.Tk()
    try:
        ttk.Style().theme_use("clam")
    except Exception:
        pass
    App(root)
    root.mainloop()
