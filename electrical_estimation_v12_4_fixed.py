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

APP_TITLE = "電気設備積算アプリ v12.2 スレッド安全版 + 再学習・再推論エンジン"
BASE_DIR = Path(r"E:\electrical_estimation_ai")
if not BASE_DIR.exists():
    BASE_DIR = Path.cwd() / "electrical_estimation_ai"
BASE_DIR.mkdir(parents=True, exist_ok=True)
DB_PATH = str(BASE_DIR / "estimation_master_integrated.db")
CAD_LIBRARY_PATH = str(BASE_DIR / "cad_library")
OUTPUT_DIR = BASE_DIR / "outputs"
CROP_DIR = OUTPUT_DIR / "symbol_crops"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
CROP_DIR.mkdir(parents=True, exist_ok=True)
os.makedirs(CAD_LIBRARY_PATH, exist_ok=True)
DEFAULT_OLLAMA_URL = "http://127.0.0.1:11434"
DEFAULT_OLLAMA_MODEL = "deepseek-r1:8b"
DEFAULT_OPENAI_MODEL = "gpt-4o-mini"
DEFAULT_CLAUDE_MODEL = "claude-sonnet-4-5"
CLAUDE_MODEL_CHOICES = ["claude-sonnet-4-5","claude-haiku-4-5","claude-opus-4-1","claude-3-5-haiku-latest","claude-3-5-sonnet-latest"]
DEFAULT_CUSTOM_OPENAI_BASE_URL = "http://127.0.0.1:8000/v1"
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

# ---------------- DB ----------------
class DatabaseManager:
    def __init__(self, db_path):
        self.db_path = db_path
        self.init_db()
    def connect(self): return sqlite3.connect(self.db_path)
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
        defaults = {"ollama_url":DEFAULT_OLLAMA_URL,"ollama_model":DEFAULT_OLLAMA_MODEL,"pdf_scale":"1.25","analysis_max_width":"1000","preview_max_width":"850","scan_step":"4","cluster_distance":"18","min_cluster_points":"3","min_box_w":"4","min_box_h":"4","max_box_w":"120","max_box_h":"120","exclude_legend":"1","legend_right_ratio":"0.24","legend_bottom_ratio":"0.28","max_detections":"140","analysis_engine":"color","template_enabled":"1","opencv_enabled":"0","ocr_enabled":"0","yolo_enabled":"0","sam_enabled":"0","cnn_vit_enabled":"0"}
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
    def annotation_summary(self):
        con=self.connect(); rows=con.execute("SELECT final_answer,COUNT(*),SUM(is_llm_correct) FROM annotation_samples GROUP BY final_answer ORDER BY final_answer").fetchall(); con.close(); return rows

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

def merge_dets(dets,dist):
    out=[]; used=[False]*len(dets)
    for i,d in enumerate(dets):
        if used[i]: continue
        group=[d]; used[i]=True; changed=True
        while changed:
            changed=False
            for j,e in enumerate(dets):
                if used[j] or e.equipment!=d.equipment: continue
                if any(near(g,e,dist) for g in group): group.append(e); used[j]=True; changed=True
        if len(group)==1: out.append(d)
        else:
            out.append(ImageDetection(0,d.page,min(g.x1 for g in group),min(g.y1 for g in group),max(g.x2 for g in group),max(g.y2 for g in group),d.color_name,d.equipment,max(g.score for g in group),sum(g.pixel_count for g in group),True,'merged',f'merged {len(group)}'))
    for idx,d in enumerate(out,1): d.det_id=idx
    return out

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
    dets=merge_dets(dets,int(dist/max(0.1,ratio))); dets=sorted(dets,key=lambda d:d.pixel_count,reverse=True)[:max_dets]; dets=sorted(dets,key=lambda d:(d.page,d.y1,d.x1))
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


def extract_symbol_feature(crop):
    """
    再学習・再推論用の軽量特徴量。
    画像を保存するだけでなく、次回推論で使える数値特徴をSQLiteへ保存する。
    - RGBヒストグラム
    - 縦横比
    - OpenCV円形度/輪郭数（入っていれば）
    """
    im = crop.convert('RGB')
    w, h = im.size
    small = im.resize((32,32), Image.Resampling.BILINEAR)
    px = small.load()

    # 4x4x4 RGBヒストグラム
    hist = [0]*64
    color_counts = {'purple':0,'yellow':0,'green':0,'cyan':0,'blue':0,'red':0,'dark':0}
    for y in range(32):
        for x in range(32):
            r,g,b = px[x,y]
            ri=min(3,r//64); gi=min(3,g//64); bi=min(3,b//64)
            hist[ri*16+gi*4+bi] += 1
            if r>120 and g<150 and b>120: color_counts['purple'] += 1
            if r>160 and g>120 and b<150: color_counts['yellow'] += 1
            if r<150 and g>100 and b<170: color_counts['green'] += 1
            if r<140 and g>120 and b>120: color_counts['cyan'] += 1
            if r<130 and g<170 and b>120: color_counts['blue'] += 1
            if r>150 and g<140 and b<140: color_counts['red'] += 1
            if r<80 and g<80 and b<80: color_counts['dark'] += 1
    total = max(1, sum(hist))
    hist = [v/total for v in hist]
    major_color = max(color_counts, key=color_counts.get)

    feat = {
        'w': w,
        'h': h,
        'aspect': w/max(1,h),
        'hist64': hist,
        'major_color': major_color,
        'color_counts': color_counts,
    }

    cvfeat = opencv_shape_features(im)
    if cvfeat and isinstance(cvfeat, dict):
        feat['opencv'] = cvfeat
    else:
        feat['opencv'] = {}
    return feat

def feature_similarity(a, b):
    """
    0.0〜1.0の類似度。
    厳密なAIではなく、過去OK/NG登録を次回判定へ反映するための軽量近傍検索。
    """
    try:
        ha = a.get('hist64', [])
        hb = b.get('hist64', [])
        if not ha or not hb or len(ha) != len(hb):
            return 0.0
        # histogram intersection
        hist_sim = sum(min(x,y) for x,y in zip(ha,hb))

        asp_a = float(a.get('aspect',1))
        asp_b = float(b.get('aspect',1))
        aspect_sim = max(0.0, 1.0 - min(abs(asp_a-asp_b)/3.0, 1.0))

        color_sim = 1.0 if a.get('major_color') == b.get('major_color') else 0.65

        cva = a.get('opencv', {}) or {}
        cvb = b.get('opencv', {}) or {}
        circ_a = float(cva.get('circularity',0) or 0)
        circ_b = float(cvb.get('circularity',0) or 0)
        circ_sim = max(0.0, 1.0 - min(abs(circ_a-circ_b), 1.0)) if (circ_a or circ_b) else 0.5

        return max(0.0, min(1.0, hist_sim*0.55 + aspect_sim*0.20 + color_sim*0.15 + circ_sim*0.10))
    except Exception:
        return 0.0

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
    def generate(self,model,prompt,images=None,timeout=10):
        requests=try_import_requests()
        if not requests: raise RuntimeError('requestsが必要です')
        payload={
            'model':model,
            'prompt':prompt,
            'stream':False,
            'keep_alive':'1m',
            'options':{'num_ctx':768,'temperature':0.1,'num_predict':64}
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
    """dictスナップショット(スレッドセーフ)またはappインスタンス(メインスレッドのみ)で初期化できる"""
    def __init__(self, app_or_dict):
        if isinstance(app_or_dict, dict):
            self._d = app_or_dict; self.app = None
        else:
            self._d = None; self.app = app_or_dict
        self.last_response = AIResponse('', 'unknown')
    def _v(self, key, default=''):
        if self._d is not None: return self._d.get(key, default)
        m = {'provider':'ai_provider_var','ollama_url':'ollama_url_var','ollama_model':'ollama_model_var',
             'openai_api_key':'openai_api_key_var','openai_model':'openai_model_var',
             'anthropic_api_key':'anthropic_api_key_var','anthropic_model':'anthropic_model_var',
             'custom_openai_base_url':'custom_openai_base_url_var',
             'custom_openai_api_key':'custom_openai_api_key_var','custom_openai_model':'custom_openai_model_var'}
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
    def generate(self,prompt,timeout=30):
        p=self.provider()
        if p=='openai': return self.generate_openai(prompt,timeout)
        if p=='anthropic': return self.generate_anthropic(prompt,timeout)
        if p=='custom_openai': return self.generate_custom_openai(prompt,timeout)
        return self.generate_ollama(prompt,timeout)
    def generate_ollama(self,prompt,timeout=30):
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
        data=r.json(); txt='\\n'.join(p.get('text','') for p in data.get('content',[]) if isinstance(p,dict) and p.get('type')=='text').strip()
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
        self.drag_start=None; self.drag_rect_id=None; self.dragging=False; self.drag_mode_enabled=True
        self._suppress_tree_select=False; self._last_selected_det_id=None
        self._preview_rendering=False; self._preview_req=None
        self.chat_history=[]; self._chat_busy=False
        # ログキューはlog_q1本のみ。process_queuesは廃止 → root.after(0,fn)方式に統一
        self.log_q=queue.Queue()
        self.build_ui(); self.refresh_all(); self.root.after(100, self._drain_log_q)

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

    def apply_learning_to_detections(self, threshold=0.82):
        """
        解析直後の全検出候補に対して、過去のOK/NG学習データを反映する。
        ここが「再学習・再推論エンジン」の中心。
        """
        try:
            rows = self.db.get_annotation_features(limit=3000)
            if not rows or not self.image_pages:
                self.log('学習特徴量なし：通常推定のみ')
                return 0
            applied = 0
            for d in self.image_dets:
                try:
                    crop = crop_symbol(self.image_pages[d.page-1], d, margin=8)
                    label, score, row = best_learned_match_for_crop(crop, rows, threshold=threshold)
                    if label:
                        d.equipment = label
                        d.enabled = True
                        d.memo = f'learned match {score:.2f}'
                        d.score = max(float(getattr(d,'score',0) or 0), score)
                        applied += 1
                except Exception:
                    continue
            self.log(f'学習済み特徴量を再推論へ反映: {applied}件')
            return applied
        except Exception:
            self.log(traceback.format_exc())
            return 0

    def rebuild_learning_features(self):
        def worker():
            return self.db.rebuild_features_from_annotations()
        def done(res):
            added, skipped = res
            self.refresh_learning_summary()
            messagebox.showinfo('再構築完了', f'特徴量を再構築しました。\n追加: {added}件\nスキップ: {skipped}件')
        self.run_in_thread(worker, done)

    def local_guess_from_detection(self, d):
        """AI未接続時/ドラッグ範囲のローカル推定。色＋テンプレート＋OpenCV特徴量。"""
        if not d:
            return '未分類設備'
        try:
            if self.image_pages:
                crop = crop_symbol(self.image_pages[d.page-1], d, margin=8)
                # まず過去学習済み特徴量で近傍検索する
                try:
                    rows = self.db.get_annotation_features(limit=3000)
                    label, score, row = best_learned_match_for_crop(crop, rows, threshold=0.82)
                    if label:
                        d.memo = f'learned local guess {score:.2f}'
                        return label
                except Exception:
                    pass
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
        except Exception as e:
            messagebox.showerror('DB保存エラー', str(e))
            return False
        try:
            self.refresh_image_view()
            self.refresh_image_tables()
            self.refresh_learning_summary()
        except Exception:
            self.log(traceback.format_exc())
        return True

    def build_ui(self):
        menubar=tk.Menu(self.root); self.root.config(menu=menubar); fm=tk.Menu(menubar,tearoff=0); menubar.add_cascade(label='ファイル',menu=fm); fm.add_command(label='積算結果CSV出力',command=self.export_result_csv); fm.add_separator(); fm.add_command(label='終了',command=self.root.quit)
        tm=tk.Menu(menubar,tearoff=0); menubar.add_cascade(label='ツール',menu=tm)
        for pkg in ['requests','pypdf','PyMuPDF','ezdxf','chardet','pillow','opencv-python','pytesseract','ultralytics']: tm.add_command(label=f'{pkg} インストール',command=lambda p=pkg:self.install_pkg_async(p))
        ttk.Label(self.root,text=APP_TITLE,font=('Arial',14,'bold')).pack(pady=5)
        self.nb=ttk.Notebook(self.root); self.nb.pack(fill=tk.BOTH,expand=True,padx=6,pady=4)
        self.tab_est=ttk.Frame(self.nb); self.tab_price=ttk.Frame(self.nb); self.tab_sym=ttk.Frame(self.nb); self.tab_cad=ttk.Frame(self.nb); self.tab_img=ttk.Frame(self.nb); self.tab_tune=ttk.Frame(self.nb); self.tab_ai=ttk.Frame(self.nb); self.tab_chat=ttk.Frame(self.nb); self.tab_learn=ttk.Frame(self.nb)
        for tab,name in [(self.tab_est,'積算実行 PDF/DXF'),(self.tab_price,'単価マスター'),(self.tab_sym,'記号パターン'),(self.tab_cad,'CADライブラリ'),(self.tab_img,'画像PDF解析・アノテーション'),(self.tab_tune,'画像チューニング'),(self.tab_ai,'Ollama解析'),(self.tab_chat,'AIチャット'),(self.tab_learn,'学習データ')]: self.nb.add(tab,text=name)
        self.build_est_tab(); self.build_price_tab(); self.build_sym_tab(); self.build_cad_tab(); self.build_image_tab(); self.build_tune_tab(); self.build_ai_tab(); self.build_chat_tab(); self.build_learning_tab()
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
    def build_image_tab(self):
        f=self.tab_img; top=ttk.LabelFrame(f,text='画像PDF/画像図面',padding=6); top.pack(fill=tk.X,padx=6,pady=4); self.image_file_var=tk.StringVar(); ttk.Entry(top,textvariable=self.image_file_var).pack(side=tk.LEFT,fill=tk.X,expand=True,padx=3); ttk.Button(top,text='選択',command=self.select_image_file).pack(side=tk.LEFT,padx=2); ttk.Button(top,text='軽量解析',command=self.start_image_analysis).pack(side=tk.LEFT,padx=2); ttk.Button(top,text='画像積算CSV',command=self.export_image_csv).pack(side=tk.LEFT,padx=2)
        main=ttk.PanedWindow(f,orient=tk.HORIZONTAL); main.pack(fill=tk.BOTH,expand=True,padx=6,pady=4); left=ttk.Frame(main); main.add(left,weight=3); bar=ttk.Frame(left); bar.pack(fill=tk.X); ttk.Button(bar,text='前ページ',command=lambda:self.change_image_page(-1)).pack(side=tk.LEFT); ttk.Button(bar,text='次ページ',command=lambda:self.change_image_page(1)).pack(side=tk.LEFT,padx=3); self.page_label=tk.StringVar(value='page -/-'); ttk.Label(bar,textvariable=self.page_label).pack(side=tk.LEFT,padx=8); ttk.Label(bar,text='左ドラッグ=範囲指定 / クリック=選択 / 右クリック・AI判定=外部AI').pack(side=tk.RIGHT)
        preview_frame=ttk.Frame(left)
        preview_frame.pack(fill=tk.BOTH,expand=True)
        self.preview_canvas=tk.Canvas(preview_frame,bg='white',width=850,height=620)
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
        right=ttk.PanedWindow(main,orient=tk.VERTICAL); main.add(right,weight=2); ann=ttk.LabelFrame(right,text='アノテーション登録・学習',padding=6); right.add(ann,weight=0); self.selected_info_var=tk.StringVar(value='未選択'); ttk.Label(ann,textvariable=self.selected_info_var).grid(row=0,column=0,columnspan=4,sticky='w'); ttk.Label(ann,text='LLM回答').grid(row=1,column=0,sticky='e'); self.llm_answer_var=tk.StringVar(); ttk.Entry(ann,textvariable=self.llm_answer_var,width=28).grid(row=1,column=1,sticky='w'); ttk.Button(ann,text='OK登録',command=self.register_llm_ok).grid(row=1,column=2,padx=3); ttk.Label(ann,text='訂正').grid(row=2,column=0,sticky='e'); self.correct_answer_var=tk.StringVar(); ttk.Combobox(ann,textvariable=self.correct_answer_var,values=['LEDダウンライト','LEDベースライト','コンセント','片切スイッチ','分電盤','換気扇','未分類設備'],width=26).grid(row=2,column=1,sticky='w'); ttk.Button(ann,text='訂正登録・学習',command=self.register_manual_correction).grid(row=2,column=2,padx=3); ttk.Button(ann,text='選択を無効化',command=self.disable_image_selected).grid(row=1,column=3,padx=3); ttk.Button(ann,text='設備名だけ変更',command=self.apply_correct_to_selected).grid(row=2,column=3,padx=3); self.drag_mode_var=tk.BooleanVar(value=True)
        ttk.Checkbutton(ann,text='左ドラッグで範囲指定',variable=self.drag_mode_var).grid(row=3,column=0,sticky='w',padx=3,pady=3)
        ttk.Button(ann,text='AI判定実行',command=self.annotate_selected_popup_async).grid(row=3,column=1,padx=3,pady=3)
        ttk.Label(ann,text='※ドラッグ範囲は即ローカル推定。必要時だけAI判定実行。').grid(row=3,column=2,columnspan=2,sticky='w')
        self.img_ai_status_var=tk.StringVar(value='AI状態: 未実行')
        ttk.Label(ann,textvariable=self.img_ai_status_var,foreground='blue').grid(row=4,column=0,columnspan=4,sticky='w',pady=2)
        self.img_ai_usage_var=tk.StringVar(value='AI使用量: 未使用')
        ttk.Label(ann,textvariable=self.img_ai_usage_var,foreground='purple').grid(row=5,column=0,columnspan=4,sticky='w',pady=2)
        detf=ttk.LabelFrame(right,text='画像検出候補',padding=4); right.add(detf,weight=3); cols=('on','id','page','equipment','color','x','y','w','h','score'); self.img_det_tree=ttk.Treeview(detf,columns=cols,show='headings',height=12)
        for c in cols: self.img_det_tree.heading(c,text=c); self.img_det_tree.column(c,width=70,anchor=tk.CENTER)
        self.img_det_tree.column('equipment',width=130); self.img_det_tree.pack(fill=tk.BOTH,expand=True); self.img_det_tree.bind('<<TreeviewSelect>>',self.on_img_tree_select); self.img_det_tree.bind('<Double-1>',lambda e:self.annotate_selected_async())
        sumf=ttk.LabelFrame(right,text='画像積算結果',padding=4); right.add(sumf,weight=2); cols2=('カテゴリ','品名','単位','数量','単価','金額'); self.img_sum_tree=ttk.Treeview(sumf,columns=cols2,show='headings',height=7)
        for c in cols2: self.img_sum_tree.heading(c,text=c); self.img_sum_tree.column(c,width=90,anchor=tk.E if c in ('数量','単価','金額') else tk.W)
        self.img_sum_tree.pack(fill=tk.BOTH,expand=True)
    def build_tune_tab(self):
        f=self.tab_tune; sf=ttk.LabelFrame(f,text='画像解析 軽量化/精度設定',padding=8); sf.pack(fill=tk.X,padx=6,pady=6); self.setting_vars={}; items=[('pdf_scale','PDF画像化倍率'),('analysis_max_width','解析最大幅'),('preview_max_width','プレビュー最大幅'),('scan_step','走査間隔 1高精度/4軽量'),('cluster_distance','結合距離'),('min_cluster_points','最小点数'),('min_box_w','最小幅'),('min_box_h','最小高'),('max_box_w','最大幅'),('max_box_h','最大高'),('exclude_legend','凡例除外 1/0'),('legend_right_ratio','凡例右比率'),('legend_bottom_ratio','凡例下比率'),('max_detections','最大候補数'),('analysis_engine','解析エンジン color/template/opencv/yolo/sam/ocr'),('template_enabled','テンプレート 1/0'),('opencv_enabled','OpenCV 1/0'),('ocr_enabled','OCR 1/0'),('yolo_enabled','YOLO 1/0'),('sam_enabled','SAM2 1/0'),('cnn_vit_enabled','CNN/ViT 1/0')]
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
        ttk.Button(conn,text='選択AIテスト',command=self.test_selected_ai_async).grid(row=0,column=3,padx=4); self.ai_usage_var=tk.StringVar(value='使用量: -'); ttk.Label(conn,textvariable=self.ai_usage_var,foreground='blue').grid(row=0,column=4,columnspan=2,sticky='w')

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

    def build_learning_tab(self):
        f=self.tab_learn
        btn=ttk.Frame(f); btn.pack(fill=tk.X,padx=6,pady=6)
        ttk.Button(btn,text='更新',command=self.refresh_learning_summary).pack(side=tk.LEFT)
        ttk.Button(btn,text='学習CSV出力',command=self.export_learning_csv).pack(side=tk.LEFT,padx=4)
        ttk.Button(btn,text='過去アノテーションから特徴量再構築',command=self.rebuild_learning_features).pack(side=tk.LEFT,padx=4)
        ttk.Button(btn,text='現在の検出へ学習反映',command=lambda:(self.apply_learning_to_detections(), self.refresh_image_view(), self.refresh_image_tables())).pack(side=tk.LEFT,padx=4)
        self.learning_text=scrolledtext.ScrolledText(f)
        self.learning_text.pack(fill=tk.BOTH,expand=True,padx=6,pady=6)

    # behaviors
    def refresh_all(self): self.price_refresh(); self.symbol_refresh(); self.cad_refresh(); self.refresh_color_rules(); self.refresh_learning_summary()
    def select_file(self):
        p=filedialog.askopenfilename(filetypes=[('PDF/DXF','*.pdf;*.dxf'),('All','*.*')]);
        if p: self.file_path_var.set(p)
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
            prompt=(
                'あなたは日本の電気設備積算の補助AIです。\n'
                '以下のPDF/DXF解析結果をもとに、過検出・不足・見積上の注意点を簡潔に指摘してください。\n'
                'なお、これは法的・最終見積ではなく、積算補助です。\n\n'
                f'対象ファイル: {p}\n'
                f'解析詳細: {details}\n'
                f'数量候補: {json.dumps(counts,ensure_ascii=False)}\n'
                f'積算行: {json.dumps(rows,ensure_ascii=False)}\n'
            )
            client=UnifiedAIClient(snap)
            ai_ans=client.generate(prompt,timeout=45)
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
        if p: self.source_image_file=p; threading.Thread(target=self.image_analysis_worker,args=(p,),daemon=True).start()
    def image_analysis_worker(self,p):
        try:
            self.log(f'画像/DXF解析開始: {p}'); s=self.db.get_settings(); all_d=[]
            if Path(p).suffix.lower()=='.dxf':
                pages, all_d = dxf_to_preview_image_and_detections(p, self.db, max_size=int(s.get('preview_max_width','850')), log=self.log)
            else:
                pages=load_pdf_or_image(p,float(s.get('pdf_scale','1.25')),log=self.log); rules=self.db.get_image_rules()
                for idx,img in enumerate(pages,1): all_d += analyze_image_page_light(img,idx,s,rules,log=self.log)
            for i,d in enumerate(all_d,1): d.det_id=i
            self.ui(self.finish_image_analysis,pages,all_d)
        except Exception: self.log(traceback.format_exc())
    def finish_image_analysis(self,pages,dets):
        self.image_pages=pages
        self.image_dets=dets
        self.current_page=1
        self.selected_det_id=None
        learned_count = self.apply_learning_to_detections(threshold=0.82)
        self.refresh_image_view()
        self.refresh_image_tables()
        self.status.config(text=f'画像解析完了: {len(dets)}件 / 学習反映 {learned_count}件')
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
                        self.preview_canvas.config(scrollregion=(0,0,preview.width,preview.height))
                        self.page_label.set(f'page {page}/{total}')
                    except Exception:
                        self.log(traceback.format_exc())
                self.ui(_up)
                if self._preview_req is None: break
        except Exception: traceback.print_exc()
        finally: self._preview_rendering=False
    def refresh_image_tables(self):
        for i in self.img_det_tree.get_children(): self.img_det_tree.delete(i)
        for d in self.image_dets[:300]: self.img_det_tree.insert('',tk.END,iid=str(d.det_id),values=('○' if d.enabled else '×',d.det_id,d.page,d.equipment,d.color_name,d.x1,d.y1,d.w,d.h,f'{d.score:.2f}'))
        for i in self.img_sum_tree.get_children(): self.img_sum_tree.delete(i)
        for r in self.image_estimate_rows(): self.img_sum_tree.insert('',tk.END,values=r)
    def image_estimate_rows(self):
        counts=defaultdict(int)
        for d in self.image_dets:
            if d.enabled: counts[d.equipment]+=1
        rows=[]
        for name,qty in sorted(counts.items()): cat,item,spec,unit,price=self.db.find_price(name); rows.append((cat,item,unit,qty,int(price),int(float(price)*qty)))
        return rows
    def change_image_page(self,delta):
        if not self.image_pages: return
        self.current_page=max(1,min(len(self.image_pages),self.current_page+delta)); self.selected_det_id=None; self.refresh_image_view()
    def canvas_xy_to_img(self,x,y): return int(self.preview_canvas.canvasx(x)/max(0.0001,self.preview_scale)), int(self.preview_canvas.canvasy(y)/max(0.0001,self.preview_scale))
    def find_det_at_canvas(self,x,y):
        ix,iy=self.canvas_xy_to_img(x,y); cand=[d for d in self.image_dets if d.page==self.current_page and d.x1<=ix<=d.x2 and d.y1<=iy<=d.y2]
        return sorted(cand,key=lambda d:d.w*d.h)[0] if cand else None
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

    def on_preview_drag(self,event):
        if not self.image_pages:
            return
        try:
            if hasattr(self,'drag_mode_var') and not self.drag_mode_var.get():
                return
        except Exception:
            pass
        cx=self.preview_canvas.canvasx(event.x)
        cy=self.preview_canvas.canvasy(event.y)
        if self.drag_start is None:
            self.drag_start=(cx,cy)
            self.dragging=False
            return
        x0,y0=self.drag_start
        if abs(cx-x0)+abs(cy-y0) < 6:
            return
        self.dragging=True
        try:
            if self.drag_rect_id:
                self.preview_canvas.delete(self.drag_rect_id)
            self.drag_rect_id=self.preview_canvas.create_rectangle(x0,y0,cx,cy,outline='red',width=2,dash=(4,2))
        except Exception:
            pass

    def on_preview_drag_release(self,event):
        if not self.image_pages:
            return
        if not getattr(self,'dragging',False):
            self.drag_start=None
            return
        try:
            if self.drag_rect_id:
                self.preview_canvas.delete(self.drag_rect_id)
                self.drag_rect_id=None
        except Exception:
            pass
        x0,y0=self.drag_start
        self.drag_start=None
        self.dragging=False
        ix0,iy0=self.canvas_xy_to_img(x0,y0)
        ix1,iy1=self.canvas_xy_to_img(event.x,event.y)
        xa,xb=sorted([ix0,ix1]); ya,yb=sorted([iy0,iy1])
        if xb-xa < 8 or yb-ya < 8:
            return
        img=self.image_pages[self.current_page-1]
        det=ImageDetection(len(self.image_dets)+1,self.current_page,max(0,xa),max(0,ya),min(img.width-1,xb),min(img.height-1,yb),'manual','未分類設備',0.5,0,True,'drag','drag annotation range')
        self.image_dets.append(det)
        self.refresh_image_tables()
        self.select_image_det(det.det_id)

        # ドラッグ範囲指定ではLLMへ自動問い合わせしない。
        # まずローカル推定を即表示し、OK/NG登録できるようにする。
        guess = self.local_guess_from_detection(det)
        det.equipment = guess
        det.memo = 'drag local annotation'
        self.llm_answer_var.set(guess)
        self.correct_answer_var.set(guess)
        self.selected_info_var.set(
            f'ドラッグ範囲 ID={det.det_id} page={det.page} 推定={guess} bbox=({det.x1},{det.y1})-({det.x2},{det.y2})'
        )
        self.refresh_image_view()
        self.refresh_image_tables()
        self.status.config(text=f'ドラッグ範囲をローカル推定: {guess}。OK/NGで学習できます。')

    def on_preview_click(self,event):
        # 左クリックは「選択だけ」。AI問い合わせや重い処理は絶対に走らせない。
        # Treeview選択イベントとの再帰を防ぐため、select_image_det側で抑制する。
        try:
            d=self.find_det_at_canvas(event.x,event.y)
            if d:
                self.select_image_det(d.det_id, update_tree=True, redraw=True)
            return 'break'
        except Exception:
            self.log(traceback.format_exc())
            return 'break'
    def on_preview_double(self,event):
        d=self.find_det_at_canvas(event.x,event.y)
        if not d and self.image_pages:
            ix,iy=self.canvas_xy_to_img(event.x,event.y); img=self.image_pages[self.current_page-1]; box=30; d=ImageDetection(len(self.image_dets)+1,self.current_page,max(0,ix-box),max(0,iy-box),min(img.width-1,ix+box),min(img.height-1,iy+box),'manual','未分類設備',0.5,0,True,'manual','manual'); self.image_dets.append(d); self.refresh_image_tables()
        if d: self.select_image_det(d.det_id); self.annotate_selected_popup_async()
    def on_preview_right_click(self,event):
        """
        右シングルクリック:
        - 既存検出枠上ならその図記号を選択
        - 検出枠がない場所なら小さい手動枠を作成
        - 選択図記号だけをOllamaへ送り、AI回答ポップアップを表示
        """
        d=self.find_det_at_canvas(event.x,event.y)
        if not d and self.image_pages:
            ix,iy=self.canvas_xy_to_img(event.x,event.y)
            img=self.image_pages[self.current_page-1]
            box=30
            d=ImageDetection(
                len(self.image_dets)+1,
                self.current_page,
                max(0,ix-box),
                max(0,iy-box),
                min(img.width-1,ix+box),
                min(img.height-1,iy+box),
                'manual',
                '未分類設備',
                0.5,
                0,
                True,
                'manual',
                'right click manual'
            )
            self.image_dets.append(d)
            self.refresh_image_tables()
        if d:
            self.select_image_det(d.det_id)
            self.annotate_selected_popup_async(event.x_root,event.y_root)

    def annotate_selected_popup_async(self,screen_x=None,screen_y=None):
        d=self.get_image_det(self.selected_det_id)
        if not d:
            messagebox.showwarning('未選択','図記号を選択してください')
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
        prompt=(
            f"あなたは日本の電気工事図面の図記号判定補助AIです。\n"
            f"候補: LEDダウンライト, LEDベースライト, コンセント, 片切スイッチ, 分電盤, 換気扇, 未分類設備\n"
            f"特徴: 色={d.color_name}, 自動推定={d.equipment}, page={d.page}, "
            f"bbox=({d.x1},{d.y1})-({d.x2},{d.y2}), 幅={d.w}, 高さ={d.h}\n"
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
        messagebox.showinfo('保存','AI設定を保存しました')

    def test_selected_ai_async(self):
        self.save_ollama_settings_silent()  # テスト前に自動保存
        snap = self.snapshot_ai_settings()
        threading.Thread(target=self.test_selected_ai_worker, args=(snap,), daemon=True).start()

    def test_selected_ai_worker(self, snap=None):
        snap = snap or {'provider':'ollama','ollama_url':DEFAULT_OLLAMA_URL,'ollama_model':DEFAULT_OLLAMA_MODEL}
        try:
            ans = UnifiedAIClient(snap).generate(
                '接続テストです。日本語で「接続OK」とだけ答えてください。',
                timeout=20
            )
            def ok():
                try:
                    if hasattr(self, 'img_ai_status_var'):
                        self.img_ai_status_var.set(f"AI状態: {snap.get('provider','AI')} 接続OK")
                except Exception:
                    pass
                messagebox.showinfo('AI接続OK', ans[:500])
            self.ui(ok)
        except Exception as e:
            msg = str(e)
            def ng(msg=msg):
                try:
                    if hasattr(self, 'img_ai_status_var'):
                        self.img_ai_status_var.set('AI状態: 接続エラー')
                except Exception:
                    pass
                provider_name=snap.get("provider","")
                # 401 = 認証エラー（APIキーが無効）
                if '401' in msg:
                    detail = (
                        f'provider={provider_name}\n\n'
                        '【原因】APIキーが無効または空です。\n\n'
                        '【確認方法】\n'
                        '1. 「AI設定保存」を押してから再テスト\n'
                        '2. 「表示」ボタンでキーの内容を確認\n'
                        '3. Anthropic/OpenAIのダッシュボードで\n'
                        '   有効なAPIキーをコピーして「貼付」\n\n'
                        f'error={msg[:400]}'
                    )
                elif '403' in msg:
                    detail = (
                        f'provider={provider_name}\n\n'
                        '【原因】APIキーの権限が不足しています。\n'
                        'APIキーのスコープを確認してください。\n\n'
                        f'error={msg[:400]}'
                    )
                elif 'timeout' in msg.lower() or 'timed out' in msg.lower():
                    detail = (
                        f'provider={provider_name}\n\n'
                        '【原因】接続タイムアウトです。\n'
                        'ネットワーク接続を確認してください。\n\n'
                        f'error={msg[:400]}'
                    )
                else:
                    detail = (
                        f'provider={provider_name}\n'
                        f'error={msg[:1200]}'
                    )
                messagebox.showwarning('AI接続エラー', '選択中のAIに接続できません。\n\n'+detail)
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
                    if current not in full:
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
        lines += [
            '',
            '使い方:',
            '1. 図記号をOK/NG登録すると、crop画像の特徴量も保存されます。',
            '2. 次回PDF/DXF/画像を解析すると、保存済み特徴量と類似検索してラベルを自動反映します。',
            '3. 既存の古い学習データは「過去アノテーションから特徴量再構築」で反映できます。',
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
    root=tk.Tk(); IntegratedApp(root); root.mainloop()
if __name__ == '__main__': main()
