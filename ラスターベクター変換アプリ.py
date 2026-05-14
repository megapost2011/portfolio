import os
import math
import time
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import cv2
import ezdxf
import numpy as np
import pypdfium2 as pdfium
import requests
from PIL import Image, ImageTk

try:
    import easyocr
    EASYOCR_AVAILABLE = True
except Exception:
    EASYOCR_AVAILABLE = False


class RasterElectricalDXFApp:
    def __init__(self, root):
        self.root = root
        self.root.title("画像/PDF/DXF → 電気設備DXF変換 + Ollama連携")
        self.root.geometry("1550x950")

        self.image_path = None
        self.input_mode = None
        self.original_image = None
        self.preview_image = None
        self.tk_img = None

        self.lines = []
        self.circles = []
        self.rectangles = []
        self.polylines = []
        self.symbols = []
        self.texts = []
        self.completed_lines = []
        self.routes = []

        self.dxf_bounds = None
        self.ocr_reader = None

        self.create_widgets()

    def create_widgets(self):
        main = tk.Frame(self.root)
        main.pack(fill=tk.BOTH, expand=True)

        top = tk.Frame(main)
        top.pack(side=tk.TOP, fill=tk.X)

        tk.Button(top, text="画像/PDF/DXFを開く", command=self.open_file).pack(side=tk.LEFT, padx=3, pady=3)
        tk.Button(top, text="ベクター検出", command=self.detect_vectors).pack(side=tk.LEFT, padx=3)
        tk.Button(top, text="欠損補完", command=self.auto_complete_shapes).pack(side=tk.LEFT, padx=3)
        tk.Button(top, text="構造解析", command=self.analyze_structure).pack(side=tk.LEFT, padx=3)
        tk.Button(top, text="OCR", command=self.run_ocr).pack(side=tk.LEFT, padx=3)
        tk.Button(top, text="DXF出力", command=self.export_dxf).pack(side=tk.LEFT, padx=3)
        tk.Button(top, text="リセット", command=self.reset).pack(side=tk.LEFT, padx=3)

        self.detect_lines_var = tk.BooleanVar(value=True)
        self.detect_circles_var = tk.BooleanVar(value=True)
        self.detect_rects_var = tk.BooleanVar(value=True)
        self.detect_curves_var = tk.BooleanVar(value=True)
        self.detect_symbols_var = tk.BooleanVar(value=True)

        for text, var in [
            ("直線", self.detect_lines_var),
            ("円", self.detect_circles_var),
            ("矩形", self.detect_rects_var),
            ("曲線", self.detect_curves_var),
            ("記号候補", self.detect_symbols_var),
        ]:
            tk.Checkbutton(top, text=text, variable=var).pack(side=tk.LEFT)

        ctrl = tk.Frame(main)
        ctrl.pack(side=tk.TOP, fill=tk.X)

        self.threshold_var = tk.IntVar(value=180)
        self.min_line_var = tk.IntVar(value=30)
        self.curve_epsilon_var = tk.DoubleVar(value=2.0)
        self.pdf_scale_var = tk.DoubleVar(value=4.0)
        self.dxf_scale_var = tk.DoubleVar(value=0.1)
        self.complete_gap_var = tk.IntVar(value=20)
        self.symbol_min_area_var = tk.IntVar(value=30)
        self.symbol_max_area_var = tk.IntVar(value=3000)

        self.add_slider(ctrl, "二値化", self.threshold_var, 50, 250, 1)
        self.add_slider(ctrl, "最小線長", self.min_line_var, 5, 500, 1)
        self.add_slider(ctrl, "曲線近似", self.curve_epsilon_var, 0.5, 10.0, 0.5)
        self.add_slider(ctrl, "PDF倍率", self.pdf_scale_var, 1.0, 6.0, 0.5)
        self.add_slider(ctrl, "DXF縮尺", self.dxf_scale_var, 0.01, 1.0, 0.01)
        self.add_slider(ctrl, "補完距離", self.complete_gap_var, 5, 100, 1)

        layer_frame = tk.LabelFrame(main, text="DXFレイヤ名")
        layer_frame.pack(side=tk.TOP, fill=tk.X, padx=4, pady=3)

        self.layer_vars = {
            "border": tk.StringVar(value="BORDER"),
            "lines": tk.StringVar(value="配線_直線"),
            "completed": tk.StringVar(value="AI補完線"),
            "circles": tk.StringVar(value="照明_円記号"),
            "rectangles": tk.StringVar(value="矩形_器具"),
            "curves": tk.StringVar(value="曲線_輪郭"),
            "symbols": tk.StringVar(value="電気記号候補"),
            "texts": tk.StringVar(value="OCR_TEXT"),
            "routes": tk.StringVar(value="配線ルート推定"),
            "dxf_import": tk.StringVar(value="DXF取込要素"),
        }

        for key, var in self.layer_vars.items():
            tk.Label(layer_frame, text=key).pack(side=tk.LEFT)
            tk.Entry(layer_frame, textvariable=var, width=12).pack(side=tk.LEFT, padx=2)

        ollama_frame = tk.LabelFrame(main, text="Ollama / LLM連携")
        ollama_frame.pack(side=tk.TOP, fill=tk.X, padx=4, pady=3)

        tk.Label(ollama_frame, text="ollama.exe").pack(side=tk.LEFT)
        self.ollama_path_var = tk.StringVar(value=r"E:\Ollama\ollama.exe")
        tk.Entry(ollama_frame, textvariable=self.ollama_path_var, width=36).pack(side=tk.LEFT, padx=3)
        tk.Button(ollama_frame, text="参照", command=self.select_ollama_exe).pack(side=tk.LEFT)
        tk.Button(ollama_frame, text="Ollama起動", command=self.start_ollama).pack(side=tk.LEFT, padx=3)
        tk.Button(ollama_frame, text="接続確認", command=self.check_ollama).pack(side=tk.LEFT, padx=3)
        tk.Button(ollama_frame, text="モデル取得", command=self.load_ollama_models).pack(side=tk.LEFT, padx=3)

        self.model_var = tk.StringVar(value="")
        self.model_combo = ttk.Combobox(ollama_frame, textvariable=self.model_var, width=24)
        self.model_combo.pack(side=tk.LEFT, padx=3)

        tk.Button(ollama_frame, text="AI解析", command=self.ai_analyze).pack(side=tk.LEFT, padx=3)
        tk.Button(ollama_frame, text="AIレイヤ提案", command=self.ai_layer_suggest).pack(side=tk.LEFT, padx=3)

        pane = tk.PanedWindow(main, orient=tk.VERTICAL)
        pane.pack(fill=tk.BOTH, expand=True)

        self.canvas = tk.Canvas(pane, bg="gray")
        pane.add(self.canvas)

        bottom = tk.Frame(pane)
        pane.add(bottom)

        tk.Label(bottom, text="ログ / AI回答").pack(anchor="w")
        self.log_text = tk.Text(bottom, height=11)
        self.log_text.pack(fill=tk.BOTH, expand=True)

        self.status = tk.StringVar(value="画像 / PDF / DXF を開いてください。")
        tk.Label(self.root, textvariable=self.status, anchor="w").pack(side=tk.BOTTOM, fill=tk.X)

    def add_slider(self, parent, label, var, frm, to, res):
        tk.Label(parent, text=label).pack(side=tk.LEFT, padx=2)
        tk.Scale(
            parent,
            from_=frm,
            to=to,
            resolution=res,
            orient=tk.HORIZONTAL,
            variable=var,
            length=105
        ).pack(side=tk.LEFT)

    def open_file(self):
        path = filedialog.askopenfilename(
            filetypes=[
                ("Supported", "*.png *.jpg *.jpeg *.bmp *.tif *.tiff *.webp *.pdf *.dxf"),
                ("Image", "*.png *.jpg *.jpeg *.bmp *.tif *.tiff *.webp"),
                ("PDF", "*.pdf"),
                ("DXF", "*.dxf"),
                ("All files", "*.*")
            ]
        )
        if not path:
            return

        ext = os.path.splitext(path)[1].lower()
        self.image_path = path

        if ext == ".dxf":
            self.load_dxf_file(path)
            return

        try:
            self.input_mode = "raster"
            self.reset_vectors()

            if ext == ".pdf":
                pdf = pdfium.PdfDocument(path)
                if len(pdf) == 0:
                    raise RuntimeError("PDFにページがありません。")
                page = pdf[0]
                bitmap = page.render(scale=float(self.pdf_scale_var.get()))
                pil_img = bitmap.to_pil().convert("RGB")
            else:
                pil_img = Image.open(path).convert("RGB")

            self.original_image = cv2.cvtColor(np.array(pil_img), cv2.COLOR_RGB2BGR)
            self.preview_image = self.original_image.copy()
            self.show_image(self.original_image)

            h, w = self.original_image.shape[:2]
            self.log(f"読み込み完了: {path}")
            self.status.set(f"読み込み完了: {w} x {h}px")

        except Exception as e:
            messagebox.showerror("読み込みエラー", str(e))

    def load_dxf_file(self, path):
        try:
            self.input_mode = "dxf"
            self.reset_vectors()

            doc = ezdxf.readfile(path)
            msp = doc.modelspace()

            min_x = min_y = 10**18
            max_x = max_y = -10**18

            def update_bounds(x, y):
                nonlocal min_x, min_y, max_x, max_y
                min_x = min(min_x, float(x))
                min_y = min(min_y, float(y))
                max_x = max(max_x, float(x))
                max_y = max(max_y, float(y))

            for e in msp:
                t = e.dxftype()

                try:
                    if t == "LINE":
                        x1, y1, _ = e.dxf.start
                        x2, y2, _ = e.dxf.end
                        self.lines.append((float(x1), float(y1), float(x2), float(y2)))
                        update_bounds(x1, y1)
                        update_bounds(x2, y2)

                    elif t == "CIRCLE":
                        x, y, _ = e.dxf.center
                        r = e.dxf.radius
                        self.circles.append((float(x), float(y), float(r)))
                        update_bounds(x - r, y - r)
                        update_bounds(x + r, y + r)

                    elif t == "ARC":
                        x, y, _ = e.dxf.center
                        r = e.dxf.radius
                        start = math.radians(float(e.dxf.start_angle))
                        end = math.radians(float(e.dxf.end_angle))
                        pts = []
                        steps = 24
                        if end < start:
                            end += math.tau
                        for i in range(steps + 1):
                            a = start + (end - start) * i / steps
                            px = x + r * math.cos(a)
                            py = y + r * math.sin(a)
                            pts.append((float(px), float(py)))
                            update_bounds(px, py)
                        self.polylines.append(pts)

                    elif t == "LWPOLYLINE":
                        pts = []
                        for p in e.get_points():
                            x, y = float(p[0]), float(p[1])
                            pts.append((x, y))
                            update_bounds(x, y)
                        if len(pts) >= 2:
                            self.polylines.append(pts)

                    elif t == "POLYLINE":
                        pts = []
                        for v in e.vertices:
                            x, y, _ = v.dxf.location
                            pts.append((float(x), float(y)))
                            update_bounds(x, y)
                        if len(pts) >= 2:
                            self.polylines.append(pts)

                    elif t == "SPLINE":
                        pts = []
                        try:
                            for p in e.flattening(distance=1.0):
                                x, y = float(p[0]), float(p[1])
                                pts.append((x, y))
                                update_bounds(x, y)
                        except Exception:
                            pass
                        if len(pts) >= 2:
                            self.polylines.append(pts)

                    elif t == "TEXT":
                        text = e.dxf.text
                        x, y, _ = e.dxf.insert
                        self.texts.append({"text": str(text), "conf": 1.0, "bbox": [(float(x), float(y))]})
                        update_bounds(x, y)

                    elif t == "MTEXT":
                        text = e.text
                        x, y, _ = e.dxf.insert
                        self.texts.append({"text": str(text), "conf": 1.0, "bbox": [(float(x), float(y))]})
                        update_bounds(x, y)

                    elif t == "INSERT":
                        x, y, _ = e.dxf.insert
                        name = e.dxf.name
                        self.symbols.append({
                            "bbox": (float(x), float(y), 10.0, 10.0),
                            "kind": f"BLOCK:{name}",
                            "points": [(float(x), float(y))]
                        })
                        update_bounds(x - 5, y - 5)
                        update_bounds(x + 5, y + 5)

                except Exception:
                    continue

            if min_x == 10**18:
                messagebox.showwarning("DXF読み込み", "DXF内に対応要素がありません。")
                return

            self.dxf_bounds = (min_x, min_y, max_x, max_y)
            self.original_image = self.render_dxf_preview(min_x, min_y, max_x, max_y)
            self.preview_image = self.original_image.copy()
            self.show_image(self.preview_image)

            msg = (
                f"DXF読み込み完了: 直線={len(self.lines)} / 円={len(self.circles)} / "
                f"ポリライン={len(self.polylines)} / 文字={len(self.texts)} / ブロック={len(self.symbols)}"
            )
            self.log(msg)
            self.status.set(msg)

        except Exception as e:
            messagebox.showerror("DXF読み込みエラー", str(e))

    def render_dxf_preview(self, min_x, min_y, max_x, max_y):
        span_x = max(max_x - min_x, 1.0)
        span_y = max(max_y - min_y, 1.0)

        canvas_w = 1200
        canvas_h = 700
        margin = 50

        scale = min((canvas_w - margin * 2) / span_x, (canvas_h - margin * 2) / span_y)
        scale = max(min(scale, 10.0), 0.01)

        img = np.ones((canvas_h, canvas_w, 3), dtype=np.uint8) * 255

        def to_img_point(x, y):
            px = int((float(x) - min_x) * scale + margin)
            py = int(canvas_h - ((float(y) - min_y) * scale + margin))
            return px, py

        for x1, y1, x2, y2 in self.lines:
            cv2.line(img, to_img_point(x1, y1), to_img_point(x2, y2), (0, 0, 0), 1)

        for x, y, r in self.circles:
            cv2.circle(img, to_img_point(x, y), max(1, int(float(r) * scale)), (255, 0, 0), 1)

        for pts in self.polylines:
            if len(pts) >= 2:
                arr = np.array([to_img_point(x, y) for x, y in pts], dtype=np.int32)
                cv2.polylines(img, [arr], False, (0, 128, 0), 1)

        for s in self.symbols:
            x, y, w, h = s["bbox"]
            p = to_img_point(x, y)
            cv2.circle(img, p, 5, (255, 0, 255), 1)
            cv2.putText(img, s["kind"][:12], p, cv2.FONT_HERSHEY_SIMPLEX, 0.4, (255, 0, 255), 1)

        for t in self.texts:
            text = t["text"]
            x, y = t["bbox"][0]
            cv2.putText(img, text[:20], to_img_point(x, y), cv2.FONT_HERSHEY_SIMPLEX, 0.5, (0, 0, 255), 1)

        return img

    def reset_vectors(self):
        self.lines = []
        self.circles = []
        self.rectangles = []
        self.polylines = []
        self.symbols = []
        self.texts = []
        self.completed_lines = []
        self.routes = []

    def preprocess(self):
        gray = cv2.cvtColor(self.original_image, cv2.COLOR_BGR2GRAY)
        _, binary = cv2.threshold(gray, int(self.threshold_var.get()), 255, cv2.THRESH_BINARY_INV)
        kernel = np.ones((2, 2), np.uint8)
        binary = cv2.morphologyEx(binary, cv2.MORPH_OPEN, kernel)
        return gray, binary

    def detect_vectors(self):
        if self.original_image is None:
            messagebox.showwarning("警告", "先に画像/PDF/DXFを開いてください。")
            return

        if self.input_mode == "dxf":
            messagebox.showinfo("DXF入力", "DXFは既にベクターデータとして読み込み済みです。AI解析やDXF出力へ進めます。")
            return

        self.reset_vectors()
        gray, binary = self.preprocess()
        result = self.original_image.copy()

        if self.detect_lines_var.get():
            self.detect_lines(binary, result)
        if self.detect_circles_var.get():
            self.detect_circles(gray, result)
        if self.detect_rects_var.get() or self.detect_curves_var.get() or self.detect_symbols_var.get():
            self.detect_contours(binary, result)

        self.preview_image = result
        self.show_image(result)

        msg = (
            f"検出完了: 直線={len(self.lines)} 円={len(self.circles)} "
            f"矩形={len(self.rectangles)} 曲線={len(self.polylines)} 記号候補={len(self.symbols)}"
        )
        self.log(msg)
        self.status.set(msg)

    def detect_lines(self, binary, result):
        lines = cv2.HoughLinesP(
            binary,
            rho=1,
            theta=np.pi / 180,
            threshold=50,
            minLineLength=int(self.min_line_var.get()),
            maxLineGap=15
        )
        if lines is None:
            return

        for line in lines:
            x1, y1, x2, y2 = line[0]
            self.lines.append((float(x1), float(y1), float(x2), float(y2)))
            cv2.line(result, (x1, y1), (x2, y2), (0, 0, 255), 2)

    def detect_circles(self, gray, result):
        blur = cv2.medianBlur(gray, 5)
        circles = cv2.HoughCircles(
            blur,
            cv2.HOUGH_GRADIENT,
            dp=1.2,
            minDist=20,
            param1=80,
            param2=25,
            minRadius=5,
            maxRadius=180
        )
        if circles is None:
            return

        circles = np.uint16(np.around(circles))
        for c in circles[0, :]:
            x, y, r = int(c[0]), int(c[1]), int(c[2])
            self.circles.append((float(x), float(y), float(r)))
            cv2.circle(result, (x, y), r, (255, 0, 0), 2)

    def detect_contours(self, binary, result):
        contours, _ = cv2.findContours(binary, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_NONE)

        for cnt in contours:
            area = cv2.contourArea(cnt)
            if area < 10:
                continue

            peri = cv2.arcLength(cnt, True)
            if peri < 10:
                continue

            approx = cv2.approxPolyDP(cnt, float(self.curve_epsilon_var.get()), True)
            points = [(float(p[0][0]), float(p[0][1])) for p in approx]
            x, y, w, h = cv2.boundingRect(cnt)

            if self.detect_symbols_var.get():
                if int(self.symbol_min_area_var.get()) <= area <= int(self.symbol_max_area_var.get()):
                    if 5 <= w <= 120 and 5 <= h <= 120:
                        kind = self.classify_symbol_candidate(points, x, y, w, h, area)
                        self.symbols.append({
                            "bbox": (float(x), float(y), float(w), float(h)),
                            "kind": kind,
                            "points": points
                        })
                        cv2.rectangle(result, (x, y), (x + w, y + h), (255, 0, 255), 2)
                        cv2.putText(result, kind[:8], (x, max(0, y - 3)),
                                    cv2.FONT_HERSHEY_SIMPLEX, 0.4, (255, 0, 255), 1)

            if len(points) == 4 and self.detect_rects_var.get():
                if w > 8 and h > 8:
                    self.rectangles.append((float(x), float(y), float(w), float(h)))
                    cv2.rectangle(result, (x, y), (x + w, y + h), (0, 255, 255), 2)
                    continue

            if self.detect_curves_var.get() and len(points) >= 2:
                self.polylines.append(points)
                for i in range(len(points) - 1):
                    p1 = (int(points[i][0]), int(points[i][1]))
                    p2 = (int(points[i + 1][0]), int(points[i + 1][1]))
                    cv2.line(result, p1, p2, (0, 255, 0), 1)

    def classify_symbol_candidate(self, points, x, y, w, h, area):
        aspect = w / max(h, 1)
        vertex_count = len(points)
        if 0.75 <= aspect <= 1.25 and vertex_count > 6:
            return "丸形記号候補"
        if vertex_count == 4:
            return "矩形器具候補"
        if aspect > 2.0:
            return "横長記号候補"
        if aspect < 0.5:
            return "縦長記号候補"
        return "電気記号候補"

    def auto_complete_shapes(self):
        if not self.lines:
            messagebox.showwarning("警告", "先にベクター検出またはDXF読み込みを実行してください。")
            return

        gap = int(self.complete_gap_var.get())
        result = self.preview_image.copy() if self.preview_image is not None else self.original_image.copy()

        endpoints = []
        for i, (x1, y1, x2, y2) in enumerate(self.lines):
            endpoints.append((x1, y1, i))
            endpoints.append((x2, y2, i))

        added = set()

        for i in range(len(endpoints)):
            x1, y1, line_i = endpoints[i]
            for j in range(i + 1, len(endpoints)):
                x2, y2, line_j = endpoints[j]
                if line_i == line_j:
                    continue

                dist = math.hypot(x1 - x2, y1 - y2)
                if 0 < dist <= gap:
                    key = tuple(sorted([(round(x1, 3), round(y1, 3)), (round(x2, 3), round(y2, 3))]))
                    if key not in added:
                        self.completed_lines.append((x1, y1, x2, y2))
                        added.add(key)

        if self.input_mode != "dxf":
            for x1, y1, x2, y2 in self.completed_lines:
                cv2.line(result, (int(x1), int(y1)), (int(x2), int(y2)), (0, 128, 255), 2)
            self.preview_image = result
            self.show_image(result)

        self.log(f"欠損補完完了: 補完線={len(self.completed_lines)}")
        self.status.set(f"欠損補完完了: {len(self.completed_lines)}本")

    def analyze_structure(self):
        if not self.lines:
            messagebox.showwarning("警告", "先にベクター検出またはDXF読み込みを実行してください。")
            return

        result = self.preview_image.copy() if self.preview_image is not None else self.original_image.copy()

        centers = []
        for x, y, r in self.circles:
            centers.append((x, y, "円記号"))

        for s in self.symbols:
            x, y, w, h = s["bbox"]
            centers.append((x + w / 2, y + h / 2, s["kind"]))

        route_count = 0

        for cx, cy, kind in centers:
            nearest = None
            nearest_dist = 999999

            for x1, y1, x2, y2 in self.lines:
                d = self.point_to_segment_distance(cx, cy, x1, y1, x2, y2)
                if d < nearest_dist:
                    nearest_dist = d
                    nearest = (x1, y1, x2, y2)

            if nearest and nearest_dist < 80:
                x1, y1, x2, y2 = nearest
                mx = (x1 + x2) / 2
                my = (y1 + y2) / 2
                self.routes.append((cx, cy, mx, my, kind))
                route_count += 1

                if self.input_mode != "dxf":
                    cv2.line(result, (int(cx), int(cy)), (int(mx), int(my)), (128, 0, 255), 2)

        if self.input_mode != "dxf":
            self.preview_image = result
            self.show_image(result)

        self.log(f"構造解析完了: 配線ルート推定={route_count}")
        self.status.set(f"構造解析完了: 配線ルート推定 {route_count}")

    def point_to_segment_distance(self, px, py, x1, y1, x2, y2):
        vx = x2 - x1
        vy = y2 - y1
        wx = px - x1
        wy = py - y1

        c1 = vx * wx + vy * wy
        if c1 <= 0:
            return math.hypot(px - x1, py - y1)

        c2 = vx * vx + vy * vy
        if c2 <= c1:
            return math.hypot(px - x2, py - y2)

        b = c1 / c2
        bx = x1 + b * vx
        by = y1 + b * vy
        return math.hypot(px - bx, py - by)

    def run_ocr(self):
        if self.original_image is None:
            messagebox.showwarning("警告", "先に画像/PDF/DXFを開いてください。")
            return

        if self.input_mode == "dxf":
            messagebox.showinfo("OCR不要", "DXF内のTEXT/MTEXTは読み込み時に取得済みです。")
            return

        if not EASYOCR_AVAILABLE:
            messagebox.showwarning(
                "OCR未導入",
                "easyocr が入っていません。\n\nOCRなしでもDXF変換は可能です。\n\n"
                "入れる場合は、Spyderを終了してからコマンドプロンプトで実行してください。"
            )
            return

        try:
            if self.ocr_reader is None:
                self.log("OCR初期化中...")
                self.ocr_reader = easyocr.Reader(["ja", "en"], gpu=False)

            img_rgb = cv2.cvtColor(self.original_image, cv2.COLOR_BGR2RGB)
            results = self.ocr_reader.readtext(img_rgb)

            result_img = self.preview_image.copy() if self.preview_image is not None else self.original_image.copy()
            self.texts = []

            for bbox, text, conf in results:
                if conf < 0.25:
                    continue

                pts = [(float(p[0]), float(p[1])) for p in bbox]
                x = min(p[0] for p in pts)
                y = min(p[1] for p in pts)

                self.texts.append({"text": text, "conf": float(conf), "bbox": pts})

                pts_i = np.array([(int(a), int(b)) for a, b in pts], dtype=np.int32)
                cv2.polylines(result_img, [pts_i], True, (0, 128, 255), 2)
                cv2.putText(result_img, text[:10], (int(x), int(y)),
                            cv2.FONT_HERSHEY_SIMPLEX, 0.5, (0, 128, 255), 1)

            self.preview_image = result_img
            self.show_image(result_img)

            self.log(f"OCR完了: {len(self.texts)}件")
            self.status.set(f"OCR完了: {len(self.texts)}件")

        except Exception as e:
            messagebox.showerror("OCRエラー", str(e))

    def export_dxf(self):
        if self.original_image is None:
            messagebox.showwarning("警告", "先に画像/PDF/DXFを開いてください。")
            return

        save_path = filedialog.asksaveasfilename(
            defaultextension=".dxf",
            filetypes=[("DXF files", "*.dxf")]
        )
        if not save_path:
            return

        try:
            doc = ezdxf.new("R2010")
            msp = doc.modelspace()

            for var in self.layer_vars.values():
                name = var.get()
                if name and name not in doc.layers:
                    doc.layers.new(name)

            if self.input_mode == "dxf" and self.dxf_bounds:
                min_x, min_y, max_x, max_y = self.dxf_bounds
                margin = 10.0

                def cad_point(x, y):
                    return (float(x), float(y))

                border = [
                    (min_x - margin, min_y - margin),
                    (max_x + margin, min_y - margin),
                    (max_x + margin, max_y + margin),
                    (min_x - margin, max_y + margin),
                    (min_x - margin, min_y - margin),
                ]

            else:
                img_h, img_w = self.original_image.shape[:2]
                scale = float(self.dxf_scale_var.get())
                margin = 10.0

                def cad_point(x, y):
                    return (
                        margin + float(x) * scale,
                        margin + float(img_h - y) * scale
                    )

                border = [
                    cad_point(0, 0),
                    cad_point(img_w, 0),
                    cad_point(img_w, img_h),
                    cad_point(0, img_h),
                    cad_point(0, 0)
                ]

            msp.add_lwpolyline(border, dxfattribs={"layer": self.layer_vars["border"].get()})

            for x1, y1, x2, y2 in self.lines:
                msp.add_line(cad_point(x1, y1), cad_point(x2, y2),
                             dxfattribs={"layer": self.layer_vars["lines"].get()})

            for x1, y1, x2, y2 in self.completed_lines:
                msp.add_line(cad_point(x1, y1), cad_point(x2, y2),
                             dxfattribs={"layer": self.layer_vars["completed"].get()})

            for x, y, r in self.circles:
                radius = float(r) if self.input_mode == "dxf" else float(r) * float(self.dxf_scale_var.get())
                msp.add_circle(center=cad_point(x, y), radius=radius,
                               dxfattribs={"layer": self.layer_vars["circles"].get()})

            for x, y, w, h in self.rectangles:
                pts = [
                    cad_point(x, y),
                    cad_point(x + w, y),
                    cad_point(x + w, y + h),
                    cad_point(x, y + h),
                    cad_point(x, y)
                ]
                msp.add_lwpolyline(pts, dxfattribs={"layer": self.layer_vars["rectangles"].get()})

            for points in self.polylines:
                if len(points) >= 2:
                    msp.add_lwpolyline([cad_point(x, y) for x, y in points],
                                       dxfattribs={"layer": self.layer_vars["curves"].get()})

            for s in self.symbols:
                x, y, w, h = s["bbox"]
                if self.input_mode == "dxf":
                    msp.add_circle(cad_point(x, y), radius=2.0,
                                   dxfattribs={"layer": self.layer_vars["symbols"].get()})
                    text_point = cad_point(x, y)
                else:
                    pts = [
                        cad_point(x, y),
                        cad_point(x + w, y),
                        cad_point(x + w, y + h),
                        cad_point(x, y + h),
                        cad_point(x, y)
                    ]
                    msp.add_lwpolyline(pts, dxfattribs={"layer": self.layer_vars["symbols"].get()})
                    text_point = cad_point(x, y - 8)

                msp.add_text(
                    s["kind"],
                    dxfattribs={"height": 2.5, "layer": self.layer_vars["symbols"].get()}
                ).set_placement(text_point)

            for t in self.texts:
                pts = t["bbox"]
                x = min(p[0] for p in pts)
                y = min(p[1] for p in pts)
                msp.add_text(
                    t["text"],
                    dxfattribs={"height": 3.0, "layer": self.layer_vars["texts"].get()}
                ).set_placement(cad_point(x, y))

            for cx, cy, mx, my, kind in self.routes:
                msp.add_line(cad_point(cx, cy), cad_point(mx, my),
                             dxfattribs={"layer": self.layer_vars["routes"].get()})

            doc.saveas(save_path)

            self.log(f"DXF出力完了: {save_path}")
            messagebox.showinfo("DXF出力完了", f"保存しました。\n\n{save_path}")

        except Exception as e:
            messagebox.showerror("DXF出力エラー", str(e))

    def select_ollama_exe(self):
        path = filedialog.askopenfilename(
            filetypes=[("ollama.exe", "ollama.exe"), ("EXE", "*.exe"), ("All", "*.*")]
        )
        if path:
            self.ollama_path_var.set(path)

    def start_ollama(self):
        exe = self.ollama_path_var.get().strip()
        if not os.path.exists(exe):
            messagebox.showerror("Ollama起動エラー", f"見つかりません:\n{exe}")
            return

        try:
            subprocess.Popen([exe, "serve"], creationflags=subprocess.CREATE_NEW_CONSOLE)
            time.sleep(1)
            self.log("Ollama serve 起動要求を送信しました。")
            self.check_ollama(show_popup=False)
        except Exception as e:
            messagebox.showerror("Ollama起動エラー", str(e))

    def check_ollama(self, show_popup=True):
        try:
            r = requests.get("http://127.0.0.1:11434/api/tags", timeout=5)
            r.raise_for_status()
            self.log("Ollama接続成功")
            if show_popup:
                messagebox.showinfo("Ollama", "接続成功")
            return True
        except Exception as e:
            self.log(f"Ollama接続失敗: {e}")
            if show_popup:
                messagebox.showerror("Ollama接続失敗", str(e))
            return False

    def load_ollama_models(self):
        if not self.check_ollama(show_popup=False):
            return

        try:
            r = requests.get("http://127.0.0.1:11434/api/tags", timeout=10)
            models = [m["name"] for m in r.json().get("models", [])]
            self.model_combo["values"] = models
            if models:
                self.model_var.set(models[0])
            self.log("モデル一覧: " + ", ".join(models))
        except Exception as e:
            messagebox.showerror("モデル取得エラー", str(e))

    def ask_ollama(self, prompt):
        model = self.model_var.get().strip()
        if not model:
            raise RuntimeError("モデルが選択されていません。")

        payload = {
            "model": model,
            "prompt": prompt,
            "stream": False,
            "options": {"temperature": 0.2, "num_ctx": 4096}
        }

        r = requests.post("http://127.0.0.1:11434/api/generate", json=payload, timeout=180)
        r.raise_for_status()
        return r.json().get("response", "")

    def ai_analyze(self):
        try:
            prompt = f"""
あなたは電気設備図面のCAD変換補助AIです。

入力形式: {self.input_mode}
検出/読取結果:
直線 {len(self.lines)}
補完線 {len(self.completed_lines)}
円 {len(self.circles)}
矩形 {len(self.rectangles)}
曲線/ポリライン {len(self.polylines)}
記号/ブロック候補 {len(self.symbols)}
文字 {len(self.texts)}
配線ルート推定 {len(self.routes)}

現在設定:
二値化 {self.threshold_var.get()}
最小線長 {self.min_line_var.get()}
曲線近似 {self.curve_epsilon_var.get()}
補完距離 {self.complete_gap_var.get()}
PDF倍率 {self.pdf_scale_var.get()}
DXF縮尺 {self.dxf_scale_var.get()}

次を日本語で簡潔に提案してください。
1. 電気設備図面としての妥当性
2. レイヤ整理案
3. DXF入力の場合の再レイヤ分け案
4. ラスター入力の場合の検出改善案
5. Jw_cadで扱いやすくする注意点
"""
            self.log("AI解析中...")
            ans = self.ask_ollama(prompt)
            self.log_text.delete("1.0", tk.END)
            self.log_text.insert(tk.END, ans)
        except Exception as e:
            messagebox.showerror("AI解析エラー", str(e))

    def ai_layer_suggest(self):
        try:
            prompt = """
電気設備図面をJw_cad/AutoCAD向けDXFに変換します。
以下のレイヤについて、実務で分かりやすい日本語レイヤ名を提案してください。

border, lines, completed, circles, rectangles, curves, symbols, texts, routes, dxf_import

出力は「キー: レイヤ名」の形式だけにしてください。
"""
            ans = self.ask_ollama(prompt)
            self.log_text.delete("1.0", tk.END)
            self.log_text.insert(tk.END, ans)
        except Exception as e:
            messagebox.showerror("AIレイヤ提案エラー", str(e))

    def show_image(self, image):
        if image is None:
            return

        rgb = cv2.cvtColor(image, cv2.COLOR_BGR2RGB)
        pil_img = Image.fromarray(rgb)

        cw = self.canvas.winfo_width()
        ch = self.canvas.winfo_height()

        if cw < 20:
            cw = 1300
        if ch < 20:
            ch = 650

        pil_img.thumbnail((cw - 20, ch - 20))
        self.tk_img = ImageTk.PhotoImage(pil_img)

        self.canvas.delete("all")
        self.canvas.create_image(10, 10, anchor=tk.NW, image=self.tk_img)

    def reset(self):
        if self.original_image is None:
            return

        self.reset_vectors()
        self.preview_image = self.original_image.copy()
        self.show_image(self.original_image)
        self.log("リセットしました。")

    def log(self, text):
        self.log_text.insert(tk.END, str(text) + "\n")
        self.log_text.see(tk.END)


if __name__ == "__main__":
    root = tk.Tk()
    app = RasterElectricalDXFApp(root)
    root.mainloop()