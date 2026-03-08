#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Python ライブラリ掃除GUI
- 失敗した可能性のあるインストール残骸を探索
- 壊れた dist-info / egg-info を検出
- pip キャッシュ、temp/build 残骸を探索
- 選択項目を隔離(推奨)または削除

対応:
- Windows
- Linux
- macOS
- Android / Pydroid3

注意:
このツールは「失敗したインストール」を100%厳密に判定するものではなく、
ヒューリスティック(経験則)で候補を洗い出します。
最初は必ず「隔離（推奨）」で試してください。
"""

from __future__ import annotations

import csv
import os
import platform
import queue
import re
import shutil
import site
import subprocess
import sys
import tempfile
import threading
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from tkinter import (
    Tk, StringVar, BooleanVar, IntVar, END,
    BOTH, LEFT, RIGHT, X, Y, filedialog, messagebox
)
import tkinter as tk
from tkinter import ttk

APP_NAME = "Pythonライブラリ掃除GUI"
APP_VERSION = "1.2.0"

SCAN_PATTERNS = [
    re.compile(r"^pip-.*", re.IGNORECASE),
    re.compile(r"^build$", re.IGNORECASE),
    re.compile(r"^build-.*", re.IGNORECASE),
    re.compile(r"^tmp.*", re.IGNORECASE),
    re.compile(r"^easy_install-.*", re.IGNORECASE),
    re.compile(r"^pybuild.*", re.IGNORECASE),
    re.compile(r"^pip-unpack-.*", re.IGNORECASE),
    re.compile(r"^pip-modern-metadata-.*", re.IGNORECASE),
    re.compile(r"^pip-req-build-.*", re.IGNORECASE),
    re.compile(r"^pip-install-.*", re.IGNORECASE),
]

ARCHIVE_EXTS = {".whl", ".zip", ".tar", ".gz", ".bz2", ".xz", ".tgz"}
QUARANTINE_ROOT = Path.home() / ".python_lib_cleanup_quarantine"
DEFAULT_EXPORT = Path.home() / "python_library_cleanup_scan_results.csv"


def safe_size(path: Path) -> int:
    try:
        if path.is_file() or path.is_symlink():
            return path.stat().st_size
        total = 0
        for root, _, files in os.walk(path, onerror=lambda e: None):
            for name in files:
                p = Path(root) / name
                try:
                    total += p.stat().st_size
                except Exception:
                    pass
        return total
    except Exception:
        return 0


def format_size(num: int) -> str:
    units = ["B", "KB", "MB", "GB", "TB"]
    size = float(num)
    for unit in units:
        if size < 1024 or unit == units[-1]:
            return f"{size:.2f} {unit}"
        size /= 1024.0
    return f"{num} B"


def get_mtime_str(path: Path) -> str:
    try:
        return datetime.fromtimestamp(path.stat().st_mtime).strftime("%Y-%m-%d %H:%M:%S")
    except Exception:
        return "-"


def normalize_name(name: str) -> str:
    return re.sub(r"[-_.]+", "-", name.strip().lower())


def extract_package_name_from_dist_info(name: str) -> str:
    if name.endswith(".dist-info"):
        core = name[:-10]
    elif name.endswith(".egg-info"):
        core = name[:-9]
    else:
        core = name
    m = re.match(r"^(.*?)-\d", core)
    if m:
        return m.group(1)
    return core


@dataclass
class ScanItem:
    category: str
    severity: str
    name: str
    path: str
    size_bytes: int
    modified: str
    reason: str
    recommended_action: str


class LibraryScanner:
    def __init__(self, log_callback=None, stop_flag: threading.Event | None = None):
        self.log_callback = log_callback or (lambda msg: None)
        self.stop_flag = stop_flag or threading.Event()
        self.items: list[ScanItem] = []

    def log(self, msg: str):
        self.log_callback(msg)

    def stopped(self) -> bool:
        return self.stop_flag.is_set()

    def add_item(self, **kwargs):
        self.items.append(ScanItem(**kwargs))

    def run(self, include_cache=True, include_temp=True, include_site=True):
        self.items.clear()
        if include_site:
            self.scan_site_packages()
        if include_cache:
            self.scan_pip_cache()
        if include_temp:
            self.scan_temp_dirs()
        return self.items

    def get_site_package_paths(self) -> list[Path]:
        paths = set()
        candidates = []
        try:
            candidates.extend(site.getsitepackages())
        except Exception:
            pass
        try:
            usp = site.getusersitepackages()
            if usp:
                candidates.append(usp)
        except Exception:
            pass
        for p in candidates:
            if p and os.path.isdir(p):
                paths.add(Path(p))
        return sorted(paths)

    def scan_site_packages(self):
        self.log("site-packages を検査中...")
        for sp in self.get_site_package_paths():
            if self.stopped():
                return
            self.log(f"  検査: {sp}")
            self.scan_one_site_packages(sp)

    def scan_one_site_packages(self, sp: Path):
        try:
            entries = list(sp.iterdir())
        except Exception as e:
            self.log(f"    読み込み失敗: {e}")
            return

        dist_infos = {}
        egg_infos = {}
        package_dirs = set()
        py_modules = set()

        for entry in entries:
            name = entry.name
            if name.endswith(".dist-info"):
                base = normalize_name(name[:-10])
                dist_infos[base] = entry
            elif name.endswith(".egg-info"):
                base = normalize_name(name[:-9])
                egg_infos[base] = entry
            elif entry.is_dir() and not name.startswith("__pycache__"):
                package_dirs.add(name)
            elif entry.is_file() and name.endswith(".py"):
                py_modules.add(name[:-3])

        for base, dist_path in dist_infos.items():
            if self.stopped():
                return
            self.inspect_dist_info(sp, base, dist_path, package_dirs, py_modules)

        for base, egg_path in egg_infos.items():
            if self.stopped():
                return
            self.inspect_egg_info(sp, base, egg_path, package_dirs, py_modules)

    def inspect_dist_info(self, sp: Path, base: str, dist_path: Path, package_dirs: set[str], py_modules: set[str]):
        metadata = dist_path / "METADATA"
        record = dist_path / "RECORD"
        top_level = dist_path / "top_level.txt"
        installer = dist_path / "INSTALLER"

        missing = []
        if not metadata.exists():
            missing.append("METADATA")
        if not record.exists():
            missing.append("RECORD")

        if missing:
            self.add_item(
                category="壊れたdist-info",
                severity="高",
                name=dist_path.name,
                path=str(dist_path),
                size_bytes=safe_size(dist_path),
                modified=get_mtime_str(dist_path),
                reason=f"dist-info に必須ファイル欠落: {', '.join(missing)}",
                recommended_action="隔離後に確認",
            )

        top_names = []
        if top_level.exists():
            try:
                txt = top_level.read_text(encoding="utf-8", errors="ignore")
                top_names = [line.strip() for line in txt.splitlines() if line.strip()]
            except Exception:
                pass

        if top_names:
            missing_targets = []
            for n in top_names:
                if n not in package_dirs and n not in py_modules:
                    missing_targets.append(n)
            if missing_targets:
                self.add_item(
                    category="孤立dist-info",
                    severity="中",
                    name=dist_path.name,
                    path=str(dist_path),
                    size_bytes=safe_size(dist_path),
                    modified=get_mtime_str(dist_path),
                    reason=f"top_level.txt に対応する本体が見当たらない: {', '.join(missing_targets)}",
                    recommended_action="隔離後に確認",
                )
        else:
            if not installer.exists() and not record.exists():
                self.add_item(
                    category="疑わしいdist-info",
                    severity="低",
                    name=dist_path.name,
                    path=str(dist_path),
                    size_bytes=safe_size(dist_path),
                    modified=get_mtime_str(dist_path),
                    reason="top_level.txt / INSTALLER / RECORD の情報不足",
                    recommended_action="確認のみ",
                )

    def inspect_egg_info(self, sp: Path, base: str, egg_path: Path, package_dirs: set[str], py_modules: set[str]):
        pkg_info = egg_path / "PKG-INFO" if egg_path.is_dir() else egg_path
        if not pkg_info.exists():
            self.add_item(
                category="壊れたegg-info",
                severity="中",
                name=egg_path.name,
                path=str(egg_path),
                size_bytes=safe_size(egg_path),
                modified=get_mtime_str(egg_path),
                reason="PKG-INFO が見当たらない",
                recommended_action="隔離後に確認",
            )

    def get_pip_cache_dir(self) -> Path | None:
        env = os.environ.get("PIP_CACHE_DIR")
        if env:
            p = Path(env)
            if p.exists():
                return p
        if platform.system() == "Windows":
            local = os.environ.get("LOCALAPPDATA")
            if local:
                p = Path(local) / "pip" / "Cache"
                if p.exists():
                    return p
        else:
            p = Path.home() / ".cache" / "pip"
            if p.exists():
                return p
        return None

    def scan_pip_cache(self):
        cache_dir = self.get_pip_cache_dir()
        if not cache_dir:
            self.log("pip キャッシュは見つかりませんでした。")
            return

        self.log(f"pip キャッシュを検査中: {cache_dir}")

        total_cache = safe_size(cache_dir)
        self.add_item(
            category="pipキャッシュ",
            severity="低",
            name="pip cache total",
            path=str(cache_dir),
            size_bytes=total_cache,
            modified=get_mtime_str(cache_dir),
            reason="pip がダウンロードしたキャッシュ全体。失敗時の残骸を含む可能性あり",
            recommended_action="不要なら削除可",
        )

        for root, _, files in os.walk(cache_dir, onerror=lambda e: None):
            if self.stopped():
                return
            root_p = Path(root)
            for name in files:
                p = root_p / name
                if p.suffix.lower() in ARCHIVE_EXTS:
                    size = safe_size(p)
                    if size >= 5 * 1024 * 1024:
                        self.add_item(
                            category="大きいキャッシュファイル",
                            severity="低",
                            name=p.name,
                            path=str(p),
                            size_bytes=size,
                            modified=get_mtime_str(p),
                            reason="大きな wheel / archive キャッシュ",
                            recommended_action="不要なら削除可",
                        )

    def scan_temp_dirs(self):
        temp_candidates = [Path(tempfile.gettempdir())]
        if platform.system() == "Windows":
            for key in ("TEMP", "TMP"):
                val = os.environ.get(key)
                if val:
                    temp_candidates.append(Path(val))

        self.log("一時フォルダを検査中...")

        visited = set()
        for temp_dir in temp_candidates:
            temp_dir = temp_dir.resolve()
            if temp_dir in visited or not temp_dir.exists():
                continue
            visited.add(temp_dir)
            self.log(f"  検査: {temp_dir}")
            try:
                for child in temp_dir.iterdir():
                    if self.stopped():
                        return
                    name = child.name
                    matched = any(pat.match(name) for pat in SCAN_PATTERNS)
                    if not matched:
                        continue
                    self.add_item(
                        category="temp/build残骸",
                        severity="中",
                        name=name,
                        path=str(child),
                        size_bytes=safe_size(child),
                        modified=get_mtime_str(child),
                        reason="pip/build/easy_install 系の一時フォルダまたは残骸の可能性",
                        recommended_action="隔離後に確認",
                    )
            except Exception as e:
                self.log(f"    読み込み失敗: {e}")


class CleanupEngine:
    def __init__(self, log_callback=None):
        self.log_callback = log_callback or (lambda msg: None)

    def log(self, msg: str):
        self.log_callback(msg)

    def quarantine_item(self, path: Path) -> tuple[bool, str]:
        try:
            if not path.exists():
                return False, "既に存在しません"
            QUARANTINE_ROOT.mkdir(parents=True, exist_ok=True)
            stamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
            dest = QUARANTINE_ROOT / f"{stamp}__{path.name}"
            shutil.move(str(path), str(dest))
            return True, f"隔離先: {dest}"
        except Exception as e:
            return False, str(e)

    def delete_item(self, path: Path) -> tuple[bool, str]:
        try:
            if not path.exists():
                return False, "既に存在しません"
            if path.is_dir() and not path.is_symlink():
                shutil.rmtree(path)
            else:
                path.unlink()
            return True, "削除しました"
        except Exception as e:
            return False, str(e)

    def pip_uninstall(self, package_name: str) -> tuple[bool, str]:
        try:
            cmd = [sys.executable, "-m", "pip", "uninstall", "-y", package_name]
            proc = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
            out = (proc.stdout or "") + "\n" + (proc.stderr or "")
            ok = proc.returncode == 0
            return ok, out.strip()
        except Exception as e:
            return False, str(e)


class App(Tk):
    def __init__(self):
        super().__init__()
        self.title(f"{APP_NAME} v{APP_VERSION}")
        self.geometry("980x720")
        self.minsize(760, 560)

        self.log_queue = queue.Queue()
        self.stop_event = threading.Event()
        self.scan_items: list[ScanItem] = []

        self.var_include_site = BooleanVar(value=True)
        self.var_include_cache = BooleanVar(value=True)
        self.var_include_temp = BooleanVar(value=True)
        self.var_action_mode = StringVar(value="quarantine")
        self.var_filter = StringVar(value="")
        self.var_only_high = BooleanVar(value=False)
        self.var_only_selected_count = IntVar(value=0)

        self.build_ui()
        self.bind_all("<MouseWheel>", self.on_mousewheel)
        self.bind_all("<Button-4>", self.on_mousewheel)
        self.bind_all("<Button-5>", self.on_mousewheel)
        self.after(150, self.process_log_queue)
        self.append_log("準備完了。まずは「スキャン開始」を押してください。")

    def build_ui(self):
        self.outer_canvas = tk.Canvas(self, highlightthickness=0)
        self.outer_vsb = ttk.Scrollbar(self, orient="vertical", command=self.outer_canvas.yview)
        self.outer_hsb = ttk.Scrollbar(self, orient="horizontal", command=self.outer_canvas.xview)
        self.outer_canvas.configure(
            yscrollcommand=self.outer_vsb.set,
            xscrollcommand=self.outer_hsb.set
        )

        self.outer_canvas.pack(side=LEFT, fill=BOTH, expand=True)
        self.outer_vsb.pack(side=RIGHT, fill=Y)
        self.outer_hsb.pack(side="bottom", fill=X)

        self.main_frame = ttk.Frame(self.outer_canvas, padding=0)
        self.canvas_window = self.outer_canvas.create_window((0, 0), window=self.main_frame, anchor="nw")

        self.main_frame.bind("<Configure>", self.on_main_frame_configure)
        self.outer_canvas.bind("<Configure>", self.on_outer_canvas_configure)

        top = ttk.Frame(self.main_frame, padding=8)
        top.pack(fill=X)

        scan_box = ttk.LabelFrame(top, text="スキャン対象", padding=8)
        scan_box.pack(fill=X, padx=(0, 0), pady=(0, 6))

        ttk.Checkbutton(scan_box, text="site-packages を検査", variable=self.var_include_site).grid(row=0, column=0, sticky="w", padx=4, pady=2)
        ttk.Checkbutton(scan_box, text="pip キャッシュを検査", variable=self.var_include_cache).grid(row=1, column=0, sticky="w", padx=4, pady=2)
        ttk.Checkbutton(scan_box, text="temp/build 残骸を検査", variable=self.var_include_temp).grid(row=2, column=0, sticky="w", padx=4, pady=2)

        action_box = ttk.LabelFrame(top, text="処理モード", padding=8)
        action_box.pack(fill=X, pady=(0, 6))
        ttk.Radiobutton(action_box, text="隔離（推奨）", value="quarantine", variable=self.var_action_mode).grid(row=0, column=0, sticky="w", padx=4, pady=2)
        ttk.Radiobutton(action_box, text="完全削除", value="delete", variable=self.var_action_mode).grid(row=1, column=0, sticky="w", padx=4, pady=2)
        ttk.Radiobutton(action_box, text="pip uninstall（dist-info名から推定）", value="pip", variable=self.var_action_mode).grid(row=2, column=0, sticky="w", padx=4, pady=2)

        button_box = ttk.LabelFrame(top, text="操作", padding=8)
        button_box.pack(fill=X, pady=(0, 6))

        row1 = ttk.Frame(button_box)
        row1.pack(fill=X, pady=2)
        ttk.Button(row1, text="スキャン開始", command=self.start_scan).pack(side=LEFT, padx=4, pady=2)
        ttk.Button(row1, text="停止", command=self.stop_scan).pack(side=LEFT, padx=4, pady=2)
        ttk.Button(row1, text="再読込", command=self.refresh_view).pack(side=LEFT, padx=4, pady=2)

        row2 = ttk.Frame(button_box)
        row2.pack(fill=X, pady=2)
        ttk.Button(row2, text="選択項目を実行", command=self.execute_selected).pack(side=LEFT, padx=4, pady=2)
        ttk.Button(row2, text="CSV出力", command=self.export_csv).pack(side=LEFT, padx=4, pady=2)

        filter_frame = ttk.Frame(self.main_frame, padding=(8, 0, 8, 4))
        filter_frame.pack(fill=X)
        ttk.Label(filter_frame, text="フィルタ:").pack(side=LEFT)
        ent = ttk.Entry(filter_frame, textvariable=self.var_filter, width=30)
        ent.pack(side=LEFT, padx=4)
        ent.bind("<KeyRelease>", lambda e: self.refresh_view())
        ttk.Checkbutton(
            filter_frame,
            text="重大度: 高のみ",
            variable=self.var_only_high,
            command=self.refresh_view
        ).pack(side=LEFT, padx=10)
        self.status_label = ttk.Label(filter_frame, text="0件")
        self.status_label.pack(side=RIGHT)

        middle = ttk.PanedWindow(self.main_frame, orient=tk.VERTICAL)
        middle.pack(fill=BOTH, expand=True, padx=8, pady=4)

        frame_table = ttk.LabelFrame(middle, text="検出結果", padding=6)
        middle.add(frame_table, weight=4)

        columns = ("selected", "category", "severity", "name", "size", "modified", "reason", "path", "action")
        self.tree = ttk.Treeview(frame_table, columns=columns, show="headings", selectmode="extended")

        self.tree_vsb = ttk.Scrollbar(frame_table, orient="vertical", command=self.tree.yview)
        self.tree_hsb = ttk.Scrollbar(frame_table, orient="horizontal", command=self.tree.xview)
        self.tree.configure(
            yscrollcommand=self.tree_vsb.set,
            xscrollcommand=self.tree_hsb.set
        )

        headings = {
            "selected": "選択",
            "category": "カテゴリ",
            "severity": "重大度",
            "name": "名前",
            "size": "サイズ",
            "modified": "更新日時",
            "reason": "理由",
            "path": "パス",
            "action": "推奨処理",
        }
        widths = {
            "selected": 60,
            "category": 110,
            "severity": 70,
            "name": 180,
            "size": 90,
            "modified": 140,
            "reason": 260,
            "path": 360,
            "action": 120,
        }

        for col in columns:
            self.tree.heading(col, text=headings[col])
            self.tree.column(col, width=widths[col], anchor="w", stretch=False)

        self.tree.pack(side=LEFT, fill=BOTH, expand=True)
        self.tree_vsb.pack(side=RIGHT, fill=Y)
        self.tree_hsb.pack(side="bottom", fill=X)

        self.tree.bind("<Double-1>", self.on_double_click)
        self.tree.tag_configure("high", background="#ffe0e0")
        self.tree.tag_configure("mid", background="#fff4d6")
        self.tree.tag_configure("low", background="#eaf6ff")

        table_btns = ttk.Frame(frame_table)
        table_btns.pack(fill=X, pady=(6, 0))
        ttk.Button(table_btns, text="全選択", command=self.select_all).pack(side=LEFT, padx=2)
        ttk.Button(table_btns, text="全解除", command=self.clear_selection_marks).pack(side=LEFT, padx=2)
        ttk.Button(table_btns, text="高のみ選択", command=lambda: self.auto_select_by_severity("高")).pack(side=LEFT, padx=2)
        ttk.Button(table_btns, text="中以上選択", command=lambda: self.auto_select_by_severity("中")).pack(side=LEFT, padx=2)
        ttk.Button(table_btns, text="存在しない項目を除外", command=self.remove_nonexistent_from_list).pack(side=LEFT, padx=2)

        frame_log = ttk.LabelFrame(middle, text="ログ", padding=6)
        middle.add(frame_log, weight=2)
        self.log_text = tk.Text(frame_log, height=12, wrap="word")
        log_vsb = ttk.Scrollbar(frame_log, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_vsb.set)
        self.log_text.pack(side=LEFT, fill=BOTH, expand=True)
        log_vsb.pack(side=RIGHT, fill=Y)

    def on_main_frame_configure(self, event=None):
        try:
            self.outer_canvas.configure(scrollregion=self.outer_canvas.bbox("all"))
        except Exception:
            pass

    def on_outer_canvas_configure(self, event):
        try:
            req_w = self.main_frame.winfo_reqwidth()
            canvas_w = event.width
            self.outer_canvas.itemconfigure(self.canvas_window, width=req_w)
            if req_w > canvas_w:
                self.outer_hsb.pack(side="bottom", fill=X)
            else:
                self.outer_hsb.pack(side="bottom", fill=X)
        except Exception:
            pass

    def on_mousewheel(self, event):
        try:
            widget = self.focus_get()
        except Exception:
            widget = None

        target = self.log_text if widget is self.log_text else self.outer_canvas

        try:
            if hasattr(event, "delta") and event.delta:
                direction = -1 if event.delta > 0 else 1
                target.yview_scroll(direction, "units")
            elif getattr(event, "num", None) == 4:
                target.yview_scroll(-1, "units")
            elif getattr(event, "num", None) == 5:
                target.yview_scroll(1, "units")
        except Exception:
            pass

    def append_log(self, msg: str):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(END, f"[{timestamp}] {msg}\n")
        self.log_text.see(END)

    def process_log_queue(self):
        while True:
            try:
                msg = self.log_queue.get_nowait()
            except queue.Empty:
                break
            self.append_log(msg)
        self.after(150, self.process_log_queue)

    def log(self, msg: str):
        self.log_queue.put(msg)

    def start_scan(self):
        self.stop_event.clear()
        self.scan_items.clear()
        self.tree.delete(*self.tree.get_children())
        self.log("スキャン開始")
        t = threading.Thread(target=self.scan_worker, daemon=True)
        t.start()

    def stop_scan(self):
        self.stop_event.set()
        self.log("停止要求を出しました。")

    def scan_worker(self):
        scanner = LibraryScanner(log_callback=self.log, stop_flag=self.stop_event)
        try:
            items = scanner.run(
                include_cache=self.var_include_cache.get(),
                include_temp=self.var_include_temp.get(),
                include_site=self.var_include_site.get(),
            )
            self.scan_items = items
            self.log(f"スキャン完了: {len(items)} 件")
            self.after(0, self.refresh_view)
        except Exception as e:
            self.log(f"スキャン中にエラー: {e}")

    def filtered_items(self) -> list[ScanItem]:
        text = self.var_filter.get().strip().lower()
        only_high = self.var_only_high.get()
        result = []
        for item in self.scan_items:
            if only_high and item.severity != "高":
                continue
            hay = " ".join([
                item.category, item.severity, item.name,
                item.path, item.reason, item.recommended_action
            ]).lower()
            if text and text not in hay:
                continue
            result.append(item)
        return result

    def refresh_view(self):
        current_states = {}
        for iid in self.tree.get_children():
            vals = self.tree.item(iid, "values")
            if vals:
                current_states[vals[7]] = vals[0]

        self.tree.delete(*self.tree.get_children())
        count = 0
        selected_count = 0
        total_size = 0

        for item in self.filtered_items():
            mark = current_states.get(item.path, "□")
            if mark == "■":
                selected_count += 1
            total_size += item.size_bytes

            tag = "low"
            if item.severity == "高":
                tag = "high"
            elif item.severity == "中":
                tag = "mid"

            self.tree.insert(
                "",
                END,
                values=(
                    mark,
                    item.category,
                    item.severity,
                    item.name,
                    format_size(item.size_bytes),
                    item.modified,
                    item.reason,
                    item.path,
                    item.recommended_action,
                ),
                tags=(tag,),
            )
            count += 1

        self.var_only_selected_count.set(selected_count)
        self.status_label.config(
            text=f"{count}件 / 選択{selected_count}件 / 合計{format_size(total_size)}"
        )

    def on_double_click(self, event):
        row_id = self.tree.identify_row(event.y)
        col = self.tree.identify_column(event.x)
        if not row_id:
            return
        vals = list(self.tree.item(row_id, "values"))
        if col == "#1":
            vals[0] = "□" if vals[0] == "■" else "■"
            self.tree.item(row_id, values=vals)
            self.refresh_status_only()
        elif col == "#8":
            path = vals[7]
            self.copy_to_clipboard(path)

    def refresh_status_only(self):
        selected_count = 0
        count = 0
        for iid in self.tree.get_children():
            vals = self.tree.item(iid, "values")
            if not vals:
                continue
            count += 1
            if vals[0] == "■":
                selected_count += 1
        self.status_label.config(text=f"{count}件 / 選択{selected_count}件")

    def copy_to_clipboard(self, text: str):
        self.clipboard_clear()
        self.clipboard_append(text)
        self.log(f"コピーしました: {text}")

    def select_all(self):
        for iid in self.tree.get_children():
            vals = list(self.tree.item(iid, "values"))
            vals[0] = "■"
            self.tree.item(iid, values=vals)
        self.refresh_status_only()

    def clear_selection_marks(self):
        for iid in self.tree.get_children():
            vals = list(self.tree.item(iid, "values"))
            vals[0] = "□"
            self.tree.item(iid, values=vals)
        self.refresh_status_only()

    def auto_select_by_severity(self, threshold: str):
        order = {"高": 3, "中": 2, "低": 1}
        th = order.get(threshold, 3)
        for iid in self.tree.get_children():
            vals = list(self.tree.item(iid, "values"))
            sev = vals[2]
            vals[0] = "■" if order.get(sev, 0) >= th else "□"
            self.tree.item(iid, values=vals)
        self.refresh_status_only()

    def remove_nonexistent_from_list(self):
        kept = []
        removed = 0
        for item in self.scan_items:
            if Path(item.path).exists():
                kept.append(item)
            else:
                removed += 1
        self.scan_items = kept
        self.refresh_view()
        self.log(f"存在しない項目を {removed} 件除外しました。")

    def get_selected_items(self) -> list[ScanItem]:
        selected_paths = []
        for iid in self.tree.get_children():
            vals = self.tree.item(iid, "values")
            if vals and vals[0] == "■":
                selected_paths.append(vals[7])
        return [item for item in self.scan_items if item.path in selected_paths]

    def execute_selected(self):
        selected = self.get_selected_items()
        if not selected:
            messagebox.showwarning(
                "未選択",
                "処理対象が選択されていません。\n一覧の先頭列をダブルクリックして選択してください。"
            )
            return

        mode = self.var_action_mode.get()
        total = sum(item.size_bytes for item in selected)
        msg = (
            f"選択件数: {len(selected)}\n"
            f"対象容量(概算): {format_size(total)}\n\n"
            f"処理モード: {mode}\n\n"
            "この処理を実行しますか？\n"
            "最初は「隔離（推奨）」で試してください。"
        )
        if not messagebox.askyesno("確認", msg):
            return

        t = threading.Thread(target=self.execute_worker, args=(selected, mode), daemon=True)
        t.start()

    def execute_worker(self, selected: list[ScanItem], mode: str):
        engine = CleanupEngine(log_callback=self.log)
        ok_count = 0
        ng_count = 0

        for item in selected:
            p = Path(item.path)
            self.log(f"処理開始: {item.path}")
            if mode == "quarantine":
                ok, detail = engine.quarantine_item(p)
            elif mode == "delete":
                ok, detail = engine.delete_item(p)
            else:
                pkg = extract_package_name_from_dist_info(item.name)
                ok, detail = engine.pip_uninstall(pkg)

            if ok:
                ok_count += 1
                self.log(f"  成功: {detail}")
            else:
                ng_count += 1
                self.log(f"  失敗: {detail}")

        self.log(f"処理完了: 成功 {ok_count} 件 / 失敗 {ng_count} 件")
        self.after(0, self.remove_nonexistent_from_list)

    def export_csv(self):
        path = filedialog.asksaveasfilename(
            title="CSV保存",
            defaultextension=".csv",
            initialfile=DEFAULT_EXPORT.name,
            filetypes=[("CSV", "*.csv"), ("All Files", "*.*")]
        )
        if not path:
            return
        try:
            with open(path, "w", newline="", encoding="utf-8-sig") as f:
                writer = csv.writer(f)
                writer.writerow([
                    "category", "severity", "name", "path",
                    "size_bytes", "modified", "reason", "recommended_action"
                ])
                for item in self.filtered_items():
                    writer.writerow([
                        item.category,
                        item.severity,
                        item.name,
                        item.path,
                        item.size_bytes,
                        item.modified,
                        item.reason,
                        item.recommended_action,
                    ])
            self.log(f"CSV出力: {path}")
            messagebox.showinfo("完了", f"CSVを保存しました。\n{path}")
        except Exception as e:
            messagebox.showerror("エラー", f"CSV出力に失敗しました。\n{e}")


def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()