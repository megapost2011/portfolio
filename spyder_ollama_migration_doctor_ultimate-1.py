#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Spyder対応 Ollama Migration Doctor Ultimate
- Windows + Spyder 前提
- /mnt/data を一切使わない
- そのまま .py として保存して実行するだけ
"""

import os
import re
import json
import time
import shutil
import queue
import ctypes
import threading
import subprocess
from pathlib import Path
from datetime import datetime

try:
    import winreg
except ImportError:
    winreg = None

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext


APP_TITLE = "Spyder対応 Ollama Migration Doctor Ultimate"
IS_WINDOWS = os.name == "nt"


# ---------------------------------------------------------------------------
# ユーティリティ
# ---------------------------------------------------------------------------

def now():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def safe_str(x):
    try:
        return str(x)
    except Exception:
        return repr(x)


def run(cmd, timeout=30, shell=False, cwd=None):
    try:
        p = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=timeout,
            shell=shell,
            cwd=cwd,
            encoding="utf-8",
            errors="replace",
        )
        return {
            "ok": p.returncode == 0,
            "returncode": p.returncode,
            "stdout": p.stdout,
            "stderr": p.stderr,
            "cmd": cmd,
        }
    except Exception as e:
        return {
            "ok": False,
            "returncode": -1,
            "stdout": "",
            "stderr": f"{type(e).__name__}: {e}",
            "cmd": cmd,
        }


def human_size(n):
    try:
        n = float(n)
    except Exception:
        return safe_str(n)
    units = ["B", "KB", "MB", "GB", "TB"]
    for u in units:
        if n < 1024 or u == units[-1]:
            return f"{n:.2f} {u}"
        n /= 1024
    return f"{n:.2f} B"


def is_admin():
    if not IS_WINDOWS:
        return False
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False


def broadcast_env_changed():
    if not IS_WINDOWS:
        return
    try:
        HWND_BROADCAST = 0xFFFF
        WM_SETTINGCHANGE = 0x001A
        SMTO_ABORTIFHUNG = 0x0002
        ctypes.windll.user32.SendMessageTimeoutW(
            HWND_BROADCAST,
            WM_SETTINGCHANGE,
            0,
            "Environment",
            SMTO_ABORTIFHUNG,
            5000,
            None,
        )
    except Exception:
        pass


def get_user_profile():
    return os.environ.get("USERPROFILE", str(Path.home()))


def get_localappdata():
    return os.environ.get("LOCALAPPDATA", "")


def default_paths():
    home = Path(get_user_profile())
    local = Path(get_localappdata())
    return {
        "home": home,
        "default_root": home / ".ollama",
        "default_models": home / ".ollama" / "models",
        "logs_dir": local / "Ollama",
        "program_dir": local / "Programs" / "Ollama",
        "cli_candidates": [
            local / "Programs" / "Ollama" / "ollama.exe",
            Path("C:/Program Files/Ollama/ollama.exe"),
        ],
        "app_candidates": [
            local / "Programs" / "Ollama" / "ollama app.exe",
            local / "Programs" / "Ollama" / "ollama.exe",
            Path("C:/Program Files/Ollama/ollama app.exe"),
            Path("C:/Program Files/Ollama/ollama.exe"),
        ],
        "report_dir": home / "Desktop",
    }


def detect_cli():
    exe = shutil.which("ollama")
    if exe:
        return exe
    for p in default_paths()["cli_candidates"]:
        if p.exists():
            return str(p)
    return ""


def detect_app():
    for p in default_paths()["app_candidates"]:
        if p.exists():
            return str(p)
    return ""


def disk_usage_for(path):
    try:
        return shutil.disk_usage(str(path))
    except Exception:
        return None


def dir_size(path, max_files=200000):
    total = 0
    count = 0
    p = Path(path)
    if not p.exists():
        return 0, 0
    try:
        for x in p.rglob("*"):
            if x.is_file():
                try:
                    total += x.stat().st_size
                except Exception:
                    pass
                count += 1
                if count >= max_files:
                    break
    except Exception:
        pass
    return total, count


def same_path(a, b):
    try:
        return Path(a).resolve() == Path(b).resolve()
    except Exception:
        return (
            os.path.normcase(os.path.abspath(str(a)))
            == os.path.normcase(os.path.abspath(str(b)))
        )


def read_tail(path, max_chars=20000):
    p = Path(path)
    if not p.exists():
        return f"[NOT FOUND] {p}"
    try:
        txt = p.read_text(encoding="utf-8", errors="replace")
        if len(txt) > max_chars:
            return txt[-max_chars:]
        return txt
    except Exception as e:
        return f"[READ ERROR] {p}\n{type(e).__name__}: {e}"


def drive_roots():
    roots = []
    for d in ["C:\\", "D:\\", "E:\\"]:
        if Path(d).exists():
            roots.append(Path(d))
    return roots


def detect_model_structure(path):
    p = Path(path)
    problems = []
    if not p.exists():
        return False, ["フォルダが存在しません"]
    if not p.is_dir():
        return False, ["フォルダではありません"]
    blobs = p / "blobs"
    manifests = p / "manifests"
    if not blobs.exists():
        problems.append("blobs がありません")
    if not manifests.exists():
        problems.append("manifests がありません")
    blob_count = 0
    mani_count = 0
    try:
        if blobs.exists():
            for _ in blobs.rglob("*"):
                blob_count += 1
                if blob_count > 5:
                    break
        if manifests.exists():
            for _ in manifests.rglob("*"):
                mani_count += 1
                if mani_count > 5:
                    break
    except Exception:
        pass
    if blobs.exists() and blob_count == 0:
        problems.append("blobs が空に見えます")
    if manifests.exists() and mani_count == 0:
        problems.append("manifests が空に見えます")
    return len(problems) == 0, problems


def search_candidates(roots, progress=None, deep=False):
    found = {"models": [], "bins": [], "logs": [], "other": []}
    seen = set()

    def add(kind, p):
        sp = str(Path(p))
        if sp not in seen:
            seen.add(sp)
            found[kind].append(sp)

    home = Path(get_user_profile())
    local = Path(get_localappdata())

    fast_paths = [
        home / ".ollama",
        home / ".ollama" / "models",
        local / "Ollama",
        local / "Ollama" / "server.log",
        local / "Ollama" / "app.log",
        local / "Programs" / "Ollama",
        local / "Programs" / "Ollama" / "ollama.exe",
        local / "Programs" / "Ollama" / "ollama app.exe",
        Path("C:/Program Files/Ollama"),
        Path("C:/Program Files/Ollama/ollama.exe"),
        Path("C:/Program Files/Ollama/ollama app.exe"),
        Path("D:/Ollama"),
        Path("D:/Ollama/models"),
        Path("E:/Ollama"),
        Path("E:/Ollama/models"),
        Path("D:/.ollama"),
        Path("D:/.ollama/models"),
        Path("E:/.ollama"),
        Path("E:/.ollama/models"),
    ]

    for p in fast_paths:
        if progress:
            progress(f"確認: {p}")
        try:
            if p.exists():
                add("other", p)
                if p.name.lower() == "models":
                    add("models", p)
                if p.is_file() and p.name.lower() in ("ollama.exe", "ollama app.exe"):
                    add("bins", p)
                if p.is_file() and p.name.lower() in (
                    "server.log",
                    "app.log",
                    "upgrade.log",
                ):
                    add("logs", p)
                if p.is_dir():
                    for name in ("ollama.exe", "ollama app.exe"):
                        pp = p / name
                        if pp.exists():
                            add("bins", pp)
                    for name in ("server.log", "app.log", "upgrade.log"):
                        pp = p / name
                        if pp.exists():
                            add("logs", pp)
                    ok, _ = detect_model_structure(p)
                    if ok:
                        add("models", p)
                    md = p / "models"
                    if md.exists():
                        add("models", md)
        except Exception:
            pass

    if deep:
        for drive_root in roots:
            if progress:
                progress(f"詳細探索: {drive_root}")
            try:
                for base, dirs, files in os.walk(str(drive_root)):
                    base_low = base.lower()
                    if any(
                        x in base_low
                        for x in [
                            "$recycle.bin",
                            "system volume information",
                            "\\windows\\winsxs",
                            "\\node_modules\\",
                            "\\.git\\",
                            "\\venv\\",
                        ]
                    ):
                        dirs[:] = []
                        continue

                    rel_depth = len(Path(base).parts) - len(Path(drive_root).parts)
                    if rel_depth > 5:
                        dirs[:] = []
                        continue

                    if "ollama" in base_low or Path(base).name.lower() == "models":
                        add("other", base)

                    lower_files = [f.lower() for f in files]
                    if "ollama.exe" in lower_files:
                        add("bins", Path(base) / "ollama.exe")
                    if "ollama app.exe" in lower_files:
                        add("bins", Path(base) / "ollama app.exe")
                    for name in ("server.log", "app.log", "upgrade.log"):
                        if name in lower_files:
                            add("logs", Path(base) / name)

                    bp = Path(base)
                    if bp.name.lower() == "models":
                        add("models", bp)
                    else:
                        ok, _ = detect_model_structure(bp)
                        if ok:
                            add("models", bp)
            except Exception:
                pass
    return found


# ---------------------------------------------------------------------------
# レジストリ
# ---------------------------------------------------------------------------

def reg_read(root, subkey, name):
    if winreg is None:
        return ""
    try:
        with winreg.OpenKey(root, subkey) as k:
            v, _ = winreg.QueryValueEx(k, name)
            return v
    except Exception:
        return ""


def reg_enum(root, subkey):
    if winreg is None:
        return []
    out = []
    try:
        with winreg.OpenKey(root, subkey) as k:
            i = 0
            while True:
                try:
                    out.append(winreg.EnumKey(k, i))
                    i += 1
                except OSError:
                    break
    except Exception:
        pass
    return out


def get_env_registry():
    data = {
        "process": {
            "OLLAMA_MODELS": os.environ.get("OLLAMA_MODELS", ""),
            "PATH": os.environ.get("PATH", ""),
        },
        "user": {},
        "system": {},
    }
    if winreg is None:
        return data
    for name in [
        "OLLAMA_MODELS",
        "OLLAMA_HOST",
        "OLLAMA_DEBUG",
        "OLLAMA_TMPDIR",
        "PATH",
    ]:
        data["user"][name] = reg_read(
            winreg.HKEY_CURRENT_USER, r"Environment", name
        )
        data["system"][name] = reg_read(
            winreg.HKEY_LOCAL_MACHINE,
            r"SYSTEM\CurrentControlSet\Control\Session Manager\Environment",
            name,
        )
    return data


def set_user_env(name, value):
    if winreg is None:
        return False, "winreg unavailable"
    try:
        with winreg.OpenKey(
            winreg.HKEY_CURRENT_USER, r"Environment", 0, winreg.KEY_SET_VALUE
        ) as k:
            winreg.SetValueEx(k, name, 0, winreg.REG_EXPAND_SZ, str(value))
        os.environ[name] = str(value)
        broadcast_env_changed()
        return True, ""
    except Exception as e:
        return False, f"{type(e).__name__}: {e}"


def set_system_env(name, value):
    if winreg is None:
        return False, "winreg unavailable"
    try:
        with winreg.OpenKey(
            winreg.HKEY_LOCAL_MACHINE,
            r"SYSTEM\CurrentControlSet\Control\Session Manager\Environment",
            0,
            winreg.KEY_SET_VALUE,
        ) as k:
            winreg.SetValueEx(k, name, 0, winreg.REG_EXPAND_SZ, str(value))
        broadcast_env_changed()
        return True, ""
    except Exception as e:
        return False, f"{type(e).__name__}: {e}"


def get_uninstall_info():
    results = []
    if winreg is None:
        return results
    bases = [
        (
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Uninstall",
        ),
        (
            winreg.HKEY_LOCAL_MACHINE,
            r"Software\Microsoft\Windows\CurrentVersion\Uninstall",
        ),
        (
            winreg.HKEY_LOCAL_MACHINE,
            r"Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall",
        ),
    ]
    for reg_root, base in bases:
        for sub in reg_enum(reg_root, base):
            full = base + "\\" + sub
            name = reg_read(reg_root, full, "DisplayName")
            if name and "ollama" in name.lower():
                results.append(
                    {
                        "root": (
                            "HKCU"
                            if reg_root == winreg.HKEY_CURRENT_USER
                            else "HKLM"
                        ),
                        "key": full,
                        "DisplayName": name,
                        "DisplayVersion": reg_read(reg_root, full, "DisplayVersion"),
                        "InstallLocation": reg_read(
                            reg_root, full, "InstallLocation"
                        ),
                        "DisplayIcon": reg_read(reg_root, full, "DisplayIcon"),
                        "UninstallString": reg_read(
                            reg_root, full, "UninstallString"
                        ),
                    }
                )
    return results


# ---------------------------------------------------------------------------
# Ollama 操作
# ---------------------------------------------------------------------------

def stop_ollama():
    return [
        run(["taskkill", "/IM", "ollama app.exe", "/F"], timeout=15),
        run(["taskkill", "/IM", "ollama.exe", "/F"], timeout=15),
    ]


def start_ollama(app_path=""):
    p = Path(app_path) if app_path else None
    if p and p.exists():
        try:
            subprocess.Popen(
                [str(p)], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL
            )
            return True, f"起動: {p}"
        except Exception as e:
            return False, f"{type(e).__name__}: {e}"
    auto = detect_app()
    if auto:
        try:
            subprocess.Popen(
                [auto], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL
            )
            return True, f"起動: {auto}"
        except Exception as e:
            return False, f"{type(e).__name__}: {e}"
    return False, "Ollama のアプリ実体が見つかりません"


def ollama_version(cli):
    if not cli or not Path(cli).exists():
        return {"ok": False, "stdout": "", "stderr": "ollama.exe が見つかりません"}
    return run([cli, "--version"], timeout=20)


def ollama_list(cli):
    if not cli or not Path(cli).exists():
        return {"ok": False, "stdout": "", "stderr": "ollama.exe が見つかりません"}
    res = run([cli, "list"], timeout=40)
    if res["ok"]:
        return res
    return run([cli, "ls"], timeout=40)


def ollama_show(cli, model_name):
    if not cli or not Path(cli).exists():
        return {"ok": False, "stdout": "", "stderr": "ollama.exe が見つかりません"}
    return run([cli, "show", model_name], timeout=40)


def api_tags():
    ps_script = (
        "try { "
        "(Invoke-WebRequest -UseBasicParsing -Uri http://localhost:11434/api/tags"
        " -TimeoutSec 8).Content"
        " } catch { $_.Exception.Message }"
    )
    return run(
        ["powershell", "-NoProfile", "-Command", ps_script], timeout=15
    )


def parse_log_model_paths(text):
    hits = []
    patterns = [
        r'OLLAMA_MODELS[=\s:"]+([A-Za-z]:\\[^"\r\n]+)',
        r'([A-Za-z]:\\[^"\r\n]*\.ollama\\models)',
        r'([A-Za-z]:\\[^"\r\n]*\\Ollama\\models)',
    ]
    for pat in patterns:
        for m in re.finditer(pat, text, re.IGNORECASE):
            p = m.group(1).strip().rstrip('".,;')
            if p not in hits:
                hits.append(p)
    return hits


def robocopy_exists():
    return (
        shutil.which("robocopy") is not None
        or Path(r"C:\Windows\System32\robocopy.exe").exists()
    )


def robocopy_copy(src, dst):
    cmd = [
        "robocopy",
        str(src),
        str(dst),
        "/E",
        "/COPY:DAT",
        "/R:1",
        "/W:1",
        "/NFL",
        "/NDL",
        "/NP",
        "/MT:8",
    ]
    res = run(cmd, timeout=3600)
    res["ok"] = 0 <= res["returncode"] <= 7
    return res


def python_copy(src, dst, progress=None, should_stop=None):
    src = Path(src)
    dst = Path(dst)
    dst.mkdir(parents=True, exist_ok=True)
    files = []
    for p in src.rglob("*"):
        if should_stop and should_stop():
            raise RuntimeError("コピーが停止されました")
        if p.is_file():
            files.append(p)
    total = len(files)
    done = 0
    for p in files:
        if should_stop and should_stop():
            raise RuntimeError("コピーが停止されました")
        rel = p.relative_to(src)
        out = dst / rel
        out.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(p, out)
        done += 1
        if progress:
            progress(done, total, str(rel))


def verify_target(path):
    ok, probs = detect_model_structure(path)
    size, count = dir_size(path, max_files=50000)
    blobs = Path(path) / "blobs"
    manifests = Path(path) / "manifests"
    blob_count = 0
    mani_count = 0
    try:
        if blobs.exists():
            for _ in blobs.rglob("*"):
                blob_count += 1
        if manifests.exists():
            for _ in manifests.rglob("*"):
                mani_count += 1
    except Exception:
        pass
    return {
        "ok": ok,
        "problems": probs,
        "size": size,
        "count": count,
        "blob_count": blob_count,
        "manifest_count": mani_count,
    }


# ---------------------------------------------------------------------------
# GUI
# ---------------------------------------------------------------------------

class App:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("1420x920")
        self.root.minsize(1180, 760)

        # ウィンドウ破棄フラグ（Spyder 多重実行対策）
        self._destroyed = False
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

        self.queue = queue.Queue()
        self.worker = None
        self.stop_flag = False

        self.cli_var = tk.StringVar(value=detect_cli())
        self.app_var = tk.StringVar(value=detect_app())
        self.src_var = tk.StringVar(value="")
        self.dst_var = tk.StringVar(
            value=r"E:\Ollama\models" if Path("E:\\").exists() else ""
        )
        self.model_var = tk.StringVar(value="gemma3:4b")
        self.deep_scan_var = tk.BooleanVar(value=False)
        self.fix_system_env_var = tk.BooleanVar(value=False)
        self.restart_var = tk.BooleanVar(value=True)
        self.backup_old_var = tk.BooleanVar(value=False)
        self.status_var = tk.StringVar(value="準備完了")

        self.last_found = {"models": [], "bins": [], "logs": [], "other": []}
        self.last_env = {}
        self.last_uninstall = []

        self.build_ui()
        self.root.after(150, self.poll)

    # -----------------------------------------------------------------------
    # ウィンドウ終了
    # -----------------------------------------------------------------------

    def on_close(self):
        """ウィンドウを閉じるときの処理。ワーカー停止 → 破棄。"""
        self._destroyed = True
        self.stop_flag = True
        try:
            self.root.destroy()
        except Exception:
            pass

    # -----------------------------------------------------------------------
    # UI 構築
    # -----------------------------------------------------------------------

    def build_ui(self):
        top = ttk.Frame(self.root, padding=8)
        top.pack(fill="x")

        rows = [
            ("Ollama CLI", self.cli_var, self.pick_cli, self.autodetect),
            ("Ollama App", self.app_var, self.pick_app, None),
            ("移行元 models", self.src_var, self.pick_src, None),
            ("移行先 models", self.dst_var, self.pick_dst, None),
        ]
        for lbl, var, browse, auto in rows:
            row = ttk.Frame(top)
            row.pack(fill="x", pady=2)
            ttk.Label(row, text=lbl, width=14).pack(side="left")
            ttk.Entry(row, textvariable=var).pack(
                side="left", fill="x", expand=True, padx=6
            )
            ttk.Button(row, text="参照", command=browse).pack(side="left", padx=2)
            if auto:
                ttk.Button(row, text="自動検出", command=auto).pack(
                    side="left", padx=2
                )

        row = ttk.Frame(top)
        row.pack(fill="x", pady=4)
        ttk.Checkbutton(row, text="詳細探索", variable=self.deep_scan_var).pack(
            side="left", padx=4
        )
        ttk.Checkbutton(
            row, text="システム環境変数も修正", variable=self.fix_system_env_var
        ).pack(side="left", padx=8)
        ttk.Checkbutton(
            row, text="修正後に Ollama を起動", variable=self.restart_var
        ).pack(side="left", padx=8)
        ttk.Checkbutton(
            row, text="旧保存先を backup_old へ退避", variable=self.backup_old_var
        ).pack(side="left", padx=8)
        ttk.Label(row, text="確認モデル").pack(side="left", padx=(20, 4))
        ttk.Entry(row, textvariable=self.model_var, width=18).pack(side="left")

        row = ttk.Frame(top)
        row.pack(fill="x", pady=6)
        buttons = [
            ("1. 全探索・診断", self.full_scan),
            ("2. 完全自動修復", self.auto_fix),
            ("3. 移行コピーのみ", self.copy_only),
            ("4. 環境変数のみ修正", self.fix_env_only),
            ("5. Ollama 停止", self.stop_ollama_action),
            ("6. Ollama 起動", self.start_ollama_action),
            ("7. 実行確認", self.verify_action),
            ("レポート保存", self.save_report),
            ("停止", self.stop_copy),
            ("クリア", self.clear_all),
        ]
        for text, cmd in buttons:
            ttk.Button(row, text=text, command=cmd).pack(side="left", padx=2)

        pane = ttk.Panedwindow(self.root, orient="horizontal")
        pane.pack(fill="both", expand=True, padx=8, pady=8)

        left = ttk.Frame(pane)
        mid = ttk.Frame(pane)
        right = ttk.Frame(pane)
        pane.add(left, weight=3)
        pane.add(mid, weight=2)
        pane.add(right, weight=2)

        ttk.Label(left, text="実行ログ").pack(anchor="w")
        self.log = scrolledtext.ScrolledText(
            left, wrap="word", font=("Consolas", 10)
        )
        self.log.pack(fill="both", expand=True, pady=4)

        ttk.Label(mid, text="検出結果").pack(anchor="w")
        self.detect = scrolledtext.ScrolledText(
            mid, wrap="word", font=("Consolas", 10)
        )
        self.detect.pack(fill="both", expand=True, pady=4)

        ttk.Label(right, text="環境・レジストリ・ログ").pack(anchor="w")
        self.detail = scrolledtext.ScrolledText(
            right, wrap="word", font=("Consolas", 10)
        )
        self.detail.pack(fill="both", expand=True, pady=4)

        bottom = ttk.Frame(self.root, padding=8)
        bottom.pack(fill="x")
        ttk.Label(bottom, textvariable=self.status_var).pack(side="left")
        ttk.Label(
            bottom,
            text=f"管理者権限: {'はい' if is_admin() else 'いいえ'}",
        ).pack(side="right")

    # -----------------------------------------------------------------------
    # UI ヘルパー
    # -----------------------------------------------------------------------

    def clear_all(self):
        self.log.delete("1.0", "end")
        self.detect.delete("1.0", "end")
        self.detail.delete("1.0", "end")

    def set_status(self, text):
        try:
            self.status_var.set(text)
        except tk.TclError:
            pass

    def append(self, widget, text):
        try:
            widget.insert("end", text + "\n")
            widget.see("end")
        except tk.TclError:
            pass

    def qput(self, kind, text):
        self.queue.put((kind, text))

    def poll(self):
        """メインスレッドでキューを処理。TclError はウィンドウ破棄を意味する。"""
        if self._destroyed:
            return
        try:
            while True:
                kind, text = self.queue.get_nowait()
                if kind == "status":
                    self.set_status(text)
                elif kind == "log":
                    self.append(self.log, text)
                elif kind == "detect":
                    self.append(self.detect, text)
                elif kind == "detail":
                    self.append(self.detail, text)
                elif kind == "error":
                    self.set_status("エラー")
                    try:
                        messagebox.showerror(APP_TITLE, text)
                    except tk.TclError:
                        pass
        except queue.Empty:
            pass
        except tk.TclError:
            # ウィンドウが破棄された → poll を止める
            self._destroyed = True
            return
        except Exception:
            pass

        if not self._destroyed:
            try:
                self.root.after(150, self.poll)
            except tk.TclError:
                self._destroyed = True

    # -----------------------------------------------------------------------
    # ワーカースレッド
    # -----------------------------------------------------------------------

    def start_worker(self, fn):
        if self.worker and self.worker.is_alive():
            messagebox.showwarning(APP_TITLE, "他の処理を実行中です。")
            return

        def _wrapped():
            try:
                fn()
            except Exception as e:
                self.qput(
                    "error",
                    f"予期しないエラーが発生しました:\n{type(e).__name__}: {e}",
                )

        self.worker = threading.Thread(target=_wrapped, daemon=True)
        self.worker.start()

    # -----------------------------------------------------------------------
    # ファイル選択
    # -----------------------------------------------------------------------

    def pick_cli(self):
        p = filedialog.askopenfilename(
            title="ollama.exe を選択",
            filetypes=[("Executable", "*.exe"), ("All files", "*.*")],
        )
        if p:
            self.cli_var.set(p)

    def pick_app(self):
        p = filedialog.askopenfilename(
            title="ollama app.exe を選択",
            filetypes=[("Executable", "*.exe"), ("All files", "*.*")],
        )
        if p:
            self.app_var.set(p)

    def pick_src(self):
        p = filedialog.askdirectory(title="移行元 models を選択")
        if p:
            self.src_var.set(p)

    def pick_dst(self):
        p = filedialog.askdirectory(title="移行先 models を選択")
        if p:
            self.dst_var.set(p)

    # -----------------------------------------------------------------------
    # 自動検出
    # -----------------------------------------------------------------------

    def autodetect(self):
        self.cli_var.set(detect_cli())
        self.app_var.set(detect_app())
        defaults = default_paths()
        if defaults["default_models"].exists() and not self.src_var.get():
            self.src_var.set(str(defaults["default_models"]))
        if not self.dst_var.get() and Path("E:\\").exists():
            self.dst_var.set(r"E:\Ollama\models")
        messagebox.showinfo(APP_TITLE, "自動検出を更新しました。")

    # -----------------------------------------------------------------------
    # 1. 全探索・診断
    # -----------------------------------------------------------------------

    def full_scan(self):
        self.start_worker(self._full_scan)

    def _full_scan(self):
        self.qput("status", "全探索・診断中...")
        self.qput("log", f"=== {now()} 全探索・診断開始 ===")

        roots = drive_roots()
        self.qput("log", "対象ドライブ: " + ", ".join(str(x) for x in roots))
        found = search_candidates(
            roots,
            progress=lambda m: self.qput("status", m),
            deep=self.deep_scan_var.get(),
        )
        self.last_found = found

        env = get_env_registry()
        self.last_env = env
        uninstall = get_uninstall_info()
        self.last_uninstall = uninstall

        self.qput("detect", "--- models 候補 ---")
        for p in found["models"]:
            ok, probs = detect_model_structure(p)
            size, count = dir_size(p, max_files=30000)
            self.qput("detect", p)
            self.qput("detect", f"  構造OK: {'はい' if ok else 'いいえ'}")
            self.qput("detect", f"  サイズ: {human_size(size)} / 走査数: {count}")
            for pr in probs:
                self.qput("detect", f"  - {pr}")

        self.qput("detect", "")
        self.qput("detect", "--- バイナリ候補 ---")
        for p in found["bins"]:
            self.qput("detect", p)

        self.qput("detect", "")
        self.qput("detect", "--- ログ候補 ---")
        for p in found["logs"]:
            self.qput("detect", p)

        defaults = default_paths()
        logs_dir = defaults["logs_dir"]
        self.qput("detail", "--- 既定パス ---")
        self.qput("detail", f"default_models = {defaults['default_models']}")
        self.qput("detail", f"logs_dir       = {logs_dir}")
        self.qput("detail", f"program_dir    = {defaults['program_dir']}")

        self.qput("detail", "")
        self.qput("detail", "--- 環境変数 ---")
        self.qput("detail", json.dumps(env, ensure_ascii=False, indent=2))

        self.qput("detail", "")
        self.qput("detail", "--- レジストリの Ollama 情報 ---")
        if uninstall:
            for item in uninstall:
                self.qput("detail", json.dumps(item, ensure_ascii=False, indent=2))
        else:
            self.qput("detail", "見つかりませんでした。")

        self.qput("detail", "")
        self.qput("detail", "--- 既定ログ抜粋 ---")
        for name in ["app.log", "server.log", "upgrade.log"]:
            p = logs_dir / name
            self.qput("detail", f"##### {p}")
            txt = read_tail(p, max_chars=10000)
            self.qput("detail", txt)
            if name == "server.log":
                hits = parse_log_model_paths(txt)
                if hits:
                    self.qput("detail", "server.log から抽出したパス候補:")
                    for h in hits:
                        self.qput("detail", f"  {h}")

        self.qput("log", "")
        self.qput("log", "--- CLI / API 確認 ---")
        cli = self.cli_var.get().strip() or detect_cli()
        self.cli_var.set(cli)
        ver = ollama_version(cli)
        self.qput(
            "log",
            f"[version] ok={ver['ok']} stdout={ver['stdout'].strip()} "
            f"stderr={ver['stderr'].strip()}",
        )

        ls = ollama_list(cli)
        self.qput("log", "[list]")
        self.qput("log", ls["stdout"].strip() or "(出力なし)")
        if ls["stderr"].strip():
            self.qput("log", ls["stderr"].strip())

        model_name = self.model_var.get().strip() or "gemma3:4b"
        sh = ollama_show(cli, model_name)
        self.qput("log", f"[show {model_name}]")
        self.qput("log", sh["stdout"].strip() or "(出力なし)")
        if sh["stderr"].strip():
            self.qput("log", sh["stderr"].strip())

        api = api_tags()
        self.qput("log", "[/api/tags]")
        self.qput("log", api["stdout"].strip() or "(出力なし)")
        if api["stderr"].strip():
            self.qput("log", api["stderr"].strip())

        self.qput("status", "全探索・診断完了")

    # -----------------------------------------------------------------------
    # 移行元推定
    # -----------------------------------------------------------------------

    def infer_source(self):
        src = self.src_var.get().strip()
        if src and Path(src).exists():
            return src
        candidates = self.last_found.get("models", [])
        for drive in ["D:\\", "C:\\", "E:\\"]:
            for c in candidates:
                if str(c).upper().startswith(drive.upper()):
                    return c
        defaults = default_paths()["default_models"]
        if defaults.exists():
            return str(defaults)
        return ""

    # -----------------------------------------------------------------------
    # 3. 移行コピーのみ
    # -----------------------------------------------------------------------

    def copy_only(self):
        self.start_worker(self._copy_only)

    def _copy_only(self):
        src = self.infer_source()
        dst = self.dst_var.get().strip()
        if not src or not Path(src).exists():
            self.qput("error", "移行元 models が見つかりません。")
            return
        if not dst:
            self.qput("error", "移行先 models を指定してください。")
            return
        if same_path(src, dst):
            self.qput("error", "移行元と移行先が同一です。")
            return

        self.src_var.set(src)
        self.qput("status", "移行コピー中...")
        self.qput("log", f"=== {now()} 移行コピー開始 ===")
        self.qput("log", f"{src} -> {dst}")

        src_size, _ = dir_size(src, max_files=100000)
        usage = disk_usage_for(Path(dst).anchor or dst)
        if usage:
            self.qput("log", f"移行元概算サイズ: {human_size(src_size)}")
            self.qput("log", f"移行先空き容量: {human_size(usage.free)}")
            if src_size and usage.free < src_size:
                self.qput("error", "移行先の空き容量が不足している可能性があります。")
                return

        Path(dst).mkdir(parents=True, exist_ok=True)
        self.stop_flag = False
        try:
            if robocopy_exists():
                rc = robocopy_copy(src, dst)
                self.qput("log", f"robocopy rc={rc['returncode']} ok={rc['ok']}")
                if rc["stdout"].strip():
                    self.qput("log", rc["stdout"][:12000])
                if not rc["ok"]:
                    self.qput("log", "robocopy 失敗のため Python コピーへフォールバック")
                    python_copy(
                        src,
                        dst,
                        progress=lambda d, t, n: self.qput(
                            "status", f"コピー中 {d}/{t}: {n}"
                        ),
                        should_stop=lambda: self.stop_flag,
                    )
            else:
                python_copy(
                    src,
                    dst,
                    progress=lambda d, t, n: self.qput(
                        "status", f"コピー中 {d}/{t}: {n}"
                    ),
                    should_stop=lambda: self.stop_flag,
                )
        except Exception as e:
            self.qput("error", f"コピー失敗: {type(e).__name__}: {e}")
            return

        v = verify_target(dst)
        self.qput("log", json.dumps(v, ensure_ascii=False, indent=2))
        self.qput("status", "移行コピー完了")

    # -----------------------------------------------------------------------
    # 4. 環境変数のみ修正
    # -----------------------------------------------------------------------

    def fix_env_only(self):
        self.start_worker(self._fix_env_only)

    def _fix_env_only(self):
        dst = self.dst_var.get().strip()
        if not dst:
            self.qput("error", "移行先 models を指定してください。")
            return
        self.qput("status", "環境変数修正中...")
        self.qput("log", f"=== {now()} 環境変数修正 ===")

        ok, err = set_user_env("OLLAMA_MODELS", dst)
        self.qput("log", f"ユーザー環境変数 OLLAMA_MODELS: {'OK' if ok else 'NG'}")
        if err:
            self.qput("log", err)

        if self.fix_system_env_var.get():
            if is_admin():
                ok2, err2 = set_system_env("OLLAMA_MODELS", dst)
                self.qput(
                    "log",
                    f"システム環境変数 OLLAMA_MODELS: {'OK' if ok2 else 'NG'}",
                )
                if err2:
                    self.qput("log", err2)
            else:
                self.qput("log", "管理者権限がないためシステム環境変数は変更しませんでした。")

        self.qput("status", "環境変数修正完了")

    # -----------------------------------------------------------------------
    # 旧フォルダ退避
    # -----------------------------------------------------------------------

    def backup_old_source(self, src):
        src_p = Path(src)
        backup = src_p.parent / "backup_old"
        if same_path(src_p, backup):
            return False, "backup_old と移行元が同一です"
        try:
            if backup.exists():
                backup = src_p.parent / (
                    f"backup_old_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
                )
            shutil.move(str(src_p), str(backup))
            return True, str(backup)
        except Exception as e:
            return False, f"{type(e).__name__}: {e}"

    # -----------------------------------------------------------------------
    # 2. 完全自動修復
    # -----------------------------------------------------------------------

    def auto_fix(self):
        self.start_worker(self._auto_fix)

    def _auto_fix(self):
        self.qput("status", "完全自動修復中...")
        self.qput("log", f"=== {now()} 完全自動修復開始 ===")

        cli = self.cli_var.get().strip() or detect_cli()
        app = self.app_var.get().strip() or detect_app()
        self.cli_var.set(cli)
        self.app_var.set(app)

        src = self.infer_source()
        dst = self.dst_var.get().strip()

        if not src or not Path(src).exists():
            self.qput(
                "error", "移行元 models が見つかりません。先に全探索・診断を実行してください。"
            )
            return
        if not dst:
            self.qput("error", "移行先 models が未指定です。")
            return
        if same_path(src, dst):
            self.qput("error", "移行元と移行先が同一です。")
            return

        self.src_var.set(src)
        self.qput("log", f"移行元: {src}")
        self.qput("log", f"移行先: {dst}")

        self.qput("log", "--- Ollama 停止 ---")
        for res in stop_ollama():
            self.qput("log", f"cmd={res['cmd']} rc={res['returncode']}")
            if res["stdout"].strip():
                self.qput("log", res["stdout"].strip())
            if res["stderr"].strip():
                self.qput("log", res["stderr"].strip())

        src_size, _ = dir_size(src, max_files=100000)
        usage = disk_usage_for(Path(dst).anchor or dst)
        if usage:
            self.qput("log", f"移行元概算サイズ: {human_size(src_size)}")
            self.qput("log", f"移行先空き容量: {human_size(usage.free)}")
            if src_size and usage.free < src_size:
                self.qput("error", "移行先の空き容量が不足しています。")
                return

        Path(dst).mkdir(parents=True, exist_ok=True)
        self.stop_flag = False
        try:
            if robocopy_exists():
                rc = robocopy_copy(src, dst)
                self.qput("log", f"robocopy rc={rc['returncode']} ok={rc['ok']}")
                if rc["stdout"].strip():
                    self.qput("log", rc["stdout"][:12000])
                if not rc["ok"]:
                    python_copy(
                        src,
                        dst,
                        progress=lambda d, t, n: self.qput(
                            "status", f"コピー中 {d}/{t}: {n}"
                        ),
                        should_stop=lambda: self.stop_flag,
                    )
            else:
                python_copy(
                    src,
                    dst,
                    progress=lambda d, t, n: self.qput(
                        "status", f"コピー中 {d}/{t}: {n}"
                    ),
                    should_stop=lambda: self.stop_flag,
                )
        except Exception as e:
            self.qput("error", f"コピー失敗: {type(e).__name__}: {e}")
            return

        verify = verify_target(dst)
        self.qput("log", "--- コピー先検証 ---")
        self.qput("log", json.dumps(verify, ensure_ascii=False, indent=2))
        if not verify["ok"]:
            self.qput("error", "コピー先の models 構造が不完全です。")
            return

        self.qput("log", "--- 環境変数修正 ---")
        ok, err = set_user_env("OLLAMA_MODELS", dst)
        self.qput("log", f"ユーザー OLLAMA_MODELS: {'OK' if ok else 'NG'}")
        if err:
            self.qput("log", err)

        if self.fix_system_env_var.get():
            if is_admin():
                ok2, err2 = set_system_env("OLLAMA_MODELS", dst)
                self.qput(
                    "log", f"システム OLLAMA_MODELS: {'OK' if ok2 else 'NG'}"
                )
                if err2:
                    self.qput("log", err2)
            else:
                self.qput("log", "管理者権限ではないためシステム環境変数は未変更です。")

        os.environ["OLLAMA_MODELS"] = dst

        if self.backup_old_var.get():
            self.qput("log", "--- 旧保存先の退避 ---")
            ok3, msg3 = self.backup_old_source(src)
            self.qput("log", f"{'OK' if ok3 else 'NG'}: {msg3}")

        if self.restart_var.get():
            self.qput("log", "--- Ollama 起動 ---")
            ok4, msg4 = start_ollama(app)
            self.qput("log", msg4)
            time.sleep(5)

        self.qput("log", "--- 実行確認 ---")
        ver = ollama_version(cli)
        self.qput(
            "log",
            f"[version] ok={ver['ok']} stdout={ver['stdout'].strip()} "
            f"stderr={ver['stderr'].strip()}",
        )

        ls = ollama_list(cli)
        self.qput("log", "[list]")
        self.qput("log", ls["stdout"].strip() or "(出力なし)")
        if ls["stderr"].strip():
            self.qput("log", ls["stderr"].strip())

        model_name = self.model_var.get().strip() or "gemma3:4b"
        sh = ollama_show(cli, model_name)
        self.qput("log", f"[show {model_name}]")
        self.qput("log", sh["stdout"].strip() or "(出力なし)")
        if sh["stderr"].strip():
            self.qput("log", sh["stderr"].strip())

        api = api_tags()
        self.qput("log", "[/api/tags]")
        self.qput("log", api["stdout"].strip() or "(出力なし)")
        if api["stderr"].strip():
            self.qput("log", api["stderr"].strip())

        logs_dir = default_paths()["logs_dir"]
        self.qput("detail", "--- 最新ログ再確認 ---")
        for name in ["server.log", "app.log"]:
            p = logs_dir / name
            txt = read_tail(p, max_chars=8000)
            self.qput("detail", f"##### {p}")
            self.qput("detail", txt)
            if name == "server.log":
                hits = parse_log_model_paths(txt)
                if hits:
                    self.qput("detail", "server.log 抽出パス:")
                    for h in hits:
                        self.qput("detail", f"  {h}")

        self.qput("status", "完全自動修復完了")

    # -----------------------------------------------------------------------
    # 5. Ollama 停止
    # -----------------------------------------------------------------------

    def stop_ollama_action(self):
        self.start_worker(self._stop_ollama)

    def _stop_ollama(self):
        self.qput("status", "Ollama 停止中...")
        self.qput("log", f"=== {now()} Ollama 停止 ===")
        for res in stop_ollama():
            self.qput("log", f"cmd={res['cmd']} rc={res['returncode']}")
            if res["stdout"].strip():
                self.qput("log", res["stdout"].strip())
            if res["stderr"].strip():
                self.qput("log", res["stderr"].strip())
        self.qput("status", "停止完了")

    # -----------------------------------------------------------------------
    # 6. Ollama 起動
    # -----------------------------------------------------------------------

    def start_ollama_action(self):
        self.start_worker(self._start_ollama)

    def _start_ollama(self):
        self.qput("status", "Ollama 起動中...")
        self.qput("log", f"=== {now()} Ollama 起動 ===")
        ok, msg = start_ollama(self.app_var.get().strip())
        self.qput("log", msg)
        self.qput("status", "起動完了")

    # -----------------------------------------------------------------------
    # 7. 実行確認
    # -----------------------------------------------------------------------

    def verify_action(self):
        self.start_worker(self._verify)

    def _verify(self):
        self.qput("status", "実行確認中...")
        self.qput("log", f"=== {now()} 実行確認 ===")

        cli = self.cli_var.get().strip() or detect_cli()
        self.cli_var.set(cli)

        ver = ollama_version(cli)
        self.qput(
            "log",
            f"[version] ok={ver['ok']} stdout={ver['stdout'].strip()} "
            f"stderr={ver['stderr'].strip()}",
        )

        ls = ollama_list(cli)
        self.qput("log", "[list]")
        self.qput("log", ls["stdout"].strip() or "(出力なし)")
        if ls["stderr"].strip():
            self.qput("log", ls["stderr"].strip())

        model_name = self.model_var.get().strip() or "gemma3:4b"
        sh = ollama_show(cli, model_name)
        self.qput("log", f"[show {model_name}]")
        self.qput("log", sh["stdout"].strip() or "(出力なし)")
        if sh["stderr"].strip():
            self.qput("log", sh["stderr"].strip())

        api = api_tags()
        self.qput("log", "[/api/tags]")
        self.qput("log", api["stdout"].strip() or "(出力なし)")
        if api["stderr"].strip():
            self.qput("log", api["stderr"].strip())

        self.qput("status", "実行確認完了")

    # -----------------------------------------------------------------------
    # コピー停止 / レポート保存
    # -----------------------------------------------------------------------

    def stop_copy(self):
        self.stop_flag = True
        self.set_status("停止要求を送信しました")

    def save_report(self):
        default_name = (
            f"ollama_migration_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
        )
        p = filedialog.asksaveasfilename(
            title="レポート保存",
            initialdir=str(default_paths()["report_dir"]),
            initialfile=default_name,
            defaultextension=".txt",
            filetypes=[("Text", "*.txt"), ("All files", "*.*")],
        )
        if not p:
            return
        content = [
            APP_TITLE,
            f"保存日時: {now()}",
            "",
            "=== 実行ログ ===",
            self.log.get("1.0", "end"),
            "",
            "=== 検出結果 ===",
            self.detect.get("1.0", "end"),
            "",
            "=== 環境・レジストリ・ログ ===",
            self.detail.get("1.0", "end"),
        ]
        Path(p).write_text("\n".join(content), encoding="utf-8", errors="replace")
        messagebox.showinfo(APP_TITLE, f"保存しました:\n{p}")


# ---------------------------------------------------------------------------
# エントリポイント
# ---------------------------------------------------------------------------

def main():
    if not IS_WINDOWS:
        print("このアプリは Windows 向けです。")
        return

    # Spyder で 2 回目以降に実行すると既存 Tk インスタンスが残る場合がある。
    # TclError をキャッチして案内メッセージを出す。
    try:
        root = tk.Tk()
    except tk.TclError as e:
        print(f"[Tkinter 初期化エラー] {e}")
        print(
            "ヒント: Spyder でこのエラーが出る場合は\n"
            "  コンソールを再起動 (Ctrl + .) してから再実行してください。"
        )
        return

    try:
        style = ttk.Style()
        for theme in ("vista", "xpnative", "clam"):
            try:
                style.theme_use(theme)
                break
            except Exception:
                pass
    except Exception:
        pass

    App(root)

    try:
        root.mainloop()
    except KeyboardInterrupt:
        pass
    except Exception as e:
        print(f"[mainloop エラー] {type(e).__name__}: {e}")


if __name__ == "__main__":
    main()
