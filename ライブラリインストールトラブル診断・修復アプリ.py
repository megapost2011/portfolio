#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
ライブラリ インストール診断 & 自動自己修復 GUI
- Pydroid / Windows / Linux を意識した単一ファイル版
- 任意のパッケージ名を入力して診断
- 診断結果に応じた修復候補を提示
- 任意の修復方法を選んで自動実行

注意:
- すべてのパッケージを自動修復できるわけではありません。
- 特に Android/Pydroid 上のネイティブ拡張ライブラリは、
  wheel が存在しないと修復不能な場合があります。
"""

import os
import re
import sys
import json
import time
import queue
import shutil
import tempfile
import threading
import subprocess
import platform
import urllib.request
import urllib.error
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

# 可能なら packaging を使う
try:
    from packaging import tags as packaging_tags
    from packaging import version as packaging_version
    from packaging.specifiers import SpecifierSet
    HAS_PACKAGING = True
except Exception:
    HAS_PACKAGING = False


APP_TITLE = "pipライブラリ診断・自動自己修復GUI"
TIMEOUT_SEC = 180


def now_str():
    return time.strftime("%Y-%m-%d %H:%M:%S")


def is_android():
    return "ANDROID_ARGUMENT" in os.environ or "android" in platform.platform().lower()


def is_pydroid():
    text = " ".join([
        os.environ.get("ANDROID_ARGUMENT", ""),
        sys.executable,
        sys.prefix,
        platform.platform(),
    ]).lower()
    return "pydroid" in text or "ru.iiec.pydroid3" in text or is_android()


def get_python_mm():
    return f"{sys.version_info.major}.{sys.version_info.minor}"


def safe_get(obj, *keys, default=None):
    cur = obj
    for k in keys:
        if isinstance(cur, dict) and k in cur:
            cur = cur[k]
        else:
            return default
    return cur


def run_cmd(cmd, timeout=TIMEOUT_SEC, cwd=None):
    try:
        p = subprocess.run(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=timeout,
            cwd=cwd,
        )
        return {
            "ok": p.returncode == 0,
            "returncode": p.returncode,
            "stdout": p.stdout,
            "stderr": p.stderr,
            "cmd": cmd,
        }
    except subprocess.TimeoutExpired as e:
        return {
            "ok": False,
            "returncode": -999,
            "stdout": e.stdout if e.stdout else "",
            "stderr": f"TimeoutExpired: {e}",
            "cmd": cmd,
        }
    except Exception as e:
        return {
            "ok": False,
            "returncode": -998,
            "stdout": "",
            "stderr": repr(e),
            "cmd": cmd,
        }


def fetch_pypi_json(package_name):
    url = f"https://pypi.org/pypi/{package_name}/json"
    req = urllib.request.Request(
        url,
        headers={
            "User-Agent": "pip-diagnostic-gui/1.0"
        }
    )
    with urllib.request.urlopen(req, timeout=20) as resp:
        return json.loads(resp.read().decode("utf-8"))


def parse_wheel_filename(filename):
    """
    ざっくり wheel 名を分解
    例:
      PyMuPDF-1.26.5-cp310-abi3-manylinux_2_28_x86_64.whl
    """
    if not filename.endswith(".whl"):
        return None
    m = re.match(r"^(?P<namever>.+)-(?P<py>[^-]+)-(?P<abi>[^-]+)-(?P<plat>[^-]+)\.whl$", filename)
    if not m:
        return None
    return m.groupdict()


def environment_summary():
    return {
        "python": sys.version.replace("\n", " "),
        "python_mm": get_python_mm(),
        "executable": sys.executable,
        "platform": platform.platform(),
        "machine": platform.machine(),
        "android": is_android(),
        "pydroid": is_pydroid(),
        "cwd": os.getcwd(),
        "prefix": sys.prefix,
    }


def get_sys_tags_text(limit=50):
    if not HAS_PACKAGING:
        return []
    out = []
    try:
        for i, t in enumerate(packaging_tags.sys_tags()):
            out.append(str(t))
            if i + 1 >= limit:
                break
    except Exception:
        pass
    return out


def python_version_satisfies(requires_python):
    if not requires_python or not HAS_PACKAGING:
        return None
    try:
        spec = SpecifierSet(requires_python)
        cur = packaging_version.Version(f"{sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}")
        return cur in spec
    except Exception:
        return None


def collect_release_files(pypi_json):
    releases = safe_get(pypi_json, "releases", default={}) or {}
    files = []
    for ver, arr in releases.items():
        for item in arr:
            files.append({
                "version": ver,
                "filename": item.get("filename", ""),
                "packagetype": item.get("packagetype", ""),
                "python_version": item.get("python_version", ""),
                "requires_python": item.get("requires_python"),
                "yanked": item.get("yanked", False),
                "url": item.get("url", ""),
            })
    return files


def has_pure_python_wheel(files):
    for f in files:
        fn = f.get("filename", "")
        if fn.endswith("py3-none-any.whl") or fn.endswith("py2.py3-none-any.whl"):
            return True
    return False


def has_any_wheel(files):
    return any(f.get("filename", "").endswith(".whl") for f in files)


def likely_native_only(files):
    if not files:
        return False
    any_whl = has_any_wheel(files)
    pure_any = has_pure_python_wheel(files)
    any_sdist = any(f.get("packagetype") == "sdist" or f.get("filename", "").endswith((".tar.gz", ".zip")) for f in files)
    return (any_whl or any_sdist) and (not pure_any)


def find_pure_python_versions(releases):
    """
    py3-none-any wheel を持つバージョンを返す
    """
    out = []
    for ver, arr in releases.items():
        for item in arr:
            fn = item.get("filename", "")
            if fn.endswith("py3-none-any.whl") or fn.endswith("py2.py3-none-any.whl"):
                out.append(ver)
                break
    def sort_key(v):
        if HAS_PACKAGING:
            try:
                return packaging_version.Version(v)
            except Exception:
                return packaging_version.Version("0")
        return v
    try:
        out = sorted(set(out), key=sort_key, reverse=True)
    except Exception:
        out = sorted(set(out), reverse=True)
    return out


def diagnose_package(package_name, ui_log=None):
    result = {
        "env": environment_summary(),
        "package": package_name,
        "diagnosis_lines": [],
        "repair_actions": [],
        "raw": {},
    }

    def log(msg):
        if ui_log:
            ui_log(msg)

    env = result["env"]
    log(f"[{now_str()}] 診断開始: {package_name}")

    # pip バージョン確認
    pip_ver = run_cmd([sys.executable, "-m", "pip", "--version"], timeout=30)
    result["raw"]["pip_version"] = pip_ver
    if pip_ver["ok"]:
        result["diagnosis_lines"].append(f"pip は利用可能です: {pip_ver['stdout'].strip()}")
    else:
        result["diagnosis_lines"].append("pip 自体の起動に失敗しています。pipの破損の可能性があります。")

    # PyPI 取得
    try:
        pypi_json = fetch_pypi_json(package_name)
        result["raw"]["pypi_json"] = pypi_json
        info = pypi_json.get("info", {})
        releases = pypi_json.get("releases", {})
        all_files = collect_release_files(pypi_json)
        latest_ver = info.get("version", "不明")
        requires_python = info.get("requires_python")

        result["diagnosis_lines"].append(f"PyPI 上でパッケージを確認しました: {package_name} {latest_ver}")
        if requires_python:
            result["diagnosis_lines"].append(f"要求Python条件: {requires_python}")
            sat = python_version_satisfies(requires_python)
            if sat is False:
                result["diagnosis_lines"].append(
                    f"現在の Python {get_python_mm()} は要求条件を満たしていない可能性があります。"
                )
            elif sat is True:
                result["diagnosis_lines"].append(
                    f"現在の Python {get_python_mm()} は要求条件に適合しています。"
                )

        latest_files = releases.get(latest_ver, [])
        latest_names = [x.get("filename", "") for x in latest_files]

        any_wheel = has_any_wheel(latest_files)
        pure_any = has_pure_python_wheel(latest_files)
        native_pkg = likely_native_only(latest_files)

        if pure_any:
            result["diagnosis_lines"].append("最新版には pure Python wheel (py3-none-any) が存在します。")
        elif any_wheel:
            result["diagnosis_lines"].append("最新版にはプラットフォーム依存 wheel があります。")
        else:
            result["diagnosis_lines"].append("最新版には wheel が見当たらず、ソースビルドが必要な可能性があります。")

        # Android/Pydroid 特化診断
        if env["android"] or env["pydroid"]:
            if native_pkg:
                result["diagnosis_lines"].append(
                    "このパッケージはネイティブ拡張を含む可能性が高く、Android/Pydroid では wheel 不足により失敗しやすいです。"
                )
            if package_name.lower() in ("pymupdf", "fitz"):
                result["diagnosis_lines"].append(
                    "PyMuPDF は Android/Pydroid では特に失敗しやすい代表例です。"
                )

        # pip download による安全寄り診断
        tmpdir = tempfile.mkdtemp(prefix="pip_diag_")
        try:
            log("pip download で事前診断中...")
            dl = run_cmd(
                [
                    sys.executable, "-m", "pip", "download",
                    "--no-deps",
                    "--disable-pip-version-check",
                    "-d", tmpdir,
                    package_name
                ],
                timeout=TIMEOUT_SEC
            )
            result["raw"]["pip_download"] = dl
            combined = (dl["stdout"] or "") + "\n" + (dl["stderr"] or "")

            if dl["ok"]:
                result["diagnosis_lines"].append("事前ダウンロード診断は成功しました。")
            else:
                result["diagnosis_lines"].append("事前ダウンロード診断で失敗を確認しました。")
                # よくあるエラー解析
                low = combined.lower()
                if "no matching distribution found" in low:
                    result["diagnosis_lines"].append(
                        "現在のPython版・OS・CPU向けの配布物が見つからない可能性があります。"
                    )
                if "requires-python" in low:
                    result["diagnosis_lines"].append(
                        "Python バージョン条件の不一致が疑われます。"
                    )
                if "failed building wheel" in low or "error: subprocess-exited-with-error" in low:
                    result["diagnosis_lines"].append(
                        "wheel の入手に失敗し、ソースビルドが始まって失敗した可能性があります。"
                    )
                if "temporary" in low or "cache" in low or "filenotfounderror" in low:
                    result["diagnosis_lines"].append(
                        "pip キャッシュ/一時ディレクトリ周りの破損が疑われます。"
                    )
        finally:
            try:
                shutil.rmtree(tmpdir, ignore_errors=True)
            except Exception:
                pass

        # 修復候補生成
        actions = []

        def add_action(title, command, description, caution=""):
            actions.append({
                "title": title,
                "command": command,
                "description": description,
                "caution": caution,
            })

        add_action(
            "pip / setuptools / wheel を更新",
            [sys.executable, "-m", "pip", "install", "--upgrade", "pip", "setuptools", "wheel"],
            "基本修復です。依存解決やビルド周りの不具合改善を狙います。",
            "Pydroid では pip 更新が逆に不安定化する場合もあります。"
        )

        add_action(
            "キャッシュを使わず再インストール",
            [sys.executable, "-m", "pip", "install", "--no-cache-dir", package_name],
            "壊れたキャッシュを避けて再取得します。"
        )

        add_action(
            "binary only で試す",
            [sys.executable, "-m", "pip", "install", "--only-binary=:all:", package_name],
            "ソースビルドを禁止し、利用可能な wheel がある場合だけ入れます。",
            "wheel が無ければ失敗します。"
        )

        # PyMuPDF 特別処方
        if package_name.lower() in ("pymupdf", "fitz"):
            add_action(
                "PyMuPDF を no-cache-dir + binary only で試す",
                [sys.executable, "-m", "pip", "install", "--no-cache-dir", "--only-binary=:all:", "PyMuPDF"],
                "Android で無理なら早めに判定できます。",
                "Android/Pydroid では成功しない可能性が高いです。"
            )
            add_action(
                "代替として pypdf をインストール",
                [sys.executable, "-m", "pip", "install", "--no-cache-dir", "pypdf"],
                "PDFの結合・分割・テキスト取得など一部用途を代替できます。",
                "PyMuPDF と完全互換ではありません。"
            )
            add_action(
                "代替として pdfplumber + pdfminer.six をインストール",
                [sys.executable, "-m", "pip", "install", "--no-cache-dir", "pdfplumber", "pdfminer.six"],
                "レイアウト寄りのテキスト抽出代替です。",
                "画像描画や高速レンダリング用途は別です。"
            )

        # pure python 互換版候補
        pure_versions = find_pure_python_versions(releases)
        if pure_versions:
            best_pure = pure_versions[0]
            add_action(
                f"pure Python wheel 版を試す: {best_pure}",
                [sys.executable, "-m", "pip", "install", "--no-cache-dir", f"{package_name}=={best_pure}"],
                "pure Python wheel が存在する過去版を優先します。"
            )

        # cache purge
        add_action(
            "pip キャッシュ削除",
            [sys.executable, "-m", "pip", "cache", "purge"],
            "pip キャッシュ破損対策です。"
        )

        # Android で native only の場合の注意候補
        if (env["android"] or env["pydroid"]) and native_pkg:
            add_action(
                "診断メモのみ表示（Android非対応の可能性）",
                ["__NOOP__"],
                "このパッケージは Android/Pydroid で wheel が不足している可能性が高く、pip 単独修復が困難です。代替ライブラリや別実行環境を検討してください。"
            )

        result["repair_actions"] = actions

    except urllib.error.HTTPError as e:
        result["diagnosis_lines"].append(f"PyPIにパッケージが見つかりませんでした: HTTP {e.code}")
        result["repair_actions"] = [
            {
                "title": "パッケージ名を再確認",
                "command": ["__NOOP__"],
                "description": "入力したライブラリ名のスペルミスの可能性があります。",
                "caution": "",
            }
        ]
    except Exception as e:
        result["diagnosis_lines"].append(f"PyPI問い合わせに失敗しました: {repr(e)}")
        result["repair_actions"] = [
            {
                "title": "ネットワーク確認",
                "command": ["__NOOP__"],
                "description": "インターネット接続や PyPI へのアクセスを確認してください。",
                "caution": "",
            },
            {
                "title": "キャッシュを使わず直接インストール",
                "command": [sys.executable, "-m", "pip", "install", "--no-cache-dir", package_name],
                "description": "まず直接再取得を試します。",
                "caution": "",
            }
        ]

    return result


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1100x780")

        self.q = queue.Queue()
        self.current_result = None
        self.repair_actions = []

        self._build_ui()
        self.after(150, self._poll_queue)

    def _build_ui(self):
        root = ttk.Frame(self, padding=8)
        root.pack(fill="both", expand=True)

        top = ttk.LabelFrame(root, text="対象ライブラリ")
        top.pack(fill="x", padx=4, pady=4)

        ttk.Label(top, text="ライブラリ名:").pack(side="left", padx=6, pady=6)

        self.pkg_var = tk.StringVar(value="PyMuPDF")
        self.pkg_entry = ttk.Entry(top, textvariable=self.pkg_var)
        self.pkg_entry.pack(side="left", fill="x", expand=True, padx=6, pady=6)

        ttk.Button(top, text="診断開始", command=self.on_diagnose).pack(side="left", padx=6, pady=6)
        ttk.Button(top, text="ログ保存", command=self.on_save_log).pack(side="left", padx=6, pady=6)
        ttk.Button(top, text="終了", command=self.destroy).pack(side="left", padx=6, pady=6)

        mid = ttk.Panedwindow(root, orient="horizontal")
        mid.pack(fill="both", expand=True, padx=4, pady=4)

        left = ttk.Frame(mid)
        right = ttk.Frame(mid)
        mid.add(left, weight=3)
        mid.add(right, weight=2)

        # 左: 診断結果
        lf1 = ttk.LabelFrame(left, text="診断結果")
        lf1.pack(fill="both", expand=True, padx=4, pady=4)

        self.diag_text = tk.Text(lf1, wrap="word")
        self.diag_text.pack(side="left", fill="both", expand=True)
        sy1 = ttk.Scrollbar(lf1, orient="vertical", command=self.diag_text.yview)
        sy1.pack(side="right", fill="y")
        self.diag_text.config(yscrollcommand=sy1.set)

        # 右上: 修復候補
        rf1 = ttk.LabelFrame(right, text="修復候補")
        rf1.pack(fill="both", expand=True, padx=4, pady=4)

        self.repair_list = tk.Listbox(rf1)
        self.repair_list.pack(side="left", fill="both", expand=True)
        sy2 = ttk.Scrollbar(rf1, orient="vertical", command=self.repair_list.yview)
        sy2.pack(side="right", fill="y")
        self.repair_list.config(yscrollcommand=sy2.set)

        btnf = ttk.Frame(right)
        btnf.pack(fill="x", padx=4, pady=4)

        ttk.Button(btnf, text="選択修復を実行", command=self.on_run_repair).pack(side="left", padx=4, pady=4)
        ttk.Button(btnf, text="候補詳細表示", command=self.on_show_action_detail).pack(side="left", padx=4, pady=4)

        # 下: 実行ログ
        lf2 = ttk.LabelFrame(root, text="実行ログ")
        lf2.pack(fill="both", expand=True, padx=4, pady=4)

        self.log_text = tk.Text(lf2, wrap="word", height=14)
        self.log_text.pack(side="left", fill="both", expand=True)
        sy3 = ttk.Scrollbar(lf2, orient="vertical", command=self.log_text.yview)
        sy3.pack(side="right", fill="y")
        self.log_text.config(yscrollcommand=sy3.set)

        self._log("アプリ起動完了。ライブラリ名を入れて「診断開始」を押してください。")

    def _log(self, msg):
        self.log_text.insert("end", msg + "\n")
        self.log_text.see("end")

    def _set_diag(self, text):
        self.diag_text.delete("1.0", "end")
        self.diag_text.insert("1.0", text)

    def _poll_queue(self):
        try:
            while True:
                kind, payload = self.q.get_nowait()
                if kind == "log":
                    self._log(payload)
                elif kind == "diag_done":
                    self.current_result = payload
                    self.show_result(payload)
                elif kind == "repair_done":
                    self._log(payload)
                elif kind == "error":
                    self._log("[ERROR] " + payload)
                    messagebox.showerror("エラー", payload)
        except queue.Empty:
            pass
        self.after(150, self._poll_queue)

    def on_diagnose(self):
        pkg = self.pkg_var.get().strip()
        if not pkg:
            messagebox.showwarning("入力不足", "ライブラリ名を入力してください。")
            return

        self._log(f"[{now_str()}] 診断要求: {pkg}")
        self._set_diag("診断中です...")
        self.repair_list.delete(0, "end")
        self.repair_actions = []

        def worker():
            try:
                res = diagnose_package(pkg, ui_log=lambda m: self.q.put(("log", m)))
                self.q.put(("diag_done", res))
            except Exception as e:
                self.q.put(("error", repr(e)))

        threading.Thread(target=worker, daemon=True).start()

    def show_result(self, res):
        env = res.get("env", {})
        lines = []
        lines.append("=== 環境情報 ===")
        lines.append(f"Python        : {env.get('python_mm')} / {env.get('python')}")
        lines.append(f"実行ファイル  : {env.get('executable')}")
        lines.append(f"OS            : {env.get('platform')}")
        lines.append(f"CPU           : {env.get('machine')}")
        lines.append(f"Android       : {env.get('android')}")
        lines.append(f"Pydroid       : {env.get('pydroid')}")
        lines.append("")
        lines.append("=== 診断結果 ===")
        for s in res.get("diagnosis_lines", []):
            lines.append(f"- {s}")

        self._set_diag("\n".join(lines))

        self.repair_actions = res.get("repair_actions", [])
        self.repair_list.delete(0, "end")
        for i, act in enumerate(self.repair_actions, start=1):
            self.repair_list.insert("end", f"{i}. {act['title']}")

        self._log(f"[{now_str()}] 診断完了。修復候補 {len(self.repair_actions)} 件。")

    def on_show_action_detail(self):
        idxs = self.repair_list.curselection()
        if not idxs:
            messagebox.showinfo("未選択", "修復候補を選択してください。")
            return
        idx = idxs[0]
        act = self.repair_actions[idx]
        msg = (
            f"タイトル:\n{act['title']}\n\n"
            f"説明:\n{act.get('description', '')}\n\n"
            f"コマンド:\n{' '.join(act.get('command', []))}\n\n"
            f"注意:\n{act.get('caution', '')}"
        )
        messagebox.showinfo("修復候補の詳細", msg)

    def on_run_repair(self):
        idxs = self.repair_list.curselection()
        if not idxs:
            messagebox.showwarning("未選択", "修復候補を選択してください。")
            return

        idx = idxs[0]
        act = self.repair_actions[idx]
        cmd = act.get("command", [])

        if not messagebox.askyesno("確認", f"次の修復を実行しますか？\n\n{act['title']}"):
            return

        def worker():
            self.q.put(("log", f"[{now_str()}] 修復開始: {act['title']}"))
            self.q.put(("log", f"COMMAND: {' '.join(cmd)}"))

            if cmd == ["__NOOP__"]:
                self.q.put(("repair_done", "この候補は実行コマンドを持たない診断メモです。"))
                return

            res = run_cmd(cmd, timeout=TIMEOUT_SEC)
            self.q.put(("log", "----- STDOUT -----"))
            self.q.put(("log", res.get("stdout", "")))
            self.q.put(("log", "----- STDERR -----"))
            self.q.put(("log", res.get("stderr", "")))
            self.q.put(("log", "------------------"))

            if res["ok"]:
                self.q.put(("repair_done", f"[{now_str()}] 修復成功: {act['title']}"))
            else:
                self.q.put(("repair_done", f"[{now_str()}] 修復失敗: {act['title']} (code={res['returncode']})"))

        threading.Thread(target=worker, daemon=True).start()

    def on_save_log(self):
        path = filedialog.asksaveasfilename(
            title="ログ保存",
            defaultextension=".txt",
            filetypes=[("Text", "*.txt"), ("All", "*.*")]
        )
        if not path:
            return
        try:
            with open(path, "w", encoding="utf-8") as f:
                f.write("=== 診断結果 ===\n")
                f.write(self.diag_text.get("1.0", "end"))
                f.write("\n=== 実行ログ ===\n")
                f.write(self.log_text.get("1.0", "end"))
            messagebox.showinfo("保存完了", f"保存しました:\n{path}")
        except Exception as e:
            messagebox.showerror("保存失敗", repr(e))


if __name__ == "__main__":
    app = App()
    app.mainloop()