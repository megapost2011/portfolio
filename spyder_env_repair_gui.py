# -*- coding: utf-8 -*-
"""
Created on Fri Apr 24 17:30:59 2026

@author: Administrator
"""

# -*- coding: utf-8 -*-
"""
Spyder環境 自動診断・修復GUI
Windows / Spyder6 対応
"""

import os
import sys
import subprocess
import threading
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext


class SpyderRepairGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Spyder環境 自動診断・修復GUI")
        self.root.geometry("900x650")

        self.python_path = tk.StringVar()
        self.package_name = tk.StringVar(value="requests")

        self.build_ui()

    def build_ui(self):
        frame = ttk.Frame(self.root, padding=10)
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text="Spyder Python パス").pack(anchor="w")

        path_frame = ttk.Frame(frame)
        path_frame.pack(fill="x", pady=5)

        self.path_entry = ttk.Entry(path_frame, textvariable=self.python_path)
        self.path_entry.pack(side="left", fill="x", expand=True)

        ttk.Button(
            path_frame,
            text="自動検出",
            command=self.detect_spyder_python
        ).pack(side="left", padx=5)

        ttk.Button(
            path_frame,
            text="現在のPythonを使用",
            command=self.use_current_python
        ).pack(side="left", padx=5)

        ttk.Label(frame, text="インストールしたいライブラリ名").pack(anchor="w", pady=(10, 0))

        pkg_frame = ttk.Frame(frame)
        pkg_frame.pack(fill="x", pady=5)

        ttk.Entry(pkg_frame, textvariable=self.package_name).pack(side="left", fill="x", expand=True)

        ttk.Button(
            pkg_frame,
            text="ライブラリをインストール",
            command=self.install_package
        ).pack(side="left", padx=5)

        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill="x", pady=10)

        ttk.Button(
            btn_frame,
            text="① 診断",
            command=self.diagnose
        ).pack(side="left", padx=5)

        ttk.Button(
            btn_frame,
            text="② pip復旧",
            command=self.repair_pip
        ).pack(side="left", padx=5)

        ttk.Button(
            btn_frame,
            text="③ pip更新",
            command=self.upgrade_pip_tools
        ).pack(side="left", padx=5)

        ttk.Button(
            btn_frame,
            text="④ requests導入テスト",
            command=self.install_requests_test
        ).pack(side="left", padx=5)

        ttk.Button(
            btn_frame,
            text="⑤ 全自動修復",
            command=self.full_repair
        ).pack(side="left", padx=5)

        ttk.Button(
            btn_frame,
            text="ログ消去",
            command=lambda: self.log_box.delete("1.0", tk.END)
        ).pack(side="right", padx=5)

        self.log_box = scrolledtext.ScrolledText(frame, wrap=tk.WORD, font=("Consolas", 10))
        self.log_box.pack(fill="both", expand=True)

        self.log("Spyder環境 自動診断・修復GUI 起動完了")
        self.log("まずは「自動検出」または「現在のPythonを使用」を押してください。")

    def log(self, text):
        self.log_box.insert(tk.END, text + "\n")
        self.log_box.see(tk.END)
        self.root.update_idletasks()

    def run_thread(self, func):
        threading.Thread(target=func, daemon=True).start()

    def run_cmd(self, args):
        self.log("")
        self.log("実行コマンド:")
        self.log(" ".join(f'"{a}"' if " " in a else a for a in args))
        self.log("-" * 80)

        try:
            proc = subprocess.Popen(
                args,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace"
            )

            for line in proc.stdout:
                self.log(line.rstrip())

            proc.wait()
            self.log("-" * 80)
            self.log(f"終了コード: {proc.returncode}")

            return proc.returncode == 0

        except Exception as e:
            self.log(f"実行エラー: {e}")
            return False

    def get_python(self):
        py = self.python_path.get().strip().strip('"')
        if not py:
            messagebox.showwarning("未指定", "Pythonパスが未指定です。")
            return None
        if not os.path.exists(py):
            messagebox.showerror("エラー", f"Pythonが見つかりません:\n{py}")
            return None
        return py

    def detect_spyder_python(self):
        candidates = [
            r"C:\ProgramData\spyder-6\envs\spyder-runtime\python.exe",
            os.path.expanduser(r"~\AppData\Local\Programs\Spyder\Python\python.exe"),
        ]

        for c in candidates:
            if os.path.exists(c):
                self.python_path.set(c)
                self.log(f"検出成功: {c}")
                return

        self.log("Spyder6 runtime python.exe を自動検出できませんでした。")
        self.log("Spyderコンソールで以下を実行して、表示されたパスを貼り付けてください。")
        self.log("import sys")
        self.log("print(sys.executable)")

    def use_current_python(self):
        self.python_path.set(sys.executable)
        self.log(f"現在このGUIを実行中のPythonを設定しました: {sys.executable}")

    def diagnose(self):
        def task():
            py = self.get_python()
            if not py:
                return

            self.log("")
            self.log("========== 診断開始 ==========")
            self.log(f"対象Python: {py}")

            if "ProgramData" in py:
                self.log("注意: ProgramData配下です。権限エラーが出る場合は管理者として実行してください。")

            self.run_cmd([py, "--version"])
            self.run_cmd([py, "-c", "import sys; print(sys.executable); print(sys.version)"])

            self.log("")
            self.log("pip確認:")
            ok = self.run_cmd([py, "-m", "pip", "--version"])

            if not ok:
                self.log("pipが存在しない可能性があります。『pip復旧』を実行してください。")
            else:
                self.log("pipは使用可能です。")

            self.log("========== 診断終了 ==========")

        self.run_thread(task)

    def repair_pip(self):
        def task():
            py = self.get_python()
            if not py:
                return

            self.log("")
            self.log("========== pip復旧開始 ==========")
            ok = self.run_cmd([py, "-m", "ensurepip", "--upgrade"])

            if ok:
                self.log("pip復旧に成功しました。")
            else:
                self.log("pip復旧に失敗しました。管理者として起動して再実行してください。")

            self.log("========== pip復旧終了 ==========")

        self.run_thread(task)

    def upgrade_pip_tools(self):
        def task():
            py = self.get_python()
            if not py:
                return

            self.log("")
            self.log("========== pip / setuptools / wheel 更新開始 ==========")

            ok = self.run_cmd([
                py, "-m", "pip", "install",
                "--upgrade", "pip", "setuptools", "wheel"
            ])

            if ok:
                self.log("基本ツール更新に成功しました。")
            else:
                self.log("基本ツール更新に失敗しました。")

            self.log("========== 更新終了 ==========")

        self.run_thread(task)

    def install_package(self):
        def task():
            py = self.get_python()
            if not py:
                return

            pkg = self.package_name.get().strip()
            if not pkg:
                messagebox.showwarning("未入力", "ライブラリ名を入力してください。")
                return

            self.log("")
            self.log(f"========== {pkg} インストール開始 ==========")

            ok = self.run_cmd([py, "-m", "pip", "install", pkg])

            if ok:
                self.log(f"{pkg} のインストールに成功しました。")
                import_name = pkg.split("[")[0].replace("-", "_")
                self.log("import確認を行います。")
                self.run_cmd([py, "-c", f"import {import_name}; print('{import_name} import OK')"])
            else:
                self.log(f"{pkg} のインストールに失敗しました。")

            self.log("========== インストール終了 ==========")

        self.run_thread(task)

    def install_requests_test(self):
        self.package_name.set("requests")
        self.install_package()

    def full_repair(self):
        def task():
            py = self.get_python()
            if not py:
                return

            self.log("")
            self.log("========== 全自動修復開始 ==========")
            self.log("1. Python診断")
            self.run_cmd([py, "--version"])
            self.run_cmd([py, "-c", "import sys; print(sys.executable); print(sys.version)"])

            self.log("")
            self.log("2. pip復旧")
            self.run_cmd([py, "-m", "ensurepip", "--upgrade"])

            self.log("")
            self.log("3. pip / setuptools / wheel 更新")
            self.run_cmd([
                py, "-m", "pip", "install",
                "--upgrade", "pip", "setuptools", "wheel"
            ])

            self.log("")
            self.log("4. requests インストールテスト")
            self.run_cmd([py, "-m", "pip", "install", "requests"])

            self.log("")
            self.log("5. import確認")
            self.run_cmd([py, "-c", "import requests; print('requests import OK'); print(requests.__version__)"])

            self.log("")
            self.log("========== 全自動修復終了 ==========")
            self.log("Spyderを再起動してから import requests を試してください。")

        self.run_thread(task)


if __name__ == "__main__":
    root = tk.Tk()
    app = SpyderRepairGUI(root)
    root.mainloop()