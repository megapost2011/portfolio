import os
import sys
import site
import json
import shutil
import queue
import ctypes
import threading
import tempfile
import subprocess
import platform
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

APP_TITLE = "Spyder / pip トラブル診断アプリ"
APP_VERSION = "1.0.0"

# -----------------------------
# 共通ユーティリティ
# -----------------------------
def safe_getenv(key, default=""):
    try:
        return os.environ.get(key, default)
    except Exception:
        return default

def is_windows():
    return os.name == "nt"

def is_admin():
    if not is_windows():
        try:
            return os.geteuid() == 0
        except Exception:
            return False
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False

def norm(p):
    try:
        return os.path.normpath(p) if p else ""
    except Exception:
        return p or ""

def file_exists(p):
    try:
        return bool(p) and os.path.exists(p)
    except Exception:
        return False

def join_lines(lines):
    return "\n".join(str(x) for x in lines if x is not None)

def short_exc(e):
    return f"{type(e).__name__}: {e}"

def which_all(name):
    results = []
    path = safe_getenv("PATH", "")
    for d in path.split(os.pathsep):
        d = d.strip('"').strip()
        if not d:
            continue
        cand = os.path.join(d, name)
        if os.path.isfile(cand):
            results.append(norm(cand))
        if is_windows():
            for ext in [".exe", ".bat", ".cmd"]:
                cand2 = cand + ext
                if os.path.isfile(cand2):
                    results.append(norm(cand2))
    # 重複排除
    seen = set()
    uniq = []
    for r in results:
        if r not in seen:
            seen.add(r)
            uniq.append(r)
    return uniq

def run_command(cmd, timeout=60, cwd=None):
    """
    subprocess 実行。出力を返す。
    """
    try:
        cp = subprocess.run(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=timeout,
            cwd=cwd,
            shell=False
        )
        return {
            "ok": cp.returncode == 0,
            "returncode": cp.returncode,
            "stdout": cp.stdout,
            "stderr": cp.stderr,
            "cmd": cmd,
        }
    except subprocess.TimeoutExpired as e:
        return {
            "ok": False,
            "returncode": -999,
            "stdout": e.stdout or "",
            "stderr": f"TimeoutExpired: {e}",
            "cmd": cmd,
        }
    except Exception as e:
        return {
            "ok": False,
            "returncode": -998,
            "stdout": "",
            "stderr": short_exc(e),
            "cmd": cmd,
        }

def run_shell_capture(command_string, timeout=30):
    """
    where / where.exe など shell 依存も使いたい時用
    """
    try:
        cp = subprocess.run(
            command_string,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=timeout,
            shell=True
        )
        return {
            "ok": cp.returncode == 0,
            "returncode": cp.returncode,
            "stdout": cp.stdout,
            "stderr": cp.stderr,
            "cmd": command_string,
        }
    except Exception as e:
        return {
            "ok": False,
            "returncode": -998,
            "stdout": "",
            "stderr": short_exc(e),
            "cmd": command_string,
        }

def parse_bool_env(v):
    if not v:
        return False
    return str(v).strip().lower() in ("1", "true", "yes", "on")

# -----------------------------
# Python / pip / Spyder 診断
# -----------------------------
class DiagnosticEngine:
    def __init__(self, log_callback=None):
        self.log_callback = log_callback

    def log(self, text=""):
        if self.log_callback:
            self.log_callback(text)

    def python_info(self, python_exe):
        code = r"""
import sys, os, site, json, platform
data = {
    "executable": sys.executable,
    "version": sys.version,
    "prefix": sys.prefix,
    "base_prefix": getattr(sys, "base_prefix", ""),
    "exec_prefix": sys.exec_prefix,
    "platform": platform.platform(),
    "python_implementation": platform.python_implementation(),
    "venv_active": (getattr(sys, "base_prefix", sys.prefix) != sys.prefix),
    "path": sys.path,
    "usersite": getattr(site, "USER_SITE", ""),
    "sitepackages": getattr(site, "getsitepackages", lambda: [])(),
}
print(json.dumps(data, ensure_ascii=False, indent=2))
"""
        return run_command([python_exe, "-c", code], timeout=30)

    def pip_version(self, python_exe):
        return run_command([python_exe, "-m", "pip", "--version"], timeout=30)

    def pip_debug(self, python_exe):
        return run_command([python_exe, "-m", "pip", "debug", "--verbose"], timeout=60)

    def pip_show(self, python_exe, package_name):
        return run_command([python_exe, "-m", "pip", "show", package_name], timeout=30)

    def pip_index_versions(self, python_exe, package_name):
        return run_command([python_exe, "-m", "pip", "index", "versions", package_name], timeout=60)

    def pip_install_drylike(self, python_exe, package_name, user_install=False):
        """
        pip --dry-run は pip バージョン差があるため、
        まず --dry-run を試し、ダメなら download ベースへフォールバック。
        """
        base = [python_exe, "-m", "pip", "install", package_name, "--disable-pip-version-check", "--no-input", "-v", "--dry-run"]
        if user_install:
            base.append("--user")
        res = run_command(base, timeout=120)
        if res["returncode"] not in (-998, -999) and ("no such option: --dry-run" not in res["stderr"].lower()):
            return res

        temp_dir = tempfile.mkdtemp(prefix="pipdiag_")
        cmd = [python_exe, "-m", "pip", "download", package_name, "--disable-pip-version-check", "--no-input", "-v", "-d", temp_dir]
        res2 = run_command(cmd, timeout=180)
        res2["stderr"] += f"\n\n[INFO] --dry-run 非対応のため download 診断へフォールバック: {temp_dir}"
        return res2

    def pip_install_real(self, python_exe, package_name, user_install=False, no_cache=False):
        cmd = [python_exe, "-m", "pip", "install", package_name, "--disable-pip-version-check", "--no-input", "-v"]
        if user_install:
            cmd.append("--user")
        if no_cache:
            cmd.append("--no-cache-dir")
        return run_command(cmd, timeout=300)

    def detect_spyder(self):
        results = {
            "which_spyder": [],
            "which_spyder3": [],
            "start_menu_guess": [],
            "module_check": None,
        }

        results["which_spyder"] = which_all("spyder")
        results["which_spyder3"] = which_all("spyder3")

        # Windows の典型的な場所をざっくり探索
        if is_windows():
            guesses = []
            for root in [
                os.path.expandvars(r"%LOCALAPPDATA%\Programs\Spyder"),
                os.path.expandvars(r"%ProgramFiles%\Spyder"),
                os.path.expandvars(r"%ProgramFiles(x86)%\Spyder"),
                os.path.expandvars(r"%USERPROFILE%\AppData\Local\Programs\Spyder"),
            ]:
                if os.path.isdir(root):
                    guesses.append(norm(root))
            results["start_menu_guess"] = guesses

        results["module_check"] = run_command([sys.executable, "-m", "pip", "show", "spyder"], timeout=30)
        return results

    def filesystem_checks(self, python_exe):
        checks = []

        py_dir = norm(os.path.dirname(python_exe))
        for path in [py_dir, safe_getenv("TEMP"), safe_getenv("TMP"), site.USER_SITE]:
            if not path:
                continue
            writable = None
            err = ""
            try:
                os.makedirs(path, exist_ok=True)
                testfile = os.path.join(path, f"pipdiag_write_test_{os.getpid()}.tmp")
                with open(testfile, "w", encoding="utf-8") as f:
                    f.write("ok")
                os.remove(testfile)
                writable = True
            except Exception as e:
                writable = False
                err = short_exc(e)
            checks.append({
                "path": norm(path),
                "writable": writable,
                "error": err,
            })
        return checks

    def collect_environment(self):
        env_keys = [
            "PATH", "PYTHONPATH", "PYTHONHOME",
            "PIP_INDEX_URL", "PIP_EXTRA_INDEX_URL", "PIP_TRUSTED_HOST",
            "HTTP_PROXY", "HTTPS_PROXY", "NO_PROXY",
            "CONDA_PREFIX", "CONDA_DEFAULT_ENV",
            "VIRTUAL_ENV",
        ]
        return {k: safe_getenv(k, "") for k in env_keys}

    def guess_causes(self, collected):
        causes = []
        actions = []

        py_info = collected.get("py_info_json", {})
        pip_ver_text = collected.get("pip_version_text", "")
        pip_cmd_path = collected.get("pip_path", "")
        python_exe = collected.get("python_exe", "")
        env = collected.get("env", {})
        fs = collected.get("fs_checks", [])
        pkg = collected.get("package_name", "")

        executable = norm(py_info.get("executable", ""))
        usersite = norm(py_info.get("usersite", ""))
        sitepackages = [norm(x) for x in py_info.get("sitepackages", [])]

        # 1. pip と python の不一致
        if pip_cmd_path and python_exe:
            pip_cmd_path_l = pip_cmd_path.lower()
            python_dir_l = norm(os.path.dirname(python_exe)).lower()
            if python_dir_l not in pip_cmd_path_l:
                causes.append("`pip` コマンドが、現在診断対象の Python と別環境のものを見ている可能性があります。")
                actions.append("`pip install ...` ではなく、必ず `python -m pip install ...` を使ってください。")

        if executable and python_exe and norm(executable).lower() != norm(python_exe).lower():
            causes.append("選択した Python と、実際に応答した Python 実体が一致していません。")
            actions.append("Spyder のインタプリタ設定を見直し、`python -c \"import sys; print(sys.executable)\"` で確認してください。")

        # 2. 管理者権限 / 書込権限
        unwritable = [x for x in fs if x.get("writable") is False]
        if unwritable:
            causes.append("Python / TEMP / ユーザーサイトの一部に書込権限がありません。")
            actions.append("管理者権限で起動するか、`--user` インストール、または書込可能な仮想環境を利用してください。")

        # 3. venv / conda かどうか
        if env.get("CONDA_PREFIX"):
            actions.append("Conda 環境なら、まず `conda install パッケージ名` を優先し、それで難しい場合に `python -m pip` を使ってください。")
        elif py_info.get("venv_active"):
            actions.append("仮想環境が有効です。Spyder 側でも同じ仮想環境の Python を指定してください。")
        else:
            causes.append("仮想環境や conda 環境ではなく、システム Python 直下に入れようとしている可能性があります。")
            actions.append("新規に venv または conda 環境を作ってからインストールするのが安全です。")

        # 4. 環境変数汚染
        if env.get("PYTHONPATH"):
            causes.append("`PYTHONPATH` が設定されており、依存解決や import 先が汚染されている可能性があります。")
            actions.append("一度 `PYTHONPATH` を外して再試行してください。")
        if env.get("PYTHONHOME"):
            causes.append("`PYTHONHOME` が設定されており、Python 実行環境が不安定になっている可能性があります。")
            actions.append("通常は `PYTHONHOME` は未設定推奨です。")

        # 5. プロキシ / ミラー問題
        if env.get("HTTP_PROXY") or env.get("HTTPS_PROXY"):
            causes.append("プロキシ設定が有効です。企業内ネットワークや不正なプロキシ設定で pip 通信が失敗することがあります。")
            actions.append("必要ならプロキシ設定を見直し、不要なら一時的に解除してください。")
        if env.get("PIP_INDEX_URL"):
            causes.append("`PIP_INDEX_URL` が設定されており、PyPI ではなく別インデックスを見ている可能性があります。")
            actions.append("ミラー先が正しいか確認してください。")

        # 6. site-packages
        if usersite and not any(usersite.lower() == sp.lower() for sp in sitepackages):
            actions.append("`--user` で入れたパッケージが現在の実行系から見えない場合があります。Spyder の interpreter を合わせてください。")

        # 7. 典型的メッセージ解析
        dry_stdout = collected.get("dry_stdout", "")
        dry_stderr = collected.get("dry_stderr", "")
        real_stdout = collected.get("real_stdout", "")
        real_stderr = collected.get("real_stderr", "")
        blob = "\n".join([dry_stdout, dry_stderr, real_stdout, real_stderr]).lower()

        patterns = [
            ("no matching distribution found", "対象 Python のバージョンや OS/アーキテクチャに対応する配布物がありません。", "Python バージョンを変えるか、対応版のパッケージ名を確認してください。"),
            ("could not find a version that satisfies the requirement", "要求バージョンが存在しない、または利用可能なインデックスにありません。", "パッケージ名の綴り・バージョン指定・インデックス先を確認してください。"),
            ("ssl", "SSL 証明書や TLS 通信まわりの問題の可能性があります。", "企業プロキシ・証明書・セキュリティソフトの HTTPS 介入を確認してください。"),
            ("permission denied", "権限不足で書き込めていません。", "`--user` か仮想環境を使うか、権限を見直してください。"),
            ("access is denied", "Windows のアクセス拒否です。", "管理者権限、セキュリティソフト、使用中ファイルを確認してください。"),
            ("externally-managed-environment", "外部管理環境として pip 書込が制限されています。", "仮想環境を作成してそこへインストールしてください。"),
            ("failed building wheel", "ビルドに失敗しています。C/C++ コンパイラやヘッダ不足の可能性があります。", "wheel がある版を選ぶ、Visual C++ Build Tools、Python 対応版を確認してください。"),
            ("microsoft visual c++", "Windows ビルドツール不足の可能性があります。", "Visual C++ Build Tools を導入してください。"),
            ("rust compiler", "Rust コンパイラが必要な依存です。", "prebuilt wheel があるバージョンを使うか、Rust を導入してください。"),
            ("subprocess-exited-with-error", "ビルドまたはセットアップサブプロセスで失敗しています。", "エラー直前のログを見て、ビルド依存を追加してください。"),
            ("proxyerror", "プロキシ経由通信に失敗しています。", "プロキシ設定を見直してください。"),
            ("read timed out", "ネットワークタイムアウトです。", "回線・ミラー・社内FW・再試行を確認してください。"),
            ("winerror 5", "Windows の権限エラーです。", "権限・ロック・セキュリティソフトを確認してください。"),
            ("requires-python", "そのパッケージが現在の Python バージョンに非対応です。", "対応する Python 版へ切り替えてください。"),
        ]
        for key, cause, action in patterns:
            if key in blob:
                causes.append(cause)
                actions.append(action)

        if pkg:
            actions.append(f"`python -m pip show {pkg}` で現在環境に入っているか再確認してください。")
            actions.append(f"Spyder のコンソールで `import {pkg.split('[')[0].split('=')[0].replace('-', '_')}` を試し、import 可否を確認してください。")

        # 重複排除
        causes = list(dict.fromkeys(causes))
        actions = list(dict.fromkeys(actions))
        return causes, actions

# -----------------------------
# GUI
# -----------------------------
class App:
    def __init__(self, root):
        self.root = root
        self.root.title(f"{APP_TITLE} v{APP_VERSION}")
        self.root.geometry("1200x760")

        self.log_queue = queue.Queue()
        self.engine = DiagnosticEngine(log_callback=self.enqueue_log)

        self.python_var = tk.StringVar(value=sys.executable)
        self.package_var = tk.StringVar(value="numpy")
        self.user_install_var = tk.BooleanVar(value=False)
        self.no_cache_var = tk.BooleanVar(value=False)
        self.auto_real_install_var = tk.BooleanVar(value=False)

        self.build_ui()
        self.root.after(100, self.process_log_queue)

    def build_ui(self):
        top = ttk.Frame(self.root, padding=8)
        top.pack(fill="x")

        ttk.Label(top, text="対象 Python 実行ファイル").grid(row=0, column=0, sticky="w")
        ttk.Entry(top, textvariable=self.python_var, width=90).grid(row=0, column=1, sticky="ew", padx=5)
        ttk.Button(top, text="参照", command=self.browse_python).grid(row=0, column=2, padx=2)
        ttk.Button(top, text="自動検出", command=self.auto_detect_python).grid(row=0, column=3, padx=2)

        ttk.Label(top, text="診断したいパッケージ").grid(row=1, column=0, sticky="w", pady=(8, 0))
        ttk.Entry(top, textvariable=self.package_var, width=40).grid(row=1, column=1, sticky="w", padx=5, pady=(8, 0))
        ttk.Checkbutton(top, text="--user で試す", variable=self.user_install_var).grid(row=1, column=2, sticky="w", pady=(8, 0))
        ttk.Checkbutton(top, text="--no-cache-dir を付与", variable=self.no_cache_var).grid(row=1, column=3, sticky="w", pady=(8, 0))

        ttk.Checkbutton(top, text="簡易診断後に実インストール試験も行う", variable=self.auto_real_install_var).grid(
            row=2, column=1, sticky="w", pady=(8, 0)
        )

        btns = ttk.Frame(self.root, padding=(8, 0))
        btns.pack(fill="x")
        ttk.Button(btns, text="基本診断", command=self.run_basic_diag).pack(side="left", padx=3, pady=6)
        ttk.Button(btns, text="パッケージ診断", command=self.run_package_diag).pack(side="left", padx=3, pady=6)
        ttk.Button(btns, text="実インストール試験", command=self.run_real_install_test).pack(side="left", padx=3, pady=6)
        ttk.Button(btns, text="Spyder検出", command=self.run_spyder_detect).pack(side="left", padx=3, pady=6)
        ttk.Button(btns, text="ログ保存", command=self.save_log).pack(side="left", padx=3, pady=6)
        ttk.Button(btns, text="クリア", command=self.clear_log).pack(side="left", padx=3, pady=6)

        # Notebook
        nb = ttk.Notebook(self.root)
        nb.pack(fill="both", expand=True, padx=8, pady=8)

        self.tab_log = ttk.Frame(nb)
        self.tab_summary = ttk.Frame(nb)
        nb.add(self.tab_log, text="詳細ログ")
        nb.add(self.tab_summary, text="診断サマリー")

        # 詳細ログタブ
        log_frame = ttk.Frame(self.tab_log)
        log_frame.pack(fill="both", expand=True)

        self.log_text = tk.Text(log_frame, wrap="none", undo=False)
        self.log_text.pack(side="left", fill="both", expand=True)

        ysb = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        ysb.pack(side="right", fill="y")
        xsb = ttk.Scrollbar(self.tab_log, orient="horizontal", command=self.log_text.xview)
        xsb.pack(fill="x")

        self.log_text.configure(yscrollcommand=ysb.set, xscrollcommand=xsb.set)

        # サマリータブ
        summary_top = ttk.Frame(self.tab_summary, padding=8)
        summary_top.pack(fill="both", expand=True)

        ttk.Label(summary_top, text="原因候補").pack(anchor="w")
        self.causes_text = tk.Text(summary_top, height=12, wrap="word")
        self.causes_text.pack(fill="x", expand=False, pady=(0, 8))

        ttk.Label(summary_top, text="推奨対処").pack(anchor="w")
        self.actions_text = tk.Text(summary_top, height=16, wrap="word")
        self.actions_text.pack(fill="both", expand=True)

        top.columnconfigure(1, weight=1)

    def browse_python(self):
        path = filedialog.askopenfilename(
            title="python.exe を選択",
            filetypes=[("Python", "*.exe" if is_windows() else "*"), ("All files", "*.*")]
        )
        if path:
            self.python_var.set(path)

    def auto_detect_python(self):
        candidates = []
        candidates.append(sys.executable)

        # PATH 上の python
        for name in ("python", "python3", "py"):
            for p in which_all(name):
                candidates.append(p)

        # Windows 典型パス
        if is_windows():
            guess_dirs = [
                os.path.expandvars(r"%LOCALAPPDATA%\Programs\Python"),
                os.path.expandvars(r"%ProgramFiles%\Python"),
                os.path.expandvars(r"%USERPROFILE%\AppData\Local\Microsoft\WindowsApps"),
                os.path.expandvars(r"%USERPROFILE%\miniconda3"),
                os.path.expandvars(r"%USERPROFILE%\anaconda3"),
            ]
            for gd in guess_dirs:
                if os.path.isdir(gd):
                    for root, dirs, files in os.walk(gd):
                        if "python.exe" in [f.lower() for f in files]:
                            candidates.append(os.path.join(root, "python.exe"))

        # 重複排除
        seen = set()
        uniq = []
        for c in candidates:
            c = norm(c)
            if c and c not in seen and os.path.exists(c):
                seen.add(c)
                uniq.append(c)

        if not uniq:
            messagebox.showwarning("自動検出", "Python を検出できませんでした。")
            return

        # 一番上に sys.executable を優先
        self.python_var.set(uniq[0])

        self.enqueue_log("[自動検出された Python 候補]")
        for c in uniq:
            self.enqueue_log(f"  {c}")
        self.enqueue_log("")

    def enqueue_log(self, text):
        self.log_queue.put(text)

    def process_log_queue(self):
        try:
            while True:
                msg = self.log_queue.get_nowait()
                self.log_text.insert("end", msg + "\n")
                self.log_text.see("end")
        except queue.Empty:
            pass
        self.root.after(100, self.process_log_queue)

    def clear_log(self):
        self.log_text.delete("1.0", "end")
        self.causes_text.delete("1.0", "end")
        self.actions_text.delete("1.0", "end")

    def save_log(self):
        path = filedialog.asksaveasfilename(
            title="診断ログ保存",
            defaultextension=".txt",
            filetypes=[("Text", "*.txt"), ("All files", "*.*")]
        )
        if not path:
            return
        try:
            content = self.log_text.get("1.0", "end")
            summary1 = self.causes_text.get("1.0", "end")
            summary2 = self.actions_text.get("1.0", "end")
            with open(path, "w", encoding="utf-8") as f:
                f.write("=== 詳細ログ ===\n")
                f.write(content)
                f.write("\n=== 原因候補 ===\n")
                f.write(summary1)
                f.write("\n=== 推奨対処 ===\n")
                f.write(summary2)
            messagebox.showinfo("保存", "ログを保存しました。")
        except Exception as e:
            messagebox.showerror("保存エラー", short_exc(e))

    def validate_python(self):
        python_exe = self.python_var.get().strip().strip('"')
        if not python_exe:
            messagebox.showwarning("入力不足", "対象 Python を指定してください。")
            return None
        if not os.path.exists(python_exe):
            messagebox.showwarning("入力エラー", f"Python が見つかりません:\n{python_exe}")
            return None
        return python_exe

    def run_in_thread(self, func):
        t = threading.Thread(target=func, daemon=True)
        t.start()

    def set_summary(self, causes, actions):
        self.causes_text.delete("1.0", "end")
        self.actions_text.delete("1.0", "end")

        if causes:
            for i, c in enumerate(causes, 1):
                self.causes_text.insert("end", f"{i}. {c}\n")
        else:
            self.causes_text.insert("end", "明確な原因候補は抽出できませんでした。\n")

        if actions:
            for i, a in enumerate(actions, 1):
                self.actions_text.insert("end", f"{i}. {a}\n")
        else:
            self.actions_text.insert("end", "推奨対処はありません。\n")

    def run_basic_diag(self):
        python_exe = self.validate_python()
        if not python_exe:
            return

        def worker():
            self.enqueue_log("=" * 80)
            self.enqueue_log("[基本診断開始]")
            self.enqueue_log(f"対象 Python: {python_exe}")
            self.enqueue_log("")

            self.enqueue_log(f"OS: {platform.platform()}")
            self.enqueue_log(f"管理者権限: {is_admin()}")
            self.enqueue_log("")

            # python info
            res_py = self.engine.python_info(python_exe)
            self.enqueue_log("[python 情報]")
            self.enqueue_log(json.dumps(res_py, ensure_ascii=False, indent=2))
            self.enqueue_log("")

            py_info_json = {}
            if res_py["ok"]:
                try:
                    py_info_json = json.loads(res_py["stdout"])
                    self.enqueue_log("[python 情報(整形)]")
                    self.enqueue_log(json.dumps(py_info_json, ensure_ascii=False, indent=2))
                    self.enqueue_log("")
                except Exception as e:
                    self.enqueue_log(f"JSON解析失敗: {short_exc(e)}")

            # pip version
            res_pv = self.engine.pip_version(python_exe)
            self.enqueue_log("[python -m pip --version]")
            self.enqueue_log(res_pv["stdout"] or res_pv["stderr"])
            self.enqueue_log("")

            # pip debug
            res_pd = self.engine.pip_debug(python_exe)
            self.enqueue_log("[python -m pip debug --verbose]")
            self.enqueue_log(res_pd["stdout"] or res_pd["stderr"])
            self.enqueue_log("")

            # where pip
            pip_path = ""
            if is_windows():
                res_where = run_shell_capture("where pip")
                self.enqueue_log("[where pip]")
                self.enqueue_log(res_where["stdout"] or res_where["stderr"])
                self.enqueue_log("")
                lines = [x.strip() for x in res_where["stdout"].splitlines() if x.strip()]
                if lines:
                    pip_path = lines[0]
            else:
                res_which = run_command(["which", "pip"], timeout=15)
                self.enqueue_log("[which pip]")
                self.enqueue_log(res_which["stdout"] or res_which["stderr"])
                self.enqueue_log("")
                pip_path = res_which["stdout"].strip()

            # filesystem check
            fs_checks = self.engine.filesystem_checks(python_exe)
            self.enqueue_log("[ファイルシステム書込チェック]")
            self.enqueue_log(json.dumps(fs_checks, ensure_ascii=False, indent=2))
            self.enqueue_log("")

            env = self.engine.collect_environment()
            self.enqueue_log("[関連環境変数]")
            self.enqueue_log(json.dumps(env, ensure_ascii=False, indent=2))
            self.enqueue_log("")

            collected = {
                "python_exe": python_exe,
                "py_info_json": py_info_json,
                "pip_version_text": (res_pv["stdout"] + "\n" + res_pv["stderr"]).strip(),
                "pip_path": pip_path,
                "env": env,
                "fs_checks": fs_checks,
                "package_name": "",
                "dry_stdout": "",
                "dry_stderr": "",
                "real_stdout": "",
                "real_stderr": "",
            }
            causes, actions = self.engine.guess_causes(collected)
            self.set_summary(causes, actions)

            self.enqueue_log("[基本診断完了]")
            self.enqueue_log("=" * 80)

        self.run_in_thread(worker)

    def run_package_diag(self):
        python_exe = self.validate_python()
        if not python_exe:
            return

        package_name = self.package_var.get().strip()
        if not package_name:
            messagebox.showwarning("入力不足", "パッケージ名を入力してください。")
            return

        def worker():
            self.enqueue_log("=" * 80)
            self.enqueue_log("[パッケージ診断開始]")
            self.enqueue_log(f"対象 Python: {python_exe}")
            self.enqueue_log(f"対象パッケージ: {package_name}")
            self.enqueue_log("")

            # python info
            res_py = self.engine.python_info(python_exe)
            py_info_json = {}
            if res_py["ok"]:
                try:
                    py_info_json = json.loads(res_py["stdout"])
                except Exception:
                    py_info_json = {}

            env = self.engine.collect_environment()
            fs_checks = self.engine.filesystem_checks(python_exe)

            # show
            res_show = self.engine.pip_show(python_exe, package_name)
            self.enqueue_log("[pip show]")
            self.enqueue_log(res_show["stdout"] or res_show["stderr"])
            self.enqueue_log("")

            # versions
            res_versions = self.engine.pip_index_versions(python_exe, package_name)
            self.enqueue_log("[pip index versions]")
            self.enqueue_log(res_versions["stdout"] or res_versions["stderr"])
            self.enqueue_log("")

            # dry-like
            res_dry = self.engine.pip_install_drylike(
                python_exe,
                package_name,
                user_install=self.user_install_var.get()
            )
            self.enqueue_log("[インストール簡易診断]")
            self.enqueue_log("CMD: " + " ".join(res_dry["cmd"]))
            self.enqueue_log(res_dry["stdout"])
            self.enqueue_log(res_dry["stderr"])
            self.enqueue_log("")

            real_stdout = ""
            real_stderr = ""
            if self.auto_real_install_var.get():
                res_real = self.engine.pip_install_real(
                    python_exe,
                    package_name,
                    user_install=self.user_install_var.get(),
                    no_cache=self.no_cache_var.get()
                )
                real_stdout = res_real["stdout"]
                real_stderr = res_real["stderr"]
                self.enqueue_log("[実インストール試験]")
                self.enqueue_log("CMD: " + " ".join(res_real["cmd"]))
                self.enqueue_log(real_stdout)
                self.enqueue_log(real_stderr)
                self.enqueue_log("")

            pip_path = ""
            if is_windows():
                res_where = run_shell_capture("where pip")
                lines = [x.strip() for x in res_where["stdout"].splitlines() if x.strip()]
                if lines:
                    pip_path = lines[0]
            else:
                res_which = run_command(["which", "pip"], timeout=15)
                pip_path = res_which["stdout"].strip()

            collected = {
                "python_exe": python_exe,
                "py_info_json": py_info_json,
                "pip_version_text": "",
                "pip_path": pip_path,
                "env": env,
                "fs_checks": fs_checks,
                "package_name": package_name,
                "dry_stdout": res_dry["stdout"],
                "dry_stderr": res_dry["stderr"],
                "real_stdout": real_stdout,
                "real_stderr": real_stderr,
            }
            causes, actions = self.engine.guess_causes(collected)
            self.set_summary(causes, actions)

            self.enqueue_log("[パッケージ診断完了]")
            self.enqueue_log("=" * 80)

        self.run_in_thread(worker)

    def run_real_install_test(self):
        python_exe = self.validate_python()
        if not python_exe:
            return
        package_name = self.package_var.get().strip()
        if not package_name:
            messagebox.showwarning("入力不足", "パッケージ名を入力してください。")
            return

        if not messagebox.askyesno(
            "確認",
            f"実際に `{package_name}` をインストールします。\n続行しますか？"
        ):
            return

        def worker():
            self.enqueue_log("=" * 80)
            self.enqueue_log("[実インストール試験開始]")
            self.enqueue_log(f"対象 Python: {python_exe}")
            self.enqueue_log(f"対象パッケージ: {package_name}")
            self.enqueue_log("")

            res_real = self.engine.pip_install_real(
                python_exe,
                package_name,
                user_install=self.user_install_var.get(),
                no_cache=self.no_cache_var.get()
            )
            self.enqueue_log("CMD: " + " ".join(res_real["cmd"]))
            self.enqueue_log(res_real["stdout"])
            self.enqueue_log(res_real["stderr"])
            self.enqueue_log("")

            self.enqueue_log("[実インストール試験完了]")
            self.enqueue_log("=" * 80)

        self.run_in_thread(worker)

    def run_spyder_detect(self):
        def worker():
            self.enqueue_log("=" * 80)
            self.enqueue_log("[Spyder 検出開始]")
            data = self.engine.detect_spyder()
            self.enqueue_log(json.dumps(data, ensure_ascii=False, indent=2))
            self.enqueue_log("")

            # 追加ヒント
            hints = [
                "Spyder 側の Python interpreter が、pip で入れた先と一致しているか確認してください。",
                "Spyder のコンソールで `import sys; print(sys.executable)` を実行し、診断対象の Python と一致するか見てください。",
                "Conda 版 Spyder なら、必要パッケージは Spyder 本体環境ではなく、実作業用環境側に入れるのが基本です。"
            ]
            self.set_summary(
                ["Spyder 自体の場所は見つかっても、実際に使う Python 環境は別である可能性があります。"],
                hints
            )

            self.enqueue_log("[Spyder 検出完了]")
            self.enqueue_log("=" * 80)

        self.run_in_thread(worker)

# -----------------------------
# main
# -----------------------------
def main():
    root = tk.Tk()
    style = ttk.Style()
    try:
        if "vista" in style.theme_names():
            style.theme_use("vista")
        elif "clam" in style.theme_names():
            style.theme_use("clam")
    except Exception:
        pass

    app = App(root)
    root.mainloop()

if __name__ == "__main__":
    main()