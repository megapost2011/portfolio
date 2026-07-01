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


APP_TITLE = "Spyder対応 Ollama Migration Doctor Ultimate SAFE"
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
        run(["taskkill", "/IM", "llama-server.exe", "/F"], timeout=15),
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
    # 新しいWindows版では GUI 実体がなくても `ollama serve` でサーバー起動できる
    cli = detect_cli()
    if cli:
        return start_ollama_serve(cli)
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



# ---------------------------------------------------------------------------
# Ollama 起動エラー診断・修復
# ---------------------------------------------------------------------------

LLAMA_SERVER_ERROR_PATTERNS = [
    "llama-server binary not found",
    "error starting llama-server",
    "Run 'cmake -S llama/server",
]


def unique_list(items):
    out = []
    seen = set()
    for x in items:
        if not x:
            continue
        sx = str(x)
        key = os.path.normcase(os.path.abspath(sx)) if IS_WINDOWS else sx
        if key not in seen:
            seen.add(key)
            out.append(sx)
    return out


def powershell_available():
    return shutil.which("powershell") or shutil.which("pwsh")


def get_path_entries():
    env = get_env_registry()
    vals = []
    for scope in ("process", "user", "system"):
        v = env.get(scope, {}).get("PATH", "")
        if v:
            vals.extend([x.strip() for x in str(v).split(os.pathsep) if x.strip()])
    return unique_list(vals)


def where_all_ollama():
    hits = []
    wh = run(["where.exe", "ollama"], timeout=15)
    for line in (wh.get("stdout") or "").splitlines():
        line = line.strip()
        if line.lower().endswith("ollama.exe"):
            hits.append(line)
    cmd = detect_cli()
    if cmd:
        hits.append(cmd)
    for p in default_paths()["cli_candidates"]:
        if p.exists():
            hits.append(str(p))
    # PATH 内に紛れた古い exe を広く確認
    for entry in get_path_entries():
        try:
            pp = Path(entry) / "ollama.exe"
            if pp.exists():
                hits.append(str(pp))
        except Exception:
            pass
    return unique_list(hits)


def ollama_program_dir_from_cli(cli):
    if cli and Path(cli).exists():
        return str(Path(cli).resolve().parent)
    return str(default_paths()["program_dir"])


def expected_llama_server_paths(cli=""):
    program_dir = Path(ollama_program_dir_from_cli(cli))
    local_programs = Path(get_localappdata()) / "Programs"
    downloads = Path(get_user_profile()) / "Downloads"
    bases = [
        program_dir / "llama-server.exe",
        program_dir / "lib" / "ollama" / "llama-server.exe",
        local_programs / "lib" / "ollama" / "llama-server.exe",
        program_dir / "build" / "lib" / "ollama" / "llama-server.exe",
        program_dir / "dist" / "windows-amd64" / "lib" / "ollama" / "llama-server.exe",
        program_dir / "dist" / "windows_amd64" / "lib" / "ollama" / "llama-server.exe",
        downloads / "build" / "lib" / "ollama" / "llama-server.exe",
        downloads / "dist" / "windows-amd64" / "lib" / "ollama" / "llama-server.exe",
        downloads / "dist" / "windows_amd64" / "lib" / "ollama" / "llama-server.exe",
    ]
    return unique_list([str(x) for x in bases])


def find_llama_server_binaries():
    hits = []
    for cli in where_all_ollama():
        for p in expected_llama_server_paths(cli):
            if Path(p).exists():
                hits.append(p)
    for base in [
        Path(get_localappdata()) / "Programs" / "Ollama",
        Path("C:/Program Files/Ollama"),
        Path(get_user_profile()) / "Downloads",
    ]:
        try:
            if base.exists():
                for p in base.rglob("llama-server.exe"):
                    hits.append(str(p))
        except Exception:
            pass
    return unique_list(hits)


def parse_ollama_version_text(text):
    # 例: "ollama version is 0.30.10\nWarning: client version is 0.30.8"
    server = ""
    client = ""
    m = re.search(r"ollama version is\s+([0-9][^\s]+)", text, re.I)
    if m:
        server = m.group(1).strip()
    m = re.search(r"client version is\s+([0-9][^\s]+)", text, re.I)
    if m:
        client = m.group(1).strip()
    return server, client


def netstat_11434():
    res = run(["netstat", "-ano"], timeout=20)
    lines = []
    for line in (res.get("stdout") or "").splitlines():
        if ":11434" in line:
            lines.append(line.strip())
    return lines


def tasklist_for_pid(pid):
    if not pid:
        return ""
    res = run(["tasklist", "/FI", f"PID eq {pid}"], timeout=15)
    return (res.get("stdout") or "").strip()


def kill_pid(pid):
    if not pid:
        return {"ok": False, "stdout": "", "stderr": "PIDなし", "returncode": -1, "cmd": ""}
    return run(["taskkill", "/F", "/PID", str(pid)], timeout=20)


def pids_using_11434():
    pids = []
    for line in netstat_11434():
        parts = line.split()
        if len(parts) >= 5 and parts[-1].isdigit():
            pids.append(parts[-1])
    return unique_list(pids)


def start_ollama_serve(cli=""):
    cli = cli or detect_cli()
    if not cli or not Path(cli).exists():
        return False, "ollama.exe が見つかりません"
    try:
        subprocess.Popen(
            [cli, "serve"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            creationflags=(subprocess.CREATE_NO_WINDOW if IS_WINDOWS else 0),
        )
        return True, f"ollama serve 起動: {cli}"
    except Exception as e:
        return False, f"{type(e).__name__}: {e}"


def ollama_probe_model(cli, model_name, timeout=90):
    """実際に llama-server を起動する確認。show/list だけでは検出できない欠落を拾う。"""
    if not cli or not Path(cli).exists():
        return {"ok": False, "stdout": "", "stderr": "ollama.exe が見つかりません", "returncode": -1, "cmd": ""}
    model_name = model_name or "gemma3:4b"
    prompt = "ping。日本語で一言だけ返事して。"
    return run([cli, "run", model_name, prompt], timeout=timeout)


def read_ollama_logs(max_chars=50000):
    logs_dir = default_paths()["logs_dir"]
    chunks = []
    for name in ["server.log", "app.log", "upgrade.log"]:
        p = logs_dir / name
        chunks.append(f"\n##### {p}\n{read_tail(p, max_chars=max_chars // 3)}")
    return "\n".join(chunks)


def classify_startup_errors(text):
    low = (text or "").lower()
    issues = []

    def add(code, title, severity, cause, fix):
        if code not in [i["code"] for i in issues]:
            issues.append({
                "code": code,
                "title": title,
                "severity": severity,
                "cause": cause,
                "fix": fix,
            })

    if any(p in low for p in LLAMA_SERVER_ERROR_PATTERNS):
        add(
            "MISSING_LLAMA_SERVER",
            "llama-server.exe 欠落 / Ollama本体破損",
            "致命的",
            "ollama.exe はあるが、モデル実行用の llama-server.exe が同梱場所にありません。ZIP/ビルド途中版、破損インストール、古いexe混入で起きます。",
            "Ollamaを停止し、公式インストーラーで本体を再インストールしてください。cmake は不要です。",
        )
    if "warning: client version" in low or "client version is" in low:
        add(
            "CLIENT_SERVER_VERSION_MISMATCH",
            "client/server バージョン不一致",
            "警告〜中",
            "起動中サーバーと、PowerShell/Spyderから呼ぶ ollama.exe の版が違います。PATHに古いollama.exeが混じっている可能性があります。",
            "全Ollamaプロセス停止後に再インストールし、PATHの先頭を AppData\\Local\\Programs\\Ollama に揃えてください。",
        )
    if "could not connect to a running ollama instance" in low or "connection refused" in low or "対象のコンピューターによって拒否" in low:
        add(
            "SERVER_NOT_RUNNING",
            "Ollamaサーバー未起動",
            "中",
            "CLIは存在しますが、localhost:11434 のOllamaサーバーへ接続できません。",
            "Ollamaアプリまたは ollama serve を起動してください。",
        )
    if "only one usage of each socket address" in low or "bind" in low and "11434" in low or "address already in use" in low:
        add(
            "PORT_11434_IN_USE",
            "11434ポート占有",
            "中",
            "別プロセスが localhost:11434 を掴んでいます。古いollama.exeが残っている場合があります。",
            "netstatでPIDを確認し、不要なOllamaプロセスを停止してください。",
        )
    if "model" in low and "not found" in low or "pull model" in low:
        add(
            "MODEL_NOT_FOUND",
            "モデル未取得 / 保存先違い",
            "中",
            "OLLAMA_MODELS が変わった、または指定モデルが未pullです。",
            "ollama list でモデルを確認し、必要なら ollama pull モデル名 を実行してください。",
        )
    if "no space left" in low or "not enough space" in low or "空き容量" in low:
        add(
            "DISK_SPACE",
            "ディスク空き容量不足",
            "致命的",
            "本体またはモデル保存先の空き容量が不足しています。Windows版Ollama本体にも数GB、モデルには数十GB以上必要です。",
            "Eドライブなど十分な空き容量のある場所へ OLLAMA_MODELS を移してください。",
        )
    if "access is denied" in low or "permission denied" in low or "アクセスが拒否" in low:
        add(
            "PERMISSION_DENIED",
            "アクセス権限エラー",
            "中",
            "モデルフォルダまたはOllama本体フォルダへの書き込み/実行権限が不足しています。",
            "ユーザー配下のフォルダを使うか、権限を修復してください。",
        )
    if "digest" in low and "not found" in low or "blob" in low and "not found" in low:
        add(
            "BROKEN_MODEL_BLOB",
            "モデルblob/manifest破損",
            "中",
            "モデルのダウンロード断片、manifestとblobの不整合、移行コピー失敗が疑われます。",
            "該当モデルを再pull、またはmodelsフォルダをrobocopyで再コピーしてください。",
        )
    if "failed to load" in low or "unable to load" in low or "cuda" in low or "vulkan" in low:
        add(
            "RUNTIME_BACKEND",
            "GPU/CPUバックエンド初期化エラー",
            "中",
            "GPUドライバ、VRAM不足、Vulkan/CUDA初期化失敗の可能性があります。",
            "小さいモデルで確認し、GPUドライバ更新、またはCPU/軽量モデルで再確認してください。",
        )
    return issues


def build_startup_diagnostic(cli="", model_name=""):
    cli = cli or detect_cli()
    model_name = model_name or "gemma3:4b"
    report = []
    all_text = []

    report.append(f"診断日時: {now()}")
    report.append(f"CLI: {cli or '(未検出)'}")
    report.append(f"App: {detect_app() or '(未検出)'}")
    report.append(f"確認モデル: {model_name}")
    report.append("")

    report.append("--- ollama.exe 候補 ---")
    for p in where_all_ollama():
        report.append(p)
    report.append("")

    report.append("--- llama-server.exe 確認 ---")
    expected = expected_llama_server_paths(cli)
    found = find_llama_server_binaries()
    for p in expected:
        report.append(f"{'OK ' if Path(p).exists() else 'NG '} {p}")
    if found:
        report.append("検出済み llama-server.exe:")
        report.extend([f"  {x}" for x in found])
    else:
        report.append("検出済み llama-server.exe: なし")
    report.append("")

    report.append("--- 11434 ポート ---")
    ns = netstat_11434()
    if ns:
        report.extend(ns)
        for pid in pids_using_11434():
            report.append(tasklist_for_pid(pid))
    else:
        report.append("11434使用プロセスなし")
    report.append("")

    ver = ollama_version(cli)
    txt = (ver.get("stdout", "") or "") + "\n" + (ver.get("stderr", "") or "")
    all_text.append(txt)
    server_v, client_v = parse_ollama_version_text(txt)
    report.append("--- version ---")
    report.append(txt.strip() or "(出力なし)")
    if server_v or client_v:
        report.append(f"解析: server={server_v or '?'} client={client_v or '?'}")
    report.append("")

    tags = api_tags()
    txt = (tags.get("stdout", "") or "") + "\n" + (tags.get("stderr", "") or "")
    all_text.append(txt)
    report.append("--- /api/tags ---")
    report.append(txt.strip() or "(出力なし)")
    report.append("")

    probe = ollama_probe_model(cli, model_name, timeout=90)
    txt = (probe.get("stdout", "") or "") + "\n" + (probe.get("stderr", "") or "")
    all_text.append(txt)
    report.append(f"--- 実モデル起動テスト: ollama run {model_name} ---")
    report.append(f"rc={probe.get('returncode')}")
    report.append(txt.strip() or "(出力なし)")
    report.append("")

    logs = read_ollama_logs(max_chars=60000)
    all_text.append(logs)
    report.append("--- Ollamaログ抜粋 ---")
    report.append(logs)
    report.append("")

    combined = "\n".join(all_text)
    issues = classify_startup_errors(combined)
    # 実ファイル検査からも補正
    if not find_llama_server_binaries() and cli and Path(cli).exists():
        issues = classify_startup_errors(combined + "\nllama-server binary not found")

    report.append("--- 判定 ---")
    if not issues:
        report.append("致命的な起動エラーは検出されませんでした。")
    else:
        for i, issue in enumerate(issues, 1):
            report.append(f"[{i}] {issue['severity']} / {issue['code']} / {issue['title']}")
            report.append(f"    原因: {issue['cause']}")
            report.append(f"    対処: {issue['fix']}")

    return {"text": "\n".join(report), "issues": issues}


def fix_user_path_prepend_ollama(program_dir=""):
    if winreg is None:
        return False, "winreg unavailable"
    program_dir = program_dir or str(default_paths()["program_dir"])
    old = reg_read(winreg.HKEY_CURRENT_USER, r"Environment", "PATH")
    if not old:
        old = os.environ.get("PATH", "")
    parts = [x for x in str(old).split(os.pathsep) if x.strip()]
    parts = [x for x in parts if os.path.normcase(os.path.abspath(x)) != os.path.normcase(os.path.abspath(program_dir))]
    new = os.pathsep.join([program_dir] + parts)
    return set_user_env("PATH", new)


def run_official_ollama_installer():
    ps = powershell_available()
    if not ps:
        return {"ok": False, "stdout": "", "stderr": "PowerShellが見つかりません", "returncode": -1, "cmd": ""}
    cmd = [ps, "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", "irm https://ollama.com/install.ps1 | iex"]
    return run(cmd, timeout=900)


def backup_program_dir(program_dir=""):
    program_dir = Path(program_dir or default_paths()["program_dir"])
    if not program_dir.exists():
        return False, f"なし: {program_dir}"
    backup = program_dir.parent / f"Ollama_broken_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    try:
        shutil.move(str(program_dir), str(backup))
        return True, str(backup)
    except Exception as e:
        return False, f"{type(e).__name__}: {e}"

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
        # 安全設定: PowerShell/公式インストーラーの自動実行は既定で禁止
        self.allow_powershell_installer_var = tk.BooleanVar(value=False)
        # 安全設定: Ollama本体フォルダ退避も既定で禁止
        self.allow_program_backup_var = tk.BooleanVar(value=False)
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
        ttk.Checkbutton(
            row, text="PowerShell公式インストーラー自動実行を許可",
            variable=self.allow_powershell_installer_var
        ).pack(side="left", padx=8)
        ttk.Checkbutton(
            row, text="Ollama本体フォルダ退避を許可",
            variable=self.allow_program_backup_var
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
            ("8. 起動エラー診断", self.startup_error_diagnosis),
            ("9. 起動エラー自動修復", self.startup_error_auto_repair),
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
    # 8. 起動エラー診断
    # -----------------------------------------------------------------------

    def startup_error_diagnosis(self):
        self.start_worker(self._startup_error_diagnosis)

    def _startup_error_diagnosis(self):
        self.qput("status", "起動エラー診断中...")
        self.qput("log", f"=== {now()} 起動エラー診断 ===")
        cli = self.cli_var.get().strip() or detect_cli()
        self.cli_var.set(cli)
        model_name = self.model_var.get().strip() or "gemma3:4b"
        diag = build_startup_diagnostic(cli, model_name)
        self.qput("detail", diag["text"])
        self.qput("log", "--- 起動エラー判定 ---")
        if not diag["issues"]:
            self.qput("log", "致命的な起動エラーは検出されませんでした。")
        else:
            for issue in diag["issues"]:
                self.qput("log", f"[{issue['severity']}] {issue['code']} - {issue['title']}")
                self.qput("log", f"  原因: {issue['cause']}")
                self.qput("log", f"  対処: {issue['fix']}")
        self.qput("status", "起動エラー診断完了")

    # -----------------------------------------------------------------------
    # 9. 起動エラー自動修復
    # -----------------------------------------------------------------------

    def startup_error_auto_repair(self):
        self.start_worker(self._startup_error_auto_repair)

    def _startup_error_auto_repair(self):
        self.qput("status", "起動エラー自動修復中...")
        self.qput("log", f"=== {now()} 起動エラー自動修復開始 ===")
        cli = self.cli_var.get().strip() or detect_cli()
        model_name = self.model_var.get().strip() or "gemma3:4b"

        self.qput("log", "--- 事前診断 ---")
        diag = build_startup_diagnostic(cli, model_name)
        self.qput("detail", diag["text"])
        codes = {x["code"] for x in diag["issues"]}
        if not codes:
            self.qput("log", "起動エラーは検出されませんでした。修復処理は不要です。")
            self.qput("status", "修復不要")
            return

        self.qput("log", "検出コード: " + ", ".join(sorted(codes)))

        self.qput("log", "--- Ollama/llama-server 停止 ---")
        for res in stop_ollama():
            self.qput("log", f"cmd={res['cmd']} rc={res['returncode']}")
            if res["stdout"].strip():
                self.qput("log", res["stdout"].strip())
            if res["stderr"].strip():
                self.qput("log", res["stderr"].strip())

        if "PORT_11434_IN_USE" in codes:
            self.qput("log", "--- 11434占有PIDの停止 ---")
            for pid in pids_using_11434():
                self.qput("log", tasklist_for_pid(pid))
                res = kill_pid(pid)
                self.qput("log", f"PID {pid} kill rc={res['returncode']}")
                if res["stdout"].strip():
                    self.qput("log", res["stdout"].strip())
                if res["stderr"].strip():
                    self.qput("log", res["stderr"].strip())

        need_reinstall = bool({"MISSING_LLAMA_SERVER", "CLIENT_SERVER_VERSION_MISMATCH"} & codes)
        if need_reinstall:
            program_dir = ollama_program_dir_from_cli(cli)
            self.qput("log", "--- Ollama本体破損/版ズレを検出 ---")
            self.qput("log", f"現在のOllama本体候補: {program_dir}")
            self.qput("log", "安全モードのため、既定ではCドライブ側のOllama本体退避やPowerShell公式インストーラー実行は行いません。")
            self.qput("log", "モデル保存先をEドライブにしていても、Windows版Ollama本体は通常 AppData\Local\Programs\Ollama 配下に置かれます。")
            self.qput("log", "手動修復コマンド案:")
            self.qput("log", "  taskkill /F /IM ollama.exe")
            self.qput("log", "  taskkill /F /IM llama-server.exe")
            self.qput("log", "  irm https://ollama.com/install.ps1 | iex")
            if not self.allow_program_backup_var.get() and not self.allow_powershell_installer_var.get():
                self.qput("log", "自動修復をここで停止しました。実行したい場合は、チェックボックス『PowerShell公式インストーラー自動実行を許可』をONにしてください。")
                self.qput("status", "安全停止: PowerShell自動実行なし")
                return

            if self.allow_program_backup_var.get():
                self.qput("log", "--- 破損したOllama本体フォルダの退避 ---")
                ok_bak, msg_bak = backup_program_dir(program_dir)
                self.qput("log", f"backup: {'OK' if ok_bak else 'SKIP/NG'} {msg_bak}")
            else:
                self.qput("log", "Ollama本体フォルダ退避は許可されていないためスキップしました。")

            if self.allow_powershell_installer_var.get():
                self.qput("log", "--- 公式インストーラー実行 ---")
                inst = run_official_ollama_installer()
                self.qput("log", f"installer rc={inst['returncode']} ok={inst['ok']}")
                if inst["stdout"].strip():
                    self.qput("log", inst["stdout"][-12000:])
                if inst["stderr"].strip():
                    self.qput("log", inst["stderr"][-12000:])
                if not inst["ok"]:
                    self.qput("error", "公式インストーラーの実行に失敗しました。ログを確認してください。")
                    return
                time.sleep(2)
                cli = detect_cli()
                self.cli_var.set(cli)
            else:
                self.qput("log", "PowerShell公式インストーラー実行は許可されていないためスキップしました。")
                self.qput("status", "安全停止: PowerShell自動実行なし")
                return

        self.qput("log", "--- PATH修正 ---")
        program_dir = ollama_program_dir_from_cli(cli)
        ok_path, msg_path = fix_user_path_prepend_ollama(program_dir)
        self.qput("log", f"User PATH: {'OK' if ok_path else 'NG'} {msg_path}")

        dst = self.dst_var.get().strip()
        if dst:
            self.qput("log", "--- OLLAMA_MODELS確認/修正 ---")
            ok_env, err_env = set_user_env("OLLAMA_MODELS", dst)
            self.qput("log", f"OLLAMA_MODELS={dst}: {'OK' if ok_env else 'NG'} {err_env}")

        if "MODEL_NOT_FOUND" in codes:
            self.qput("log", f"--- モデル再取得候補: {model_name} ---")
            self.qput("log", "自動pullは大容量通信になるため実行しません。必要なら PowerShell で以下を実行してください。")
            self.qput("log", f"ollama pull {model_name}")

        self.qput("log", "--- Ollama再起動 ---")
        ok_start, msg_start = start_ollama(self.app_var.get().strip())
        self.qput("log", f"start: {'OK' if ok_start else 'NG'} {msg_start}")
        time.sleep(6)

        self.qput("log", "--- 修復後診断 ---")
        diag2 = build_startup_diagnostic(cli, model_name)
        self.qput("detail", diag2["text"])
        if not diag2["issues"]:
            self.qput("log", "修復後: 致命的な起動エラーは検出されませんでした。")
        else:
            self.qput("log", "修復後も残っている問題:")
            for issue in diag2["issues"]:
                self.qput("log", f"[{issue['severity']}] {issue['code']} - {issue['title']}")
                self.qput("log", f"  対処: {issue['fix']}")
        self.qput("status", "起動エラー自動修復完了")

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
