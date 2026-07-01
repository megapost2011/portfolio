# -*- coding: utf-8 -*-
"""
Created on Sun Jun 28 16:09:34 2026

@author: megap
"""

# -*- coding: utf-8 -*-
"""
OpenOffice Cドライブ -> Eドライブ マイグレーションツール 完全修正版

推奨方式:
  C:\Program Files\OpenOffice 4 の実体を E:\OpenOffice 4 にコピー
  その後、C側にジャンクションを作成する

これにより、レジストリや既存ショートカットが C:\Program Files\OpenOffice 4 を指していても、
実体は E:\OpenOffice 4 に置かれるため、失敗しにくい。

実行条件:
  - Windows
  - 管理者権限
  - OpenOfficeを終了しておく
  - pywin32 は任意。なくても動くが、ショートカット更新はスキップされる。
"""

import ctypes
import datetime
import os
import queue
import shutil
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

try:
    import winreg
except ImportError:
    winreg = None

try:
    import win32com.client
except ImportError:
    win32com = None


APP_TITLE = "OpenOffice マイグレーションツール 完全修正版"


def is_windows():
    return os.name == "nt"


def is_admin():
    if not is_windows():
        return False
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False


def norm_path(path):
    return os.path.normcase(os.path.abspath(os.path.expandvars(os.path.expanduser(path.strip()))))


def is_junction_or_symlink(path):
    return os.path.exists(path) and os.path.islink(path)


def run_cmd(cmd):
    completed = subprocess.run(
        cmd,
        shell=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        encoding="cp932",
        errors="replace",
    )
    return completed.returncode, completed.stdout, completed.stderr


def process_exists(names):
    """
    names例: ["soffice.exe", "soffice.bin"]
    """
    try:
        code, out, err = run_cmd("tasklist")
        if code != 0:
            return False
        lower = out.lower()
        return any(name.lower() in lower for name in names)
    except Exception:
        return False


def get_tree_stats(path):
    """
    フォルダ内のファイル数・合計サイズを返す
    """
    file_count = 0
    total_size = 0

    for root, dirs, files in os.walk(path):
        for name in files:
            full = os.path.join(root, name)
            try:
                file_count += 1
                total_size += os.path.getsize(full)
            except OSError:
                pass

    return file_count, total_size


def safe_makedirs(path):
    os.makedirs(path, exist_ok=True)


def copy_tree_with_progress(src, dst, log_func, progress_func):
    """
    shutil.copytreeだけだと進捗が出ず、失敗箇所も分かりにくいので手動コピーする。
    """
    src = norm_path(src)
    dst = norm_path(dst)

    total_files, total_size = get_tree_stats(src)
    if total_files == 0:
        raise RuntimeError("移動元フォルダにコピー対象ファイルがありません。OpenOfficeのパスが正しいか確認してください。")

    copied_files = 0
    copied_size = 0

    log_func(f"コピー対象: {total_files} ファイル / {total_size:,} bytes")

    for root, dirs, files in os.walk(src):
        rel = os.path.relpath(root, src)
        target_root = dst if rel == "." else os.path.join(dst, rel)

        safe_makedirs(target_root)

        for d in dirs:
            safe_makedirs(os.path.join(target_root, d))

        for name in files:
            src_file = os.path.join(root, name)
            dst_file = os.path.join(target_root, name)

            try:
                safe_makedirs(os.path.dirname(dst_file))
                shutil.copy2(src_file, dst_file)

                copied_files += 1
                try:
                    copied_size += os.path.getsize(src_file)
                except OSError:
                    pass

                if copied_files % 20 == 0 or copied_files == total_files:
                    percent = 10 + int((copied_files / total_files) * 50)
                    progress_func(percent)
                    log_func(f"コピー中: {copied_files}/{total_files} files")

            except Exception as e:
                raise RuntimeError(
                    f"コピーに失敗しました。\n"
                    f"元: {src_file}\n"
                    f"先: {dst_file}\n"
                    f"理由: {e}"
                )

    return total_files, total_size, copied_files, copied_size


def create_junction(link_path, target_path, log_func):
    """
    mklink /J "C:\\Program Files\\OpenOffice 4" "E:\\OpenOffice 4"
    """
    link_path = norm_path(link_path)
    target_path = norm_path(target_path)

    if os.path.exists(link_path):
        raise RuntimeError(f"ジャンクション作成先が既に存在します: {link_path}")

    cmd = f'mklink /J "{link_path}" "{target_path}"'
    log_func(f"ジャンクション作成: {cmd}")

    code, out, err = run_cmd(f'cmd /c {cmd}')

    if code != 0:
        raise RuntimeError(
            "ジャンクション作成に失敗しました。\n"
            f"command: {cmd}\n"
            f"stdout:\n{out}\n"
            f"stderr:\n{err}"
        )

    log_func(out.strip())

    if not os.path.exists(link_path):
        raise RuntimeError("mklinkは成功したように見えますが、ジャンクションが確認できません。")

    return True


def backup_registry(log_func):
    """
    念のためOpenOffice関連レジストリを .reg にバックアップする。
    失敗しても致命傷にはしない。
    """
    backup_dir = os.path.join(os.path.expanduser("~"), "Desktop", "OpenOffice_migration_registry_backup")
    safe_makedirs(backup_dir)

    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")

    targets = [
        (r"HKLM\SOFTWARE\OpenOffice.org", f"HKLM_OpenOffice_{ts}.reg"),
        (r"HKLM\SOFTWARE\WOW6432Node\OpenOffice.org", f"HKLM_WOW6432Node_OpenOffice_{ts}.reg"),
        (r"HKCU\Software\OpenOffice.org", f"HKCU_OpenOffice_{ts}.reg"),
    ]

    for key, filename in targets:
        out_file = os.path.join(backup_dir, filename)
        cmd = f'reg export "{key}" "{out_file}" /y'
        code, out, err = run_cmd(cmd)
        if code == 0:
            log_func(f"レジストリバックアップ作成: {out_file}")
        else:
            log_func(f"レジストリバックアップ対象なし又は失敗: {key}")


def replace_registry_strings(root, subkey, old, new, log_func):
    """
    指定レジストリ配下を再帰的に見て、REG_SZ / REG_EXPAND_SZ 内の old を new に置換する。
    ジャンクション方式では必須ではないため、オプション扱い。
    """
    if winreg is None:
        log_func("winregが使えないため、レジストリ更新をスキップしました。")
        return 0

    changed = 0

    try:
        key = winreg.OpenKey(root, subkey, 0, winreg.KEY_READ | winreg.KEY_WRITE)
    except FileNotFoundError:
        return 0
    except PermissionError:
        log_func(f"レジストリ権限なし: {subkey}")
        return 0
    except OSError as e:
        log_func(f"レジストリを開けません: {subkey} / {e}")
        return 0

    try:
        value_count = winreg.QueryInfoKey(key)[1]

        for i in range(value_count):
            try:
                name, data, typ = winreg.EnumValue(key, i)

                if typ in (winreg.REG_SZ, winreg.REG_EXPAND_SZ) and isinstance(data, str):
                    if old.lower() in data.lower():
                        replaced = data.replace(old, new)

                        # 大文字小文字違い対策
                        if replaced == data:
                            replaced = data.replace(old.replace("\\", "\\\\"), new.replace("\\", "\\\\"))

                        if replaced != data:
                            winreg.SetValueEx(key, name, 0, typ, replaced)
                            changed += 1
                            log_func(f"Registry更新: {subkey}\\{name}")

            except OSError:
                continue

        subkey_count = winreg.QueryInfoKey(key)[0]
        subkeys = []

        for i in range(subkey_count):
            try:
                child = winreg.EnumKey(key, i)
                subkeys.append(child)
            except OSError:
                pass

    finally:
        winreg.CloseKey(key)

    for child in subkeys:
        changed += replace_registry_strings(root, subkey + "\\" + child, old, new, log_func)

    return changed


def update_openoffice_registry_paths(old, new, log_func):
    """
    OpenOffice関連だけ置換。
    """
    if winreg is None:
        return

    roots = [
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\OpenOffice.org"),
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\OpenOffice.org"),
        (winreg.HKEY_CURRENT_USER, r"Software\OpenOffice.org"),
    ]

    total = 0
    for root, subkey in roots:
        total += replace_registry_strings(root, subkey, old, new, log_func)

    log_func(f"レジストリ置換完了: {total} 箇所")


def update_shortcuts(old, new, log_func):
    """
    .lnkショートカットのリンク先と作業フォルダを置換する。
    pywin32 がない場合はスキップ。
    """
    if win32com is None:
        log_func("pywin32が無いため、ショートカット更新はスキップします。")
        log_func("必要なら後で: pip install pywin32")
        return 0

    shell = win32com.client.Dispatch("WScript.Shell")

    candidates = [
        os.path.join(os.path.expanduser("~"), "Desktop"),
        os.path.join(os.path.expanduser("~"), "AppData", "Roaming", "Microsoft", "Windows", "Start Menu", "Programs"),
        os.path.join(os.environ.get("PUBLIC", r"C:\Users\Public"), "Desktop"),
        r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs",
    ]

    changed = 0
    old_norm = old.lower().replace("/", "\\")
    new_norm = new.replace("/", "\\")

    for base in candidates:
        if not os.path.isdir(base):
            continue

        for root, dirs, files in os.walk(base):
            for name in files:
                if not name.lower().endswith(".lnk"):
                    continue

                lnk = os.path.join(root, name)

                try:
                    sc = shell.CreateShortcut(lnk)

                    target = sc.Targetpath or ""
                    workdir = sc.WorkingDirectory or ""

                    updated = False

                    if target.lower().replace("/", "\\").startswith(old_norm):
                        sc.Targetpath = target.replace(old, new_norm)
                        updated = True

                    if workdir.lower().replace("/", "\\").startswith(old_norm):
                        sc.WorkingDirectory = workdir.replace(old, new_norm)
                        updated = True

                    if updated:
                        sc.Save()
                        changed += 1
                        log_func(f"ショートカット更新: {lnk}")

                except Exception as e:
                    log_func(f"ショートカット更新失敗: {lnk} / {e}")

    log_func(f"ショートカット更新完了: {changed} 件")
    return changed


class OpenOfficeMigratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("760x620")

        self.queue = queue.Queue()
        self.worker = None

        self.source_var = tk.StringVar(value=r"C:\Program Files\OpenOffice 4")
        self.dest_var = tk.StringVar(value=r"E:\OpenOffice 4")

        self.keep_backup_var = tk.BooleanVar(value=True)
        self.update_shortcuts_var = tk.BooleanVar(value=True)
        self.update_registry_var = tk.BooleanVar(value=False)
        self.allow_existing_dest_var = tk.BooleanVar(value=True)

        self.status_var = tk.StringVar(value="準備完了")

        self.build_ui()
        self.root.after(100, self.poll_queue)

    def build_ui(self):
        main = ttk.Frame(self.root)
        main.pack(fill="both", expand=True, padx=12, pady=12)

        frame_config = ttk.LabelFrame(main, text="設定")
        frame_config.pack(fill="x", pady=6)

        frame_config.columnconfigure(1, weight=1)

        ttk.Label(frame_config, text="移動元 OpenOffice フォルダ").grid(row=0, column=0, sticky="w", padx=8, pady=6)
        ttk.Entry(frame_config, textvariable=self.source_var).grid(row=0, column=1, sticky="ew", padx=8, pady=6)
        ttk.Button(frame_config, text="参照", command=self.browse_source).grid(row=0, column=2, padx=8, pady=6)

        ttk.Label(frame_config, text="移動先 Eドライブ フォルダ").grid(row=1, column=0, sticky="w", padx=8, pady=6)
        ttk.Entry(frame_config, textvariable=self.dest_var).grid(row=1, column=1, sticky="ew", padx=8, pady=6)
        ttk.Button(frame_config, text="参照", command=self.browse_dest).grid(row=1, column=2, padx=8, pady=6)

        ttk.Checkbutton(
            frame_config,
            text="移動先フォルダが既に存在しても利用する",
            variable=self.allow_existing_dest_var
        ).grid(row=2, column=0, columnspan=3, sticky="w", padx=8, pady=3)

        ttk.Checkbutton(
            frame_config,
            text="Cドライブ側の元フォルダをバックアップとして残す",
            variable=self.keep_backup_var
        ).grid(row=3, column=0, columnspan=3, sticky="w", padx=8, pady=3)

        ttk.Checkbutton(
            frame_config,
            text="ショートカットもEドライブ向けに更新する",
            variable=self.update_shortcuts_var
        ).grid(row=4, column=0, columnspan=3, sticky="w", padx=8, pady=3)

        ttk.Checkbutton(
            frame_config,
            text="OpenOffice関連レジストリ内のパスもEドライブ向けに置換する（通常は不要）",
            variable=self.update_registry_var
        ).grid(row=5, column=0, columnspan=3, sticky="w", padx=8, pady=3)

        frame_action = ttk.LabelFrame(main, text="実行")
        frame_action.pack(fill="x", pady=6)

        ttk.Label(frame_action, textvariable=self.status_var).pack(anchor="w", padx=8, pady=6)

        self.progress = ttk.Progressbar(frame_action, orient="horizontal", mode="determinate", maximum=100)
        self.progress.pack(fill="x", padx=8, pady=6)

        btns = ttk.Frame(frame_action)
        btns.pack(fill="x", padx=8, pady=8)

        self.btn_start = ttk.Button(btns, text="マイグレーション開始", command=self.start)
        self.btn_start.pack(side="left", padx=4)

        ttk.Button(btns, text="終了", command=self.root.destroy).pack(side="left", padx=4)

        frame_log = ttk.LabelFrame(main, text="ログ")
        frame_log.pack(fill="both", expand=True, pady=6)

        self.text_log = tk.Text(frame_log, height=18, wrap="word")
        self.text_log.pack(side="left", fill="both", expand=True)

        scroll = ttk.Scrollbar(frame_log, command=self.text_log.yview)
        scroll.pack(side="right", fill="y")
        self.text_log.configure(yscrollcommand=scroll.set)

        self.log("このツールは OpenOffice を Eドライブへコピーし、Cドライブ側にジャンクションを作成します。")
        self.log("必ず OpenOffice を終了し、管理者権限で実行してください。")

    def browse_source(self):
        path = filedialog.askdirectory(title="移動元 OpenOffice フォルダを選択")
        if path:
            self.source_var.set(path)

    def browse_dest(self):
        path = filedialog.askdirectory(title="移動先フォルダを選択")
        if path:
            self.dest_var.set(path)

    def log(self, msg):
        ts = datetime.datetime.now().strftime("%H:%M:%S")
        self.text_log.insert("end", f"[{ts}] {msg}\n")
        self.text_log.see("end")
        self.root.update_idletasks()

    def qlog(self, msg):
        self.queue.put(("log", msg))

    def qstatus(self, msg):
        self.queue.put(("status", msg))

    def qprogress(self, value):
        self.queue.put(("progress", value))

    def poll_queue(self):
        try:
            while True:
                kind, value = self.queue.get_nowait()

                if kind == "log":
                    self.log(value)
                elif kind == "status":
                    self.status_var.set(value)
                elif kind == "progress":
                    self.progress["value"] = value
                elif kind == "done":
                    self.btn_start.config(state="normal")
                    messagebox.showinfo("完了", value)
                elif kind == "error":
                    self.btn_start.config(state="normal")
                    messagebox.showerror("エラー", value)

        except queue.Empty:
            pass

        self.root.after(100, self.poll_queue)

    def start(self):
        if self.worker and self.worker.is_alive():
            messagebox.showwarning("実行中", "現在処理中です。")
            return

        source = norm_path(self.source_var.get())
        dest = norm_path(self.dest_var.get())

        if not is_windows():
            messagebox.showerror("エラー", "このツールはWindows専用です。")
            return

        if not is_admin():
            messagebox.showerror(
                "管理者権限が必要",
                "Program Files配下のリネームとジャンクション作成には管理者権限が必要です。\n\n"
                "PowerShellまたはコマンドプロンプトを「管理者として実行」し、そこから起動してください。"
            )
            return

        if process_exists(["soffice.exe", "soffice.bin", "swriter.exe", "scalc.exe"]):
            messagebox.showerror(
                "OpenOfficeが起動中",
                "OpenOffice関連プロセスが起動中です。\n"
                "Writer、Calc、Quickstarter等をすべて終了してから再実行してください。"
            )
            return

        if not os.path.isdir(source):
            messagebox.showerror("エラー", f"移動元フォルダが存在しません。\n{source}")
            return

        if os.path.abspath(source).lower() == os.path.abspath(dest).lower():
            messagebox.showerror("エラー", "移動元と移動先が同じです。")
            return

        if dest.lower().startswith(source.lower() + os.sep):
            messagebox.showerror("エラー", "移動先を移動元フォルダの中にすることはできません。")
            return

        if os.path.exists(dest) and not self.allow_existing_dest_var.get():
            messagebox.showerror(
                "エラー",
                f"移動先フォルダが既に存在します。\n{dest}\n\n"
                "既存フォルダを利用する場合はチェックを入れてください。"
            )
            return

        msg = (
            "以下の処理を実行します。\n\n"
            f"移動元: {source}\n"
            f"移動先: {dest}\n\n"
            "1. EドライブへOpenOfficeをコピー\n"
            "2. コピー検証\n"
            "3. Cドライブ側の元フォルダをバックアップ名に変更\n"
            "4. Cドライブ側にEドライブへのジャンクションを作成\n\n"
            "続行しますか？"
        )

        if not messagebox.askyesno("確認", msg):
            return

        self.btn_start.config(state="disabled")
        self.progress["value"] = 0
        self.text_log.delete("1.0", "end")

        self.worker = threading.Thread(target=self.do_migration, daemon=True)
        self.worker.start()

    def do_migration(self):
        source = norm_path(self.source_var.get())
        dest = norm_path(self.dest_var.get())

        try:
            self.qstatus("開始")
            self.qprogress(0)
            self.qlog("マイグレーション開始")

            self.qlog(f"移動元: {source}")
            self.qlog(f"移動先: {dest}")

            if is_junction_or_symlink(source):
                raise RuntimeError(
                    "移動元が既にジャンクションまたはシンボリックリンクです。\n"
                    "既に移行済みの可能性があります。"
                )

            self.qstatus("移動先フォルダを準備中")
            self.qprogress(5)
            safe_makedirs(dest)

            self.qstatus("レジストリバックアップ中")
            self.qprogress(8)
            backup_registry(self.qlog)

            self.qstatus("ファイルコピー中")
            copied = copy_tree_with_progress(source, dest, self.qlog, self.qprogress)

            src_files, src_size, copied_files, copied_size = copied

            self.qstatus("コピー検証中")
            self.qprogress(65)

            dst_files, dst_size = get_tree_stats(dest)

            self.qlog(f"移動元: {src_files} files / {src_size:,} bytes")
            self.qlog(f"移動先: {dst_files} files / {dst_size:,} bytes")

            if dst_files < src_files:
                raise RuntimeError(
                    "コピー後のファイル数が不足しています。\n"
                    f"移動元: {src_files}\n"
                    f"移動先: {dst_files}\n"
                    "コピーが不完全なため中止します。"
                )

            if dst_size < int(src_size * 0.98):
                raise RuntimeError(
                    "コピー後の合計サイズが不足しています。\n"
                    f"移動元: {src_size:,} bytes\n"
                    f"移動先: {dst_size:,} bytes\n"
                    "コピーが不完全な可能性が高いため中止します。"
                )

            self.qlog("コピー検証OK")

            ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_path = source.rstrip("\\/") + f"_backup_{ts}"

            self.qstatus("Cドライブ側の元フォルダをバックアップ名へ変更中")
            self.qprogress(72)

            if os.path.exists(backup_path):
                raise RuntimeError(f"バックアップ先が既に存在します: {backup_path}")

            self.qlog(f"元フォルダをリネーム: {source} -> {backup_path}")
            os.rename(source, backup_path)

            try:
                self.qstatus("ジャンクション作成中")
                self.qprogress(78)
                create_junction(source, dest, self.qlog)
            except Exception as e:
                self.qlog("ジャンクション作成に失敗したため、元フォルダを復元します。")
                if not os.path.exists(source):
                    os.rename(backup_path, source)
                raise e

            self.qprogress(85)
            self.qlog("ジャンクション作成OK")
            self.qlog(f"C側の見かけのパス: {source}")
            self.qlog(f"実体の保存先: {dest}")

            if self.update_registry_var.get():
                self.qstatus("レジストリ内パスを更新中")
                self.qprogress(88)
                update_openoffice_registry_paths(source, dest, self.qlog)
            else:
                self.qlog("レジストリ置換はスキップしました。ジャンクション方式なので通常は問題ありません。")

            if self.update_shortcuts_var.get():
                self.qstatus("ショートカット更新中")
                self.qprogress(92)
                update_shortcuts(source, dest, self.qlog)
            else:
                self.qlog("ショートカット更新はスキップしました。")

            if not self.keep_backup_var.get():
                self.qstatus("Cドライブ側バックアップを削除中")
                self.qprogress(96)
                self.qlog(f"バックアップ削除: {backup_path}")
                shutil.rmtree(backup_path, ignore_errors=False)
                self.qlog("バックアップ削除完了")
            else:
                self.qlog(f"Cドライブ側バックアップを残しました: {backup_path}")
                self.qlog("容量を本当に空けたい場合は、動作確認後にこのバックアップフォルダを手動削除してください。")

            self.qstatus("完了")
            self.qprogress(100)

            done_msg = (
                "OpenOfficeのマイグレーションが完了しました。\n\n"
                f"実体: {dest}\n"
                f"C側ジャンクション: {source}\n\n"
                "OpenOfficeが起動するか確認してください。"
            )
            self.queue.put(("done", done_msg))

        except Exception as e:
            self.qstatus("エラー")
            self.qlog(f"ERROR: {e}")
            self.queue.put(("error", str(e)))


def main():
    root = tk.Tk()
    app = OpenOfficeMigratorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()