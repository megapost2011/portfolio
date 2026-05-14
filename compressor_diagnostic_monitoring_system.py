# -*- coding: utf-8 -*-
"""
Created on Thu May  7 20:08:02 2026

@author: User
"""

#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
メーカー非依存・旧型機対応 コンプレッサー故障診断GUI
1ファイル完結版

対応モード:
1. シミュレーション
2. Modbus TCP
3. 外付けセンサーCSV

CSV形式例:
timestamp,current_a,pressure_bar,discharge_temp_c,ambient_temp_c,vibration_g,sound_db,run_hours,start_count,is_running
2026-05-07 10:00:00,52.1,7.1,82.0,25.0,0.8,62,1234,850,1
"""

import csv
import os
import random
import time
from dataclasses import dataclass
from typing import List, Optional

import tkinter as tk
from tkinter import ttk, filedialog, messagebox


# ==========================================================
# データモデル
# ==========================================================

@dataclass
class CompressorState:
    current_a: float = 0.0
    pressure_bar: float = 0.0
    discharge_temp_c: float = 0.0
    ambient_temp_c: float = 0.0
    vibration_g: float = 0.0
    sound_db: float = 0.0
    run_hours: float = 0.0
    start_count: int = 0
    is_running: bool = False
    source: str = "UNKNOWN"


@dataclass
class FaultCode:
    code: str
    severity: str
    message: str
    causes: List[str]
    actions: List[str]


# ==========================================================
# デバイス共通インターフェース
# ==========================================================

class BaseDevice:
    def connect(self) -> bool:
        return True

    def disconnect(self):
        pass

    def read_state(self) -> CompressorState:
        raise NotImplementedError


# ==========================================================
# 1. シミュレーションデバイス
# ==========================================================

class SimulationDevice(BaseDevice):
    def __init__(self):
        self.run_hours = 1234.0
        self.start_count = 850

    def read_state(self) -> CompressorState:
        is_running = random.random() > 0.15

        current = random.uniform(38, 58) if is_running else 0.0
        pressure = random.uniform(6.5, 7.5) if is_running else 0.0
        temp = random.uniform(75, 88) if is_running else random.uniform(25, 45)
        ambient = random.uniform(15, 35)
        vibration = random.uniform(0.3, 1.2) if is_running else random.uniform(0.0, 0.2)
        sound = random.uniform(58, 72) if is_running else random.uniform(35, 45)

        # 異常イベントを時々発生
        r = random.random()
        if r < 0.08 and is_running:
            temp += random.uniform(12, 25)          # 過熱
        elif r < 0.16 and is_running:
            current += random.uniform(12, 25)       # 過負荷
        elif r < 0.24 and is_running:
            pressure = random.uniform(3.5, 5.8)     # 圧力不足
        elif r < 0.32 and is_running:
            vibration += random.uniform(2.0, 4.0)   # 振動異常
        elif r < 0.40 and is_running:
            sound += random.uniform(18, 30)         # 異音

        if is_running:
            self.run_hours += 0.02

        if random.random() < 0.02:
            self.start_count += 1

        return CompressorState(
            current_a=current,
            pressure_bar=pressure,
            discharge_temp_c=temp,
            ambient_temp_c=ambient,
            vibration_g=vibration,
            sound_db=sound,
            run_hours=self.run_hours,
            start_count=self.start_count,
            is_running=is_running,
            source="Simulation"
        )


# ==========================================================
# 2. Modbus TCP デバイス
# ==========================================================

class ModbusTcpDevice(BaseDevice):
    """
    Modbus TCP接続用。
    pymodbus が未インストールでもアプリ全体は動くようにしている。

    想定レジスタ:
    address 0: 電流 x10
    address 1: 圧力 x100
    address 2: 吐出温度 x10
    address 3: 周囲温度 x10
    address 4: 振動 x100
    address 5: 音量 dB x10
    address 6: 運転時間 下位
    address 7: 起動回数
    address 8: 運転フラグ 0/1
    """

    def __init__(self, host: str, port: int, unit_id: int):
        self.host = host
        self.port = port
        self.unit_id = unit_id
        self.client = None

    def connect(self) -> bool:
        try:
            from pymodbus.client import ModbusTcpClient
        except Exception:
            raise RuntimeError(
                "pymodbus がインストールされていません。\n"
                "Modbus TCPを使う場合は以下を実行してください:\n\n"
                "pip install pymodbus"
            )

        self.client = ModbusTcpClient(host=self.host, port=self.port, timeout=2)
        return self.client.connect()

    def disconnect(self):
        if self.client:
            try:
                self.client.close()
            except Exception:
                pass

    def read_state(self) -> CompressorState:
        if not self.client:
            raise RuntimeError("Modbusクライアントが未接続です。")

        rr = self.client.read_holding_registers(
            address=0,
            count=9,
            slave=self.unit_id
        )

        if rr.isError():
            raise RuntimeError("Modbusレジスタ読取に失敗しました。")

        regs = rr.registers

        return CompressorState(
            current_a=regs[0] / 10.0,
            pressure_bar=regs[1] / 100.0,
            discharge_temp_c=regs[2] / 10.0,
            ambient_temp_c=regs[3] / 10.0,
            vibration_g=regs[4] / 100.0,
            sound_db=regs[5] / 10.0,
            run_hours=float(regs[6]),
            start_count=int(regs[7]),
            is_running=bool(regs[8]),
            source=f"Modbus TCP {self.host}:{self.port}"
        )


# ==========================================================
# 3. 外付けセンサーCSVデバイス
# ==========================================================

class CsvSensorDevice(BaseDevice):
    def __init__(self, csv_path: str):
        self.csv_path = csv_path
        self.rows = []
        self.index = 0

    def connect(self) -> bool:
        if not os.path.exists(self.csv_path):
            raise FileNotFoundError(self.csv_path)

        with open(self.csv_path, "r", encoding="utf-8-sig", newline="") as f:
            reader = csv.DictReader(f)
            self.rows = list(reader)

        if not self.rows:
            raise RuntimeError("CSVにデータ行がありません。")

        self.index = 0
        return True

    def _get_float(self, row, key, default=0.0):
        try:
            return float(row.get(key, default))
        except Exception:
            return default

    def _get_int(self, row, key, default=0):
        try:
            return int(float(row.get(key, default)))
        except Exception:
            return default

    def read_state(self) -> CompressorState:
        if not self.rows:
            raise RuntimeError("CSVが読み込まれていません。")

        row = self.rows[self.index]
        self.index = (self.index + 1) % len(self.rows)

        return CompressorState(
            current_a=self._get_float(row, "current_a"),
            pressure_bar=self._get_float(row, "pressure_bar"),
            discharge_temp_c=self._get_float(row, "discharge_temp_c"),
            ambient_temp_c=self._get_float(row, "ambient_temp_c"),
            vibration_g=self._get_float(row, "vibration_g"),
            sound_db=self._get_float(row, "sound_db"),
            run_hours=self._get_float(row, "run_hours"),
            start_count=self._get_int(row, "start_count"),
            is_running=bool(self._get_int(row, "is_running")),
            source=f"External Sensor CSV: {os.path.basename(self.csv_path)}"
        )


# ==========================================================
# 診断エンジン
# ==========================================================

class DiagnosticEngine:
    def __init__(self):
        self.max_temp = 95.0
        self.max_current = 65.0
        self.min_pressure = 6.0
        self.max_vibration = 2.5
        self.max_sound = 85.0
        self.oil_change_hours = 4000.0

    def diagnose(self, s: CompressorState) -> List[FaultCode]:
        faults = []

        if s.is_running and s.discharge_temp_c > self.max_temp:
            faults.append(FaultCode(
                "A0101", "HIGH",
                "吐出温度過昇。冷却不良の疑いがあります。",
                ["冷却器フィン目詰まり", "冷却ファン異常", "換気不良", "オイル劣化", "周囲温度過大"],
                ["冷却器清掃", "ファン回転確認", "設置場所の換気確認", "オイル状態確認"]
            ))

        if s.is_running and s.pressure_bar < self.min_pressure:
            faults.append(FaultCode(
                "A0201", "MEDIUM",
                "吐出圧力不足。エア漏れまたは吸込不良の疑いがあります。",
                ["配管エア漏れ", "吸込フィルタ詰まり", "アンロード弁不良", "圧力設定不良"],
                ["配管漏れ点検", "吸込フィルタ清掃", "アンロード弁点検", "設定圧確認"]
            ))

        if s.is_running and s.current_a > self.max_current:
            faults.append(FaultCode(
                "A0301", "HIGH",
                "モータ過負荷傾向。電気系または機械的負荷異常の疑いがあります。",
                ["電源電圧低下", "欠相", "軸受異常", "圧縮機本体の固着", "ベルト張り過大"],
                ["三相電圧測定", "電流バランス確認", "軸受音確認", "絶縁抵抗測定", "ベルト張力確認"]
            ))

        if s.is_running and s.vibration_g > self.max_vibration:
            faults.append(FaultCode(
                "A0401", "HIGH",
                "振動異常。軸受・芯ずれ・アンバランスの疑いがあります。",
                ["ベアリング摩耗", "芯ずれ", "基礎ボルト緩み", "回転体アンバランス"],
                ["振動測定", "固定ボルト増し締め", "軸受点検", "芯出し確認"]
            ))

        if s.is_running and s.sound_db > self.max_sound:
            faults.append(FaultCode(
                "A0501", "MEDIUM",
                "異音・騒音増加。ベルト鳴き、軸受異常、エア漏れの疑いがあります。",
                ["ベルト鳴き", "軸受異音", "配管エア漏れ", "吸込音増大"],
                ["聴診棒等で異音箇所確認", "ベルト点検", "漏れ点検", "軸受温度確認"]
            ))

        if s.run_hours > self.oil_change_hours:
            faults.append(FaultCode(
                "A0601", "LOW",
                "保守時期超過。オイル交換または定期点検時期です。",
                ["オイル交換未実施", "フィルタ交換未実施", "点検周期超過"],
                ["オイル交換", "オイルフィルタ交換", "点検記録確認"]
            ))

        # 複合診断
        if s.is_running and s.current_a > 60 and s.pressure_bar < 6:
            faults.append(FaultCode(
                "A0701", "HIGH",
                "電流上昇＋圧力不足。負荷異常またはアンロード系不良の可能性が高いです。",
                ["アンロード弁不良", "圧縮機内部抵抗増加", "配管漏れ", "制御不良"],
                ["アンロード動作確認", "圧力上昇時間確認", "配管漏れ確認", "制御盤アラーム履歴確認"]
            ))

        if s.is_running and s.discharge_temp_c > 90 and s.vibration_g > 2.0:
            faults.append(FaultCode(
                "A0801", "HIGH",
                "温度上昇＋振動増加。潤滑不良または軸受異常の疑いがあります。",
                ["潤滑不良", "オイル劣化", "軸受摩耗", "芯ずれ"],
                ["オイル量・色・臭い確認", "軸受温度確認", "振動傾向確認", "早期停止点検を検討"]
            ))

        return faults


# ==========================================================
# GUI
# ==========================================================

class CompressorDiagnosticApp:
    UPDATE_INTERVAL_MS = 2000

    def __init__(self, root):
        self.root = root
        self.root.title("メーカー非依存・旧型機対応 コンプレッサー故障診断システム")
        self.root.geometry("1050x720")

        self.device: Optional[BaseDevice] = None
        self.engine = DiagnosticEngine()
        self.current_faults: List[FaultCode] = []
        self.running = False

        self._build_ui()

    def _build_ui(self):
        # 接続設定
        conn = ttk.LabelFrame(self.root, text="接続モード", padding=10)
        conn.pack(fill=tk.X, padx=10, pady=5)

        self.mode_var = tk.StringVar(value="simulation")
        ttk.Radiobutton(conn, text="シミュレーション", variable=self.mode_var, value="simulation").grid(row=0, column=0, padx=5)
        ttk.Radiobutton(conn, text="Modbus TCP", variable=self.mode_var, value="modbus").grid(row=0, column=1, padx=5)
        ttk.Radiobutton(conn, text="外付けセンサーCSV", variable=self.mode_var, value="csv").grid(row=0, column=2, padx=5)

        ttk.Label(conn, text="IP:").grid(row=1, column=0, sticky=tk.E)
        self.ip_var = tk.StringVar(value="192.168.1.100")
        ttk.Entry(conn, textvariable=self.ip_var, width=16).grid(row=1, column=1, sticky=tk.W)

        ttk.Label(conn, text="Port:").grid(row=1, column=2, sticky=tk.E)
        self.port_var = tk.StringVar(value="502")
        ttk.Entry(conn, textvariable=self.port_var, width=8).grid(row=1, column=3, sticky=tk.W)

        ttk.Label(conn, text="Unit ID:").grid(row=1, column=4, sticky=tk.E)
        self.unit_var = tk.StringVar(value="1")
        ttk.Entry(conn, textvariable=self.unit_var, width=6).grid(row=1, column=5, sticky=tk.W)

        ttk.Label(conn, text="CSV:").grid(row=2, column=0, sticky=tk.E)
        self.csv_var = tk.StringVar()
        ttk.Entry(conn, textvariable=self.csv_var, width=60).grid(row=2, column=1, columnspan=4, sticky=tk.W)
        ttk.Button(conn, text="参照", command=self.select_csv).grid(row=2, column=5, padx=5)

        ttk.Button(conn, text="接続開始", command=self.start_connection).grid(row=0, column=6, padx=10)
        ttk.Button(conn, text="停止", command=self.stop_connection).grid(row=0, column=7, padx=5)
        ttk.Button(conn, text="サンプルCSV作成", command=self.create_sample_csv).grid(row=2, column=6, padx=10)

        # 状態表示
        status = ttk.LabelFrame(self.root, text="ライブデータ", padding=10)
        status.pack(fill=tk.X, padx=10, pady=5)

        self.values = {}
        labels = [
            ("source", "データ元"),
            ("is_running", "運転状態"),
            ("current_a", "モータ電流[A]"),
            ("pressure_bar", "吐出圧力[bar]"),
            ("discharge_temp_c", "吐出温度[℃]"),
            ("ambient_temp_c", "周囲温度[℃]"),
            ("vibration_g", "振動[g]"),
            ("sound_db", "騒音[dB]"),
            ("run_hours", "運転時間[h]"),
            ("start_count", "起動回数"),
        ]

        for i, (key, label) in enumerate(labels):
            r = i // 2
            c = (i % 2) * 2
            ttk.Label(status, text=label + ":").grid(row=r, column=c, sticky=tk.E, padx=5, pady=3)
            var = tk.StringVar(value="-")
            self.values[key] = var
            ttk.Label(status, textvariable=var, width=35).grid(row=r, column=c + 1, sticky=tk.W, padx=5)

        # 故障コード一覧
        mid = ttk.LabelFrame(self.root, text="診断結果", padding=10)
        mid.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        cols = ("code", "severity", "message")
        self.tree = ttk.Treeview(mid, columns=cols, show="headings", height=10)
        self.tree.heading("code", text="コード")
        self.tree.heading("severity", text="重要度")
        self.tree.heading("message", text="内容")
        self.tree.column("code", width=90, anchor=tk.CENTER)
        self.tree.column("severity", width=90, anchor=tk.CENTER)
        self.tree.column("message", width=760, anchor=tk.W)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        sb = ttk.Scrollbar(mid, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        sb.pack(side=tk.RIGHT, fill=tk.Y)

        self.tree.bind("<<TreeviewSelect>>", self.show_fault_detail)

        # 詳細欄
        bottom = ttk.LabelFrame(self.root, text="故障詳細・点検手順", padding=10)
        bottom.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        self.detail = tk.Text(bottom, height=9, wrap=tk.WORD)
        self.detail.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        dsb = ttk.Scrollbar(bottom, orient=tk.VERTICAL, command=self.detail.yview)
        self.detail.configure(yscrollcommand=dsb.set)
        dsb.pack(side=tk.RIGHT, fill=tk.Y)

        # ログ欄
        logf = ttk.LabelFrame(self.root, text="通信・動作ログ", padding=10)
        logf.pack(fill=tk.X, padx=10, pady=5)

        self.log_var = tk.StringVar(value="未接続")
        ttk.Label(logf, textvariable=self.log_var).pack(anchor=tk.W)

    def select_csv(self):
        path = filedialog.askopenfilename(
            title="外付けセンサーCSVを選択",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if path:
            self.csv_var.set(path)

    def create_sample_csv(self):
        path = filedialog.asksaveasfilename(
            title="サンプルCSV保存先",
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")]
        )
        if not path:
            return

        headers = [
            "timestamp", "current_a", "pressure_bar", "discharge_temp_c",
            "ambient_temp_c", "vibration_g", "sound_db",
            "run_hours", "start_count", "is_running"
        ]

        with open(path, "w", encoding="utf-8-sig", newline="") as f:
            writer = csv.writer(f)
            writer.writerow(headers)

            run_hours = 1200
            start_count = 800

            for i in range(60):
                is_running = 1
                current = random.uniform(42, 58)
                pressure = random.uniform(6.6, 7.4)
                temp = random.uniform(76, 88)
                ambient = random.uniform(20, 32)
                vib = random.uniform(0.4, 1.2)
                sound = random.uniform(60, 72)

                if 15 <= i <= 20:
                    temp = random.uniform(98, 112)
                if 30 <= i <= 35:
                    pressure = random.uniform(4.5, 5.6)
                    current = random.uniform(62, 75)
                if 45 <= i <= 50:
                    vib = random.uniform(2.8, 4.5)
                    sound = random.uniform(86, 98)

                run_hours += 0.1

                writer.writerow([
                    f"2026-05-07 10:{i:02d}:00",
                    round(current, 1),
                    round(pressure, 2),
                    round(temp, 1),
                    round(ambient, 1),
                    round(vib, 2),
                    round(sound, 1),
                    round(run_hours, 1),
                    start_count,
                    is_running
                ])

        self.csv_var.set(path)
        messagebox.showinfo("作成完了", "サンプルCSVを作成しました。")

    def start_connection(self):
        self.stop_connection()

        mode = self.mode_var.get()

        try:
            if mode == "simulation":
                self.device = SimulationDevice()

            elif mode == "modbus":
                host = self.ip_var.get().strip()
                port = int(self.port_var.get())
                unit = int(self.unit_var.get())
                self.device = ModbusTcpDevice(host, port, unit)

            elif mode == "csv":
                path = self.csv_var.get().strip()
                if not path:
                    messagebox.showwarning("CSV未選択", "CSVファイルを選択してください。")
                    return
                self.device = CsvSensorDevice(path)

            else:
                raise RuntimeError("不明なモードです。")

            ok = self.device.connect()
            if not ok:
                raise RuntimeError("接続に失敗しました。")

            self.running = True
            self.log(f"接続開始: {mode}")
            self.update_loop()

        except Exception as e:
            self.device = None
            self.running = False
            messagebox.showerror("接続エラー", str(e))
            self.log(f"接続エラー: {e}")

    def stop_connection(self):
        self.running = False
        if self.device:
            try:
                self.device.disconnect()
            except Exception:
                pass
        self.device = None
        self.log("停止しました。")

    def update_loop(self):
        if not self.running or not self.device:
            return

        try:
            state = self.device.read_state()
            faults = self.engine.diagnose(state)
            self.current_faults = faults

            self.update_values(state)
            self.update_faults(faults)

        except Exception as e:
            self.log(f"読取エラー: {e}")

        self.root.after(self.UPDATE_INTERVAL_MS, self.update_loop)

    def update_values(self, s: CompressorState):
        self.values["source"].set(s.source)
        self.values["is_running"].set("運転中" if s.is_running else "停止中")
        self.values["current_a"].set(f"{s.current_a:.1f}")
        self.values["pressure_bar"].set(f"{s.pressure_bar:.2f}")
        self.values["discharge_temp_c"].set(f"{s.discharge_temp_c:.1f}")
        self.values["ambient_temp_c"].set(f"{s.ambient_temp_c:.1f}")
        self.values["vibration_g"].set(f"{s.vibration_g:.2f}")
        self.values["sound_db"].set(f"{s.sound_db:.1f}")
        self.values["run_hours"].set(f"{s.run_hours:.1f}")
        self.values["start_count"].set(str(s.start_count))

    def update_faults(self, faults: List[FaultCode]):
        for item in self.tree.get_children():
            self.tree.delete(item)

        self.detail.delete("1.0", tk.END)

        if not faults:
            self.tree.insert("", "end", iid="no_fault", values=("-", "-", "異常は検出されていません。"))
            self.detail.insert(tk.END, "現在、診断上の異常は検出されていません。")
            return

        for i, f in enumerate(faults):
            self.tree.insert("", "end", iid=str(i), values=(f.code, f.severity, f.message))

    def show_fault_detail(self, event=None):
        sel = self.tree.selection()
        if not sel:
            return

        item_id = sel[0]
        if item_id == "no_fault":
            return

        try:
            idx = int(item_id)
            f = self.current_faults[idx]
        except Exception:
            return

        self.detail.delete("1.0", tk.END)
        self.detail.insert(tk.END, f"[{f.severity}] {f.code}\n")
        self.detail.insert(tk.END, f"{f.message}\n\n")

        self.detail.insert(tk.END, "■ 想定原因\n")
        for c in f.causes:
            self.detail.insert(tk.END, f"・{c}\n")

        self.detail.insert(tk.END, "\n■ 推奨点検・対処\n")
        for a in f.actions:
            self.detail.insert(tk.END, f"・{a}\n")

    def log(self, msg: str):
        now = time.strftime("%H:%M:%S")
        self.log_var.set(f"[{now}] {msg}")


# ==========================================================
# メイン
# ==========================================================

def main():
    root = tk.Tk()

    try:
        style = ttk.Style()
        if "clam" in style.theme_names():
            style.theme_use("clam")
    except Exception:
        pass

    app = CompressorDiagnosticApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()