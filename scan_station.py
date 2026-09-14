# -*- coding: utf-8 -*-
"""
QR SCAN STATION
- 모델 선택(S-FRONT / S-REAR / R-FRONT / R-REAR) 후 바코드 스캔
- 앞 10자리 인식코드 일치 여부로 OK/NG 판정
- Label QR(필드 6개 이상) 스캔 시 그 이전 미그룹 DMC들을 하나로 묶음
- 모델별로 별도의 엑셀 파일(S-FRONT.xlsx 등)에 실시간 저장
- NG 발생 시 6자리 비밀번호를 입력해야 잠금 해제
"""

import json
import os
import sys
import tkinter as tk
from tkinter import ttk
from datetime import datetime
from pathlib import Path

try:
    import openpyxl
    from openpyxl.styles import Font, Alignment, PatternFill
    from openpyxl.utils import get_column_letter
except ImportError:
    raise SystemExit(
        "openpyxl 모듈이 필요합니다.\n"
        "명령 프롬프트에서 아래 명령을 실행한 뒤 다시 실행하세요:\n"
        "    pip install openpyxl"
    )

# EXE로 빌드된 경우 실행 파일이 있는 폴더에, 아니면 이 스크립트가 있는 폴더에 데이터를 저장합니다.
if getattr(sys, "frozen", False):
    DATA_DIR = Path(sys.executable).parent
else:
    DATA_DIR = Path(os.path.dirname(os.path.abspath(__file__)))

CONFIG_PATH = DATA_DIR / "scan_station_config.json"

MODELS = {
    "S-FRONT": "MPL02916AD",
    "S-REAR": "MPL02915AD",
    "R-FRONT": "MPL02926AD",
    "R-REAR": "MPL02925AD",
}

DEFAULT_PASSWORD = "000000"
HEADERS = ["DAY", "TIME", "Label QR", "DMC", "RESULT"]

# ---------- 설정(비밀번호) ----------

def load_config():
    if CONFIG_PATH.exists():
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
                if "password" in data:
                    return data
        except Exception:
            pass
    data = {"password": DEFAULT_PASSWORD}
    save_config(data)
    return data


def save_config(data):
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


# ---------- 엑셀 입출력 ----------

def excel_path(model):
    return DATA_DIR / f"{model}.xlsx"


def load_records(model):
    """모델의 엑셀 파일에서 기존 기록을 불러오고, 병합된 Label QR 셀은 값을 채워서 복원합니다."""
    path = excel_path(model)
    records = []
    if not path.exists():
        return records
    try:
        wb = openpyxl.load_workbook(path)
        ws = wb.active
        merge_map = {}
        for merged in ws.merged_cells.ranges:
            if merged.min_col == 3 and merged.max_col == 3:
                top_value = ws.cell(row=merged.min_row, column=3).value
                for r in range(merged.min_row, merged.max_row + 1):
                    merge_map[r] = top_value
        for row in range(2, ws.max_row + 1):
            day = ws.cell(row=row, column=1).value
            time_ = ws.cell(row=row, column=2).value
            label = merge_map.get(row, ws.cell(row=row, column=3).value)
            dmc = ws.cell(row=row, column=4).value
            result = ws.cell(row=row, column=5).value
            if day is None and dmc is None:
                continue
            records.append({
                "day": day or "",
                "time": time_ or "",
                "label": label or "",
                "dmc": dmc or "",
                "result": result or "",
            })
    except Exception as e:
        print("엑셀 로드 오류:", e)
    return records


def compute_groups(records):
    groups = []
    for r in records:
        if groups and groups[-1]["label"] == r["label"]:
            groups[-1]["items"].append(r)
        else:
            groups.append({"label": r["label"], "items": [r]})
    return groups


def save_records(model, records):
    path = excel_path(model)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = model
    header_fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
    for c, h in enumerate(HEADERS, start=1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.font = Font(bold=True)
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")

    row_idx = 2
    groups = compute_groups(records)
    for g in groups:
        start_row = row_idx
        for i, it in enumerate(g["items"]):
            ws.cell(row=row_idx, column=1, value=it["day"])
            ws.cell(row=row_idx, column=2, value=it["time"])
            ws.cell(row=row_idx, column=3, value=g["label"] if i == 0 else None)
            ws.cell(row=row_idx, column=4, value=it["dmc"])
            result_cell = ws.cell(row=row_idx, column=5, value=it["result"])
            if it["result"] == "NG":
                result_cell.font = Font(color="C00000", bold=True)
            else:
                result_cell.font = Font(color="1E7B34", bold=True)
            row_idx += 1
        if g["label"] and len(g["items"]) > 1:
            ws.merge_cells(start_row=start_row, start_column=3, end_row=row_idx - 1, end_column=3)
            ws.cell(row=start_row, column=3).alignment = Alignment(vertical="center", wrap_text=True)

    widths = [12, 10, 55, 26, 9]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    wb.save(path)


# ---------- GUI ----------

BG = "#1b1f22"
PANEL = "#23282c"
PANEL2 = "#2a3035"
LINE = "#3a4147"
TEXT = "#e7ebee"
TEXT_DIM = "#93a0a8"
OK_C = "#35c471"
OK_BG = "#1d3a2a"
NG_C = "#e5484d"
NG_BG = "#3f1f21"
AMBER = "#dfa538"
MONO = "Consolas"


class ScanStationApp:
    def __init__(self, root):
        self.root = root
        self.root.title("QR SCAN STATION")
        self.root.geometry("1080x680")
        self.root.configure(bg=BG)

        self.config_data = load_config()
        self.model_var = tk.StringVar(value="")
        self.records = []
        self.locked = False

        self._build_style()
        self._build_topbar()
        self.body = tk.Frame(self.root, bg=BG)
        self.body.pack(fill="both", expand=True, padx=16, pady=(0, 16))
        self.render_body()

    def _build_style(self):
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure("TCombobox", fieldbackground=PANEL2, background=PANEL2, foreground=TEXT)
        style.configure("Treeview", background="#20242a", fieldbackground="#20242a",
                         foreground=TEXT, rowheight=24, bordercolor=LINE, borderwidth=0)
        style.configure("Treeview.Heading", background=PANEL2, foreground=TEXT_DIM, relief="flat")
        style.map("Treeview", background=[("selected", LINE)])

    def _build_topbar(self):
        bar = tk.Frame(self.root, bg=BG)
        bar.pack(fill="x", padx=16, pady=14)

        left = tk.Frame(bar, bg=BG)
        left.pack(side="left")

        tk.Label(left, text="QR SCAN STATION", bg=BG, fg=TEXT_DIM,
                 font=(MONO, 11, "bold")).pack(side="left", padx=(0, 12))

        combo = ttk.Combobox(left, textvariable=self.model_var, values=list(MODELS.keys()),
                              state="readonly", width=14, font=(MONO, 11, "bold"))
        combo.pack(side="left")
        combo.bind("<<ComboboxSelected>>", self.on_model_change)

        tk.Button(bar, text="⚙ 비밀번호 설정", command=self.open_settings,
                  bg=PANEL2, fg=TEXT_DIM, relief="flat", padx=12, pady=6,
                  activebackground=LINE, activeforeground=TEXT).pack(side="right")

    def render_body(self):
        for w in self.body.winfo_children():
            w.destroy()

        model = self.model_var.get()
        if not model:
            tk.Label(self.body, text="모델을 선택하면 검사를 시작합니다",
                     bg=BG, fg=TEXT, font=(MONO, 14, "bold")).pack(expand=True)
            return

        # 좌측 상태 패널
        left = tk.Frame(self.body, bg=PANEL, width=320)
        left.pack(side="left", fill="y", padx=(0, 14))
        left.pack_propagate(False)

        code = MODELS[model]
        tk.Label(left, text=f"모델 {model}\n인식코드 {code}", bg=PANEL, fg=AMBER,
                 font=(MONO, 10, "bold"), justify="left").pack(anchor="w", padx=16, pady=(16, 10))

        self.result_label = tk.Label(left, text="—", bg=PANEL2, fg="#4a5158",
                                      font=(MONO, 40, "bold"), width=8, height=3)
        self.result_label.pack(padx=16, pady=6)

        self.last_scan_label = tk.Label(left, text="스캔 대기 중...", bg=PANEL, fg=TEXT_DIM,
                                         font=(MONO, 9), wraplength=280, justify="left")
        self.last_scan_label.pack(anchor="w", padx=16, pady=(2, 10))

        tk.Label(left, text="바코드 스캔 입력 (포커스 유지)", bg=PANEL, fg=TEXT_DIM,
                 font=(MONO, 9)).pack(anchor="w", padx=16)
        self.scan_entry = tk.Entry(left, bg=PANEL2, fg=TEXT, insertbackground=TEXT,
                                    font=(MONO, 12), relief="flat")
        self.scan_entry.pack(fill="x", padx=16, pady=(2, 14), ipady=4)
        self.scan_entry.bind("<Return>", self.on_scan_enter)
        self.scan_entry.focus_set()

        stats = tk.Frame(left, bg=PANEL)
        stats.pack(fill="x", padx=16)
        self.stat_total = self._stat_box(stats, "TOTAL", TEXT)
        self.stat_ok = self._stat_box(stats, "OK", OK_C)
        self.stat_ng = self._stat_box(stats, "NG", NG_C)

        self.open_hint = tk.Label(left, text="", bg=PANEL, fg=TEXT_DIM, font=(MONO, 8),
                                   wraplength=280, justify="left")
        self.open_hint.pack(anchor="w", padx=16, pady=(10, 16))

        # 우측 기록 테이블
        right = tk.Frame(self.body, bg=PANEL)
        right.pack(side="left", fill="both", expand=True)

        head = tk.Frame(right, bg=PANEL)
        head.pack(fill="x", padx=16, pady=(16, 6))
        tk.Label(head, text=f"{model} 기록   (저장 파일: {model}.xlsx)", bg=PANEL, fg=TEXT_DIM,
                 font=(MONO, 10)).pack(side="left")

        columns = ("day", "time", "label", "dmc", "result")
        self.tree = ttk.Treeview(right, columns=columns, show="headings", height=20)
        widths = {"day": 90, "time": 70, "label": 380, "dmc": 200, "result": 70}
        headers = {"day": "DAY", "time": "TIME", "label": "Label QR", "dmc": "DMC", "result": "RESULT"}
        for c in columns:
            self.tree.heading(c, text=headers[c])
            self.tree.column(c, width=widths[c], anchor="w")
        self.tree.tag_configure("ok", foreground=OK_C)
        self.tree.tag_configure("ng", foreground=NG_C)
        self.tree.pack(fill="both", expand=True, padx=16, pady=(0, 16))

        self.records = load_records(model)
        self.refresh_view()

    def _stat_box(self, parent, label, color):
        box = tk.Frame(parent, bg=PANEL2)
        box.pack(side="left", expand=True, fill="x", padx=3)
        n = tk.Label(box, text="0", bg=PANEL2, fg=color, font=(MONO, 16, "bold"))
        n.pack(pady=(8, 0))
        tk.Label(box, text=label, bg=PANEL2, fg=TEXT_DIM, font=(MONO, 8)).pack(pady=(0, 8))
        return n

    def on_model_change(self, event=None):
        self.locked = False
        self.render_body()

    def refresh_view(self):
        self.tree.delete(*self.tree.get_children())
        groups = compute_groups(self.records)
        for g in groups:
            for i, it in enumerate(g["items"]):
                label_disp = g["label"] if i == 0 else ""
                tag = "ok" if it["result"] == "OK" else "ng"
                self.tree.insert("", "end",
                                  values=(it["day"], it["time"], label_disp, it["dmc"], it["result"]),
                                  tags=(tag,))
        children = self.tree.get_children()
        if children:
            self.tree.see(children[-1])

        total = len(self.records)
        ok = sum(1 for r in self.records if r["result"] == "OK")
        ng = total - ok
        open_cnt = sum(1 for r in self.records if not r["label"])
        self.stat_total.config(text=str(total))
        self.stat_ok.config(text=str(ok))
        self.stat_ng.config(text=str(ng))
        self.open_hint.config(text=f"미그룹 스캔 {open_cnt}건 — Label QR 대기 중" if open_cnt else "")

    def on_scan_enter(self, event=None):
        if self.locked:
            return
        value = self.scan_entry.get().strip()
        self.scan_entry.delete(0, "end")
        if not value:
            return
        self.process_scan(value)

    def process_scan(self, value):
        model = self.model_var.get()
        code = MODELS[model]
        fields = value.split(";")
        is_label = len(fields) > 2
        prefix10 = value[:10].upper()
        match = prefix10 == code
        result = "OK" if match else "NG"
        now = datetime.now()
        day = now.strftime("%Y-%m-%d")
        time_ = now.strftime("%H:%M:%S")

        if is_label:
            if match:
                for r in self.records:
                    if not r["label"]:
                        r["label"] = value
                save_records(model, self.records)
                self.refresh_view()
        else:
            self.records.append({"day": day, "time": time_, "label": "", "dmc": value, "result": result})
            save_records(model, self.records)
            self.refresh_view()

        self.last_scan_label.config(text=f"마지막 스캔: {value}")
        if result == "OK":
            self.result_label.config(text="OK", bg=OK_BG, fg=OK_C)
        else:
            self.result_label.config(text="NG", bg=NG_BG, fg=NG_C)

        if result == "NG":
            self.lock()

    def lock(self):
        self.locked = True
        self.scan_entry.config(state="disabled")
        self.open_password_dialog()

    def open_password_dialog(self):
        dlg = tk.Toplevel(self.root)
        dlg.title("NG - 비밀번호 입력")
        dlg.configure(bg=PANEL)
        dlg.geometry("300x230")
        dlg.resizable(False, False)
        dlg.transient(self.root)
        dlg.grab_set()
        dlg.protocol("WM_DELETE_WINDOW", lambda: None)

        tk.Label(dlg, text="NG", bg=PANEL, fg=NG_C, font=(MONO, 26, "bold")).pack(pady=(18, 4))
        tk.Label(dlg, text="인식 코드가 일치하지 않습니다.\n비밀번호 6자리를 입력하세요.",
                 bg=PANEL, fg=TEXT_DIM, font=(MONO, 9), justify="center").pack(pady=(0, 10))

        pw_var = tk.StringVar()
        entry = tk.Entry(dlg, textvariable=pw_var, show="●", justify="center",
                          font=(MONO, 16), bg=PANEL2, fg=TEXT, relief="flat")
        entry.pack(pady=4, ipady=4, padx=30, fill="x")
        entry.focus_set()

        err = tk.Label(dlg, text="", bg=PANEL, fg=NG_C, font=(MONO, 9))
        err.pack(pady=4)

        def try_unlock(event=None):
            if pw_var.get() == self.config_data.get("password", DEFAULT_PASSWORD):
                self.locked = False
                self.scan_entry.config(state="normal")
                dlg.grab_release()
                dlg.destroy()
                self.scan_entry.focus_set()
            else:
                err.config(text="비밀번호가 올바르지 않습니다")
                pw_var.set("")

        entry.bind("<Return>", try_unlock)
        tk.Button(dlg, text="확인", command=try_unlock, bg=NG_C, fg="white",
                  relief="flat", padx=10, pady=6).pack(pady=8)

    def open_settings(self):
        dlg = tk.Toplevel(self.root)
        dlg.title("비밀번호 설정")
        dlg.configure(bg=PANEL)
        dlg.geometry("300x360")
        dlg.resizable(False, False)
        dlg.transient(self.root)
        dlg.grab_set()

        tk.Label(dlg, text="비밀번호 설정", bg=PANEL, fg=TEXT, font=(MONO, 12, "bold")).pack(pady=(16, 10))

        def labeled_entry(text):
            tk.Label(dlg, text=text, bg=PANEL, fg=TEXT_DIM, font=(MONO, 9)).pack(anchor="w", padx=24)
            var = tk.StringVar()
            e = tk.Entry(dlg, textvariable=var, show="●", justify="center", font=(MONO, 13),
                         bg=PANEL2, fg=TEXT, relief="flat")
            e.pack(padx=24, pady=(2, 10), fill="x", ipady=3)
            return var

        cur_var = labeled_entry("현재 비밀번호")
        new1_var = labeled_entry("새 비밀번호 (숫자 6자리)")
        new2_var = labeled_entry("새 비밀번호 확인")

        err = tk.Label(dlg, text="", bg=PANEL, fg=NG_C, font=(MONO, 9), wraplength=250)
        err.pack(pady=4)

        def save_pw():
            cur = cur_var.get()
            n1 = new1_var.get()
            n2 = new2_var.get()
            if cur != self.config_data.get("password", DEFAULT_PASSWORD):
                err.config(fg=NG_C, text="현재 비밀번호가 올바르지 않습니다")
                return
            if not (n1.isdigit() and len(n1) == 6):
                err.config(fg=NG_C, text="새 비밀번호는 숫자 6자리여야 합니다")
                return
            if n1 != n2:
                err.config(fg=NG_C, text="새 비밀번호가 일치하지 않습니다")
                return
            self.config_data["password"] = n1
            save_config(self.config_data)
            err.config(fg=OK_C, text="저장되었습니다")

        tk.Button(dlg, text="저장", command=save_pw, bg=AMBER, fg="#241a05",
                  relief="flat", padx=10, pady=6).pack(pady=8)
        tk.Button(dlg, text="닫기", command=dlg.destroy, bg=PANEL2, fg=TEXT_DIM,
                  relief="flat", padx=10, pady=6).pack()


def main():
    root = tk.Tk()
    ScanStationApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
