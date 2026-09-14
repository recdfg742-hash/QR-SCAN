import os
import sys
import json
import threading
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ==========================================
# 1. 모델 매핑 및 기본 환경 설정
# ==========================================
MODEL_CONFIG = {
    'S-FRONT': 'MPL02916AD',
    'S-REAR':  'MPL02915AD',
    'R-FRONT': 'MPL02926AD',
    'R-REAR':  'MPL02925AD'
}

CODE_TO_MODEL = {v: k for k, v in MODEL_CONFIG.items()}
DEFAULT_PASSWORD = "123456"
MAX_ITEMS_PER_BOX = 10

def get_base_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

BASE_DIR = get_base_dir()
COUNT_FILE = os.path.join(BASE_DIR, "counts.json")

BG_MAIN = "#1a1f26"
BG_PANEL = "#222731"
BG_INPUT = "#15181e"
TEXT_COLOR = "#e1e4ea"
TEXT_MUTED = "#8b949e"
ACCENT_YELLOW = "#f59f00"
COLOR_OK = "#28a745"
COLOR_NG = "#dc3545"


class QRScanStationApp:
    def __init__(self, root):
        self.root = root
        self.root.title("QR SCAN STATION")
        self.root.geometry("1300x800")
        self.root.minsize(1160, 720)
        self.root.configure(bg=BG_MAIN)

        self.current_model = tk.StringVar(value='R-FRONT')
        self.admin_password = DEFAULT_PASSWORD
        self.model_session_id = 0

        self.model_counts = self.load_model_counts()
        self.scanned_history_by_model = {m: set() for m in MODEL_CONFIG}
        self.scanned_label_by_model = {m: set() for m in MODEL_CONFIG}

        self.pending_items = []
        self.pending_tree_ids = []
        self.file_lock = threading.Lock()
        self.global_scan_buffer = []

        self.setup_custom_styles()
        self.setup_ui()
        self.setup_global_key_listener()
        self.on_model_changed()

    def load_model_counts(self):
        default_counts = {m: {"total": 0, "ok": 0, "ng": 0} for m in MODEL_CONFIG}
        if os.path.exists(COUNT_FILE):
            try:
                with open(COUNT_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    for m in default_counts:
                        if m in data:
                            default_counts[m] = data[m]
                    return default_counts
            except Exception:
                return default_counts
        return default_counts

    def save_model_counts(self):
        try:
            with open(COUNT_FILE, "w", encoding="utf-8") as f:
                json.dump(self.model_counts, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"[카운트 저장 실패]: {e}")

    def setup_custom_styles(self):
        style = ttk.Style()
        style.theme_use("clam")

        style.configure("Dark.TNotebook", background=BG_MAIN, borderwidth=0)
        style.configure("Dark.TNotebook.Tab", background="#2a2f3a", foreground=TEXT_MUTED,
                        font=("맑은 고딕", 10, "bold"), padding=[15, 5])
        style.map("Dark.TNotebook.Tab",
                  background=[("selected", BG_PANEL)],
                  foreground=[("selected", "#ffffff")])

        style.configure("Dark.TCombobox", 
                        fieldbackground=BG_INPUT, 
                        background="#303642", 
                        foreground="#ffffff", 
                        arrowcolor="#ffffff",
                        darkcolor=BG_INPUT, 
                        lightcolor=BG_INPUT)

        style.configure("Dark.Treeview",
                        background="#1e232d",
                        foreground=TEXT_COLOR,
                        fieldbackground="#1e232d",
                        rowheight=26,
                        font=("맑은 고딕", 9))
        style.configure("Dark.Treeview.Heading",
                        background="#282e3a",
                        foreground="#adb5bd",
                        font=("맑은 고딕", 9, "bold"),
                        relief="flat")
        style.map("Dark.Treeview",
                  background=[("selected", "#2b5278")],
                  foreground=[("selected", "#ffffff")])
        style.map("Dark.Treeview.Heading",
                  background=[("active", "#343c4c")])

    def setup_ui(self):
        header_frame = tk.Frame(self.root, bg=BG_MAIN, height=45)
        header_frame.pack(fill=tk.X, padx=20, pady=(10, 4))

        tk.Label(header_frame, text="QR  SCAN  STATION", font=("Arial", 12, "bold"), 
                 fg=TEXT_COLOR, bg=BG_MAIN).pack(side=tk.LEFT, padx=(0, 15))

        self.model_combo = ttk.Combobox(
            header_frame, 
            textvariable=self.current_model, 
            values=list(MODEL_CONFIG.keys()), 
            state="readonly", 
            font=("맑은 고딕", 10, "bold"), 
            width=13,
            style="Dark.TCombobox"
        )
        self.model_combo.pack(side=tk.LEFT)
        self.model_combo.bind("<<ComboboxSelected>>", self.on_model_changed)

        btn_pw = tk.Button(
            header_frame, text="⚙ 비밀번호 설정", command=self.change_password_dialog,
            bg="#2c323d", fg=TEXT_COLOR, activebackground="#3a4250", activeforeground="#ffffff",
            relief="flat", font=("맑은 고딕", 9), padx=10, pady=3, cursor="hand2"
        )
        btn_pw.pack(side=tk.RIGHT)

        self.notebook = ttk.Notebook(self.root, style="Dark.TNotebook")
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=20, pady=(0, 10))

        self.tab_scan = tk.Frame(self.notebook, bg=BG_MAIN)
        self.notebook.add(self.tab_scan, text="  QR Scan  ")

        self.tab_recode = tk.Frame(self.notebook, bg=BG_MAIN)
        self.notebook.add(self.tab_recode, text="  Re-code  ")

        self.build_scan_tab()
        self.build_recode_tab()

    def build_scan_tab(self):
        main_frame = tk.Frame(self.tab_scan, bg=BG_MAIN)
        main_frame.pack(fill=tk.BOTH, expand=True, pady=10)

        # 좌측 컨트롤 패널 (320px -> 440px 대폭 확장)
        left_panel = tk.Frame(main_frame, bg=BG_PANEL, width=440)
        left_panel.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 15))
        left_panel.pack_propagate(False)

        self.lbl_model_info = tk.Label(
            left_panel, text="모델 R-FRONT\n인식코드 MPL02926AD", 
            font=("맑은 고딕", 11, "bold"), fg=ACCENT_YELLOW, bg=BG_PANEL, justify=tk.LEFT
        )
        self.lbl_model_info.pack(anchor="w", padx=20, pady=(15, 10))

        # 대형 상태 배너
        self.status_box = tk.Label(
            left_panel, text="READY", font=("Arial", 38, "bold"),
            fg="#adb5bd", bg="#2a2e37", height=3, relief="flat"
        )
        self.status_box.pack(fill=tk.X, padx=20, pady=6)

        self.lbl_last_scan = tk.Label(
            left_panel, text="마지막 스캔: -", font=("맑은 고딕", 9),
            fg=TEXT_MUTED, bg=BG_PANEL, anchor="w"
        )
        self.lbl_last_scan.pack(fill=tk.X, padx=20, pady=(10, 3))

        tk.Label(left_panel, text="바코드 스캔 입력 (어느 화면에서나 스캔 가능)", font=("맑은 고딕", 9),
                 fg=TEXT_MUTED, bg=BG_PANEL, anchor="w").pack(fill=tk.X, padx=20)

        self.scan_entry = tk.Entry(
            left_panel, font=("Consolas", 11), bg=BG_INPUT, fg="#ffffff",
            insertbackground="#ffffff", relief="flat", highlightthickness=1,
            highlightbackground="#343c4c", highlightcolor="#3b82f6"
        )
        self.scan_entry.pack(fill=tk.X, padx=20, pady=(4, 15), ipady=5)
        self.scan_entry.bind("<Return>", lambda e: self.process_scan(self.scan_entry.get()))

        stats_frame = tk.Frame(left_panel, bg=BG_PANEL)
        stats_frame.pack(fill=tk.X, padx=20, pady=3)
        stats_frame.columnconfigure((0, 1, 2), weight=1)

        card_total = tk.Frame(stats_frame, bg="#1a1e26", pady=6)
        card_total.grid(row=0, column=0, padx=2, sticky="nsew")
        self.lbl_total_val = tk.Label(card_total, text="0", font=("Arial", 16, "bold"), fg=TEXT_COLOR, bg="#1a1e26")
        self.lbl_total_val.pack()
        tk.Label(card_total, text="TOTAL", font=("Arial", 8, "bold"), fg=TEXT_MUTED, bg="#1a1e26").pack()

        card_ok = tk.Frame(stats_frame, bg="#1a1e26", pady=6)
        card_ok.grid(row=0, column=1, padx=2, sticky="nsew")
        self.lbl_ok_val = tk.Label(card_ok, text="0", font=("Arial", 16, "bold"), fg="#28a745", bg="#1a1e26")
        self.lbl_ok_val.pack()
        tk.Label(card_ok, text="OK", font=("Arial", 8, "bold"), fg=TEXT_MUTED, bg="#1a1e26").pack()

        card_ng = tk.Frame(stats_frame, bg="#1a1e26", pady=6)
        card_ng.grid(row=0, column=2, padx=2, sticky="nsew")
        self.lbl_ng_val = tk.Label(card_ng, text="0", font=("Arial", 16, "bold"), fg="#dc3545", bg="#1a1e26")
        self.lbl_ng_val.pack()
        tk.Label(card_ng, text="NG", font=("Arial", 8, "bold"), fg=TEXT_MUTED, bg="#1a1e26").pack()

        btn_reset = tk.Button(
            left_panel, text="RESET (카운터 초기화)", command=self.open_reset_dialog,
            bg="#2c323d", fg="#ff8787", activebackground="#3d2729", activeforeground="#ff6b6b",
            relief="flat", font=("맑은 고딕", 9, "bold"), pady=4, cursor="hand2"
        )
        btn_reset.pack(fill=tk.X, padx=20, pady=(8, 10))

        self.lbl_pending_status = tk.Label(
            left_panel, text="미그룹 스캔 0건 – Label QR 대기 중",
            font=("맑은 고딕", 9), fg=TEXT_MUTED, bg=BG_PANEL, anchor="w"
        )
        self.lbl_pending_status.pack(fill=tk.X, padx=20, pady=(5, 5))

        # 우측 테이블 패널 (너비 슬림화 적용)
        right_panel = tk.Frame(main_frame, bg=BG_MAIN)
        right_panel.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        self.lbl_right_header = tk.Label(
            right_panel, text="R-FRONT 기록",
            font=("맑은 고딕", 10, "bold"), fg="#9aa0a6", bg=BG_MAIN, anchor="w"
        )
        self.lbl_right_header.pack(fill=tk.X, pady=(0, 6))

        columns = ("DAY", "TIME", "Label QR", "DMC", "JUDGMENT")
        self.tree = ttk.Treeview(right_panel, columns=columns, show="headings", style="Dark.Treeview")

        self.tree.heading("DAY", text="DAY")
        self.tree.heading("TIME", text="TIME")
        self.tree.heading("Label QR", text="Label QR")
        self.tree.heading("DMC", text="DMC")
        self.tree.heading("JUDGMENT", text="JUDGMENT")

        # DAY, TIME, DMC 슬림화로 화면 밸런스 유지
        self.tree.column("DAY", width=85, anchor="center")
        self.tree.column("TIME", width=75, anchor="center")
        self.tree.column("Label QR", width=250, anchor="w")
        self.tree.column("DMC", width=240, anchor="w")
        self.tree.column("JUDGMENT", width=80, anchor="center")

        tree_scroll = ttk.Scrollbar(right_panel, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscroll=tree_scroll.set)

        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)

    def build_recode_tab(self):
        recode_frame = tk.Frame(self.tab_recode, bg=BG_MAIN)
        recode_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        filter_bar = tk.Frame(recode_frame, bg=BG_PANEL, pady=8, padx=15)
        filter_bar.pack(fill=tk.X, pady=(0, 10))

        today_str = datetime.now().strftime("%Y-%m-%d")

        tk.Label(filter_bar, text="조회 일자:", font=("맑은 고딕", 9, "bold"), fg=TEXT_COLOR, bg=BG_PANEL).pack(side=tk.LEFT)
        self.entry_start_day = tk.Entry(filter_bar, width=11, font=("맑은 고딕", 9), justify="center", bg=BG_INPUT, fg="#ffffff")
        self.entry_start_day.insert(0, today_str)
        self.entry_start_day.pack(side=tk.LEFT, padx=5)

        tk.Label(filter_bar, text="~", fg=TEXT_MUTED, bg=BG_PANEL).pack(side=tk.LEFT)
        self.entry_end_day = tk.Entry(filter_bar, width=11, font=("맑은 고딕", 9), justify="center", bg=BG_INPUT, fg="#ffffff")
        self.entry_end_day.insert(0, today_str)
        self.entry_end_day.pack(side=tk.LEFT, padx=5)

        tk.Label(filter_bar, text="시간대:", font=("맑은 고딕", 9, "bold"), fg=TEXT_COLOR, bg=BG_PANEL).pack(side=tk.LEFT, padx=(15, 0))
        self.entry_start_time = tk.Entry(filter_bar, width=9, font=("맑은 고딕", 9), justify="center", bg=BG_INPUT, fg="#ffffff")
        self.entry_start_time.insert(0, "00:00:00")
        self.entry_start_time.pack(side=tk.LEFT, padx=5)

        tk.Label(filter_bar, text="~", fg=TEXT_MUTED, bg=BG_PANEL).pack(side=tk.LEFT)
        self.entry_end_time = tk.Entry(filter_bar, width=9, font=("맑은 고딕", 9), justify="center", bg=BG_INPUT, fg="#ffffff")
        self.entry_end_time.insert(0, "23:59:59")
        self.entry_end_time.pack(side=tk.LEFT, padx=5)

        btn_search = tk.Button(
            filter_bar, text="🔍 검색", command=self.apply_recode_filter,
            bg="#2b5278", fg="#ffffff", relief="flat", font=("맑은 고딕", 9, "bold"), padx=12, pady=2, cursor="hand2"
        )
        btn_search.pack(side=tk.LEFT, padx=15)

        btn_save = tk.Button(
            filter_bar, text="💾 Save (엑셀 저장)", command=self.save_recode_to_excel,
            bg="#198754", fg="#ffffff", relief="flat", font=("맑은 고딕", 9, "bold"), padx=12, pady=2, cursor="hand2"
        )
        btn_save.pack(side=tk.RIGHT)

        cols = ("DAY", "TIME", "Label QR", "DMC", "JUDGMENT")
        self.tree_recode = ttk.Treeview(recode_frame, columns=cols, show="headings", style="Dark.Treeview")

        self.tree_recode.heading("DAY", text="DAY")
        self.tree_recode.heading("TIME", text="TIME")
        self.tree_recode.heading("Label QR", text="Label QR")
        self.tree_recode.heading("DMC", text="DMC")
        self.tree_recode.heading("JUDGMENT", text="JUDGMENT")

        self.tree_recode.column("DAY", width=85, anchor="center")
        self.tree_recode.column("TIME", width=75, anchor="center")
        self.tree_recode.column("Label QR", width=260, anchor="w")
        self.tree_recode.column("DMC", width=260, anchor="w")
        self.tree_recode.column("JUDGMENT", width=80, anchor="center")

        scroll_r = ttk.Scrollbar(recode_frame, orient=tk.VERTICAL, command=self.tree_recode.yview)
        self.tree_recode.configure(yscroll=scroll_r.set)

        self.tree_recode.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll_r.pack(side=tk.RIGHT, fill=tk.Y)

    def setup_global_key_listener(self):
        def _on_key_press(event):
            focused = self.root.focus_get()
            if isinstance(focused, tk.Entry) and focused != self.scan_entry:
                return

            if event.keysym in ("Return", "KP_Enter"):
                if self.global_scan_buffer:
                    scanned_text = "".join(self.global_scan_buffer).strip()
                    self.global_scan_buffer.clear()
                    self.scan_entry.delete(0, tk.END)
                    if scanned_text:
                        self.process_scan(scanned_text)
            elif event.char and event.char.isprintable():
                self.global_scan_buffer.append(event.char)

        self.root.bind_all("<Key>", _on_key_press)

    # 상태 배너 텍스트 표시 및 글자 크기 가변 조정 함수
    def set_status(self, text, fg_color, bg_color):
        if len(text) <= 2:  # OK, NG
            font_size = 46
        elif len(text) <= 6:  # READY, QR NG
            font_size = 36
        elif len(text) <= 11:  # Grouping NG
            font_size = 30
        else:  # Label QR NG 등 긴 텍스트
            font_size = 26

        self.status_box.config(text=text, fg=fg_color, bg=bg_color, font=("Arial", font_size, "bold"))

    # 팝업창 정중앙 배치 함수
    def center_popup(self, dialog, width, height):
        self.root.update_idletasks()
        rx = self.root.winfo_x()
        ry = self.root.winfo_y()
        rw = self.root.winfo_width()
        rh = self.root.winfo_height()

        x = rx + (rw - width) // 2
        y = ry + (rh - height) // 2
        dialog.geometry(f"{width}x{height}+{max(0, x)}+{max(0, y)}")

    # ==========================================
    # 4. 모델 변경 및 엑셀 로더
    # ==========================================
    def on_model_changed(self, event=None):
        self.model_session_id += 1
        current_session = self.model_session_id

        model = self.current_model.get()
        target_code = MODEL_CONFIG[model]

        self.lbl_model_info.config(text=f"모델 {model}\n인식코드 {target_code}")
        self.lbl_right_header.config(text=f"{model} 기록")
        
        self.pending_items.clear()
        self.pending_tree_ids.clear()
        self.lbl_pending_status.config(text="미그룹 스캔 0건 – Label QR 대기 중")

        self.update_stat_cards()
        self.load_history_from_excel(model, target_code, current_session)
        self.scan_entry.focus_set()

    def load_history_from_excel(self, model_name, target_code, session_id):
        self.tree.delete(*self.tree.get_children())

        filepath = os.path.join(BASE_DIR, f"{model_name}.xlsx")
        if not os.path.exists(filepath):
            return

        def _loader():
            try:
                wb = openpyxl.load_workbook(filepath, data_only=True)
                ws = wb.active

                rows_to_insert = []
                target_upper = target_code.upper()
                last_label = ""
                loaded_pending = []

                for row in ws.iter_rows(min_row=2, values_only=True):
                    if not row or len(row) < 6:
                        continue
                    lbl_val, box_time, seq_val, dmc_val, dmc_time, res = row[:6]

                    lbl_str = str(lbl_val).strip() if lbl_val and str(lbl_val).strip() != "-" else ""
                    dmc_str = str(dmc_val).strip() if dmc_val and str(dmc_val).strip() != "-" else ""
                    res_str = str(res).strip() if res else "OK"

                    if seq_val == "HEADER" or dmc_str.startswith("[박스 묶음 완료"):
                        t_parts = str(box_time).split()
                        day_val = t_parts[0] if len(t_parts) > 0 else ""
                        time_val = t_parts[1] if len(t_parts) > 1 else ""
                        rows_to_insert.append((day_val, time_val, lbl_str, dmc_str, "OK"))
                        if lbl_str:
                            self.scanned_label_by_model[model_name].add(lbl_str)
                        last_label = ""
                        loaded_pending.clear()
                        continue

                    ts_str = str(dmc_time if dmc_time and str(dmc_time).strip() != "-" else box_time)
                    if not ts_str or ts_str == "-":
                        continue

                    t_parts = ts_str.split()
                    day_val = t_parts[0] if len(t_parts) > 0 else ""
                    time_val = t_parts[1] if len(t_parts) > 1 else ""

                    if not dmc_str.upper().startswith(target_upper):
                        continue

                    if lbl_str:
                        last_label = lbl_str
                        final_label = lbl_str
                    else:
                        final_label = ""
                        if res_str == "OK":
                            loaded_pending.append({
                                "day": day_val, "time": time_val, "code": dmc_str, "result": "OK"
                            })

                    if final_label and final_label.upper().startswith(target_upper):
                        self.scanned_label_by_model[model_name].add(final_label)
                    if dmc_str and dmc_str.upper().startswith(target_upper) and res_str == "OK":
                        self.scanned_history_by_model[model_name].add(dmc_str)

                    rows_to_insert.append((day_val, time_val, final_label, dmc_str, res_str))

                if session_id != self.model_session_id:
                    return

                def _populate():
                    if session_id == self.model_session_id:
                        for r in reversed(rows_to_insert):
                            self.tree.insert("", tk.END, values=r)
                        
                        if loaded_pending:
                            self.pending_items = list(loaded_pending)
                            self.lbl_pending_status.config(text=f"미그룹 스캔 {len(self.pending_items)}건 – Label QR 대기 중")

                self.root.after(0, _populate)

            except Exception as e:
                print(f"[히스토리 로드 오류]: {e}")

        threading.Thread(target=_loader, daemon=True).start()

    # ==========================================
    # 5. 스캔 판정 및 실시간 저장
    # ==========================================
    def process_scan(self, raw_code):
        raw_code = raw_code.strip()
        self.scan_entry.delete(0, tk.END)
        self.global_scan_buffer.clear()

        if not raw_code:
            return

        now = datetime.now()
        day_str = now.strftime("%Y-%m-%d")
        time_str = now.strftime("%H:%M:%S")
        timestamp_full = f"{day_str} {time_str}"

        curr_model = self.current_model.get()
        target_code = MODEL_CONFIG[curr_model].upper()
        scanned_prefix = raw_code[:10].upper()

        is_label_qr = (raw_code.count(';') >= 3)
        self.lbl_last_scan.config(text=f"마지막 스캔: {raw_code}")

        # ---------------- [검증 1] 모델 코드 불일치 NG ----------------
        if scanned_prefix != target_code:
            self.set_status("NG", "#dc3545", "#3a1c1f")
            hint_model = CODE_TO_MODEL.get(scanned_prefix, "알 수 없음")
            self.open_lock_popup(
                title_text="⚠️ NG - 모델 불일치",
                msg=(
                    f"[NG: 선택 모델과 바코드 코드가 일치하지 않습니다]\n\n"
                    f"현재 선택 모델: {curr_model} ({target_code})\n"
                    f"스캔된 코드 접두: {raw_code[:10]}\n"
                    f"참고: 스캔된 코드는 [{hint_model}] 전용 코드입니다.\n\n"
                    f"관리자 비밀번호 6자리를 입력하여 해제하세요."
                ),
                header_bg="#2d1d20", header_fg="#f87171"
            )
            return

        # ---------------- [검증 2] Label QR 전용 검증 ----------------
        if is_label_qr:
            if raw_code in self.scanned_label_by_model[curr_model]:
                self.set_status("Label QR NG", "#dc3545", "#3a1c1f")
                self.open_lock_popup(
                    title_text="⚠️ Label QR NG - 중복 스캔",
                    msg=(
                        f"[Label QR NG: 이미 사용된 Label QR입니다]\n\n"
                        f"스캔 Label QR: {raw_code[:35]}...\n"
                        f"이미 등록/포장 완료된 중복 라벨입니다.\n\n"
                        f"관리자 비밀번호 6자리를 입력하여 해제하세요."
                    ),
                    header_bg="#2d1d20", header_fg="#f87171"
                )
                return

            tokens = raw_code.split(';')
            expected_qty = None
            if len(tokens) >= 2 and tokens[1].isdigit():
                expected_qty = int(tokens[1])

            current_scanned_qty = len(self.pending_items)

            if expected_qty is not None and expected_qty != current_scanned_qty:
                self.set_status("Grouping NG", "#dc3545", "#3a1c1f")
                self.open_lock_popup(
                    title_text="⚠️ Grouping NG - 수량 불일치",
                    msg=(
                        f"[Grouping NG: 단품 수량과 Label 포장 수량 불일치]\n\n"
                        f"Label QR 지정 수량: {expected_qty}개\n"
                        f"현재 스캔된 단품 수량: {current_scanned_qty}개\n\n"
                        f"수량이 일치하지 않아 묶음을 진행할 수 없습니다.\n"
                        f"관리자 비밀번호 6자리를 입력하여 해제하세요."
                    ),
                    header_bg="#2d1d20", header_fg="#f87171"
                )
                return

        # ---------------- [검증 3] 단품 QR 전용 체크 ----------------
        if not is_label_qr:
            if len(self.pending_items) >= MAX_ITEMS_PER_BOX:
                self.set_status("NG", "#dc3545", "#3a1c1f")
                self.open_lock_popup(
                    title_text="⚠️ NG - Label QR 미스캔",
                    msg=(
                        f"[NG 발생: Label QR 미스캔]\n\n"
                        f"단품이 이미 {MAX_ITEMS_PER_BOX}개 모두 스캔되었습니다.\n"
                        f"Label QR을 먼저 스캔하여 박스 묶음을 완료하십시오.\n\n"
                        f"관리자 비밀번호 6자리를 입력하여 해제하세요."
                    ),
                    header_bg="#2d1d20", header_fg="#f87171"
                )
                return

            # DMC 중복 스캔 시에만 NG 카운팅
            if raw_code in self.scanned_history_by_model[curr_model]:
                self.model_counts[curr_model]["ng"] += 1
                self.model_counts[curr_model]["total"] += 1
                self.save_model_counts()
                self.update_stat_cards()

                self.set_status("QR NG", "#fd7e14", "#3d2716")
                self.open_lock_popup(
                    title_text="🚫 QR NG - 중복 바코드 감지",
                    msg=(
                        f"[QR NG 발생: 이미 스캔된 바코드입니다]\n\n"
                        f"스캔 바코드: {raw_code}\n"
                        f"동일 제품의 중복 스캔이 감지되었습니다.\n\n"
                        f"관리자 비밀번호 6자리를 입력하여 해제하세요."
                    ),
                    header_bg="#352316", header_fg="#fb923c"
                )
                return

        # ---------------- [검증 통과] 정상 OK 처리 ----------------
        self.set_status("OK", "#28a745", "#193322")

        if not is_label_qr:
            self.model_counts[curr_model]["ok"] += 1
            self.model_counts[curr_model]["total"] += 1
            self.save_model_counts()
            self.update_stat_cards()

            self.scanned_history_by_model[curr_model].add(raw_code)
            item_data = {
                "day": day_str,
                "time": time_str,
                "code": raw_code,
                "result": "OK"
            }
            self.pending_items.append(item_data)
            
            item_id = self.tree.insert("", 0, values=(day_str, time_str, "", raw_code, "OK"))
            self.pending_tree_ids.append(item_id)
            self.lbl_pending_status.config(text=f"미그룹 스캔 {len(self.pending_items)}건 – Label QR 대기 중")

            # 단품 1개 즉시 엑셀 저장
            threading.Thread(
                target=self.async_append_single_item,
                args=(curr_model, item_data),
                daemon=True
            ).start()

        else:
            self.scanned_label_by_model[curr_model].add(raw_code)

            for t_id in self.pending_tree_ids:
                curr_vals = self.tree.item(t_id, "values")
                if curr_vals:
                    self.tree.item(t_id, values=(curr_vals[0], curr_vals[1], raw_code, curr_vals[3], curr_vals[4]))

            items_to_bundle = list(self.pending_items)
            bundle_count = len(items_to_bundle)
            header_text = f"[박스 묶음 완료: {bundle_count}건]"

            self.tree.insert("", 0, values=(day_str, time_str, raw_code, header_text, "OK"))

            self.pending_items.clear()
            self.pending_tree_ids.clear()
            self.lbl_pending_status.config(text="미그룹 스캔 0건 – Label QR 대기 중")
            self.root.update_idletasks()

            threading.Thread(
                target=self.async_finalize_excel_group,
                args=(curr_model, raw_code, timestamp_full, items_to_bundle, header_text),
                daemon=True
            ).start()

        self.scan_entry.focus_set()

    def update_stat_cards(self):
        curr_model = self.current_model.get()
        counts = self.model_counts.get(curr_model, {"total": 0, "ok": 0, "ng": 0})
        self.lbl_total_val.config(text=str(counts["total"]))
        self.lbl_ok_val.config(text=str(counts["ok"]))
        self.lbl_ng_val.config(text=str(counts["ng"]))

    def open_reset_dialog(self):
        win = tk.Toplevel(self.root)
        win.title("카운터 초기화")
        win.configure(bg=BG_PANEL)
        win.transient(self.root)
        win.grab_set()

        # 화면 정중앙 배치
        self.center_popup(win, 360, 210)

        curr_model = self.current_model.get()
        tk.Label(win, text=f"[{curr_model}] 카운터를 초기화하려면\n관리자 비밀번호를 입력하세요.", 
                 font=("맑은 고딕", 10, "bold"), fg=TEXT_COLOR, bg=BG_PANEL).pack(pady=(15, 8))

        pw_entry = tk.Entry(win, show="*", font=("Arial", 14), justify="center", bg=BG_INPUT, fg="#ffffff")
        pw_entry.pack(pady=5)
        pw_entry.focus_set()

        lbl_err = tk.Label(win, text="", font=("맑은 고딕", 9), fg="#ff6b6b", bg=BG_PANEL)
        lbl_err.pack()

        def do_reset(event=None):
            if pw_entry.get() == self.admin_password:
                self.model_counts[curr_model] = {"total": 0, "ok": 0, "ng": 0}
                self.save_model_counts()
                self.update_stat_cards()
                win.destroy()
                messagebox.showinfo("완료", f"{curr_model} 카운터가 초기화되었습니다.")
                self.scan_entry.focus_set()
            else:
                lbl_err.config(text="비밀번호가 올바르지 않습니다.")
                pw_entry.delete(0, tk.END)

        pw_entry.bind("<Return>", do_reset)
        tk.Button(win, text="초기화 실행", command=do_reset, bg="#dc3545", fg="#ffffff",
                  relief="flat", font=("맑은 고딕", 10, "bold"), padx=15, pady=3).pack(pady=10)

    def apply_recode_filter(self):
        self.tree_recode.delete(*self.tree_recode.get_children())
        model = self.current_model.get()
        target_upper = MODEL_CONFIG[model].upper()
        filepath = os.path.join(BASE_DIR, f"{model}.xlsx")

        if not os.path.exists(filepath):
            messagebox.showinfo("알림", f"{model}.xlsx 파일이 존재하지 않습니다.")
            return

        start_day = self.entry_start_day.get().strip()
        end_day = self.entry_end_day.get().strip()
        start_time = self.entry_start_time.get().strip()
        end_time = self.entry_end_time.get().strip()

        try:
            wb = openpyxl.load_workbook(filepath, data_only=True)
            ws = wb.active

            matched = []
            last_known_label_qr = ""

            for row in ws.iter_rows(min_row=2, values_only=True):
                if not row or len(row) < 6:
                    continue
                label_qr, box_time, seq_val, dmc_code, dmc_time, res = row[:6]

                if label_qr and str(label_qr).strip() != "-":
                    last_known_label_qr = str(label_qr).strip()

                ts_str = str(dmc_time if dmc_time and dmc_time != "-" else box_time)
                if not ts_str or ts_str == "-":
                    continue

                parts = ts_str.split()
                r_day = parts[0] if len(parts) > 0 else ""
                r_time = parts[1] if len(parts) > 1 else "00:00:00"

                if start_day and r_day < start_day:
                    continue
                if end_day and r_day > end_day:
                    continue
                if start_time and r_time < start_time:
                    continue
                if end_time and r_time > end_time:
                    continue

                lbl_val = last_known_label_qr
                dmc_val = "" if not dmc_code or dmc_code == "-" else str(dmc_code)

                if not (dmc_val.upper().startswith(target_upper) or lbl_val.upper().startswith(target_upper)):
                    continue

                matched.append((r_day, r_time, lbl_val, dmc_val, str(res)))

            for m in reversed(matched):
                self.tree_recode.insert("", tk.END, values=m)

            if not matched:
                messagebox.showinfo("조회 결과", "조건에 일치하는 데이터가 없습니다.")

        except Exception as e:
            messagebox.showerror("오류", f"데이터 조회 중 오류 발생: {e}")

    def save_recode_to_excel(self):
        items = self.tree_recode.get_children()
        if not items:
            messagebox.showwarning("경고", "저장할 조회 데이터가 없습니다. 먼저 검색을 수행하세요.")
            return

        now_ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        initial_name = f"{self.current_model.get()}_Recode_{now_ts}.xlsx"
        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
            initialfile=initial_name,
            title="조회 기록 별도 저장"
        )

        if not save_path:
            return

        try:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Re-code검색결과"

            headers = ["DAY", "TIME", "Label QR", "DMC", "판정"]
            ws.append(headers)

            header_fill = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
            header_font = Font(name="맑은 고딕", size=11, bold=True, color="FFFFFF")
            thin_border = Border(
                left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
                top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9')
            )

            for col in range(1, 6):
                cell = ws.cell(row=1, column=col)
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center")

            for idx, item_id in enumerate(items, start=2):
                row_vals = self.tree_recode.item(item_id, "values")
                ws.append(list(row_vals))
                for col in range(1, 6):
                    c = ws.cell(row=idx, column=col)
                    c.border = thin_border
                    c.alignment = Alignment(horizontal="center" if col in [1, 2, 5] else "left", vertical="center")

            ws.column_dimensions['A'].width = 14
            ws.column_dimensions['B'].width = 14
            ws.column_dimensions['C'].width = 46
            ws.column_dimensions['D'].width = 34
            ws.column_dimensions['E'].width = 12

            wb.save(save_path)
            messagebox.showinfo("성공", f"성공적으로 저장되었습니다:\n{save_path}")

        except Exception as e:
            messagebox.showerror("오류", f"파일 저장 실패: {e}")

    # ==========================================
    # 6. 정중앙 락 팝업 및 비밀번호 관리
    # ==========================================
    def open_lock_popup(self, title_text, msg, header_bg, header_fg):
        dialog = tk.Toplevel(self.root)
        dialog.title(title_text)
        dialog.resizable(False, False)
        dialog.configure(bg=header_bg)

        dialog.transient(self.root)
        dialog.grab_set()
        dialog.protocol("WM_DELETE_WINDOW", lambda: None)

        # 화면 정중앙 배치 (520x320)
        self.center_popup(dialog, 520, 320)

        tk.Label(dialog, text=msg, font=("맑은 고딕", 11), bg=header_bg, fg=header_fg, justify=tk.LEFT).pack(pady=15)

        pw_entry = tk.Entry(dialog, show="*", font=("Arial", 16), justify="center", width=14,
                            bg=BG_INPUT, fg="#ffffff", insertbackground="#ffffff")
        pw_entry.pack(pady=5)
        pw_entry.focus_set()

        lbl_err = tk.Label(dialog, text="", font=("맑은 고딕", 10, "bold"), fg="#ff6b6b", bg=header_bg)
        lbl_err.pack()

        def unlock(event=None):
            if pw_entry.get() == self.admin_password:
                dialog.grab_release()
                dialog.destroy()
                self.set_status("READY", "#adb5bd", "#2a2e37")
                self.scan_entry.focus_set()
            else:
                lbl_err.config(text="비밀번호가 올바르지 않습니다.")
                pw_entry.delete(0, tk.END)

        pw_entry.bind("<Return>", unlock)
        tk.Button(dialog, text="확인 및 잠금 해제", command=unlock,
                  font=("맑은 고딕", 11, "bold"), bg=header_fg, fg="#ffffff",
                  relief="flat", padx=16, pady=5, cursor="hand2").pack(pady=12)

    def change_password_dialog(self):
        win = tk.Toplevel(self.root)
        win.title("관리자 비밀번호 설정")
        win.configure(bg=BG_PANEL)
        win.transient(self.root)
        win.grab_set()

        # 화면 정중앙 배치
        self.center_popup(win, 340, 220)

        tk.Label(win, text="현재 비밀번호", font=("맑은 고딕", 9), fg=TEXT_COLOR, bg=BG_PANEL).pack(pady=(15, 2))
        curr_entry = tk.Entry(win, show="*", font=("Arial", 11), justify="center", bg=BG_INPUT, fg="#ffffff")
        curr_entry.pack()

        tk.Label(win, text="새 6자리 숫자 비밀번호", font=("맑은 고딕", 9), fg=TEXT_COLOR, bg=BG_PANEL).pack(pady=(10, 2))
        new_entry = tk.Entry(win, show="*", font=("Arial", 11), justify="center", bg=BG_INPUT, fg="#ffffff")
        new_entry.pack()

        def apply_pw():
            if curr_entry.get() != self.admin_password:
                messagebox.showerror("오류", "현재 비밀번호가 일치하지 않습니다.", parent=win)
                return
            new_val = new_entry.get()
            if len(new_val) != 6 or not new_val.isdigit():
                messagebox.showerror("오류", "비밀번호는 숫자 6자리로 지정해주세요.", parent=win)
                return
            self.admin_password = new_val
            messagebox.showinfo("성공", "관리자 비밀번호가 변경되었습니다.", parent=win)
            win.destroy()

        tk.Button(win, text="변경 적용", command=apply_pw, bg="#2b5278", fg="#ffffff",
                  relief="flat", font=("맑은 고딕", 10, "bold"), padx=15, pady=4).pack(pady=15)

    # ==========================================
    # 7. 엑셀 워크북 저장
    # ==========================================
    def get_or_create_workbook(self, filename):
        filepath = os.path.join(BASE_DIR, filename)
        if os.path.exists(filepath):
            wb = openpyxl.load_workbook(filepath)
            ws = wb.active
        else:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "스캔실적"

            headers = ["Label QR (Box/Lot)", "Label 스캔일시", "단품 순번", "단품 DMC", "단품 스캔일시", "판정"]
            ws.append(headers)

            header_fill = PatternFill(start_color="1F242D", end_color="1F242D", fill_type="solid")
            header_font = Font(name="맑은 고딕", size=11, bold=True, color="FFFFFF")

            for col in range(1, 7):
                cell = ws.cell(row=1, column=col)
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center")

            ws.column_dimensions['A'].width = 46
            ws.column_dimensions['B'].width = 20
            ws.column_dimensions['C'].width = 12
            ws.column_dimensions['D'].width = 34
            ws.column_dimensions['E'].width = 20
            ws.column_dimensions['F'].width = 16

        return wb, ws, filepath

    def async_append_single_item(self, model_name, item):
        with self.file_lock:
            try:
                filename = f"{model_name}.xlsx"
                wb, ws, filepath = self.get_or_create_workbook(filename)

                thin_border = Border(
                    left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
                    top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9')
                )
                ok_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")

                start_row = ws.max_row + 1
                ts_full = f"{item['day']} {item['time']}"
                row_data = ["-", "-", "-", item["code"], ts_full, item["result"]]
                ws.append(row_data)

                for col in range(1, 7):
                    c = ws.cell(row=start_row, column=col)
                    c.border = thin_border
                    c.alignment = Alignment(horizontal="center" if col in [2, 3, 5, 6] else "left", vertical="center")
                    if col == 6:
                        c.fill = ok_fill

                wb.save(filepath)
            except Exception as e:
                print(f"[단품 즉시 저장 실패]: {e}")

    def async_finalize_excel_group(self, model_name, box_qr, box_time, items, header_text):
        with self.file_lock:
            try:
                filename = f"{model_name}.xlsx"
                wb, ws, filepath = self.get_or_create_workbook(filename)

                item_codes = set(it["code"] for it in items)

                for row in reversed(list(ws.iter_rows(min_row=2, max_row=ws.max_row))):
                    dmc_val = str(row[3].value).strip() if row[3].value else ""
                    if dmc_val in item_codes:
                        row[0].value = box_qr
                        row[1].value = box_time
                        item_codes.remove(dmc_val)
                    if not item_codes:
                        break

                thin_border = Border(
                    left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
                    top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9')
                )

                header_row = [box_qr, box_time, "HEADER", header_text, box_time, "OK"]
                ws.append(header_row)
                h_row_idx = ws.max_row
                for col in range(1, 7):
                    c = ws.cell(row=h_row_idx, column=col)
                    c.border = thin_border
                    c.alignment = Alignment(horizontal="center" if col in [2, 3, 5, 6] else "left", vertical="center")

                wb.save(filepath)
            except Exception as e:
                print(f"[그룹핑 저장 실패]: {e}")


if __name__ == "__main__":
    root = tk.Tk()
    app = QRScanStationApp(root)
    root.mainloop()
