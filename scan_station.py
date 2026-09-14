import os
import sys
import threading
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ==========================================
# 1. 모델 매핑 및 기본 설정
# ==========================================
MODEL_CONFIG = {
    'S-FRONT': 'MPL02916AD',
    'S-REAR':  'MPL02915AD',
    'R-FRONT': 'MPL02926AD',
    'R-REAR':  'MPL02925AD'
}

CODE_TO_MODEL = {v: k for k, v in MODEL_CONFIG.items()}
DEFAULT_PASSWORD = "123456"
MAX_ITEMS_PER_BOX = 10  # 10개 초과 시 라벨 미스캔 락

def get_base_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

BASE_DIR = get_base_dir()

# 색상 테마 정의 (Dark Industrial Theme)
BG_MAIN = "#1a1f26"
BG_PANEL = "#222731"
BG_INPUT = "#15181e"
TEXT_COLOR = "#e1e4ea"
TEXT_MUTED = "#8b949e"
ACCENT_YELLOW = "#f59f00"
COLOR_OK = "#28a745"
COLOR_NG = "#dc3545"
COLOR_WARN = "#fd7e14"


class QRScanStationApp:
    def __init__(self, root):
        self.root = root
        self.root.title("QR SCAN STATION")
        self.root.geometry("1100x720")
        self.root.minsize(1020, 680)
        self.root.configure(bg=BG_MAIN)

        self.current_model = tk.StringVar(value='R-FRONT')
        self.admin_password = DEFAULT_PASSWORD
        
        # 실적 통계 카운터
        self.total_count = 0
        self.ok_count = 0
        self.ng_count = 0

        # 작업 대기 리스트
        self.pending_items = []
        self.scanned_history = set()
        self.file_lock = threading.Lock()

        self.setup_custom_styles()
        self.setup_ui()
        self.on_model_changed()

    def setup_custom_styles(self):
        style = ttk.Style()
        style.theme_use("clam")

        # 콤보박스 다크 스타일
        style.configure("Dark.TCombobox", 
                        fieldbackground=BG_INPUT, 
                        background="#303642", 
                        foreground="#ffffff", 
                        arrowcolor="#ffffff",
                        darkcolor=BG_INPUT, 
                        lightcolor=BG_INPUT)

        # 테이블(Treeview) 다크 스타일
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
        # ---------------- 1. 최상단 네비게이션 헤더 ----------------
        header_frame = tk.Frame(self.root, bg=BG_MAIN, height=50)
        header_frame.pack(fill=tk.X, padx=20, pady=(12, 5))

        tk.Label(header_frame, text="QR  SCAN  STATION", font=("Arial", 12, "bold"), 
                 fg=TEXT_COLOR, bg=BG_MAIN).pack(side=tk.LEFT, padx=(0, 15))

        self.model_combo = ttk.Combobox(
            header_frame, 
            textvariable=self.current_model, 
            values=list(MODEL_CONFIG.keys()), 
            state="readonly", 
            font=("맑은 고딕", 10, "bold"), 
            width=14,
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

        # ---------------- 2. 메인 컨텐츠 영역 (좌/우 분할) ----------------
        main_content = tk.Frame(self.root, bg=BG_MAIN)
        main_content.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        # ====== [좌측 패널: 작업 및 상태 표시] ======
        left_panel = tk.Frame(main_content, bg=BG_PANEL, width=320)
        left_panel.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 15))
        left_panel.pack_propagate(False)

        # 모델 정보 라벨
        self.lbl_model_info = tk.Label(
            left_panel, text="모델 R-FRONT\n인식코드 MPL02926AD", 
            font=("맑은 고딕", 11, "bold"), fg=ACCENT_YELLOW, bg=BG_PANEL, justify=tk.LEFT
        )
        self.lbl_model_info.pack(anchor="w", padx=20, pady=(20, 15))

        # OK / NG 대형 상태 배너
        self.status_box = tk.Label(
            left_panel, text="OK", font=("Arial", 46, "bold"),
            fg="#28a745", bg="#193322", width=10, height=3, relief="flat"
        )
        self.status_box.pack(fill=tk.X, padx=20, pady=10)

        # 마지막 스캔 표시
        self.lbl_last_scan = tk.Label(
            left_panel, text="마지막 스캔: -", font=("맑은 고딕", 9),
            fg=TEXT_MUTED, bg=BG_PANEL, anchor="w"
        )
        self.lbl_last_scan.pack(fill=tk.X, padx=20, pady=(15, 4))

        tk.Label(left_panel, text="바코드 스캔 입력 (포커스 유지)", font=("맑은 고딕", 9),
                 fg=TEXT_MUTED, bg=BG_PANEL, anchor="w").pack(fill=tk.X, padx=20)

        # 입력창 (스캐너 입력 대기)
        self.scan_entry = tk.Entry(
            left_panel, font=("Consolas", 11), bg=BG_INPUT, fg="#ffffff",
            insertbackground="#ffffff", relief="flat", highlightthickness=1,
            highlightbackground="#343c4c", highlightcolor="#3b82f6"
        )
        self.scan_entry.pack(fill=tk.X, padx=20, pady=(5, 20), ipady=5)
        self.scan_entry.bind("<Return>", self.process_scan)
        self.scan_entry.focus_set()

        # TOTAL / OK / NG 카운트 타일
        stats_frame = tk.Frame(left_panel, bg=BG_PANEL)
        stats_frame.pack(fill=tk.X, padx=20, pady=5)
        stats_frame.columnconfigure((0, 1, 2), weight=1)

        # TOTAL 카드
        card_total = tk.Frame(stats_frame, bg="#1a1e26", pady=8)
        card_total.grid(row=0, column=0, padx=2, sticky="nsew")
        self.lbl_total_val = tk.Label(card_total, text="0", font=("Arial", 16, "bold"), fg=TEXT_COLOR, bg="#1a1e26")
        self.lbl_total_val.pack()
        tk.Label(card_total, text="TOTAL", font=("Arial", 8, "bold"), fg=TEXT_MUTED, bg="#1a1e26").pack()

        # OK 카드
        card_ok = tk.Frame(stats_frame, bg="#1a1e26", pady=8)
        card_ok.grid(row=0, column=1, padx=2, sticky="nsew")
        self.lbl_ok_val = tk.Label(card_ok, text="0", font=("Arial", 16, "bold"), fg="#28a745", bg="#1a1e26")
        self.lbl_ok_val.pack()
        tk.Label(card_ok, text="OK", font=("Arial", 8, "bold"), fg=TEXT_MUTED, bg="#1a1e26").pack()

        # NG 카드
        card_ng = tk.Frame(stats_frame, bg="#1a1e26", pady=8)
        card_ng.grid(row=0, column=2, padx=2, sticky="nsew")
        self.lbl_ng_val = tk.Label(card_ng, text="0", font=("Arial", 16, "bold"), fg="#dc3545", bg="#1a1e26")
        self.lbl_ng_val.pack()
        tk.Label(card_ng, text="NG", font=("Arial", 8, "bold"), fg=TEXT_MUTED, bg="#1a1e26").pack()

        # 미그룹 단품 대기 상태 표시
        self.lbl_pending_status = tk.Label(
            left_panel, text="미그룹 스캔 0건 – Label QR 대기 중",
            font=("맑은 고딕", 9), fg=TEXT_MUTED, bg=BG_PANEL, anchor="w"
        )
        self.lbl_pending_status.pack(fill=tk.X, padx=20, pady=(20, 10))

        # ====== [우측 패널: 실적 테이블] ======
        right_panel = tk.Frame(main_content, bg=BG_MAIN)
        right_panel.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        # 우측 상단 정보 바
        self.lbl_right_header = tk.Label(
            right_panel, text="R-FRONT 기록     (저장 파일: R-FRONT.xlsx)",
            font=("맑은 고딕", 10, "bold"), fg="#9aa0a6", bg=BG_MAIN, anchor="w"
        )
        self.lbl_right_header.pack(fill=tk.X, pady=(0, 6))

        # 테이블
        columns = ("DAY", "TIME", "Label QR", "DMC")
        self.tree = ttk.Treeview(right_panel, columns=columns, show="headings", style="Dark.Treeview")

        self.tree.heading("DAY", text="DAY")
        self.tree.heading("TIME", text="TIME")
        self.tree.heading("Label QR", text="Label QR")
        self.tree.heading("DMC", text="DMC")

        self.tree.column("DAY", width=100, anchor="center")
        self.tree.column("TIME", width=90, anchor="center")
        self.tree.column("Label QR", width=250, anchor="w")
        self.tree.column("DMC", width=280, anchor="w")

        tree_scroll = ttk.Scrollbar(right_panel, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscroll=tree_scroll.set)

        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)

    # ==========================================
    # 2. 모델 변경 이벤트
    # ==========================================
    def on_model_changed(self, event=None):
        model = self.current_model.get()
        target_code = MODEL_CONFIG[model]

        self.lbl_model_info.config(text=f"모델 {model}\n인식코드 {target_code}")
        self.lbl_right_header.config(text=f"{model} 기록     (저장 파일: {model}.xlsx)")
        
        # 포커스 유지
        self.scan_entry.focus_set()

    # ==========================================
    # 3. 바코드 판정 및 스캔 로직
    # ==========================================
    def process_scan(self, event=None):
        raw_code = self.scan_entry.get().strip()
        self.scan_entry.delete(0, tk.END)

        if not raw_code:
            return

        now = datetime.now()
        day_str = now.strftime("%Y-%m-%d")
        time_str = now.strftime("%H:%M:%S")
        timestamp_full = f"{day_str} {time_str}"

        target_code = MODEL_CONFIG[self.current_model.get()].upper()
        scanned_prefix = raw_code[:10].upper()

        # 3개 이상의 세미콜론 -> 박스/로트용 롱 라벨 QR
        is_label_qr = (raw_code.count(';') >= 3)
        self.lbl_last_scan.config(text=f"마지막 스캔: {raw_code}")

        # ---------------- [검증 1] 모델 코드 불일치 NG ----------------
        if scanned_prefix != target_code:
            self.ng_count += 1
            self.total_count += 1
            self.update_stat_cards()

            self.status_box.config(text="NG", fg="#dc3545", bg="#3a1c1f")
            
            # 로그 비동기 저장
            qr_type = "Label QR" if is_label_qr else "단품 DMC"
            threading.Thread(target=self.async_log_ng, args=(timestamp_full, qr_type, raw_code, "NG(모델불일치)"), daemon=True).start()

            hint_model = CODE_TO_MODEL.get(scanned_prefix, "알 수 없음")
            self.open_lock_popup(
                title_text="⚠️ NG - 모델 불일치",
                msg=(
                    f"[NG: 선택 모델과 바코드 코드가 일치하지 않습니다]\n\n"
                    f"선택 모델: {self.current_model.get()} ({target_code})\n"
                    f"스캔 코드: {raw_code[:10]}\n"
                    f"참고: 스캔된 코드는 [{hint_model}] 전용 코드입니다.\n\n"
                    f"관리자 비밀번호 6자리를 입력하여 해제하세요."
                ),
                header_bg="#2d1d20", header_fg="#f87171"
            )
            return

        # ---------------- [검증 2] 단품 QR 전용 체크 ----------------
        if not is_label_qr:
            # 1) 10개 초과(11번째 이상) 연속 스캔 시 -> Label QR 미스캔 NG 발생
            if len(self.pending_items) >= MAX_ITEMS_PER_BOX:
                self.ng_count += 1
                self.total_count += 1
                self.update_stat_cards()

                self.status_box.config(text="NG", fg="#dc3545", bg="#3a1c1f")
                threading.Thread(target=self.async_log_ng, args=(timestamp_full, "단품 DMC", raw_code, "NG(Label QR 미스캔)"), daemon=True).start()

                self.open_lock_popup(
                    title_text="⚠️ NG - Label QR 미스캔",
                    msg=(
                        f"[NG 발생: Label QR 미스캔]\n\n"
                        f"이미 단품이 {MAX_ITEMS_PER_BOX}개 모두 스캔되었습니다.\n"
                        f"Label QR을 먼저 스캔하여 박스 묶음을 완료하십시오.\n\n"
                        f"관리자 비밀번호 6자리를 입력하여 해제하세요."
                    ),
                    header_bg="#2d1d20", header_fg="#f87171"
                )
                return

            # 2) 중복 바코드 스캔 -> QR NG 발생
            if raw_code in self.scanned_history:
                self.ng_count += 1
                self.total_count += 1
                self.update_stat_cards()

                self.status_box.config(text="QR NG", fg="#fd7e14", bg="#3d2716")
                threading.Thread(target=self.async_log_ng, args=(timestamp_full, "단품 DMC", raw_code, "QR NG(중복스캔)"), daemon=True).start()

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

        # ---------------- [검증 3] 정상 OK 처리 ----------------
        self.ok_count += 1
        self.total_count += 1
        self.update_stat_cards()
        self.status_box.config(text="OK", fg="#28a745", bg="#193322")

        if not is_label_qr:
            # 단품 리스트 및 테이블 추가 (Label QR은 대기 중 공란 처리)
            self.scanned_history.add(raw_code)
            item_data = {
                "day": day_str,
                "time": time_str,
                "code": raw_code,
                "result": "OK"
            }
            self.pending_items.append(item_data)
            
            # 우측 Treeview 추가 (최신 항목이 맨 위에 오도록 insert 0)
            self.tree.insert("", 0, values=(day_str, time_str, "", raw_code))
            self.lbl_pending_status.config(text=f"미그룹 스캔 {len(self.pending_items)}건 – Label QR 대기 중")

        else:
            # Label QR 스캔 시 -> 즉시 직전 단품들과 매핑하여 저장
            items_to_bundle = list(self.pending_items)
            self.pending_items.clear()
            self.lbl_pending_status.config(text="미그룹 스캔 0건 – Label QR 대기 중")

            # 우측 Treeview에 Label 완료 행 표기
            self.tree.insert("", 0, values=(day_str, time_str, raw_code, f"[박스 묶음 완료: {len(items_to_bundle)}건]"))
            self.root.update_idletasks()

            # 백그라운드 비동기 엑셀 저장
            curr_model = self.current_model.get()
            threading.Thread(
                target=self.async_save_excel_group,
                args=(curr_model, raw_code, timestamp_full, items_to_bundle),
                daemon=True
            ).start()

        self.scan_entry.focus_set()

    def update_stat_cards(self):
        self.lbl_total_val.config(text=str(self.total_count))
        self.lbl_ok_val.config(text=str(self.ok_count))
        self.lbl_ng_val.config(text=str(self.ng_count))

    # ==========================================
    # 4. 비밀번호 잠금 팝업
    # ==========================================
    def open_lock_popup(self, title_text, msg, header_bg, header_fg):
        dialog = tk.Toplevel(self.root)
        dialog.title(title_text)
        dialog.geometry("490x300")
        dialog.resizable(False, False)
        dialog.configure(bg=header_bg)

        dialog.transient(self.root)
        dialog.grab_set()
        dialog.protocol("WM_DELETE_WINDOW", lambda: None)  # X 버튼 닫기 무효화

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
                self.status_box.config(text="READY", fg="#adb5bd", bg="#2a2e37")
                self.scan_entry.focus_set()
            else:
                lbl_err.config(text="비밀번호가 올바르지 않습니다.")
                pw_entry.delete(0, tk.END)

        pw_entry.bind("<Return>", unlock)
        tk.Button(dialog, text="확인 및 잠금 해제", command=unlock,
                  font=("맑은 고딕", 11, "bold"), bg=header_fg, fg="#ffffff",
                  relief="flat", padx=16, pady=5, cursor="hand2").pack(pady=12)

    # ==========================================
    # 5. 관리자 비밀번호 변경 창
    # ==========================================
    def change_password_dialog(self):
        win = tk.Toplevel(self.root)
        win.title("관리자 비밀번호 설정")
        win.geometry("320x210")
        win.configure(bg=BG_PANEL)
        win.transient(self.root)
        win.grab_set()

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
    # 6. 엑셀 워크북 제어 및 비동기 저장
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

    def async_save_excel_group(self, model_name, box_qr, box_time, items):
        with self.file_lock:
            try:
                filename = f"{model_name}.xlsx"
                wb, ws, filepath = self.get_or_create_workbook(filename)

                start_row = ws.max_row + 1
                item_count = len(items)

                thin_border = Border(
                    left=Side(style='thin', color='D9D9D9'),
                    right=Side(style='thin', color='D9D9D9'),
                    top=Side(style='thin', color='D9D9D9'),
                    bottom=Side(style='thin', color='D9D9D9')
                )
                ok_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")

                if item_count == 0:
                    row_data = [box_qr, box_time, "-", "-", "-", "OK"]
                    ws.append(row_data)
                    for col in range(1, 7):
                        c = ws.cell(row=start_row, column=col)
                        c.border = thin_border
                        c.alignment = Alignment(horizontal="center", vertical="center")
                else:
                    for idx, item in enumerate(items, start=1):
                        curr_row = start_row + idx - 1
                        ts_full = f"{item['day']} {item['time']}"
                        row_data = [box_qr, box_time, idx, item["code"], ts_full, item["result"]]
                        ws.append(row_data)

                        for col in range(1, 7):
                            c = ws.cell(row=curr_row, column=col)
                            c.border = thin_border
                            c.alignment = Alignment(horizontal="center" if col in [2, 3, 5, 6] else "left", vertical="center")
                            if col == 6:
                                c.fill = ok_fill

                    # 박스 라벨 및 시간 열 병합 (Grouping)
                    if item_count > 1:
                        end_row = start_row + item_count - 1
                        ws.merge_cells(start_row=start_row, start_column=1, end_row=end_row, end_column=1)
                        ws.merge_cells(start_row=start_row, start_column=2, end_row=end_row, end_column=2)
                        ws.cell(row=start_row, column=1).alignment = Alignment(horizontal="left", vertical="center")
                        ws.cell(row=start_row, column=2).alignment = Alignment(horizontal="center", vertical="center")

                wb.save(filepath)
            except Exception as e:
                print(f"[엑셀 저장 실패] 파일이 열려있는지 확인하세요: {e}")

    def async_log_ng(self, timestamp, qr_type, raw_code, reason):
        with self.file_lock:
            try:
                filename = f"{self.current_model.get()}.xlsx"
                wb, ws, filepath = self.get_or_create_workbook(filename)

                start_row = ws.max_row + 1
                ng_fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
                thin_border = Border(
                    left=Side(style='thin', color='D9D9D9'),
                    right=Side(style='thin', color='D9D9D9'),
                    top=Side(style='thin', color='D9D9D9'),
                    bottom=Side(style='thin', color='D9D9D9')
                )

                row_data = ["-", "-", f"[{reason}]", raw_code, timestamp, "NG"]
                ws.append(row_data)

                for col in range(1, 7):
                    c = ws.cell(row=start_row, column=col)
                    c.border = thin_border
                    c.fill = ng_fill
                    c.alignment = Alignment(horizontal="center" if col in [1, 2, 3, 5, 6] else "left", vertical="center")

                wb.save(filepath)
            except Exception as e:
                print(f"[NG 로그 저장 오류]: {e}")


if __name__ == "__main__":
    root = tk.Tk()
    app = QRScanStationApp(root)
    root.mainloop()
