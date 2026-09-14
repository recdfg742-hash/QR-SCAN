import os
import sys
import threading
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ==========================================
# 1. 모델 설정 및 매핑
# ==========================================
MODEL_CONFIG = {
    'S-FRONT': 'MPL02916AD',
    'S-REAR':  'MPL02915AD',
    'R-FRONT': 'MPL02926AD',
    'R-REAR':  'MPL02925AD'
}

# 역방향 매핑 (코드 -> 모델명)
CODE_TO_MODEL = {v: k for k, v in MODEL_CONFIG.items()}

DEFAULT_PASSWORD = "123456"

def get_base_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

BASE_DIR = get_base_dir()

class ScanStationApp:
    def __init__(self, root):
        self.root = root
        self.root.title("바코드/QR 실적 집계 시스템")
        self.root.geometry("1020x760")
        self.root.minsize(920, 680)

        self.current_model = tk.StringVar(value='S-FRONT')
        self.admin_password = DEFAULT_PASSWORD
        
        # 데이터 관리용 변수
        self.pending_items = []      # 라벨 QR 전까지 누적되는 단품 리스트
        self.scanned_history = set() # 중복 검사용 집합(Set)
        self.lock = threading.Lock() # 엑셀 동시 쓰기 방지 락

        self.setup_ui()
        self.on_model_changed()

    def setup_ui(self):
        # 상단 제어 바
        top_frame = tk.Frame(self.root, bg="#f0f2f5", pady=10, padx=15)
        top_frame.pack(fill=tk.X)

        tk.Label(top_frame, text="선택 모델:", font=("맑은 고딕", 12, "bold"), bg="#f0f2f5").pack(side=tk.LEFT)
        
        self.model_combo = ttk.Combobox(
            top_frame, 
            textvariable=self.current_model, 
            values=list(MODEL_CONFIG.keys()), 
            state="readonly", 
            font=("맑은 고딕", 12, "bold"), 
            width=12
        )
        self.model_combo.pack(side=tk.LEFT, padx=10)
        self.model_combo.bind("<<ComboboxSelected>>", self.on_model_changed)

        self.lbl_target_code = tk.Label(top_frame, text="", font=("맑은 고딕", 12), fg="#1a73e8", bg="#f0f2f5")
        self.lbl_target_code.pack(side=tk.LEFT, padx=10)

        self.lbl_pending_count = tk.Label(top_frame, text="[대기 단품: 0개]", font=("맑은 고딕", 11, "bold"), fg="#e65100", bg="#f0f2f5")
        self.lbl_pending_count.pack(side=tk.LEFT, padx=15)

        btn_change_pw = tk.Button(top_frame, text="비밀번호 변경", command=self.change_password_dialog)
        btn_change_pw.pack(side=tk.RIGHT)

        # 상태 대형 배너 (OK / NG / QR NG / READY)
        self.status_banner = tk.Label(
            self.root, 
            text="READY", 
            font=("Arial", 54, "bold"), 
            fg="white", 
            bg="#495057", 
            height=2
        )
        self.status_banner.pack(fill=tk.X, padx=20, pady=10)

        # 바코드 리더기 입력창
        scan_frame = tk.Frame(self.root, pady=6)
        scan_frame.pack(fill=tk.X, padx=20)

        tk.Label(scan_frame, text="QR 스캔 입력:", font=("맑은 고딕", 12, "bold")).pack(side=tk.LEFT)
        self.scan_entry = tk.Entry(scan_frame, font=("맑은 고딕", 13), width=60)
        self.scan_entry.pack(side=tk.LEFT, padx=10, fill=tk.X, expand=True)
        self.scan_entry.bind("<Return>", self.process_scan)
        self.scan_entry.focus_set()

        # 스캔 현황 테이블 (Treeview)
        table_frame = tk.Frame(self.root)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        columns = ("no", "timestamp", "type", "qr_code", "result")
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings", height=10)
        
        self.tree.heading("no", text="순번")
        self.tree.heading("timestamp", text="스캔 일시")
        self.tree.heading("type", text="구분")
        self.tree.heading("qr_code", text="스캔 QR 데이터")
        self.tree.heading("result", text="판정")

        self.tree.column("no", width=60, anchor="center")
        self.tree.column("timestamp", width=160, anchor="center")
        self.tree.column("type", width=100, anchor="center")
        self.tree.column("qr_code", width=520, anchor="w")
        self.tree.column("result", width=90, anchor="center")

        scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        lbl_info = tk.Label(
            self.root, 
            text="* 단품 QR들을 순서대로 스캔 후, Label QR을 스캔하면 즉시 하나의 Box로 묶여 엑셀에 기록됩니다.",
            font=("맑은 고딕", 9), 
            fg="#666666"
        )
        lbl_info.pack(side=tk.BOTTOM, pady=5)

    def on_model_changed(self, event=None):
        model = self.current_model.get()
        target_code = MODEL_CONFIG[model]
        self.lbl_target_code.config(text=f"[인식 코드: {target_code}]")
        self.scan_entry.focus_set()

    # ==========================================
    # 2. 바코드 판정 로직
    # ==========================================
    def process_scan(self, event=None):
        raw_code = self.scan_entry.get().strip()
        self.scan_entry.delete(0, tk.END)

        if not raw_code:
            return

        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        target_code = MODEL_CONFIG[self.current_model.get()].upper()
        scanned_prefix = raw_code[:10].upper()

        # 세미콜론이 3개 이상이면 박스 라벨 QR로 구분
        is_label_qr = (raw_code.count(';') >= 3)
        qr_type_str = "Label QR" if is_label_qr else "단품 QR"

        # ---------------- 1) 모델 코드 일치 여부 판정 ----------------
        if scanned_prefix != target_code:
            # 모델 불일치 NG
            self.status_banner.config(text="NG", bg="#dc3545")
            self.tree.insert("", 0, values=("-", now_str, qr_type_str, raw_code, "NG"))
            
            # 백그라운드로 NG 로그 기록
            threading.Thread(target=self.async_log_ng, args=(now_str, qr_type_str, raw_code, "NG(모델불일치)"), daemon=True).start()
            
            # 실제 스캔된 코드가 어느 모델에 속하는지 힌트 파악
            hint_model = CODE_TO_MODEL.get(scanned_prefix, "알 수 없는 모델")
            self.open_lock_dialog(
                title_text="⚠️ NG - 모델 코드 불일치",
                msg_text=(
                    f"[NG 발생: 선택 모델과 바코드 코드가 다릅니다]\n\n"
                    f"현재 선택 모델: {self.current_model.get()} ({target_code})\n"
                    f"스캔된 코드 접두: {raw_code[:10]}\n"
                    f"참고: 스캔된 코드는 [{hint_model}] 전용 코드입니다.\n\n"
                    f"관리자 비밀번호 6자리를 입력하여 해제하세요."
                ),
                header_bg="#ffebee",
                header_fg="#c62828"
            )
            return

        # ---------------- 2) 중복 스캔 검사 (단품 QR 대상) ----------------
        if not is_label_qr:
            if raw_code in self.scanned_history:
                # 중복 스캔 발생 -> QR NG 처리
                self.status_banner.config(text="QR NG", bg="#d9534f")
                self.tree.insert("", 0, values=("-", now_str, qr_type_str, raw_code, "QR NG"))
                
                threading.Thread(target=self.async_log_ng, args=(now_str, qr_type_str, raw_code, "QR NG(중복스캔)"), daemon=True).start()

                self.open_lock_dialog(
                    title_text="🚫 QR NG - 중복 바코드 스캔 감지",
                    msg_text=(
                        f"[QR NG 발생: 이미 스캔된 바코드입니다]\n\n"
                        f"스캔 바코드: {raw_code}\n"
                        f"동일 제품의 중복 스캔이 감지되었습니다.\n\n"
                        f"관리자 비밀번호 6자리를 입력하여 해제하세요."
                    ),
                    header_bg="#fff3e0",
                    header_fg="#e65100"
                )
                return

        # ---------------- 3) 정상 OK 처리 ----------------
        self.status_banner.config(text="OK", bg="#28a745")

        if not is_label_qr:
            # 단품 등록
            self.scanned_history.add(raw_code)
            self.pending_items.append({
                "timestamp": now_str,
                "code": raw_code,
                "result": "OK"
            })
            item_seq = len(self.pending_items)
            self.tree.insert("", 0, values=(item_seq, now_str, qr_type_str, raw_code, "OK"))
            self.lbl_pending_count.config(text=f"[대기 단품: {item_seq}개]")
        else:
            # Label QR 등록 -> 즉시 UI 화면 갱신 후 백그라운드 엑셀 저장 (지연 렉 제거)
            items_to_save = list(self.pending_items)
            self.pending_items.clear()
            self.lbl_pending_count.config(text="[대기 단품: 0개]")

            self.tree.insert("", 0, values=("-", now_str, "Label QR(완료)", raw_code, "OK"))
            self.root.update_idletasks() # UI 즉시 새로고침

            # 백그라운드 스레드에서 엑셀 그룹핑 저장 실행
            current_model_name = self.current_model.get()
            threading.Thread(
                target=self.async_save_group,
                args=(current_model_name, raw_code, now_str, items_to_save),
                daemon=True
            ).start()

        self.scan_entry.focus_set()

    # ==========================================
    # 3. 모달 잠금 팝업 (비밀번호 입력 창)
    # ==========================================
    def open_lock_dialog(self, title_text, msg_text, header_bg, header_fg):
        dialog = tk.Toplevel(self.root)
        dialog.title(title_text)
        dialog.geometry("500x300")
        dialog.resizable(False, False)
        dialog.configure(bg=header_bg)

        dialog.transient(self.root)
        dialog.grab_set()
        dialog.protocol("WM_DELETE_WINDOW", lambda: None) # 닫기 방지

        tk.Label(dialog, text=msg_text, font=("맑은 고딕", 11), bg=header_bg, fg=header_fg, justify=tk.LEFT).pack(pady=15)

        pw_entry = tk.Entry(dialog, show="*", font=("Arial", 16), justify="center", width=15)
        pw_entry.pack(pady=5)
        pw_entry.focus_set()

        lbl_error = tk.Label(dialog, text="", font=("맑은 고딕", 10, "bold"), fg="red", bg=header_bg)
        lbl_error.pack()

        def verify_password(event=None):
            if pw_entry.get() == self.admin_password:
                dialog.grab_release()
                dialog.destroy()
                self.status_banner.config(text="READY", bg="#495057")
                self.scan_entry.focus_set()
            else:
                lbl_error.config(text="비밀번호가 일치하지 않습니다. 다시 입력하세요.")
                pw_entry.delete(0, tk.END)

        pw_entry.bind("<Return>", verify_password)
        tk.Button(
            dialog, text="확인 및 잠금 해제", command=verify_password,
            font=("맑은 고딕", 11, "bold"), bg=header_fg, fg="white", padx=15, pady=5
        ).pack(pady=10)

    # ==========================================
    # 4. 비밀번호 변경 창
    # ==========================================
    def change_password_dialog(self):
        pw_win = tk.Toplevel(self.root)
        pw_win.title("관리자 비밀번호 변경")
        pw_win.geometry("320x210")
        pw_win.transient(self.root)
        pw_win.grab_set()

        tk.Label(pw_win, text="현재 비밀번호:").pack(pady=5)
        curr_entry = tk.Entry(pw_win, show="*")
        curr_entry.pack()

        tk.Label(pw_win, text="새 6자리 비밀번호:").pack(pady=5)
        new_entry = tk.Entry(pw_win, show="*")
        new_entry.pack()

        def apply_change():
            if curr_entry.get() != self.admin_password:
                messagebox.showerror("오류", "현재 비밀번호가 일치하지 않습니다.", parent=pw_win)
                return
            new_pw = new_entry.get()
            if len(new_pw) != 6 or not new_pw.isdigit():
                messagebox.showerror("오류", "새 비밀번호는 숫자 6자리여야 합니다.", parent=pw_win)
                return
            self.admin_password = new_pw
            messagebox.showinfo("성공", "비밀번호가 변경되었습니다.", parent=pw_win)
            pw_win.destroy()

        tk.Button(pw_win, text="변경 완료", command=apply_change, pady=4).pack(pady=15)

    # ==========================================
    # 5. 엑셀 워크북 초기화
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

            headers = ["Label QR (Box/Lot)", "Label 스캔일시", "단품 순번", "단품 QR", "단품 스캔일시", "판정"]
            ws.append(headers)

            header_fill = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
            header_font = Font(name="맑은 고딕", size=11, bold=True, color="FFFFFF")

            for col_num in range(1, len(headers) + 1):
                cell = ws.cell(row=1, column=col_num)
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center")

            ws.column_dimensions['A'].width = 46
            ws.column_dimensions['B'].width = 20
            ws.column_dimensions['C'].width = 12
            ws.column_dimensions['D'].width = 34
            ws.column_dimensions['E'].width = 20
            ws.column_dimensions['F'].width = 14

        return wb, ws, filepath

    # ==========================================
    # 6. 비동기 엑셀 묶음 저장 (렉 방지)
    # ==========================================
    def async_save_group(self, model_name, box_qr, box_time, items):
        with self.lock:
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
                        cell = ws.cell(row=start_row, column=col)
                        cell.border = thin_border
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                else:
                    for idx, item in enumerate(items, start=1):
                        curr_row = start_row + idx - 1
                        row_data = [box_qr, box_time, idx, item["code"], item["timestamp"], item["result"]]
                        ws.append(row_data)

                        for col in range(1, 7):
                            cell = ws.cell(row=curr_row, column=col)
                            cell.border = thin_border
                            cell.alignment = Alignment(horizontal="center" if col in [2, 3, 5, 6] else "left", vertical="center")
                            if col == 6:
                                cell.fill = ok_fill

                    # Label QR 및 시간 컬럼 셀 병합 (묶음 시각화)
                    if item_count > 1:
                        end_row = start_row + item_count - 1
                        ws.merge_cells(start_row=start_row, start_column=1, end_row=end_row, end_column=1)
                        ws.merge_cells(start_row=start_row, start_column=2, end_row=end_row, end_column=2)
                        ws.cell(row=start_row, column=1).alignment = Alignment(horizontal="left", vertical="center")
                        ws.cell(row=start_row, column=2).alignment = Alignment(horizontal="center", vertical="center")

                wb.save(filepath)
            except Exception as e:
                print(f"[저장 오류] 엑셀 파일이 열려있는지 확인하세요: {e}")

    # ==========================================
    # 7. NG 로그 비동기 저장
    # ==========================================
    def async_log_ng(self, timestamp, qr_type, raw_code, reason):
        with self.lock:
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
                    cell = ws.cell(row=start_row, column=col)
                    cell.border = thin_border
                    cell.fill = ng_fill
                    cell.alignment = Alignment(horizontal="center" if col in [1, 2, 3, 5, 6] else "left", vertical="center")

                wb.save(filepath)
            except Exception as e:
                print(f"[로그 기록 오류]: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = ScanStationApp(root)
    root.mainloop()
