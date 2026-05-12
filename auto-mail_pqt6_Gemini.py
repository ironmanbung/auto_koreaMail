import sys
import os
import json
import time
import threading
import ctypes
import shutil
import openpyxl
import pyautogui
import pyperclip
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
    QLabel, QLineEdit, QPushButton, QTableWidget, QTableWidgetItem, 
    QFileDialog, QMessageBox, QHeaderView, QGroupBox, QSizePolicy, QFrame, QTextEdit
)
from PyQt6.QtCore import Qt, QTimer, QUrl, pyqtSlot, pyqtSignal
from PyQt6.QtGui import QFont, QPixmap, QDesktopServices, QColor

# ── 리소스 경로 설정 (PyInstaller 대응) ────────────────────────
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

CONFIG_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "mail_auto_config.json")
COLUMNS = ["연번", "받는사람이메일", "메일제목", "메일본문", "첨부파일명"]
pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.05

# ── 고급 스타일시트 ──────────────────────────────────────────
STYLE_SHEET = """
QMainWindow { background-color: #1a1b26; }

QGroupBox { 
    color: ghostwhite; 
    font-family: 'Malgun Gothic'; 
    font-weight: bold; 
    font-size: 15px;
    border: 1px solid #292e42; 
    border-radius: 10px; 
    margin-top: 12px;
    padding-top: 10px;
    background-color: #24283b;
}
QGroupBox::title { subcontrol-origin: margin; left: 15px; padding: 0 5px; }

QLabel { color: #c0caf5; font-size: 13px; }

QLineEdit { 
    background-color: #1a1b26; 
    color: #c0caf5; 
    border: 1px solid #414868; 
    border-radius: 5px; 
    padding: 5px;
}
QLineEdit:focus { border: 1px solid #7aa2f7; }

QPushButton { 
    background-color: #414868; 
    color: #c0caf5; 
    border-radius: 6px; 
    padding: 6px 12px; 
    font-size: 13px;
    font-weight: bold;
    border: 1px solid #292e42;
}
QPushButton:hover { background-color: #565f89; }
QPushButton:pressed { background-color: #7aa2f7; color: #1a1b26; }

QPushButton#Primary { background-color: #7aa2f7; color: #1a1b26; border: none; }
QPushButton#Primary:hover { background-color: #89ddff; }

QPushButton#Success { background-color: #9ece6a; color: #1a1b26; border: none; }
QPushButton#Success:hover { background-color: #b9f27c; }

QPushButton#Danger { background-color: #f7768e; color: #1a1b26; border: none; }
QPushButton#Danger:hover { background-color: #ff9e64; }

QPushButton#CaptureActive { background-color: #e0af68; color: #1a1b26; border: none; }

QTableWidget { 
    background-color: #1a1b26; 
    color: #c0caf5; 
    border: 1px solid #292e42;
    gridline-color: #292e42;
    border-radius: 5px;
    selection-background-color: #33467c;
}
QHeaderView::section {
    background-color: #24283b;
    color: #7aa2f7;
    padding: 5px;
    border: 1px solid #1a1b26;
    font-weight: bold;
}

QScrollBar:vertical {
    border: none; background: #1a1b26; width: 10px; margin: 0;
}
QScrollBar::handle:vertical {
    background: #414868; min-height: 20px; border-radius: 5px;
}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0px; }

QTextEdit {
    background-color: #16161e;
    color: #9ece6a;
    border: 1px solid #292e42;
    border-radius: 8px;
    font-family: 'Consolas', 'Malgun Gothic';
    font-size: 11px;
    padding: 5px;
}

/* 팝업창 스타일 */
QMessageBox { background-color: #f0f0f0; }
QMessageBox QLabel { color: #000000; }
QMessageBox QPushButton { color: #000000; background-color: #e0e0e0; min-width: 80px; }
"""

class MailAutomationApp(QMainWindow):
    error_signal = pyqtSignal(str)

    def __init__(self):
        super().__init__()
        self.setWindowTitle("공직자통합메일 자동발송 프로그램")
        self.stop_flag = False
        self.capturing_idx = None
        self.load_config()
        self.init_ui()
        self.apply_config()
        self.error_signal.connect(self.show_error_message)

    def init_ui(self):
        self.setStyleSheet(STYLE_SHEET)
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        main_layout = QVBoxLayout(central_widget)
        main_layout.setSpacing(10) # 그룹 간 상하 간격
        main_layout.setContentsMargins(15, 0, 15, 15)

        # ── 상단 타이틀바 (고급 그라데이션) ──────────────────────────
        self.top_bar = QFrame()
        self.top_bar.setFixedHeight(60)
        self.top_bar.setStyleSheet("""
            QFrame {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1, 
                            stop:0 #1a1b26, stop:1 #24283b);
                border-bottom: 2px solid #7aa2f7;
            }
        """)
        top_layout = QHBoxLayout(self.top_bar)
        top_layout.setContentsMargins(20, 0, 20, 0)

        title_label = QLabel("공직자통합메일 일괄발송 프로그램")
        title_label.setStyleSheet("color: aqua; font-size: 22px; font-weight: bold; background: transparent; border: none;")
        sub_title_label = QLabel("Ver 0.2 | 개발: 'AI요리사'팀_'26년 학습조직_리더 대전청 양모세주무관")
        sub_title_label.setStyleSheet("color: floralwhite; font-size: 13px; background: transparent; border: none;")

        top_layout.addWidget(title_label)
        top_layout.addSpacing(15)
        top_layout.addWidget(sub_title_label)
        top_layout.addStretch()
        
        # 메인 레이아웃의 마진 수정을 위해 최상단에 배치
        main_layout.parentWidget().layout().setContentsMargins(0,0,0,0) if main_layout.parentWidget() else None
        main_layout.insertWidget(0, self.top_bar)

        content_layout = QVBoxLayout()
        content_layout.setSpacing(8) # 위젯 내부 요소 간격
        main_layout.addLayout(content_layout)

        # 1. 엑셀 관리
        self.group1 = QGroupBox(" ① 엑셀 파일 관리")
        layout1 = QVBoxLayout(self.group1)
        layout1.setContentsMargins(15, 10, 15, 15)
        
        self.download_label = QLabel("<u>📥 엑셀 표준서식 다운로드 및 바로 실행</u>")
        self.download_label.setStyleSheet("color: #7aa2f7; font-size: 12px; cursor: pointer;")
        self.download_label.mousePressEvent = self.download_template
        
        h1 = QHBoxLayout()
        h1.setSpacing(8)
        self.excel_path_input = QLineEdit()
        self.excel_path_input.setPlaceholderText("엑셀 파일을 선택하세요...")
        btn_load = QPushButton("📂 파일 찾기")
        btn_load.clicked.connect(self.load_excel)
        h1.addWidget(self.excel_path_input); h1.addWidget(btn_load)
        
        layout1.addWidget(self.download_label, alignment=Qt.AlignmentFlag.AlignRight)
        layout1.addLayout(h1)
        content_layout.addWidget(self.group1)

        # 2. 데이터 테이블
        self.group2 = QGroupBox(" ② 데이터 테이블")
        layout2 = QVBoxLayout(self.group2)
        layout2.setContentsMargins(12, 10, 12, 12)
        layout2.setSpacing(10)
        
        self.table = QTableWidget(0, 5)
        self.table.setHorizontalHeaderLabels(COLUMNS)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.table.setFixedHeight(220)
        self.table.itemSelectionChanged.connect(self.handle_selection_change)
        
        h2 = QHBoxLayout()
        h2.setSpacing(10)
        btn_add = QPushButton("➕ 행 추가")
        btn_del = QPushButton("➖ 행 삭제")
        btn_del.clicked.connect(lambda: self.table.removeRow(self.table.currentRow()))
        btn_add.clicked.connect(self.add_row)
        h2.addWidget(btn_add); h2.addWidget(btn_del); h2.addStretch()
        
        layout2.addWidget(self.table); layout2.addLayout(h2)
        content_layout.addWidget(self.group2)

        # 3. 첨부파일 폴더
        self.group3 = QGroupBox(" ③ 첨부파일 폴더")
        layout3 = QHBoxLayout(self.group3)
        layout3.setContentsMargins(15, 12, 15, 15)
        self.folder_input = QLineEdit()
        self.folder_input.setPlaceholderText("첨부파일이 보관된 폴더를 선택하세요...")
        btn_folder = QPushButton("📁 폴더 선택")
        btn_folder.clicked.connect(self.select_folder)
        layout3.addWidget(self.folder_input); layout3.addWidget(btn_folder)
        content_layout.addWidget(self.group3)

        # 4. 위치정보 설정
        self.group4 = QGroupBox(" ④ 위치정보 설정")
        layout4 = QHBoxLayout(self.group4)
        layout4.setContentsMargins(15, 10, 15, 15)
        l_pos, r_pos = QVBoxLayout(), QVBoxLayout()
        l_pos.setSpacing(5); r_pos.setSpacing(5)
        self.pos_btns, self.pos_inputs, self.memo_inputs = [], [], []
        
        for i in range(10):
            row = QHBoxLayout()
            btn = QPushButton(f"위치{i+1}"); btn.setFixedWidth(65)
            btn.clicked.connect(lambda chk, idx=i: self.start_capture(idx))
            pi = QLineEdit(); pi.setFixedWidth(80)
            mi = QLineEdit(); mi.setPlaceholderText(f"클릭위치 메모")
            row.addWidget(btn); row.addWidget(pi); row.addWidget(mi)
            if i < 5: l_pos.addLayout(row)
            else: r_pos.addLayout(row)
            self.pos_btns.append(btn); self.pos_inputs.append(pi); self.memo_inputs.append(mi)
        
        layout4.addLayout(l_pos); layout4.addLayout(r_pos)
        content_layout.addWidget(self.group4)

        # 5. 제어 및 로그
        self.group5 = QGroupBox(" ⑤ 시스템 제어 및 실시간 로그")
        layout5 = QVBoxLayout(self.group5)
        layout5.setContentsMargins(15, 15, 15, 15)
        layout5.setSpacing(12)
        
        h5 = QHBoxLayout()
        h5.setSpacing(15)
        btn_one = QPushButton("📧 선택 행 1건 발송"); btn_one.setObjectName("Success"); btn_one.setFixedHeight(40)
        btn_one.clicked.connect(self.send_selected)
        btn_all = QPushButton("📨 전체 자동 발송 시작"); btn_all.setObjectName("Primary"); btn_all.setFixedHeight(40)
        btn_all.clicked.connect(self.send_all)
        btn_stop = QPushButton("⛔ 긴급 중단"); btn_stop.setObjectName("Danger"); btn_stop.setFixedHeight(40)
        btn_stop.clicked.connect(self.stop_sending)
        h5.addWidget(btn_one); h5.addWidget(btn_all); h5.addWidget(btn_stop)
        
        self.log_display = QTextEdit()
        self.log_display.setReadOnly(True)
        self.log_display.setMinimumHeight(100)
        
        layout5.addLayout(h5); layout5.addWidget(self.log_display)
        content_layout.addWidget(self.group5)

        # 초기 윈도우 배치 (좌측 절반)
        screen = QApplication.primaryScreen().availableGeometry()
        self.setGeometry(0, 0, screen.width() // 2, screen.height())

    @pyqtSlot(str)
    def show_error_message(self, message):
        QMessageBox.critical(self, "중단 알림", message)
        self.add_log("🛑 작업이 일시 중단되었습니다.")

    def add_log(self, text):
        self.log_display.append(f"<span style='color:#565f89;'>[{time.strftime('%H:%M:%S')}]</span> {text}")
        self.log_display.moveCursor(self.log_display.textCursor().MoveOperation.End)

    def automate_mail_logic(self, email, subject, body, filename, folder):
        if str(filename).strip().lower() != "모두":
            full_path = os.path.join(folder, str(filename).strip())
            if not os.path.exists(full_path):
                raise FileNotFoundError(f"'{filename}' 파일을 지정된 폴더에서 찾을 수 없습니다.")

        def click_idx(idx):
            val = self.pos_inputs[idx].text()
            if not val or ',' not in val: return
            x, y = map(int, val.split(','))
            pyautogui.click(x, y)

        def paste(txt):
            pyperclip.copy(str(txt))
            time.sleep(0.2)
            pyautogui.hotkey("ctrl", "v")

        # 자동화 시퀀스
        click_idx(0); time.sleep(1.2); pyautogui.press('enter')
        click_idx(1); paste(email); time.sleep(0.5); pyautogui.press('enter')
        click_idx(2); paste(subject); time.sleep(0.5)
        click_idx(3); paste(body); time.sleep(0.5)
        click_idx(4); time.sleep(0.5); pyautogui.press("end"); time.sleep(0.5)
        
        click_idx(5); time.sleep(1.5)
        pyautogui.hotkey('alt', 'd'); time.sleep(0.3)
        paste(folder); pyautogui.press('enter'); time.sleep(0.5)

        if str(filename).strip() == "모두":
            pyautogui.press('tab', presses=4, interval=0.1)
            time.sleep(0.3); pyautogui.hotkey('ctrl', 'a'); time.sleep(0.5)
        else:
            pyautogui.hotkey('alt', 'n'); time.sleep(0.3)
            paste(filename); pyautogui.press('enter')
        
        pyautogui.hotkey('alt', 'o'); time.sleep(2.0)
        click_idx(8); time.sleep(10) # 발송 대기
        click_idx(9); time.sleep(1)

    def run_automation_thread(self, indices):
        folder = self.folder_input.text()
        for idx in indices:
            if self.stop_flag: break
            
            for c in range(5):
                item = self.table.item(idx, c)
                if item: item.setBackground(QColor("#e0af68")) # 작업중 주황색

            data = [self.table.item(idx, c).text() if self.table.item(idx, c) else "" for c in range(5)]
            self.add_log(f"진행 중: {data[1]}")
            
            try:
                self.automate_mail_logic(data[1], data[2], data[3], data[4], folder)
                for c in range(5):
                    item = self.table.item(idx, c)
                    if item: item.setBackground(QColor("#2e3c2e")) # 성공 초록색
                self.add_log(f"✅ 발송 성공: {data[1]}")
            except Exception as e:
                self.error_signal.emit(str(e))
                return 
                
            time.sleep(2.0)
        self.add_log("🏁 자동화 작업이 완료되었습니다.")

    def handle_selection_change(self):
        for r in range(self.table.rowCount()):
            color = QColor("#33467c") if r == self.table.currentRow() else QColor("#1a1b26")
            for c in range(self.table.columnCount()):
                item = self.table.item(r, c)
                if item: item.setBackground(color)

    def download_template(self, event):
        source = resource_path("add_files/서식_엑셀표준서식.xlsx")
        save_path, _ = QFileDialog.getSaveFileName(self, "표준서식 저장", "서식_엑셀표준서식.xlsx", "Excel Files (*.xlsx)")
        if save_path:
            try:
                shutil.copy2(source, save_path)
                os.startfile(save_path)
                self.add_log("📥 표준 서식 다운로드 및 실행")
            except Exception as e:
                QMessageBox.critical(self, "오류", f"파일 실행 중 문제가 발생했습니다: {e}")

    def load_excel(self):
        path, _ = QFileDialog.getOpenFileName(self, "엑셀 선택", "", "Excel Files (*.xlsx)")
        if path:
            self.excel_path_input.setText(path)
            try:
                wb = openpyxl.load_workbook(path, data_only=True); ws = wb.active
                self.table.setRowCount(0)
                for row in ws.iter_rows(min_row=2, values_only=True):
                    r = self.table.rowCount(); self.table.insertRow(r)
                    for c in range(5):
                        val = row[c] if c < len(row) else ""
                        self.table.setItem(r, c, QTableWidgetItem(str(val) if val is not None else ""))
                self.add_log(f"📂 엑셀 데이터 로드 완료 ({self.table.rowCount()}건)")
            except Exception as e: self.add_log(f"❌ 엑셀 로드 실패: {e}")

    def send_selected(self):
        curr = self.table.currentRow()
        if curr >= 0:
            self.stop_flag = False
            threading.Thread(target=self.run_automation_thread, args=([curr],), daemon=True).start()

    def send_all(self):
        indices = list(range(self.table.rowCount()))
        if indices:
            self.stop_flag = False
            threading.Thread(target=self.run_automation_thread, args=(indices,), daemon=True).start()

    def load_config(self):
        if os.path.exists(CONFIG_PATH):
            try:
                with open(CONFIG_PATH, "r", encoding="utf-8") as f: self.cfg = json.load(f)
            except: self.cfg = {}
        else: self.cfg = {}

    def save_config(self):
        self.cfg.update({
            "excel_path": self.excel_path_input.text(),
            "attachment_folder": self.folder_input.text(),
            "positions": [i.text() for i in self.pos_inputs],
            "memos": [i.text() for i in self.memo_inputs]
        })
        try:
            with open(CONFIG_PATH, "w", encoding="utf-8") as f: json.dump(self.cfg, f, ensure_ascii=False, indent=2)
        except: pass

    def apply_config(self):
        self.excel_path_input.setText(self.cfg.get("excel_path", ""))
        self.folder_input.setText(self.cfg.get("attachment_folder", ""))
        pos, memo = self.cfg.get("positions", [""]*10), self.cfg.get("memos", [""]*10)
        for i in range(10):
            if i < len(pos): self.pos_inputs[i].setText(pos[i])
            if i < len(memo): self.memo_inputs[i].setText(memo[i])

    def add_row(self):
        r = self.table.rowCount(); self.table.insertRow(r)
        self.table.setItem(r, 0, QTableWidgetItem(str(r + 1)))

    def select_folder(self):
        d = QFileDialog.getExistingDirectory(self, "폴더 선택")
        if d: self.folder_input.setText(d)

    def start_capture(self, idx):
        if self.capturing_idx is not None: return
        self.capturing_idx = idx
        self.pos_btns[idx].setObjectName("CaptureActive")
        self.pos_btns[idx].setStyle(self.pos_btns[idx].style())
        QTimer.singleShot(1000, lambda: self.poll_mouse(idx))

    def poll_mouse(self, idx):
        if bool(ctypes.windll.user32.GetAsyncKeyState(0x01) & 0x8000):
            x, y = pyautogui.position()
            self.pos_inputs[idx].setText(f"{x},{y}")
            self.pos_btns[idx].setObjectName("")
            self.pos_btns[idx].setStyle(self.pos_btns[idx].style())
            self.capturing_idx = None
            self.save_config()
        else:
            QTimer.singleShot(50, lambda: self.poll_mouse(idx))

    def stop_sending(self): 
        self.stop_flag = True
        self.add_log("🛑 사용자에 의해 중단 요청됨")

    def closeEvent(self, event):
        self.save_config()
        event.accept()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setFont(QFont("Malgun Gothic", 9))
    window = MailAutomationApp()
    window.show()
    sys.exit(app.exec())