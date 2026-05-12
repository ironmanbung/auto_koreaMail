import sys
import os
import json
import time
import threading
import ctypes
import openpyxl
import pyautogui
import pyperclip
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QTableWidget, QTableWidgetItem,
    QFileDialog, QMessageBox, QHeaderView, QGroupBox, QSizePolicy,
    QFrame, QTextEdit, QScrollArea, QGridLayout, QInputDialog
)
from PyQt6.QtCore import Qt, QTimer, pyqtSlot, pyqtSignal, QMimeData
from PyQt6.QtGui import QFont, QColor, QDrag

def get_config_path():
    if getattr(sys, 'frozen', False):
        base = os.path.dirname(sys.executable)
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, "mail_auto_config.json")

CONFIG_PATH = get_config_path()

pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.05

# ══════════════════════════════════════════════════════════════
#  스타일시트
# ══════════════════════════════════════════════════════════════
STYLE_SHEET = """
QMainWindow { background-color: black; }

QWidget#ScrollContent { background-color: black; }

QGroupBox {
    color: aliceblue;
    font-family: 'Malgun Gothic';
    font-weight: bold;
    font-size: 17px;
    margin-top: 12px;
    padding-top: 10px;
    background-color: black;
}
QGroupBox::title { subcontrol-origin: margin; left: 15px; padding: 0 5px; }

QLabel { color: #c0caf5; font-size: 13px; }

QLineEdit {
    background-color: #1a1b26;
    color: #c0caf5;
    border: 1px solid #414868;
    border-radius: 5px;
    padding: 4px;
}
QLineEdit:focus { border: 1px solid #4169e1; }

/* ── 기본 버튼: royalblue 계열 ── */
QPushButton {
    background-color: #4169e1;
    color: #ffffff;
    border-radius: 6px;
    padding: 5px 10px;
    font-size: 12px;
    font-weight: bold;
    border: none;
}
QPushButton:hover {
    background-color: #5a7ff0;
    color: #ffffff;
}
QPushButton:pressed {
    background-color: #2851c8;
    color: #ffffff;
}

/* ── Primary (실행): deepskyblue 계열 ── */
QPushButton#Primary {
    background-color: #00bfff;
    color: #0a1628;
    border: none;
}
QPushButton#Primary:hover {
    background-color: #33ccff;
    color: #000000;
}
QPushButton#Primary:pressed {
    background-color: #0099cc;
    color: #ffffff;
}

/* ── Success (성공/실행): turquoise 계열 ── */
QPushButton#Success {
    background-color: #40e0d0;
    color: #0a1628;
    border: none;
}
QPushButton#Success:hover {
    background-color: #5eeedd;
    color: #000000;
}
QPushButton#Success:pressed {
    background-color: #2dc4b5;
    color: #ffffff;
}

/* ── Danger (위험/중단/초기화): 붉은 계열 ── */
QPushButton#Danger {
    background-color: #7f1d1d;
    color: #fee2e2;
    border: none;
}
QPushButton#Danger:hover {
    background-color: #ef4444;
    color: #ffffff;
}
QPushButton#Danger:pressed {
    background-color: #fca5a5;
    color: #450a0a;
}

/* ── 위치 버튼 (PosBtn): paleturquoise 계열 ── */
QPushButton#PosBtn {
    background-color: #2e8b8b;
    color: #afeeee;
    border: none;
    border-radius: 5px;
    padding: 1px 3px;
    font-size: 11px;
    font-weight: bold;
}
QPushButton#PosBtn:hover {
    background-color: #afeeee;
    color: #0a2020;
}
QPushButton#PosBtn:pressed {
    background-color: #7fd4d4;
    color: #000000;
}

/* ── CaptureActive: 캡처 중 ── */
QPushButton#CaptureActive {
    background-color: #fbbf24;
    color: #1c1917;
    border: 2px solid #ffffff;
    font-weight: bold;
}

/* ── 키보드 버튼 (KbdBtn): PosBtn과 동일한 paleturquoise 계열 ── */
QPushButton#KbdBtn {
    background-color: #2e8b8b;
    color: #afeeee;
    border: none;
    border-radius: 5px;
    padding: 2px 6px;
    font-size: 11px;
    font-weight: bold;
}
QPushButton#KbdBtn:hover {
    background-color: #afeeee;
    color: #0a2020;
}
QPushButton#KbdBtn:pressed {
    background-color: #7fd4d4;
    color: #000000;
}

/* ── 변수명 버튼 (VarBtn): PosBtn과 동일한 paleturquoise 계열 ── */
QPushButton#VarBtn {
    background-color: #2e8b8b;
    color: #afeeee;
    border: none;
    border-radius: 5px;
    padding: 2px 6px;
    font-size: 11px;
    font-weight: bold;
}
QPushButton#VarBtn:hover {
    background-color: #afeeee;
    color: #0a2020;
}
QPushButton#VarBtn:pressed {
    background-color: #7fd4d4;
    color: #000000;
}

/* ── DIY단독실행 버튼 (DiyOnly): 주황/골드 계열 ── */
QPushButton#DiyOnly {
    background-color: #b45309;
    color: #fef3c7;
    border: none;
    border-radius: 6px;
    padding: 5px 10px;
    font-size: 12px;
    font-weight: bold;
}
QPushButton#DiyOnly:hover {
    background-color: #f59e0b;
    color: #1c1917;
}
QPushButton#DiyOnly:pressed {
    background-color: #78350f;
    color: #fef3c7;
}

QTableWidget {
    background-color: #1a1b26;
    color: #c0caf5;
    border: 1px solid #292e42;
    gridline-color: #292e42;
    border-radius: 5px;
    selection-background-color: #33467c;
}
QTableWidget QTableCornerButton::section {
    background-color: #24283b;
    border: 1px solid #1a1b26;
}
QTableWidget > QAbstractScrollArea {
    background-color: #1a1b26;
}
QHeaderView::section {
    background-color: #24283b;
    color: #7aa2f7;
    padding: 5px;
    border: 1px solid #1a1b26;
    font-weight: bold;
}
QHeaderView {
    background-color: #24283b;
}

QScrollBar:vertical {
    border: none; background: #1a1b26; width: 10px; margin: 0;
}
QScrollBar::handle:vertical {
    background: #414868; min-height: 20px; border-radius: 5px;
}
QScrollBar::handle:vertical:hover { background: #4169e1; }
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0px; }

QScrollBar:horizontal {
    border: none; background: #1a1b26; height: 10px; margin: 0;
}
QScrollBar::handle:horizontal {
    background: #414868; min-width: 20px; border-radius: 5px;
}
QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal { width: 0px; }

QTextEdit {
    background-color: #16161e;
    color: #9ece6a;
    border: 1px solid #292e42;
    border-radius: 8px;
    font-family: 'Consolas', 'Malgun Gothic';
    font-size: 11px;
    padding: 5px;
}

QMessageBox { background-color: #f0f0f0; }
QMessageBox QLabel { color: #000000; }
QMessageBox QPushButton { color: #000000; background-color: #e0e0e0; min-width: 80px; border: 1px solid #aaaaaa; }

QScrollArea { border: none; background: transparent; }
"""


# ══════════════════════════════════════════════════════════════
#  드래그 가능한 소스 버튼
# ══════════════════════════════════════════════════════════════
class DraggableButton(QPushButton):
    """드래그앤드롭으로 DIY 슬롯에 복사할 수 있는 버튼"""
    def __init__(self, label, action_key, parent=None):
        super().__init__(label, parent)
        self.action_key = action_key
        self._drag_start = None
        self.setAcceptDrops(False)

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self._drag_start = event.position().toPoint()
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if not (event.buttons() & Qt.MouseButton.LeftButton):
            return
        if self._drag_start is None:
            return
        if (event.position().toPoint() - self._drag_start).manhattanLength() < 8:
            return
        drag = QDrag(self)
        mime = QMimeData()
        mime.setText(f"{self.action_key}|{self.text()}")
        drag.setMimeData(mime)
        drag.exec(Qt.DropAction.CopyAction)


# ══════════════════════════════════════════════════════════════
#  DIY 슬롯 위젯
#  - 드래그 드롭으로 채우기
#  - 더블클릭으로 비우기
#  - 다른 슬롯으로 드래그하여 내보내기(빼기) 가능
# ══════════════════════════════════════════════════════════════
class DIYSlot(QLabel):
    def __init__(self, index, on_filled_changed=None, parent=None):
        super().__init__(parent)
        self.index = index
        self.action_key = None
        self.action_label = None
        self.on_filled_changed = on_filled_changed
        self.setAcceptDrops(True)
        self.setFixedSize(80, 44)
        self.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.setWordWrap(True)
        self._drag_start = None
        self._refresh_style()

    def is_filled(self):
        return self.action_key is not None

    def _refresh_style(self):
        if self.action_key:
            self.setStyleSheet("""
                QLabel {
                    background-color: darkblue;
                    color: yellow;
                    border: 2px solid yellow;
                    border-radius: 8px;
                    font-size: 12px;
                    font-weight: bold;
                    padding: 1px;
                }
            """)
            self.setText(self.action_label or "")
        else:
            self.setStyleSheet("""
                QLabel {
                    background-color: #1a1b26;
                    color: #414868;
                    border: 2px dashed #292e42;
                    border-radius: 8px;
                    font-size: 11px;
                }
            """)
            self.setText(f"({self.index + 1})")

    # ── 드래그 소스 (채워진 슬롯을 드래그해서 빼기) ───────────
    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton and self.action_key:
            self._drag_start = event.position().toPoint()
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if not (event.buttons() & Qt.MouseButton.LeftButton):
            return
        if self._drag_start is None or not self.action_key:
            return
        if (event.position().toPoint() - self._drag_start).manhattanLength() < 8:
            return
        # 드래그 시작 → 자신 비우기
        drag = QDrag(self)
        mime = QMimeData()
        mime.setText(f"{self.action_key}|{self.action_label}")
        drag.setMimeData(mime)
        self.action_key = None
        self.action_label = None
        self._refresh_style()
        if self.on_filled_changed:
            self.on_filled_changed()
        drag.exec(Qt.DropAction.MoveAction)

    def dragEnterEvent(self, event):
        if event.mimeData().hasText():
            event.acceptProposedAction()
            self.setStyleSheet("""
                QLabel {
                    background-color: #1e3a1e;
                    border: 2px solid #40e0d0;
                    border-radius: 8px;
                    color: #6ee7b7;
                    font-size: 11px;
                }
            """)

    def dragLeaveEvent(self, event):
        self._refresh_style()

    def dropEvent(self, event):
        text = event.mimeData().text()
        if "|" in text:
            key, label = text.split("|", 1)
            self.action_key = key
            self.action_label = label
        self._refresh_style()
        event.acceptProposedAction()
        if self.on_filled_changed:
            self.on_filled_changed()

    def mouseDoubleClickEvent(self, event):
        """더블클릭으로 슬롯 비우기"""
        self.action_key = None
        self.action_label = None
        self._refresh_style()
        if self.on_filled_changed:
            self.on_filled_changed()


# ══════════════════════════════════════════════════════════════
#  DIY 다이어그램 위젯  (10열 × 3행)
# ══════════════════════════════════════════════════════════════
class DIYDiagramWidget(QWidget):
    SLOT_COUNT = 30   # 10 × 3

    def __init__(self, on_changed=None, parent=None):
        super().__init__(parent)
        self.slots = []
        self.arrows = []
        self.on_changed = on_changed   # 슬롯 변경 시 콜백 (저장용)
        self._build_ui()

    def _build_ui(self):
        outer = QVBoxLayout(self)
        outer.setContentsMargins(4, 4, 4, 4)
        outer.setSpacing(3)

        for row in range(3):
            row_widget = QWidget()
            row_widget.setStyleSheet("background: transparent;")
            h = QHBoxLayout(row_widget)
            h.setContentsMargins(2, 2, 2, 2)
            h.setSpacing(0)

            for col in range(10):
                slot_idx = row * 10 + col
                slot = DIYSlot(slot_idx, on_filled_changed=self._slot_changed)
                self.slots.append(slot)
                h.addWidget(slot)

                if col < 9:
                    arrow = QLabel("→")
                    arrow.setFixedWidth(14)
                    arrow.setAlignment(Qt.AlignmentFlag.AlignCenter)
                    arrow.setStyleSheet("color: #414868; font-size: 12px; font-weight: bold;")
                    h.addWidget(arrow)
                    self.arrows.append((row, col, arrow))

            outer.addWidget(row_widget)

    def _slot_changed(self):
        """슬롯 변경 시 화살표 업데이트 + 외부 콜백 호출"""
        self._update_arrows()
        if self.on_changed:
            self.on_changed()

    def _update_arrows(self):
        for row, col, arrow in self.arrows:
            next_slot = self.slots[row * 10 + col + 1]
            if next_slot.is_filled():
                arrow.setStyleSheet("color: #ef4444; font-size: 13px; font-weight: bold;")
            else:
                arrow.setStyleSheet("color: #414868; font-size: 12px; font-weight: bold;")

    def get_actions(self):
        return [(s.action_key, s.action_label) for s in self.slots if s.action_key]

    def get_slot_data(self):
        """슬롯 전체 상태를 직렬화 가능한 리스트로 반환"""
        return [
            {"key": s.action_key, "label": s.action_label}
            for s in self.slots
        ]

    def set_slot_data(self, data):
        """저장된 슬롯 데이터를 복원"""
        for i, slot in enumerate(self.slots):
            if i < len(data):
                slot.action_key   = data[i].get("key")
                slot.action_label = data[i].get("label")
            else:
                slot.action_key   = None
                slot.action_label = None
            slot._refresh_style()
        self._update_arrows()

    def reset(self):
        for s in self.slots:
            s.action_key = None
            s.action_label = None
            s._refresh_style()
        self._update_arrows()
        if self.on_changed:
            self.on_changed()


# ══════════════════════════════════════════════════════════════
#  메인 앱
# ══════════════════════════════════════════════════════════════
class MailAutomationApp(QMainWindow):
    error_signal = pyqtSignal(str)

    def __init__(self):
        super().__init__()
        self.setWindowTitle("(RPA) 업무자동화 프로그램_v0.4")
        self.stop_flag = False
        self.capturing_idx = None
        self.excel_columns = []
        self.var_btns = []
        self.var_btn_widget = None
        self.load_config()
        self.init_ui()
        self.apply_config()
        self.error_signal.connect(self.show_error_message)

    # ──────────────────────────────────────────────────────────
    def init_ui(self):
        self.setStyleSheet(STYLE_SHEET)

        # ★ [수정1] 최소 창 크기를 명시적으로 작게 설정 → 사용자가 창을 자유롭게 줄일 수 있음
        self.setMinimumSize(400, 300)

        central = QWidget()
        self.setCentralWidget(central)
        outer_layout = QVBoxLayout(central)
        outer_layout.setContentsMargins(0, 0, 0, 0)
        outer_layout.setSpacing(0)

        # ── 상단 타이틀바 ─────────────────────────────────────
        top_bar = QFrame()
        top_bar.setFixedHeight(58)
        top_bar.setStyleSheet("""
            QFrame {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
                            stop:0 #1a1b26, stop:1 #24283b);
                border-bottom: 2px solid #4169e1;
            }
        """)
        tb_layout = QHBoxLayout(top_bar)
        tb_layout.setContentsMargins(20, 0, 20, 0)
        title_lbl = QLabel("(RPA) 업무자동화(엑셀자료와 좌표 기반) 프로그램")
        title_lbl.setStyleSheet(
            "color: deepskyblue; font-size: 21px; font-weight: bold;"
            " background: transparent; border: none;"
        )
        sub_lbl = QLabel("Ver 0.4 | 개발: 'AI요리사'팀_'26년 학습조직_리더 대전청 양모세주무관")
        sub_lbl.setStyleSheet(
            "color: floralwhite; font-size: 12px;"
            " background: transparent; border: none;"
        )
        tb_layout.addWidget(title_lbl)
        tb_layout.addSpacing(15)
        tb_layout.addWidget(sub_lbl)
        tb_layout.addStretch()
        outer_layout.addWidget(top_bar)

        # ── 스크롤 영역 ────────────────────────────────────────
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        scroll_area.setStyleSheet("QScrollArea { border: none; background: #1a1b26; }")
        outer_layout.addWidget(scroll_area)

        # ── 하단 푸터바 ────────────────────────────────────────
        footer_bar = QFrame()
        footer_bar.setFixedHeight(48)
        footer_bar.setStyleSheet("""
            QFrame {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                            stop:0 #1a1b26, stop:1 #1e2030);
                border-top: 1px solid #292e42;
            }
        """)
        footer_layout = QHBoxLayout(footer_bar)
        footer_layout.setContentsMargins(20, 6, 20, 6)
        footer_layout.addStretch()

        footer_vbox = QVBoxLayout()
        footer_vbox.setSpacing(2)
        footer_vbox.setContentsMargins(0, 0, 0, 0)

        footer_line1 = QLabel(
            "© 2026 Public Domain — 저작권 제한 없이 누구나 자유롭게 사용·복제·배포·수정할 수 있습니다."
        )
        footer_line1.setStyleSheet(
            "color: snow; font-size: 12px; background: transparent; border: none;"
        )
        footer_line1.setAlignment(Qt.AlignmentFlag.AlignRight)

        footer_line2 = QLabel("※ 개발 : AI요리사 (대전지방고용노동청 지역협력과)")
        footer_line2.setStyleSheet(
            "color: white; font-size: 12px; background: transparent; border: none;"
        )
        footer_line2.setAlignment(Qt.AlignmentFlag.AlignRight)

        footer_vbox.addWidget(footer_line1)
        footer_vbox.addWidget(footer_line2)
        footer_layout.addLayout(footer_vbox)
        outer_layout.addWidget(footer_bar)

        scroll_content = QWidget()
        scroll_content.setObjectName("ScrollContent")
        # ★ [수정1] minimumWidth를 줄여 창 축소 시 스크롤바로 대응하도록 변경
        scroll_content.setMinimumWidth(600)
        scroll_area.setWidget(scroll_content)

        content_layout = QVBoxLayout(scroll_content)
        content_layout.setContentsMargins(15, 8, 15, 15)
        content_layout.setSpacing(8)

        # ══════════════════════════════════════════════════════
        #  1. 엑셀 파일 관리 + 데이터 테이블
        # ══════════════════════════════════════════════════════
        g1 = QGroupBox("Ⅰ. 엑셀 파일 관리")
        l1 = QVBoxLayout(g1)
        l1.setContentsMargins(15, 10, 15, 12)
        l1.setSpacing(8)

        title_row = QHBoxLayout()
        title_row.setSpacing(8)

        h1 = QHBoxLayout()
        h1.setSpacing(8)
        self.excel_path_input = QLineEdit()
        self.excel_path_input.setPlaceholderText("엑셀 파일을 선택하세요...")
        btn_load = QPushButton("📂 파일 찾기")
        btn_load.clicked.connect(self.load_excel)

        btn_excel_reset = QPushButton("🔄 초기화")
        btn_excel_reset.setObjectName("Danger")
        btn_excel_reset.setFixedHeight(28)
        btn_excel_reset.clicked.connect(self.reset_excel)

        h1.addWidget(self.excel_path_input)
        h1.addWidget(btn_load)
        h1.addWidget(btn_excel_reset)

        l1.addLayout(h1)

        # ── 데이터 테이블 ─────────────────────────────────────
        self.table = QTableWidget(0, 0)
        hdr = self.table.horizontalHeader()
        hdr.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        hdr.setMinimumSectionSize(30)
        hdr.setDefaultSectionSize(120)
        hdr.setStretchLastSection(True)
        self.table.setFixedHeight(152)
        self.table.itemSelectionChanged.connect(self.handle_selection_change)

        h2 = QHBoxLayout()
        h2.setSpacing(8)
        btn_add = QPushButton("➕ 행 추가")
        btn_del = QPushButton("➖ 행 삭제")
        btn_del.clicked.connect(lambda: self.table.removeRow(self.table.currentRow()))
        btn_add.clicked.connect(self.add_row)
        h2.addWidget(btn_add)
        h2.addWidget(btn_del)
        h2.addStretch()

        l1.addWidget(self.table)
        l1.addLayout(h2)

        content_layout.addWidget(g1)

        # ══════════════════════════════════════════════════════
        #  2. 변수명, 위치 좌표 및 키보드작업 버튼화
        # ══════════════════════════════════════════════════════
        g4 = QGroupBox("Ⅱ. 자동화 작업 '버튼 모음'")
        l4 = QVBoxLayout(g4)
        l4.setContentsMargins(12, 8, 12, 10)
        l4.setSpacing(6)

        # ── Ⅱ 섹션 헤더 행 (제목 레이블 + 초기화 버튼) ──────────
        sec2_header = QHBoxLayout()
        sec2_header.setSpacing(8)
        sec2_header.addStretch()
        btn_pos_reset = QPushButton("🔄 초기화")
        btn_pos_reset.setObjectName("Danger")
        btn_pos_reset.setFixedHeight(28)
        btn_pos_reset.setToolTip("위치 좌표 및 메모 전체 초기화")
        btn_pos_reset.clicked.connect(self.reset_positions)
        sec2_header.addWidget(btn_pos_reset)
        l4.addLayout(sec2_header)

        # ── ① 변수명 버튼화 ──────────────────────────────────
        var_label = QLabel("① 변수명 버튼화")
        var_label.setStyleSheet(
            "color: #bae6fd; font-weight: bold; font-size: 13px; padding: 2px 0;"
        )
        l4.addWidget(var_label)

        self.var_btn_container = QWidget()
        self.var_btn_container.setStyleSheet("background: transparent;")
        self.var_btn_grid = QGridLayout(self.var_btn_container)
        self.var_btn_grid.setHorizontalSpacing(4)
        self.var_btn_grid.setVerticalSpacing(3)
        self.var_btn_grid.setContentsMargins(0, 0, 0, 0)

        no_col_lbl = QLabel("   ※ 엑셀 파일을 로드하면 '변수명별'로 버튼이 생성됩니다.")
        no_col_lbl.setStyleSheet("color: lavender; font-size: 12px;")
        self.var_btn_grid.addWidget(no_col_lbl, 0, 0)
        self.var_placeholder_lbl = no_col_lbl

        l4.addWidget(self.var_btn_container)

        # 구분선
        sep1 = QFrame()
        sep1.setFrameShape(QFrame.Shape.HLine)
        sep1.setStyleSheet("color: #292e42; margin-top: 2px; margin-bottom: 2px;")
        l4.addWidget(sep1)

        # ── ② 위치 좌표 버튼화 ───────────────────────────────
        pos_label = QLabel("② 위치 좌표 버튼화")
        pos_label.setStyleSheet(
            "color: #afeeee; font-weight: bold; font-size: 13px; padding: 2px 0;"
        )
        l4.addWidget(pos_label)

        self.pos_btns, self.pos_inputs, self.memo_inputs = [], [], []

        pos_grid = QGridLayout()
        pos_grid.setHorizontalSpacing(18)
        pos_grid.setVerticalSpacing(2)

        for i in range(15):
            r, c = i // 3, i % 3

            cell_w = QWidget()
            cell_h = QHBoxLayout(cell_w)
            cell_h.setContentsMargins(0, 0, 0, 0)
            cell_h.setSpacing(3)

            btn = DraggableButton(f"위치{i+1}", f"pos_{i}")
            btn.setObjectName("PosBtn")
            btn.setFixedWidth(56)
            btn.setFixedHeight(24)
            btn.setFont(QFont("Malgun Gothic", 9, QFont.Weight.Bold))
            btn.clicked.connect(lambda chk, idx=i: self.start_capture(idx))

            pi = QLineEdit()
            pi.setFixedWidth(72)
            pi.setFixedHeight(24)

            mi = QLineEdit()
            mi.setPlaceholderText("메모")
            mi.setFixedWidth(90)
            mi.setFixedHeight(24)

            cell_h.addWidget(btn)
            cell_h.addWidget(pi)
            cell_h.addWidget(mi)
            pos_grid.addWidget(cell_w, r, c)

            self.pos_btns.append(btn)
            self.pos_inputs.append(pi)
            self.memo_inputs.append(mi)

        l4.addLayout(pos_grid)

        # 구분선
        sep2 = QFrame()
        sep2.setFrameShape(QFrame.Shape.HLine)
        sep2.setStyleSheet("color: #292e42; margin-top: 2px; margin-bottom: 2px;")
        l4.addWidget(sep2)

        # ── ③ 키보드 작업 버튼화 ─────────────────────────────
        kbd_label = QLabel("③ 키보드 작업 버튼화")
        kbd_label.setStyleSheet(
            "color: #bfdbfe; font-weight: bold; font-size: 13px; padding: 2px 0;"
        )
        l4.addWidget(kbd_label)

        KBD_ACTIONS = [
            ("Enter",    "kbd_enter",     "Enter 키"),
            ("Tab",      "kbd_tab",       "Tab 키"),
            ("ESC",      "kbd_esc",       "Escape 키"),
            ("Space",    "kbd_space",     "Space 키"),
            ("Back",     "kbd_backspace", "Backspace 키"),
            ("Del",      "kbd_delete",    "Delete 키"),
            ("↑",        "kbd_up",        "위 방향키"),
            ("↓",        "kbd_down",      "아래 방향키"),
            ("←",        "kbd_left",      "왼쪽 방향키"),
            ("→",        "kbd_right",     "오른쪽 방향키"),
            ("Home",     "kbd_home",      "Home 키"),
            ("End",      "kbd_end",       "End 키"),
            ("PgUp",     "kbd_pageup",    "Page Up"),
            ("PgDn",     "kbd_pagedown",  "Page Down"),
            ("Ctrl+A",   "kbd_ctrl_a",    "전체 선택"),
            ("Ctrl+C",   "kbd_ctrl_c",    "복사"),
            ("Ctrl+V",   "kbd_ctrl_v",    "붙여넣기"),
            ("Ctrl+Z",   "kbd_ctrl_z",    "실행취소"),
            ("Ctrl+S",   "kbd_ctrl_s",    "저장"),
            ("Alt+F4",   "kbd_alt_f4",    "창 닫기"),
            ("Alt+O",    "kbd_alt_o",     "열기(파일대화상자)"),
            ("Alt+D",    "kbd_alt_d",     "주소표시줄 포커스"),
            ("Win+D",    "kbd_win_d",     "바탕화면 보기"),
            ("Alt+Tab",  "kbd_alt_tab",   "창 전환(Alt+Tab)"),
            ("F5",       "kbd_f5",        "새로고침"),
            ("대기0.5s", "wait_0.5",      "0.5초 대기"),
            ("대기1s",   "wait_1",        "1초 대기"),
            ("대기2s",   "wait_2",        "2초 대기"),
            ("대기3s",   "wait_3",        "3초 대기"),
            ("대기5s",   "wait_5",        "5초 대기"),
        ]
        self.kbd_action_map = {a[1]: a for a in KBD_ACTIONS}

        kbd_grid = QGridLayout()
        kbd_grid.setHorizontalSpacing(4)
        kbd_grid.setVerticalSpacing(3)
        COLS_PER_ROW = 14
        self.kbd_btns = []
        for i, (label, key, tip) in enumerate(KBD_ACTIONS):
            btn = DraggableButton(label, key)
            btn.setObjectName("KbdBtn")
            btn.setFixedHeight(24)
            btn.setToolTip(tip)
            kbd_grid.addWidget(btn, i // COLS_PER_ROW, i % COLS_PER_ROW)
            self.kbd_btns.append(btn)

        l4.addLayout(kbd_grid)
        content_layout.addWidget(g4)

        # ══════════════════════════════════════════════════════
        #  3. DIY 작업 조합 (10열 × 3행)
        # ══════════════════════════════════════════════════════
        g5 = QGroupBox("Ⅲ. DIY 작업 조합")
        l5 = QVBoxLayout(g5)
        l5.setContentsMargins(12, 8, 12, 8)
        l5.setSpacing(4)

        diy_top = QHBoxLayout()
        diy_help = QLabel(
            "   ※ 위 버튼들을 드래그해서 아래 슬롯에 놓으세요. | 슬롯 간 이동, 빼기(더블클릭 또는 빈공간 이동 시) 가능합니다. | 전체 슬롯 순서대로 실행됩니다"
        )
        diy_help.setStyleSheet("color: lavender; font-size: 12px;")
        diy_top.addWidget(diy_help)
        diy_top.addStretch()

        btn_diy_reset = QPushButton("🔄 초기화")
        btn_diy_reset.setObjectName("Danger")
        btn_diy_reset.setFixedHeight(28)
        btn_diy_reset.clicked.connect(self.reset_diy)
        diy_top.addWidget(btn_diy_reset)

        l5.addLayout(diy_top)

        self.diy_diagram = DIYDiagramWidget(on_changed=self.save_config)
        l5.addWidget(self.diy_diagram)

        content_layout.addWidget(g5)

        # ══════════════════════════════════════════════════════
        #  4. 시스템 제어 및 실시간 로그
        # ══════════════════════════════════════════════════════
        g6 = QGroupBox("Ⅳ. 시스템 제어 및 실시간 로그")
        l6 = QVBoxLayout(g6)
        l6.setContentsMargins(15, 12, 15, 12)
        l6.setSpacing(8)

        h6 = QHBoxLayout()
        h6.setSpacing(10)

        # ★ [수정2] DIY 작업조합만 실행(엑셀파일 없이 실행) 버튼 추가 — btn_one 왼쪽에 배치
        btn_diy_only = QPushButton("⚡ DIY 작업조합만 실행\n(엑셀파일 없이 실행)")
        btn_diy_only.setObjectName("DiyOnly")
        btn_diy_only.setFixedHeight(38)
        btn_diy_only.clicked.connect(self.run_diy_only)

        btn_one = QPushButton("📧 선택 행 1건 실행")
        btn_one.setObjectName("Primary")
        btn_one.setFixedHeight(38)
        btn_one.clicked.connect(self.send_selected)

        btn_all = QPushButton("📨 전체 자동 실행")
        btn_all.setObjectName("Primary")
        btn_all.setFixedHeight(38)
        btn_all.clicked.connect(self.send_all)

        btn_stop = QPushButton("⛔ 긴급 중단")
        btn_stop.setObjectName("Danger")
        btn_stop.setFixedHeight(38)
        btn_stop.clicked.connect(self.stop_sending)

        btn_clear_log = QPushButton("🗑 로그 지우기")
        btn_clear_log.setObjectName("Danger")
        btn_clear_log.setFixedHeight(38)
        btn_clear_log.clicked.connect(self.clear_log)

        # ★ [수정2] DIY단독 버튼을 btn_one 왼쪽에 먼저 추가
        h6.addWidget(btn_diy_only)
        h6.addWidget(btn_one)
        h6.addWidget(btn_all)
        h6.addWidget(btn_stop)
        h6.addWidget(btn_clear_log)

        self.log_display = QTextEdit()
        self.log_display.setReadOnly(True)
        self.log_display.setMinimumHeight(110)

        l6.addLayout(h6)
        l6.addWidget(self.log_display)
        content_layout.addWidget(g6)

        content_layout.addStretch()

        # ★ [수정1] 화면 절반 크기로 초기 창 설정 (최소 크기는 setMinimumSize로 따로 제한)
        screen = QApplication.primaryScreen().availableGeometry()
        self.setGeometry(0, 0, screen.width() // 2, screen.height())

    # ══════════════════════════════════════════════════════════
    #  엑셀 파일 초기화
    # ══════════════════════════════════════════════════════════
    def reset_excel(self):
        self.excel_path_input.clear()
        self.table.setRowCount(0)
        self.table.setColumnCount(0)
        self.excel_columns = []
        # 변수명 버튼 초기화
        for i in reversed(range(self.var_btn_grid.count())):
            w = self.var_btn_grid.itemAt(i).widget()
            if w:
                w.deleteLater()
        self.var_btns = []
        no_col_lbl = QLabel("   ※ 엑셀 파일을 로드하면 컬럼 버튼이 생성됩니다.")
        no_col_lbl.setStyleSheet("color: lavender; font-size: 12px;")
        self.var_btn_grid.addWidget(no_col_lbl, 0, 0)
        self.var_placeholder_lbl = no_col_lbl
        self.add_log("🔄 엑셀 파일 관리 초기화")

    # ══════════════════════════════════════════════════════════
    #  엑셀 로드
    # ══════════════════════════════════════════════════════════
    def load_excel(self):
        path, _ = QFileDialog.getOpenFileName(self, "엑셀 선택", "", "Excel Files (*.xlsx)")
        if not path:
            return
        self.excel_path_input.setText(path)
        try:
            wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
            ws = wb.active

            all_rows = list(ws.iter_rows(values_only=True))
            wb.close()

            if not all_rows:
                self.add_log("❌ 엑셀 파일이 비어 있습니다.")
                return

            first_row = all_rows[0]
            columns = [str(v) if v is not None else f"Col{i+1}"
                    for i, v in enumerate(first_row)]
            self.excel_columns = columns
            col_count = len(columns)

            self.table.setColumnCount(col_count)
            self.table.setHorizontalHeaderLabels(columns)
            hdr = self.table.horizontalHeader()
            hdr.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
            hdr.setMinimumSectionSize(30)
            hdr.setStretchLastSection(True)
            QTimer.singleShot(0, lambda: self._set_initial_column_widths(col_count))

            self.table.setRowCount(0)
            for row in all_rows[1:]:
                r = self.table.rowCount()
                self.table.insertRow(r)
                for c in range(col_count):
                    val = row[c] if c < len(row) else ""
                    self.table.setItem(r, c, QTableWidgetItem(
                        str(val) if val is not None else ""
                    ))

            self.add_log(f"📂 엑셀 데이터 로드 완료 ({self.table.rowCount()}건, 컬럼: {col_count}개)")
            self._rebuild_var_buttons(columns)

        except Exception as e:
            self.add_log(f"❌ 엑셀 로드 실패: {e}")

    def _set_initial_column_widths(self, col_count):
        """초기 컬럼 너비를 균등 배분 (Interactive 모드용)"""
        if col_count <= 0:
            return
        total_w = self.table.viewport().width()
        if col_count > 1:
            col_w = max(60, total_w // col_count)
            hdr = self.table.horizontalHeader()
            for i in range(col_count - 1):
                hdr.resizeSection(i, col_w)

    def _rebuild_var_buttons(self, columns):
        """변수명 버튼을 동적으로 생성"""
        for i in reversed(range(self.var_btn_grid.count())):
            w = self.var_btn_grid.itemAt(i).widget()
            if w:
                w.deleteLater()

        self.var_btns = []
        COLS_PER_ROW = 12

        VAR_BTN_STYLE = """
            QPushButton {
                background-color: #2e8b8b;
                color: #afeeee;
                border: none;
                border-radius: 5px;
                padding: 2px 6px;
                font-size: 11px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #afeeee;
                color: #0a2020;
            }
            QPushButton:pressed {
                background-color: #7fd4d4;
                color: #000000;
            }
        """

        for i, col_name in enumerate(columns):
            r, c = i // COLS_PER_ROW, i % COLS_PER_ROW
            btn = DraggableButton(col_name, f"var_{col_name}")
            btn.setFixedHeight(24)
            btn.setStyleSheet(VAR_BTN_STYLE)
            btn.setToolTip(f"컬럼: {col_name}  →  해당 레코드의 셀 값 복사")
            self.var_btn_grid.addWidget(btn, r, c)
            self.var_btns.append(btn)

    # ══════════════════════════════════════════════════════════
    #  DIY 실행 (내부 공통)
    # ══════════════════════════════════════════════════════════
    def _diy_thread(self, actions, row_index=None):
        for key, label in actions:
            if self.stop_flag:
                self.add_log("🛑 DIY 작업 중단")
                return
            try:
                self._execute_action(key, label, row_index)
                self.add_log(f"  ✅ [{label}] 완료")
                time.sleep(0.5)
            except Exception as e:
                self.add_log(f"  ❌ [{label}] 오류: {e}")
        self.add_log("🏁 DIY 작업 완료")

    def _execute_action(self, key, label, row_index=None):
        # 변수명 버튼 → 해당 레코드·컬럼 셀값 복사
        if key.startswith("var_"):
            col_name = key[4:]
            if col_name in self.excel_columns:
                col_idx = self.excel_columns.index(col_name)
                if row_index is not None and row_index < self.table.rowCount():
                    item = self.table.item(row_index, col_idx)
                    cell_val = item.text() if item else ""
                else:
                    cell_val = ""
                pyperclip.copy(str(cell_val))
                time.sleep(0.2)
                pyautogui.hotkey("ctrl", "v")
            else:
                self.add_log(f"  ⚠ [{label}] 컬럼 '{col_name}'를 찾을 수 없습니다.")
            return

        if key.startswith("pos_"):
            idx = int(key.split("_")[1])
            val = self.pos_inputs[idx].text()
            label_text = self.memo_inputs[idx].text() or f"위치{idx+1}"
            if val and ',' in val:
                try:
                    x, y = map(lambda v: int(float(v)), val.split(','))
                    pyautogui.click(x, y)
                except ValueError:
                    self.add_log(f"  ⚠ [{label}] 좌표 형식 오류: '{val}'")
                time.sleep(0.3)
            else:
                self.add_log(f"  ⚠ [{label}] 좌표 미설정, 건너뜁니다.")
        elif key == "kbd_enter":      pyautogui.press('enter')
        elif key == "kbd_tab":        pyautogui.press('tab')
        elif key == "kbd_esc":        pyautogui.press('escape')
        elif key == "kbd_space":      pyautogui.press('space')
        elif key == "kbd_backspace":  pyautogui.press('backspace')
        elif key == "kbd_delete":     pyautogui.press('delete')
        elif key == "kbd_up":         pyautogui.press('up')
        elif key == "kbd_down":       pyautogui.press('down')
        elif key == "kbd_left":       pyautogui.press('left')
        elif key == "kbd_right":      pyautogui.press('right')
        elif key == "kbd_home":       pyautogui.press('home')
        elif key == "kbd_end":        pyautogui.press('end')
        elif key == "kbd_pageup":     pyautogui.press('pageup')
        elif key == "kbd_pagedown":   pyautogui.press('pagedown')
        elif key == "kbd_ctrl_a":     pyautogui.hotkey('ctrl', 'a')
        elif key == "kbd_ctrl_c":     pyautogui.hotkey('ctrl', 'c')
        elif key == "kbd_ctrl_v":     pyautogui.hotkey('ctrl', 'v')
        elif key == "kbd_ctrl_z":     pyautogui.hotkey('ctrl', 'z')
        elif key == "kbd_ctrl_s":     pyautogui.hotkey('ctrl', 's')
        elif key == "kbd_alt_f4":     pyautogui.hotkey('alt', 'f4')
        elif key == "kbd_alt_o":      pyautogui.hotkey('alt', 'o')
        elif key == "kbd_alt_d":      pyautogui.hotkey('alt', 'd')
        elif key == "kbd_win_d":      pyautogui.hotkey('win', 'd')
        elif key == "kbd_alt_tab":    pyautogui.hotkey('alt', 'tab')
        elif key == "kbd_f5":         pyautogui.press('f5')
        elif key == "wait_0.5":       time.sleep(0.5)
        elif key == "wait_1":         time.sleep(1.0)
        elif key == "wait_2":         time.sleep(2.0)
        elif key == "wait_3":         time.sleep(3.0)
        elif key == "wait_5":         time.sleep(5.0)

    def reset_diy(self):
        self.diy_diagram.reset()
        self.add_log("🔄 DIY 작업 조합 초기화")

    # ══════════════════════════════════════════════════════════
    #  Ⅱ섹션 위치 좌표 초기화
    # ══════════════════════════════════════════════════════════
    def reset_positions(self):
        for pi in self.pos_inputs:
            pi.clear()
        for mi in self.memo_inputs:
            mi.clear()
        self.save_config()
        self.add_log("🔄 위치 좌표 및 메모 전체 초기화")

    # ══════════════════════════════════════════════════════════
    #  ★ [수정2] DIY 작업조합만 실행 (엑셀 없이) — 반복 횟수 입력 포함
    # ══════════════════════════════════════════════════════════
    def run_diy_only(self):
        actions = self.diy_diagram.get_actions()
        if not actions:
            QMessageBox.information(
                self, "DIY 단독 실행",
                "Ⅲ. DIY 작업 조합에 배치된 슬롯이 없습니다.\n"
                "버튼을 드래그하여 슬롯에 먼저 배치하세요."
            )
            return

        # ── 커스텀 다이얼로그 ──────────────────────────────────
        from PyQt6.QtWidgets import QDialog, QSpinBox, QDialogButtonBox

        dlg = QDialog(self)
        dlg.setWindowTitle("반복 횟수 입력")
        dlg.setFixedWidth(320)
        dlg.setStyleSheet("""
            QDialog { background-color: #ffffff; }
            QLabel  { color: #000000; font-size: 14px; font-weight: bold; }
            QSpinBox {
                color: #000000;
                background-color: #f5f5f5;
                border: 1px solid #aaaaaa;
                border-radius: 4px;
                font-size: 16px;
                min-height: 40px;
                padding: 4px 8px;
            }
            QSpinBox::up-button, QSpinBox::down-button { width: 24px; }
            QDialogButtonBox QPushButton {
                color: #000000;
                background-color: #e0e0e0;
                border: 1px solid #aaaaaa;
                min-width: 70px;
                min-height: 25px;
                border-radius: 4px;
                font-size: 13px;
            }
            QDialogButtonBox QPushButton:hover { background-color: #c8c8c8; }
        """)

        vbox = QVBoxLayout(dlg)
        vbox.setContentsMargins(20, 20, 20, 16)
        vbox.setSpacing(12)

        lbl = QLabel("DIY 작업 조합을 몇 번 반복하시겠습니까?")
        lbl.setWordWrap(True)
        vbox.addWidget(lbl)

        spin = QSpinBox()
        spin.setRange(1, 9999)
        spin.setValue(1)
        vbox.addWidget(spin)

        btn_box = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        btn_box.accepted.connect(dlg.accept)
        btn_box.rejected.connect(dlg.reject)
        vbox.addWidget(btn_box)

        if dlg.exec() != QDialog.DialogCode.Accepted:
            self.add_log("⚠ DIY 단독 실행 취소됨")
            return

        repeat_count = spin.value()
        # ──────────────────────────────────────────────────────

        self.stop_flag = False
        self.add_log(
            f"⚡ DIY 작업조합 단독 실행 시작 "
            f"(슬롯 {len(actions)}단계 × {repeat_count}회 반복)"
        )
        threading.Thread(
            target=self._run_diy_only_thread,
            args=(actions, repeat_count),
            daemon=True
        ).start()

    def _run_diy_only_thread(self, actions, repeat_count):
        """DIY 단독 실행 스레드: row_index=None 으로 _diy_thread 호출"""
        for i in range(repeat_count):
            if self.stop_flag:
                self.add_log("🛑 DIY 단독 실행 중단")
                return
            if repeat_count > 1:
                self.add_log(f"🔁 {i + 1} / {repeat_count} 회차 실행 중...")
            self._diy_thread(actions, row_index=None)
            # 마지막 회차가 아니면 회차 간 짧은 대기
            if i < repeat_count - 1 and not self.stop_flag:
                time.sleep(1.0)
        self.add_log(
            f"🏁 DIY 단독 실행 완료 (총 {repeat_count}회)"
        )

    # ══════════════════════════════════════════════════════════
    #  실행 버튼
    # ══════════════════════════════════════════════════════════
    def send_selected(self):
        curr = self.table.currentRow()
        if curr < 0:
            QMessageBox.warning(self, "행 선택 필요", "테이블에서 실행할 행을 먼저 선택하세요.")
            return
        actions = self.diy_diagram.get_actions()
        if not actions:
            QMessageBox.information(self, "DIY 실행",
                "Ⅲ. DIY 작업 조합에 배치된 슬롯이 없습니다.\n"
                "버튼을 드래그하여 슬롯에 먼저 배치하세요.")
            return
        self.stop_flag = False
        self.add_log(f"📧 선택 행 ({curr+1}번) DIY 실행 시작 ({len(actions)}단계)")
        for c in range(self.table.columnCount()):
            item = self.table.item(curr, c)
            if item: item.setBackground(QColor("#e0af68"))
        threading.Thread(
            target=self._run_single_row_diy,
            args=(actions, curr),
            daemon=True
        ).start()

    def _run_single_row_diy(self, actions, row_index):
        self._diy_thread(actions, row_index)
        for c in range(self.table.columnCount()):
            item = self.table.item(row_index, c)
            if item: item.setBackground(QColor("#2e3c2e"))

    def send_all(self):
        row_count = self.table.rowCount()
        if row_count == 0:
            QMessageBox.warning(self, "데이터 없음", "테이블에 데이터가 없습니다.")
            return
        actions = self.diy_diagram.get_actions()
        if not actions:
            QMessageBox.information(self, "DIY 실행",
                "Ⅲ. DIY 작업 조합에 배치된 슬롯이 없습니다.\n"
                "버튼을 드래그하여 슬롯에 먼저 배치하세요.")
            return
        self.stop_flag = False
        self.add_log(f"📨 전체 일괄 자동실행 시작 (총 {row_count}건, 각 {len(actions)}단계)")
        threading.Thread(
            target=self._run_all_rows_diy,
            args=(actions, list(range(row_count))),
            daemon=True
        ).start()

    def _run_all_rows_diy(self, actions, indices):
        for row_idx in indices:
            if self.stop_flag:
                self.add_log("🛑 전체 실행 중단")
                return
            self.add_log(f"▶ {row_idx+1}번 행 DIY 실행 중...")
            for c in range(self.table.columnCount()):
                item = self.table.item(row_idx, c)
                if item: item.setBackground(QColor("#e0af68"))
            self._diy_thread(actions, row_idx)
            for c in range(self.table.columnCount()):
                item = self.table.item(row_idx, c)
                if item: item.setBackground(QColor("#2e3c2e"))
            time.sleep(2.0)
        self.add_log("🏁 전체 일괄 자동실행 완료")

    # ══════════════════════════════════════════════════════════
    #  자동화 관련 (공통)
    # ══════════════════════════════════════════════════════════
    @pyqtSlot(str)
    def show_error_message(self, message):
        QMessageBox.critical(self, "중단 알림", message)
        self.add_log("🛑 작업이 일시 중단되었습니다.")

    def add_log(self, text):
        self.log_display.append(
            f"<span style='color:#565f89;'>[{time.strftime('%H:%M:%S')}]</span> {text}"
        )
        self.log_display.moveCursor(self.log_display.textCursor().MoveOperation.End)

    def clear_log(self):
        self.log_display.clear()

    def handle_selection_change(self):
        for r in range(self.table.rowCount()):
            color = (QColor("#33467c") if r == self.table.currentRow()
                     else QColor("#1a1b26"))
            for c in range(self.table.columnCount()):
                item = self.table.item(r, c)
                if item: item.setBackground(color)

    # ══════════════════════════════════════════════════════════
    #  설정 저장/불러오기
    # ══════════════════════════════════════════════════════════
    def load_config(self):
        if os.path.exists(CONFIG_PATH):
            try:
                with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                    self.cfg = json.load(f)
            except Exception:
                self.cfg = {}
        else:
            self.cfg = {}

    def save_config(self):
        """현재 UI 상태 전체를 JSON에 저장 (종료 시 + DIY 슬롯 변경 시 자동 호출)"""
        diy_data = self.diy_diagram.get_slot_data() if hasattr(self, 'diy_diagram') else []
        self.cfg.update({
            "excel_path": self.excel_path_input.text(),
            "positions":  [i.text() for i in self.pos_inputs],
            "memos":      [i.text() for i in self.memo_inputs],
            "diy_slots":  diy_data,
        })
        try:
            with open(CONFIG_PATH, "w", encoding="utf-8") as f:
                json.dump(self.cfg, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def apply_config(self):
        """저장된 설정을 UI에 복원"""
        self.excel_path_input.setText(self.cfg.get("excel_path", ""))
        pos  = self.cfg.get("positions", [""] * 15)
        memo = self.cfg.get("memos",     [""] * 15)
        for i in range(15):
            if i < len(pos):  self.pos_inputs[i].setText(pos[i])
            if i < len(memo): self.memo_inputs[i].setText(memo[i])

        diy_slots = self.cfg.get("diy_slots", [])
        if diy_slots:
            self.diy_diagram.set_slot_data(diy_slots)

    # ══════════════════════════════════════════════════════════
    #  기타 유틸
    # ══════════════════════════════════════════════════════════
    def add_row(self):
        r = self.table.rowCount()
        self.table.insertRow(r)
        for c in range(self.table.columnCount()):
            self.table.setItem(r, c, QTableWidgetItem(""))

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
            self.pos_btns[idx].setObjectName("PosBtn")
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


# ══════════════════════════════════════════════════════════════
if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setFont(QFont("Malgun Gothic", 9))
    window = MailAutomationApp()
    window.show()
    sys.exit(app.exec())