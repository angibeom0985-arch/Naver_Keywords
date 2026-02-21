import os
import sys
import json
import ast
import random
import socket
import requests
import re
import base64
import hmac
import hashlib
import glob
import atexit
import signal
import threading
import traceback
import subprocess
import io
import builtins
import uuid
import winreg
import ctypes
from collections import deque
from pathlib import Path
from datetime import datetime, timedelta
import time
import urllib.parse
from urllib.parse import quote, unquote
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from webdriver_manager.chrome import ChromeDriverManager

# comment removed (encoding issue)
try:
    from bs4 import BeautifulSoup
    BEAUTIFULSOUP_AVAILABLE = True
except ImportError:
    print("BeautifulSoup가 설치되지 않았습니다: pip install beautifulsoup4")
    BEAUTIFULSOUP_AVAILABLE = False

# comment removed (encoding issue)

# comment removed (encoding issue)
try:
    import PyQt6
    qt_plugin_path = os.path.join(
        os.path.dirname(PyQt6.__file__),
        'Qt6',
        'plugins')
    if os.path.exists(qt_plugin_path):
        os.environ['QT_PLUGIN_PATH'] = qt_plugin_path
        print(f"Qt 플러그인 경로 설정: {qt_plugin_path}")
    else:
        # comment removed (encoding issue)
        alt_path = os.path.join(
            os.path.dirname(PyQt6.__file__),
            'Qt',
            'plugins')
        if os.path.exists(alt_path):
            os.environ['QT_PLUGIN_PATH'] = alt_path
            print(f"Qt 플러그인 대체 경로 설정: {alt_path}")
except ImportError:
    pass

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QDialog, QWidget, QVBoxLayout, QHBoxLayout, 
    QLabel, QLineEdit, QPushButton, QTextEdit, QMessageBox, 
    QScrollArea, QFrame, QGridLayout, QGroupBox, QComboBox, 
    QCheckBox, QFileDialog, QProgressBar, QStatusBar, QSizePolicy,
    QTabWidget, QTabBar, QSpinBox, QDoubleSpinBox, QTableWidget, QTableWidgetItem, QHeaderView,
    QAbstractItemView, QScroller, QStackedLayout
)
from PyQt6.QtCore import QThread, pyqtSignal, Qt, QEvent, QSettings, QDir, QTimer, QUrl, QRect
from PyQt6.QtGui import (
    QPixmap, QKeySequence, QFont, QTransform, QIcon, QShortcut,
    QPainter, QColor, QDesktopServices, QCursor, QPen
)

import pandas as pd

# comment removed (encoding issue)
try:
    from bs4 import BeautifulSoup
    BEAUTIFULSOUP_AVAILABLE = True
except ImportError:
    print("BeautifulSoup가 설치되지 않았습니다: pip install beautifulsoup4")
    BEAUTIFULSOUP_AVAILABLE = False

# comment removed (encoding issue)
OPENAI_AVAILABLE = False

# comment removed (encoding issue)
NAVER_GREEN = "#03c75a"
NAVER_GREEN_DARK = "#028a4a"
NAVER_GREEN_LIGHT = "#e8f5f0"
NAVER_GREEN_ULTRA_LIGHT = "#f0faf7"
WHITE_COLOR = "#ffffff"
TEXT_PRIMARY = "#212529"
TEXT_SECONDARY = "#6c757d"
BACKGROUND_MAIN = "#f0faf7"
BACKGROUND_CARD = "#ffffff"
BORDER_COLOR = "#d4edda"
BORDER_FOCUS = "#03c75a"
PLACEHOLDER_COLOR = "#8a8a8a"
SELENIUM_HEADLESS = False

# comment removed (encoding issue)
_current_window = None
_crash_save_enabled = True
MACHINE_ID_GUARD_HASH = "9808ecd261f917072ef4aa92222de467ff3950794597460c6b0503ad1417806c"
MACHINE_ID_APPROVAL_FILE = 'machine_id_change_approval.txt'
MACHINE_ID_APPROVAL_TOKEN = 'I_APPROVE_MACHINE_ID_CHANGE'


class ApiUsageReporter:
    def __init__(self):
        self._lock = threading.Lock()
        self.machine_id = ""
        self.webhook_url = ""
        self.webhook_token = ""
        self.local_total = 0

    def configure(self, machine_id="", webhook_url="", webhook_token=""):
        with self._lock:
            if machine_id:
                self.machine_id = str(machine_id).strip()
            self.webhook_url = str(webhook_url or "").strip()
            self.webhook_token = str(webhook_token or "").strip()

    def increment(self, delta=1):
        with self._lock:
            self.local_total += int(delta)
            machine_id = self.machine_id
            webhook_url = self.webhook_url
            webhook_token = self.webhook_token
            local_total = self.local_total

        # 로컬 누적 저장
        try:
            usage_file = get_app_base_dir() / "api_usage_local.json"
            payload_local = {
                "machine_id": machine_id,
                "local_total": local_total,
                "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            }
            with open(usage_file, "w", encoding="utf-8") as f:
                json.dump(payload_local, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

        # 스프레드시트 업데이트(선택): Apps Script 웹앱 URL 필요
        if not webhook_url:
            return
        try:
            body = {
                "machine_id": machine_id,
                "delta": int(delta),
                "local_total": local_total,
                "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            }
            headers = {"Content-Type": "application/json"}
            if webhook_token:
                headers["X-Usage-Token"] = webhook_token
            requests.post(webhook_url, headers=headers, json=body, timeout=2.5)
        except Exception:
            pass


API_USAGE_REPORTER = ApiUsageReporter()

def activate_korean_input_method():
    """Best-effort: switch current input language to Korean on Windows."""
    if os.name != "nt":
        return
    try:
        user32 = ctypes.WinDLL("user32", use_last_error=True)
        hkl = user32.LoadKeyboardLayoutW("00000412", 0x00000001)  # KLF_ACTIVATE
        if not hkl:
            return
        user32.ActivateKeyboardLayout(hkl, 0)
        hwnd = user32.GetForegroundWindow()
        if hwnd:
            user32.PostMessageW(hwnd, 0x0050, 0, hkl)  # WM_INPUTLANGCHANGEREQUEST
    except Exception:
        pass


class KoreanDefaultLineEdit(QLineEdit):
    """LineEdit that tries to default to Korean input when focused."""

    def focusInEvent(self, event):
        super().focusInEvent(event)
        activate_korean_input_method()


class SpiralSpinner(QWidget):
    """Fixed-size painted spinner to avoid glyph width jitter."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self._angle = 0
        self._mode = "light"
        self.setObjectName("loadingSpinner")
        self.setFixedSize(16, 16)
        self.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)

    def set_mode(self, mode):
        self._mode = "dark" if str(mode).strip().lower() == "dark" else "light"
        self.update()

    def step(self):
        self._angle = (self._angle + 24) % 360
        self.update()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)

        size = min(self.width(), self.height())
        pad = 2
        rect = QRect(pad, pad, size - (pad * 2), size - (pad * 2))

        if self._mode == "dark":
            base_color = QColor("#355444")
            spin_color = QColor("#88e0b4")
            tail_color = QColor("#4fc892")
        else:
            base_color = QColor("#cde9d8")
            spin_color = QColor("#1f6a49")
            tail_color = QColor("#03c75a")

        base_pen = QPen(base_color, 1.8)
        base_pen.setCapStyle(Qt.PenCapStyle.RoundCap)
        painter.setPen(base_pen)
        painter.drawEllipse(rect)

        head_pen = QPen(spin_color, 2.2)
        head_pen.setCapStyle(Qt.PenCapStyle.RoundCap)
        painter.setPen(head_pen)
        start = int((90 - self._angle) * 16)
        painter.drawArc(rect, start, int(-210 * 16))

        inner = QRect(rect.x() + 3, rect.y() + 3, rect.width() - 6, rect.height() - 6)
        tail_pen = QPen(tail_color, 1.6)
        tail_pen.setCapStyle(Qt.PenCapStyle.RoundCap)
        painter.setPen(tail_pen)
        tail_start = int((140 - self._angle) * 16)
        painter.drawArc(inner, tail_start, int(-150 * 16))


# comment removed (encoding issue)
def get_icon_path():
    """Description"""
    try:
        # comment removed (encoding issue)
        meipass = getattr(sys, "_MEIPASS", None)
        if meipass:
            # comment removed (encoding issue)
            icon_path = os.path.join(str(meipass), 'auto_naver.ico')
            if os.path.exists(icon_path):
                return icon_path
        
        # comment removed (encoding issue)
        if getattr(sys, 'frozen', False):
            # comment removed (encoding issue)
            exe_dir = os.path.dirname(sys.executable)
            icon_path = os.path.join(exe_dir, 'auto_naver.ico')
            if os.path.exists(icon_path):
                return icon_path
        
        # comment removed (encoding issue)
        script_dir = os.path.dirname(os.path.abspath(__file__))
        icon_path = os.path.join(script_dir, 'auto_naver.ico')
        if os.path.exists(icon_path):
            return icon_path
        
        # comment removed (encoding issue)
        assets_icon = os.path.join(script_dir, 'assets', 'auto_naver.ico')
        if os.path.exists(assets_icon):
            return assets_icon
        
        # comment removed (encoding issue)
        cwd_icon = os.path.join(os.getcwd(), 'auto_naver.ico')
        if os.path.exists(cwd_icon):
            return cwd_icon
            
    except Exception as e:
        safe_print(f"아이콘 경로 확인 중 오류: {e}")
    
    return None


def get_app_base_dir():
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def sanitize_display_text(text):
    if text is None:
        return ""
    clean = str(text).replace("�", " ").replace("?", " ")
    clean = re.sub(r"[^\w\s가-힣.,:;!()\[\]\-_/+|%#@'\"~]", " ", clean)
    clean = re.sub(r"\s{2,}", " ", clean).strip()
    return clean or "진행 상태 업데이트"


def safe_print(*args, **kwargs):
    normalized = [sanitize_display_text(arg) for arg in args]
    builtins.print(*normalized, **kwargs)


def get_embedded_api_credentials():
    return {
        "searchad_access_key": "01000000000e64e897c68f6e79f36bca6a49962aa41145858136b2e00f482bc1677f3a1446",
        "searchad_secret_key": "AQAAAABfmxB5yE7SY0Bij8RNJKkZ4af7UA++fTEhxfgv/FKteQ==",
        "searchad_customer_id": "3010221",
        "naver_client_id": "GSiFqhyeZrtRo1PAR0RF",
        "naver_client_secret": "1TV2afJdhU",
        "usage_webhook_url": "https://script.google.com/macros/s/AKfycbz3WZH9J1uRwXzzsFLlyH3gZkwI7eFrO_fSxDdOa7bLk0TU0_WXZaa3XC1marNnRBebVw/exec?token=david_usage_2026_01",
        "usage_webhook_token": "david_usage_2026_01",
    }


def load_api_credentials_from_file():
    required_keys = [
        "searchad_access_key",
        "searchad_secret_key",
        "searchad_customer_id",
        "naver_client_id",
        "naver_client_secret",
    ]
    api_file = get_app_base_dir() / "api_keys.json"
    embedded = get_embedded_api_credentials()
    credentials = {key: str(embedded.get(key, "")).strip() for key in required_keys}
    credentials["usage_webhook_url"] = str(embedded.get("usage_webhook_url", "")).strip()
    credentials["usage_webhook_token"] = str(embedded.get("usage_webhook_token", "")).strip()

    if api_file.exists():
        try:
            with open(api_file, "r", encoding="utf-8-sig") as f:
                data = json.load(f)
            if not isinstance(data, dict):
                raise ValueError("api_keys.json 형식이 올바르지 않습니다.")

            for key in required_keys:
                value = str(data.get(key, "")).strip()
                if value:
                    credentials[key] = value
            for key in ["usage_webhook_url", "usage_webhook_token"]:
                value = str(data.get(key, "")).strip()
                if value:
                    credentials[key] = value
        except Exception as e:
            safe_print(f"api_keys.json 읽기 실패(내장 키 사용): {e}")

    missing = []
    for key in required_keys:
        if not credentials.get(key):
            missing.append(key)

    if missing:
        raise ValueError("필수 API 키 누락: " + ", ".join(missing))
    return credentials, api_file


def get_machine_id():
    """안정적인 머신 ID 생성/조회 (업데이트/재빌드 시에도 동일 PC면 유지)."""
    cache_path = Path.home() / ".auto_naver_machine_id.txt"

    # 0) 캐시 우선 사용: 기능 업데이트/빌드 변경으로 추출 경로가 달라도 ID가 유지됨
    try:
        if cache_path.exists():
            cached = cache_path.read_text(encoding="utf-8-sig").strip()
            if cached:
                return cached
    except Exception:
        pass

    parts = []

    # 1) Windows MachineGuid (가장 안정적)
    try:
        with winreg.OpenKey(
            winreg.HKEY_LOCAL_MACHINE,
            r"SOFTWARE\Microsoft\Cryptography"
        ) as key:
            machine_guid, _ = winreg.QueryValueEx(key, "MachineGuid")
            machine_guid = str(machine_guid).strip()
            if machine_guid:
                parts.append(f"GUID:{machine_guid}")
    except Exception:
        pass

    # 2) 시스템 UUID
    try:
        cmd = 'powershell "Get-CimInstance -Class Win32_ComputerSystemProduct | Select-Object -ExpandProperty UUID"'
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        output = subprocess.check_output(cmd, startupinfo=startupinfo, shell=True).decode("utf-8", errors="ignore").strip()
        if output and len(output) > 10:
            parts.append(f"CSUUID:{output}")
    except Exception:
        pass

    # 3) BIOS 시리얼
    try:
        output = subprocess.check_output("wmic bios get serialnumber", shell=True).decode("utf-8", errors="ignore")
        bios = "".join(output.splitlines()[1:]).strip()
        if bios:
            parts.append(f"BIOS:{bios}")
    except Exception:
        pass

    # 4) MAC (가능한 경우)
    try:
        mac = uuid.getnode()
        if mac:
            parts.append(f"MAC:{mac:012x}")
    except Exception:
        pass

    if not parts:
        parts.append(f"FALLBACK:{uuid.uuid4()}")

    raw = "|".join(parts)
    stable_id = "MID-" + hashlib.sha256(raw.encode("utf-8")).hexdigest()[:32].upper()

    try:
        cache_path.write_text(stable_id, encoding="utf-8")
    except Exception:
        pass

    return stable_id


def check_license_from_sheet(machine_id):
    """Description"""
    sheet_url = "https://docs.google.com/spreadsheets/d/1YxiLMs7NEbKj0ZuEhx8zdKH-Hz0co2dd3OFxbZSEQS0/export?format=csv&gid=0"
    try:
        safe_print(f"라이선스 확인 중... ID: {machine_id}")
        response = requests.get(sheet_url, timeout=5)
        if response.status_code == 200:
            # comment removed (encoding issue)
            df = pd.read_csv(io.StringIO(response.text))
            
            # comment removed (encoding issue)
            if len(df.columns) >= 4:
                # comment removed (encoding issue)
                df.iloc[:, 2] = df.iloc[:, 2].astype(str).str.strip()
                target_row = df[df.iloc[:, 2] == str(machine_id)]
                
                if not target_row.empty:
                    expiration_date = str(target_row.iloc[0, 3]).strip()
                    return expiration_date
                
        return None
    except Exception as e:
        safe_print(f"라이선스 확인 실패: {e}")
        return None


def verify_machine_id_guard():
    try:
        # onefile/onedir EXE에서는 소스 파일 직접 읽기가 불안정하므로 런타임 검증을 건너뜀
        if getattr(sys, "frozen", False):
            return True
        source_text = Path(__file__).read_text(encoding="utf-8-sig")
        tree = ast.parse(source_text)
        lines = source_text.splitlines()
        targets = {}
        for node in tree.body:
            if isinstance(node, ast.FunctionDef) and node.name in ("get_machine_id", "check_license_from_sheet"):
                targets[node.name] = "\n".join(lines[node.lineno - 1: node.end_lineno])
        if len(targets) != 2:
            raise RuntimeError("머신 ID 보호 함수 파싱 실패")

        payload = "\n\n".join(targets[k] for k in sorted(targets))
        current_hash = hashlib.sha256(payload.encode("utf-8")).hexdigest()
        if current_hash == MACHINE_ID_GUARD_HASH:
            return True

        approval_file = get_app_base_dir() / MACHINE_ID_APPROVAL_FILE
        approved = False
        if approval_file.exists():
            approved = MACHINE_ID_APPROVAL_TOKEN in approval_file.read_text(encoding="utf-8-sig")
        if approved:
            return True

        raise RuntimeError(
            "머신ID/사용기간 제한 로직 변경이 감지되었습니다. "
            "소유자 승인 파일이 없어 실행을 중단합니다."
        )
    except Exception as e:
        safe_print(f"보호 검증 실패: {e}")
        return False


class UnregisteredDialog(QDialog):
    """미등록 기기 안내 다이얼로그"""
    def __init__(self, machine_id):
        super().__init__()
        self.setWindowTitle("프로그램 사용 권한")
        self.setFixedSize(500, 380)
        self.setWindowFlags(Qt.WindowType.WindowStaysOnTopHint | Qt.WindowType.WindowCloseButtonHint)
        self.setStyleSheet("QDialog { background-color: #ffffff; }")
        
        layout = QVBoxLayout(self)
        layout.setSpacing(20)
        layout.setContentsMargins(30, 30, 30, 30)
        
        # comment removed (encoding issue)
        warning_layout = QHBoxLayout()
        warning_icon = QLabel("⚠")
        warning_icon.setStyleSheet("font-size: 40px; background-color: transparent;")
        warning_layout.addWidget(warning_icon)
        
        warning_text = QLabel("등록되지 않은 사용자입니다.")
        warning_text.setStyleSheet("font-size: 20px; font-weight: bold; color: #FF4444; background-color: transparent;")
        warning_layout.addWidget(warning_text)
        warning_layout.addStretch()
        warning_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addLayout(warning_layout)
        
        # comment removed (encoding issue)
        blue_box = QFrame()
        blue_box.setStyleSheet("""
            QFrame {
                background-color: #E8F4FD;
                border-radius: 10px;
            }
        """)
        box_layout = QVBoxLayout(blue_box)
        box_layout.setSpacing(15)
        box_layout.setContentsMargins(20, 20, 20, 20)
        
        info_label = QLabel("아래 머신 ID를 판매자에게 전달해 주세요.")
        info_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        info_label.setStyleSheet("font-size: 15px; font-weight: bold; color: #0066CC; border: none;")
        box_layout.addWidget(info_label)
        
        # comment removed (encoding issue)
        id_layout = QHBoxLayout()
        self.id_input = QLineEdit(machine_id)
        self.id_input.setReadOnly(True)
        self.id_input.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.id_input.setStyleSheet("""
            QLineEdit {
                padding: 10px;
                font-size: 14px;
                color: #555555;
                background-color: white;
                border: 1px solid #D1D1D1;
                border-radius: 5px;
            }
        """)
        
        copy_btn = QPushButton("복사")
        copy_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        copy_btn.setFixedWidth(80)
        copy_btn.setStyleSheet("""
            QPushButton {
                background-color: #1A73E8;
                color: white;
                font-size: 14px;
                font-weight: bold;
                border-radius: 5px;
                padding: 10px;
            }
            QPushButton:hover {
                background-color: #1557B0;
            }
        """)
        copy_btn.clicked.connect(self.copy_to_clipboard)
        
        id_layout.addWidget(self.id_input)
        id_layout.addWidget(copy_btn)
        box_layout.addLayout(id_layout)
        
        layout.addWidget(blue_box)
        
        # comment removed (encoding issue)
        note_layout = QHBoxLayout()
        bulb_icon = QLabel("💡")
        bulb_icon.setStyleSheet("font-size: 16px; background-color: transparent;")
        note_text = QLabel("참고: PC를 변경하면 머신 ID가 바뀔 수 있습니다.")
        note_text.setStyleSheet("font-size: 13px; color: #888888; background-color: transparent;")
        note_layout.addWidget(bulb_icon)
        note_layout.addWidget(note_text)
        note_layout.addStretch()
        layout.addLayout(note_layout)
        
        # comment removed (encoding issue)
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        ok_btn = QPushButton("확인")
        ok_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        ok_btn.setFixedWidth(100)
        ok_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: {NAVER_GREEN};
                color: white;
                font-size: 15px;
                font-weight: bold;
                border-radius: 5px;
                padding: 10px 20px;
            }}
            QPushButton:hover {{
                background-color: {NAVER_GREEN_DARK};
            }}
        """)
        ok_btn.clicked.connect(self.accept)
        btn_layout.addWidget(ok_btn)
        layout.addLayout(btn_layout)

    def copy_to_clipboard(self):
        clipboard = QApplication.clipboard()
        if clipboard is not None:
            clipboard.setText(self.id_input.text())
        sender = self.sender()
        if isinstance(sender, QPushButton):
            sender.setText("완료")
            sender.setEnabled(False)
            QTimer.singleShot(2000, lambda: self._reset_btn(sender))

    def _reset_btn(self, btn):
        btn.setText("복사")
        btn.setEnabled(True)


class ExpiredDialog(QDialog):
    """사용 기간 만료 안내 다이얼로그"""
    def __init__(self, expiry_date):
        super().__init__()
        self.setWindowTitle("프로그램 사용 권한")
        self.setFixedSize(500, 380)
        self.setWindowFlags(Qt.WindowType.WindowStaysOnTopHint | Qt.WindowType.WindowCloseButtonHint)
        self.setStyleSheet("QDialog { background-color: #ffffff; }")
        
        layout = QVBoxLayout(self)
        layout.setSpacing(25)
        layout.setContentsMargins(30, 30, 30, 30)
        
        # comment removed (encoding issue)
        warning_layout = QHBoxLayout()
        warning_icon = QLabel("⚠")
        warning_icon.setStyleSheet("font-size: 40px; background-color: transparent;")
        warning_layout.addWidget(warning_icon)
        
        warning_text = QLabel("사용 기간이 만료되었습니다.")
        warning_text.setStyleSheet("font-size: 20px; font-weight: bold; color: #FF4444; background-color: transparent;")
        warning_layout.addWidget(warning_text)
        warning_layout.addStretch()
        warning_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addLayout(warning_layout)
        
        # comment removed (encoding issue)
        yellow_box = QFrame()
        yellow_box.setStyleSheet("""
            QFrame {
                background-color: #FFF8E1;
                border-radius: 10px;
            }
        """)
        box_layout = QVBoxLayout(yellow_box)
        box_layout.setSpacing(15)
        box_layout.setContentsMargins(20, 20, 20, 20)
        
        info_label = QLabel("기간 연장이 필요합니다. 아래 카카오톡으로 문의해 주세요.")
        info_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        info_label.setStyleSheet("font-size: 15px; color: #333333; border: none;")
        box_layout.addWidget(info_label)
        
        kakao_btn = QPushButton("카카오톡 바로가기")
        kakao_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        kakao_btn.setMinimumHeight(60)
        kakao_btn.setStyleSheet("""
            QPushButton {
                background-color: #1E6ECA;
                color: white;
                font-size: 18px;
                font-weight: bold;
                border-radius: 8px;
                padding: 15px;
                border: none;
            }
            QPushButton:hover {
                background-color: #1557B0;
            }
        """)
        kakao_btn.clicked.connect(self.open_kakao)
        box_layout.addWidget(kakao_btn)
        
        layout.addWidget(yellow_box)
        
        # comment removed (encoding issue)
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        ok_btn = QPushButton("확인")
        ok_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        ok_btn.setFixedWidth(100)
        ok_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: {NAVER_GREEN};
                color: white;
                font-size: 15px;
                font-weight: bold;
                border-radius: 5px;
                padding: 10px 20px;
            }}
            QPushButton:hover {{
                background-color: {NAVER_GREEN_DARK};
            }}
        """)
        ok_btn.clicked.connect(self.accept)
        btn_layout.addWidget(ok_btn)
        layout.addLayout(btn_layout)

    def open_kakao(self):
        url = QUrl("https://open.kakao.com/me/david0985")
        QDesktopServices.openUrl(url)


class ResizableTextEdit(QTextEdit):
    """Resizable text edit widget."""
    
    def __init__(self, parent=None, min_height=200, max_height=800):
        super().__init__(parent)
        self.min_height = min_height
        self.max_height = max_height
        self.resize_step = 30
        
    def wheelEvent(self, event):
        # comment removed (encoding issue)
        if event.modifiers() & Qt.KeyboardModifier.ControlModifier:
            # comment removed (encoding issue)
            delta = event.angleDelta().y()
            current_height = self.height()
            
            if delta > 0:
                new_height = min(current_height + self.resize_step, self.max_height)
            else:
                new_height = max(current_height - self.resize_step, self.min_height)
            
            # comment removed (encoding issue)
            self.setMaximumHeight(new_height)
            self.setMinimumHeight(new_height)
            
            # comment removed (encoding issue)
            event.accept()
        else:
            # comment removed (encoding issue)
            super().wheelEvent(event)


class SmartProgressTextEdit(ResizableTextEdit):
    """Description"""
    
    def __init__(self, parent=None, min_height=200, max_height=800):
        super().__init__(parent, min_height, max_height)
        self.user_is_scrolling = False
        self.last_scroll_time = 0
        self.auto_scroll_enabled = True
        self.search_widget = None
        self.last_search_text = ""
        
        # comment removed (encoding issue)
        scrollbar = self.verticalScrollBar()
        if scrollbar:
            scrollbar.valueChanged.connect(self._on_scroll_changed)
            scrollbar.sliderPressed.connect(self._on_user_scroll_start)
            scrollbar.sliderReleased.connect(self._on_user_scroll_end)
        
        # comment removed (encoding issue)
        self.search_shortcut = QShortcut(QKeySequence("Ctrl+F"), self)
        self.search_shortcut.activated.connect(self.show_search_dialog)
        
    def _on_scroll_changed(self, value):
        """Description"""
        import time
        current_time = time.time()
        
        # comment removed (encoding issue)
        if not self.user_is_scrolling and current_time - self.last_scroll_time > 1.0:
            self.auto_scroll_enabled = True
            
    def _on_user_scroll_start(self):
        """User started manual scrollbar interaction."""
        self.user_is_scrolling = True
        self.auto_scroll_enabled = False
        
    def _on_user_scroll_end(self):
        """User ended manual scrollbar interaction."""
        import time
        self.user_is_scrolling = False
        self.last_scroll_time = time.time()
        
        # comment removed (encoding issue)
        QTimer.singleShot(3000, self._enable_auto_scroll)
        
    def _enable_auto_scroll(self):
        """Description"""
        if not self.user_is_scrolling:
            self.auto_scroll_enabled = True
            
            # comment removed (encoding issue)
            QTimer.singleShot(3000, self._enable_auto_scroll)
            
    def wheelEvent(self, event):
        """Handle wheel event."""
        # comment removed (encoding issue)
        if not (event.modifiers() & Qt.KeyboardModifier.ControlModifier):
            import time
            self.auto_scroll_enabled = False
            self.last_scroll_time = time.time()
            # comment removed (encoding issue)
            QTimer.singleShot(3000, self._enable_auto_scroll)
            
        super().wheelEvent(event)
        
    def append_with_smart_scroll(self, text):
        """Description"""
        # comment removed (encoding issue)
        scrollbar = self.verticalScrollBar()
        was_at_bottom = False
        if scrollbar:
            was_at_bottom = scrollbar.value() >= scrollbar.maximum() - 10
        
        # comment removed (encoding issue)
        self.append(text)
        
        # comment removed (encoding issue)
        if scrollbar and self.auto_scroll_enabled and (was_at_bottom or scrollbar.maximum() == 0):
            scrollbar.setValue(scrollbar.maximum())
    
    def show_search_dialog(self):
        """로그 검색 다이얼로그 표시"""
        from PyQt6.QtWidgets import QInputDialog
        
        text, ok = QInputDialog.getText(
            self, 
            "로그 검색",
            "찾을 텍스트를 입력하세요:",
            text=self.last_search_text
        )
        
        if ok and text:
            self.last_search_text = text
            self.search_in_text(text)
    
    def search_in_text(self, search_text):
        """로그 텍스트 내 검색."""
        if not search_text:
            return
        
        # comment removed (encoding issue)
        text_content = self.toPlainText()
        
        # comment removed (encoding issue)
        cursor = self.textCursor()
        current_position = cursor.position()
        
        # comment removed (encoding issue)
        found_index = text_content.find(search_text, current_position)
        
        if found_index == -1:
            # comment removed (encoding issue)
            found_index = text_content.find(search_text)
            
        if found_index != -1:
            # comment removed (encoding issue)
            cursor.setPosition(found_index)
            cursor.setPosition(found_index + len(search_text), cursor.MoveMode.KeepAnchor)
            self.setTextCursor(cursor)
            self.ensureCursorVisible()
            QMessageBox.information(self, "검색 완료", f"'{search_text}' 텍스트를 찾았습니다.")
        else:
            QMessageBox.information(self, "검색 결과 없음", f"'{search_text}' 텍스트를 찾을 수 없습니다.")


class SortableNumericItem(QTableWidgetItem):
    def __init__(self, text, numeric_value):
        super().__init__(text)
        self._numeric_value = float(numeric_value)

    def __lt__(self, other):
        if isinstance(other, SortableNumericItem):
            return self._numeric_value < other._numeric_value
        return super().__lt__(other)


class InsightChartWidget(QWidget):
    def __init__(self, title, chart_type="line", parent=None):
        super().__init__(parent)
        self.title = title
        self.chart_type = chart_type  # "line" or "bar"
        self.labels = []
        self.values = []
        self.setMinimumHeight(220)

    def set_data(self, labels, values):
        self.labels = [str(x) for x in labels]
        self.values = [float(v) for v in values]
        self.update()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)
        rect = self.rect()

        painter.fillRect(rect, QColor("#ffffff"))
        panel_rect = rect.adjusted(0, 14, -1, -1)
        painter.setPen(QPen(QColor("#d4edda"), 1))
        painter.setBrush(QColor("#ffffff"))
        painter.drawRoundedRect(panel_rect, 8, 8)

        # Title badge (matches "인사이트" group title style)
        painter.setFont(QFont("Malgun Gothic", 10, QFont.Weight.Bold))
        metrics = painter.fontMetrics()
        title_w = metrics.horizontalAdvance(self.title) + 26
        title_h = 28
        title_x = rect.left() + 12
        title_y = rect.top()
        title_rect = QRect(title_x, title_y, title_w, title_h)
        painter.setPen(QPen(QColor("#03c75a"), 2))
        painter.setBrush(QColor("#ffffff"))
        painter.drawRoundedRect(title_rect, 8, 8)
        painter.setPen(QPen(QColor("#185a3a"), 1))
        painter.drawText(title_rect, Qt.AlignmentFlag.AlignCenter, self.title)

        if not self.values:
            painter.setPen(QPen(QColor("#6c757d"), 1))
            painter.setFont(QFont("Malgun Gothic", 9))
            painter.drawText(panel_rect, Qt.AlignmentFlag.AlignCenter, "데이터 없음")
            return

        left = panel_rect.left() + 68
        right = panel_rect.right() - 16
        top = panel_rect.top() + 34
        bottom = panel_rect.bottom() - 38
        plot_w = max(1, right - left)
        plot_h = max(1, bottom - top)

        max_val = max(self.values) if self.values else 1.0
        if max_val <= 0:
            max_val = 1.0

        painter.setPen(QPen(QColor("#e9ecef"), 1))
        for i in range(5):
            y = top + int(plot_h * i / 4)
            painter.drawLine(left, y, right, y)

        painter.setPen(QPen(QColor("#6c757d"), 1))
        painter.setFont(QFont("Malgun Gothic", 8))
        for i in range(5):
            y = top + int(plot_h * i / 4)
            v = max_val * (4 - i) / 4
            painter.drawText(
                panel_rect.left() + 2,
                y - 6,
                54,
                12,
                Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter,
                f"{v:.1f}"
            )

        painter.setPen(QPen(QColor("#adb5bd"), 1))
        painter.drawLine(left, top, left, bottom)
        painter.drawLine(left, bottom, right, bottom)

        if self.chart_type == "line":
            points = []
            count = len(self.values)
            for i, value in enumerate(self.values):
                x = left + int(plot_w * i / max(1, count - 1))
                y = bottom - int((value / max_val) * plot_h)
                points.append((x, y))
            painter.setPen(QPen(QColor("#2f9e44"), 2))
            for i in range(1, len(points)):
                painter.drawLine(points[i - 1][0], points[i - 1][1], points[i][0], points[i][1])
            painter.setBrush(QColor("#2f9e44"))
            for x, y in points:
                painter.drawEllipse(x - 2, y - 2, 4, 4)
        else:
            count = len(self.values)
            bar_w = max(10, int(plot_w / max(1, count * 1.7)))
            gap = max(4, int((plot_w - bar_w * count) / max(1, count - 1))) if count > 1 else 0
            x = left
            painter.setBrush(QColor("#4dabf7"))
            painter.setPen(QPen(QColor("#339af0"), 1))
            for value in self.values:
                h = int((value / max_val) * plot_h)
                painter.drawRect(x, bottom - h, bar_w, h)
                x += bar_w + gap

        painter.setPen(QPen(QColor("#6c757d"), 1))
        painter.setFont(QFont("Malgun Gothic", 8))
        if self.labels:
            if self.chart_type == "bar" and len(self.labels) <= 12:
                step = 1
            else:
                step = max(1, len(self.labels) // 6)
            if self.chart_type == "line":
                for i in range(0, len(self.labels), step):
                    x = left + int(plot_w * i / max(1, len(self.labels) - 1))
                    painter.drawText(x - 20, bottom + 16, 40, 12, Qt.AlignmentFlag.AlignCenter, self.labels[i])
            else:
                count = len(self.values)
                bar_w = max(10, int(plot_w / max(1, count * 1.7)))
                gap = max(4, int((plot_w - bar_w * count) / max(1, count - 1))) if count > 1 else 0
                x = left + bar_w // 2
                for i in range(0, len(self.labels), step):
                    xpos = x + i * (bar_w + gap)
                    painter.drawText(xpos - 32, bottom + 18, 64, 14, Qt.AlignmentFlag.AlignCenter, self.labels[i])


def emergency_save_data():
    """Emergency backup for crash situations."""
    global _current_window, _crash_save_enabled
    
    if not _crash_save_enabled or not _current_window:
        return
    
    try:
        safe_print("log update")
        
        saved_count = 0
        
        # comment removed (encoding issue)
        if hasattr(_current_window, 'active_threads') and _current_window.active_threads:
            save_dir = _current_window.save_path_input.text().strip()
            if not save_dir:
                # comment removed (encoding issue)
                desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
                save_dir = os.path.join(desktop_path, "keyword_results")
                try:
                    os.makedirs(save_dir, exist_ok=True)
                except Exception:
                    # comment removed (encoding issue)
                    save_dir = os.path.join(os.path.expanduser("~"), "Documents", "keyword_results")
                    os.makedirs(save_dir, exist_ok=True)
            
            current_time = datetime.now().strftime('%Y%m%d_%H%M%S')
            
            for thread in _current_window.active_threads:
                try:
                    if (thread.isRunning() and 
                        hasattr(thread, 'searcher') and 
                        thread.searcher and 
                        hasattr(thread.searcher, 'all_related_keywords') and 
                        thread.searcher.all_related_keywords):
                        
                        searcher = thread.searcher
                        base_keyword = thread.keyword
                        
                        # comment removed (encoding issue)
                        safe_keyword = re.sub(r'[^\w가-힣\s]', '', base_keyword).strip()[:20]
                        if not safe_keyword:
                            safe_keyword = "응급저장"
                        
                        emergency_file = os.path.join(save_dir, f"{safe_keyword}_응급저장_{current_time}.xlsx")
                        
                        # comment removed (encoding issue)
                        if searcher.save_recursive_results_to_excel(emergency_file):
                            safe_print("log update")
                            saved_count += 1
                except Exception as inner_e:
                    safe_print("log update")
                    continue
            
            if saved_count > 0:
                safe_print("log update")
            else:
                safe_print("log update")
            
    except Exception as e:
        safe_print("log update")
        # comment removed (encoding issue)
        try:
            search_thread = getattr(_current_window, "search_thread", None) if _current_window else None
            searcher = getattr(search_thread, "searcher", None) if search_thread else None
            all_keywords = getattr(searcher, "all_related_keywords", None) if searcher else None
            if all_keywords:
                
                backup_dir = os.path.join(os.getcwd(), "emergency_backup")
                os.makedirs(backup_dir, exist_ok=True)
                
                backup_file = os.path.join(backup_dir, f"emergency_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json")
                
                with open(backup_file, 'w', encoding='utf-8') as f:
                    json.dump(all_keywords, f, 
                             ensure_ascii=False, indent=2)
                
                safe_print("log update")
        except:
            safe_print("log update")


def handle_exception(exc_type, exc_value, exc_traceback):
    """Handle uncaught exceptions and trigger emergency backup."""
    global _crash_save_enabled
    
    if _crash_save_enabled:
        safe_print("log update")
        safe_print("log update")
        safe_print("log update")
        
        # comment removed (encoding issue)
        emergency_save_data()
        
        # comment removed (encoding issue)
        try:
            crash_dir = os.path.join(os.getcwd(), "crash_logs")
            os.makedirs(crash_dir, exist_ok=True)
            
            current_time = datetime.now().strftime('%Y%m%d_%H%M%S')
            crash_file = os.path.join(
    crash_dir, f"crash_log_{current_time}.txt")
            
            with open(crash_file, 'w', encoding='utf-8') as f:
                f.write(f"Crash time: {datetime.now()}\n")
                f.write(f"Exception type: {exc_type.__name__}\n")
                f.write(f"Exception message: {str(exc_value)}\n\n")
                f.write("Stack trace:\n")
                traceback.print_exception(
    exc_type, exc_value, exc_traceback, file=f)
            
            safe_print("log update")
        except:
            pass
    
    # comment removed (encoding issue)
    sys.__excepthook__(exc_type, exc_value, exc_traceback)


def handle_signal(signum, frame):
    """Description"""
    signal_names = {
        signal.SIGINT: "SIGINT (Ctrl+C)",
        signal.SIGTERM: "SIGTERM (terminate request)"
    }
    
    signal_name = signal_names.get(signum, f"Signal {signum}")
    safe_print("log update")
    
    emergency_save_data()
    
    # comment removed (encoding issue)
    if _current_window:
        _current_window.close()
    
    sys.exit(0)


class MultiKeywordTextEdit(QTextEdit):
    """Description"""
    search_requested = pyqtSignal()
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.min_height = 200
        self.max_height = 800
        self.resize_step = 30
        self._placeholder_text = ""
        
        # comment removed (encoding issue)
        self._cta_text = "키워드 공부하러 가기"
        self._cta_url = "https://cafe.naver.com/f-e/cafes/31118881/articles/2036?menuid=12&referrerAllArticles=false"
        self._link_rect = None
        
        # comment removed (encoding issue)
        self.setMouseTracking(True)
        
        # comment removed (encoding issue)
        self.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        
        # comment removed (encoding issue)
        self.setStyleSheet(f"""
            QTextEdit {{
                background-color: {BACKGROUND_CARD} !important;
                color: #333333 !important;
                border: 3px solid {BORDER_COLOR} !important;
                border-radius: 12px !important;
                padding: 16px !important;
                font-size: 14px !important;
                font-weight: normal !important;
                selection-background-color: {NAVER_GREEN_LIGHT} !important;
                selection-color: #333333 !important;
                line-height: 180%;
            }}
            QTextEdit:focus {{
                border: 3px solid {NAVER_GREEN} !important;
            }}
            QScrollBar:vertical {{
                background: {NAVER_GREEN_LIGHT};
                width: 12px;
                border-radius: 6px;
            }}
            QScrollBar::handle:vertical {{
                background: {NAVER_GREEN};
                border-radius: 6px;
                min-height: 20px;
            }}
            QScrollBar::handle:vertical:hover {{
                background: {NAVER_GREEN_DARK};
            }}
        """)
        
    def setPlaceholderText(self, text):
        """Description"""
        self._placeholder_text = text
        self.update()
        
    def paintEvent(self, event):
        """Custom paint event for placeholder and CTA."""
        super().paintEvent(event)
        
        # comment removed (encoding issue)
        if not self.toPlainText().strip() and not self.hasFocus():
            viewport = self.viewport()
            if viewport is None:
                return
            painter = QPainter(viewport)
            painter.setRenderHint(QPainter.RenderHint.Antialiasing)
            
            # comment removed (encoding issue)
            painter.setPen(QColor("#777777"))
            font = self.font()
            font.setPointSize(11)
            painter.setFont(font)
            
            lines = self._placeholder_text.split('\n')
            metrics = painter.fontMetrics()
            line_height = metrics.height()
            line_spacing = 2.2
            text_block_height = (len(lines) * line_height * line_spacing)
            
            viewport_rect = viewport.rect()
            padding_left = 60
            
            # comment removed (encoding issue)
            link_h = 40
            spacing_between = 20
            
            # comment removed (encoding issue)
            total_content_height = text_block_height + spacing_between + link_h
            start_y = (viewport_rect.height() - total_content_height) / 2 + metrics.ascent()
            
            current_y = start_y
            
            # comment removed (encoding issue)
            for i, line in enumerate(lines):
                painter.drawText(int(viewport_rect.left() + padding_left), int(current_y), line)
                current_y += (line_height * line_spacing)
            
            current_y += spacing_between
            
            # comment removed (encoding issue)
            
            # comment removed (encoding issue)
            link_font = self.font()
            link_font.setPointSize(11)
            link_font.setUnderline(True)
            link_font.setBold(True)
            painter.setFont(link_font)
            painter.setPen(QColor("#0066CC"))
            
            link_metrics = painter.fontMetrics()
            link_width = link_metrics.horizontalAdvance(self._cta_text)
            link_x = (viewport_rect.width() - link_width) / 2
            
            painter.drawText(int(link_x), int(current_y + link_metrics.ascent()), self._cta_text)
            
            # comment removed (encoding issue)
            self._link_rect = QRect(int(link_x), int(current_y), int(link_width), int(link_metrics.height() + 10))

    def mouseMoveEvent(self, event):
        """Handle mouse move for link hover."""
        viewport = self.viewport()
        if viewport is None:
            return
        if isinstance(self._link_rect, QRect) and self._link_rect.contains(event.pos()) and not self.toPlainText().strip():
            viewport.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        else:
            viewport.setCursor(QCursor(Qt.CursorShape.IBeamCursor))
        super().mouseMoveEvent(event)

    def mousePressEvent(self, event):
        """Description"""
        if isinstance(self._link_rect, QRect) and self._link_rect.contains(event.pos()) and not self.toPlainText().strip():
            QDesktopServices.openUrl(QUrl(self._cta_url))
            return

        super().mousePressEvent(event)
    
    def keyPressEvent(self, event):
        """Description"""
        if event.key() == Qt.Key.Key_Return or event.key() == Qt.Key.Key_Enter:
            if event.modifiers() == Qt.KeyboardModifier.ShiftModifier:
                # comment removed (encoding issue)
                super().keyPressEvent(event)
            else:
                # comment removed (encoding issue)
                self.search_requested.emit()
                event.accept()
        else:
            super().keyPressEvent(event)

    def wheelEvent(self, event):
        """Description"""
        # comment removed (encoding issue)
        if event.modifiers() & Qt.KeyboardModifier.ControlModifier:
            # comment removed (encoding issue)
            delta = event.angleDelta().y()
            current_height = self.height()
            
            if delta > 0:
                new_height = min(current_height + self.resize_step, self.max_height)
            else:
                new_height = max(current_height - self.resize_step, self.min_height)
            
            # comment removed (encoding issue)
            self.setMaximumHeight(new_height)
            self.setMinimumHeight(new_height)
            
            # comment removed (encoding issue)
            event.accept()
        else:
            # comment removed (encoding issue)
            super().wheelEvent(event)


class NaverMobileSearchScraper:
    """Description"""
    
    def __init__(self, driver=None):
        self.session = requests.Session()
        self.driver = driver
        self.results = []
        self.searched_keywords = set()
        self.save_dir = ""
        self.extracted_keywords = set()
        self.is_running = True
        self.all_related_keywords = []
        self.base_keyword = ""
        self.processed_keywords = set()
        self.search_thread: QThread | None = None
        
        # comment removed (encoding issue)
        self.session.headers.update({
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8',
            'Accept-Language': 'ko-KR,ko;q=0.8,en-US;q=0.5,en;q=0.3',
            'Accept-Encoding': 'gzip, deflate, br',
            'DNT': '1',
            'Connection': 'keep-alive',
            'Upgrade-Insecure-Requests': '1',
        })

    def search_keyword(self, keyword, progress_callback=None):
        """Description"""
        try:
            if progress_callback:
                progress_callback("progress update")
            
            encoded_keyword = urllib.parse.quote(keyword)
            search_url = f"https://search.naver.com/search.naver?where=nexearch&sm=top_hty&fbm=0&ie=utf8&query={encoded_keyword}"
            
            response = self.session.get(search_url, timeout=10)
            response.raise_for_status()
            
            if progress_callback:
                progress_callback("progress update")
            
            return response.text
            
        except Exception as e:
            if progress_callback:
                progress_callback("progress update")
            return None

    # extract_autocomplete_keywords (requests version) removed to avoid duplication
    pass

    def extract_related_keywords(self, keyword, progress_callback=None):
        """Description"""
        keywords = []
        
        try:
            if not BEAUTIFULSOUP_AVAILABLE:
                if progress_callback:
                    progress_callback("progress update")
                return keywords
            
            html_content = self.search_keyword(keyword, progress_callback)
            if not html_content:
                return keywords
            
            soup = BeautifulSoup(html_content, 'html.parser')
            
            if progress_callback:
                progress_callback("progress update")
            
            # comment removed (encoding issue)
            related_selectors = [
                '.related_srch a',
                '.lst_related a', 
                '.keyword a',
                '.related_keyword a',
                'a[data-area="relate"]'
            ]
            
            found_count = 0
            for selector in related_selectors:
                elements = soup.select(selector)
                if elements and progress_callback:
                    progress_callback("progress update")
                
                for element in elements:
                    try:
                        keyword_text = element.get_text(strip=True)
                        if keyword_text and keyword_text not in keywords:
                            keywords.append(keyword_text)
                            found_count += 1
                            if progress_callback:
                                progress_callback("progress update")
                    except:
                        continue
            
            if progress_callback:
                progress_callback("progress update")
            
            return keywords
            
        except Exception as e:
            if progress_callback:
                progress_callback("progress update")
            return []

    def check_internet_connection(self):
        """Description"""
        try:
            response = requests.get("https://www.naver.com", timeout=5)
            return response.status_code == 200
        except:
            try:
                socket.create_connection(("8.8.8.8", 53), timeout=3)
                return True
            except:
                return False

    def check_pause_status(self, progress_callback=None):
        """Description"""
        if not self.is_running:
            return False
        
        if self.search_thread:
            if bool(getattr(self.search_thread, "is_paused", False)):
                if progress_callback:
                    progress_callback("progress update")
                
                while bool(getattr(self.search_thread, "is_paused", False)) and self.is_running:
                    time.sleep(0.5)
                
                if not self.is_running:
                    return False
                
                if progress_callback:
                    progress_callback("progress update")
        
        if not self.check_internet_connection():
            if progress_callback:
                progress_callback("progress update")
            
            connection_wait_count = 0
            while not self.check_internet_connection() and self.is_running:
                time.sleep(2)
                connection_wait_count += 1
                
                if self.search_thread and bool(getattr(self.search_thread, "is_paused", False)):
                    if progress_callback:
                        progress_callback("인터넷 연결 대기 중 일시정지됨")
                    break
                
                if connection_wait_count % 5 == 0 and progress_callback:
                    progress_callback("progress update")
            
            if not self.is_running:
                return False
            
            if self.check_internet_connection() and progress_callback:
                progress_callback("progress update")
        
        return True

    def initialize_browser(self):
        """Description"""
        try:
            safe_print("log update")
            
            driver_path = None
            
            # comment removed (encoding issue)
            import shutil
            
            # comment removed (encoding issue)
            base_paths = []
            if getattr(sys, 'frozen', False):
                base_paths.append(os.path.dirname(sys.executable))
                meipass = getattr(sys, "_MEIPASS", None)
                if meipass:
                    base_paths.append(str(meipass))
            base_paths.append(os.getcwd())
            
            for base_path in base_paths:
                local_driver = os.path.join(base_path, "chromedriver.exe")
                if os.path.exists(local_driver):
                    safe_print("log update")
                    driver_path = local_driver
                    break
            
            # comment removed (encoding issue)
            if not driver_path:
                try:
                    from webdriver_manager.chrome import ChromeDriverManager
                    safe_print("log update")
                    # comment removed (encoding issue)
                    driver_path = ChromeDriverManager().install()
                    safe_print("log update")
                except Exception as e:
                    safe_print("log update")
            
            # comment removed (encoding issue)
            if not driver_path and shutil.which("chromedriver"):
                driver_path = "chromedriver"
                safe_print("log update")
            
            if not driver_path:
                raise Exception(
                    "ChromeDriver를 찾을 수 없습니다.\n"
                    "프로그램 폴더에 'chromedriver.exe'를 넣거나\n"
                    "인터넷 연결 상태를 확인해 주세요."
                )

            # comment removed (encoding issue)
            try:
                service = Service(driver_path)
            except:
                if driver_path == "chromedriver":
                    service = Service()
                else:
                    service = Service(executable_path=driver_path)

            # comment removed (encoding issue)
            if os.name == 'nt':
                try:
                    startupopt = subprocess.STARTUPINFO()
                    startupopt.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                    service.creation_flags = subprocess.CREATE_NO_WINDOW
                except:
                    pass

            options = webdriver.ChromeOptions()
            
            # comment removed (encoding issue)
            if SELENIUM_HEADLESS:
                options.add_argument("--headless")
                safe_print("Selenium headless mode enabled")
            else:
                safe_print("Selenium visible mode enabled")
            
            options.add_argument("--window-size=1920,1080")
            options.add_argument("--start-maximized")
            
            # comment removed (encoding issue)
            options.add_argument(
                "--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")

            # comment removed (encoding issue)
            options.add_argument("--disable-web-security")
            options.add_argument("--allow-running-insecure-content")
            options.add_argument("--force-device-scale-factor=1")
            options.add_argument("--disable-features=VizDisplayCompositor")
            
            # comment removed (encoding issue)
            options.add_argument("--no-sandbox")
            options.add_argument("--disable-dev-shm-usage") 
            options.add_argument("--disable-gpu")
            options.add_argument("--disable-extensions")
            options.add_argument("--disable-plugins")
            options.add_argument("--disable-images")
            
            # comment removed (encoding issue)
            options.add_argument("--disable-background-networking")
            options.add_argument("--disable-background-timer-throttling")
            options.add_argument("--disable-renderer-backgrounding")
            options.add_argument("--disable-backgrounding-occluded-windows")
            options.add_argument("--disable-features=TranslateUI")
            options.add_argument("--disable-default-apps")
            options.add_argument("--disable-sync")
            options.add_argument("--disable-notifications")
            options.add_argument("--disable-popup-blocking")
            options.add_argument("--aggressive-cache-discard")
            options.add_argument("--memory-pressure-off")
            options.add_argument("--max_old_space_size=4096")
            
            # comment removed (encoding issue)
            options.add_argument("--lang=en-US")
            options.add_argument("--disable-logging")
            options.add_argument("--disable-gpu-sandbox")
            options.add_argument("--log-level=3")
            options.add_argument("--silent")
            # comment removed (encoding issue)
            options.add_argument("--disable-features=TranslateUI,VizDisplayCompositor")
            options.add_argument("--disable-ipc-flooding-protection")
            
            # comment removed (encoding issue)
            options.add_experimental_option('useAutomationExtension', False)
            options.add_experimental_option("excludeSwitches", ["enable-logging", "enable-automation"])
            options.add_experimental_option("detach", True)
            options.add_argument("--disable-blink-features=AutomationControlled")
            
            # comment removed (encoding issue)
            options.add_argument("--network-timeout=15")
            options.add_argument("--page-load-strategy=eager")
            options.add_argument("--timeout=15000")
            options.add_argument("--dns-prefetch-disable")
            
            # comment removed (encoding issue)
            self.driver = webdriver.Chrome(service=service, options=options)
            safe_print("log update")
            
            # comment removed (encoding issue)
            self.driver.set_page_load_timeout(15)
            self.driver.implicitly_wait(3)
            
            safe_print("log update")
            
            # comment removed (encoding issue)
            max_retries = 3
            for attempt in range(max_retries):
                try:
                    # comment removed (encoding issue)
                    self.driver.get("https://m.naver.com")
                    
                    # comment removed (encoding issue)
                    self.driver.execute_script("""
                        // Read browser viewport size
                        var windowWidth = window.innerWidth;
                        var windowHeight = window.innerHeight;
                        
                        // Set viewport meta to current browser size
                        var existingMeta = document.querySelector('meta[name="viewport"]');
                        if (existingMeta) {
                            existingMeta.remove();
                        }
                        var meta = document.createElement('meta');
                        meta.name = 'viewport';
                        meta.content = 'width=' + windowWidth + ', initial-scale=1.0, maximum-scale=3.0, user-scalable=yes';
                        document.getElementsByTagName('head')[0].appendChild(meta);
                        
                        // Fit page layout to browser viewport
                        document.documentElement.style.width = '100%';
                        document.documentElement.style.height = '100%';
                        document.body.style.minWidth = windowWidth + 'px';
                        document.body.style.width = '100%';
                        document.body.style.height = '100vh';
                        document.body.style.transform = 'none';
                        document.body.style.transformOrigin = 'top left';
                        document.body.style.margin = '0';
                        document.body.style.padding = '0';
                        document.body.style.fontSize = Math.max(14, windowWidth / 100) + 'px';  // dynamic font size
                        
                        // Expand main containers to full width
                        var containers = document.querySelectorAll('.container, .wrap, .content_area, #wrap, .nx_wrap');
                        containers.forEach(function(container) {
                            container.style.maxWidth = '100%';
                            container.style.width = '100%';
                            container.style.minWidth = windowWidth + 'px';
                        });
                        
                        // Adjust search result area layout
                        var searchArea = document.querySelector('.TF7QLJYoGthrUnoIpxEj, .api_subject_bx, .search_result');
                        if (searchArea) {
                            searchArea.style.minHeight = (windowHeight - 200) + 'px';
                            searchArea.style.width = '100%';
                            searchArea.style.overflow = 'visible';
                            searchArea.style.maxWidth = 'none';
                        }
                        
                        console.log('Page layout adjusted for viewport: ' + windowWidth + 'x' + windowHeight);
                    """)
                    
                    # comment removed (encoding issue)
                    time.sleep(2)
                    
                    WebDriverWait(self.driver, 10).until(
                        EC.presence_of_element_located((By.TAG_NAME, "body"))
                    )
                    safe_print("log update")
                    time.sleep(2)
                    return True
                except Exception as e:
                    if attempt < max_retries - 1:
                        safe_print("log update")
                        time.sleep(3)
                    else:
                        safe_print("log update")
                        return False
            
            return True

        except Exception as e:
            error_msg = (
                f"브라우저 초기화 오류:\n{str(e)}\n\n"
                "Chrome 브라우저 설치 여부를 확인해 주세요."
            )
            safe_print("log update")
            
            # comment removed (encoding issue)
            try:
                global _current_window
                if _current_window:
                    # comment removed (encoding issue)
                    # comment removed (encoding issue)
                    # comment removed (encoding issue)
                    # comment removed (encoding issue)
                    from PyQt6.QtCore import QTimer
                    QTimer.singleShot(0, lambda: QMessageBox.critical(
                        _current_window, "브라우저 오류", error_msg))
            except:
                pass
                
            return False

    def search_keyword_mobile(self, keyword, progress_callback=None):
        """Search keyword on Naver mobile."""
        max_retries = 3
        
        for attempt in range(max_retries):
            try:
                if progress_callback:
                    if attempt > 0:
                        progress_callback("progress update")
                    else:
                        progress_callback("progress update")
                
                encoded_keyword = urllib.parse.quote(keyword)
                search_url = f"https://m.search.naver.com/search.naver?where=m&sm=mtp_hty.top&query={encoded_keyword}"
                
                if self.driver:
                    self.driver.get(search_url)
                
                if progress_callback:
                    progress_callback("progress update")
                
                # comment removed (encoding issue)
                time.sleep(random.uniform(1.5, 2.5))
                
                try:
                    if self.driver:
                        WebDriverWait(self.driver, 5).until(
                            EC.presence_of_element_located((By.TAG_NAME, "body"))
                        )
                except TimeoutException:
                    if attempt < max_retries - 1:
                        if progress_callback:
                            progress_callback("progress update")
                        continue
                    else:
                        if progress_callback:
                            progress_callback("progress update")
                        return False
                
                time.sleep(1)
                
                if progress_callback:
                    progress_callback("progress update")
                
                return True
                
            except Exception as e:
                error_msg = str(e)
                
                # comment removed (encoding issue)
                if "invalid session id" in error_msg.lower() or "no such session" in error_msg.lower() or "no such window" in error_msg.lower():
                    if progress_callback:
                        progress_callback("progress update")
                    
                    # comment removed (encoding issue)
                    if self.initialize_browser():
                        if progress_callback:
                            progress_callback("progress update")
                        continue
                    else:
                        if progress_callback:
                            progress_callback("progress update")
                        return False
                
                # comment removed (encoding issue)
                elif "timeout" in error_msg.lower() and "renderer" in error_msg.lower():
                    if progress_callback:
                        progress_callback("progress update")
                    
                    if attempt < max_retries - 1:
                        if progress_callback:
                            progress_callback("progress update")
                        
                        # comment removed (encoding issue)
                        if self.initialize_browser():
                            if progress_callback:
                                progress_callback("progress update")
                            continue
                        else:
                            if progress_callback:
                                progress_callback("progress update")
                    else:
                        if progress_callback:
                            progress_callback("progress update")
                        return False
                
                # comment removed (encoding issue)
                elif "timeout" in error_msg.lower():
                    if progress_callback:
                        progress_callback("progress update")
                    
                    if attempt < max_retries - 1:
                        if progress_callback:
                            progress_callback("progress update")
                        time.sleep(5)
                        continue
                    else:
                        if progress_callback:
                            progress_callback("progress update")
                        return False
                
                # comment removed (encoding issue)
                if attempt < max_retries - 1:
                    if progress_callback:
                        progress_callback("progress update")
                    time.sleep(3)
                    continue
                else:
                    if progress_callback:
                        progress_callback("progress update")
                    return False
        
        return False

    def extract_autocomplete_keywords(self, keyword, progress_callback=None):
        """Extract autocomplete keywords."""
        keywords = []
        max_retries = 2
        
        for attempt in range(max_retries):
            try:
                if not self.driver:
                    # comment removed (encoding issue)
                    if not self.initialize_browser():
                        return keywords
                
                if progress_callback:
                    if attempt > 0:
                        progress_callback("progress update")
                    else:
                        progress_callback("progress update")
                
                # comment removed (encoding issue)
                try:
                    driver = self.driver
                    if not driver:
                        return keywords
                    driver.set_page_load_timeout(15)
                    driver.get("https://m.naver.com")
                except TimeoutException:
                    if progress_callback:
                        progress_callback("progress update")
                    try:
                        if self.driver:
                            self.driver.execute_script("window.stop();")
                    except:
                        pass
                except Exception as e:
                    # comment removed (encoding issue)
                    raise e
                
                time.sleep(2)
            
                # comment removed (encoding issue)
                try:
                    from selenium.webdriver.support.ui import WebDriverWait
                    from selenium.webdriver.support import expected_conditions as EC
                    
                    # comment removed (encoding issue)
                    try:
                        driver = self.driver
                        if not driver:
                            return keywords
                        wait = WebDriverWait(driver, 10)
                        wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
                        if progress_callback:
                            progress_callback("progress update")
                    except TimeoutException:
                        if progress_callback:
                            progress_callback("progress update")
                            
                except Exception as e:
                    if progress_callback:
                        progress_callback("progress update")
            
                # comment removed (encoding issue)
                search_input = None
                search_selectors = [
                    '#nx_query',
                    'input.search_input',
                    'input[name="query"]',
                    'input[type="search"]'
                ]
                
                for selector in search_selectors:
                    try:
                        driver = self.driver
                        if not driver:
                            return keywords
                        search_input = driver.find_element(By.CSS_SELECTOR, selector)
                        if search_input and search_input.is_enabled():
                            if progress_callback:
                                progress_callback("progress update")
                            break
                    except:
                        continue
                
                if not search_input:
                    if progress_callback:
                        progress_callback("progress update")
                    return keywords
            
                # comment removed (encoding issue)
                try:
                    # comment removed (encoding issue)
                    driver = self.driver
                    if not driver:
                        return keywords
                    driver.execute_script("""
                        var input = arguments[0];
                        var keyword = arguments[1];
                        input.focus();
                        input.value = keyword;
                        input.dispatchEvent(new Event('input', { bubbles: true }));
                        input.dispatchEvent(new Event('keyup', { bubbles: true }));
                    """, search_input, keyword)
                    
                    # comment removed (encoding issue)
                    time.sleep(2)
                    
                    if progress_callback:
                        progress_callback("progress update")
                        
                except Exception as input_error:
                    if progress_callback:
                        progress_callback("progress update")
                        return keywords
            
                # comment removed (encoding issue)
                autocomplete_selectors = [
                    '#_nx_ac_layer_wrap ._nx_ac_text',
                    '._nx_ac_text',
                    '.u_atcp_txt',
                    '.autocomplete .item',
                    '.search_list li'
                ]
                
                found_count = 0
                for selector in autocomplete_selectors:
                    try:
                        driver = self.driver
                        if not driver:
                            return keywords
                        elements = driver.find_elements(By.CSS_SELECTOR, selector)
                        
                        for element in elements:
                            try:
                                if not element.is_displayed():
                                    continue
                                    
                                # comment removed (encoding issue)
                                keyword_text = element.get_attribute("textContent") or element.text
                                
                                if keyword_text:
                                    keyword_text = keyword_text.strip()
                                    
                                    # comment removed (encoding issue)
                                    if '<' in keyword_text:
                                        import re
                                        keyword_text = re.sub(r'<[^>]+>', '', keyword_text)
                                        keyword_text = keyword_text.strip()
                                    
                                    # comment removed (encoding issue)
                                        keyword_text = self.clean_duplicate_text(keyword_text)
                                        
                                    # comment removed (encoding issue)
                                    if (keyword.lower() in keyword_text.lower() and 
                                        keyword_text not in keywords and
                                        len(keyword_text) <= 50 and
                                        len(keyword_text) > 1):
                                            keywords.append(keyword_text)
                                            found_count += 1
                                            if progress_callback:
                                                progress_callback("progress update")
                            except Exception:
                                continue
                    except Exception:
                        continue
                
                # comment removed (encoding issue)
                keywords = list(set(keywords))
                keywords.sort()
                
                if progress_callback:
                    progress_callback("progress update")
                
                return keywords
            
            except Exception as e:
                error_msg = str(e)
                if "no such window" in error_msg.lower() or "invalid session id" in error_msg.lower():
                    if progress_callback:
                        progress_callback("progress update")
                    
                    if attempt < max_retries - 1:
                        if self.initialize_browser():
                            if progress_callback:
                                progress_callback("progress update")
                            continue
                
                if progress_callback:
                    progress_callback("progress update")
                
                if attempt == max_retries - 1:
                    return []
        
        return keywords

    def extract_related_keywords_new(self, current_keyword, progress_callback=None):
        """Extract related keywords from current page."""
        keywords = []
        
        try:
            if not self.driver:
                return keywords
            
            # comment removed (encoding issue)
            related_selectors = [
                '#_related_keywords .keyword a',
                '.related_srch .lst a',
                '.related_keyword a',
                '.lst_related a',
                '.keyword_area a',
                '.related_search a',
                '.keyword a'
            ]
            
            if progress_callback:
                progress_callback("progress update")
            
            found_count = 0
            
            for selector in related_selectors:
                try:
                    elements = self.driver.find_elements(By.CSS_SELECTOR, selector)
                    if len(elements) > 0 and progress_callback:
                        progress_callback("progress update")
                    
                    for element in elements:
                        try:
                            # comment removed (encoding issue)
                            keyword_text = ""
                            
                            # comment removed (encoding issue)
                            try:
                                keyword_text = element.text
                                if keyword_text:
                                    keyword_text = keyword_text.strip()
                            except:
                                keyword_text = ""

                            # comment removed (encoding issue)
                            if not keyword_text:
                                try:
                                    keyword_text = element.get_attribute("textContent")
                                    if keyword_text:
                                        keyword_text = keyword_text.strip()
                                except:
                                    keyword_text = ""

                            # comment removed (encoding issue)
                            if not keyword_text:
                                try:
                                    keyword_text = element.get_attribute("innerText")
                                    if keyword_text:
                                        keyword_text = keyword_text.strip()
                                except:
                                    keyword_text = ""

                            # comment removed (encoding issue)
                            if not keyword_text:
                                try:
                                    keyword_text = self.driver.execute_script("""
                                        var element = arguments[0];
                                        if (!element) return '';
                                        
                                            // Extract direct text from anchor element
                                            var textContent = element.textContent || element.innerText || '';
                                            
                                            // Trim and normalize whitespace
                                            return textContent.replace(/\\s+/g, ' ').trim();
                                    """, element)
                                except:
                                    keyword_text = ""
                            
                            # comment removed (encoding issue)
                            if keyword_text:
                                import re
                                # comment removed (encoding issue)
                                keyword_text = re.sub(r'<[^>]+>', '', keyword_text)
                                # comment removed (encoding issue)
                                keyword_text = re.sub(r'\s+', ' ', keyword_text)
                                # comment removed (encoding issue)
                                keyword_text = re.sub(r'[\u200b-\u200d\ufeff]', '', keyword_text)
                                # comment removed (encoding issue)
                                keyword_text = re.sub(r'\s+[가-힣]{1}$', '', keyword_text)
                                keyword_text = re.sub(r'\s+[a-zA-Z]{1}$', '', keyword_text)
                                keyword_text = keyword_text.strip()
                                
                            if keyword_text:
                                # comment removed (encoding issue)
                                keyword_text = self.clean_duplicate_text(keyword_text)
                                    
                                # comment removed (encoding issue)
                                # comment removed (encoding issue)
                                if keyword_text and not re.search(r'[가-힣]{1}$|[a-zA-Z]{1}$', keyword_text):
                                    # comment removed (encoding issue)
                                    if (keyword_text not in keywords and
                                        len(keyword_text) <= 50 and
                                        len(keyword_text) > 1):
                                        keywords.append(keyword_text)
                                        found_count += 1
                                        if progress_callback:
                                            progress_callback("progress update")
                                elif keyword_text and len(keyword_text) > 3:
                                    if (keyword_text not in keywords and 
                                        len(keyword_text) <= 50 and 
                                        len(keyword_text) > 1):
                                        keywords.append(keyword_text)
                                        found_count += 1
                                        if progress_callback:
                                            progress_callback("progress update")
                        except Exception as e:
                            continue
                except Exception as e:
                    continue
            
            # comment removed (encoding issue)
            keywords = list(set(keywords))
            keywords.sort()
            
            if progress_callback:
                progress_callback("progress update")
            
            return keywords
            
        except Exception as e:
            if progress_callback:
                progress_callback("progress update")
            return []

    def clean_duplicate_text(self, text):
        """Description"""
        if not text:
            return text

        text = text.strip()
        # comment removed (encoding issue)
        text = re.sub(r'\s+', ' ', text)
        
        # comment removed (encoding issue)
        words = text.split()
        unique_words = []
        seen_words = set()
        
        for word in words:
            word_lower = word.lower()
            if word_lower not in seen_words:
                unique_words.append(word)
                seen_words.add(word_lower)
        
        # comment removed (encoding issue)
        result = ' '.join(unique_words)
        return result

    def extract_together_keywords(self, current_keyword, progress_callback=None):
        """Description"""
        keywords = []
        try:
            if progress_callback:
                progress_callback("progress update")
            
            # comment removed (encoding issue)
            selectors = [
                'a[data-template-type="alsoSearch"]',
                '.related_keyword a',
                '.keyword a'
            ]
            
            for selector in selectors:
                try:
                    if not self.driver:
                        continue
                    elements = self.driver.find_elements(By.CSS_SELECTOR, selector)
                    for element in elements:
                        try:
                            text = element.text or element.get_attribute("textContent")
                            if text and text.strip():
                                text = text.strip()
                                if text not in keywords and len(text) > 1 and len(text) <= 50:
                                    keywords.append(text)
                        except:
                            continue
                except:
                    continue
            
            if progress_callback:
                progress_callback("progress update")
        
            return list(set(keywords))
        except Exception as e:
            if progress_callback:
                progress_callback("progress update")
            return []
            
    def extract_popular_topics(self, current_keyword, progress_callback=None):
        """Description"""
        keywords = []
        try:
            if progress_callback:
                progress_callback("progress update")
            
            # comment removed (encoding issue)
            selectors = [
                '.fds-comps-keyword-chip-text',
                '.keyword-chip .text',
                '.popular-keyword a'
            ]
            
            for selector in selectors:
                try:
                    if not self.driver:
                        continue
                    elements = self.driver.find_elements(By.CSS_SELECTOR, selector)
                    for element in elements:
                        try:
                            text = element.text or element.get_attribute("textContent")
                            if text and text.strip():
                                text = text.strip()
                                if text not in keywords and len(text) > 1 and len(text) <= 50:
                                    keywords.append(text)
                        except:
                            continue
                except:
                    continue
            
            if progress_callback:
                progress_callback("progress update")
            
            return list(set(keywords))
        except Exception as e:
            if progress_callback:
                progress_callback("progress update")
            return []

    def recursive_keyword_extraction(self, initial_keyword, progress_callback=None, extract_autocomplete=True):
        """Description"""
        if not self.driver:
            if progress_callback:
                progress_callback("progress update")
            return False
        
        self.base_keyword = initial_keyword
        self.all_related_keywords = []
        self.processed_autocomplete_keywords = set()
        
        if progress_callback:
            progress_callback("progress update")

        # comment removed (encoding issue)
        success = self._extract_all_keyword_types(
            initial_keyword, 
            parent_keyword=initial_keyword, 
            depth=0, 
            progress_callback=progress_callback
        )
        
        if not success:
            return False
            
        # comment removed (encoding issue)
        if extract_autocomplete:
            autocomplete_keywords = self.extract_autocomplete_keywords(initial_keyword, progress_callback)
            
            # comment removed (encoding issue)
            for keyword in autocomplete_keywords:
                self.all_related_keywords.append({
                    'depth': 0,
                    'parent_keyword': initial_keyword,
                    'current_keyword': initial_keyword,
                    'related_keyword': keyword,
                    'keyword_type': '자동완성',
                    'source_type': '자동완성검색어',
                    'extracted_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                })
                
            # comment removed (encoding issue)
            keywords_for_recursion = []
            for keyword in autocomplete_keywords:
                # comment removed (encoding issue)
                if keyword.lower().strip() != initial_keyword.lower().strip():
                    keywords_for_recursion.append(keyword)
            
            if progress_callback:
                progress_callback("progress update")
                if keywords_for_recursion:
                    progress_callback("progress update")
            
            # comment removed (encoding issue)
            if keywords_for_recursion:
                self._recursive_autocomplete_extraction(
                    keywords_for_recursion, 
                    initial_keyword, 
                    depth=1, 
                    progress_callback=progress_callback
                )
            else:
                if progress_callback:
                    progress_callback("progress update")
            
        if progress_callback:
            progress_callback(f"'{initial_keyword}' 키워드 추출 완료: 총 {len(self.all_related_keywords)}개")

        return True

    def _extract_all_keyword_types(self, current_keyword, parent_keyword, depth, progress_callback=None):
        """Description"""
        try:
            if not self.is_running:
                return False
            
            # comment removed (encoding issue)
            if not self.check_pause_status(progress_callback):
                return False
                
            # comment removed (encoding issue)
            if not self.search_keyword_mobile(current_keyword, progress_callback):
                return False
            
            # comment removed (encoding issue)
            if not self.check_pause_status(progress_callback):
                return False
            
            # comment removed (encoding issue)
            related_keywords = self.extract_related_keywords_new(current_keyword, progress_callback)
            
            # comment removed (encoding issue)
            if not self.check_pause_status(progress_callback):
                return False
                
            together_keywords = self.extract_together_keywords(current_keyword, progress_callback)
            
            # comment removed (encoding issue)
            if not self.check_pause_status(progress_callback):
                return False
            
            popular_keywords = self.extract_popular_topics(current_keyword, progress_callback)

            # comment removed (encoding issue)
            all_extracted = []
            
            for keyword in related_keywords:
                entry = {
                    'depth': depth,
                    'parent_keyword': parent_keyword,
                    'current_keyword': current_keyword,
                    'related_keyword': keyword,
                    'keyword_type': '연관검색어',
                    'source_type': '연관검색어',
                    'extracted_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                }
                self.all_related_keywords.append(entry)
                all_extracted.append(keyword)
            
            for keyword in together_keywords:
                entry = {
                    'depth': depth,
                    'parent_keyword': parent_keyword,
                    'current_keyword': current_keyword,
                    'related_keyword': keyword,
                    'keyword_type': '함께많이찾는',
                    'source_type': '함께많이찾는',
                    'extracted_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                }
                self.all_related_keywords.append(entry)
                all_extracted.append(keyword)
                        
            for keyword in popular_keywords:
                entry = {
                    'depth': depth,
                    'parent_keyword': parent_keyword,
                    'current_keyword': current_keyword,
                    'related_keyword': keyword,
                    'keyword_type': '인기주제',
                    'source_type': '인기주제',
                    'extracted_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                }
                self.all_related_keywords.append(entry)
                all_extracted.append(keyword)
            
            total_extracted = len(related_keywords) + len(together_keywords) + len(popular_keywords)
            if progress_callback:
                progress_callback(
                    f"'{current_keyword}' (depth={depth}): "
                    f"연관 {len(related_keywords)} + 함께 {len(together_keywords)} + 인기 {len(popular_keywords)} = 총 {total_extracted}개"
                )
            
            return True
            
        except Exception as e:
            if progress_callback:
                progress_callback("progress update")
            return False

    def _recursive_autocomplete_extraction(self, keywords_to_process, original_keyword, depth, progress_callback=None, max_depth=5):
        """Description"""
        
        if depth > max_depth:
            if progress_callback:
                progress_callback("progress update")
            return
                
        if not self.is_running:
            return
            
        for i, current_keyword in enumerate(keywords_to_process):
            
            if not self.is_running:
                break
                    
            # comment removed (encoding issue)
            if current_keyword.lower() in self.processed_autocomplete_keywords:
                if progress_callback:
                    progress_callback("progress update")
                continue
                    
            # comment removed (encoding issue)
            self.processed_autocomplete_keywords.add(current_keyword.lower())
            
            if progress_callback:
                progress_callback("progress update")
            
            # comment removed (encoding issue)
            self._extract_all_keyword_types(
                current_keyword, 
                parent_keyword=current_keyword, 
                depth=depth, 
                progress_callback=progress_callback
            )
            
            # comment removed (encoding issue)
            new_autocomplete_keywords = self.extract_autocomplete_keywords(current_keyword, progress_callback)
            
            # comment removed (encoding issue)
            for keyword in new_autocomplete_keywords:
                self.all_related_keywords.append({
                    'depth': depth,
                    'parent_keyword': current_keyword,
                    'current_keyword': current_keyword,
                    'related_keyword': keyword,
                    'keyword_type': '자동완성',
                    'source_type': '자동완성검색어',
                    'extracted_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                })
            
            # comment removed (encoding issue)
            if new_autocomplete_keywords:
                # comment removed (encoding issue)
                filtered_keywords = []
                for keyword in new_autocomplete_keywords:
                    # comment removed (encoding issue)
                    if keyword.lower().strip() == current_keyword.lower().strip():
                        continue
                    
                    # comment removed (encoding issue)
                    if keyword.lower() not in self.processed_autocomplete_keywords:
                        # comment removed (encoding issue)
                        if self.base_keyword.lower() in keyword.lower() or len(filtered_keywords) < 20:
                            filtered_keywords.append(keyword)
                
                if filtered_keywords:
                    if progress_callback:
                        progress_callback("progress update")
                        if len(filtered_keywords) <= 10:
                            progress_callback("progress update")
                        else:
                            progress_callback("progress update")
                    
                    # comment removed (encoding issue)
                    self._recursive_autocomplete_extraction(
                        filtered_keywords, 
                        original_keyword, 
                        depth + 1, 
                        progress_callback, 
                        max_depth
                    )
                else:
                    if progress_callback:
                        progress_callback("progress update")
            else:
                if progress_callback:
                    progress_callback("progress update")
            
        if progress_callback:
            progress_callback("progress update")

    def save_recursive_results_to_excel(self, save_path=None, progress_callback=None):
        """Save extraction results to file."""
        try:
            if not hasattr(self, 'all_related_keywords') or not self.all_related_keywords:
                if progress_callback:
                    progress_callback("progress update")
                return False
            
            if not save_path:
                if not self.save_dir:
                    self.save_dir = "keyword_results"
                    os.makedirs(self.save_dir, exist_ok=True)
                
                current_time = datetime.now().strftime('%Y%m%d_%H%M%S')
                base_keyword = getattr(self, 'base_keyword', 'keyword_extraction')
                save_path = os.path.join(self.save_dir, f"{base_keyword}_{current_time}.xlsx")
            
            # comment removed (encoding issue)
            df = pd.DataFrame({
                '추출된_키워드': [item['related_keyword'] for item in self.all_related_keywords]
            })

            # comment removed (encoding issue)
            df = df.drop_duplicates(subset=['추출된_키워드'], keep='first').reset_index(drop=True)

            # comment removed (encoding issue)
            try:
                df.to_excel(save_path, index=False, engine='openpyxl')
                
                if os.path.exists(save_path) and os.path.getsize(save_path) > 0:
                    if progress_callback:
                        progress_callback("progress update")
                        progress_callback(f"저장된 키워드 수: {len(df)}")
                    return True
                else:
                    raise Exception("저장 파일 생성 실패")
                
            except Exception as excel_error:
                # comment removed (encoding issue)
                csv_path = save_path.rsplit('.', 1)[0] + '.csv'
                df.to_csv(csv_path, index=False, encoding='utf-8-sig')
                    
                if progress_callback:
                    progress_callback("progress update")
                return True
            
        except Exception as e:
            if progress_callback:
                progress_callback("progress update")
            return False

    def close(self):
        """Description"""
        if self.driver:
            try:
                self.driver.quit()
            except:
                pass
            self.driver = None


class Settings:
    def __init__(self):
        self.settings_file = Path.home() / '.keyword_extractor_settings.json'
        self.settings = self.load_settings()

    def load_settings(self):
        default_settings = {
            "save_dir": "",
            "remember_dir": False,
            "remember_api_keys": False,
            "theme_mode": "light",
            "blog_count_mode": "monthly",
            "searchad_access_key": "",
            "searchad_secret_key": "",
            "searchad_customer_id": "",
            "naver_client_id": "",
            "naver_client_secret": "",
            "api_keys_file": str(Path.home() / ".naver_keyword_api_keys.json")
        }
        if self.settings_file.exists():
            try:
                with open(self.settings_file, 'r', encoding='utf-8') as f:
                    loaded = json.load(f)
                    if not isinstance(loaded, dict):
                        return default_settings
                    default_settings.update(loaded)
                    return default_settings
            except:
                return default_settings
        return default_settings

    def save_settings(self):
        with open(self.settings_file, 'w', encoding='utf-8') as f:
            json.dump(self.settings, f, ensure_ascii=False, indent=2)

    def get_save_dir(self):
        return self.settings.get("save_dir", "")

    def set_save_dir(self, directory, remember=None):
        if remember is not None:
            self.settings["remember_dir"] = remember
        if remember or (remember is None and self.settings.get("remember_dir", False)):
            self.settings["save_dir"] = directory
            self.save_settings()

    def should_remember_dir(self):
        return self.settings.get("remember_dir", False)

    def get_api_credentials(self):
        return {
            "searchad_access_key": self.settings.get("searchad_access_key", ""),
            "searchad_secret_key": self.settings.get("searchad_secret_key", ""),
            "searchad_customer_id": self.settings.get("searchad_customer_id", ""),
            "naver_client_id": self.settings.get("naver_client_id", ""),
            "naver_client_secret": self.settings.get("naver_client_secret", "")
        }

    def set_api_credentials(self, credentials, remember=False):
        self.settings["remember_api_keys"] = bool(remember)
        if remember:
            self.settings.update(credentials)
        else:
            self.settings["searchad_access_key"] = ""
            self.settings["searchad_secret_key"] = ""
            self.settings["searchad_customer_id"] = ""
            self.settings["naver_client_id"] = ""
            self.settings["naver_client_secret"] = ""
        self.save_settings()

    def should_remember_api_keys(self):
        return self.settings.get("remember_api_keys", False)

    def get_api_keys_file(self):
        default_path = str(Path.home() / ".naver_keyword_api_keys.json")
        return self.settings.get("api_keys_file", default_path)

    def set_api_keys_file(self, file_path):
        self.settings["api_keys_file"] = file_path
        self.save_settings()

    def get_theme_mode(self):
        mode = str(self.settings.get("theme_mode", "light")).strip().lower()
        return "dark" if mode == "dark" else "light"

    def set_theme_mode(self, mode):
        self.settings["theme_mode"] = "dark" if str(mode).strip().lower() == "dark" else "light"
        self.save_settings()

    def get_blog_count_mode(self):
        mode = str(self.settings.get("blog_count_mode", "monthly")).strip().lower()
        return "total" if mode == "total" else "monthly"

    def set_blog_count_mode(self, mode):
        self.settings["blog_count_mode"] = "total" if str(mode).strip().lower() == "total" else "monthly"
        self.save_settings()


class KeywordHunter:
    """Naver API based golden keyword analyzer."""
    SEARCHAD_BASE_URL = "https://api.searchad.naver.com"
    SEARCHAD_KEYWORD_URI = "/keywordstool"
    NAVER_BLOG_SEARCH_URL = "https://openapi.naver.com/v1/search/blog.json"
    NAVER_DATALAB_SEARCH_URL = "https://openapi.naver.com/v1/datalab/search"

    def __init__(self, access_key, secret_key, customer_id, client_id, client_secret, usage_callback=None):
        self.access_key = access_key.strip()
        self.secret_key = secret_key.strip()
        self.customer_id = customer_id.strip()
        self.client_id = client_id.strip()
        self.client_secret = client_secret.strip()
        self.searchad_cache = {}
        self.blog_count_cache = {}
        self.keyword_insight_cache = {}
        self.blog_count_driver = None
        self.usage_callback = usage_callback
        self._last_request_at = {"검색광고 API": 0.0, "네이버 검색 API": 0.0}

    def close(self):
        if self.blog_count_driver:
            try:
                self.blog_count_driver.quit()
            except Exception:
                pass
            self.blog_count_driver = None

    def _get_blog_count_driver(self):
        if self.blog_count_driver:
            return self.blog_count_driver
        # 월간 발행량 수집용 크롬은 항상 백그라운드(headless)로 실행
        self.blog_count_driver = create_chrome_driver(force_headless=True)
        return self.blog_count_driver

    def _extract_blog_post_key(self, href):
        href = str(href or "").strip()
        if not href:
            return None
        try:
            parsed = urllib.parse.urlparse(href)
            host = (parsed.netloc or "").lower()
            if "blog.naver.com" not in host:
                return None

            # 최신 검색결과에서 주로 쓰는 형태: /{blog_id}/{log_no}
            path = (parsed.path or "").strip("/")
            m = re.match(r"^([^/]+)/(\d+)$", path)
            if m:
                return f"{m.group(1).lower()}:{m.group(2)}"

            # 구형 링크 형태: /PostView.naver?blogId=...&logNo=...
            if path.lower() == "postview.naver":
                q = urllib.parse.parse_qs(parsed.query or "")
                blog_id = str((q.get("blogId") or [""])[0]).strip().lower()
                log_no = str((q.get("logNo") or [""])[0]).strip()
                if blog_id and log_no.isdigit():
                    return f"{blog_id}:{log_no}"
        except Exception:
            return None
        return None

    def _collect_visible_blog_post_keys(self, driver):
        post_keys = set()
        try:
            hrefs = driver.execute_script(
                """
                const selectors = [
                  "div[data-template-id='ugcItem'] a[href]",
                  ".lst_view a[href]",
                  "a.title_link[href]"
                ];
                const out = new Set();
                for (const sel of selectors) {
                  const nodes = document.querySelectorAll(sel);
                  for (const el of nodes) {
                    const href = (el && el.href) ? String(el.href).trim() : "";
                    if (href) out.add(href);
                  }
                }
                return Array.from(out);
                """
            ) or []
        except Exception:
            hrefs = []

        for href in hrefs:
            key = self._extract_blog_post_key(href)
            if key:
                post_keys.add(key)
        return post_keys

    def _count_visible_blog_cards(self, driver):
        return len(self._collect_visible_blog_post_keys(driver))

    def _count_monthly_blog_posts_by_scrolling(self, keyword):
        try:
            driver = self._get_blog_count_driver()
            if not driver:
                return None

            url = "https://search.naver.com/search.naver?ssc=tab.blog.all&query=&sm=tab_opt&nso=so%3Ar%2Cp%3A1m"
            driver.get(url)

            search_input = WebDriverWait(driver, 8).until(
                EC.element_to_be_clickable((By.ID, "nx_query"))
            )
            search_input.click()
            search_input.clear()
            search_input.send_keys(keyword)
            search_input.send_keys(Keys.ENTER)

            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
            time.sleep(0.8)

            prev_height = 0
            prev_y = -1
            no_progress_rounds = 0
            max_rounds = 2500
            safety_max_duration_sec = 300.0
            started_at = time.time()
            body = driver.find_element(By.TAG_NAME, "body")
            seen_post_keys = set()

            for i in range(max_rounds):
                # 하이브리드: JS 대점프 + PageDown 보조로 속도/안정성 균형
                step_info = driver.execute_script(
                    """
                    const se = document.scrollingElement || document.documentElement || document.body;
                    const h = Math.max(se.scrollHeight || 0, document.documentElement.scrollHeight || 0, document.body.scrollHeight || 0);
                    const y = Math.max(se.scrollTop || 0, window.pageYOffset || 0, document.documentElement.scrollTop || 0, document.body.scrollTop || 0);
                    const vh = (window.innerHeight || document.documentElement.clientHeight || 0);
                    const nearBottom = (y + vh) >= (h - Math.max(600, vh * 1.5));
                    const step = nearBottom ? Math.max(320, Math.floor(vh * 0.9)) : Math.max(1500, Math.floor(vh * 2.6));
                    window.scrollBy(0, step);
                    return [h, y, vh, Math.max(0, h - vh), nearBottom ? 1 : 0];
                    """
                ) or [0, 0, 0, 0, 0]

                # 키 이벤트 기반 lazy-load 트리거 보조
                for _ in range(36):
                    body.send_keys(Keys.PAGE_DOWN)
                if i % 6 == 0:
                    body.send_keys(Keys.END)
                time.sleep(0.008)

                metrics = driver.execute_script(
                    """
                    const se = document.scrollingElement || document.documentElement || document.body;
                    const h = Math.max(se.scrollHeight || 0, document.documentElement.scrollHeight || 0, document.body.scrollHeight || 0);
                    const y = Math.max(se.scrollTop || 0, window.pageYOffset || 0, document.documentElement.scrollTop || 0, document.body.scrollTop || 0);
                    const vh = (window.innerHeight || document.documentElement.clientHeight || 0);
                    return [h, y, vh, Math.max(0, h - vh)];
                    """
                ) or [0, 0, 0, 0]
                current_height = int(metrics[0] or 0)
                current_y = int(metrics[1] or 0)
                max_scroll_top = int(metrics[3] or 0)
                at_bottom = current_y >= max(0, max_scroll_top - 2)

                if i % 2 == 0 or at_bottom:
                    seen_post_keys.update(self._collect_visible_blog_post_keys(driver))

                if current_height == prev_height and current_y == prev_y and not at_bottom:
                    no_progress_rounds += 1
                else:
                    no_progress_rounds = 0

                prev_height = current_height
                prev_y = current_y

                # 하단으로 보이면 강제 바닥 점프 + PageDown 재확인 후 종료 판단
                if at_bottom:
                    probe_changed = False
                    probe_height = current_height
                    probe_y = current_y
                    for _ in range(4):
                        driver.execute_script(
                            """
                            const se = document.scrollingElement || document.documentElement || document.body;
                            se.scrollTop = se.scrollHeight;
                            window.scrollTo(0, se.scrollHeight);
                            """
                        )
                        for _ in range(80):
                            body.send_keys(Keys.PAGE_DOWN)
                        time.sleep(0.012)
                        probe_metrics = driver.execute_script(
                            """
                            const se = document.scrollingElement || document.documentElement || document.body;
                            const h = Math.max(se.scrollHeight || 0, document.documentElement.scrollHeight || 0, document.body.scrollHeight || 0);
                            const y = Math.max(se.scrollTop || 0, window.pageYOffset || 0, document.documentElement.scrollTop || 0, document.body.scrollTop || 0);
                            const vh = (window.innerHeight || document.documentElement.clientHeight || 0);
                            return [h, y, Math.max(0, h - vh)];
                            """
                        ) or [0, 0, 0]
                        probe_h = int(probe_metrics[0] or 0)
                        probe_y_new = int(probe_metrics[1] or 0)
                        probe_max = int(probe_metrics[2] or 0)
                        seen_post_keys.update(self._collect_visible_blog_post_keys(driver))
                        if probe_h != probe_height or probe_y_new != probe_y:
                            probe_changed = True
                            break
                        if probe_y_new < max(0, probe_max - 2):
                            probe_changed = True
                            break
                    if not probe_changed:
                        break

                # 스크롤이 남았는데 정체되면 포커스 재정렬 후 크게 밀어준다
                if no_progress_rounds >= 4:
                    try:
                        body = driver.find_element(By.TAG_NAME, "body")
                        driver.execute_script(
                            """
                            const se = document.scrollingElement || document.documentElement || document.body;
                            if (document && document.body) { document.body.focus(); }
                            window.focus();
                            se.focus && se.focus();
                            window.scrollBy(0, (window.innerHeight || 800) * 3);
                            """
                        )
                        body.click()
                        for _ in range(140):
                            body.send_keys(Keys.PAGE_DOWN)
                    except Exception:
                        pass
                    no_progress_rounds = 0

                # 비정상 무한루프 방지용 안전장치
                if (time.time() - started_at) >= safety_max_duration_sec:
                    break

            final_count = len(seen_post_keys) if seen_post_keys else self._count_visible_blog_cards(driver)
            return int(final_count)
        except Exception:
            return None

    def _normalize_keyword(self, text):
        value = str(text or "")
        value = unquote(value).replace("+", " ")
        value = re.sub(r"\s+", " ", value).strip()
        return value

    def _keyword_key(self, text):
        return self._normalize_keyword(text).replace(" ", "").lower()

    def _throttle(self, api_name):
        min_interval_map = {
            # API는 너무 빠른 연속 호출을 피해서 429/탐지 리스크를 줄인다.
            "검색광고 API": 0.60,
            "네이버 검색 API": 0.75,
        }
        min_interval = min_interval_map.get(api_name, 0.5)
        now = time.time()
        elapsed = now - self._last_request_at.get(api_name, 0.0)
        if elapsed < min_interval:
            time.sleep(min_interval - elapsed)
        self._last_request_at[api_name] = time.time()

    def _request_with_retry(self, url, headers, params, timeout, api_name):
        max_retries = 5
        backoff = 1.2
        last_response = None

        for attempt in range(1, max_retries + 1):
            self._throttle(api_name)
            if self.usage_callback:
                try:
                    self.usage_callback(1)
                except Exception:
                    pass
            response = requests.get(url, headers=headers, params=params, timeout=timeout)
            last_response = response
            if response.status_code == 200:
                return response

            retriable = response.status_code == 429 or 500 <= response.status_code < 600
            if not retriable or attempt == max_retries:
                break

            retry_after = response.headers.get("Retry-After")
            if retry_after and retry_after.isdigit():
                wait_seconds = float(retry_after)
            else:
                if response.status_code == 429:
                    wait_seconds = (backoff * attempt) + (2.0 * attempt)
                else:
                    wait_seconds = backoff * attempt
            time.sleep(min(wait_seconds, 18.0))

        if last_response is None:
            raise ValueError(f"{api_name} 요청 실패: 응답 없음")
        raise ValueError(
            f"{api_name} 요청 실패 ({last_response.status_code}): {last_response.text[:220]}"
        )

    def get_signature(self, method, uri, timestamp):
        message = f"{timestamp}.{method}.{uri}"
        digest = hmac.new(
            self.secret_key.encode("utf-8"),
            message.encode("utf-8"),
            hashlib.sha256
        ).digest()
        return base64.b64encode(digest).decode("utf-8")

    def _parse_count(self, value):
        if value is None:
            return 0
        if isinstance(value, int):
            return value
        value_text = str(value).replace(",", "").strip()
        if value_text.startswith("<"):
            digits = re.sub(r"[^0-9]", "", value_text)
            return int(digits) if digits else 0
        digits = re.sub(r"[^0-9]", "", value_text)
        return int(digits) if digits else 0

    def get_searchad_related_keywords(self, keyword):
        keyword = self._normalize_keyword(keyword)
        hint_keyword = self._keyword_key(keyword)
        if not hint_keyword:
            return []
        cache_key = hint_keyword
        if cache_key in self.searchad_cache:
            return [dict(row) for row in self.searchad_cache[cache_key]]

        timestamp = str(int(time.time() * 1000))
        method = "GET"
        uri = self.SEARCHAD_KEYWORD_URI
        signature = self.get_signature(method, uri, timestamp)
        headers = {
            "X-Timestamp": timestamp,
            "X-API-KEY": self.access_key,
            "X-Customer": self.customer_id,
            "X-Signature": signature
        }
        params = {
            "hintKeywords": hint_keyword,
            "showDetail": 1
        }

        response = self._request_with_retry(
            f"{self.SEARCHAD_BASE_URL}{uri}",
            headers=headers,
            params=params,
            timeout=15,
            api_name="검색광고 API"
        )

        payload = response.json()
        keyword_list = payload.get("keywordList", [])
        results = []

        for item in keyword_list:
            rel_keyword = self._normalize_keyword(item.get("relKeyword", ""))
            if not rel_keyword:
                continue
            pc_count = self._parse_count(item.get("monthlyPcQcCnt"))
            mobile_count = self._parse_count(item.get("monthlyMobileQcCnt"))
            search_volume = pc_count + mobile_count
            results.append({
                "keyword": rel_keyword,
                "monthly_pc_search": pc_count,
                "monthly_mobile_search": mobile_count,
                "monthly_total_search": search_volume
            })

        self.searchad_cache[cache_key] = [dict(row) for row in results]
        return results

    def _get_total_blog_document_count(self, keyword, headers):
        keyword = self._normalize_keyword(keyword)
        cache_key = f"total:{keyword.strip().lower()}"
        if cache_key in self.blog_count_cache:
            return int(self.blog_count_cache[cache_key])

        # 1순위: 네이버 OpenAPI total (기간 제한 없음)
        try:
            params = {
                "query": keyword,
                "display": 1,
                "start": 1,
                "sort": "sim"
            }
            response = self._request_with_retry(
                self.NAVER_BLOG_SEARCH_URL,
                headers=headers,
                params=params,
                timeout=10,
                api_name="네이버 검색 API"
            )
            total = int(response.json().get("total", 0))
            self.blog_count_cache[cache_key] = total
            return total
        except Exception:
            pass

        # 2순위 fallback: 네이버 웹검색 블로그 total 파싱 (기간 제한 없음)
        try:
            web_url = "https://search.naver.com/search.naver"
            web_headers = {"User-Agent": "Mozilla/5.0"}
            web_params = {
                "where": "blog",
                "query": keyword,
                "sm": "tab_opt",
                "dup_remove": "1",
            }
            self._throttle("네이버 검색 API")
            web_resp = requests.get(web_url, params=web_params, headers=web_headers, timeout=10)
            total = self._extract_naver_total_count(web_resp.text)
            if total > 0:
                self.blog_count_cache[cache_key] = total
                return total
        except Exception:
            pass

        self.blog_count_cache[cache_key] = 0
        return 0

    def get_blog_document_count(self, keyword, count_mode="monthly"):
        keyword = self._normalize_keyword(keyword)
        mode = "total" if str(count_mode).strip().lower() == "total" else "monthly"
        cache_key = f"{mode}:{keyword.strip().lower()}"
        if cache_key in self.blog_count_cache:
            return int(self.blog_count_cache[cache_key])

        headers = {
            "X-Naver-Client-Id": self.client_id,
            "X-Naver-Client-Secret": self.client_secret
        }
        if mode == "total":
            return self._get_total_blog_document_count(keyword, headers)

        # 1순위: 블로그 탭(최근 1개월) 접속 후 실제 스크롤 로딩된 카드 수를 카운트
        scrolled_count = self._count_monthly_blog_posts_by_scrolling(keyword)
        if scrolled_count is not None:
            self.blog_count_cache[cache_key] = int(scrolled_count)
            return int(scrolled_count)

        # 2순위: 최근 1개월 + 완전일치("키워드") 기준 블로그 문서 수
        # 완전일치 결과가 없으면 같은 1개월 범위의 일반 쿼리 결과를 사용한다.
        web_url = "https://search.naver.com/search.naver"
        web_headers = {"User-Agent": "Mozilla/5.0"}
        month_params_common = {
            "where": "blog",
            "sm": "tab_opt",
            "dup_remove": "1",
            "nso": "so:dd,p:1m,a:all",
        }
        try:
            exact_query = f"\"{keyword}\""
            exact_params = dict(month_params_common, query=exact_query)
            self._throttle("네이버 검색 API")
            exact_resp = requests.get(web_url, params=exact_params, headers=web_headers, timeout=10)
            exact_total = self._extract_naver_total_count(exact_resp.text)
            if exact_total > 0:
                self.blog_count_cache[cache_key] = exact_total
                return exact_total
        except Exception:
            pass

        try:
            broad_params = dict(month_params_common, query=keyword)
            self._throttle("네이버 검색 API")
            broad_resp = requests.get(web_url, params=broad_params, headers=web_headers, timeout=10)
            broad_total = self._extract_naver_total_count(broad_resp.text)
            if broad_total > 0:
                self.blog_count_cache[cache_key] = broad_total
                return broad_total
        except Exception:
            pass

        # 3순위: 네이버 OpenAPI total (기간 제한 없음)
        try:
            params = {
                "query": keyword,
                "display": 1,
                "start": 1,
                "sort": "sim"
            }
            response = self._request_with_retry(
                self.NAVER_BLOG_SEARCH_URL,
                headers=headers,
                params=params,
                timeout=10,
                api_name="네이버 검색 API"
            )
            total = int(response.json().get("total", 0))
            self.blog_count_cache[cache_key] = total
            return total
        except Exception:
            pass

        # 4순위 fallback: 네이버 웹검색 블로그 total 파싱 (기간 제한 없음)
        try:
            web_params = {
                "where": "blog",
                "query": keyword,
                "sm": "tab_opt",
                "dup_remove": "1",
            }
            self._throttle("네이버 검색 API")
            web_resp = requests.get(web_url, params=web_params, headers=web_headers, timeout=10)
            total = self._extract_naver_total_count(web_resp.text)
            if total > 0:
                self.blog_count_cache[cache_key] = total
                return total
        except Exception:
            pass

        self.blog_count_cache[cache_key] = 0
        return 0

    def _extract_naver_total_count(self, html):
        patterns = [
            r"/\s*([\d,]+)\s*건",
            r"약\s*([\d,]+)\s*건",
            r"([\d,]+)\s*건의\s*검색결과",
            r'"total"\s*:\s*"?([\d,]+)"?',
            r'"totalCount"\s*:\s*"?([\d,]+)"?',
        ]
        for pattern in patterns:
            m = re.search(pattern, html)
            if not m:
                continue
            try:
                return int(str(m.group(1)).replace(",", ""))
            except Exception:
                continue
        return 0

    def _request_post_with_retry(self, url, headers, payload, timeout, api_name):
        max_retries = 5
        last_response = None
        for attempt in range(1, max_retries + 1):
            self._throttle(api_name)
            if self.usage_callback:
                try:
                    self.usage_callback(1)
                except Exception:
                    pass
            response = requests.post(url, headers=headers, json=payload, timeout=timeout)
            last_response = response
            if response.status_code == 200:
                return response
            retriable = response.status_code == 429 or 500 <= response.status_code < 600
            if not retriable or attempt == max_retries:
                break
            time.sleep(min(2.0 * attempt, 12.0))
        if last_response is None:
            raise ValueError(f"{api_name} 요청 실패: 응답 없음")
        raise ValueError(f"{api_name} 요청 실패 ({last_response.status_code}): {last_response.text[:200]}")

    def get_keyword_insight(self, keyword):
        key = keyword.strip()
        cache_key = key.lower()
        if cache_key in self.keyword_insight_cache:
            return self.keyword_insight_cache[cache_key]

        headers = {
            "X-Naver-Client-Id": self.client_id,
            "X-Naver-Client-Secret": self.client_secret
        }
        today = datetime.now().date()
        start_month = (today.replace(day=1) - timedelta(days=330)).replace(day=1)
        monthly_payload = {
            "startDate": start_month.strftime("%Y-%m-%d"),
            "endDate": today.strftime("%Y-%m-%d"),
            "timeUnit": "month",
            "keywordGroups": [{"groupName": key, "keywords": [key]}]
        }
        monthly_resp = self._request_post_with_retry(
            self.NAVER_DATALAB_SEARCH_URL, headers, monthly_payload, 12, "네이버 검색 API"
        )
        monthly_data = monthly_resp.json().get("results", [{}])[0].get("data", [])

        trend = []
        month_ratio = []
        month_sum = sum(float(x.get("ratio", 0.0)) for x in monthly_data) or 1.0
        for row in monthly_data:
            period = str(row.get("period", ""))[:7]
            value = float(row.get("ratio", 0.0))
            trend.append({"label": period, "value": value})
            month_ratio.append({"label": period, "value": (value / month_sum) * 100.0})

        daily_payload = {
            "startDate": (today - timedelta(days=89)).strftime("%Y-%m-%d"),
            "endDate": today.strftime("%Y-%m-%d"),
            "timeUnit": "date",
            "keywordGroups": [{"groupName": key, "keywords": [key]}]
        }
        daily_resp = self._request_post_with_retry(
            self.NAVER_DATALAB_SEARCH_URL, headers, daily_payload, 12, "네이버 검색 API"
        )
        daily_data = daily_resp.json().get("results", [{}])[0].get("data", [])
        names = ["월", "화", "수", "목", "금", "토", "일"]
        wd_sum_map = {name: 0.0 for name in names}
        for row in daily_data:
            p = str(row.get("period", ""))
            try:
                wd = datetime.strptime(p, "%Y-%m-%d").weekday()
                wd_sum_map[names[wd]] += float(row.get("ratio", 0.0))
            except Exception:
                continue
        wd_total = sum(wd_sum_map.values()) or 1.0
        weekday_ratio = [{"label": k, "value": (v / wd_total) * 100.0} for k, v in wd_sum_map.items()]

        age_groups = [
            ("10대", ["2", "3"]),
            ("20대", ["4", "5"]),
            ("30대", ["6", "7"]),
            ("40대", ["8", "9"]),
            ("50대 이상", ["10", "11"]),
        ]
        age_raw = {}
        for label, ages in age_groups:
            age_payload = {
                "startDate": (today - timedelta(days=89)).strftime("%Y-%m-%d"),
                "endDate": today.strftime("%Y-%m-%d"),
                "timeUnit": "date",
                "ages": ages,
                "keywordGroups": [{"groupName": key, "keywords": [key]}]
            }
            try:
                age_resp = self._request_post_with_retry(
                    self.NAVER_DATALAB_SEARCH_URL, headers, age_payload, 12, "네이버 검색 API"
                )
                age_data = age_resp.json().get("results", [{}])[0].get("data", [])
                age_raw[label] = sum(float(x.get("ratio", 0.0)) for x in age_data)
            except Exception:
                age_raw[label] = 0.0
        age_total = sum(age_raw.values()) or 1.0
        age_ratio = [{"label": k, "value": (v / age_total) * 100.0} for k, v in age_raw.items()]

        insight = {
            "keyword": key,
            "trend": trend,
            "month_ratio": month_ratio,
            "weekday_ratio": weekday_ratio,
            "age_ratio": age_ratio
        }
        self.keyword_insight_cache[cache_key] = insight
        return insight

    def calculate_content_saturation_index(self, monthly_search, blog_docs):
        # 콘텐츠 포화지수(%) = (콘텐츠양 / 월 검색량) * 100
        if monthly_search <= 0:
            return 9999.0 if blog_docs > 0 else 0.0
        return (blog_docs / monthly_search) * 100.0

    def _score_keyword_rows(self, keyword_rows, limit, offset=0, progress_callback=None, blog_count_mode="monthly"):
        scored = []
        start = max(0, int(offset))
        end = start + max(1, int(limit))
        batch_rows = keyword_rows[start:end]
        if progress_callback and start > 0:
            progress_callback(f"배치 오프셋 {start}부터 {len(batch_rows)}개를 분석합니다.")
        for idx, row in enumerate(batch_rows, start=1):
            if progress_callback:
                progress_callback(f"[EVAL {idx}/{len(batch_rows)}] {row['keyword']} 지표 계산 시작")
            blog_docs = self.get_blog_document_count(row["keyword"], count_mode=blog_count_mode)
            saturation = self.calculate_content_saturation_index(
                row["monthly_total_search"], blog_docs
            )
            result = {
                "keyword": row["keyword"],
                "monthly_pc_search": row["monthly_pc_search"],
                "monthly_mobile_search": row["monthly_mobile_search"],
                "monthly_total_search": row["monthly_total_search"],
                "blog_document_count": blog_docs,
                "content_saturation_index": saturation,
                "section_position": "확인 필요",
            }
            scored.append(result)
            if progress_callback:
                progress_callback(f"[EVAL {idx}/{len(batch_rows)}] {row['keyword']} 계산 완료")
        return scored

    def _tokenize_text(self, text):
        return [t for t in re.split(r"[\s/,_\-]+", str(text).lower()) if t]

    def _seed_ngrams(self, seed_text, n=3):
        seed = self._normalize_keyword(seed_text).replace(" ", "").lower()
        if len(seed) < n:
            return {seed} if seed else set()
        return {seed[i:i+n] for i in range(0, len(seed) - n + 1)}

    def _build_seed_variants(self, seed_text):
        seed = self._normalize_keyword(seed_text)
        compact = seed.replace(" ", "")
        variants = []
        for item in [seed, compact]:
            if item and item not in variants:
                variants.append(item)
        if len(compact) >= 5:
            for cut in [1, 2, 3]:
                if len(compact) - cut >= 4:
                    v = compact[:-cut]
                    if v not in variants:
                        variants.append(v)
        return variants[:5]

    def _is_related_match(self, seed_text, keyword_text):
        seed = self._normalize_keyword(seed_text).replace(" ", "").lower()
        key = self._normalize_keyword(keyword_text).replace(" ", "").lower()
        if not seed or not key:
            return False
        return seed in key

    def _fetch_autocomplete_candidates(self, seed_text):
        seed = self._normalize_keyword(seed_text)
        if not seed:
            return []
        url = "https://ac.search.naver.com/nx/ac"
        query_variants = [seed]
        compact = seed.replace(" ", "")
        if compact and compact != seed:
            query_variants.append(compact)
        if not seed.endswith(" "):
            query_variants.append(seed + " ")
        headers = {"User-Agent": "Mozilla/5.0"}
        terms = []
        for q in query_variants:
            params = {
                "q": q,
                "con": 0,
                "frm": "nv",
                "ans": 2,
                "r_format": "json",
                "r_enc": "UTF-8",
                "r_unicode": 0,
                "t_koreng": 1,
                "run": 2,
                "rev": 4,
            }
            try:
                self._throttle("네이버 검색 API")
                resp = requests.get(url, params=params, headers=headers, timeout=8)
                payload = resp.json()
                items = payload.get("items", [])
                if items and isinstance(items[0], list):
                    for row in items[0]:
                        if isinstance(row, (list, tuple)) and row:
                            t = self._normalize_keyword(row[0])
                            if t:
                                terms.append(t)
            except Exception:
                continue

        dedup = []
        seen = set()
        for t in terms:
            k = t.lower()
            if k in seen:
                continue
            seen.add(k)
            dedup.append(t)
        return dedup

    def _fetch_serp_candidates(self, seed_text):
        seed = self._normalize_keyword(seed_text)
        if not seed:
            return []
        if not BEAUTIFULSOUP_AVAILABLE:
            return []

        generic_stop = {
            "뉴스", "블로그", "카페", "이미지", "동영상", "쇼핑", "지도",
            "지식in", "지식인", "어학사전", "웹사이트", "인플루언서",
        }
        urls = [
            f"https://search.naver.com/search.naver?where=nexearch&query={quote(seed)}",
            f"https://m.search.naver.com/search.naver?query={quote(seed)}",
        ]
        headers = {"User-Agent": "Mozilla/5.0"}
        candidates = []

        for url in urls:
            try:
                self._throttle("네이버 검색 API")
                resp = requests.get(url, headers=headers, timeout=10)
                html = resp.text
                soup = BeautifulSoup(html, "html.parser")

                # data-query 기반
                for node in soup.select("[data-query]"):
                    q = self._normalize_keyword(node.get("data-query", ""))
                    if q:
                        candidates.append(q)

                # href query 파라미터 기반
                for a in soup.select("a[href*='query=']"):
                    href = str(a.get("href", ""))
                    try:
                        parsed = urllib.parse.urlparse(href)
                        qv = urllib.parse.parse_qs(parsed.query).get("query", [])
                        if qv:
                            q = self._normalize_keyword(qv[0])
                            if q:
                                candidates.append(q)
                    except Exception:
                        continue
            except Exception:
                continue

        dedup = []
        seen = set()
        for c in candidates:
            key = c.replace(" ", "").lower()
            if not key or key in seen:
                continue
            if c in generic_stop:
                continue
            seen.add(key)
            dedup.append(c)
        return dedup

    def _sanitize_candidate_keyword(self, text):
        t = self._normalize_keyword(text)
        if not t:
            return ""
        t = re.sub(r"https?://\S+", " ", t, flags=re.IGNORECASE)
        t = re.sub(r"\bwww\.\S+", " ", t, flags=re.IGNORECASE)
        t = re.sub(r"\bsite\s*:\s*\S+", " ", t, flags=re.IGNORECASE)
        t = re.sub(r"[\"'|(){}\[\]]", " ", t)
        t = re.sub(r"[^0-9A-Za-z가-힣\s]", " ", t)
        t = re.sub(r"\s+", " ", t).strip()
        if not t or len(t) < 2:
            return ""
        if "site" in t.lower():
            return ""
        return t

    def _strip_html(self, text):
        return re.sub(r"<[^>]+>", " ", str(text or ""))

    def _fetch_blog_title_candidates(self, seed_text):
        seed = self._normalize_keyword(seed_text)
        if not seed:
            return []
        headers = {
            "X-Naver-Client-Id": self.client_id,
            "X-Naver-Client-Secret": self.client_secret
        }
        terms = []
        for start in [1, 101, 201]:
            params = {"query": seed, "display": 100, "start": start, "sort": "date"}
            try:
                resp = self._request_with_retry(
                    self.NAVER_BLOG_SEARCH_URL,
                    headers=headers,
                    params=params,
                    timeout=10,
                    api_name="네이버 검색 API"
                )
                items = resp.json().get("items", []) or []
            except Exception:
                items = []
            if not items:
                break
            for item in items:
                title = self._normalize_keyword(self._strip_html(item.get("title", "")))
                if not title:
                    continue
                tnorm = title.replace(" ", "").lower()
                snorm = seed.replace(" ", "").lower()
                if snorm not in tnorm:
                    continue
                words = [w for w in re.split(r"\s+", title) if w]
                # title에서 seed가 포함된 구간 주변 단어를 후보로 생성
                for i, w in enumerate(words):
                    wn = w.replace(" ", "").lower()
                    if snorm not in wn and snorm not in "".join(words[max(0, i-1):i+2]).replace(" ", "").lower():
                        continue
                    for span in [2, 3, 4]:
                        s = max(0, i - 1)
                        e = min(len(words), s + span)
                        cand = self._normalize_keyword(" ".join(words[s:e]))
                        if snorm in cand.replace(" ", "").lower():
                            terms.append(cand)

        dedup = []
        seen = set()
        for t in terms:
            k = t.replace(" ", "").lower()
            if not k or k in seen:
                continue
            if len(t) > 40 or len(t) < 2:
                continue
            seen.add(k)
            dedup.append(t)
        return dedup

    def _resolve_term_volume_row(self, term):
        term_norm = self._normalize_keyword(term)
        if not term_norm:
            return None
        rows = self.get_searchad_related_keywords(term_norm)
        if not rows:
            return None
        term_key = term_norm.replace(" ", "").lower()
        best = None

        def grams2(s):
            s = s.replace(" ", "").lower()
            if len(s) < 2:
                return {s} if s else set()
            return {s[i:i+2] for i in range(len(s) - 1)}

        term_grams = grams2(term_norm)
        for row in rows:
            kw = self._normalize_keyword(row.get("keyword", ""))
            kw_key = kw.replace(" ", "").lower()
            if kw_key == term_key:
                return {
                    "keyword": term_norm,
                    "monthly_pc_search": int(row.get("monthly_pc_search", 0)),
                    "monthly_mobile_search": int(row.get("monthly_mobile_search", 0)),
                    "monthly_total_search": int(row.get("monthly_total_search", 0)),
                }
            if term_key in kw_key:
                cand = {
                    "keyword": term_norm,
                    "monthly_pc_search": int(row.get("monthly_pc_search", 0)),
                    "monthly_mobile_search": int(row.get("monthly_mobile_search", 0)),
                    "monthly_total_search": int(row.get("monthly_total_search", 0)),
                }
                if best is None or cand["monthly_total_search"] > best["monthly_total_search"]:
                    best = cand
                continue

            # 정확일치가 없어도 유사도가 높은 검색광고 행을 후보 키워드에 매핑
            kw_grams = grams2(kw)
            overlap = len(term_grams & kw_grams) if term_grams and kw_grams else 0
            if overlap <= 0:
                continue
            cand = {
                "keyword": term_norm,
                "monthly_pc_search": int(row.get("monthly_pc_search", 0)),
                "monthly_mobile_search": int(row.get("monthly_mobile_search", 0)),
                "monthly_total_search": int(row.get("monthly_total_search", 0)),
                "_overlap": overlap
            }
            if best is None:
                best = cand
            else:
                prev_overlap = int(best.get("_overlap", 0))
                if overlap > prev_overlap or (
                    overlap == prev_overlap and cand["monthly_total_search"] > best["monthly_total_search"]
                ):
                    best = cand
        if best and "_overlap" in best:
            best.pop("_overlap", None)
        return best

    def _collect_naver_candidates(self, seed_text, progress_callback=None, required_key=None):
        seed = self._normalize_keyword(seed_text)
        if not seed:
            return []
        seed_key = self._keyword_key(required_key if required_key is not None else seed)
        if progress_callback:
            progress_callback("1단계 네이버 검색결과/자동완성 키워드 수집 중...")

        # 네이버 검색결과 연관검색어 + 자동완성검색어만 사용
        ac_terms = self._fetch_autocomplete_candidates(seed)
        extractor_related = []
        serp_related = []
        try:
            scraper = NaverMobileSearchScraper(driver=None)
            extractor_related = scraper.extract_related_keywords(seed, progress_callback=None) or []
        except Exception:
            extractor_related = []
        try:
            # 검색결과 페이지에서 추출한 관련 쿼리(related source)
            serp_related = self._fetch_serp_candidates(seed) or []
        except Exception:
            serp_related = []

        # 수집이 너무 적으면 같은 소스(연관/자동완성)를 Selenium으로 재시도
        selenium_related = []
        selenium_ac = []
        if len(extractor_related) + len(serp_related) + len(ac_terms) < 3:
            try:
                driver = create_chrome_driver()
                if driver:
                    sc = NaverMobileSearchScraper(driver)
                    if sc.search_keyword_mobile(seed, progress_callback=None):
                        selenium_related = sc.extract_related_keywords_new(seed, progress_callback=None) or []
                    selenium_ac = sc.extract_autocomplete_keywords(seed, progress_callback=None) or []
                    try:
                        driver.quit()
                    except Exception:
                        pass
            except Exception:
                pass

        related_all = [
            self._normalize_keyword(x) for x in (extractor_related + serp_related + selenium_related)
            if self._normalize_keyword(x)
        ]
        ac_all = [self._normalize_keyword(x) for x in (ac_terms + selenium_ac) if self._normalize_keyword(x)]

        candidates = []
        seen = {}
        for term in related_all:
            t = self._sanitize_candidate_keyword(term)
            if not t:
                continue
            key = t.replace(" ", "").lower()
            if seed_key and seed_key not in key:
                continue
            if key in seen:
                if seen[key] != "연관+자동완성":
                    seen[key] = "연관"
                continue
            seen[key] = "연관"
            candidates.append({"keyword": t, "source": "연관"})

        for term in ac_all:
            t = self._sanitize_candidate_keyword(term)
            if not t:
                continue
            key = t.replace(" ", "").lower()
            if seed_key and seed_key not in key:
                continue
            if key in seen:
                for c in candidates:
                    if c["keyword"].replace(" ", "").lower() == key:
                        c["source"] = "연관+자동완성" if c["source"] == "연관" else "자동완성"
                        break
                continue
            seen[key] = "자동완성"
            candidates.append({"keyword": t, "source": "자동완성"})

        if not candidates:
            # 최후 fallback: 입력 키워드라도 분석 대상으로 유지
            candidates.append({"keyword": seed, "source": "자동완성"})

        if progress_callback:
            progress_callback(
                "키워드 수집 완료: "
                f"검색결과 {len(related_all)}개, 자동완성 {len(ac_all)}개, "
                f"요청기반 연관 {len(extractor_related)}개"
            )
        return candidates

    def _category_relevance(self, seed, keyword):
        seed_tokens = self._tokenize_text(seed)
        key_tokens = self._tokenize_text(keyword)
        if not seed_tokens or not key_tokens:
            return 0
        score = 0
        merged_keyword = "".join(key_tokens)
        for token in seed_tokens:
            if token in merged_keyword:
                score += 2
            if token in key_tokens:
                score += 1
        return score

    def analyze_related_keywords_with_content(self, seed_keyword, limit=30, offset=0, progress_callback=None, expand_seeds=None, blog_count_mode="monthly"):
        seed = self._normalize_keyword(seed_keyword)
        if not seed:
            raise ValueError("키워드를 입력해 주세요.")

        if progress_callback:
            progress_callback(f"'{seed}' 분석을 시작합니다.")

        # 1~4) 네이버 검색결과/자동완성 후보 수집
        required_key = self._keyword_key(seed)
        web_terms = self._collect_naver_candidates(seed, progress_callback=progress_callback, required_key=required_key)

        # "더 많은 연관 키워드 보기"에서는 초기 결과(검색량 내림차순) 씨드로
        # 연관의 연관을 재귀 확장한다. 재귀는 월 검색량이 1,000 초과인 키워드만 계속 진행한다.
        expand_pool = []
        if expand_seeds:
            seen_expand = set()
            for raw in expand_seeds:
                t = self._normalize_keyword(raw)
                if not t:
                    continue
                k = self._keyword_key(t)
                if not k or k in seen_expand:
                    continue
                seen_expand.add(k)
                expand_pool.append(t)

        if expand_pool:
            volume_threshold = 1000
            max_nodes_per_root = 120
            max_children_per_node = 80
            seen_recursive_terms = {
                self._keyword_key(item.get("keyword", "")) for item in web_terms if item.get("keyword")
            }
            seen_recursive_terms.discard("")

            if progress_callback:
                progress_callback(
                    f"1단계 재귀 확장 수집 시작... ({len(expand_pool)}개 루트, 임계값 {volume_threshold:,})"
                )

            for root_idx, root_seed in enumerate(expand_pool, start=1):
                root_key = self._keyword_key(root_seed)
                if not root_key:
                    continue
                queue = deque([(root_seed, 0)])
                visited_chain = {root_key}
                expanded_nodes = 0

                if progress_callback:
                    progress_callback(f"[루트 {root_idx}/{len(expand_pool)}] '{root_seed}' 재귀 확장")

                while queue and expanded_nodes < max_nodes_per_root:
                    current_seed, depth = queue.popleft()
                    expanded_nodes += 1

                    try:
                        rel_rows = self.get_searchad_related_keywords(current_seed) or []
                    except Exception:
                        rel_rows = []

                    if not rel_rows:
                        continue

                    for rel in rel_rows[:max_children_per_node]:
                        term = self._normalize_keyword(rel.get("keyword", ""))
                        if not term:
                            continue
                        tkey = self._keyword_key(term)
                        if not tkey or tkey == required_key:
                            continue

                        if tkey not in seen_recursive_terms:
                            web_terms.append({"keyword": term, "source": "요청기반 연관(재귀)"})
                            seen_recursive_terms.add(tkey)

                        monthly_total = int(rel.get("monthly_total_search", 0))
                        if monthly_total > volume_threshold and tkey not in visited_chain:
                            visited_chain.add(tkey)
                            queue.append((term, depth + 1))

                    time.sleep(0.02)

        if web_terms:
            dedup_terms = []
            seen_terms = set()
            for item in web_terms:
                k = self._keyword_key(item.get("keyword", ""))
                if not k or k in seen_terms:
                    continue
                seen_terms.add(k)
                dedup_terms.append(item)
            web_terms = dedup_terms
        if not web_terms:
            return []

        # 5) 수집한 키워드를 검색광고 API로 검색량 분석
        merged = {}
        total_terms = len(web_terms)
        if progress_callback:
            progress_callback(f"2단계 검색광고 검색량 분석 중... ({total_terms}개)")

        # 속도 개선: seed 기준 검색광고 1회 조회 결과를 우선 매핑
        base_volume_map = {}
        try:
            base_rows = self.get_searchad_related_keywords(seed)
            for r in base_rows:
                k = self._keyword_key(r.get("keyword", ""))
                if not k:
                    continue
                base_volume_map[k] = {
                    "keyword": self._normalize_keyword(r.get("keyword", "")),
                    "monthly_pc_search": int(r.get("monthly_pc_search", 0)),
                    "monthly_mobile_search": int(r.get("monthly_mobile_search", 0)),
                    "monthly_total_search": int(r.get("monthly_total_search", 0)),
                }
        except Exception:
            base_volume_map = {}

        for idx, candidate in enumerate(web_terms, start=1):
            term = self._normalize_keyword(candidate.get("keyword", ""))
            if not term:
                continue
            if progress_callback:
                progress_callback(f"[{idx}/{total_terms}] {term} 검색량 조회")
            tnorm = self._normalize_keyword(term)
            tkey = self._keyword_key(tnorm)
            row = base_volume_map.get(tkey)
            if row:
                row = {
                    "keyword": tnorm,
                    "monthly_pc_search": int(row.get("monthly_pc_search", 0)),
                    "monthly_mobile_search": int(row.get("monthly_mobile_search", 0)),
                    "monthly_total_search": int(row.get("monthly_total_search", 0)),
                }
            else:
                try:
                    row = self._resolve_term_volume_row(term)
                except Exception:
                    row = None
            if row:
                key = self._keyword_key(row.get("keyword", ""))
                payload = dict(row)
            else:
                # 수집된 키워드는 유지하되, 검색량 0으로 반영
                key = self._keyword_key(term)
                payload = {
                    "keyword": self._normalize_keyword(term),
                    "monthly_pc_search": 0,
                    "monthly_mobile_search": 0,
                    "monthly_total_search": 0,
                }
            if not key:
                continue
            prev = merged.get(key)
            if prev is None or int(payload.get("monthly_total_search", 0)) > int(prev.get("monthly_total_search", 0)):
                merged[key] = payload

        keyword_rows = list(merged.values())
        if not keyword_rows:
            return []

        # 네이버 검색결과/자동완성으로 수집된 후보 자체를 분석 대상으로 사용
        source_rows = keyword_rows
        source_rows = sorted(source_rows, key=lambda x: -x["monthly_total_search"])
        if progress_callback:
            start = max(0, int(offset))
            remaining = max(0, len(source_rows) - start)
            batch_count = min(remaining, max(1, int(limit)))
            progress_callback(f"수집 키워드 {len(source_rows)}개 중 이번 배치 {batch_count}개를 분석합니다.")

        return self._score_keyword_rows(
            source_rows,
            limit,
            offset=offset,
            progress_callback=progress_callback,
            blog_count_mode=blog_count_mode
        )

    def analyze_single_keyword_with_content(self, keyword, progress_callback=None, blog_count_mode="monthly"):
        seed = self._normalize_keyword(keyword)
        if not seed:
            raise ValueError("키워드를 입력해 주세요.")
        if progress_callback:
            progress_callback(f"'{seed}' 단일 키워드 지표 계산을 시작합니다.")

        row = self._resolve_term_volume_row(seed)
        monthly_pc = int(row.get("monthly_pc_search", 0)) if row else 0
        monthly_mobile = int(row.get("monthly_mobile_search", 0)) if row else 0
        monthly_total = int(row.get("monthly_total_search", 0)) if row else 0

        blog_docs = int(self.get_blog_document_count(seed, count_mode=blog_count_mode))
        saturation = self.calculate_content_saturation_index(monthly_total, blog_docs)
        result = [{
            "keyword": seed,
            "monthly_pc_search": monthly_pc,
            "monthly_mobile_search": monthly_mobile,
            "monthly_total_search": monthly_total,
            "blog_document_count": blog_docs,
            "content_saturation_index": float(saturation),
            "section_position": "확인 필요",
        }]
        if progress_callback:
            progress_callback(f"'{seed}' 단일 키워드 지표 계산 완료")
        return result

    def find_golden_keywords(self, category_keyword, seed_keywords=None, max_candidates=30, offset=0, progress_callback=None, blog_count_mode="monthly"):
        category_name = category_keyword.strip()
        if not category_name:
            raise ValueError("카테고리를 선택해 주세요.")

        normalized_seeds = []
        if seed_keywords:
            for seed in seed_keywords:
                text = str(seed).strip()
                if text and text not in normalized_seeds:
                    normalized_seeds.append(text)
        if not normalized_seeds:
            normalized_seeds = [category_name]

        seed_limit = min(len(normalized_seeds), 3 if max_candidates <= 30 else 4 if max_candidates <= 80 else 6)
        selected_seeds = normalized_seeds[:seed_limit]
        if progress_callback:
            progress_callback(
                f"카테고리 '{category_name}' 대표 키워드 {len(selected_seeds)}개 기반으로 연관 키워드를 수집합니다."
            )

        merged = {}
        for idx, seed in enumerate(selected_seeds, start=1):
            if progress_callback:
                progress_callback(f"[{idx}/{len(selected_seeds)}] '{seed}' 연관 키워드 수집 중")
            rows = self.get_searchad_related_keywords(seed)
            for row in rows:
                key = row["keyword"]
                prev = merged.get(key)
                if prev is None or int(row["monthly_total_search"]) > int(prev["monthly_total_search"]):
                    merged[key] = row
            time.sleep(0.05)

        keyword_rows = list(merged.values())
        if not keyword_rows:
            return []

        relevance_text = " ".join([category_name] + normalized_seeds)
        ranked_rows = sorted(
            keyword_rows,
            key=lambda x: (
                -self._category_relevance(relevance_text, x["keyword"]),
                -x["monthly_total_search"]
            )
        )
        related_first = [r for r in ranked_rows if self._category_relevance(relevance_text, r["keyword"]) > 0]
        source_rows = related_first if related_first else ranked_rows
        prefilter_limit = min(len(source_rows), max(max_candidates * 4, 150))
        source_rows = source_rows[:prefilter_limit]
        if progress_callback:
            progress_callback(f"수집 키워드 {len(source_rows)}개를 1차 선별 후, 황금 키워드를 계산합니다.")

        scored = self._score_keyword_rows(
            source_rows,
            max_candidates,
            offset=offset,
            progress_callback=progress_callback,
            blog_count_mode=blog_count_mode
        )
        # 포화 지수는 낮을수록 유리, 동률이면 검색량 높은 순
        scored.sort(
            key=lambda x: (x["content_saturation_index"], -x["monthly_total_search"])
        )
        return scored


class GoldenKeywordThread(QThread):
    finished = pyqtSignal(list)
    insight = pyqtSignal(dict)
    error = pyqtSignal(str)
    log = pyqtSignal(str)

    def __init__(
        self,
        analysis_type,
        keyword,
        limit,
        offset,
        credentials,
        category_seeds=None,
        blog_count_mode="monthly",
        single_keyword_mode=False
    ):
        super().__init__()
        self.analysis_type = analysis_type
        self.keyword = keyword
        self.limit = limit
        self.offset = offset
        self.credentials = credentials
        self.category_seeds = category_seeds or []
        self.blog_count_mode = "total" if str(blog_count_mode).strip().lower() == "total" else "monthly"
        self.single_keyword_mode = bool(single_keyword_mode)

    def run(self):
        hunter = None
        try:
            hunter = KeywordHunter(
                access_key=self.credentials["searchad_access_key"],
                secret_key=self.credentials["searchad_secret_key"],
                customer_id=self.credentials["searchad_customer_id"],
                client_id=self.credentials["naver_client_id"],
                client_secret=self.credentials["naver_client_secret"],
                usage_callback=lambda delta: API_USAGE_REPORTER.increment(delta)
            )
            if self.analysis_type == "related":
                if self.single_keyword_mode:
                    results = hunter.analyze_single_keyword_with_content(
                        self.keyword,
                        blog_count_mode=self.blog_count_mode,
                        progress_callback=lambda msg: self.log.emit(msg)
                    )
                else:
                    results = hunter.analyze_related_keywords_with_content(
                        self.keyword,
                        limit=self.limit,
                        offset=self.offset,
                        expand_seeds=self.category_seeds,
                        blog_count_mode=self.blog_count_mode,
                        progress_callback=lambda msg: self.log.emit(msg)
                    )
                try:
                    self.insight.emit(hunter.get_keyword_insight(self.keyword))
                except Exception:
                    self.insight.emit({})
            else:
                results = hunter.find_golden_keywords(
                    self.keyword,
                    seed_keywords=self.category_seeds,
                    max_candidates=self.limit,
                    offset=self.offset,
                    blog_count_mode=self.blog_count_mode,
                    progress_callback=lambda msg: self.log.emit(msg)
                )
            self.finished.emit(results)
        except Exception as e:
            self.error.emit(str(e))
        finally:
            if hunter:
                hunter.close()


class FileKeywordAnalysisThread(QThread):
    finished = pyqtSignal(list, str)
    error = pyqtSignal(str)
    log = pyqtSignal(str)
    progress = pyqtSignal(int, int)

    def __init__(self, file_path, credentials, blog_count_mode="monthly"):
        super().__init__()
        self.file_path = file_path
        self.credentials = credentials
        self.blog_count_mode = "total" if str(blog_count_mode).strip().lower() == "total" else "monthly"

    def _load_keywords(self):
        ext = os.path.splitext(self.file_path)[1].lower()
        if ext == ".csv":
            try:
                df = pd.read_csv(self.file_path, encoding="utf-8-sig")
            except Exception:
                df = pd.read_csv(self.file_path, encoding="cp949")
        elif ext in [".xlsx", ".xls"]:
            df = pd.read_excel(self.file_path)
        else:
            raise ValueError("지원 파일 형식은 xlsx, csv만 가능합니다.")

        if df is None or df.empty:
            return []

        first_col = df.iloc[:, 0].tolist()
        keywords = []
        seen = set()
        for raw in first_col:
            text = str(raw or "").strip()
            if not text or text.lower() == "nan":
                continue
            if text in ["키워드", "keyword"]:
                continue
            key = text.replace(" ", "").lower()
            if key in seen:
                continue
            seen.add(key)
            keywords.append(text)
        return keywords

    def _save_output(self, rows):
        blog_col = "전체 발행량" if self.blog_count_mode == "total" else "월간 발행량"
        out_df = pd.DataFrame({
            "키워드": [r.get("keyword", "") for r in rows],
            "월 검색량": [int(r.get("monthly_total_search", 0)) for r in rows],
            blog_col: [int(r.get("blog_document_count", 0)) for r in rows],
            "콘텐츠 포화 지수": [round(float(r.get("content_saturation_index", 0.0)), 2) for r in rows],
        })
        base, ext = os.path.splitext(self.file_path)
        if ext.lower() == ".csv":
            output_path = f"{base}_분석결과.csv"
            out_df.to_csv(output_path, index=False, encoding="utf-8-sig")
        else:
            output_path = f"{base}_분석결과.xlsx"
            out_df.to_excel(output_path, index=False)
        return output_path

    def run(self):
        hunter = None
        try:
            hunter = KeywordHunter(
                access_key=self.credentials["searchad_access_key"],
                secret_key=self.credentials["searchad_secret_key"],
                customer_id=self.credentials["searchad_customer_id"],
                client_id=self.credentials["naver_client_id"],
                client_secret=self.credentials["naver_client_secret"],
                usage_callback=lambda delta: API_USAGE_REPORTER.increment(delta)
            )
            keywords = self._load_keywords()
            if not keywords:
                raise ValueError("업로드 파일 A열에서 키워드를 찾지 못했습니다.")

            self.log.emit(f"파일 분석 시작: {len(keywords)}개 키워드")
            results = []
            total = len(keywords)
            for idx, kw in enumerate(keywords, start=1):
                self.progress.emit(idx, total)
                row = hunter._resolve_term_volume_row(kw)
                monthly_pc = int(row.get("monthly_pc_search", 0)) if row else 0
                monthly_mobile = int(row.get("monthly_mobile_search", 0)) if row else 0
                monthly_total = int(row.get("monthly_total_search", 0)) if row else 0
                blog_docs = int(hunter.get_blog_document_count(kw, count_mode=self.blog_count_mode))
                saturation = float(hunter.calculate_content_saturation_index(monthly_total, blog_docs))
                results.append({
                    "keyword": str(kw).replace("+", " ").strip(),
                    "monthly_pc_search": monthly_pc,
                    "monthly_mobile_search": monthly_mobile,
                    "monthly_total_search": monthly_total,
                    "blog_document_count": blog_docs,
                    "content_saturation_index": saturation,
                    "section_position": "확인 필요",
                })

            output_path = self._save_output(results)
            self.finished.emit(results, output_path)
        except Exception as e:
            self.error.emit(str(e))
        finally:
            if hunter:
                hunter.close()


class ParallelKeywordThread(QThread):
    finished = pyqtSignal(str)
    error = pyqtSignal(str)
    log = pyqtSignal(str, str)
    
    def __init__(self, keyword, save_dir, extract_autocomplete=True):
        super().__init__()
        self.keyword = keyword
        self.save_dir = save_dir
        self.extract_autocomplete = extract_autocomplete
        self.driver = None
        self.searcher = None
        self.is_running = True
        
    def run(self):
        try:
            self.log.emit(self.keyword, f"'{self.keyword}' 검색을 시작합니다...")
            
            # comment removed (encoding issue)
            self.driver = create_chrome_driver()
            if not self.driver:
                self.error.emit(f"'{self.keyword}' 브라우저 생성 실패")
                return

            # comment removed (encoding issue)
            searcher = NaverMobileSearchScraper(driver=self.driver)
            searcher.save_dir = self.save_dir
            searcher.is_running = self.is_running
            searcher.search_thread = self
            self.searcher = searcher
            
            # comment removed (encoding issue)
            success = searcher.recursive_keyword_extraction(
                self.keyword, 
                progress_callback=self._log_wrapper,
                extract_autocomplete=self.extract_autocomplete
            )
            
            current_time = datetime.now().strftime('%Y%m%d_%H%M%S')
            safe_keyword = re.sub(r"[^\w가-힣\s]", "", self.keyword).strip()[:20] or "키워드"
            filename_suffix = "키워드추출" if self.is_running and success else "중간저장"
            filename = f"{safe_keyword}_{filename_suffix}_{current_time}.xlsx"
            save_path = os.path.join(self.save_dir, filename)

            saved_path = None
            has_partial_result = bool(
                searcher and
                hasattr(searcher, "all_related_keywords") and
                searcher.all_related_keywords
            )
            if has_partial_result:
                if searcher.save_recursive_results_to_excel(save_path, self._log_wrapper):
                    saved_path = save_path

            if success and self.is_running:
                if saved_path:
                    self.finished.emit(saved_path)
                    self.log.emit(self.keyword, f"'{self.keyword}' 처리 완료: {filename}")
                else:
                    self.error.emit(f"'{self.keyword}' 저장할 결과가 없습니다.")
            elif not self.is_running:
                if saved_path:
                    self.finished.emit(saved_path)
                    self.log.emit(self.keyword, f"'{self.keyword}' 중단됨 - 중간 결과 저장 완료")
                else:
                    self.error.emit(f"'{self.keyword}' 중단됨 (저장할 결과 없음)")
            else:
                self.error.emit(f"'{self.keyword}' 추출 실패 또는 중단됨")
                
        except Exception as e:
            self.error.emit(f"'{self.keyword}' 작업 중 오류: {str(e)}")
            
        finally:
            # comment removed (encoding issue)
            if self.driver:
                try:
                    self.driver.quit()
                except:
                    pass
                self.driver = None

    def _log_wrapper(self, msg):
        """Description"""
        self.log.emit(self.keyword, msg)

    def stop(self):
        """Description"""
        self.is_running = False
        if self.searcher:
            self.searcher.is_running = False
        if self.driver:
            try:
                self.driver.quit()
            except Exception:
                pass

        

STYLESHEET = f"""
    QMainWindow, QWidget {{
        background-color: {BACKGROUND_MAIN};
        color: #333333 !important;
    }}

    QGroupBox {{
        font-weight: bold;
        background-color: {BACKGROUND_CARD};
        border: 2px solid {BORDER_COLOR};
        border-radius: 16px;
        padding-top: 30px;
        margin-top: 20px;
        color: #333333 !important;
        font-size: 16px;
    }}
    QGroupBox::title {{
        subcontrol-origin: margin;
        left: 25px;
        padding: 8px 15px;
        color: {NAVER_GREEN} !important;
        font-weight: bold;
        font-size: 18px;
        background-color: {BACKGROUND_CARD};
        border: 2px solid {NAVER_GREEN};
        border-radius: 8px;
    }}
    QGroupBox#extractSearchGroup::title,
    QGroupBox#extractProgressGroup::title,
    QGroupBox#extractSaveGroup::title {{
        color: #111111 !important;
    }}

    QPushButton {{
        background-color: {NAVER_GREEN};
        color: {WHITE_COLOR} !important;
        border-radius: 8px;
        padding: 10px 20px;
        font-size: 14px;
        font-weight: bold;
        border: none;
    }}
    QPushButton:hover {{
        background-color: {NAVER_GREEN_DARK};
        color: {WHITE_COLOR} !important;
    }}
    QPushButton:disabled {{
        background-color: #dcdcdc;
        color: #666666 !important;
    }}

    QLineEdit {{
        padding: 10px;
        font-size: 14px;
        border: 1px solid {BORDER_COLOR};
        border-radius: 8px;
        background-color: {WHITE_COLOR};
        color: #333333 !important;
    }}
    QLineEdit:read-only {{
        background-color: {NAVER_GREEN_LIGHT};
        color: #333333 !important;
    }}

    QTextEdit {{
        font-size: 14px;
        border: 1px solid {BORDER_COLOR};
        border-radius: 8px;
        padding: 10px;
        background-color: {WHITE_COLOR};
        color: #333333 !important;
    }}

    QCheckBox {{
        color: #333333 !important;
        font-size: 14px;
        font-weight: 600;
        spacing: 15px;
    }}
    QCheckBox::indicator {{
        width: 24px;
        height: 24px;
        border: 3px solid {BORDER_COLOR};
        border-radius: 6px;
        background-color: {BACKGROUND_CARD};
    }}
    QCheckBox::indicator:checked {{
        background-color: {NAVER_GREEN};
        border: 3px solid {NAVER_GREEN};
    }}

    QStatusBar {{
        background: {BACKGROUND_CARD};
        color: #333333 !important;
        border-top: 3px solid {NAVER_GREEN};
        padding: 12px;
        font-size: 14px;
        font-weight: 600;
    }}

    QLabel {{
        color: #333333 !important;
    }}
"""

DARK_STYLESHEET = """
    QMainWindow, QWidget {
        background-color: #12161b;
        color: #e6edf3 !important;
    }

    QGroupBox {
        font-weight: bold;
        background-color: #1b222b;
        border: 2px solid #30404d;
        border-radius: 16px;
        padding-top: 30px;
        margin-top: 20px;
        color: #e6edf3 !important;
        font-size: 16px;
    }
    QGroupBox::title {
        subcontrol-origin: margin;
        left: 25px;
        padding: 8px 15px;
        color: #9be2bc !important;
        font-weight: bold;
        font-size: 18px;
        background-color: #1b222b;
        border: 2px solid #00a83a;
        border-radius: 8px;
    }

    QPushButton {
        background-color: #00a83a;
        color: #ffffff !important;
        border-radius: 8px;
        padding: 10px 20px;
        font-size: 14px;
        font-weight: bold;
        border: none;
    }
    QPushButton:hover {
        background-color: #078f37;
        color: #ffffff !important;
    }
    QPushButton:disabled {
        background-color: #3a4652;
        color: #9aa8b6 !important;
    }

    QLineEdit, QTextEdit, QComboBox, QSpinBox, QDoubleSpinBox {
        padding: 10px;
        font-size: 14px;
        border: 1px solid #3a4652;
        border-radius: 8px;
        background-color: #171d24;
        color: #e6edf3 !important;
    }
    QLineEdit:read-only {
        background-color: #1f2a24;
        color: #d3dfd8 !important;
    }

    QCheckBox {
        color: #dbe5ef !important;
        font-size: 14px;
        font-weight: 600;
        spacing: 15px;
    }
    QCheckBox::indicator {
        width: 24px;
        height: 24px;
        border: 3px solid #3a4652;
        border-radius: 6px;
        background-color: #1b222b;
    }
    QCheckBox::indicator:checked {
        background-color: #00a83a;
        border: 3px solid #00a83a;
    }

    QStatusBar {
        background: #1b222b;
        color: #d7e2ed !important;
        border-top: 3px solid #00a83a;
        padding: 12px;
        font-size: 14px;
        font-weight: 600;
    }

    QLabel {
        color: #dce6f0 !important;
    }
"""


def create_chrome_driver(log_callback=None, force_headless=False):
    """Description"""
    try:
        if log_callback:
            log_callback("log update")
        
        try:
            driver_path = ChromeDriverManager().install()
            if log_callback:
                log_callback("log update")
        except Exception as e:
            if log_callback:
                log_callback("log update")
            driver_path = None
        
        options = webdriver.ChromeOptions()
        if force_headless or SELENIUM_HEADLESS:
            options.add_argument("--headless=new")
        options.add_argument("--window-size=1920,1080")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--disable-gpu")
        options.add_argument("--log-level=3")
        options.add_argument("--silent")

        if driver_path:
            service = Service(driver_path)
        else:
            service = Service()
        if os.name == "nt":
            try:
                service.creation_flags = subprocess.CREATE_NO_WINDOW
            except Exception:
                pass
        
        driver = webdriver.Chrome(service=service, options=options)
        driver.set_page_load_timeout(20)
        driver.implicitly_wait(5)
        
        driver.get("about:blank")
        
        if log_callback:
            log_callback("log update")
            
        return driver

    except Exception as e:
        if log_callback:
            log_callback("log update")
        raise e


class KeywordExtractorMainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        
        # comment removed (encoding issue)
        icon_path = get_icon_path()
        if icon_path:
            self.setWindowIcon(QIcon(icon_path))
            safe_print(f"아이콘 설정 완료: {icon_path}")
        else:
            safe_print("아이콘 파일을 찾을 수 없습니다.")
        
        self.setWindowTitle("네이버 연관키워드 추출기")
        self.resize(1200, 800)
        self.showMaximized()
        
        # comment removed (encoding issue)
        self.settings = Settings()
        self.driver = None
        # comment removed (encoding issue)
        self.active_threads = []
        self.completed_threads = 0
        self.total_threads = 0
        self.stop_requested = False
        self.related_keyword_results = []
        self.category_keyword_results = []
        self.golden_keyword_thread = None
        self.file_keyword_thread = None
        self.last_analysis_keyword = {"related": "", "category": ""}
        self.analysis_offset = {"related": 0, "category": 0}
        self.analysis_keep_existing = {"related": False, "category": False}
        self.current_analysis_mode = ""
        self.blog_count_mode = self.settings.get_blog_count_mode()
        self.related_single_mode = False
        self.related_progress_total = 0
        self.related_spinner_timer = QTimer(self)
        self.related_spinner_timer.setInterval(120)
        self.related_spinner_timer.timeout.connect(self._tick_related_spinner)
        
        # comment removed (encoding issue)
        self.setup_crash_protection()
        
        # comment removed (encoding issue)
        self.init_ui()
        self.setup_chrome_driver()
        self.current_theme_mode = "light"
        self.apply_theme(self.settings.get_theme_mode(), save=False)

    def check_license_info(self):
        """Description"""
        # comment removed (encoding issue)
        machine_id = get_machine_id()
        
        # comment removed (encoding issue)
        expiration_date = check_license_from_sheet(machine_id)
        
        if expiration_date:
            try:
                # comment removed (encoding issue)
                exp_date = datetime.strptime(str(expiration_date).strip(), '%Y-%m-%d')
                today = datetime.now()
                
                if exp_date < today:
                    # comment removed (encoding issue)
                    self.show_license_dialog(machine_id, expired=True)
                else:
                    # comment removed (encoding issue)
                    self.usage_label.setText(f"사용 기간: {expiration_date}까지")
                    self.usage_label.setStyleSheet(f"font-size: 13px; font-weight: bold; color: {NAVER_GREEN};")
            except:
                # comment removed (encoding issue)
                # comment removed (encoding issue)
                self.usage_label.setText(f"사용 기간: {expiration_date}")
                self.usage_label.setStyleSheet(f"font-size: 13px; font-weight: bold; color: {NAVER_GREEN};")
        else:
            # comment removed (encoding issue)
            self.show_license_dialog(machine_id)
            
    def show_license_dialog(self, machine_id, expired=False):
        """라이선스 상태에 맞는 안내창 표시"""
        if expired:
            dialog = ExpiredDialog(datetime.now().strftime("%Y-%m-%d"))
        else:
            dialog = UnregisteredDialog(machine_id)
        dialog.exec()
        sys.exit(0)

    def setup_crash_protection(self):
        """Description"""
        global _current_window
        _current_window = self
        sys.excepthook = handle_exception
        
        # comment removed (encoding issue)
        QTimer.singleShot(100, self.check_license_info)
        
        try:
            signal.signal(signal.SIGINT, handle_signal)
            if hasattr(signal, 'SIGBREAK'):
                signal.signal(signal.SIGBREAK, handle_signal)
        except Exception as e:
            safe_print(f"신호 핸들러 설정 실패: {str(e)}")
        
        atexit.register(emergency_save_data)
        safe_print("크래시 보호 시스템이 활성화되었습니다.")

    def setup_chrome_driver(self):
        """Description"""
        # comment removed (encoding issue)
        # comment removed (encoding issue)
        # comment removed (encoding issue)
        pass

    def init_ui(self):
        """Initialize UI."""
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        central_layout = QVBoxLayout(central_widget)
        central_layout.setContentsMargins(0, 0, 0, 0)

        self.nav_widget = QWidget()
        nav_layout = QHBoxLayout(self.nav_widget)
        nav_layout.setContentsMargins(10, 6, 10, 4)
        nav_layout.setSpacing(8)
        left_slot = QWidget()
        left_slot.setFixedWidth(240)
        left_slot_layout = QHBoxLayout(left_slot)
        left_slot_layout.setContentsMargins(0, 0, 0, 0)
        left_slot_layout.setSpacing(0)
        self.theme_toggle_button = QPushButton(self._theme_button_text("light"))
        self.theme_toggle_button.setObjectName("themeToggleButton")
        self.theme_toggle_button.setFixedWidth(240)
        self.theme_toggle_button.clicked.connect(self.toggle_theme_mode)
        left_slot_layout.addWidget(self.theme_toggle_button)
        self.section_related_button = QPushButton("연관 키워드 추출")
        self.section_related_button.setCheckable(True)
        self.section_gold_button = QPushButton("황금 키워드 분석")
        self.section_gold_button.setCheckable(True)
        self.usage_label = QLabel("사용 기간: 확인 중...")
        self.usage_label.setStyleSheet("font-size: 13px; font-weight: 700; color: #555555;")
        self.usage_label.setFixedWidth(240)
        self.usage_label.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        self.section_related_button.clicked.connect(lambda: self.switch_main_section(0))
        self.section_gold_button.clicked.connect(lambda: self.switch_main_section(1))
        nav_layout.addWidget(left_slot)
        nav_layout.addStretch(1)
        nav_layout.addWidget(self.section_related_button)
        nav_layout.addWidget(self.section_gold_button)
        nav_layout.addStretch(1)
        nav_layout.addWidget(self.usage_label)
        self.nav_widget.setStyleSheet(self._nav_stylesheet("light"))
        central_layout.addWidget(self.nav_widget)

        self.main_tabs = QTabWidget()
        self.main_tabs.setObjectName("mainNavigationTabs")
        self.main_tabs.setStyleSheet(self._main_tabs_stylesheet("light"))
        main_tab_bar = self.main_tabs.tabBar()
        if main_tab_bar is not None:
            main_tab_bar.hide()
        self.main_tabs.currentChanged.connect(self._sync_main_section_buttons)
        central_layout.addWidget(self.main_tabs)

        extractor_tab = QWidget()
        extractor_tab_layout = QVBoxLayout(extractor_tab)
        extractor_tab_layout.setContentsMargins(0, 0, 0, 0)

        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        extractor_tab_layout.addWidget(scroll_area)

        scroll_content = QWidget()
        main_layout = QVBoxLayout(scroll_content)
        main_layout.setSpacing(10)
        main_layout.setContentsMargins(10, 10, 10, 10)
        scroll_area.setWidget(scroll_content)
        
        # comment removed (encoding issue)
        top_section_layout = QHBoxLayout()
        
        # comment removed (encoding issue)
        self.setup_search_section(top_section_layout)
        
        # comment removed (encoding issue)
        self.setup_progress_section(top_section_layout)
        
        main_layout.addLayout(top_section_layout)
        
        # comment removed (encoding issue)
        self.setup_save_section(main_layout)
        self.main_tabs.addTab(extractor_tab, "연관 키워드 추출")

        analysis_tab = QWidget()
        analysis_tab_layout = QVBoxLayout(analysis_tab)
        analysis_tab_layout.setContentsMargins(10, 10, 10, 10)
        self.setup_golden_keyword_section(analysis_tab_layout)
        self.main_tabs.addTab(analysis_tab, "황금 키워드 분석")
        self.switch_main_section(0)
            
        # comment removed (encoding issue)
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("준비 완료")

    def switch_main_section(self, index):
        self.main_tabs.setCurrentIndex(index)
        self._sync_main_section_buttons(index)

    def _sync_main_section_buttons(self, index):
        self.section_related_button.setChecked(index == 0)
        self.section_gold_button.setChecked(index == 1)

    def _theme_button_text(self, mode):
        return "테마: 다크" if mode == "dark" else "테마: 라이트"

    def toggle_theme_mode(self):
        next_mode = "dark" if getattr(self, "current_theme_mode", "light") == "light" else "light"
        self.apply_theme(next_mode)

    def apply_theme(self, mode, save=True):
        mode = "dark" if str(mode).strip().lower() == "dark" else "light"
        self.current_theme_mode = mode
        self.setStyleSheet(DARK_STYLESHEET if mode == "dark" else STYLESHEET)

        if hasattr(self, "theme_toggle_button"):
            self.theme_toggle_button.setText(self._theme_button_text(mode))
        if hasattr(self, "nav_widget"):
            self.nav_widget.setStyleSheet(self._nav_stylesheet(mode))
        if hasattr(self, "main_tabs"):
            self.main_tabs.setStyleSheet(self._main_tabs_stylesheet(mode))
        if hasattr(self, "progress_tabs"):
            self.progress_tabs.setStyleSheet(self._progress_tabs_stylesheet(mode))
        if hasattr(self, "golden_root_widget"):
            self.golden_root_widget.setStyleSheet(self._golden_root_stylesheet(mode))
        if hasattr(self, "related_table"):
            self.related_table.setStyleSheet(self._result_table_stylesheet(mode))
        if hasattr(self, "related_table_placeholder"):
            self.related_table_placeholder.setStyleSheet(self._result_table_stylesheet(mode))
            self._populate_related_guide_table()
        if hasattr(self, "category_table"):
            self.category_table.setStyleSheet(self._result_table_stylesheet(mode))
        if hasattr(self, "related_spinner"):
            self.related_spinner.set_mode(mode)

        if save and hasattr(self, "settings"):
            self.settings.set_theme_mode(mode)

    def _nav_stylesheet(self, mode):
        if mode == "dark":
            return """
                QPushButton {
                    min-width: 180px;
                    min-height: 34px;
                    padding: 6px 12px;
                    border-radius: 8px;
                    font-size: 14px;
                    font-weight: 800;
                    color: #d8f7e8;
                    background: #1f3a2f;
                    border: 1px solid #2f5b47;
                }
                QPushButton:checked {
                    color: #ffffff;
                    background: #00a83a;
                    border: 1px solid #04c14a;
                }
                QPushButton:hover:!checked {
                    background: #28483a;
                }
            """
        return """
            QPushButton {
                min-width: 180px;
                min-height: 34px;
                padding: 6px 12px;
                border-radius: 8px;
                font-size: 14px;
                font-weight: 800;
                color: #1f5b40;
                background: #e7f2ec;
                border: 1px solid #b9ddcc;
            }
            QPushButton:checked {
                color: #ffffff;
                background: #00c73c;
                border: 1px solid #00ae35;
            }
            QPushButton:hover:!checked {
                background: #dff0e7;
            }
        """

    def _main_tabs_stylesheet(self, mode):
        if mode == "dark":
            return """
                QTabWidget#mainNavigationTabs::pane {
                    border: none;
                    top: -1px;
                }
                QTabWidget#mainNavigationTabs QTabBar::tab {
                    min-width: 220px;
                    min-height: 40px;
                    padding: 8px 18px;
                    margin: 8px 6px 2px 6px;
                    border-radius: 10px;
                    font-size: 15px;
                    font-weight: 800;
                    color: #cceedd;
                    background: #223029;
                    border: 1px solid #355444;
                }
                QTabWidget#mainNavigationTabs QTabBar::tab:selected {
                    color: #ffffff;
                    background: #00a83a;
                    border: 1px solid #04c14a;
                }
                QTabWidget#mainNavigationTabs QTabBar::tab:hover:!selected {
                    background: #2a3a32;
                }
                QTabWidget#mainNavigationTabs QTabWidget::tab-bar {
                    alignment: center;
                }
            """
        return """
            QTabWidget#mainNavigationTabs::pane {
                border: none;
                top: -1px;
            }
            QTabWidget#mainNavigationTabs QTabBar::tab {
                min-width: 220px;
                min-height: 40px;
                padding: 8px 18px;
                margin: 8px 6px 2px 6px;
                border-radius: 10px;
                font-size: 15px;
                font-weight: 800;
                color: #2b5a47;
                background: #e9f6ef;
                border: 1px solid #bfe3cf;
            }
            QTabWidget#mainNavigationTabs QTabBar::tab:selected {
                color: #ffffff;
                background: #00c73c;
                border: 1px solid #00ae35;
            }
            QTabWidget#mainNavigationTabs QTabBar::tab:hover:!selected {
                background: #dcf2e6;
            }
            QTabWidget#mainNavigationTabs QTabWidget::tab-bar {
                alignment: center;
            }
        """

    def _progress_tabs_stylesheet(self, mode):
        if mode == "dark":
            return """
                QTabWidget#progressTabs {
                    background: #171b20;
                }
                QTabWidget#progressTabs::pane {
                    border: 1px solid #3a4652;
                    border-radius: 8px;
                    top: 0px;
                }
                QTabWidget#progressTabs::tab-bar { alignment: left; }
                QTabWidget#progressTabs QTabBar {
                    background: #171b20;
                    border: none;
                }
                QTabWidget#progressTabs QTabBar::tab {
                    background: #242b33;
                    color: #d9e3ec;
                    padding: 1px 6px;
                    min-width: 40px;
                    max-width: 68px;
                    min-height: 20px;
                    font-size: 12px;
                    font-weight: 400;
                    border: 1px solid #3a4652;
                    border-top-left-radius: 6px;
                    border-top-right-radius: 6px;
                    margin-right: 2px;
                }
                QTabWidget#progressTabs QTabBar::tab:selected {
                    background: #2b4b66;
                    color: #9fd0ff;
                    border: 1px solid #4a7aa4;
                    border-bottom: 1px solid #4a7aa4;
                }
            """
        return """
            QTabWidget#progressTabs {
                background: #ffffff;
            }
            QTabWidget#progressTabs::pane {
                border: 1px solid #cccccc;
                border-radius: 8px;
                top: 0px;
            }
            QTabWidget#progressTabs::tab-bar { alignment: left; }
            QTabWidget#progressTabs QTabBar {
                background: #ffffff;
                border: none;
            }
            QTabWidget#progressTabs QTabBar::tab {
                background: #f0f0f0;
                padding: 1px 6px;
                min-width: 40px;
                max-width: 68px;
                min-height: 20px;
                font-size: 12px;
                font-weight: 400;
                border: 1px solid #d0d0d0;
                border-top-left-radius: 6px;
                border-top-right-radius: 6px;
                margin-right: 2px;
            }
            QTabWidget#progressTabs QTabBar::tab:selected {
                background: #E6F0FD;
                color: #1E6ECA;
                border: 1px solid #1E6ECA;
                border-bottom: 1px solid #1E6ECA;
            }
        """

    def _result_table_stylesheet(self, mode):
        if mode == "dark":
            return """
                QTableWidget {
                    gridline-color: #31404d;
                    border: 1px solid #3a4652;
                    border-radius: 8px;
                    background: #171b20;
                    color: #e3ecf5;
                    alternate-background-color: #1f252d;
                }
                QHeaderView::section {
                    background-color: #23303c;
                    color: #d2e7ff;
                    font-weight: 700;
                    padding: 6px;
                    border: 0px;
                    border-bottom: 1px solid #3f5364;
                }
            """
        return """
            QTableWidget {
                gridline-color: #d9e9e0;
                border: 1px solid #d4edda;
                border-radius: 8px;
                background: #ffffff;
                alternate-background-color: #f6fbf8;
            }
            QHeaderView::section {
                background-color: #e8f5f0;
                color: #1f5136;
                font-weight: 700;
                padding: 6px;
                border: 0px;
                border-bottom: 1px solid #cde9d8;
            }
        """

    def _golden_root_stylesheet(self, mode):
        if mode == "dark":
            return """
                QGroupBox { font-size: 14px; font-weight: 700; }
                QGroupBox::title {
                    subcontrol-origin: margin;
                    left: 25px;
                    padding: 8px 15px;
                    color: #bdecd0;
                    font-size: 18px;
                    font-weight: 700;
                    background-color: #171b20;
                    border: 2px solid #00a83a;
                    border-radius: 8px;
                }
                QGroupBox#leftPanel, QGroupBox#rightPanel {
                    background: #1d242c;
                    border: 1px solid #2f3a44;
                    border-radius: 10px;
                    margin-top: 8px;
                }
            """
        return """
            QGroupBox { font-size: 14px; font-weight: 700; }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 25px;
                padding: 8px 15px;
                color: #185a3a;
                font-size: 18px;
                font-weight: 700;
                background-color: #ffffff;
                border: 2px solid #03c75a;
                border-radius: 8px;
            }
            QGroupBox#leftPanel, QGroupBox#rightPanel {
                background: #fbfdfc;
                border: 1px solid #cfe8d7;
                border-radius: 10px;
                margin-top: 8px;
            }
        """

    def setup_search_section(self, main_layout):
        """검색 섹션 설정"""
        search_group = QGroupBox("키워드 검색")
        search_group.setObjectName("extractSearchGroup")
        search_layout = QVBoxLayout(search_group)
        
        # comment removed (encoding issue)
        self.search_input = MultiKeywordTextEdit()
        self.search_input.setPlaceholderText(
            "사용 방법\n"
            "1. 키워드를 한 줄에 하나씩 입력하세요.\n"
            "2. Enter로 바로 추출 시작, Shift+Enter로 줄바꿈합니다.\n"
            "3. 여러 키워드를 동시에 병렬 처리합니다."
        )
        # comment removed (encoding issue)
        self.search_input.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self.search_input.search_requested.connect(self.start_search)
        search_layout.addWidget(self.search_input)
        
        # comment removed (encoding issue)
        search_layout.setContentsMargins(10, 10, 10, 10)
        search_layout.setSpacing(5)
        
        # comment removed (encoding issue)
        button_layout = QHBoxLayout()
        
        self.start_button = QPushButton("키워드 추출 시작")
        self.start_button.clicked.connect(self.start_search)
        
        self.pause_button = QPushButton("일시정지")
        self.pause_button.clicked.connect(self.pause_resume_search)
        self.pause_button.setEnabled(False)
        
        self.stop_button = QPushButton("중단")
        self.stop_button.clicked.connect(self.stop_search)
        self.stop_button.setEnabled(False)
        
        button_layout.addWidget(self.start_button)
        button_layout.addWidget(self.pause_button)
        button_layout.addWidget(self.stop_button)
        search_layout.addLayout(button_layout)
        
        main_layout.addWidget(search_group)

    def setup_save_section(self, main_layout):
        """저장 위치 섹션 설정"""
        save_group = QGroupBox("저장 위치")
        save_group.setObjectName("extractSaveGroup")
        save_layout = QVBoxLayout(save_group)
        
        path_layout = QHBoxLayout()
        self.save_path_input = QLineEdit()
        self.save_path_input.setReadOnly(True)
        self.save_path_input.setPlaceholderText("저장할 폴더를 선택하세요.")
        
        self.browse_button = QPushButton("폴더 선택")
        self.browse_button.clicked.connect(self.change_save_directory)
        
        path_layout.addWidget(self.save_path_input)
        path_layout.addWidget(self.browse_button)
        
        # comment removed (encoding issue)
        saved_dir = self.settings.get_save_dir()
        if saved_dir and os.path.exists(saved_dir):
            self.save_path_input.setText(saved_dir)
        else:
            # comment removed (encoding issue)
            desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
            default_dir = os.path.join(desktop_path, "keyword_results")
            
            try:
                os.makedirs(default_dir, exist_ok=True)
            except Exception:
                # comment removed (encoding issue)
                default_dir = os.path.join(os.path.expanduser("~"), "Documents", "keyword_results")
                os.makedirs(default_dir, exist_ok=True)
                
            self.save_path_input.setText(default_dir)
        
        save_layout.addLayout(path_layout)
        
        self.remember_checkbox = QCheckBox("저장 경로 기억하기")
        self.remember_checkbox.setChecked(self.settings.should_remember_dir())
        save_layout.addWidget(self.remember_checkbox)
        
        main_layout.addWidget(save_group)

    def setup_progress_section(self, main_layout):
        """진행 상황 섹션 설정"""
        progress_group = QGroupBox("진행 상황")
        progress_group.setObjectName("extractProgressGroup")
        progress_layout = QVBoxLayout(progress_group)
        progress_layout.setContentsMargins(10, 2, 10, 10)
        progress_layout.setSpacing(0)
        
        # comment removed (encoding issue)
        self.progress_tabs = QTabWidget()
        self.progress_tabs.setObjectName("progressTabs")
        progress_tab_bar = self.progress_tabs.tabBar()
        if progress_tab_bar is not None:
            progress_tab_bar.setExpanding(False)
            progress_tab_bar.setUsesScrollButtons(False)
        self.progress_tabs.setStyleSheet(self._progress_tabs_stylesheet("light"))
        
        # comment removed (encoding issue)
        self.total_log_text = SmartProgressTextEdit(min_height=100, max_height=800)
        self.total_log_text.setReadOnly(True)
        self.total_log_text.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self.total_log_text.setPlaceholderText("여기에 전체 진행 로그가 표시됩니다.")
        self.progress_tabs.addTab(self.total_log_text, "전체 로그")
        
        # comment removed (encoding issue)
        self.log_widgets = {"전체": self.total_log_text}
        
        progress_layout.addWidget(self.progress_tabs)
        main_layout.addWidget(progress_group)



    def change_save_directory(self):
        """Change save directory."""
        current_dir = self.save_path_input.text() or os.getcwd()
        directory = QFileDialog.getExistingDirectory(self, "저장할 폴더 선택", current_dir)
        if directory:
            self.save_path_input.setText(directory)
            if self.remember_checkbox.isChecked():
                self.settings.set_save_dir(directory, True)
                self.update_progress(f"저장 위치가 기억되었습니다: {directory}")

    def setup_golden_keyword_section(self, parent_layout):
        root_widget = QWidget()
        self.golden_root_widget = root_widget
        root = QVBoxLayout(root_widget)
        root.setSpacing(10)
        root.setContentsMargins(0, 0, 0, 0)

        self.category_seed_map = {
            "생활/리빙": ["생활꿀팁", "청소", "수납", "정리정돈", "살림", "집꾸미기"],
            "건강/의료": ["건강관리", "영양제", "혈압", "다이어트", "건강검진", "병원예약"],
            "교육/학습": ["공부법", "인강", "자격증", "영어회화", "국어문법", "수학문제"],
            "금융/재테크": ["재테크", "연말정산", "세금환급", "절세", "예적금", "신용점수"],
            "부동산": ["부동산", "전세", "월세", "청약", "아파트", "대출금리"],
            "자동차": ["자동차", "중고차", "보험료", "연비", "정비", "타이어"],
            "여행/숙박": ["국내여행", "해외여행", "호텔", "항공권", "여행코스", "여행준비물"],
            "뷰티/미용": ["피부관리", "선크림", "여드름", "헤어스타일", "메이크업", "화장품"],
            "패션": ["데일리룩", "코디", "운동화", "가을패션", "겨울코트", "쇼핑몰"],
            "식품/요리": ["집밥레시피", "다이어트식단", "반찬", "에어프라이어", "밀프렙", "요리팁"],
            "IT/전자": ["IT기기", "스마트폰", "노트북", "태블릿", "무선이어폰", "게이밍마우스"],
            "육아/아동": ["육아정보", "아기수면", "이유식", "어린이집", "초등학습", "육아템"],
            "취업/자격증": ["취업준비", "이력서", "면접", "자소서", "국비지원", "자격증시험"],
            "법률/행정": ["행정서류", "민원24", "등본", "가족관계증명서", "계약서", "소송절차"],
            "반려동물": ["반려동물", "강아지사료", "고양이사료", "예방접종", "산책", "훈련"],
            "스포츠/레저": ["홈트", "헬스", "러닝", "필라테스", "자전거", "등산"],
            "문화/공연": ["전시공연", "뮤지컬", "콘서트", "영화개봉", "연극", "전시회"],
            "인테리어": ["집꾸미기", "셀프인테리어", "조명", "커튼", "가구배치", "리모델링"],
            "청소/가사": ["청소팁", "세탁", "욕실청소", "주방청소", "정리정돈", "살림템"],
            "기타 서비스": ["생활서비스", "세무상담", "노무상담", "이사견적", "보험상담", "법무사"],
        }

        split = QHBoxLayout()
        split.setSpacing(10)

        # Left panel: related keyword analysis
        left_group = QGroupBox("연관 키워드 분석")
        left_group.setObjectName("leftPanel")
        left_layout = QVBoxLayout(left_group)
        left_layout.setSpacing(8)
        left_layout.setContentsMargins(10, 10, 10, 10)

        left_top = QHBoxLayout()
        self.related_keyword_input = KoreanDefaultLineEdit()
        self.related_keyword_input.setPlaceholderText("예: 연말정산")
        self.related_keyword_input.returnPressed.connect(self.start_related_keyword_analysis)
        left_top.addWidget(self.related_keyword_input, 1)

        self.blog_count_mode_combo = QComboBox()
        self.blog_count_mode_combo.setObjectName("blogCountModeCombo")
        self.blog_count_mode_combo.addItem("월간 발행량", "monthly")
        self.blog_count_mode_combo.addItem("전체 발행량", "total")
        self.blog_count_mode_combo.setCurrentIndex(0 if self.blog_count_mode == "monthly" else 1)
        self.blog_count_mode_combo.currentIndexChanged.connect(self.on_blog_count_mode_changed)
        left_top.addWidget(self.blog_count_mode_combo)

        self.related_batch_size = 10

        self.related_upload_button = QPushButton("파일 업로드")
        self.related_upload_button.setObjectName("uploadButton")
        self.related_upload_button.clicked.connect(self.start_related_file_analysis)
        left_top.addWidget(self.related_upload_button)

        self.related_keyword_button = QPushButton("분석 실행")
        self.related_keyword_button.clicked.connect(self.start_related_keyword_analysis)
        self.related_single_button = QPushButton("단일 키워드")
        self.related_single_button.clicked.connect(self.start_single_keyword_analysis)
        self.related_save_button = QPushButton("저장")
        self.related_save_button.setEnabled(False)
        self.related_save_button.clicked.connect(lambda: self.save_results_for_mode("related"))
        left_top.addWidget(self.related_keyword_button)
        left_top.addWidget(self.related_single_button)
        left_top.addWidget(self.related_save_button)
        left_layout.addLayout(left_top)

        self.blog_count_mode_hint = QLabel("")
        self.blog_count_mode_hint.setObjectName("summaryHint")
        left_layout.addWidget(self.blog_count_mode_hint)
        self._update_blog_count_mode_hint()

        self.related_sort_hint = QLabel("")
        self.related_sort_hint.setObjectName("sortHint")
        self.related_sort_hint.setVisible(False)
        left_layout.addWidget(self.related_sort_hint)
        self.related_loading_widget = QWidget()
        related_loading_layout = QHBoxLayout(self.related_loading_widget)
        related_loading_layout.setContentsMargins(4, 2, 4, 2)
        related_loading_layout.setSpacing(8)
        self.related_spinner = SpiralSpinner()
        self.related_loading_text = QLabel("작업 중...")
        self.related_loading_text.setObjectName("loadingText")
        self.related_progress_bar = QProgressBar()
        self.related_progress_bar.setMinimum(0)
        self.related_progress_bar.setMaximum(100)
        self.related_progress_bar.setValue(0)
        self.related_progress_bar.setTextVisible(True)
        related_loading_layout.addWidget(self.related_spinner)
        related_loading_layout.addWidget(self.related_loading_text)
        related_loading_layout.addWidget(self.related_progress_bar, 1)
        self.related_loading_widget.setVisible(False)
        left_layout.addWidget(self.related_loading_widget)

        self.related_table = QTableWidget()
        self._init_result_table(self.related_table)
        self.related_table_placeholder = QTableWidget()
        self.related_table_placeholder.setObjectName("relatedGuideTable")
        self._init_result_table(self.related_table_placeholder)
        self._populate_related_guide_table()

        self.related_table_stack_widget = QWidget()
        self.related_table_stack = QStackedLayout(self.related_table_stack_widget)
        self.related_table_stack.setContentsMargins(0, 0, 0, 0)
        self.related_table_stack.addWidget(self.related_table_placeholder)
        self.related_table_stack.addWidget(self.related_table)
        self.related_table_stack.setCurrentIndex(0)
        left_layout.addWidget(self.related_table_stack_widget)
        related_more_row = QHBoxLayout()
        related_more_row.addStretch(1)
        self.related_more_button = QPushButton("더 많은 연관 키워드 보기")
        self.related_more_button.clicked.connect(self.start_related_keyword_analysis_more)
        self.related_more_button.setEnabled(False)
        self.related_more_button.setObjectName("moreLinkButton")
        related_more_row.addWidget(self.related_more_button)
        related_more_row.addStretch(1)
        left_layout.addLayout(related_more_row)
        split.addWidget(left_group, 1)

        # Right panel: category golden keyword recommendation (read-only for now)
        right_group = QGroupBox("카테고리 황금키워드 추천")
        right_group.setObjectName("rightPanel")
        right_layout = QVBoxLayout(right_group)
        right_layout.setSpacing(8)
        right_layout.setContentsMargins(10, 10, 10, 10)

        right_top = QHBoxLayout()
        right_top.setSpacing(10)
        self.golden_category_combo = QComboBox()
        self.golden_category_combo.addItems(list(self.category_seed_map.keys()))
        self.golden_category_combo.setObjectName("categoryCenterCombo")
        self.golden_category_combo.setEditable(True)
        combo_line_edit = self.golden_category_combo.lineEdit()
        if combo_line_edit is not None:
            combo_line_edit.setReadOnly(True)
            combo_line_edit.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.golden_category_combo.setInsertPolicy(QComboBox.InsertPolicy.NoInsert)
        right_top.addWidget(self.golden_category_combo, 1)

        right_limit_card = QFrame()
        right_limit_card.setObjectName("metricBlue")
        right_limit_card.setMinimumWidth(210)
        right_limit_layout = QHBoxLayout(right_limit_card)
        right_limit_layout.setContentsMargins(8, 6, 8, 6)
        right_limit_layout.setSpacing(8)
        right_limit_layout.addWidget(QLabel("추천 개수"))
        self.category_limit_spin = QSpinBox()
        self.category_limit_spin.setObjectName("categoryCenterSpin")
        self.category_limit_spin.setRange(5, 100)
        self.category_limit_spin.setValue(10)
        self.category_limit_spin.setSuffix("개")
        self.category_limit_spin.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.category_limit_spin.setMinimumWidth(90)
        right_limit_layout.addWidget(self.category_limit_spin)
        right_top.addWidget(right_limit_card)

        self.golden_start_button = QPushButton("추천 실행")
        self.golden_start_button.setObjectName("categoryActionButton")
        self.golden_start_button.setMinimumWidth(120)
        self.golden_start_button.clicked.connect(self.start_category_golden_keyword_search)
        self.category_continue_button = QPushButton("이어서 실행")
        self.category_continue_button.setObjectName("categoryActionButton")
        self.category_continue_button.setMinimumWidth(130)
        self.category_continue_button.clicked.connect(lambda: self._expand_analysis_scope("category"))
        self.category_continue_button.setEnabled(False)
        self.category_save_button = QPushButton("저장")
        self.category_save_button.setObjectName("categoryActionButton")
        self.category_save_button.setMinimumWidth(90)
        self.category_save_button.setEnabled(False)
        self.category_save_button.clicked.connect(lambda: self.save_results_for_mode("category"))
        right_top.addWidget(self.golden_start_button)
        right_top.addWidget(self.category_continue_button)
        right_top.addWidget(self.category_save_button)
        right_layout.addLayout(right_top)

        self.category_sort_hint = QLabel("")
        self.category_sort_hint.setObjectName("sortHint")
        self.category_sort_hint.setVisible(False)
        right_layout.addWidget(self.category_sort_hint)

        self.category_table = QTableWidget()
        self._init_result_table(self.category_table)
        right_layout.addWidget(self.category_table)

        for w in [
            self.golden_category_combo,
            self.category_limit_spin,
            self.golden_start_button,
            self.category_continue_button,
            self.category_save_button,
            self.category_table
        ]:
            w.setEnabled(False)

        self.category_notice_card = QFrame()
        self.category_notice_card.setObjectName("categoryNoticeCard")
        notice_layout = QVBoxLayout(self.category_notice_card)
        notice_layout.setContentsMargins(16, 18, 16, 18)
        notice_layout.setSpacing(10)
        notice_title = QLabel("업데이트 예정")
        notice_title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        notice_title.setObjectName("categoryNoticeTitle")
        self.category_notice_text = QLabel(
            "카테고리 황금키워드 추천 기능은 현재 개선 중입니다.\n"
            "다음 업데이트에서 새 로직으로 제공됩니다."
        )
        self.category_notice_text.setWordWrap(True)
        self.category_notice_text.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.category_notice_text.setObjectName("categoryNoticeText")
        notice_layout.addStretch(1)
        notice_layout.addWidget(notice_title)
        notice_layout.addWidget(self.category_notice_text)
        notice_layout.addStretch(1)
        right_layout.addWidget(self.category_notice_card)
        split.addWidget(right_group, 1)

        root.addLayout(split)

        self.insight_group = QGroupBox("인사이트")
        insight_layout = QVBoxLayout(self.insight_group)
        self.insight_title_label = QLabel("")
        self.insight_title_label.setObjectName("insightTitle")
        self.insight_title_label.setVisible(False)
        insight_layout.addWidget(self.insight_title_label)

        ratio_row = QHBoxLayout()
        self.month_ratio_chart = InsightChartWidget("월별 검색 비율", chart_type="bar")
        self.weekday_ratio_chart = InsightChartWidget("요일별 검색 비율", chart_type="bar")
        self.age_ratio_chart = InsightChartWidget("연령별 검색 비율", chart_type="bar")
        ratio_row.addWidget(self.month_ratio_chart, 1)
        ratio_row.addWidget(self.weekday_ratio_chart, 1)
        ratio_row.addWidget(self.age_ratio_chart, 1)
        insight_layout.addLayout(ratio_row)
        root.addWidget(self.insight_group)
        root_widget.setStyleSheet("""
            QGroupBox { font-size: 14px; font-weight: 700; }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 25px;
                padding: 8px 15px;
                color: #185a3a;
                font-size: 18px;
                font-weight: 700;
                background-color: #ffffff;
                border: 2px solid #03c75a;
                border-radius: 8px;
            }
            QGroupBox#leftPanel, QGroupBox#rightPanel {
                background: #fbfdfc;
                border: 1px solid #cfe8d7;
                border-radius: 10px;
                margin-top: 8px;
            }
            QFrame#metricBlue { background: #eaf2ff; border: 1px solid #c7d8ff; border-radius: 8px; }
            QFrame#metricGreen { background: #e9f9f0; border: 1px solid #c8ebd7; border-radius: 8px; }
            QFrame#metricAmber { background: #fff6e8; border: 1px solid #f0d9ad; border-radius: 8px; }
            QFrame#metricBlue QLabel, QFrame#metricGreen QLabel, QFrame#metricAmber QLabel {
                background: transparent;
                color: #27403a;
                font-weight: 700;
                border: none;
            }
            QFrame#metricBlue QSpinBox, QFrame#metricGreen QSpinBox, QFrame#metricAmber QDoubleSpinBox {
                background: #ffffff;
                border: 1px solid #cfd8d3;
                border-radius: 6px;
            }
            QLabel {
                background: transparent;
                border: none;
            }
            QLabel#insightTitle {
                color: #1f5b40;
                font-weight: 800;
            }
            QLabel#sortHint {
                color: #31745a;
                font-size: 12px;
                font-weight: 700;
                padding-left: 2px;
            }
            QLabel#summaryHint {
                color: #4f6b5b;
                font-size: 12px;
                font-weight: 700;
                padding-left: 2px;
            }
            QFrame#categoryNoticeCard {
                background: #fef9eb;
                border: 1px solid #f3d48b;
                border-radius: 10px;
                min-height: 230px;
            }
            QLabel#categoryNoticeTitle {
                color: #9a6a00;
                font-size: 16px;
                font-weight: 800;
            }
            QLabel#categoryNoticeText {
                color: #6f4e00;
                font-size: 15px;
                font-weight: 700;
                line-height: 1.5;
            }
            QLabel#loadingSpinner {
                color: #1f6a49;
                font-size: 14px;
                font-weight: 800;
                min-width: 14px;
            }
            QLabel#loadingText {
                color: #2f5f4a;
                font-size: 12px;
                font-weight: 700;
            }
            QFrame#relatedGuideCard {
                background: #f8fcfa;
                border: 1px dashed #bfe5ce;
                border-radius: 8px;
            }
            QLabel#relatedGuideTitle {
                color: #1f5136;
                font-size: 15px;
                font-weight: 800;
            }
            QLabel#relatedGuideText {
                color: #4b6b5b;
                font-size: 13px;
                font-weight: 600;
                line-height: 1.45;
            }
            QPushButton { font-size: 13px; font-weight: 700; min-height: 34px; border-radius: 8px; }
            QPushButton#categoryActionButton {
                min-height: 36px;
                padding: 6px 14px;
                text-align: center;
            }
            QPushButton#uploadButton {
                background: #ecf3ff;
                color: #1f4f8a !important;
                border: 1px solid #c6d9f5;
                padding: 6px 12px;
            }
            QPushButton#uploadButton:hover {
                background: #e2eeff;
                color: #123f74 !important;
            }
            QPushButton#moreLinkButton {
                min-height: 34px;
                padding: 8px 18px;
                border: none;
                border-radius: 8px;
                background: #03c75a;
                color: #ffffff !important;
                font-size: 13px;
                font-weight: 700;
            }
            QPushButton#moreLinkButton:hover {
                background: #028a4a;
                color: #ffffff !important;
            }
            QPushButton#moreLinkButton:disabled {
                background: #dcdcdc;
                color: #666666 !important;
            }
            QLineEdit, QComboBox, QSpinBox, QDoubleSpinBox {
                min-height: 34px;
                font-size: 13px;
                background: #ffffff;
                border: 1px solid #cfd8d3;
                border-radius: 7px;
            }
            QProgressBar {
                min-height: 16px;
                max-height: 16px;
                border: 1px solid #cfd8d3;
                border-radius: 8px;
                background: #f3f7f5;
                text-align: center;
                color: #1f5136;
                font-size: 11px;
                font-weight: 700;
            }
            QProgressBar::chunk {
                border-radius: 7px;
                background: #03c75a;
            }
            QComboBox {
                padding-left: 10px;
                padding-right: 28px;
            }
            QComboBox#categoryCenterCombo {
                text-align: center;
                padding-left: 12px;
                padding-right: 30px;
            }
            QComboBox#categoryCenterCombo QLineEdit {
                qproperty-alignment: 'AlignCenter';
            }
            QComboBox::drop-down {
                border: none;
                width: 26px;
            }
            QComboBox::down-arrow {
                image: none;
                width: 0px;
                height: 0px;
                border-left: 5px solid transparent;
                border-right: 5px solid transparent;
                border-top: 7px solid #4f6f61;
                margin-right: 8px;
            }
            QSpinBox::up-button, QDoubleSpinBox::up-button {
                subcontrol-origin: border;
                subcontrol-position: top right;
                width: 18px;
                border-left: 1px solid #cfd8d3;
                border-bottom: 1px solid #cfd8d3;
                border-top-right-radius: 7px;
                background: #f2f7f4;
            }
            QSpinBox::down-button, QDoubleSpinBox::down-button {
                subcontrol-origin: border;
                subcontrol-position: bottom right;
                width: 18px;
                border-left: 1px solid #cfd8d3;
                border-bottom-right-radius: 7px;
                background: #f2f7f4;
            }
            QSpinBox::up-button:hover, QDoubleSpinBox::up-button:hover,
            QSpinBox::down-button:hover, QDoubleSpinBox::down-button:hover {
                background: #e4efe8;
            }
            QSpinBox::up-arrow, QDoubleSpinBox::up-arrow,
            QSpinBox::down-arrow, QDoubleSpinBox::down-arrow {
                image: none;
                width: 0px;
                height: 0px;
            }
            QSpinBox::up-arrow, QDoubleSpinBox::up-arrow {
                border-left: 5px solid transparent;
                border-right: 5px solid transparent;
                border-bottom: 7px solid #4f6f61;
            }
            QSpinBox::down-arrow, QDoubleSpinBox::down-arrow {
                border-left: 5px solid transparent;
                border-right: 5px solid transparent;
                border-top: 7px solid #4f6f61;
            }
            QSpinBox#categoryCenterSpin {
                qproperty-alignment: 'AlignCenter';
                padding-left: 8px;
                padding-right: 24px;
            }
        """)

        root_widget.setStyleSheet(self._golden_root_stylesheet("light"))
        parent_layout.addWidget(root_widget)

    def _init_result_table(self, table_widget):
        table_widget.setColumnCount(4)
        table_widget.setHorizontalHeaderLabels(
            ["키워드 ↕", "월 검색량 ↕", "월 블로그 발행량 ↕", "콘텐츠 포화 지수 ↕"]
        )
        header = table_widget.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        table_widget.setMinimumHeight(300)
        table_widget.setAlternatingRowColors(True)
        table_widget.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        table_widget.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        # 드래그로 연속 구간을 쉽게 선택할 수 있도록 행 단위 선택을 사용
        table_widget.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        table_widget.setVerticalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
        table_widget.setHorizontalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
        table_widget._copy_shortcut = QShortcut(QKeySequence("Ctrl+C"), table_widget)
        table_widget._copy_shortcut.activated.connect(
            lambda tw=table_widget: self.copy_selected_table_cells(tw)
        )
        table_widget.setSortingEnabled(True)
        table_widget.setStyleSheet(self._result_table_stylesheet(getattr(self, "current_theme_mode", "light")))
        self._update_blog_count_column_header()

    def _populate_related_guide_table(self):
        guide_table = getattr(self, "related_table_placeholder", None)
        if not isinstance(guide_table, QTableWidget):
            return

        lines = [
            "연관 키워드 분석 사용법",
            "1. 실행 전 설정: '월간 발행량 / 전체 발행량' 중 하나를 먼저 선택하세요.",
            "2. 설정 차이: 월간 발행량은 수집 정밀도가 높지만 시간이 더 걸리고, 전체 발행량은 더 빠르게 분석됩니다.",
            "3. 버튼 차이: '분석 실행'은 입력창의 여러 키워드를 일괄 분석하고, '단일 키워드'는 1개 키워드만 빠르게 분석합니다.",
            "4. 분석 시작: 키워드 입력 후 버튼을 누르거나, 여러 키워드는 '파일 업로드'를 사용하세요. (xlsx/csv, A열 기준)",
            "5. 진행 확인: 하단 진행 막대와 로그 탭에서 단계별 상태를 확인할 수 있습니다.",
            "6. 결과 표 항목: 키워드 / 월 검색량 / 월 블로그 발행량 / 콘텐츠 포화 지수",
            "7. 표 정렬: 각 열 제목의 ↕를 클릭하면 오름차순/내림차순으로 정렬됩니다.",
            "8. 추가 확장: '더 많은 연관 키워드 보기'를 누르면 현재 결과 기반으로 재분석합니다.",
            "9. 저장: 분석 완료 후 '저장' 버튼으로 현재 결과를 파일로 저장하세요.",
        ]

        guide_table.setSortingEnabled(False)
        guide_table.clearContents()
        guide_table.setRowCount(len(lines))
        guide_table.setSelectionMode(QAbstractItemView.SelectionMode.NoSelection)
        guide_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        guide_table.setFocusPolicy(Qt.FocusPolicy.NoFocus)

        gray_title = QColor("#5f6368")
        gray_body = QColor("#6f7680")

        for row, text in enumerate(lines):
            guide_table.setSpan(row, 0, 1, 4)
            item = QTableWidgetItem(text)
            item.setTextAlignment(int(Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignLeft))
            item.setForeground(gray_title if row == 0 else gray_body)
            if row == 0:
                f = item.font()
                f.setBold(True)
                f.setPointSize(max(11, f.pointSize()))
                item.setFont(f)
            guide_table.setItem(row, 0, item)
            guide_table.setRowHeight(row, 38 if row == 0 else 34)

    def copy_selected_table_cells(self, table_widget):
        indexes = table_widget.selectedIndexes()
        if not indexes:
            return
        rows = sorted({idx.row() for idx in indexes})
        cols = sorted({idx.column() for idx in indexes})
        selected = {(idx.row(), idx.column()): idx for idx in indexes}
        lines = []
        for r in rows:
            values = []
            for c in cols:
                idx = selected.get((r, c))
                values.append(idx.data() if idx is not None else "")
            lines.append("\t".join(values))
        clipboard = QApplication.clipboard()
        if clipboard is not None:
            clipboard.setText("\n".join(lines))

    def _tick_related_spinner(self):
        if not hasattr(self, "related_spinner"):
            return
        self.related_spinner.step()

    def _show_related_loading(self, text="작업 중...", indeterminate=True):
        self.related_loading_text.setText(text)
        if indeterminate:
            self.related_progress_bar.setRange(0, 0)
        else:
            self.related_progress_bar.setRange(0, max(1, self.related_progress_total or 100))
            self.related_progress_bar.setValue(0)
        self.related_loading_widget.setVisible(True)
        if not self.related_spinner_timer.isActive():
            self.related_spinner_timer.start()

    def _update_related_loading_progress(self, current, total, text=None):
        total = max(1, int(total))
        current = max(0, min(int(current), total))
        self.related_progress_total = total
        self.related_progress_bar.setRange(0, total)
        self.related_progress_bar.setValue(current)
        if text:
            self.related_loading_text.setText(text)

    def _hide_related_loading(self):
        if self.related_spinner_timer.isActive():
            self.related_spinner_timer.stop()
        self.related_loading_widget.setVisible(False)

    def _set_related_table_guide_visible(self, visible):
        if not hasattr(self, "related_table_stack"):
            return
        self.related_table_stack.setCurrentIndex(0 if visible else 1)

    def _get_blog_count_mode(self):
        if hasattr(self, "blog_count_mode_combo"):
            value = self.blog_count_mode_combo.currentData()
            if value in ("monthly", "total"):
                return value
        return "total" if str(getattr(self, "blog_count_mode", "monthly")).lower() == "total" else "monthly"

    def _get_blog_count_display_name(self):
        return "전체 발행량" if self._get_blog_count_mode() == "total" else "월간 발행량"

    def _update_blog_count_mode_hint(self):
        if not hasattr(self, "blog_count_mode_hint"):
            return
        self.blog_count_mode_hint.setText("")
        self.blog_count_mode_hint.setVisible(False)

    def _update_blog_count_column_header(self):
        title = f"{self._get_blog_count_display_name()} ↕"
        for table in [
            getattr(self, "related_table", None),
            getattr(self, "category_table", None),
            getattr(self, "related_table_placeholder", None),
        ]:
            if not table:
                continue
            header_item = table.horizontalHeaderItem(2)
            if header_item:
                header_item.setText(title)

    def on_blog_count_mode_changed(self):
        self.blog_count_mode = self._get_blog_count_mode()
        self.settings.set_blog_count_mode(self.blog_count_mode)
        self._update_blog_count_column_header()
        self._update_blog_count_mode_hint()

    def _expand_analysis_scope(self, mode):
        batch_size = int(self.related_batch_size if mode == "related" else self.category_limit_spin.value())
        self.analysis_offset[mode] = int(self.analysis_offset.get(mode, 0)) + batch_size

        keyword = self.last_analysis_keyword.get(mode, "").strip()
        if not keyword:
            if mode == "related":
                keyword = self.related_keyword_input.text().strip()
            else:
                selected = self.golden_category_combo.currentText().strip()
                keyword = selected
        self._start_golden_analysis(
            mode,
            keyword,
            self.category_seed_map.get(keyword, []),
            offset=int(self.analysis_offset.get(mode, 0))
        )

    def _start_golden_analysis(
        self,
        analysis_type,
        keyword,
        category_seeds=None,
        offset=0,
        keep_existing=False,
        single_keyword_mode=False
    ):
        if (self.golden_keyword_thread and self.golden_keyword_thread.isRunning()) or (
            self.file_keyword_thread and self.file_keyword_thread.isRunning()
        ):
            QMessageBox.information(self, "진행 중", "다른 분석이 이미 실행 중입니다.")
            return
        if not keyword:
            QMessageBox.warning(self, "입력 오류", "키워드 또는 카테고리를 선택해 주세요.")
            return
        try:
            credentials, _ = load_api_credentials_from_file()
        except FileNotFoundError as e:
            QMessageBox.warning(self, "설정 필요", "분석에 필요한 API 설정을 확인해 주세요.")
            return
        except Exception as e:
            QMessageBox.warning(self, "API 키 오류", str(e))
            return

        API_USAGE_REPORTER.configure(
            machine_id=get_machine_id(),
            webhook_url=credentials.get("usage_webhook_url", ""),
            webhook_token=credentials.get("usage_webhook_token", ""),
        )

        self.last_analysis_keyword[analysis_type] = keyword
        self.analysis_offset[analysis_type] = int(offset)
        self.analysis_keep_existing[analysis_type] = bool(keep_existing)
        self.current_analysis_mode = analysis_type
        self.related_single_mode = bool(single_keyword_mode) if analysis_type == "related" else False
        if analysis_type == "related":
            if not keep_existing:
                self.related_keyword_results = []
                self.related_table.setRowCount(0)
                self.related_save_button.setEnabled(False)
                self._set_related_table_guide_visible(True)
            self.related_keyword_button.setEnabled(False)
            self.related_single_button.setEnabled(False)
            self.related_upload_button.setEnabled(False)
            self.related_more_button.setEnabled(False)
            self.related_progress_total = 0
            loading_text = "작업 중... 단일 키워드 지표를 계산하고 있습니다." if self.related_single_mode else "작업 중... 키워드를 수집하고 있습니다."
            self._show_related_loading(loading_text, indeterminate=True)
            self.insight_title_label.setText("키워드 인사이트를 불러오는 중...")
            self.month_ratio_chart.set_data([], [])
            self.weekday_ratio_chart.set_data([], [])
            self.age_ratio_chart.set_data([], [])
        else:
            self.category_keyword_results = []
            self.category_table.setRowCount(0)
            self.category_save_button.setEnabled(False)
            self.golden_start_button.setEnabled(False)
            self.category_continue_button.setEnabled(False)
        self.status_bar.showMessage("키워드 분석 중...")
        selected_mode = self._get_blog_count_mode()
        if analysis_type == "related" and self.related_single_mode:
            self.on_golden_keyword_log("단일 키워드 모드로 실행합니다.")

        self.golden_keyword_thread = GoldenKeywordThread(
            analysis_type=analysis_type,
            keyword=keyword,
            limit=int(100000 if analysis_type == "related" else self.category_limit_spin.value()),
            offset=int(self.analysis_offset.get(analysis_type, 0)),
            credentials=credentials,
            category_seeds=category_seeds or [],
            blog_count_mode=selected_mode,
            single_keyword_mode=self.related_single_mode
        )
        self.golden_keyword_thread.log.connect(self.on_golden_keyword_log)
        self.golden_keyword_thread.insight.connect(self.on_keyword_insight)
        self.golden_keyword_thread.finished.connect(
            lambda rows, mode=analysis_type: self.on_golden_keyword_finished(rows, mode)
        )
        self.golden_keyword_thread.error.connect(self.on_golden_keyword_error)
        self.golden_keyword_thread.start()

    def start_related_keyword_analysis(self):
        self._start_golden_analysis(
            "related",
            self.related_keyword_input.text().strip(),
            offset=0,
            keep_existing=False,
            single_keyword_mode=False
        )

    def start_single_keyword_analysis(self):
        self._start_golden_analysis(
            "related",
            self.related_keyword_input.text().strip(),
            offset=0,
            keep_existing=False,
            single_keyword_mode=True
        )

    def start_related_file_analysis(self):
        if (self.golden_keyword_thread and self.golden_keyword_thread.isRunning()) or (
            self.file_keyword_thread and self.file_keyword_thread.isRunning()
        ):
            QMessageBox.information(self, "진행 중", "다른 분석이 이미 실행 중입니다.")
            return

        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "키워드 파일 선택",
            "",
            "Excel/CSV Files (*.xlsx *.csv)"
        )
        if not file_path:
            return

        ext = os.path.splitext(file_path)[1].lower()
        if ext not in [".xlsx", ".csv"]:
            QMessageBox.warning(self, "파일 형식 오류", "xlsx, csv 파일만 업로드할 수 있습니다.")
            return

        try:
            credentials, _ = load_api_credentials_from_file()
        except FileNotFoundError:
            QMessageBox.warning(self, "설정 필요", "분석에 필요한 API 설정을 확인해 주세요.")
            return
        except Exception as e:
            QMessageBox.warning(self, "API 키 오류", str(e))
            return

        API_USAGE_REPORTER.configure(
            machine_id=get_machine_id(),
            webhook_url=credentials.get("usage_webhook_url", ""),
            webhook_token=credentials.get("usage_webhook_token", ""),
        )

        self.current_analysis_mode = "related"
        self.related_keyword_results = []
        self.related_table.setRowCount(0)
        self._set_related_table_guide_visible(True)
        self.related_save_button.setEnabled(False)
        self.related_more_button.setEnabled(False)
        self.related_keyword_button.setEnabled(False)
        self.related_upload_button.setEnabled(False)
        self.related_progress_total = 0
        self._show_related_loading("작업 중... 업로드 파일 분석 준비 중", indeterminate=True)
        self.status_bar.showMessage("업로드 파일 분석 중...")
        selected_mode = self._get_blog_count_mode()

        self.file_keyword_thread = FileKeywordAnalysisThread(
            file_path,
            credentials,
            blog_count_mode=selected_mode
        )
        self.file_keyword_thread.log.connect(self.on_golden_keyword_log)
        self.file_keyword_thread.progress.connect(self.on_related_file_progress)
        self.file_keyword_thread.finished.connect(self.on_related_file_finished)
        self.file_keyword_thread.error.connect(self.on_related_file_error)
        self.file_keyword_thread.start()

    def on_related_file_progress(self, current, total):
        self._update_related_loading_progress(
            current,
            total,
            f"작업 중... {current}/{total}"
        )

    def on_related_file_finished(self, results, output_path):
        self._hide_related_loading()
        self.current_analysis_mode = ""
        self.related_keyword_button.setEnabled(True)
        self.related_single_button.setEnabled(True)
        self.related_upload_button.setEnabled(True)
        self.related_save_button.setEnabled(True)
        self.related_more_button.setEnabled(False)
        self.related_keyword_results = list(results or [])
        self._set_related_table_guide_visible(not bool(self.related_keyword_results))
        self.apply_filters_for_mode("related")
        self.status_bar.showMessage(f"파일 분석 완료 ({len(self.related_keyword_results)}개)")
        QMessageBox.information(
            self,
            "파일 분석 완료",
            f"분석 결과 파일 저장 완료:\n{output_path}\n\n"
            "A열 키워드 기준으로\n"
            f"B열 월 검색량, C열 {self._get_blog_count_display_name()}, D열 콘텐츠 포화 지수로 생성했습니다."
        )

    def on_related_file_error(self, error_message):
        self._hide_related_loading()
        self.current_analysis_mode = ""
        self.related_keyword_button.setEnabled(True)
        self.related_single_button.setEnabled(True)
        self.related_upload_button.setEnabled(True)
        self.related_save_button.setEnabled(bool(self.related_keyword_results))
        self.related_more_button.setEnabled(bool(self.last_analysis_keyword.get("related", "").strip()))
        self.status_bar.showMessage("파일 분석 실패")
        QMessageBox.critical(self, "파일 분석 오류", str(error_message))

    def start_related_keyword_analysis_more(self):
        keyword = self.last_analysis_keyword.get("related", "").strip() or self.related_keyword_input.text().strip()
        if not keyword:
            QMessageBox.warning(self, "입력 오류", "먼저 연관 키워드 분석을 실행해 주세요.")
            return

        rows = [r for r in (self.related_keyword_results or []) if str(r.get("keyword", "")).strip()]
        if not rows:
            QMessageBox.warning(self, "안내", "표 결과가 없어 확장 분석을 진행할 수 없습니다.")
            return

        sorted_rows = sorted(
            rows,
            key=lambda r: int(r.get("monthly_total_search", 0)),
            reverse=True
        )

        # 표 결과를 검색량 내림차순으로 모두 순회하되, 1,000 초과 키워드만 재귀 확장 대상으로 사용
        expand_seeds = []
        seen_seed_keys = set()
        for row in sorted_rows:
            monthly = int(row.get("monthly_total_search", 0))
            if monthly <= 1000:
                continue
            seed_text = str(row.get("keyword", "")).replace("+", " ").strip()
            seed_key = seed_text.replace(" ", "").lower()
            if not seed_text or seed_key in seen_seed_keys:
                continue
            seen_seed_keys.add(seed_key)
            expand_seeds.append(seed_text)

        if not expand_seeds:
            QMessageBox.information(
                self,
                "안내",
                "현재 표 결과에서 월 검색량 1,000 초과 키워드가 없어\n"
                "'더 많은 연관 키워드 보기' 재귀 확장을 적용할 수 없습니다."
            )
            return

        self.status_bar.showMessage(
            f"확장 루트 {len(expand_seeds)}개 선택 (월 검색량 1,000 초과)"
        )
        self._start_golden_analysis(
            "related",
            keyword,
            category_seeds=expand_seeds,
            offset=0,
            keep_existing=True
        )

    def start_category_golden_keyword_search(self):
        QMessageBox.information(self, "안내", "카테고리 황금키워드 추천 기능은 업데이트 예정입니다.")

    def on_golden_keyword_log(self, message):
        self.status_bar.showMessage(sanitize_display_text(message))
        text = sanitize_display_text(message)
        if self.current_analysis_mode != "related":
            return
        m_batch = re.search(r"(?:후보|수집 키워드)\s+(\d+)개\s+중\s+이번\s*배치\s+(\d+)개|(?:후보|수집 키워드)\s+(\d+)개\s+중\s+배치\s+(\d+)개", text)
        if m_batch:
            batch = int(m_batch.group(2) or m_batch.group(4))
            self.related_progress_total = max(1, batch)
            self._update_related_loading_progress(
                0,
                self.related_progress_total,
                f"작업 중... 0/{self.related_progress_total}"
            )
            return
        m_step = re.search(r"\[(\d+)/(\d+)\]", text)
        if m_step:
            cur = int(m_step.group(1))
            total = int(m_step.group(2))
            # 1단계: 후보 검색량 조회(0~75%)
            pct = int((cur / max(1, total)) * 75)
            self.related_progress_bar.setRange(0, 100)
            self.related_progress_bar.setValue(pct)
            self.related_loading_text.setText(f"1단계 검색량 조회 중... {cur}/{total}")
            return
        m_eval = re.search(r"\[EVAL\s+(\d+)/(\d+)\]", text)
        if m_eval:
            cur = int(m_eval.group(1))
            total = int(m_eval.group(2))
            # 2단계: 블로그/포화지수 계산(75~100%)
            pct = 75 + int((cur / max(1, total)) * 25)
            self.related_progress_bar.setRange(0, 100)
            self.related_progress_bar.setValue(min(100, pct))
            self.related_loading_text.setText(f"2단계 콘텐츠 지표 계산 중... {cur}/{total}")
            return
        if "단일 키워드" in text:
            self._show_related_loading("작업 중... 단일 키워드 지표 계산 중", indeterminate=True)
            return
        if "키워드 수집 완료" in text:
            self._show_related_loading("작업 중... 검색량 분석 준비 중", indeterminate=True)

    def on_keyword_insight(self, insight):
        if not insight:
            self.insight_title_label.setText("인사이트 데이터를 불러오지 못했습니다.")
            return

        self.insight_title_label.setText("인사이트")
        month_ratio = insight.get("month_ratio", [])
        weekday_ratio = insight.get("weekday_ratio", [])
        age_ratio = insight.get("age_ratio", [])
        self.month_ratio_chart.set_data(
            [str(x.get("label", "")) for x in month_ratio],
            [float(x.get("value", 0.0)) for x in month_ratio]
        )
        self.weekday_ratio_chart.set_data(
            [str(x.get("label", "")) for x in weekday_ratio],
            [float(x.get("value", 0.0)) for x in weekday_ratio]
        )
        self.age_ratio_chart.set_data(
            [str(x.get("label", "")) for x in age_ratio],
            [float(x.get("value", 0.0)) for x in age_ratio]
        )

    def on_golden_keyword_finished(self, results, analysis_type):
        self.related_keyword_button.setEnabled(True)
        self.related_upload_button.setEnabled(True)
        self.related_single_button.setEnabled(True)
        self.golden_start_button.setEnabled(False)
        if analysis_type == "related":
            self._hide_related_loading()
        self.current_analysis_mode = ""
        keep_existing = bool(self.analysis_keep_existing.get(analysis_type, False))
        if analysis_type == "related":
            incoming = results or []
            if keep_existing and self.related_keyword_results:
                existing_keys = {
                    str(r.get("keyword", "")).replace("+", " ").strip().lower()
                    for r in self.related_keyword_results
                }
                for row in incoming:
                    key = str(row.get("keyword", "")).replace("+", " ").strip().lower()
                    if key and key not in existing_keys:
                        row["keyword"] = str(row.get("keyword", "")).replace("+", " ").strip()
                        self.related_keyword_results.append(row)
                        existing_keys.add(key)
            else:
                for row in incoming:
                    row["keyword"] = str(row.get("keyword", "")).replace("+", " ").strip()
                self.related_keyword_results = incoming
            current_rows = self.related_keyword_results
        else:
            self.category_keyword_results = results or []
            current_rows = self.category_keyword_results

        if not current_rows:
            self.status_bar.showMessage("분석 완료 (결과 없음)")
            if analysis_type == "related":
                self.related_more_button.setEnabled(not self.related_single_mode)
                self.related_keyword_results = []
                self.apply_filters_for_mode("related")
                QMessageBox.information(
                    self,
                    "안내",
                    "수집된 연관/자동완성 키워드가 충분하지 않아 결과가 비었습니다.\n"
                    "키워드 띄어쓰기 없이 다시 시도하거나, '더 많은 연관 키워드 보기'를 눌러 추가 수집을 진행해 주세요."
                )
                self.related_single_mode = False
            else:
                self.category_continue_button.setEnabled(True)
            return

        if analysis_type == "related":
            self.related_save_button.setEnabled(True)
            self.related_more_button.setEnabled(not self.related_single_mode)
            self.related_single_mode = False
        else:
            self.category_save_button.setEnabled(False)
            self.category_continue_button.setEnabled(False)
        self.apply_filters_for_mode(analysis_type)
        mode_name = "연관 키워드 분석" if analysis_type == "related" else "카테고리 황금키워드 추천"
        self.status_bar.showMessage(f"{mode_name} 완료 ({len(current_rows)}개)")

    def on_golden_keyword_error(self, error_message):
        self.related_keyword_button.setEnabled(True)
        self.related_upload_button.setEnabled(True)
        self.related_single_button.setEnabled(True)
        self.golden_start_button.setEnabled(False)
        self._hide_related_loading()
        self.current_analysis_mode = ""
        self.related_more_button.setEnabled(bool(self.last_analysis_keyword.get("related", "").strip()))
        self.related_single_mode = False
        self.category_continue_button.setEnabled(False)
        self.related_save_button.setEnabled(bool(self.related_keyword_results))
        self.category_save_button.setEnabled(False)
        self.status_bar.showMessage("분석 실패")
        if "(429)" in str(error_message) or "too many" in str(error_message).lower():
            QMessageBox.warning(
                self,
                "요청 제한 안내",
                "API 요청 제한(429)에 도달했습니다.\n"
                "1-2분 후 다시 시도해 주세요."
            )
            return
        QMessageBox.critical(self, "키워드 분석 오류", error_message)

    def render_results_for_mode(self, mode, results):
        table = self.related_table if mode == "related" else self.category_table
        if mode == "related":
            self._set_related_table_guide_visible(len(results) == 0)
        table.setSortingEnabled(False)
        table.setRowCount(len(results))
        for idx, row in enumerate(results, start=1):
            saturation = float(row.get("content_saturation_index", 0.0))
            monthly = int(row.get("monthly_total_search", 0))
            blog_count = int(row.get("blog_document_count", 0))
            items = [
                QTableWidgetItem(str(row["keyword"]).replace("+", " ").strip()),
                SortableNumericItem(f"{monthly:,}", monthly),
                SortableNumericItem(f"{blog_count:,}", blog_count),
                SortableNumericItem(f"{saturation:.2f}%", saturation),
            ]
            for col, item in enumerate(items):
                align = Qt.AlignmentFlag.AlignCenter
                if col == 0:
                    align = Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignLeft
                if col == 3 and saturation < 10.0:
                    item.setForeground(QColor("#d93025"))
                item.setTextAlignment(int(align))
                table.setItem(idx - 1, col, item)

        table.resizeRowsToContents()
        table.setSortingEnabled(True)

    def _get_sorted_filtered_results_for_mode(self, mode):
        return list(self.related_keyword_results if mode == "related" else self.category_keyword_results)

    def _update_related_summary(self):
        return

    def apply_filters_for_mode(self, mode):
        filtered = self._get_sorted_filtered_results_for_mode(mode)
        save_button = self.related_save_button if mode == "related" else self.category_save_button
        if mode == "related":
            self._update_related_summary()
        self.render_results_for_mode(mode, filtered)
        save_button.setEnabled(bool(filtered))

    def save_results_for_mode(self, mode):
        filtered_results = self._get_sorted_filtered_results_for_mode(mode)
        if not filtered_results:
            QMessageBox.warning(self, "데이터 없음", "저장할 황금 키워드 결과가 없습니다.")
            return

        save_dir = self.save_path_input.text().strip() or os.getcwd()
        os.makedirs(save_dir, exist_ok=True)
        filename = "related_keywords.txt" if mode == "related" else "category_golden_keywords.txt"
        save_path = os.path.join(save_dir, filename)

        top_keywords = [str(row["keyword"]).replace("+", " ").strip() for row in filtered_results]
        with open(save_path, "w", encoding="utf-8") as f:
            f.write("\n".join(top_keywords))

        self.update_progress("전체", f"키워드 {len(top_keywords)}개 저장: {save_path}")
        QMessageBox.information(self, "저장 완료", f"저장 위치:\n{save_path}")

    def start_search(self):
        """검색 시작"""
        keywords_text = self.search_input.toPlainText().strip()
        if not keywords_text:
            QMessageBox.warning(self, "입력 오류", "키워드를 입력해 주세요.")
            return
        
        keywords = [keyword.strip() for keyword in keywords_text.split('\n') if keyword.strip()]
        if not keywords:
            QMessageBox.warning(self, "입력 오류", "유효한 키워드를 입력해 주세요.")
            return
        
        save_dir = self.save_path_input.text().strip()
        if not save_dir:
            QMessageBox.warning(self, "경로 오류", "저장할 폴더를 선택해 주세요.")
            return
        
        if not os.path.exists(save_dir):
            try:
                os.makedirs(save_dir, exist_ok=True)
            except Exception as e:
                QMessageBox.critical(self, "폴더 생성 오류", f"폴더 생성 실패:\n{str(e)}")
                return
        
        # if not self.driver check removed as threads handle their own drivers

        
        # comment removed (encoding issue)
        self.start_button.setEnabled(False)
        self.pause_button.setEnabled(True)
        self.pause_button.setText("일시정지")
        self.stop_button.setEnabled(True)
        self.search_input.setEnabled(False)
        # comment removed (encoding issue)
        self.status_bar.showMessage("키워드 추출 중...")
        
        self.update_progress(f"키워드 추출 작업 시작 (총 {len(keywords)}개)")
        self.update_progress(f"입력 키워드: {', '.join(keywords)}")
        self.update_progress(f"저장 폴더: {save_dir}")
        self.update_progress("병렬 처리 모드로 실행합니다.")
        
        # comment removed (encoding issue)
        while self.progress_tabs.count() > 1:
            self.progress_tabs.removeTab(1)
            
        self.log_widgets = {"전체": self.total_log_text}
        self.total_log_text.clear()
        
        # comment removed (encoding issue)
        self.active_threads = []
        self.completed_threads = 0
        self.total_threads = len(keywords)
        self.stop_requested = False
        
        # comment removed (encoding issue)
        for keyword in keywords:
            log_widget = SmartProgressTextEdit(min_height=100, max_height=800)
            log_widget.setReadOnly(True)
            self.progress_tabs.addTab(log_widget, keyword)
            self.log_widgets[keyword] = log_widget
            
            # comment removed (encoding issue)
            if keywords.index(keyword) == 0:
                self.progress_tabs.setCurrentIndex(1)
        
        for i, keyword in enumerate(keywords):
            thread = ParallelKeywordThread(keyword, save_dir, True)
            thread.finished.connect(self.on_thread_finished)
            thread.error.connect(self.on_thread_error)
            thread.log.connect(self.update_progress)
            
            self.active_threads.append(thread)
            thread.start()
            self.update_progress(keyword, f"'{keyword}' 작업 시작...")

    def on_thread_finished(self, save_path):
        """Description"""
        if save_path:
            self.update_progress("전체", f"저장 완료: {save_path}")
        self.completed_threads += 1
        self.check_all_threads_finished()

    def on_thread_error(self, error_msg):
        """스레드 에러 처리"""
        self.update_progress("전체", error_msg)
        self.completed_threads += 1
        self.check_all_threads_finished()
        
    def check_all_threads_finished(self):
        """Description"""
        if self.completed_threads >= self.total_threads:
            if self.stop_requested:
                self.search_finished("중단 요청된 작업이 모두 종료되었습니다.")
            else:
                self.search_finished("모든 작업이 완료되었습니다.")

    def search_finished(self, message):
        """검색 완료 처리"""
        self.reset_ui_state()
        self.update_progress("전체", f"완료: {message}")
        self.status_bar.showMessage("완료")
        QMessageBox.information(self, "완료", message)
        self.active_threads = []
        self.total_threads = 0
        self.completed_threads = 0
        self.stop_requested = False

    def stop_search(self):
        """검색 중단"""
        if not self.active_threads:
            return
            
        self.update_progress("전체", "작업 중단을 요청했습니다...")
        self.stop_requested = True
        
        for thread in self.active_threads:
            if thread and thread.isRunning():
                thread.stop()
                
        self.stop_button.setEnabled(False)
        self.pause_button.setEnabled(False)
        self.update_progress("전체", "중단 신호를 전송했습니다. 실행 중인 스레드가 순차 종료됩니다.")

    def search_error(self, error_msg):
        """검색 오류 처리"""
        self.reset_ui_state()
        self.status_bar.showMessage("추출 실패")
        QMessageBox.critical(self, "추출 오류", f"연관키워드 추출 중 오류가 발생했습니다:\n{error_msg}")

    def update_progress(self, keyword_or_msg, message=None):
        """Description"""
        current_time = datetime.now().strftime('%H:%M:%S')
        
        # comment removed (encoding issue)
        if message is None:
            # comment removed (encoding issue)
            target_keyword = "전체"
            msg_content = keyword_or_msg
        else:
            # comment removed (encoding issue)
            target_keyword = keyword_or_msg
            msg_content = message

        target_keyword = sanitize_display_text(target_keyword)
        msg_content = sanitize_display_text(msg_content)
        formatted_message = f"[{current_time}] {msg_content}"
        
        # comment removed (encoding issue)
        if target_keyword in self.log_widgets:
            widget = self.log_widgets[target_keyword]
            if hasattr(widget, 'append_with_smart_scroll'):
                widget.append_with_smart_scroll(formatted_message)
            else:
                widget.append(formatted_message)
        
        # comment removed (encoding issue)
        if target_keyword != "전체":
            formatted_total_msg = f"[{current_time}] [{target_keyword}] {msg_content}"
            if hasattr(self.total_log_text, 'append_with_smart_scroll'):
                self.total_log_text.append_with_smart_scroll(formatted_total_msg)
            else:
                self.total_log_text.append(formatted_total_msg)

    def pause_resume_search(self):
        """Description"""
        # comment removed (encoding issue)
        # comment removed (encoding issue)
        # comment removed (encoding issue)
        pass

    def on_search_paused(self):
        """Description"""
        self.pause_button.setText("재개")
        self.status_bar.showMessage("키워드 추출이 일시정지되었습니다.")
    
    def on_search_resumed(self):
        """Description"""
        self.pause_button.setText("일시정지")
        self.status_bar.showMessage("키워드 추출 중...")

    def reset_ui_state(self):
        """Description"""
        self.start_button.setEnabled(True)
        self.pause_button.setEnabled(False)
        self.pause_button.setText("일시정지")
        self.stop_button.setEnabled(False)
        self.search_input.setEnabled(True)

    def closeEvent(self, event):
        """Handle window close event and cleanup."""
        global _crash_save_enabled
        
        # comment removed (encoding issue)
        if hasattr(self, 'active_threads') and self.active_threads:
            running_threads = [t for t in self.active_threads if t.isRunning()]
            
            if running_threads:
                reply = QMessageBox.question(
                    self, "확인",
                    f"현재 {len(running_threads)}개의 키워드 추출 작업이 진행 중입니다.\n\n작업을 중단하고 종료하시겠습니까?",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                    QMessageBox.StandardButton.Yes
                )
                
                if reply == QMessageBox.StandardButton.Yes:
                    self.update_progress("전체", "프로그램 종료를 위해 작업을 정리하고 있습니다...")
                    
                    # comment removed (encoding issue)
                    for thread in running_threads:
                        thread.stop()
                        thread.wait(1000)
                    
                    self.active_threads = []
                else:
                    event.ignore()
                    return

        if hasattr(self, "golden_keyword_thread") and self.golden_keyword_thread:
            if self.golden_keyword_thread.isRunning():
                self.golden_keyword_thread.quit()
                self.golden_keyword_thread.wait(1000)
        if hasattr(self, "file_keyword_thread") and self.file_keyword_thread:
            if self.file_keyword_thread.isRunning():
                self.file_keyword_thread.quit()
                self.file_keyword_thread.wait(1000)
        
        _crash_save_enabled = False
        
        if self.driver:
            pass
        
        event.accept()


def main():
    if sys.platform == 'win32':
        import ctypes
        ctypes.windll.kernel32.SetConsoleOutputCP(65001)
        ctypes.windll.kernel32.SetConsoleCP(65001)

    if not verify_machine_id_guard():
        sys.exit(1)
    
    app = QApplication(sys.argv)
    
    # comment removed (encoding issue)
    icon_path = get_icon_path()
    if icon_path:
        app.setWindowIcon(QIcon(icon_path))
        safe_print(f"아이콘 설정 완료: {icon_path}")
    else:
        safe_print("아이콘 파일을 찾을 수 없습니다.")
    
    app.setApplicationName("네이버 연관키워드 추출기")
    
    # comment removed (encoding issue)
    machine_id = get_machine_id()
    safe_print(f"Machine ID: {machine_id}")
    
    # comment removed (encoding issue)
    # comment removed (encoding issue)
    expiry_date_str = check_license_from_sheet(machine_id)
    
    if expiry_date_str:
        try:
            # comment removed (encoding issue)
            expiry_date = datetime.strptime(expiry_date_str, '%Y-%m-%d')
            current_date = datetime.now()
            
            # comment removed (encoding issue)
            if current_date > expiry_date + pd.Timedelta(days=1):
                safe_print(f"라이선스 만료: {expiry_date_str}")
                app_dummy = QApplication.instance() or QApplication(sys.argv)
                
                # comment removed (encoding issue)
                dialog = ExpiredDialog(expiry_date_str)
                dialog.exec()
                sys.exit(0)
            
            # comment removed (encoding issue)
            safe_print(f"라이선스 확인 완료: {expiry_date_str}")
            window = KeywordExtractorMainWindow()
            window.usage_label.setText(f"사용 기간: {expiry_date_str}")
            window.usage_label.setStyleSheet(f"font-size: 13px; font-weight: bold; color: {NAVER_GREEN};")
            window.showMaximized()
            
            try:
                sys.exit(app.exec())
            except Exception as e:
                emergency_save_data()
                raise
                
        except ValueError:
            # comment removed (encoding issue)
            safe_print(f"라이선스 날짜 형식 확인 필요: {expiry_date_str}")
            window = KeywordExtractorMainWindow()
            window.usage_label.setText(f"사용 기간: {expiry_date_str}")
            window.showMaximized()
            sys.exit(app.exec())
            
    else:
        # comment removed (encoding issue)
        safe_print("미등록 기기 - 실행 차단")
        dialog = UnregisteredDialog(machine_id)
        dialog.exec()
        sys.exit(0)

if __name__ == "__main__":
    main() 





