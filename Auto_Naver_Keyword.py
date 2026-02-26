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
from typing import Any, Optional
from pathlib import Path
from datetime import datetime
import time
import urllib.parse
from urllib.parse import quote, unquote
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.remote.webdriver import WebDriver
from webdriver_manager.chrome import ChromeDriverManager

# comment cleaned (encoding issue)
try:
    from bs4 import BeautifulSoup
    BEAUTIFULSOUP_AVAILABLE = True
except ImportError:
    print("BeautifulSoup가 설치되지 않았습니다: pip install beautifulsoup4")
    BEAUTIFULSOUP_AVAILABLE = False

# comment cleaned (encoding issue)

# comment cleaned (encoding issue)
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
        # comment cleaned (encoding issue)
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
    QTabWidget, QTabBar, QSpinBox, QDoubleSpinBox, QTableWidget, QTableWidgetItem, QHeaderView
)
from PyQt6.QtCore import QThread, pyqtSignal, Qt, QEvent, QSettings, QDir, QTimer, QUrl, QRect
from PyQt6.QtGui import (
    QPixmap, QKeySequence, QFont, QTransform, QIcon, QShortcut,
    QPainter, QColor, QDesktopServices, QCursor
)

import pandas as pd

# comment cleaned (encoding issue)
try:
    from bs4 import BeautifulSoup
    BEAUTIFULSOUP_AVAILABLE = True
except ImportError:
    print("BeautifulSoup가 설치되지 않았습니다: pip install beautifulsoup4")
    BEAUTIFULSOUP_AVAILABLE = False

# comment cleaned (encoding issue)
OPENAI_AVAILABLE = False

# comment cleaned (encoding issue)
NAVER_GREEN = "#03c75a"              # 메인 이그린
NAVER_GREEN_DARK = "#028a4a"         # 진한 그린 (hover)
NAVER_GREEN_LIGHT = "#e8f5f0"        # 한 그린 (배경)
NAVER_GREEN_ULTRA_LIGHT = "#f0faf7"  # 매우 한 그린 (체 배경)
WHITE_COLOR = "#ffffff"               # 백
TEXT_PRIMARY = "#212529"             # 진한 스
TEXT_SECONDARY = "#6c757d"           # 보조 스
BACKGROUND_MAIN = "#f0faf7"          # 메인 배경 (한 그린)
BACKGROUND_CARD = "#ffffff"          # 카드 배경
BORDER_COLOR = "#d4edda"             # 한 그린 두
BORDER_FOCUS = "#03c75a"             # 커두
PLACEHOLDER_COLOR = "#8a8a8a"        # 리시

# comment cleaned (encoding issue)
_current_window = None
_crash_save_enabled = True
MACHINE_ID_GUARD_HASH = "39185df9b843b979ce5f989e26ae7e692407c83a5ea380e8dbf7c986e444e375"
MACHINE_ID_APPROVAL_FILE = "machine_id_change_approval.txt"
MACHINE_ID_APPROVAL_TOKEN = "I_APPROVE_MACHINE_ID_CHANGE"
MACHINE_ID_PREFIX = "Keyword-"
# comment cleaned (encoding issue)
def get_icon_path():
    """Text cleaned due to encoding issue."""
    try:
        # comment cleaned (encoding issue)
        meipass_dir = getattr(sys, "_MEIPASS", None)
        if meipass_dir:
            # comment cleaned (encoding issue)
            icon_path = os.path.join(meipass_dir, 'auto_naver.ico')
            if os.path.exists(icon_path):
                return icon_path
        
        # comment cleaned (encoding issue)
        if getattr(sys, 'frozen', False):
            # comment cleaned (encoding issue)
            exe_dir = os.path.dirname(sys.executable)
            icon_path = os.path.join(exe_dir, 'auto_naver.ico')
            if os.path.exists(icon_path):
                return icon_path
        
        # comment cleaned (encoding issue)
        script_dir = os.path.dirname(os.path.abspath(__file__))
        icon_path = os.path.join(script_dir, 'auto_naver.ico')
        if os.path.exists(icon_path):
            return icon_path
        
        # comment cleaned (encoding issue)
        assets_icon = os.path.join(script_dir, 'assets', 'auto_naver.ico')
        if os.path.exists(assets_icon):
            return assets_icon
        
        # comment cleaned (encoding issue)
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
        except Exception as e:
            safe_print(f"api_keys.json 읽기 실패(내장 키 사용): {e}")

    missing = []
    for key in required_keys:
        if not credentials.get(key):
            missing.append(key)

    if missing:
        raise ValueError("필수 API 키 누락: " + ", ".join(missing))
    return credentials, api_file


def _machine_id_cache_paths():
    paths = [Path.home() / ".auto_naver_machine_id.txt"]
    appdata = os.getenv("APPDATA", "").strip()
    if appdata:
        paths.append(Path(appdata) / "AutoNaverKeyword" / "machine_id.txt")
    return paths


def _load_persisted_machine_id():
    for path in _machine_id_cache_paths():
        try:
            if path.exists():
                cached = path.read_text(encoding="utf-8-sig").strip()
                if cached.startswith("MID-"):
                    return cached
        except Exception:
            pass
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\AutoNaverKeyword") as key:
            cached, _ = winreg.QueryValueEx(key, "MachineId")
            cached = str(cached).strip()
            if cached.startswith("MID-"):
                return cached
    except Exception:
        pass
    return None


def _save_persisted_machine_id(machine_id):
    for path in _machine_id_cache_paths():
        try:
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text(machine_id, encoding="utf-8")
        except Exception:
            pass
    try:
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, r"Software\AutoNaverKeyword") as key:
            winreg.SetValueEx(key, "MachineId", 0, winreg.REG_SZ, machine_id)
    except Exception:
        pass


def get_machine_id():
    """안정적인 머신 ID 생성/조회 (업데이트/재빌드 시에도 동일 PC면 유지)."""
    cached = _load_persisted_machine_id()
    if cached:
        return cached if cached.startswith(MACHINE_ID_PREFIX) else f"{MACHINE_ID_PREFIX}{cached}"

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

    # 5) 결정적 fallback (랜덤 금지: 재빌드/재실행에도 동일)
    if not parts:
        fallback = "|".join([
            os.getenv("COMPUTERNAME", ""),
            os.getenv("USERDOMAIN", ""),
            os.getenv("PROCESSOR_IDENTIFIER", ""),
            os.getenv("SystemDrive", ""),
            f"{uuid.getnode():012x}",
        ])
        parts.append(f"FALLBACK:{fallback}")

    raw = "|".join(parts)
    stable_id = "MID-" + hashlib.sha256(raw.encode("utf-8")).hexdigest()[:32].upper()
    _save_persisted_machine_id(stable_id)
    return f"{MACHINE_ID_PREFIX}{stable_id}"


def check_license_from_sheet(machine_id):
    """Text cleaned due to encoding issue."""
    sheet_url = "https://docs.google.com/spreadsheets/d/10-AseeTNvE97wo29HT2ajui918bg5ICj5L9UOYV0NBo/export?format=csv&gid=0"
    try:
        safe_print(f"라이선스 확인 중... ID: {machine_id}")
        response = requests.get(sheet_url, timeout=5)
        if response.status_code == 200:
            # comment cleaned (encoding issue)
            df = pd.read_csv(io.StringIO(response.text))
            
            # comment cleaned (encoding issue)
            if len(df.columns) >= 4:
                # comment cleaned (encoding issue)
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
        
        # comment cleaned (encoding issue)
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
        
        # comment cleaned (encoding issue)
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
        
        # comment cleaned (encoding issue)
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
        
        # comment cleaned (encoding issue)
        note_layout = QHBoxLayout()
        bulb_icon = QLabel("💡")
        bulb_icon.setStyleSheet("font-size: 16px; background-color: transparent;")
        note_text = QLabel("참고: PC를 변경하면 머신 ID가 바뀔 수 있습니다.")
        note_text.setStyleSheet("font-size: 13px; color: #888888; background-color: transparent;")
        note_layout.addWidget(bulb_icon)
        note_layout.addWidget(note_text)
        note_layout.addStretch()
        layout.addLayout(note_layout)
        
        # comment cleaned (encoding issue)
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
        if clipboard is None:
            return
        clipboard.setText(self.id_input.text())
        sender = self.sender()
        if isinstance(sender, QPushButton):
            sender.setText("완료")
            sender.setEnabled(False)
            QTimer.singleShot(2000, lambda: self._reset_btn(sender))

    def _reset_btn(self, btn: QPushButton):
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
        
        # comment cleaned (encoding issue)
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
        
        # comment cleaned (encoding issue)
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
        kakao_btn.setMinimumHeight(60)  # 명시이 정로 기 보
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
        
        # comment cleaned (encoding issue)
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
        self.resize_step = 30  # 크롤당 기 량
        
    def wheelEvent(self, event):
        # comment cleaned (encoding issue)
        if event.modifiers() & Qt.KeyboardModifier.ControlModifier:
            # comment cleaned (encoding issue)
            delta = event.angleDelta().y()
            current_height = self.height()
            
            if delta > 0:  # 로 크(기 증)
                new_height = min(current_height + self.resize_step, self.max_height)
            else:  # 래크(기 감소)
                new_height = max(current_height - self.resize_step, self.min_height)
            
            # comment cleaned (encoding issue)
            self.setMaximumHeight(new_height)
            self.setMinimumHeight(new_height)
            
            # comment cleaned (encoding issue)
            event.accept()
        else:
            # comment cleaned (encoding issue)
            super().wheelEvent(event)


class SmartProgressTextEdit(ResizableTextEdit):
    """Text cleaned due to encoding issue."""
    
    def __init__(self, parent=None, min_height=200, max_height=800):
        super().__init__(parent, min_height, max_height)
        self.user_is_scrolling = False
        self.last_scroll_time = 0
        self.auto_scroll_enabled = True
        self.search_widget = None
        self.last_search_text = ""
        
        # comment cleaned (encoding issue)
        scrollbar = self.verticalScrollBar()
        if scrollbar:
            scrollbar.valueChanged.connect(self._on_scroll_changed)
            scrollbar.sliderPressed.connect(self._on_user_scroll_start)
            scrollbar.sliderReleased.connect(self._on_user_scroll_end)
        
        # comment cleaned (encoding issue)
        self.search_shortcut = QShortcut(QKeySequence("Ctrl+F"), self)
        self.search_shortcut.activated.connect(self.show_search_dialog)
        
    def _on_scroll_changed(self, value):
        """Text cleaned due to encoding issue."""
        import time
        current_time = time.time()
        
        # comment cleaned (encoding issue)
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
        
        # comment cleaned (encoding issue)
        QTimer.singleShot(3000, self._enable_auto_scroll)
        
    def _enable_auto_scroll(self):
        """Text cleaned due to encoding issue."""
        if not self.user_is_scrolling:
            self.auto_scroll_enabled = True
            
            # comment cleaned (encoding issue)
            QTimer.singleShot(3000, self._enable_auto_scroll)
            
    def wheelEvent(self, event):
        """Handle wheel event."""
        # comment cleaned (encoding issue)
        if not (event.modifiers() & Qt.KeyboardModifier.ControlModifier):
            import time
            self.auto_scroll_enabled = False
            self.last_scroll_time = time.time()
            # comment cleaned (encoding issue)
            QTimer.singleShot(3000, self._enable_auto_scroll)
            
        super().wheelEvent(event)
        
    def append_with_smart_scroll(self, text):
        """Text cleaned due to encoding issue."""
        # comment cleaned (encoding issue)
        scrollbar = self.verticalScrollBar()
        was_at_bottom = False
        if scrollbar:
            was_at_bottom = scrollbar.value() >= scrollbar.maximum() - 10
        
        # comment cleaned (encoding issue)
        self.append(text)
        
        # comment cleaned (encoding issue)
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
        
        # comment cleaned (encoding issue)
        text_content = self.toPlainText()
        
        # comment cleaned (encoding issue)
        cursor = self.textCursor()
        current_position = cursor.position()
        
        # comment cleaned (encoding issue)
        found_index = text_content.find(search_text, current_position)
        
        if found_index == -1:
            # comment cleaned (encoding issue)
            found_index = text_content.find(search_text)
            
        if found_index != -1:
            # comment cleaned (encoding issue)
            cursor.setPosition(found_index)
            cursor.setPosition(found_index + len(search_text), cursor.MoveMode.KeepAnchor)
            self.setTextCursor(cursor)
            self.ensureCursorVisible()
            QMessageBox.information(self, "검색 완료", f"'{search_text}' 텍스트를 찾았습니다.")
        else:
            QMessageBox.information(self, "검색 결과 없음", f"'{search_text}' 텍스트를 찾을 수 없습니다.")


def emergency_save_data():
    """Emergency backup for crash situations."""
    global _current_window, _crash_save_enabled
    
    if not _crash_save_enabled or not _current_window:
        return
    
    try:
        safe_print(" ...")
        
        saved_count = 0
        
        # comment cleaned (encoding issue)
        if hasattr(_current_window, 'active_threads') and _current_window.active_threads:
            save_dir = _current_window.save_path_input.text().strip()
            if not save_dir:
                # comment cleaned (encoding issue)
                desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
                save_dir = os.path.join(desktop_path, "keyword_results")
                try:
                    os.makedirs(save_dir, exist_ok=True)
                except Exception:
                    # comment cleaned (encoding issue)
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
                        
                        # comment cleaned (encoding issue)
                        safe_keyword = re.sub(r'[^\w가-힣\s]', '', base_keyword).strip()[:20]
                        if not safe_keyword:
                            safe_keyword = "응급저장"
                        
                        emergency_file = os.path.join(save_dir, f"{safe_keyword}_응급저장_{current_time}.xlsx")
                        
                        # comment cleaned (encoding issue)
                        if searcher.save_recursive_results_to_excel(emergency_file):
                            safe_print(f" ...")
                            saved_count += 1
                except Exception as inner_e:
                    safe_print(f" ...")
                    continue
            
            if saved_count > 0:
                safe_print(f" ...")
            else:
                safe_print(" ...")
            
    except Exception as e:
        safe_print(f": {str(e)}")
        # comment cleaned (encoding issue)
        try:
            search_thread = getattr(_current_window, "search_thread", None)
            searcher = getattr(search_thread, "searcher", None)
            all_related_keywords = getattr(searcher, "all_related_keywords", None)
            if all_related_keywords:
                
                backup_dir = os.path.join(os.getcwd(), "emergency_backup")
                os.makedirs(backup_dir, exist_ok=True)
                
                backup_file = os.path.join(backup_dir, f"emergency_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json")
                
                with open(backup_file, 'w', encoding='utf-8') as f:
                    json.dump(all_related_keywords, f, 
                             ensure_ascii=False, indent=2)
                
                safe_print(f" ...")
        except:
            safe_print(" ...")


def handle_exception(exc_type, exc_value, exc_traceback):
    """Handle uncaught exceptions and trigger emergency backup."""
    global _crash_save_enabled
    
    if _crash_save_enabled:
        safe_print(" ...")
        safe_print(f" ...")
        safe_print(f" ...")
        
        # comment cleaned (encoding issue)
        emergency_save_data()
        
        # comment cleaned (encoding issue)
        try:
            crash_dir = os.path.join(os.getcwd(), "crash_logs")
            os.makedirs(crash_dir, exist_ok=True)
            
            current_time = datetime.now().strftime('%Y%m%d_%H%M%S')
            crash_file = os.path.join(
    crash_dir, f"crash_log_{current_time}.txt")
            
            with open(crash_file, 'w', encoding='utf-8') as f:
                f.write(f"래발생 간: {datetime.now()}\n")
                f.write(f"외  {exc_type.__name__}\n")
                f.write(f"외 용: {str(exc_value)}\n\n")
                f.write("택 레스:\n")
                traceback.print_exception(
    exc_type, exc_value, exc_traceback, file=f)
            
            safe_print(f" ...")
        except:
            pass
    
    # comment cleaned (encoding issue)
    sys.__excepthook__(exc_type, exc_value, exc_traceback)


def handle_signal(signum, frame):
    """Text cleaned due to encoding issue."""
    signal_names = {
        signal.SIGINT: "SIGINT (Ctrl+C)",
        signal.SIGTERM: "SIGTERM (종료 청)"
    }
    
    signal_name = signal_names.get(signum, f"Signal {signum}")
    safe_print(f" ...")
    
    emergency_save_data()
    
    # comment cleaned (encoding issue)
    if _current_window:
        _current_window.close()
    
    sys.exit(0)


class MultiKeywordTextEdit(QTextEdit):
    """Text cleaned due to encoding issue."""
    search_requested = pyqtSignal()
    _link_rect: Optional[QRect]
    resize_step: int
    min_height: int
    max_height: int
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self._placeholder_text = ""
        
        # comment cleaned (encoding issue)
        self._cta_text = "키워드 공부하러 가기"
        self._cta_url = "https://cafe.naver.com/f-e/cafes/31118881/articles/2036?menuid=12&referrerAllArticles=false"
        self._link_rect = None
        
        # comment cleaned (encoding issue)
        self.setMouseTracking(True)
        
        # comment cleaned (encoding issue)
        self.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        
        # comment cleaned (encoding issue)
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
        """Text cleaned due to encoding issue."""
        self._placeholder_text = text
        self.update()
        
    def paintEvent(self, event):
        """Custom paint event for placeholder and CTA."""
        super().paintEvent(event)
        
        # comment cleaned (encoding issue)
        if not self.toPlainText().strip() and not self.hasFocus():
            viewport = self.viewport()
            if viewport is None:
                return
            painter = QPainter(viewport)
            painter.setRenderHint(QPainter.RenderHint.Antialiasing)
            
            # comment cleaned (encoding issue)
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
            
            # comment cleaned (encoding issue)
            link_h = 40
            spacing_between = 20
            
            # comment cleaned (encoding issue)
            total_content_height = text_block_height + spacing_between + link_h
            start_y = (viewport_rect.height() - total_content_height) / 2 + metrics.ascent()
            
            current_y = start_y
            
            # comment cleaned (encoding issue)
            for i, line in enumerate(lines):
                painter.drawText(int(viewport_rect.left() + padding_left), int(current_y), line)
                current_y += (line_height * line_spacing)
            
            current_y += spacing_between
            
            # comment cleaned (encoding issue)
            
            # comment cleaned (encoding issue)
            link_font = self.font()
            link_font.setPointSize(11)
            link_font.setUnderline(True)
            link_font.setBold(True)
            painter.setFont(link_font)
            painter.setPen(QColor("#0066CC")) # 링크
            
            link_metrics = painter.fontMetrics()
            link_width = link_metrics.horizontalAdvance(self._cta_text)
            link_x = (viewport_rect.width() - link_width) / 2
            
            painter.drawText(int(link_x), int(current_y + link_metrics.ascent()), self._cta_text)
            
            # comment cleaned (encoding issue)
            self._link_rect = QRect(int(link_x), int(current_y), int(link_width), int(link_metrics.height() + 10))

    def mouseMoveEvent(self, event):
        """Handle mouse move for link hover."""
        viewport = self.viewport()
        if viewport is None:
            super().mouseMoveEvent(event)
            return
        if self._link_rect is not None and self._link_rect.contains(event.pos()) and not self.toPlainText().strip():
            viewport.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        else:
            viewport.setCursor(QCursor(Qt.CursorShape.IBeamCursor))
        super().mouseMoveEvent(event)

    def mousePressEvent(self, event):
        """Text cleaned due to encoding issue."""
        if self._link_rect is not None and self._link_rect.contains(event.pos()) and not self.toPlainText().strip():
            QDesktopServices.openUrl(QUrl(self._cta_url))
            return # 링크 릭 커 음

        super().mousePressEvent(event)
    
    def keyPressEvent(self, event):
        """Text cleaned due to encoding issue."""
        if event.key() == Qt.Key.Key_Return or event.key() == Qt.Key.Key_Enter:
            if event.modifiers() == Qt.KeyboardModifier.ShiftModifier:
                # comment cleaned (encoding issue)
                super().keyPressEvent(event)
            else:
                # comment cleaned (encoding issue)
                self.search_requested.emit()
                event.accept()
        else:
            super().keyPressEvent(event)

    def wheelEvent(self, event):
        """Text cleaned due to encoding issue."""
        # comment cleaned (encoding issue)
        if event.modifiers() & Qt.KeyboardModifier.ControlModifier:
            # comment cleaned (encoding issue)
            delta = event.angleDelta().y()
            current_height = self.height()
            
            if delta > 0:  # 로 크(기 증)
                new_height = min(current_height + self.resize_step, self.max_height)
            else:  # 래크(기 감소)
                new_height = max(current_height - self.resize_step, self.min_height)
            
            # comment cleaned (encoding issue)
            self.setMaximumHeight(new_height)
            self.setMinimumHeight(new_height)
            
            # comment cleaned (encoding issue)
            event.accept()
        else:
            # comment cleaned (encoding issue)
            super().wheelEvent(event)


class NaverMobileSearchScraper:
    """Text cleaned due to encoding issue."""
    
    def __init__(self, driver: Optional[WebDriver] = None):
        self.session = requests.Session()
        self.driver: Optional[WebDriver] = driver
        self.results = []
        self.searched_keywords = set()
        self.save_dir = ""
        self.extracted_keywords = set()
        self.is_running = True
        self.all_related_keywords = []
        self.base_keyword = ""
        self.processed_keywords = set()
        self.search_thread: Optional[Any] = None
        
        # comment cleaned (encoding issue)
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
        """네이버에서 키워드를 검색하고 HTML을 반환합니다."""
        try:
            if progress_callback:
                progress_callback(f"'{keyword}' 검색 시작...")
            
            encoded_keyword = urllib.parse.quote(keyword)
            search_url = f"https://search.naver.com/search.naver?where=nexearch&sm=top_hty&fbm=0&ie=utf8&query={encoded_keyword}"
            
            response = self.session.get(search_url, timeout=10)
            response.raise_for_status()
            
            if progress_callback:
                progress_callback("검색 완료")
            
            return response.text
            
        except Exception as e:
            if progress_callback:
                progress_callback(f"검색 오류: {str(e)}")
            return None

    def extract_related_keywords(self, keyword, progress_callback=None):
        """연관 키워드를 HTML 파싱으로 추출합니다."""
        keywords = []
        
        try:
            if not BEAUTIFULSOUP_AVAILABLE:
                if progress_callback:
                    progress_callback("BeautifulSoup가 설치되지 않았습니다.")
                return keywords
            
            html_content = self.search_keyword(keyword, progress_callback)
            if not html_content:
                return keywords
            
            soup = BeautifulSoup(html_content, 'html.parser')
            
            if progress_callback:
                progress_callback(f"'{keyword}' 페이지에서 연관어를 추출하는 중...")
            
            # comment cleaned (encoding issue)
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
                    progress_callback(f"선택자 '{selector}'에서 {len(elements)}개 요소를 발견했습니다.")
                
                for element in elements:
                    try:
                        keyword_text = element.get_text(strip=True)
                        if keyword_text and keyword_text not in keywords:
                            keywords.append(keyword_text)
                            found_count += 1
                            if progress_callback:
                                progress_callback(f"연관 키워드 발견 ({found_count}): {keyword_text}")
                    except:
                        continue
            
            if progress_callback:
                progress_callback(f"연관 키워드 {len(keywords)}개 추출 완료")
            
            return keywords
            
        except Exception as e:
            if progress_callback:
                progress_callback(f"연관 키워드 추출 오류: {str(e)}")
            return []

    def check_internet_connection(self):
        """인터넷 연결 상태를 확인합니다."""
        try:
            # comment cleaned (encoding issue)
            response = requests.get("https://www.naver.com", timeout=5)
            return response.status_code == 200
        except:
            try:
                # comment cleaned (encoding issue)
                import socket
                socket.create_connection(("8.8.8.8", 53), timeout=3)
                return True
            except:
                return False

    def check_pause_status(self, progress_callback=None):
        """일시정지 상태와 인터넷 연결 상태를 확인합니다."""
        # comment cleaned (encoding issue)
        if not self.is_running:
            return False
        
        # comment cleaned (encoding issue)
        search_thread = self.search_thread
        if search_thread and hasattr(search_thread, 'is_paused'):
            if search_thread.is_paused:
                if progress_callback:
                    progress_callback("작업이 일시정지되었습니다. '재개' 버튼을 눌러주세요.")
                
                # comment cleaned (encoding issue)
                while search_thread.is_paused and self.is_running:
                    time.sleep(0.5)
                
                if not self.is_running:
                    return False
                
                if progress_callback:
                    progress_callback("작업을 재개합니다.")
        
        # comment cleaned (encoding issue)
        if not self.check_internet_connection():
            if progress_callback:
                progress_callback("인터넷 연결이 끊어졌습니다. 연결 복구를 기다립니다...")
            
            # comment cleaned (encoding issue)
            connection_wait_count = 0
            while not self.check_internet_connection() and self.is_running:
                time.sleep(2)
                connection_wait_count += 1
                
                # comment cleaned (encoding issue)
                if search_thread and hasattr(search_thread, 'is_paused') and search_thread.is_paused:
                    if progress_callback:
                        progress_callback("인터넷 연결 대기 중 일시정지됨")
                    break
                
                # comment cleaned (encoding issue)
                if connection_wait_count % 5 == 0 and progress_callback:
                    progress_callback(f"인터넷 재연결 시도 중... ({connection_wait_count * 2}초 경과)")
            
            if not self.is_running:
                return False
            
            # comment cleaned (encoding issue)
            if self.check_internet_connection():
                if progress_callback:
                    progress_callback("인터넷 연결이 복구되었습니다. 작업을 계속합니다.")
        
        # comment cleaned (encoding issue)
        return True

    def initialize_browser(self):
        """Text cleaned due to encoding issue."""
        try:
            safe_print("브라우저 초기화를 시작합니다...")
            driver = create_chrome_driver(log_callback=safe_print)
            if not driver:
                safe_print("Chrome 드라이버 생성 실패")
                return False
            self.driver = driver
            try:
                self.driver.get("https://m.naver.com")
            except TimeoutException:
                try:
                    self.driver.execute_script("window.stop();")
                except Exception:
                    pass
            return True
             
            driver_path = None
            
            # comment cleaned (encoding issue)
            import shutil
            
            # comment cleaned (encoding issue)
            base_paths = []
            if getattr(sys, 'frozen', False):
                base_paths.append(os.path.dirname(sys.executable))
                meipass_dir = getattr(sys, "_MEIPASS", None)
                if meipass_dir:
                    base_paths.append(meipass_dir)
            base_paths.append(os.getcwd())
            
            for base_path in base_paths:
                local_driver = os.path.join(base_path, "chromedriver.exe")
                if os.path.exists(local_driver):
                    safe_print(f" ...")
                    driver_path = local_driver
                    break
            
            # comment cleaned (encoding issue)
            if not driver_path:
                try:
                    from webdriver_manager.chrome import ChromeDriverManager
                    safe_print(" ...")
                    # comment cleaned (encoding issue)
                    driver_path = ChromeDriverManager().install()
                    safe_print(f" ...")
                except Exception as e:
                    safe_print(f": {str(e)}")
            
            # comment cleaned (encoding issue)
            if not driver_path and shutil.which("chromedriver"):
                driver_path = "chromedriver"
                safe_print(" ...")
            
            if not driver_path:
                raise Exception(
                    "ChromeDriver를 찾을 수 없습니다.\n"
                    "프로그램 폴더에 'chromedriver.exe'를 넣어주거나\n"
                    "인터넷 연결 상태를 확인해주세요."
                )

            # comment cleaned (encoding issue)
            try:
                service = Service(driver_path)
            except:
                if driver_path == "chromedriver":
                    service = Service()
                else:
                    service = Service(executable_path=driver_path)

            # comment cleaned (encoding issue)
            if os.name == 'nt':
                try:
                    startupopt = subprocess.STARTUPINFO()
                    startupopt.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                    service.creation_flags = subprocess.CREATE_NO_WINDOW
                except:
                    pass

            options = webdriver.ChromeOptions()
            
            # comment cleaned (encoding issue)
            options.add_argument("--headless")  # 드리스 모드 성
            safe_print(" ...")
            
            options.add_argument("--window-size=1920,1080")  #  FHD 상로 정
            options.add_argument("--start-maximized")  # 브라 최(드리스서효)
            
            # comment cleaned (encoding issue)
            options.add_argument(
                "--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")

            # comment cleaned (encoding issue)
            options.add_argument("--disable-web-security")  # CORS 제
            options.add_argument("--allow-running-insecure-content")  # 합 콘텐용
            options.add_argument("--force-device-scale-factor=1")  # 터 고정
            options.add_argument("--disable-features=VizDisplayCompositor")  # 렌더링 최적화
            
            # comment cleaned (encoding issue)
            options.add_argument("--no-sandbox")
            options.add_argument("--disable-dev-shm-usage") 
            options.add_argument("--disable-gpu")
            options.add_argument("--disable-extensions")
            options.add_argument("--disable-plugins")
            options.add_argument("--disable-images")
            
            # comment cleaned (encoding issue)
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
            
            # comment cleaned (encoding issue)
            options.add_argument("--lang=en-US")  # 어정
            options.add_argument("--disable-logging")  # 모든 로그 비활화
            options.add_argument("--disable-gpu-sandbox")
            options.add_argument("--log-level=3")  # 각류
            options.add_argument("--silent")  # 조용모드
            # comment cleaned (encoding issue)
            options.add_argument("--disable-features=TranslateUI,VizDisplayCompositor")
            options.add_argument("--disable-ipc-flooding-protection")  # IPC 러보호 비활화
            
            # comment cleaned (encoding issue)
            options.add_experimental_option('useAutomationExtension', False)
            options.add_experimental_option("excludeSwitches", ["enable-logging", "enable-automation"])
            options.add_experimental_option("detach", True)
            options.add_argument("--disable-blink-features=AutomationControlled")
            
            # comment cleaned (encoding issue)
            options.add_argument("--network-timeout=15")  # 3015
            options.add_argument("--page-load-strategy=eager")
            options.add_argument("--timeout=15000")  # 15아
            options.add_argument("--dns-prefetch-disable")  # DNS 리치 비활화
            
            # comment cleaned (encoding issue)
            self.driver = webdriver.Chrome(service=service, options=options)
            safe_print(" ...")
            
            # comment cleaned (encoding issue)
            self.driver.set_page_load_timeout(15)  # 이 로딩 아15초로 축
            self.driver.implicitly_wait(3)  # 시3초로 축
            
            safe_print(" ...")
            
            # comment cleaned (encoding issue)
            max_retries = 3
            for attempt in range(max_retries):
                try:
                    # comment cleaned (encoding issue)
                    self.driver.get("https://m.naver.com")
                    
                    # comment cleaned (encoding issue)
                    self.driver.execute_script("""
                        // 브라? ?기 ?오?
                        var windowWidth = window.innerWidth;
                        var windowHeight = window.innerHeight;
                        
                        // viewport 메? ?그?브라? ?기?맞게 ?정
                        var existingMeta = document.querySelector('meta[name="viewport"]');
                        if (existingMeta) {
                            existingMeta.remove();
                        }
                        var meta = document.createElement('meta');
                        meta.name = 'viewport';
                        meta.content = 'width=' + windowWidth + ', initial-scale=1.0, maximum-scale=3.0, user-scalable=yes';
                        document.getElementsByTagName('head')[0].appendChild(meta);
                        
                        // ?이 ?체?브라? ?기?맞게 조정
                        document.documentElement.style.width = '100%';
                        document.documentElement.style.height = '100%';
                        document.body.style.minWidth = windowWidth + 'px';
                        document.body.style.width = '100%';
                        document.body.style.height = '100vh';
                        document.body.style.transform = 'none';
                        document.body.style.transformOrigin = 'top left';
                        document.body.style.margin = '0';
                        document.body.style.padding = '0';
                        document.body.style.fontSize = Math.max(14, windowWidth / 100) + 'px';  // 기른 트 조정
                        
                        // 메인 컨테?너?을 브라? ?체 ?기?장
                        var containers = document.querySelectorAll('.container, .wrap, .content_area, #wrap, .nx_wrap');
                        containers.forEach(function(container) {
                            container.style.maxWidth = '100%';
                            container.style.width = '100%';
                            container.style.minWidth = windowWidth + 'px';
                        });
                        
                        // ?역?브라? 창에 맞게 ?
                        var searchArea = document.querySelector('.TF7QLJYoGthrUnoIpxEj, .api_subject_bx, .search_result');
                        if (searchArea) {
                            searchArea.style.minHeight = (windowHeight - 200) + 'px';
                            searchArea.style.width = '100%';
                            searchArea.style.overflow = 'visible';
                            searchArea.style.maxWidth = 'none';
                        }
                        
                        console.log('이 브라 기(' + windowWidth + 'x' + windowHeight + ')맞게 조정었니');
                    """)
                    
                    # comment cleaned (encoding issue)
                    time.sleep(2)
                    
                    WebDriverWait(self.driver, 10).until(
                        EC.presence_of_element_located((By.TAG_NAME, "body"))
                    )
                    safe_print(" ...")
                    time.sleep(2)
                    return True
                except Exception as e:
                    if attempt < max_retries - 1:
                        safe_print(f" ...")
                        time.sleep(3)
                    else:
                        safe_print(f": {str(e)}")
                        return False
            
            return True

        except Exception as e:
            error_msg = f"브라 초기류:\n{str(e)}\n\nChrome 브라 치어 는 인주요."
            safe_print(f" ...")
            
            # comment cleaned (encoding issue)
            try:
                global _current_window
                if _current_window:
                    # comment cleaned (encoding issue)
                    # comment cleaned (encoding issue)
                    # comment cleaned (encoding issue)
                    # comment cleaned (encoding issue)
                    from PyQt6.QtCore import QTimer
                    QTimer.singleShot(0, lambda: QMessageBox.critical(
                        _current_window, "브라 류", error_msg))
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
                        progress_callback(f"{keyword}  ...")
                    else:
                        progress_callback(f"{keyword}  ...")
                
                encoded_keyword = urllib.parse.quote(keyword)
                search_url = f"https://m.search.naver.com/search.naver?where=m&sm=mtp_hty.top&query={encoded_keyword}"
                
                if self.driver:
                    self.driver.get(search_url)
                
                if progress_callback:
                    progress_callback(" ...")
                
                # comment cleaned (encoding issue)
                time.sleep(random.uniform(1.5, 2.5))  #  방 + 당최적
                
                try:
                    if self.driver:
                        WebDriverWait(self.driver, 5).until(  # 85초로 축
                            EC.presence_of_element_located((By.TAG_NAME, "body"))
                        )
                except TimeoutException:
                    if attempt < max_retries - 1:
                        if progress_callback:
                            progress_callback(f" ...")
                        continue
                    else:
                        if progress_callback:
                            progress_callback(" ...")
                        return False
                
                time.sleep(1)
                
                if progress_callback:
                    progress_callback(" ...")
                
                return True
                
            except Exception as e:
                error_msg = str(e)
                
                # comment cleaned (encoding issue)
                if "invalid session id" in error_msg.lower() or "no such session" in error_msg.lower() or "no such window" in error_msg.lower():
                    if progress_callback:
                        progress_callback(f" ...")
                    
                    # comment cleaned (encoding issue)
                    if self.initialize_browser():
                        if progress_callback:
                            progress_callback(f" ...")
                        continue
                    else:
                        if progress_callback:
                            progress_callback(f" ...")
                        return False
                
                # comment cleaned (encoding issue)
                elif "timeout" in error_msg.lower() and "renderer" in error_msg.lower():
                    if progress_callback:
                        progress_callback(f" ...")
                    
                    if attempt < max_retries - 1:
                        if progress_callback:
                            progress_callback(f" ...")
                        
                        # comment cleaned (encoding issue)
                        if self.initialize_browser():
                            if progress_callback:
                                progress_callback(f" ...")
                            continue
                        else:
                            if progress_callback:
                                progress_callback(f" ...")
                    else:
                        if progress_callback:
                            progress_callback(f" ...")
                        return False
                
                # comment cleaned (encoding issue)
                elif "timeout" in error_msg.lower():
                    if progress_callback:
                        progress_callback(f" ...")
                    
                    if attempt < max_retries - 1:
                        if progress_callback:
                            progress_callback(f" ...")
                        time.sleep(5)  # 아의 경우 조금 
                        continue
                    else:
                        if progress_callback:
                            progress_callback(f" ...")
                        return False
                
                # comment cleaned (encoding issue)
                if attempt < max_retries - 1:
                    if progress_callback:
                        progress_callback(f": {str(e)}")
                    time.sleep(3)
                    continue
                else:
                    if progress_callback:
                        progress_callback(f": {str(e)}")
                    return False
        
        return False

    def extract_autocomplete_keywords(self, keyword, progress_callback=None):
        """Extract autocomplete keywords."""
        keywords = []
        max_retries = 2
        
        for attempt in range(max_retries):
            try:
                if not self.driver:
                    # comment cleaned (encoding issue)
                    if not self.initialize_browser():
                        return keywords
                driver = self.driver
                if driver is None:
                    return keywords
                
                if progress_callback:
                    if attempt > 0:
                        progress_callback(f"{keyword}  ...")
                    else:
                        progress_callback(f"{keyword}  ...")
                
                # comment cleaned (encoding issue)
                try:
                    driver.set_page_load_timeout(15)  # 15한
                    driver.get("https://m.naver.com")
                except TimeoutException:
                    if progress_callback:
                        progress_callback(" ...")
                    try:
                        driver.execute_script("window.stop();")
                    except:
                        pass
                except Exception as e:
                    # comment cleaned (encoding issue)
                    raise e
                
                time.sleep(2)
            
                # comment cleaned (encoding issue)
                try:
                    from selenium.webdriver.support.ui import WebDriverWait
                    from selenium.webdriver.support import expected_conditions as EC
                    
                    # comment cleaned (encoding issue)
                    try:
                        wait = WebDriverWait(driver, 10)
                        wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
                        if progress_callback:
                            progress_callback(" ...")
                    except TimeoutException:
                        if progress_callback:
                            progress_callback(" ...")
                            
                except Exception as e:
                    if progress_callback:
                        progress_callback(f": {str(e)}")
            
                # comment cleaned (encoding issue)
                search_input = None
                search_selectors = [
                    '#nx_query',
                    'input.search_input',
                    'input[name="query"]',
                    'input[type="search"]'
                ]
                
                for selector in search_selectors:
                    try:
                        search_input = driver.find_element(By.CSS_SELECTOR, selector)
                        if search_input and search_input.is_enabled():
                            if progress_callback:
                                progress_callback(f" ...")
                            break
                    except:
                        continue
                
                if not search_input:
                    if progress_callback:
                        progress_callback(" ...")
                    return keywords
            
                # comment cleaned (encoding issue)
                try:
                    # comment cleaned (encoding issue)
                    driver.execute_script("""
                        var input = arguments[0];
                        var keyword = arguments[1];
                        input.focus();
                        input.value = keyword;
                        input.dispatchEvent(new Event('input', { bubbles: true }));
                        input.dispatchEvent(new Event('keyup', { bubbles: true }));
                    """, search_input, keyword)
                    
                    # comment cleaned (encoding issue)
                    time.sleep(2)
                    
                    if progress_callback:
                        progress_callback(f"{keyword}  ...")
                        
                except Exception as input_error:
                    if progress_callback:
                        progress_callback(f" ...")
                        return keywords
            
                # comment cleaned (encoding issue)
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
                        elements = driver.find_elements(By.CSS_SELECTOR, selector)
                        
                        for element in elements:
                            try:
                                if not element.is_displayed():
                                    continue
                                    
                                # comment cleaned (encoding issue)
                                keyword_text = element.get_attribute("textContent") or element.text
                                
                                if keyword_text:
                                    keyword_text = keyword_text.strip()
                                    
                                    # comment cleaned (encoding issue)
                                    if '<' in keyword_text:
                                        import re
                                        keyword_text = re.sub(r'<[^>]+>', '', keyword_text)
                                        keyword_text = keyword_text.strip()
                                    
                                    # comment cleaned (encoding issue)
                                        keyword_text = self.clean_duplicate_text(keyword_text)
                                        
                                    # comment cleaned (encoding issue)
                                    if (keyword.lower() in keyword_text.lower() and 
                                        keyword_text not in keywords and
                                        len(keyword_text) <= 50 and
                                        len(keyword_text) > 1):
                                            keywords.append(keyword_text)
                                            found_count += 1
                                            if progress_callback:
                                                progress_callback(f" ...")
                            except Exception:
                                continue
                    except Exception:
                        continue
                
                # comment cleaned (encoding issue)
                keywords = list(set(keywords))
                keywords.sort()
                
                if progress_callback:
                    progress_callback(f"  : {len(keywords)}")
                
                return keywords
            
            except Exception as e:
                error_msg = str(e)
                if "no such window" in error_msg.lower() or "invalid session id" in error_msg.lower():
                    if progress_callback:
                        progress_callback(" ...")
                    
                    if attempt < max_retries - 1:
                        if self.initialize_browser():
                            if progress_callback:
                                progress_callback(" ...")
                            continue
                
                if progress_callback:
                    progress_callback(f": {str(e)}")
                
                if attempt == max_retries - 1:
                    return []
        
        return keywords

    def extract_related_keywords_new(self, current_keyword, progress_callback=None):
        """Extract related keywords from current page."""
        keywords = []
        
        try:
            if not self.driver:
                return keywords
            
            # comment cleaned (encoding issue)
            related_selectors = [
                '#_related_keywords .keyword a',  # 메인 택
                '.related_srch .lst a',
                '.related_keyword a',
                '.lst_related a',
                '.keyword_area a',
                '.related_search a',
                '.keyword a'
            ]
            
            if progress_callback:
                progress_callback(f" ...")
            
            found_count = 0
            
            for selector in related_selectors:
                try:
                    elements = self.driver.find_elements(By.CSS_SELECTOR, selector)
                    if len(elements) > 0 and progress_callback:
                        progress_callback(f"  : {len(keywords)}")
                    
                    for element in elements:
                        try:
                            # comment cleaned (encoding issue)
                            keyword_text = ""
                            
                            # comment cleaned (encoding issue)
                            try:
                                keyword_text = element.text
                                if keyword_text:
                                    keyword_text = keyword_text.strip()
                            except:
                                keyword_text = ""

                            # comment cleaned (encoding issue)
                            if not keyword_text:
                                try:
                                    keyword_text = element.get_attribute("textContent")
                                    if keyword_text:
                                        keyword_text = keyword_text.strip()
                                except:
                                    keyword_text = ""

                            # comment cleaned (encoding issue)
                            if not keyword_text:
                                try:
                                    keyword_text = element.get_attribute("innerText")
                                    if keyword_text:
                                        keyword_text = keyword_text.strip()
                                except:
                                    keyword_text = ""

                            # comment cleaned (encoding issue)
                            if not keyword_text:
                                try:
                                    keyword_text = self.driver.execute_script("""
                                        var element = arguments[0];
                                        if (!element) return '';
                                        
                                            // 링크 ?소?직접?인 ?스?만 추출
                                            var textContent = element.textContent || element.innerText || '';
                                            
                                            // ?뒤 공백 ?거 ?속 공백 ?리
                                            return textContent.replace(/\\s+/g, ' ').trim();
                                    """, element)
                                except:
                                    keyword_text = ""
                            
                            # comment cleaned (encoding issue)
                            if keyword_text:
                                import re
                                # comment cleaned (encoding issue)
                                keyword_text = re.sub(r'<[^>]+>', '', keyword_text)
                                # comment cleaned (encoding issue)
                                keyword_text = re.sub(r'\s+', ' ', keyword_text)
                                # comment cleaned (encoding issue)
                                keyword_text = re.sub(r'[\u200b-\u200d\ufeff]', '', keyword_text)  # 제로폭 문자 제거
                                # comment cleaned (encoding issue)
                                keyword_text = re.sub(r'\s+[가-힣]{1}$', '', keyword_text)
                                keyword_text = re.sub(r'\s+[a-zA-Z]{1}$', '', keyword_text)  # 끝에 영문 1자만 남는 경우 제거
                                keyword_text = keyword_text.strip()
                                
                            if keyword_text:
                                # comment cleaned (encoding issue)
                                keyword_text = self.clean_duplicate_text(keyword_text)
                                    
                                # comment cleaned (encoding issue)
                                # comment cleaned (encoding issue)
                                if keyword_text and not re.search(r'[가-힣]{1}$|[a-zA-Z]{1}$', keyword_text):
                                    # comment cleaned (encoding issue)
                                    if (keyword_text not in keywords and
                                        len(keyword_text) <= 50 and
                                        len(keyword_text) > 1):
                                        keywords.append(keyword_text)
                                        found_count += 1
                                        if progress_callback:
                                            progress_callback(f" ...")
                                elif keyword_text and len(keyword_text) > 3:  # 3상면 용
                                    if (keyword_text not in keywords and 
                                        len(keyword_text) <= 50 and 
                                        len(keyword_text) > 1):
                                        keywords.append(keyword_text)
                                        found_count += 1
                                        if progress_callback:
                                            progress_callback(f" ...")
                        except Exception as e:
                            continue
                except Exception as e:
                    continue
            
            # comment cleaned (encoding issue)
            keywords = list(set(keywords))
            keywords.sort()
            
            if progress_callback:
                progress_callback(f"  : {len(keywords)}")
            
            return keywords
            
        except Exception as e:
            if progress_callback:
                progress_callback(f": {str(e)}")
            return []

    def clean_duplicate_text(self, text):
        """Text cleaned due to encoding issue."""
        if not text:
            return text

        text = text.strip()
        # comment cleaned (encoding issue)
        text = re.sub(r'\s+', ' ', text)
        
        # comment cleaned (encoding issue)
        words = text.split()
        unique_words = []
        seen_words = set()
        
        for word in words:
            word_lower = word.lower()
            if word_lower not in seen_words:
                unique_words.append(word)
                seen_words.add(word_lower)
        
        # comment cleaned (encoding issue)
        result = ' '.join(unique_words)
        return result

    def extract_together_keywords(self, current_keyword, progress_callback=None):
        """Text cleaned due to encoding issue."""
        keywords = []
        try:
            if progress_callback:
                progress_callback(f" ...")
            
            # comment cleaned (encoding issue)
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
                progress_callback(f"  : {len(keywords)}")
        
            return list(set(keywords))
        except Exception as e:
            if progress_callback:
                progress_callback(f": {str(e)}")
            return []
            
    def extract_popular_topics(self, current_keyword, progress_callback=None):
        """Text cleaned due to encoding issue."""
        keywords = []
        try:
            if progress_callback:
                progress_callback(f" ...")
            
            # comment cleaned (encoding issue)
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
                progress_callback(f"  : {len(keywords)}")
            
            return list(set(keywords))
        except Exception as e:
            if progress_callback:
                progress_callback(f": {str(e)}")
            return []

    def recursive_keyword_extraction(self, initial_keyword, progress_callback=None, extract_autocomplete=True):
        """Text cleaned due to encoding issue."""
        if not self.driver:
            if progress_callback:
                progress_callback(" ...")
            return False
        
        self.base_keyword = initial_keyword
        self.all_related_keywords = []
        self.processed_autocomplete_keywords = set()  # 처리동성어 추적
        
        if progress_callback:
            progress_callback(f" ...")

        # comment cleaned (encoding issue)
        success = self._extract_all_keyword_types(
            initial_keyword, 
            parent_keyword=initial_keyword, 
            depth=0, 
            progress_callback=progress_callback
        )
        
        if not success:
            return False
            
        # comment cleaned (encoding issue)
        if extract_autocomplete:
            autocomplete_keywords = self.extract_autocomplete_keywords(initial_keyword, progress_callback)
            
            # comment cleaned (encoding issue)
            for keyword in autocomplete_keywords:
                self.all_related_keywords.append({
                    'depth': 0,
                    'parent_keyword': initial_keyword,
                    'current_keyword': initial_keyword,
                    'related_keyword': keyword,
                    'keyword_type': '동성',
                    'source_type': '동성어',
                    'extracted_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                })
                
            # comment cleaned (encoding issue)
            keywords_for_recursion = []
            for keyword in autocomplete_keywords:
                # comment cleaned (encoding issue)
                if keyword.lower().strip() != initial_keyword.lower().strip():
                    keywords_for_recursion.append(keyword)
            
            if progress_callback:
                progress_callback(f"재귀 대상 자동완성 키워드 수: {len(keywords_for_recursion)}개")
                if keywords_for_recursion:
                    progress_callback("자동완성 재귀 추출을 시작합니다.")
            
            # comment cleaned (encoding issue)
            if keywords_for_recursion:
                self._recursive_autocomplete_extraction(
                    keywords_for_recursion, 
                    initial_keyword, 
                    depth=1, 
                    progress_callback=progress_callback
                )
            else:
                if progress_callback:
                    progress_callback(" ...")
            
        if progress_callback:
            progress_callback(f"'{initial_keyword}' 키워드 추출 완료: 총 {len(self.all_related_keywords)}개")

        return True

    def _extract_all_keyword_types(self, current_keyword, parent_keyword, depth, progress_callback=None):
        """Text cleaned due to encoding issue."""
        try:
            if not self.is_running:
                return False
            
            # comment cleaned (encoding issue)
            if not self.check_pause_status(progress_callback):
                return False
                
            # comment cleaned (encoding issue)
            if not self.search_keyword_mobile(current_keyword, progress_callback):
                return False
            
            # comment cleaned (encoding issue)
            if not self.check_pause_status(progress_callback):
                return False
            
            # comment cleaned (encoding issue)
            related_keywords = self.extract_related_keywords_new(current_keyword, progress_callback)
            
            # comment cleaned (encoding issue)
            if not self.check_pause_status(progress_callback):
                return False
                
            together_keywords = self.extract_together_keywords(current_keyword, progress_callback)
            
            # comment cleaned (encoding issue)
            if not self.check_pause_status(progress_callback):
                return False
            
            popular_keywords = self.extract_popular_topics(current_keyword, progress_callback)

            # comment cleaned (encoding issue)
            all_extracted = []
            
            for keyword in related_keywords:
                entry = {
                    'depth': depth,
                    'parent_keyword': parent_keyword,
                    'current_keyword': current_keyword,
                    'related_keyword': keyword,
                    'keyword_type': '어',
                    'source_type': '어',
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
                    'keyword_type': '께많이찾는',
                    'source_type': '께많이찾는',
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
                    'keyword_type': '기주제',
                    'source_type': '기주제',
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
                progress_callback(f": {str(e)}")
            return False

    def _recursive_autocomplete_extraction(self, keywords_to_process, original_keyword, depth, progress_callback=None, max_depth=5):
        """Text cleaned due to encoding issue."""
        
        if depth > max_depth:
            if progress_callback:
                progress_callback(f" ...")
            return
                
        if not self.is_running:
            return
            
        for i, current_keyword in enumerate(keywords_to_process):
            
            if not self.is_running:
                break
                    
            # comment cleaned (encoding issue)
            if current_keyword.lower() in self.processed_autocomplete_keywords:
                if progress_callback:
                    progress_callback(f" ...")
                continue
                    
            # comment cleaned (encoding issue)
            self.processed_autocomplete_keywords.add(current_keyword.lower())
            
            if progress_callback:
                progress_callback(
                    f"[{depth}단계] {i + 1}/{len(keywords_to_process)} 진행 중: '{current_keyword}'"
                )
            
            # comment cleaned (encoding issue)
            self._extract_all_keyword_types(
                current_keyword, 
                parent_keyword=current_keyword, 
                depth=depth, 
                progress_callback=progress_callback
            )
            
            # comment cleaned (encoding issue)
            new_autocomplete_keywords = self.extract_autocomplete_keywords(current_keyword, progress_callback)
            
            # comment cleaned (encoding issue)
            for keyword in new_autocomplete_keywords:
                self.all_related_keywords.append({
                    'depth': depth,
                    'parent_keyword': current_keyword,
                    'current_keyword': current_keyword,
                    'related_keyword': keyword,
                    'keyword_type': '동성',
                    'source_type': '동성어',
                    'extracted_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                })
            
            # comment cleaned (encoding issue)
            if new_autocomplete_keywords:
                # comment cleaned (encoding issue)
                filtered_keywords = []
                for keyword in new_autocomplete_keywords:
                    # comment cleaned (encoding issue)
                    if keyword.lower().strip() == current_keyword.lower().strip():
                        continue
                    
                    # comment cleaned (encoding issue)
                    if keyword.lower() not in self.processed_autocomplete_keywords:
                        # comment cleaned (encoding issue)
                        if self.base_keyword.lower() in keyword.lower() or len(filtered_keywords) < 20:  # 무 많 워방
                            filtered_keywords.append(keyword)
                
                if filtered_keywords:
                    if progress_callback:
                        progress_callback(
                            f"'{current_keyword}'에서 재귀 대상 {len(filtered_keywords)}개 발견"
                        )
                        if len(filtered_keywords) <= 10:  # 10하모두 시
                            progress_callback(f"다음 처리: {', '.join(filtered_keywords)}")
                        else:  # 10초과처음 10개만 시
                            progress_callback(
                                f"다음 처리: {', '.join(filtered_keywords[:10])} ... (총 {len(filtered_keywords)}개)"
                            )
                    
                    # comment cleaned (encoding issue)
                    self._recursive_autocomplete_extraction(
                        filtered_keywords, 
                        original_keyword, 
                        depth + 1, 
                        progress_callback, 
                        max_depth
                    )
                else:
                    if progress_callback:
                        progress_callback(f" ...")
            else:
                if progress_callback:
                    progress_callback(f" ...")
            
        if progress_callback:
            progress_callback(f" ...")

    def save_recursive_results_to_excel(self, save_path=None, progress_callback=None):
        """Save extraction results to file."""
        try:
            if not hasattr(self, 'all_related_keywords') or not self.all_related_keywords:
                if progress_callback:
                    progress_callback(" ...")
                return False
            
            if not save_path:
                if not self.save_dir:
                    self.save_dir = "keyword_results"
                    os.makedirs(self.save_dir, exist_ok=True)
                
                current_time = datetime.now().strftime('%Y%m%d_%H%M%S')
                base_keyword = getattr(self, 'base_keyword', 'keyword_extraction')
                save_path = os.path.join(self.save_dir, f"{base_keyword}_{current_time}.xlsx")
            
            # comment cleaned (encoding issue)
            df = pd.DataFrame({
                '추출된_키워드': [item['related_keyword'] for item in self.all_related_keywords]
            })

            # comment cleaned (encoding issue)
            df = df.drop_duplicates(subset=['추출된_키워드'], keep='first').reset_index(drop=True)

            # comment cleaned (encoding issue)
            try:
                df.to_excel(save_path, index=False, engine='openpyxl')
                
                if os.path.exists(save_path) and os.path.getsize(save_path) > 0:
                    if progress_callback:
                        progress_callback(f" ...")
                        progress_callback(f"저장된 키워드 수: {len(df)}")
                    return True
                else:
                    raise Exception(" 일 성 패")
                
            except Exception as excel_error:
                # comment cleaned (encoding issue)
                csv_path = save_path.rsplit('.', 1)[0] + '.csv'
                df.to_csv(csv_path, index=False, encoding='utf-8-sig')
                    
                if progress_callback:
                    progress_callback(f" ...")
                return True
            
        except Exception as e:
            if progress_callback:
                progress_callback(f": {str(e)}")
            return False

    def close(self):
        """Text cleaned due to encoding issue."""
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


class ParallelKeywordThread(QThread):
    finished = pyqtSignal(str)              # 료 된 일 경로 그
    error = pyqtSignal(str)                 # 러 그
    log = pyqtSignal(str, str)              # 로그 그(워 메시)
    
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
            
            # comment cleaned (encoding issue)
            self.driver = create_chrome_driver(
                log_callback=lambda msg: self.log.emit(self.keyword, str(msg))
            )
            if not self.driver:
                self.error.emit(f"'{self.keyword}' 브라우저 생성 실패")
                return

            # comment cleaned (encoding issue)
            self.searcher = NaverMobileSearchScraper(driver=self.driver)
            self.searcher.save_dir = self.save_dir
            self.searcher.is_running = self.is_running
            self.searcher.search_thread = self
            
            # comment cleaned (encoding issue)
            success = self.searcher.recursive_keyword_extraction(
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
                self.searcher and
                hasattr(self.searcher, "all_related_keywords") and
                self.searcher.all_related_keywords
            )
            if has_partial_result:
                if self.searcher.save_recursive_results_to_excel(save_path, self._log_wrapper):
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
            # comment cleaned (encoding issue)
            if self.driver:
                try:
                    self.driver.quit()
                except:
                    pass
                self.driver = None

    def _log_wrapper(self, msg):
        """Text cleaned due to encoding issue."""
        self.log.emit(self.keyword, msg)

    def stop(self):
        """Text cleaned due to encoding issue."""
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


def create_chrome_driver(log_callback=None):
    """Chrome WebDriver를 exe 환경에서도 안정적으로 생성한다."""
    errors = []

    def _log(msg):
        if log_callback:
            try:
                log_callback(str(msg))
            except Exception:
                pass

    def _try_create(service=None):
        if service is None:
            driver = webdriver.Chrome(options=options)
        else:
            driver = webdriver.Chrome(service=service, options=options)
        driver.set_page_load_timeout(20)
        driver.implicitly_wait(5)
        driver.get("about:blank")
        return driver

    options = webdriver.ChromeOptions()
    options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--disable-extensions")
    options.add_argument("--disable-logging")
    options.add_argument("--log-level=3")
    options.add_argument("--window-size=1920,1080")

    # 1) 로컬/번들 chromedriver 우선
    local_candidates = []
    try:
        local_candidates.append(os.path.join(get_app_base_dir(), "chromedriver.exe"))
    except Exception:
        pass
    try:
        local_candidates.append(os.path.join(os.getcwd(), "chromedriver.exe"))
    except Exception:
        pass
    try:
        local_candidates.append(os.path.join(os.path.dirname(sys.executable), "chromedriver.exe"))
    except Exception:
        pass
    try:
        meipass_dir = getattr(sys, "_MEIPASS", None)
        if meipass_dir:
            local_candidates.append(os.path.join(meipass_dir, "chromedriver.exe"))
    except Exception:
        pass

    for candidate in local_candidates:
        if not candidate or not os.path.exists(candidate):
            continue
        try:
            _log(f"로컬 ChromeDriver 사용: {candidate}")
            return _try_create(Service(candidate))
        except Exception as e:
            errors.append(f"local({candidate}): {e}")

    # 2) webdriver-manager
    try:
        _log("webdriver-manager로 ChromeDriver 설치 시도")
        driver_path = ChromeDriverManager().install()
        _log(f"설치된 ChromeDriver: {driver_path}")
        return _try_create(Service(driver_path))
    except Exception as e:
        errors.append(f"webdriver_manager: {e}")

    # 3) PATH chromedriver
    try:
        import shutil
        path_driver = shutil.which("chromedriver")
        if path_driver:
            _log(f"PATH ChromeDriver 사용: {path_driver}")
            return _try_create(Service(path_driver))
    except Exception as e:
        errors.append(f"path: {e}")

    # 4) Selenium Manager 자동 설치(최종 fallback)
    try:
        _log("Selenium Manager fallback 시도")
        return _try_create()
    except Exception as e:
        errors.append(f"selenium_manager: {e}")

    chrome_paths = [
        os.path.join(os.environ.get("PROGRAMFILES", r"C:\Program Files"), "Google", "Chrome", "Application", "chrome.exe"),
        os.path.join(os.environ.get("PROGRAMFILES(X86)", r"C:\Program Files (x86)"), "Google", "Chrome", "Application", "chrome.exe"),
        os.path.join(os.environ.get("LOCALAPPDATA", ""), "Google", "Chrome", "Application", "chrome.exe"),
    ]
    chrome_installed = any(p and os.path.exists(p) for p in chrome_paths)
    detail = "\n".join(errors[-5:]) if errors else "unknown"
    if not chrome_installed:
        raise RuntimeError(
            "Chrome 브라우저가 설치되어 있지 않거나 경로를 찾지 못했습니다.\n"
            "Google Chrome 설치 후 다시 실행해주세요.\n\n"
            f"{detail}"
        )
    raise RuntimeError(f"Chrome WebDriver 생성 실패\n{detail}")


class KeywordExtractorMainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        
        # comment cleaned (encoding issue)
        icon_path = get_icon_path()
        if icon_path:
            self.setWindowIcon(QIcon(icon_path))
            safe_print(f"아이콘 설정 완료: {icon_path}")
        else:
            safe_print("아이콘 파일을 찾을 수 없습니다.")
        
        self.setWindowTitle("네이버 연관키워드 추출기")
        self.resize(1200, 800) # 기본 기 정
        self.showMaximized()   # 로그램 행 체 면로 작
        
        # comment cleaned (encoding issue)
        self.settings = Settings()
        self.driver = None
        # comment cleaned (encoding issue)
        self.active_threads = []     #   
        self.completed_threads = 0   #   
        self.total_threads = 0       #   
        self.stop_requested = False
        
        # comment cleaned (encoding issue)
        self.setup_crash_protection()
        
        # comment cleaned (encoding issue)
        self.init_ui()
        self.setup_chrome_driver()
        
        # comment cleaned (encoding issue)
        self.setStyleSheet(STYLESHEET)

    def check_license_info(self):
        """Text cleaned due to encoding issue."""
        # comment cleaned (encoding issue)
        machine_id = get_machine_id()
        
        # comment cleaned (encoding issue)
        expiration_date = check_license_from_sheet(machine_id)
        
        if expiration_date:
            try:
                # comment cleaned (encoding issue)
                exp_date = datetime.strptime(str(expiration_date).strip(), '%Y-%m-%d')
                today = datetime.now()
                
                if exp_date < today:
                    # comment cleaned (encoding issue)
                    self.show_license_dialog(machine_id, expired=True)
                else:
                    # comment cleaned (encoding issue)
                    self.usage_label.setText(f"사용 기간: {expiration_date}까지")
                    self.usage_label.setStyleSheet(f"font-size: 13px; font-weight: bold; color: {NAVER_GREEN};")
            except:
                # comment cleaned (encoding issue)
                # comment cleaned (encoding issue)
                self.usage_label.setText(f"사용 기간: {expiration_date}")
                self.usage_label.setStyleSheet(f"font-size: 13px; font-weight: bold; color: {NAVER_GREEN};")
        else:
            # comment cleaned (encoding issue)
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
        """Text cleaned due to encoding issue."""
        global _current_window
        _current_window = self
        sys.excepthook = handle_exception
        
        # comment cleaned (encoding issue)
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
        """Text cleaned due to encoding issue."""
        # comment cleaned (encoding issue)
        # comment cleaned (encoding issue)
        # comment cleaned (encoding issue)
        pass

    def init_ui(self):
        """Initialize UI."""
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        central_layout = QVBoxLayout(central_widget)
        central_layout.setContentsMargins(0, 0, 0, 0)
        self.main_tabs = QTabWidget()
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

        # comment cleaned (encoding issue)
        header_widget = QWidget()
        header_widget.setStyleSheet(f"""
            QWidget {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0, 
                    stop:0 {NAVER_GREEN_LIGHT}, stop:0.5 {BACKGROUND_CARD}, stop:1 {NAVER_GREEN_LIGHT});
                border-radius: 15px;
                border: 2px solid {BORDER_COLOR};
            }}
            QLabel {{
                background: transparent;
                border: none;
            }}
        """)
        header_layout = QHBoxLayout(header_widget)
        header_layout.setContentsMargins(20, 15, 20, 15)
        
        # comment cleaned (encoding issue)
        title_label = QLabel("네이버 연관키워드 추출기")
        title_label.setStyleSheet(f"font-size: 16px; font-weight: 800; color: {NAVER_GREEN};")
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        # comment cleaned (encoding issue)
        self.usage_label = QLabel("사용 기간: 확인 중...")
        self.usage_label.setStyleSheet(f"font-size: 13px; font-weight: bold; color: #555555;")
        
        # comment cleaned (encoding issue)
        header_layout.addStretch()
        header_layout.addWidget(title_label)
        header_layout.addStretch()
        header_layout.addWidget(self.usage_label)
        
        main_layout.addWidget(header_widget)
        
        # comment cleaned (encoding issue)
        top_section_layout = QHBoxLayout()
        
        # comment cleaned (encoding issue)
        self.setup_search_section(top_section_layout)
        
        # comment cleaned (encoding issue)
        self.setup_progress_section(top_section_layout)
        
        main_layout.addLayout(top_section_layout)
        
        # comment cleaned (encoding issue)
        self.setup_save_section(main_layout)
        self.main_tabs.addTab(extractor_tab, "연관 키워드 추출")
            
        # comment cleaned (encoding issue)
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("준비 완료")

    def setup_search_section(self, main_layout):
        """검색 섹션 설정"""
        search_group = QGroupBox("키워드 검색")
        search_layout = QVBoxLayout(search_group)
        
        # comment cleaned (encoding issue)
        self.search_input = MultiKeywordTextEdit()
        self.search_input.setPlaceholderText(
            "사용 방법\n"
            "1. 키워드를 한 줄에 하나씩 입력하세요.\n"
            "2. Enter로 바로 추출 시작, Shift+Enter로 줄바꿈합니다.\n"
            "3. 여러 키워드를 동시에 병렬 처리합니다."
        )
        # comment cleaned (encoding issue)
        self.search_input.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self.search_input.search_requested.connect(self.start_search)
        search_layout.addWidget(self.search_input)
        
        # comment cleaned (encoding issue)
        search_layout.setContentsMargins(10, 10, 10, 10)
        search_layout.setSpacing(5)
        
        # comment cleaned (encoding issue)
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
        save_layout = QVBoxLayout(save_group)
        
        path_layout = QHBoxLayout()
        self.save_path_input = QLineEdit()
        self.save_path_input.setReadOnly(True)
        self.save_path_input.setPlaceholderText("저장할 폴더를 선택하세요.")
        
        self.browse_button = QPushButton("폴더 선택")
        self.browse_button.clicked.connect(self.change_save_directory)
        
        path_layout.addWidget(self.save_path_input)
        path_layout.addWidget(self.browse_button)
        
        # comment cleaned (encoding issue)
        saved_dir = self.settings.get_save_dir()
        if saved_dir and os.path.exists(saved_dir):
            self.save_path_input.setText(saved_dir)
        else:
            # comment cleaned (encoding issue)
            desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
            default_dir = os.path.join(desktop_path, "keyword_results")
            
            try:
                os.makedirs(default_dir, exist_ok=True)
            except Exception:
                # comment cleaned (encoding issue)
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
        progress_layout = QVBoxLayout(progress_group)
        progress_layout.setContentsMargins(10, 10, 10, 10) # 백 치
        
        # comment cleaned (encoding issue)
        self.progress_tabs = QTabWidget()
        self.progress_tabs.setStyleSheet("""
            QTabWidget::pane { border: 1px solid #CCCCCC; border-radius: 8px; }
            QTabBar::tab { background: #f0f0f0; padding: 8px 12px; border-top-left-radius: 6px; border-top-right-radius: 6px; margin-right: 2px; }
            QTabBar::tab:selected { background: #E6F0FD; font-weight: bold; color: #1E6ECA; border-bottom: 2px solid #1E6ECA; }
        """)
        
        # comment cleaned (encoding issue)
        self.total_log_text = SmartProgressTextEdit(min_height=100, max_height=800)
        self.total_log_text.setReadOnly(True)
        self.total_log_text.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self.total_log_text.setPlaceholderText("여기에 전체 진행 로그가 표시됩니다.")
        self.progress_tabs.addTab(self.total_log_text, "전체 로그")
        
        # comment cleaned (encoding issue)
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

        
        # comment cleaned (encoding issue)
        self.start_button.setEnabled(False)
        self.pause_button.setEnabled(True)
        self.pause_button.setText("일시정지")
        self.stop_button.setEnabled(True)
        self.search_input.setEnabled(False)
        # comment cleaned (encoding issue)
        self.status_bar.showMessage("키워드 추출 중...")
        
        self.update_progress(f"키워드 추출 작업 시작 (총 {len(keywords)}개)")
        self.update_progress(f"입력 키워드: {', '.join(keywords)}")
        self.update_progress(f"저장 폴더: {save_dir}")
        self.update_progress("병렬 처리 모드로 실행합니다.")
        
        # comment cleaned (encoding issue)
        while self.progress_tabs.count() > 1:
            self.progress_tabs.removeTab(1)
            
        self.log_widgets = {"전체": self.total_log_text}
        self.total_log_text.clear()
        
        # comment cleaned (encoding issue)
        self.active_threads = []
        self.completed_threads = 0
        self.total_threads = len(keywords)
        self.stop_requested = False
        
        # comment cleaned (encoding issue)
        for keyword in keywords:
            log_widget = SmartProgressTextEdit(min_height=100, max_height=800)
            log_widget.setReadOnly(True)
            self.progress_tabs.addTab(log_widget, keyword)
            self.log_widgets[keyword] = log_widget
            
            # comment cleaned (encoding issue)
            if keywords.index(keyword) == 0:
                self.progress_tabs.setCurrentIndex(1)
        
        for i, keyword in enumerate(keywords):
            thread = ParallelKeywordThread(keyword, save_dir, True)
            thread.finished.connect(self.on_thread_finished)
            thread.error.connect(self.on_thread_error)
            thread.log.connect(self.update_progress) # 그매핑 동 처리
            
            self.active_threads.append(thread)
            thread.start()
            self.update_progress(keyword, f"'{keyword}' 작업 시작...")

    def on_thread_finished(self, save_path):
        """Text cleaned due to encoding issue."""
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
        """Text cleaned due to encoding issue."""
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
        """Text cleaned due to encoding issue."""
        current_time = datetime.now().strftime('%H:%M:%S')
        
        # comment cleaned (encoding issue)
        if message is None:
            # comment cleaned (encoding issue)
            target_keyword = "전체"
            msg_content = keyword_or_msg
        else:
            # comment cleaned (encoding issue)
            target_keyword = keyword_or_msg
            msg_content = message

        target_keyword = sanitize_display_text(target_keyword)
        msg_content = sanitize_display_text(msg_content)
        formatted_message = f"[{current_time}] {msg_content}"
        
        # comment cleaned (encoding issue)
        if target_keyword in self.log_widgets:
            widget = self.log_widgets[target_keyword]
            if hasattr(widget, 'append_with_smart_scroll'):
                widget.append_with_smart_scroll(formatted_message)
            else:
                widget.append(formatted_message)
        
        # comment cleaned (encoding issue)
        if target_keyword != "전체":
            formatted_total_msg = f"[{current_time}] [{target_keyword}] {msg_content}"
            if hasattr(self.total_log_text, 'append_with_smart_scroll'):
                self.total_log_text.append_with_smart_scroll(formatted_total_msg)
            else:
                self.total_log_text.append(formatted_total_msg)

    def pause_resume_search(self):
        """Text cleaned due to encoding issue."""
        # comment cleaned (encoding issue)
        # comment cleaned (encoding issue)
        # comment cleaned (encoding issue)
        pass

    def on_search_paused(self):
        """Text cleaned due to encoding issue."""
        self.pause_button.setText("재개")
        self.status_bar.showMessage("키워드 추출이 일시정지되었습니다.")
    
    def on_search_resumed(self):
        """Text cleaned due to encoding issue."""
        self.pause_button.setText("일시정지")
        self.status_bar.showMessage("키워드 추출 중...")

    def reset_ui_state(self):
        """Text cleaned due to encoding issue."""
        self.start_button.setEnabled(True)
        self.pause_button.setEnabled(False)
        self.pause_button.setText("일시정지")
        self.stop_button.setEnabled(False)
        self.search_input.setEnabled(True)

    def closeEvent(self, event):
        """Handle window close event and cleanup."""
        global _crash_save_enabled
        
        # comment cleaned (encoding issue)
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
                    
                    # comment cleaned (encoding issue)
                    for thread in running_threads:
                        thread.stop()
                        thread.wait(1000)
                    
                    self.active_threads = []
                else:
                    event.ignore()
                    return

        _crash_save_enabled = False
        
        if self.driver:
            pass  # 메인 라버상 용 음
        
        event.accept()


def main():
    if sys.platform == 'win32':
        import ctypes
        ctypes.windll.kernel32.SetConsoleOutputCP(65001)
        ctypes.windll.kernel32.SetConsoleCP(65001)

    if not verify_machine_id_guard():
        sys.exit(1)
    
    app = QApplication(sys.argv)
    
    # comment cleaned (encoding issue)
    icon_path = get_icon_path()
    if icon_path:
        app.setWindowIcon(QIcon(icon_path))
        safe_print(f"아이콘 설정 완료: {icon_path}")
    else:
        safe_print("아이콘 파일을 찾을 수 없습니다.")
    
    app.setApplicationName("네이버 연관키워드 추출기")
    
    # comment cleaned (encoding issue)
    machine_id = get_machine_id()
    safe_print(f"Machine ID: {machine_id}")
    
    # comment cleaned (encoding issue)
    # comment cleaned (encoding issue)
    expiry_date_str = check_license_from_sheet(machine_id)
    
    if expiry_date_str:
        try:
            # comment cleaned (encoding issue)
            expiry_date = datetime.strptime(expiry_date_str, '%Y-%m-%d')
            current_date = datetime.now()
            
            # comment cleaned (encoding issue)
            if current_date > expiry_date + pd.Timedelta(days=1):
                safe_print(f"라이선스 만료: {expiry_date_str}")
                app_dummy = QApplication.instance() or QApplication(sys.argv)
                
                # comment cleaned (encoding issue)
                dialog = ExpiredDialog(expiry_date_str)
                dialog.exec()
                sys.exit(0)
            
            # comment cleaned (encoding issue)
            safe_print(f"라이선스 확인 완료: {expiry_date_str}")
            window = KeywordExtractorMainWindow()
            window.usage_label.setText(f"사용 기간: {expiry_date_str}")
            window.usage_label.setStyleSheet(f"font-size: 13px; font-weight: bold; color: {NAVER_GREEN};")
            window.show()
            
            try:
                sys.exit(app.exec())
            except Exception as e:
                emergency_save_data()
                raise
                
        except ValueError:
            # comment cleaned (encoding issue)
            safe_print(f"라이선스 날짜 형식 확인 필요: {expiry_date_str}")
            window = KeywordExtractorMainWindow()
            window.usage_label.setText(f"사용 기간: {expiry_date_str}")
            window.show()
            sys.exit(app.exec())
            
    else:
        # comment cleaned (encoding issue)
        safe_print("미등록 기기 - 실행 차단")
        dialog = UnregisteredDialog(machine_id)
        dialog.exec()
        sys.exit(0)

if __name__ == "__main__":
    main() 







