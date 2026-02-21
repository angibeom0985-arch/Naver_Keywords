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
from webdriver_manager.chrome import ChromeDriverManager

# BeautifulSoup for HTML parsing (釉뚮씪?곗? ?놁씠 HTML ?뚯떛)
try:
    from bs4 import BeautifulSoup
    BEAUTIFULSOUP_AVAILABLE = True
except ImportError:
    print("BeautifulSoup가 설치되지 않았습니다: pip install beautifulsoup4")
    BEAUTIFULSOUP_AVAILABLE = False

# 釉뚮씪?곗? ?놁씠 ?묒뾽?섍린 ?꾪빐 selenium ?쒓굅

# Qt ?뚮윭洹몄씤 寃쎈줈 ?ㅼ젙 (PyQt6 ?ㅻ쪟 ?닿껐)
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
        # ???寃쎈줈 ?쒕룄
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

# BeautifulSoup for HTML parsing (釉뚮씪?곗? ?놁씠 HTML ?뚯떛)
try:
    from bs4 import BeautifulSoup
    BEAUTIFULSOUP_AVAILABLE = True
except ImportError:
    print("BeautifulSoup가 설치되지 않았습니다: pip install beautifulsoup4")
    BEAUTIFULSOUP_AVAILABLE = False

# ?ㅼ썙??異붿텧 ?꾩슜 - OpenAI API 遺덊븘??
OPENAI_AVAILABLE = False

# ?ㅼ씠踰?釉뚮옖??湲곕컲 議고솕濡쒖슫 ?됱긽 ?붾젅??
NAVER_GREEN = "#03c75a"              # 硫붿씤 ?ㅼ씠踰?洹몃┛
NAVER_GREEN_DARK = "#028a4a"         # 吏꾪븳 洹몃┛ (hover)
NAVER_GREEN_LIGHT = "#e8f5f0"        # ?고븳 洹몃┛ (諛곌꼍)
NAVER_GREEN_ULTRA_LIGHT = "#f0faf7"  # 留ㅼ슦 ?고븳 洹몃┛ (?꾩껜 諛곌꼍)
WHITE_COLOR = "#ffffff"               # ?쒕갚??
TEXT_PRIMARY = "#212529"             # 吏꾪븳 ?띿뒪??
TEXT_SECONDARY = "#6c757d"           # 蹂댁“ ?띿뒪??
BACKGROUND_MAIN = "#f0faf7"          # 硫붿씤 諛곌꼍 (?고븳 洹몃┛)
BACKGROUND_CARD = "#ffffff"          # 移대뱶 諛곌꼍
BORDER_COLOR = "#d4edda"             # ?고븳 洹몃┛ ?뚮몢由?
BORDER_FOCUS = "#03c75a"             # ?ъ빱???뚮몢由?
PLACEHOLDER_COLOR = "#8a8a8a"        # ?먮━?쒖떆??

# ?꾩뿭 蹂?? ?щ옒??蹂댄샇瑜??꾪븳 ?꾩옱 ?묒뾽 ?곹깭 異붿쟻
_current_window = None
_crash_save_enabled = True
MACHINE_ID_GUARD_HASH = "9491ed89095c9822c512bd386b2a54102992e3466af1d351361903eacb79f585"
MACHINE_ID_APPROVAL_FILE = "machine_id_change_approval.txt"
MACHINE_ID_APPROVAL_TOKEN = "I_APPROVE_MACHINE_ID_CHANGE"
# ?꾩씠肄?寃쎈줈 ?ㅼ젙 (exe ?뚯씪 吏??
def get_icon_path():
    """?꾩씠肄??뚯씪 寃쎈줈瑜?諛섑솚 (exe? py 紐⑤몢 吏?? - ?낅┰ ?ㅽ뻾 媛쒖꽑"""
    try:
        # PyInstaller濡?鍮뚮뱶??exe ?뚯씪??寃쎌슦 (理쒖슦??
        if hasattr(sys, '_MEIPASS'):
            # PyInstaller媛 ?앹꽦???꾩떆 ?대뜑?먯꽌 李얘린
            icon_path = os.path.join(sys._MEIPASS, 'auto_naver.ico')
            if os.path.exists(icon_path):
                return icon_path
        
        # exe ?뚯씪怨?媛숈? ?꾩튂?먯꽌 李얘린 (諛고룷 ??
        if getattr(sys, 'frozen', False):
            # exe ?뚯씪???덈뒗 ?붾젆?좊━
            exe_dir = os.path.dirname(sys.executable)
            icon_path = os.path.join(exe_dir, 'auto_naver.ico')
            if os.path.exists(icon_path):
                return icon_path
        
        # ?쇰컲 Python ?ㅽ겕由쏀듃 ?ㅽ뻾??寃쎌슦
        script_dir = os.path.dirname(os.path.abspath(__file__))
        icon_path = os.path.join(script_dir, 'auto_naver.ico')
        if os.path.exists(icon_path):
            return icon_path
        
        # assets ?대뜑?먯꽌 李얘린
        assets_icon = os.path.join(script_dir, 'assets', 'auto_naver.ico')
        if os.path.exists(assets_icon):
            return assets_icon
        
        # ?꾩옱 ?묒뾽 ?붾젆?좊━?먯꽌 李얘린
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
    """援ш? ?쒗듃?먯꽌 ?쇱씠?좎뒪 ?뺣낫 ?뺤씤"""
    sheet_url = "https://docs.google.com/spreadsheets/d/10-AseeTNvE97wo29HT2ajui918bg5ICj5L9UOYV0NBo/export?format=csv&gid=0"
    try:
        safe_print(f"라이선스 확인 중... ID: {machine_id}")
        response = requests.get(sheet_url, timeout=5)
        if response.status_code == 200:
            # CSV ?뚯떛
            df = pd.read_csv(io.StringIO(response.text))
            
            # 癒몄떊 ID 而щ읆 李얘린 (3踰덉㎏ 而щ읆 媛??
            if len(df.columns) >= 4:
                # 怨듬갚 ?쒓굅 諛?臾몄옄??蹂????鍮꾧탳
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
        
        # 1. 寃쎄퀬 ?꾩씠肄?諛??띿뒪??
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
        
        # 2. 釉붾（ 諛뺤뒪 ?곸뿭
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
        
        # 癒몄떊 ID ?낅젰李?+ 蹂듭궗 踰꾪듉
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
        
        # 3. ?섎떒 李멸퀬 臾멸뎄
        note_layout = QHBoxLayout()
        bulb_icon = QLabel("💡")
        bulb_icon.setStyleSheet("font-size: 16px; background-color: transparent;")
        note_text = QLabel("참고: PC를 변경하면 머신 ID가 바뀔 수 있습니다.")
        note_text.setStyleSheet("font-size: 13px; color: #888888; background-color: transparent;")
        note_layout.addWidget(bulb_icon)
        note_layout.addWidget(note_text)
        note_layout.addStretch()
        layout.addLayout(note_layout)
        
        # 4. ?뺤씤 踰꾪듉 (?곗륫 ?섎떒)
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
        clipboard.setText(self.id_input.text())
        sender = self.sender()
        if sender:
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
        
        # 1. 寃쎄퀬 ?꾩씠肄?諛??띿뒪??
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
        
        # 2. 踰좎씠吏 諛뺤뒪 ?곸뿭
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
        kakao_btn.setMinimumHeight(60)  # 紐낆떆???믪씠 ?ㅼ젙?쇰줈 ?ш린 ?뺣낫
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
        
        # 3. ?뺤씤 踰꾪듉 (?곗륫 ?섎떒)
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
        self.resize_step = 30  # ?ㅽ겕濡ㅻ떦 ?ш린 蹂?붾웾
        
    def wheelEvent(self, event):
        # Ctrl ?ㅺ? ?뚮┛ ?곹깭?먯꽌 留덉슦?????대깽??泥섎━
        if event.modifiers() & Qt.KeyboardModifier.ControlModifier:
            # ??諛⑺뼢 ?뺤씤
            delta = event.angleDelta().y()
            current_height = self.height()
            
            if delta > 0:  # ?꾨줈 ?ㅽ겕濡?(李??ш린 利앷?)
                new_height = min(current_height + self.resize_step, self.max_height)
            else:  # ?꾨옒濡??ㅽ겕濡?(李??ш린 媛먯냼)
                new_height = max(current_height - self.resize_step, self.min_height)
            
            # 理쒕? ?믪씠 ?ㅼ젙 ?낅뜲?댄듃
            self.setMaximumHeight(new_height)
            self.setMinimumHeight(new_height)
            
            # ?대깽??泥섎━ ?꾨즺
            event.accept()
        else:
            # ?쇰컲 ?ㅽ겕濡?泥섎━
            super().wheelEvent(event)


class SmartProgressTextEdit(ResizableTextEdit):
    """?ㅻ쭏???먮룞 ?ㅽ겕濡??쒖뼱媛 ?덈뒗 吏꾪뻾?곹솴 ?띿뒪???먮뵒??- 寃??湲곕뒫 ?ы븿"""
    
    def __init__(self, parent=None, min_height=200, max_height=800):
        super().__init__(parent, min_height, max_height)
        self.user_is_scrolling = False
        self.last_scroll_time = 0
        self.auto_scroll_enabled = True
        self.search_widget = None
        self.last_search_text = ""
        
        # ?ㅽ겕濡ㅻ컮 蹂寃??대깽???곌껐
        scrollbar = self.verticalScrollBar()
        if scrollbar:
            scrollbar.valueChanged.connect(self._on_scroll_changed)
            scrollbar.sliderPressed.connect(self._on_user_scroll_start)
            scrollbar.sliderReleased.connect(self._on_user_scroll_end)
        
        # Ctrl+F ?⑥텞???ㅼ젙
        self.search_shortcut = QShortcut(QKeySequence("Ctrl+F"), self)
        self.search_shortcut.activated.connect(self.show_search_dialog)
        
    def _on_scroll_changed(self, value):
        """?ㅽ겕濡??꾩튂 蹂寃????몄텧"""
        import time
        current_time = time.time()
        
        # ?ъ슜?먭? ?ㅽ겕濡?以묒씠 ?꾨땲怨? 留덉?留??ㅽ겕濡ㅻ줈遺??1珥덇? 吏?ъ쑝硫??먮룞 ?ㅽ겕濡??ы솢?깊솕
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
        
        # 3珥????먮룞 ?ㅽ겕濡??ы솢?깊솕
        QTimer.singleShot(3000, self._enable_auto_scroll)
        
    def _enable_auto_scroll(self):
        """?먮룞 ?ㅽ겕濡??ы솢?깊솕"""
        if not self.user_is_scrolling:
            self.auto_scroll_enabled = True
            
            # 3珥????먮룞 ?ㅽ겕濡??ы솢?깊솕
            QTimer.singleShot(3000, self._enable_auto_scroll)
            
    def wheelEvent(self, event):
        """Handle wheel event."""
        # ?ъ슜?먭? ?좊줈 ?ㅽ겕濡ㅽ븯??寃쎌슦 (Ctrl???뚮━吏 ?딆븯????
        if not (event.modifiers() & Qt.KeyboardModifier.ControlModifier):
            import time
            self.auto_scroll_enabled = False
            self.last_scroll_time = time.time()
            # 3珥????먮룞 ?ㅽ겕濡??ы솢?깊솕
            QTimer.singleShot(3000, self._enable_auto_scroll)
            
        super().wheelEvent(event)
        
    def append_with_smart_scroll(self, text):
        """?ㅻ쭏???ㅽ겕濡ㅼ씠 ?덈뒗 ?띿뒪??異붽?"""
        # ?ㅽ겕濡ㅻ컮媛 留??꾨옒???덈뒗吏 ?뺤씤
        scrollbar = self.verticalScrollBar()
        was_at_bottom = False
        if scrollbar:
            was_at_bottom = scrollbar.value() >= scrollbar.maximum() - 10
        
        # ?띿뒪??異붽?
        self.append(text)
        
        # ?먮룞 ?ㅽ겕濡ㅼ씠 ?쒖꽦?붾릺???덇퀬, ?댁쟾??留??꾨옒???덉뿀?ㅻ㈃ ?ㅽ겕濡?
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
        
        # ?꾩껜 ?띿뒪?몄뿉??寃??
        text_content = self.toPlainText()
        
        # ?꾩옱 而ㅼ꽌 ?꾩튂 媛?몄삤湲?
        cursor = self.textCursor()
        current_position = cursor.position()
        
        # ?꾩옱 ?꾩튂遺??寃??
        found_index = text_content.find(search_text, current_position)
        
        if found_index == -1:
            # 泥섏쓬遺???ㅼ떆 寃??
            found_index = text_content.find(search_text)
            
        if found_index != -1:
            # 寃??寃곌낵 ?섏씠?쇱씠??
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
        safe_print("?슚 ?묎툒 ????쒖옉...")
        
        saved_count = 0
        
        # ?쒖꽦 ?ㅻ젅???뺤씤 (蹂묐젹 泥섎━ 吏??
        if hasattr(_current_window, 'active_threads') and _current_window.active_threads:
            save_dir = _current_window.save_path_input.text().strip()
            if not save_dir:
                # ?ъ슜?먮퀎 諛뷀깢?붾㈃ 寃쎈줈 ?숈쟻 ?앹꽦
                desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
                save_dir = os.path.join(desktop_path, "keyword_results")
                try:
                    os.makedirs(save_dir, exist_ok=True)
                except Exception:
                    # 諛뷀깢?붾㈃ ?묎렐 ?ㅽ뙣 ????臾몄꽌濡??泥?
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
                        
                        # ?덉쟾???뚯씪紐??앹꽦
                        safe_keyword = re.sub(r'[^\w가-힣\s]', '', base_keyword).strip()[:20]
                        if not safe_keyword:
                            safe_keyword = "응급저장"
                        
                        emergency_file = os.path.join(save_dir, f"{safe_keyword}_응급저장_{current_time}.xlsx")
                        
                        # ?곗씠?????
                        if searcher.save_recursive_results_to_excel(emergency_file):
                            safe_print(f"???묎툒 ????꾨즺 ({base_keyword}): {emergency_file}")
                            saved_count += 1
                except Exception as inner_e:
                    safe_print(f"?좑툘 媛쒕퀎 ?ㅻ젅??????ㅽ뙣: {str(inner_e)}")
                    continue
            
            if saved_count > 0:
                safe_print(f"?뱤 珥?{saved_count}媛쒖쓽 ?묒뾽???묎툒 ??λ릺?덉뒿?덈떎.")
            else:
                safe_print("?좑툘 ??ν븷 ?곗씠?곌? ?녾굅???ㅽ뙣?덉뒿?덈떎.")
            
    except Exception as e:
        safe_print(f"???묎툒 ???珥덇린???ㅽ뙣: {str(e)}")
        
    except Exception as e:
        safe_print(f"???묎툒 ????ㅽ뙣: {str(e)}")
        # ?묎툒 ??λ룄 ?ㅽ뙣??寃쎌슦 理쒖냼??JSON?쇰줈?쇰룄 ????쒕룄
        try:
            if (_current_window and _current_window.search_thread and 
                hasattr(_current_window.search_thread, 'searcher') and
                hasattr(_current_window.search_thread.searcher, 'all_related_keywords')):
                
                backup_dir = os.path.join(os.getcwd(), "emergency_backup")
                os.makedirs(backup_dir, exist_ok=True)
                
                backup_file = os.path.join(backup_dir, f"emergency_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json")
                
                with open(backup_file, 'w', encoding='utf-8') as f:
                    json.dump(_current_window.search_thread.searcher.all_related_keywords, f, 
                             ensure_ascii=False, indent=2)
                
                safe_print(f"?뱞 JSON 諛깆뾽 ????꾨즺: {backup_file}")
        except:
            safe_print("??JSON 諛깆뾽 ??λ룄 ?ㅽ뙣")


def handle_exception(exc_type, exc_value, exc_traceback):
    """Handle uncaught exceptions and trigger emergency backup."""
    global _crash_save_enabled
    
    if _crash_save_enabled:
        safe_print("?슚 泥섎━?섏? ?딆? ?덉쇅 諛쒖깮!")
        safe_print(f"?덉쇅 ??? {exc_type.__name__}")
        safe_print(f"?덉쇅 ?댁슜: {str(exc_value)}")
        
        # ?묎툒 ????섑뻾
        emergency_save_data()
        
        # ?덉쇅 ?뺣낫瑜??뚯씪濡????
        try:
            crash_dir = os.path.join(os.getcwd(), "crash_logs")
            os.makedirs(crash_dir, exist_ok=True)
            
            current_time = datetime.now().strftime('%Y%m%d_%H%M%S')
            crash_file = os.path.join(
    crash_dir, f"crash_log_{current_time}.txt")
            
            with open(crash_file, 'w', encoding='utf-8') as f:
                f.write(f"?щ옒??諛쒖깮 ?쒓컙: {datetime.now()}\n")
                f.write(f"?덉쇅 ??? {exc_type.__name__}\n")
                f.write(f"?덉쇅 ?댁슜: {str(exc_value)}\n\n")
                f.write("?ㅽ깮 ?몃젅?댁뒪:\n")
                traceback.print_exception(
    exc_type, exc_value, exc_traceback, file=f)
            
            safe_print(f"?뱷 ?щ옒??濡쒓렇 ??? {crash_file}")
        except:
            pass
    
    # 湲곕낯 ?덉쇅 泥섎━湲??몄텧
    sys.__excepthook__(exc_type, exc_value, exc_traceback)


def handle_signal(signum, frame):
    """?쒓렇???몃뱾??(Ctrl+C, 媛뺤젣 醫낅즺 ??"""
    signal_names = {
        signal.SIGINT: "SIGINT (Ctrl+C)",
        signal.SIGTERM: "SIGTERM (醫낅즺 ?붿껌)"
    }
    
    signal_name = signal_names.get(signum, f"Signal {signum}")
    safe_print(f"?슚 {signal_name} ?좏샇 ?섏떊! ?묎툒 ???以?..")
    
    emergency_save_data()
    
    # ?뺤긽 醫낅즺
    if _current_window:
        _current_window.close()
    
    sys.exit(0)


class MultiKeywordTextEdit(QTextEdit):
    """?щ윭 ?ㅼ썙???낅젰???꾪븳 而ㅼ뒪? ?띿뒪???먮뵒??- paintEvent 湲곕컲 placeholder 諛?媛?낆꽦 媛쒖꽑"""
    search_requested = pyqtSignal()
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self._placeholder_text = ""
        
        # ?대?吏 諛?留곹겕 ?ㅼ젙
        self._cta_text = "키워드 공부하러 가기"
        self._cta_url = "https://cafe.naver.com/f-e/cafes/31118881/articles/2036?menuid=12&referrerAllArticles=false"
        self._link_rect = None
        
        # 二쇱냼 ?쒖떆以?而ㅼ꽌 蹂寃쎌쓣 ?꾪븳 ?몃옒???쒖꽦??
        self.setMouseTracking(True)
        
        # ?ㅽ겕濡??뺤콉 ?ㅼ젙
        self.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        
        # ?ㅼ씠踰?洹몃┛ ?뚮쭏 湲곕낯 ?ㅽ????곸슜 (?됯컙 180% 異붽?)
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
        """Placeholder ?띿뒪???ㅼ젙"""
        self._placeholder_text = text
        self.update()
        
    def paintEvent(self, event):
        """Custom paint event for placeholder and CTA."""
        super().paintEvent(event)
        
        # ?ъ빱?ㅺ? ?녾퀬 ?띿뒪?멸? 鍮꾩뼱?덉쓣 ?뚮쭔 placeholder ?쒖떆
        if not self.toPlainText().strip() and not self.hasFocus():
            painter = QPainter(self.viewport())
            painter.setRenderHint(QPainter.RenderHint.Antialiasing)
            
            # --- 1. ?띿뒪??洹몃━湲?---
            painter.setPen(QColor("#777777"))
            font = self.font()
            font.setPointSize(11)
            painter.setFont(font)
            
            lines = self._placeholder_text.split('\n')
            metrics = painter.fontMetrics()
            line_height = metrics.height()
            line_spacing = 2.2
            text_block_height = (len(lines) * line_height * line_spacing)
            
            viewport_rect = self.viewport().rect()
            padding_left = 60
            
            # ?띿뒪?? ?대?吏, 留곹겕 ?꾩껜 ?믪씠 怨꾩궛 (???
            link_h = 40
            spacing_between = 20
            
            # ?꾩껜 而⑦뀗痢좎쓽 ?쒖옉 Y (?붾㈃ 以묒븰 ?뺣젹)
            total_content_height = text_block_height + spacing_between + link_h
            start_y = (viewport_rect.height() - total_content_height) / 2 + metrics.ascent()
            
            current_y = start_y
            
            # ?띿뒪??洹몃━湲?
            for i, line in enumerate(lines):
                painter.drawText(int(viewport_rect.left() + padding_left), int(current_y), line)
                current_y += (line_height * line_spacing)
            
            current_y += spacing_between
            
            # --- 2. ?대?吏 洹몃━湲?(?쒓굅?? ---
            
            # --- 3. 留곹겕 洹몃━湲?---
            link_font = self.font()
            link_font.setPointSize(11)
            link_font.setUnderline(True)
            link_font.setBold(True)
            painter.setFont(link_font)
            painter.setPen(QColor("#0066CC")) # ?뚮???留곹겕
            
            link_metrics = painter.fontMetrics()
            link_width = link_metrics.horizontalAdvance(self._cta_text)
            link_x = (viewport_rect.width() - link_width) / 2
            
            painter.drawText(int(link_x), int(current_y + link_metrics.ascent()), self._cta_text)
            
            # 留곹겕 ?곸뿭 ???(?대┃ 媛먯???
            self._link_rect = QRect(int(link_x), int(current_y), int(link_width), int(link_metrics.height() + 10))

    def mouseMoveEvent(self, event):
        """Handle mouse move for link hover."""
        if self._link_rect and self._link_rect.contains(event.pos()) and not self.toPlainText().strip():
            self.viewport().setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        else:
            self.viewport().setCursor(QCursor(Qt.CursorShape.IBeamCursor))
        super().mouseMoveEvent(event)

    def mousePressEvent(self, event):
        """留덉슦???대┃ ??留곹겕 ?닿린"""
        if self._link_rect and self._link_rect.contains(event.pos()) and not self.toPlainText().strip():
            QDesktopServices.openUrl(QUrl(self._cta_url))
            return # 留곹겕 ?대┃ ???ъ빱???≪? ?딆쓬

        super().mousePressEvent(event)
    
    def keyPressEvent(self, event):
        """???낅젰 ?대깽??泥섎━"""
        if event.key() == Qt.Key.Key_Return or event.key() == Qt.Key.Key_Enter:
            if event.modifiers() == Qt.KeyboardModifier.ShiftModifier:
                # Shift+Enter: 以꾨컮轅?
                super().keyPressEvent(event)
            else:
                # Enter: 寃???쒖옉 ?좏샇 諛쒖깮
                self.search_requested.emit()
                event.accept()
        else:
            super().keyPressEvent(event)

    def wheelEvent(self, event):
        """留덉슦?????대깽??泥섎━ - Ctrl+?좊줈 ?ш린 議곗젅 湲곕뒫 異붽?"""
        # Ctrl ?ㅺ? ?뚮┛ ?곹깭?먯꽌 留덉슦?????대깽??泥섎━ (?ш린 議곗젅)
        if event.modifiers() & Qt.KeyboardModifier.ControlModifier:
            # ??諛⑺뼢 ?뺤씤
            delta = event.angleDelta().y()
            current_height = self.height()
            
            if delta > 0:  # ?꾨줈 ?ㅽ겕濡?(李??ш린 利앷?)
                new_height = min(current_height + self.resize_step, self.max_height)
            else:  # ?꾨옒濡??ㅽ겕濡?(李??ш린 媛먯냼)
                new_height = max(current_height - self.resize_step, self.min_height)
            
            # 理쒕?/理쒖냼 ?믪씠 ?ㅼ젙 ?낅뜲?댄듃
            self.setMaximumHeight(new_height)
            self.setMinimumHeight(new_height)
            
            # ?대깽??泥섎━ ?꾨즺
            event.accept()
        else:
            # ?쇰컲 ?ㅽ겕濡?泥섎━
            super().wheelEvent(event)


class NaverMobileSearchScraper:
    """釉뚮씪?곗? 湲곕컲 ?ㅼ씠踰??ㅼ썙??異붿텧 (媛쒖꽑??"""
    
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
        self.search_thread = None
        
        # User-Agent ?ㅼ젙 (?ㅼ젣 釉뚮씪?곗?泥섎읆 蹂댁씠寃?
        self.session.headers.update({
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8',
            'Accept-Language': 'ko-KR,ko;q=0.8,en-US;q=0.5,en;q=0.3',
            'Accept-Encoding': 'gzip, deflate, br',
            'DNT': '1',
            'Connection': 'keep-alive',
            'Upgrade-Insecure-Requests': '1',
        })

    def check_internet_connection(self):
        """?명꽣???곌껐 ?곹깭 ?뺤씤"""
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
        """?쇱떆?뺤? ?곹깭 泥댄겕 諛??명꽣???곌껐 ?뺤씤"""
        if not self.is_running:
            return False
        
        # ?쇱떆?뺤? ?곹깭 ?뺤씤
        if self.search_thread and hasattr(self.search_thread, 'is_paused'):
            if self.search_thread.is_paused:
                if progress_callback:
                    progress_callback("?몌툘 ?묒뾽???쇱떆?뺤??섏뿀?듬땲?? '?ш컻' 踰꾪듉???뚮윭二쇱꽭??")
                
                while self.search_thread.is_paused and self.is_running:
                    time.sleep(0.5)
                
                if not self.is_running:
                    return False
                
                if progress_callback:
                    progress_callback("?띰툘 ?묒뾽???ш컻?⑸땲??")
        
        # ?명꽣???곌껐 ?곹깭 ?뺤씤
        if not self.check_internet_connection():
            if progress_callback:
                progress_callback("?뙋 ?명꽣???곌껐???딆뼱議뚯뒿?덈떎. ?곌껐??湲곕떎由щ뒗 以?..")
            
            connection_wait_count = 0
            while not self.check_internet_connection() and self.is_running:
                time.sleep(2)
                connection_wait_count += 1
                
                if connection_wait_count % 5 == 0 and progress_callback:
                    progress_callback(f"?봽 ?명꽣???곌껐 ?쒕룄 以?.. ({connection_wait_count * 2}珥?寃쎄낵)")
            
            if not self.is_running:
                return False
            
            if self.check_internet_connection():
                if progress_callback:
                    progress_callback("???명꽣???곌껐??蹂듦뎄?섏뿀?듬땲?? ?묒뾽??怨꾩냽 吏꾪뻾?⑸땲??")
        
        return True

    def search_keyword(self, keyword, progress_callback=None):
        """?ㅼ씠踰꾩뿉???ㅼ썙??寃??(HTTP ?붿껌?쇰줈)"""
        try:
            if progress_callback:
                progress_callback(f"'{keyword}' 寃???쒖옉... (釉뚮씪?곗? ?놁씠)")
            
            encoded_keyword = urllib.parse.quote(keyword)
            search_url = f"https://search.naver.com/search.naver?where=nexearch&sm=top_hty&fbm=0&ie=utf8&query={encoded_keyword}"
            
            response = self.session.get(search_url, timeout=10)
            response.raise_for_status()
            
            if progress_callback:
                progress_callback("寃???꾨즺")
            
            return response.text
            
        except Exception as e:
            if progress_callback:
                progress_callback(f"寃???ㅻ쪟: {str(e)}")
            return None

    # extract_autocomplete_keywords (requests version) removed to avoid duplication
    pass

    def extract_related_keywords(self, keyword, progress_callback=None):
        """?곌?寃?됱뼱 異붿텧 (HTML ?뚯떛)"""
        keywords = []
        
        try:
            if not BEAUTIFULSOUP_AVAILABLE:
                if progress_callback:
                    progress_callback("??BeautifulSoup???ㅼ튂?섏? ?딆븯?듬땲??")
                return keywords
            
            html_content = self.search_keyword(keyword, progress_callback)
            if not html_content:
                return keywords
            
            soup = BeautifulSoup(html_content, 'html.parser')
            
            if progress_callback:
                progress_callback(f"'{keyword}' ?섏씠吏?먯꽌 ?곌?寃?됱뼱 異붿텧 以?..")
            
            # ?곌?寃?됱뼱 ?좏깮?먮뱾
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
                    progress_callback(f"?좏깮??'{selector}'?먯꽌 {len(elements)}媛??붿냼 諛쒓껄")
                
                for element in elements:
                    try:
                        keyword_text = element.get_text(strip=True)
                        if keyword_text and keyword_text not in keywords:
                            keywords.append(keyword_text)
                            found_count += 1
                            if progress_callback:
                                progress_callback(f"???곌??ㅼ썙??諛쒓껄 ({found_count}): {keyword_text}")
                    except:
                        continue
            
            if progress_callback:
                progress_callback(f"珥?{len(keywords)}媛쒖쓽 ?곌??ㅼ썙?쒕? 異붿텧?덉뒿?덈떎.")
            
            return keywords
            
        except Exception as e:
            if progress_callback:
                progress_callback(f"?곌??ㅼ썙??異붿텧 ?ㅻ쪟: {str(e)}")
            return []

    def check_internet_connection(self):
        """?명꽣???곌껐 ?곹깭 ?뺤씤"""
        try:
            # ?ㅼ씠踰꾩뿉 媛꾨떒???붿껌?쇰줈 ?곌껐 ?뺤씤
            response = requests.get("https://www.naver.com", timeout=5)
            return response.status_code == 200
        except:
            try:
                # ??덉쑝濡?援ш? DNS ?뺤씤
                import socket
                socket.create_connection(("8.8.8.8", 53), timeout=3)
                return True
            except:
                return False

    def check_pause_status(self, progress_callback=None):
        """?쇱떆?뺤? ?곹깭 泥댄겕 諛??명꽣???곌껐 ?뺤씤 - 媛쒖꽑??踰꾩쟾"""
        # 1. ?꾨줈洹몃옩 以묐떒 ?곹깭 ?뺤씤
        if not self.is_running:
            return False
        
        # 2. ?쇱떆?뺤? ?곹깭 ?뺤씤
        if self.search_thread and hasattr(self.search_thread, 'is_paused'):
            if self.search_thread.is_paused:
                if progress_callback:
                    progress_callback("?몌툘 ?묒뾽???쇱떆?뺤??섏뿀?듬땲?? '?ш컻' 踰꾪듉???뚮윭二쇱꽭??")
                
                # ?쇱떆?뺤? ?곹깭?먯꽌 ?湲?
                while self.search_thread.is_paused and self.is_running:
                    time.sleep(0.5)
                
                if not self.is_running:
                    return False
                
                if progress_callback:
                    progress_callback("?띰툘 ?묒뾽???ш컻?⑸땲??")
        
        # 3. ?명꽣???곌껐 ?곹깭 ?뺤씤
        if not self.check_internet_connection():
            if progress_callback:
                progress_callback("?뙋 ?명꽣???곌껐???딆뼱議뚯뒿?덈떎. ?곌껐??湲곕떎由щ뒗 以?..")
            
            # ?명꽣???곌껐??蹂듦뎄???뚭퉴吏 ?湲?
            connection_wait_count = 0
            while not self.check_internet_connection() and self.is_running:
                time.sleep(2)
                connection_wait_count += 1
                
                # ?쇱떆?뺤? ?곹깭???④퍡 ?뺤씤
                if self.search_thread and hasattr(self.search_thread, 'is_paused') and self.search_thread.is_paused:
                    if progress_callback:
                        progress_callback("인터넷 연결 대기 중 일시정지됨")
                    break
                
                # 10珥덈쭏???곌껐 ?쒕룄 硫붿떆吏
                if connection_wait_count % 5 == 0 and progress_callback:
                    progress_callback(f"?봽 ?명꽣???곌껐 ?쒕룄 以?.. ({connection_wait_count * 2}珥?寃쎄낵)")
            
            if not self.is_running:
                return False
            
            # ?명꽣?룹씠 ?ㅼ떆 ?곌껐??寃쎌슦
            if self.check_internet_connection():
                if progress_callback:
                    progress_callback("???명꽣???곌껐??蹂듦뎄?섏뿀?듬땲?? ?묒뾽??怨꾩냽 吏꾪뻾?⑸땲??")
        
        # 紐⑤뱺 ?뺤씤???꾨즺?섎㈃ ?뺤긽 吏꾪뻾
        return True

    def initialize_browser(self):
        """釉뚮씪?곗? 珥덇린??- 諛깃렇?쇱슫??紐⑤뱶 ?꾩슜 (媛쒖꽑??踰꾩쟾)"""
        try:
            safe_print("?봽 釉뚮씪?곗? 珥덇린?붾? ?쒖옉?⑸땲??.. (諛깃렇?쇱슫??紐⑤뱶)")
            
            driver_path = None
            
            # 1. 濡쒖뺄 ?쒕씪?대쾭 ?뺤씤 (媛???곗꽑 - 諛고룷 ?섍꼍 ???
            import shutil
            
            # EXE ?ㅽ뻾 ?꾩튂 ?먮뒗 ?꾩옱 ?묒뾽 ?붾젆?좊━ ?뺤씤
            base_paths = []
            if getattr(sys, 'frozen', False):
                base_paths.append(os.path.dirname(sys.executable))
                if hasattr(sys, '_MEIPASS'):
                    base_paths.append(sys._MEIPASS)
            base_paths.append(os.getcwd())
            
            for base_path in base_paths:
                local_driver = os.path.join(base_path, "chromedriver.exe")
                if os.path.exists(local_driver):
                    safe_print(f"?뱛 濡쒖뺄 ?쒕씪?대쾭 諛쒓껄: {local_driver}")
                    driver_path = local_driver
                    break
            
            # 2. ChromeDriverManager ?ъ슜 (濡쒖뺄???놁쓣 寃쎌슦)
            if not driver_path:
                try:
                    from webdriver_manager.chrome import ChromeDriverManager
                    safe_print("燧뉛툘 ChromeDriverManager濡??쒕씪?대쾭 ?ㅼ튂/?뺤씤 以?..")
                    # cache_valid_range=1濡??ㅼ젙?섏뿬 留ㅻ쾲 泥댄겕?섏? ?딅룄濡?理쒖쟻??
                    driver_path = ChromeDriverManager().install()
                    safe_print(f"???쒕씪?대쾭 寃쎈줈 ?뺣낫: {driver_path}")
                except Exception as e:
                    safe_print(f"?좑툘 ChromeDriverManager ?ㅽ뙣: {str(e)}")
            
            # 3. ?쒖뒪??PATH ?뺤씤 (理쒗썑???섎떒)
            if not driver_path and shutil.which("chromedriver"):
                driver_path = "chromedriver"
                safe_print("???쒖뒪??PATH?먯꽌 chromedriver 諛쒓껄")
            
            if not driver_path:
                raise Exception("ChromeDriver瑜?李얠쓣 ???놁뒿?덈떎.\n?꾨줈洹몃옩 ?대뜑??'chromedriver.exe'瑜??ｌ뼱二쇨굅??\n?명꽣???곌껐???뺤씤?댁＜?몄슂.")

            # Service ?ㅼ젙
            try:
                service = Service(driver_path)
            except:
                if driver_path == "chromedriver":
                    service = Service()
                else:
                    service = Service(executable_path=driver_path)

            # 肄섏넄 李??④린湲?(Windows ?꾩슜)
            if os.name == 'nt':
                try:
                    startupopt = subprocess.STARTUPINFO()
                    startupopt.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                    service.creation_flags = subprocess.CREATE_NO_WINDOW
                except:
                    pass

            options = webdriver.ChromeOptions()
            
            # ?ㅻ뱶由ъ뒪 紐⑤뱶 媛뺤젣 ?쒖꽦??(??긽 諛깃렇?쇱슫??紐⑤뱶)
            options.add_argument("--headless")  # ?ㅻ뱶由ъ뒪 紐⑤뱶 ?쒖꽦??
            safe_print("?뵁 諛깃렇?쇱슫??紐⑤뱶: 釉뚮씪?곗? 李쎌씠 ?④꺼吏묐땲??")
            
            options.add_argument("--window-size=1920,1080")  # ?쒖? FHD ?댁긽?꾨줈 ?ㅼ젙
            options.add_argument("--start-maximized")  # 釉뚮씪?곗? 理쒕???(?ㅻ뱶由ъ뒪?먯꽌???좏슚)
            
            # ?곗뒪?ы넲 User-Agent ?ㅼ젙?쇰줈 釉뚮씪?곗? 李??ш린??留욌뒗 諛섏쓳????吏??
            options.add_argument(
                "--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")

            # 諛섏쓳????吏?먯쓣 ?꾪븳 理쒖쟻???ㅼ젙
            options.add_argument("--disable-web-security")  # CORS ?댁젣
            options.add_argument("--allow-running-insecure-content")  # ?쇳빀 肄섑뀗痢??덉슜
            options.add_argument("--force-device-scale-factor=1")  # ?ㅼ????⑺꽣 怨좎젙
            options.add_argument("--disable-features=VizDisplayCompositor")  # ?뚮뜑留?理쒖쟻??
            
            # ?덉젙???μ긽 ?듭뀡??
            options.add_argument("--no-sandbox")
            options.add_argument("--disable-dev-shm-usage") 
            options.add_argument("--disable-gpu")
            options.add_argument("--disable-extensions")
            options.add_argument("--disable-plugins")
            options.add_argument("--disable-images")
            
            # ?깅뒫 理쒖쟻???듭뀡??(?띾룄 ?μ긽 + 以묐났 ?쒓굅)
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
            
            # ?쒓? 源⑥쭚 諛⑹? - 紐⑤뱺 濡쒓렇? ?먮윭 硫붿떆吏 ?꾩쟾 李⑤떒
            options.add_argument("--lang=en-US")  # ?곸뼱濡??ㅼ젙
            options.add_argument("--disable-logging")  # 紐⑤뱺 濡쒓렇 鍮꾪솢?깊솕
            options.add_argument("--disable-gpu-sandbox")
            options.add_argument("--log-level=3")  # ?ш컖???ㅻ쪟留?
            options.add_argument("--silent")  # 議곗슜??紐⑤뱶
            # 踰덉뿭 UI 諛??뚮뜑留?鍮꾪솢?깊솕
            options.add_argument("--disable-features=TranslateUI,VizDisplayCompositor")
            options.add_argument("--disable-ipc-flooding-protection")  # IPC ?뚮윭??蹂댄샇 鍮꾪솢?깊솕
            
            # ?깅뒫 理쒖쟻??諛?濡쒓렇 ?꾩쟾 李⑤떒
            options.add_experimental_option('useAutomationExtension', False)
            options.add_experimental_option("excludeSwitches", ["enable-logging", "enable-automation"])
            options.add_experimental_option("detach", True)
            options.add_argument("--disable-blink-features=AutomationControlled")
            
            # ?ㅽ듃?뚰겕 諛???꾩븘???ㅼ젙 (???곴레??
            options.add_argument("--network-timeout=15")  # 30珥???15珥?
            options.add_argument("--page-load-strategy=eager")
            options.add_argument("--timeout=15000")  # 15珥???꾩븘??
            options.add_argument("--dns-prefetch-disable")  # DNS ?꾨━?섏튂 鍮꾪솢?깊솕
            
            # ?쒕씪?대쾭 ?앹꽦
            self.driver = webdriver.Chrome(service=service, options=options)
            safe_print("??Chrome ?쒕씪?대쾭 ?앹꽦 ?깃났!")
            
            # WebDriver ??꾩븘???ㅼ젙 (理쒖쟻??
            self.driver.set_page_load_timeout(15)  # ?섏씠吏 濡쒕뵫 ??꾩븘??15珥덈줈 ?⑥텞
            self.driver.implicitly_wait(3)  # ?붿떆???湲?3珥덈줈 ?⑥텞
            
            safe_print("釉뚮씪?곗?媛 ?ㅽ뻾?섏뿀?듬땲??")
            
            # ?ㅼ씠踰??묒냽 ?쒕룄
            max_retries = 3
            for attempt in range(max_retries):
                try:
                    # ?ㅼ씠踰?紐⑤컮??踰꾩쟾?쇰줈 ?묒냽?섎릺 ?곗뒪?ы넲 User-Agent濡?諛섏쓳???쒖떆
                    self.driver.get("https://m.naver.com")
                    
                    # viewport 諛??섏씠吏 ?ㅽ??쇱쓣 釉뚮씪?곗? 李??ш린???꾩쟾??留욊쾶 理쒖쟻??
                    self.driver.execute_script("""
                        // 釉뚮씪?곗? 李??ш린 媛?몄삤湲?
                        var windowWidth = window.innerWidth;
                        var windowHeight = window.innerHeight;
                        
                        // viewport 硫뷀? ?쒓렇瑜?釉뚮씪?곗? 李??ш린??留욊쾶 ?ㅼ젙
                        var existingMeta = document.querySelector('meta[name="viewport"]');
                        if (existingMeta) {
                            existingMeta.remove();
                        }
                        var meta = document.createElement('meta');
                        meta.name = 'viewport';
                        meta.content = 'width=' + windowWidth + ', initial-scale=1.0, maximum-scale=3.0, user-scalable=yes';
                        document.getElementsByTagName('head')[0].appendChild(meta);
                        
                        // ?섏씠吏 ?꾩껜瑜?釉뚮씪?곗? 李??ш린??留욊쾶 議곗젙
                        document.documentElement.style.width = '100%';
                        document.documentElement.style.height = '100%';
                        document.body.style.minWidth = windowWidth + 'px';
                        document.body.style.width = '100%';
                        document.body.style.height = '100vh';
                        document.body.style.transform = 'none';
                        document.body.style.transformOrigin = 'top left';
                        document.body.style.margin = '0';
                        document.body.style.padding = '0';
                        document.body.style.fontSize = Math.max(14, windowWidth / 100) + 'px';  // 李??ш린???곕Ⅸ ?고듃 議곗젙
                        
                        // 硫붿씤 而⑦뀒?대꼫?ㅼ쓣 釉뚮씪?곗? ?꾩껜 ?ш린濡??뺤옣
                        var containers = document.querySelectorAll('.container, .wrap, .content_area, #wrap, .nx_wrap');
                        containers.forEach(function(container) {
                            container.style.maxWidth = '100%';
                            container.style.width = '100%';
                            container.style.minWidth = windowWidth + 'px';
                        });
                        
                        // 寃???곸뿭??釉뚮씪?곗? 李쎌뿉 留욊쾶 ?뺣?
                        var searchArea = document.querySelector('.TF7QLJYoGthrUnoIpxEj, .api_subject_bx, .search_result');
                        if (searchArea) {
                            searchArea.style.minHeight = (windowHeight - 200) + 'px';
                            searchArea.style.width = '100%';
                            searchArea.style.overflow = 'visible';
                            searchArea.style.maxWidth = 'none';
                        }
                        
                        console.log('?섏씠吏媛 釉뚮씪?곗? 李??ш린(' + windowWidth + 'x' + windowHeight + ')??留욊쾶 議곗젙?섏뿀?듬땲??');
                    """)
                    
                    # ?섏씠吏 濡쒕뵫 ?湲?
                    time.sleep(2)
                    
                    WebDriverWait(self.driver, 10).until(
                        EC.presence_of_element_located((By.TAG_NAME, "body"))
                    )
                    safe_print("?ㅼ씠踰?紐⑤컮?쇱뿉 ?묒냽?덉뒿?덈떎. (釉뚮씪?곗? 李??ш린??留욊쾶 理쒖쟻?붾맖)")
                    time.sleep(2)
                    return True
                except Exception as e:
                    if attempt < max_retries - 1:
                        safe_print(f"?ㅼ씠踰??묒냽 ?쒕룄 {attempt + 1} ?ㅽ뙣, ?ъ떆??以?..")
                        time.sleep(3)
                    else:
                        safe_print(f"?ㅼ씠踰??묒냽 理쒖쥌 ?ㅽ뙣: {str(e)}")
                        return False
            
            return True

        except Exception as e:
            error_msg = f"釉뚮씪?곗? 珥덇린???ㅻ쪟:\n{str(e)}\n\nChrome 釉뚮씪?곗?媛 ?ㅼ튂?섏뼱 ?덈뒗吏 ?뺤씤?댁＜?몄슂."
            safe_print(f"??{error_msg}")
            
            # GUI ?ㅻ젅?쒖뿉??硫붿떆吏 諛뺤뒪 ?쒖떆 ?쒕룄
            try:
                global _current_window
                if _current_window:
                    # 諛⑸쾿 1: 吏곸젒 ?쒖떆 (??대㉧ ?ъ슜?쇰줈 硫붿씤 猷⑦봽?먯꽌 ?ㅽ뻾?섎룄濡??좊룄)
                    # 硫붿씤 ?ㅻ젅?쒖뿉???ㅽ뻾?섏? ?딆쓣 ?꾪뿕???덉?留? 蹂댄넻 Qt??寃쎄퀬留??섍퀬 ?숈옉?섍굅???щ옒?쒕맖
                    # ?덉쟾???꾪빐 QMetaObject.invokeMethod媛 ?뺤꽍?댁?留?Python?먯꽌??蹂듭옟??
                    # QTimer.singleShot(0, ...) ?⑦꽩 ?ъ슜
                    from PyQt6.QtCore import QTimer
                    QTimer.singleShot(0, lambda: QMessageBox.critical(
                        _current_window, "釉뚮씪?곗? ?ㅻ쪟", error_msg))
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
                        progress_callback(f"'{keyword}' 寃???ъ떆??({attempt + 1}/{max_retries})...")
                    else:
                        progress_callback(f"'{keyword}' 寃???쒖옉...")
                
                encoded_keyword = urllib.parse.quote(keyword)
                search_url = f"https://m.search.naver.com/search.naver?where=m&sm=mtp_hty.top&query={encoded_keyword}"
                
                if self.driver:
                    self.driver.get(search_url)
                
                if progress_callback:
                    progress_callback("?섏씠吏 濡쒕뵫 以?..")
                
                # ?섏씠吏 濡쒕뵫 ?湲?(?덉쟾??理쒖쟻??
                time.sleep(random.uniform(1.5, 2.5))  # 遊??먯? 諛⑹? + ?곷떦??理쒖쟻??
                
                try:
                    if self.driver:
                        WebDriverWait(self.driver, 5).until(  # 8珥???5珥덈줈 ?⑥텞
                            EC.presence_of_element_located((By.TAG_NAME, "body"))
                        )
                except TimeoutException:
                    if attempt < max_retries - 1:
                        if progress_callback:
                            progress_callback(f"???섏씠吏 濡쒕뵫 ?쒓컙 珥덇낵 - ?ъ떆??以?..")
                        continue
                    else:
                        if progress_callback:
                            progress_callback("???섏씠吏 濡쒕뵫 ?쒓컙 珥덇낵 - ?대떦 ?ㅼ썙???ㅽ궢")
                        return False
                
                time.sleep(1)
                
                if progress_callback:
                    progress_callback("寃???꾨즺")
                
                return True
                
            except Exception as e:
                error_msg = str(e)
                
                # 1. invalid session id ?먮뒗 no such window ?ㅻ쪟 媛먯?
                if "invalid session id" in error_msg.lower() or "no such session" in error_msg.lower() or "no such window" in error_msg.lower():
                    if progress_callback:
                        progress_callback(f"?봽 ?щ＼ ?쒕씪?대쾭 ?몄뀡/李?臾몄젣 媛먯?. ?ъ떆??以?..")
                    
                    # ?쒕씪?대쾭 ?ъ떆???쒕룄
                    if self.initialize_browser():
                        if progress_callback:
                            progress_callback(f"???щ＼ ?쒕씪?대쾭 ?ъ떆???깃났. 寃???ъ떆??..")
                        continue
                    else:
                        if progress_callback:
                            progress_callback(f"???щ＼ ?쒕씪?대쾭 ?ъ떆???ㅽ뙣")
                        return False
                
                # 2. ?뚮뜑????꾩븘???ㅻ쪟 媛먯? 諛?泥섎━
                elif "timeout" in error_msg.lower() and "renderer" in error_msg.lower():
                    if progress_callback:
                        progress_callback(f"???뚮뜑????꾩븘??媛먯? (釉뚮씪?곗? ?묐떟 ?놁쓬)")
                    
                    if attempt < max_retries - 1:
                        if progress_callback:
                            progress_callback(f"?봽 ?쒕씪?대쾭 ?ъ떆?????ъ떆??..")
                        
                        # ?뚮뜑????꾩븘?껋쓽 寃쎌슦 ?쒕씪?대쾭 ?ъ떆?묒씠 ?④낵??
                        if self.initialize_browser():
                            if progress_callback:
                                progress_callback(f"???쒕씪?대쾭 ?ъ떆???꾨즺. 寃???ъ떆??..")
                            continue
                        else:
                            if progress_callback:
                                progress_callback(f"???쒕씪?대쾭 ?ъ떆???ㅽ뙣")
                    else:
                        if progress_callback:
                            progress_callback(f"???뚮뜑????꾩븘??理쒖쥌 ?ㅽ뙣 - ?대떦 ?ㅼ썙???ㅽ궢")
                        return False
                
                # 3. ?쇰컲?곸씤 ??꾩븘???ㅻ쪟
                elif "timeout" in error_msg.lower():
                    if progress_callback:
                        progress_callback(f"????꾩븘???ㅻ쪟 媛먯?")
                    
                    if attempt < max_retries - 1:
                        if progress_callback:
                            progress_callback(f"?깍툘 ?좎떆 ?湲????ъ떆??..")
                        time.sleep(5)  # ??꾩븘?껋쓽 寃쎌슦 議곌툑 ???湲?
                        continue
                    else:
                        if progress_callback:
                            progress_callback(f"????꾩븘??理쒖쥌 ?ㅽ뙣 - ?대떦 ?ㅼ썙???ㅽ궢")
                        return False
                
                # 4. 湲고? ?ㅻ쪟
                if attempt < max_retries - 1:
                    if progress_callback:
                        progress_callback(f"寃???ㅻ쪟 - ?ъ떆??以? {str(e)}")
                    time.sleep(3)
                    continue
                else:
                    if progress_callback:
                        progress_callback(f"寃??理쒖쥌 ?ㅽ뙣: {str(e)}")
                    return False
        
        return False

    def extract_autocomplete_keywords(self, keyword, progress_callback=None):
        """Extract autocomplete keywords."""
        keywords = []
        max_retries = 2
        
        for attempt in range(max_retries):
            try:
                if not self.driver:
                    # ?쒕씪?대쾭媛 ?놁쑝硫?珥덇린???쒕룄
                    if not self.initialize_browser():
                        return keywords
                
                if progress_callback:
                    if attempt > 0:
                        progress_callback(f"'{keyword}' ?먮룞?꾩꽦寃?됱뼱 異붿텧 ?ъ떆??({attempt + 1}/{max_retries})...")
                    else:
                        progress_callback(f"'{keyword}' ?먮룞?꾩꽦寃?됱뼱 異붿텧 ?쒖옉...")
                
                # ?ㅼ씠踰?硫붿씤 ?섏씠吏濡??대룞
                try:
                    self.driver.set_page_load_timeout(15)  # 15珥??쒗븳
                    self.driver.get("https://m.naver.com")
                except TimeoutException:
                    if progress_callback:
                        progress_callback("?좑툘 ?섏씠吏 濡쒕뵫 吏?? 怨꾩냽 吏꾪뻾?⑸땲??..")
                    try:
                        self.driver.execute_script("window.stop();")
                    except:
                        pass
                except Exception as e:
                    # ?대룞 以??먮윭 諛쒖깮 ??(no such window ?? ?덉쇅瑜??곸쐞濡??꾪뙆?섏뿬 泥섎━
                    raise e
                
                time.sleep(2)
            
                # ?섏씠吏 濡쒕뵫 ?湲?
                try:
                    from selenium.webdriver.support.ui import WebDriverWait
                    from selenium.webdriver.support import expected_conditions as EC
                    
                    # ??꾩븘???덉쇅 泥섎━ 異붽?
                    try:
                        wait = WebDriverWait(self.driver, 10)
                        wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
                        if progress_callback:
                            progress_callback("?ㅼ씠踰?硫붿씤 ?섏씠吏 濡쒕뵫 ?꾨즺")
                    except TimeoutException:
                        if progress_callback:
                            progress_callback("?좑툘 ?섏씠吏 ?붿냼 濡쒕뵫 ?쒓컙 珥덇낵 (臾댁떆?섍퀬 吏꾪뻾)")
                            
                except Exception as e:
                    if progress_callback:
                        progress_callback(f"?섏씠吏 濡쒕뵫 ?湲?以??ㅻ쪟: {str(e)}")
            
                # 寃?됱갹 李얘린
                search_input = None
                search_selectors = [
                    '#nx_query',
                    'input.search_input',
                    'input[name="query"]',
                    'input[type="search"]'
                ]
                
                for selector in search_selectors:
                    try:
                        search_input = self.driver.find_element(By.CSS_SELECTOR, selector)
                        if search_input and search_input.is_enabled():
                            if progress_callback:
                                progress_callback(f"寃?됱갹 諛쒓껄: {selector}")
                            break
                    except:
                        continue
                
                if not search_input:
                    if progress_callback:
                        progress_callback("??寃?됱갹??李얠쓣 ???놁뒿?덈떎.")
                    return keywords
            
                # 寃?됱갹???ㅼ썙???낅젰
                try:
                    # JavaScript濡??덉쟾?섍쾶 ?낅젰
                    self.driver.execute_script("""
                        var input = arguments[0];
                        var keyword = arguments[1];
                        input.focus();
                        input.value = keyword;
                        input.dispatchEvent(new Event('input', { bubbles: true }));
                        input.dispatchEvent(new Event('keyup', { bubbles: true }));
                    """, search_input, keyword)
                    
                    # ?먮룞?꾩꽦 濡쒕뵫 ?湲?
                    time.sleep(2)
                    
                    if progress_callback:
                        progress_callback(f"'{keyword}' ?낅젰 ?꾨즺, ?먮룞?꾩꽦 ?湲?以?..")
                        
                except Exception as input_error:
                    if progress_callback:
                        progress_callback(f"?ㅼ썙???낅젰 ?ㅽ뙣: {str(input_error)}")
                        return keywords
            
                # ?먮룞?꾩꽦 ?ㅼ썙??異붿텧
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
                        elements = self.driver.find_elements(By.CSS_SELECTOR, selector)
                        
                        for element in elements:
                            try:
                                if not element.is_displayed():
                                    continue
                                    
                                # ?띿뒪??異붿텧
                                keyword_text = element.get_attribute("textContent") or element.text
                                
                                if keyword_text:
                                    keyword_text = keyword_text.strip()
                                    
                                    # HTML ?쒓렇 ?쒓굅
                                    if '<' in keyword_text:
                                        import re
                                        keyword_text = re.sub(r'<[^>]+>', '', keyword_text)
                                        keyword_text = keyword_text.strip()
                                    
                                    # ?띿뒪???뺤젣
                                        keyword_text = self.clean_duplicate_text(keyword_text)
                                        
                                    # ?좏슚??寃利?
                                    if (keyword.lower() in keyword_text.lower() and 
                                        keyword_text not in keywords and
                                        len(keyword_text) <= 50 and
                                        len(keyword_text) > 1):
                                            keywords.append(keyword_text)
                                            found_count += 1
                                            if progress_callback:
                                                progress_callback(f"???먮룞?꾩꽦?ㅼ썙??諛쒓껄 ({found_count}): {keyword_text}")
                            except Exception:
                                continue
                    except Exception:
                        continue
                
                # 以묐났 ?쒓굅 諛??뺣젹
                keywords = list(set(keywords))
                keywords.sort()
                
                if progress_callback:
                    progress_callback(f"珥?{len(keywords)}媛쒖쓽 ?먮룞?꾩꽦?ㅼ썙?쒕? 異붿텧?덉뒿?덈떎.")
                
                return keywords
            
            except Exception as e:
                error_msg = str(e)
                if "no such window" in error_msg.lower() or "invalid session id" in error_msg.lower():
                    if progress_callback:
                        progress_callback("?좑툘 釉뚮씪?곗? 李쎌씠 ?ロ삍嫄곕굹 ?몄뀡??留뚮즺?섏뿀?듬땲?? 蹂듦뎄 ?쒕룄 以?..")
                    
                    if attempt < max_retries - 1:
                        if self.initialize_browser():
                            if progress_callback:
                                progress_callback("??釉뚮씪?곗? 蹂듦뎄 ?깃났. ?ъ떆?꾪빀?덈떎.")
                            continue
                
                if progress_callback:
                    progress_callback(f"?먮룞?꾩꽦?ㅼ썙??異붿텧 ?ㅻ쪟: {str(e)}")
                
                if attempt == max_retries - 1:
                    return []
        
        return keywords

    def extract_related_keywords_new(self, current_keyword, progress_callback=None):
        """Extract related keywords from current page."""
        keywords = []
        
        try:
            if not self.driver:
                return keywords
            
            # ?곌?寃?됱뼱 ?좏깮?먮뱾 - 2024??理쒖떊 ?ㅼ씠踰?紐⑤컮??
            related_selectors = [
                '#_related_keywords .keyword a',  # 硫붿씤 ?좏깮??
                '.related_srch .lst a',
                '.related_keyword a',
                '.lst_related a',
                '.keyword_area a',
                '.related_search a',
                '.keyword a'
            ]
            
            if progress_callback:
                progress_callback(f"'{current_keyword}' ?섏씠吏?먯꽌 ?곌?寃?됱뼱 異붿텧 ?쒖옉...")
            
            found_count = 0
            
            for selector in related_selectors:
                try:
                    elements = self.driver.find_elements(By.CSS_SELECTOR, selector)
                    if len(elements) > 0 and progress_callback:
                        progress_callback(f"?좏깮??'{selector}'?먯꽌 {len(elements)}媛??붿냼 諛쒓껄")
                    
                    for element in elements:
                        try:
                            # ???뺥솗???띿뒪??異붿텧???꾪븳 ?μ긽??濡쒖쭅
                            keyword_text = ""
                            
                            # 1李? 媛???덉쟾???띿뒪??異붿텧 - element.text 癒쇱? ?쒕룄
                            try:
                                keyword_text = element.text
                                if keyword_text:
                                    keyword_text = keyword_text.strip()
                            except:
                                keyword_text = ""

                            # 2李? textContent ?띿꽦?쇰줈 諛깆뾽 異붿텧
                            if not keyword_text:
                                try:
                                    keyword_text = element.get_attribute("textContent")
                                    if keyword_text:
                                        keyword_text = keyword_text.strip()
                                except:
                                    keyword_text = ""

                            # 3李? innerText ?띿꽦?쇰줈 諛깆뾽 異붿텧
                            if not keyword_text:
                                try:
                                    keyword_text = element.get_attribute("innerText")
                                    if keyword_text:
                                        keyword_text = keyword_text.strip()
                                except:
                                    keyword_text = ""

                            # 4李? JavaScript濡??뺥솗???띿뒪??異붿텧 (留덉?留??섎떒)
                            if not keyword_text:
                                try:
                                    keyword_text = self.driver.execute_script("""
                                        var element = arguments[0];
                                        if (!element) return '';
                                        
                                            // 留곹겕 ?붿냼??吏곸젒?곸씤 ?띿뒪?몃쭔 異붿텧
                                            var textContent = element.textContent || element.innerText || '';
                                            
                                            // ?욌뮘 怨듬갚 ?쒓굅 諛??곗냽 怨듬갚 ?뺣━
                                            return textContent.replace(/\\s+/g, ' ').trim();
                                    """, element)
                                except:
                                    keyword_text = ""
                            
                            # HTML ?쒓렇 ?쒓굅 諛??뱀닔臾몄옄 ?뺣━
                            if keyword_text:
                                import re
                                # HTML ?쒓렇 ?쒓굅
                                keyword_text = re.sub(r'<[^>]+>', '', keyword_text)
                                # ?곗냽??怨듬갚 ?쒓굅
                                keyword_text = re.sub(r'\s+', ' ', keyword_text)
                                # ?뱀닔臾몄옄 ?뺣━
                                keyword_text = re.sub(r'[\u200b-\u200d\ufeff]', '', keyword_text)  # ?쒕줈??臾몄옄 ?쒓굅
                                # 遺덉셿?꾪븳 ?띿뒪???뺣━ (?앹뿉 ?ㅻ뒗 遺덉셿?꾪븳 ?⑥뼱 ?쒓굅)
                                keyword_text = re.sub(r'\s+[가-힣]{1}$', '', keyword_text)
                                keyword_text = re.sub(r'\s+[a-zA-Z]{1}$', '', keyword_text)  # ?앹뿉 ?곷Ц 1湲?먮쭔 ?덈뒗 寃쎌슦 ?쒓굅
                                keyword_text = keyword_text.strip()
                                
                            if keyword_text:
                                # ?띿뒪???뺤젣
                                keyword_text = self.clean_duplicate_text(keyword_text)
                                    
                                # 異붽? 寃利? 遺덉셿?꾪븳 ?ㅼ썙???꾪꽣留?
                                # ?섎??덈뒗 ?⑥뼱濡??앸굹?붿? ?뺤씤
                                if keyword_text and not re.search(r'[가-힣]{1}$|[a-zA-Z]{1}$', keyword_text):
                                    # ?좏슚???ㅼ썙?쒖씤吏 ?뺤씤 (以묐났 ?쒓굅 諛?湲몄씠 泥댄겕)
                                    if (keyword_text not in keywords and
                                        len(keyword_text) <= 50 and
                                        len(keyword_text) > 1):
                                        keywords.append(keyword_text)
                                        found_count += 1
                                        if progress_callback:
                                            progress_callback(f"???곌??ㅼ썙??諛쒓껄 ({found_count}): {keyword_text}")
                                elif keyword_text and len(keyword_text) > 3:  # 3湲???댁긽?대㈃ ?덉슜
                                    if (keyword_text not in keywords and 
                                        len(keyword_text) <= 50 and 
                                        len(keyword_text) > 1):
                                        keywords.append(keyword_text)
                                        found_count += 1
                                        if progress_callback:
                                            progress_callback(f"???곌??ㅼ썙??諛쒓껄 ({found_count}): {keyword_text}")
                        except Exception as e:
                            continue
                except Exception as e:
                    continue
            
            # 以묐났 ?쒓굅 諛??뺣젹
            keywords = list(set(keywords))
            keywords.sort()
            
            if progress_callback:
                progress_callback(f"珥?{len(keywords)}媛쒖쓽 ?곌??ㅼ썙?쒕? 異붿텧?덉뒿?덈떎.")
            
            return keywords
            
        except Exception as e:
            if progress_callback:
                progress_callback(f"?곌??ㅼ썙??異붿텧 ?ㅻ쪟: {str(e)}")
            return []

    def clean_duplicate_text(self, text):
        """?띿뒪???뺣━ 諛?以묐났 ?쒓굅 - 媛쒖꽑??踰꾩쟾"""
        if not text:
            return text

        text = text.strip()
        # ?곗냽??怨듬갚???섎굹濡??듯빀
        text = re.sub(r'\s+', ' ', text)
        
        # ?⑥뼱 ?⑥쐞濡?遺꾨━?섏뿬 以묐났 ?쒓굅
        words = text.split()
        unique_words = []
        seen_words = set()
        
        for word in words:
            word_lower = word.lower()
            if word_lower not in seen_words:
                unique_words.append(word)
                seen_words.add(word_lower)
        
        # 寃곌낵 ?띿뒪???ъ“??
        result = ' '.join(unique_words)
        return result

    def extract_together_keywords(self, current_keyword, progress_callback=None):
        """?④퍡 留롮씠 李얜뒗 ?ㅼ썙??異붿텧"""
        keywords = []
        try:
            if progress_callback:
                progress_callback(f"'{current_keyword}' ?④퍡 留롮씠 李얜뒗 ?ㅼ썙??異붿텧 以?..")
            
            # 媛꾨떒??CSS ?좏깮?먮줈 ?붿냼 李얘린
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
                progress_callback(f"?④퍡 留롮씠 李얜뒗 ?ㅼ썙??{len(keywords)}媛?異붿텧")
        
            return list(set(keywords))
        except Exception as e:
            if progress_callback:
                progress_callback(f"?④퍡 留롮씠 李얜뒗 ?ㅼ썙??異붿텧 ?ㅻ쪟: {str(e)}")
            return []
            
    def extract_popular_topics(self, current_keyword, progress_callback=None):
        """?멸린二쇱젣 ?ㅼ썙??異붿텧"""
        keywords = []
        try:
            if progress_callback:
                progress_callback(f"'{current_keyword}' ?멸린二쇱젣 ?ㅼ썙??異붿텧 以?..")
            
            # 媛꾨떒??CSS ?좏깮?먮줈 ?붿냼 李얘린
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
                progress_callback(f"?멸린二쇱젣 ?ㅼ썙??{len(keywords)}媛?異붿텧")
            
            return list(set(keywords))
        except Exception as e:
            if progress_callback:
                progress_callback(f"?멸린二쇱젣 ?ㅼ썙??異붿텧 ?ㅻ쪟: {str(e)}")
            return []

    def recursive_keyword_extraction(self, initial_keyword, progress_callback=None, extract_autocomplete=True):
        """?ш????ㅼ썙??異붿텧 ?꾨줈?몄뒪 - ?꾩쟾 ?ш? 踰꾩쟾"""
        if not self.driver:
            if progress_callback:
                progress_callback("釉뚮씪?곗?媛 珥덇린?붾릺吏 ?딆븯?듬땲??")
            return False
        
        self.base_keyword = initial_keyword
        self.all_related_keywords = []
        self.processed_autocomplete_keywords = set()  # 泥섎━???먮룞?꾩꽦寃?됱뼱 異붿쟻
        
        if progress_callback:
            progress_callback(f"?? '{initial_keyword}' ?꾩쟾 ?ш????ㅼ썙??異붿텧???쒖옉?⑸땲??")

        # 1?④퀎: 湲곕낯 ?ㅼ썙?쒕줈 寃??諛?紐⑤뱺 ?ㅼ썙??異붿텧
        success = self._extract_all_keyword_types(
            initial_keyword, 
            parent_keyword=initial_keyword, 
            depth=0, 
            progress_callback=progress_callback
        )
        
        if not success:
            return False
            
        # 2?④퀎: ?먮룞?꾩꽦寃?됱뼱 ?꾩쟾 ?ш???異붿텧
        if extract_autocomplete:
            autocomplete_keywords = self.extract_autocomplete_keywords(initial_keyword, progress_callback)
            
            # ?먮룞?꾩꽦寃?됱뼱 寃곌낵 ???
            for keyword in autocomplete_keywords:
                self.all_related_keywords.append({
                    'depth': 0,
                    'parent_keyword': initial_keyword,
                    'current_keyword': initial_keyword,
                    'related_keyword': keyword,
                    'keyword_type': '?먮룞?꾩꽦',
                    'source_type': '?먮룞?꾩꽦寃?됱뼱',
                    'extracted_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                })
                
            # ?먮낯 ?ㅼ썙???쒖쇅?섍퀬 ?ш? 泥섎━???ㅼ썙?쒕뱾 以鍮?
            keywords_for_recursion = []
            for keyword in autocomplete_keywords:
                # ?먮낯 ?ㅼ썙?쒖? ?뺥솗???쇱튂?섎뒗 寃쎌슦???쒖쇅
                if keyword.lower().strip() != initial_keyword.lower().strip():
                    keywords_for_recursion.append(keyword)
            
            if progress_callback:
                progress_callback(f"?뱥 ?먮낯 ?ㅼ썙??'{initial_keyword}' ?쒖쇅, {len(keywords_for_recursion)}媛??ㅼ썙?쒕? ?쒖꽌?濡??ш? 泥섎━?⑸땲??")
                if keywords_for_recursion:
                    progress_callback(f"?봽 ?ш? 泥섎━ ?쒖꽌: {', '.join(keywords_for_recursion[:5])}{'...' if len(keywords_for_recursion) > 5 else ''}")
            
            # ?꾩쟾 ?ш????먮룞?꾩꽦寃?됱뼱 異붿텧 ?쒖옉 (?먮낯 ?ㅼ썙???쒖쇅)
            if keywords_for_recursion:
                self._recursive_autocomplete_extraction(
                    keywords_for_recursion, 
                    initial_keyword, 
                    depth=1, 
                    progress_callback=progress_callback
                )
            else:
                if progress_callback:
                    progress_callback("?좑툘 ?먮낯 ?ㅼ썙???몄뿉 ?ш? 泥섎━???먮룞?꾩꽦寃?됱뼱媛 ?놁뒿?덈떎.")
            
        if progress_callback:
            progress_callback(f"'{initial_keyword}' 키워드 추출 완료: 총 {len(self.all_related_keywords)}개")

        return True

    def _extract_all_keyword_types(self, current_keyword, parent_keyword, depth, progress_callback=None):
        """?꾩옱 ?ㅼ썙?쒖뿉 ???紐⑤뱺 ?좏삎???ㅼ썙??異붿텧 (?곌?寃?됱뼱, ?④퍡留롮씠李얜뒗, ?멸린二쇱젣)"""
        try:
            if not self.is_running:
                return False
            
            # ?쇱떆?뺤? 諛??명꽣???곌껐 ?곹깭 ?뺤씤
            if not self.check_pause_status(progress_callback):
                return False
                
            # ?ㅼ썙??寃??
            if not self.search_keyword_mobile(current_keyword, progress_callback):
                return False
            
            # ?쇱떆?뺤? ?곹깭 ?ы솗??
            if not self.check_pause_status(progress_callback):
                return False
            
            # 紐⑤뱺 ?좏삎???ㅼ썙??異붿텧
            related_keywords = self.extract_related_keywords_new(current_keyword, progress_callback)
            
            # ?쇱떆?뺤? ?곹깭 ?뺤씤
            if not self.check_pause_status(progress_callback):
                return False
                
            together_keywords = self.extract_together_keywords(current_keyword, progress_callback)
            
            # ?쇱떆?뺤? ?곹깭 ?뺤씤
            if not self.check_pause_status(progress_callback):
                return False
            
            popular_keywords = self.extract_popular_topics(current_keyword, progress_callback)

            # 寃곌낵 ???
            all_extracted = []
            
            for keyword in related_keywords:
                entry = {
                    'depth': depth,
                    'parent_keyword': parent_keyword,
                    'current_keyword': current_keyword,
                    'related_keyword': keyword,
                    'keyword_type': '?곌?寃?됱뼱',
                    'source_type': '?곌?寃?됱뼱',
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
                    'keyword_type': '?④퍡留롮씠李얜뒗',
                    'source_type': '?④퍡留롮씠李얜뒗',
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
                    'keyword_type': '?멸린二쇱젣',
                    'source_type': '?멸린二쇱젣',
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
                progress_callback(f"??'{current_keyword}' ?ㅼ썙??異붿텧 以??ㅻ쪟: {str(e)}")
            return False

    def _recursive_autocomplete_extraction(self, keywords_to_process, original_keyword, depth, progress_callback=None, max_depth=5):
        """?먮룞?꾩꽦寃?됱뼱 ?꾩쟾 ?ш???異붿텧"""
        
        if depth > max_depth:
            if progress_callback:
                progress_callback(f"?좑툘 理쒕? depth({max_depth}) ?꾨떖濡??ш? 以묐떒")
            return
                
        if not self.is_running:
            return
            
        for i, current_keyword in enumerate(keywords_to_process):
            
            if not self.is_running:
                break
                    
            # ?대? 泥섎━???ㅼ썙?쒕뒗 ?ㅽ궢
            if current_keyword.lower() in self.processed_autocomplete_keywords:
                if progress_callback:
                    progress_callback(f"??툘 '{current_keyword}' ?대? 泥섎━??- ?ㅽ궢")
                continue
                    
            # 泥섎━???ㅼ썙?쒕줈 異붽?
            self.processed_autocomplete_keywords.add(current_keyword.lower())
            
            if progress_callback:
                progress_callback(f"\n?뵇 [{depth}?④퀎] [{i+1}/{len(keywords_to_process)}] '{current_keyword}' ?ш? 泥섎━ 以?..")
            
            # 1. ?꾩옱 ?ㅼ썙?쒕줈 紐⑤뱺 ?좏삎 ?ㅼ썙??異붿텧 (?곌?寃?됱뼱, ?④퍡留롮씠李얜뒗, ?멸린二쇱젣)
            self._extract_all_keyword_types(
                current_keyword, 
                parent_keyword=current_keyword, 
                depth=depth, 
                progress_callback=progress_callback
            )
            
            # 2. ?꾩옱 ?ㅼ썙?쒖쓽 ?먮룞?꾩꽦寃?됱뼱 異붿텧
            new_autocomplete_keywords = self.extract_autocomplete_keywords(current_keyword, progress_callback)
            
            # ?먮룞?꾩꽦寃?됱뼱 寃곌낵 ???
            for keyword in new_autocomplete_keywords:
                self.all_related_keywords.append({
                    'depth': depth,
                    'parent_keyword': current_keyword,
                    'current_keyword': current_keyword,
                    'related_keyword': keyword,
                    'keyword_type': '?먮룞?꾩꽦',
                    'source_type': '?먮룞?꾩꽦寃?됱뼱',
                    'extracted_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                })
            
            # 3. ?덈줈???먮룞?꾩꽦寃?됱뼱媛 ?덉쑝硫??ш? ?몄텧
            if new_autocomplete_keywords:
                # 以묐났 ?쒓굅 諛??꾪꽣留?
                filtered_keywords = []
                for keyword in new_autocomplete_keywords:
                    # ?꾩옱 泥섎━ 以묒씤 ?ㅼ썙?쒖? ?숈씪??寃쎌슦 ?쒖쇅
                    if keyword.lower().strip() == current_keyword.lower().strip():
                        continue
                    
                    # ?대? 泥섎━?섏? ?딆? ?ㅼ썙?쒕쭔 異붽?
                    if keyword.lower() not in self.processed_autocomplete_keywords:
                        # ?먮낯 ?ㅼ썙?쒖? 愿?⑥꽦???덈뒗 ?ㅼ썙?쒕쭔 異붽? (?좏깮?ы빆)
                        if self.base_keyword.lower() in keyword.lower() or len(filtered_keywords) < 20:  # ?덈Т 留롮? ?ㅼ썙??諛⑹?
                            filtered_keywords.append(keyword)
                
                if filtered_keywords:
                    if progress_callback:
                        progress_callback(f"?봽 '{current_keyword}'?먯꽌 {len(filtered_keywords)}媛????먮룞?꾩꽦寃?됱뼱 諛쒓껄 ??{depth+1}?④퀎 ?ш? 吏꾪뻾")
                        if len(filtered_keywords) <= 10:  # 10媛??댄븯硫?紐⑤몢 ?쒖떆
                            progress_callback(f"?뱷 ?ㅼ쓬 ?쒖꽌濡?泥섎━: {', '.join(filtered_keywords)}")
                        else:  # 10媛?珥덇낵硫?泥섏쓬 10媛쒕쭔 ?쒖떆
                            progress_callback(f"?뱷 ?ㅼ쓬 ?쒖꽌濡?泥섎━: {', '.join(filtered_keywords[:10])} ... (珥?{len(filtered_keywords)}媛?")
                    
                    # ?ш? ?몄텧
                    self._recursive_autocomplete_extraction(
                        filtered_keywords, 
                        original_keyword, 
                        depth + 1, 
                        progress_callback, 
                        max_depth
                    )
                else:
                    if progress_callback:
                        progress_callback(f"??'{current_keyword}' - ?덈줈???먮룞?꾩꽦寃?됱뼱 ?놁쓬")
            else:
                if progress_callback:
                    progress_callback(f"??'{current_keyword}' - ?먮룞?꾩꽦寃?됱뼱 ?놁쓬")
            
        if progress_callback:
            progress_callback(f"?뢾 {depth}?④퀎 ?ш? 泥섎━ ?꾨즺!")

    def save_recursive_results_to_excel(self, save_path=None, progress_callback=None):
        """Save extraction results to file."""
        try:
            if not hasattr(self, 'all_related_keywords') or not self.all_related_keywords:
                if progress_callback:
                    progress_callback("????ν븷 ?ㅼ썙?쒓? ?놁뒿?덈떎.")
                return False
            
            if not save_path:
                if not self.save_dir:
                    self.save_dir = "keyword_results"
                    os.makedirs(self.save_dir, exist_ok=True)
                
                current_time = datetime.now().strftime('%Y%m%d_%H%M%S')
                base_keyword = getattr(self, 'base_keyword', 'keyword_extraction')
                save_path = os.path.join(self.save_dir, f"{base_keyword}_{current_time}.xlsx")
            
            # ?곗씠?고봽?덉엫 ?앹꽦
            df = pd.DataFrame({
                '추출된_키워드': [item['related_keyword'] for item in self.all_related_keywords]
            })

            # 以묐났 ?쒓굅
            df = df.drop_duplicates(subset=['추출된_키워드'], keep='first').reset_index(drop=True)

            # ?묒? ???
            try:
                df.to_excel(save_path, index=False, engine='openpyxl')
                
                if os.path.exists(save_path) and os.path.getsize(save_path) > 0:
                    if progress_callback:
                        progress_callback(f"???묒? ?뚯씪 ????꾨즺: {save_path}")
                        progress_callback(f"저장된 키워드 수: {len(df)}")
                    return True
                else:
                    raise Exception("?묒? ?뚯씪 ?앹꽦 ?ㅽ뙣")
                
            except Exception as excel_error:
                # CSV濡?諛깆뾽 ???
                csv_path = save_path.rsplit('.', 1)[0] + '.csv'
                df.to_csv(csv_path, index=False, encoding='utf-8-sig')
                    
                if progress_callback:
                    progress_callback(f"?좑툘 ?묒? ????ㅽ뙣, CSV濡???? {csv_path}")
                return True
            
        except Exception as e:
            if progress_callback:
                progress_callback(f"???뚯씪 ????ㅻ쪟: {str(e)}")
            return False

    def close(self):
        """釉뚮씪?곗? 醫낅즺"""
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
    finished = pyqtSignal(str)              # ?꾨즺 ????λ맂 ?뚯씪 寃쎈줈 ?쒓렇??
    error = pyqtSignal(str)                 # ?먮윭 ?쒓렇??
    log = pyqtSignal(str, str)              # 濡쒓렇 ?쒓렇??(?ㅼ썙?? 硫붿떆吏)
    
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
            
            # 釉뚮씪?곗? ?앹꽦
            self.driver = create_chrome_driver()
            if not self.driver:
                self.error.emit(f"'{self.keyword}' 브라우저 생성 실패")
                return

            # 寃?됯린 珥덇린??
            self.searcher = NaverMobileSearchScraper(driver=self.driver)
            self.searcher.save_dir = self.save_dir
            self.searcher.is_running = self.is_running
            self.searcher.search_thread = self
            
            # ?ㅼ썙??異붿텧 ?ㅽ뻾
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
            # 釉뚮씪?곗? 醫낅즺 諛??뺣━
            if self.driver:
                try:
                    self.driver.quit()
                except:
                    pass
                self.driver = None

    def _log_wrapper(self, msg):
        """濡쒓렇 ?섑띁: ?ㅼ썙???앸퀎??異붽?"""
        self.log.emit(self.keyword, msg)

    def stop(self):
        """?묒뾽 以묐떒"""
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
    """Chrome WebDriver ?앹꽦 諛??ㅼ젙 (?낅┰?곸쑝濡??ㅽ뻾 媛??"""
    try:
        if log_callback:
            log_callback("Chrome ?쒕씪?대쾭 ?ㅼ젙???쒖옉?⑸땲??..")
        
        try:
            driver_path = ChromeDriverManager().install()
            if log_callback:
                log_callback(f"Chrome ?쒕씪?대쾭 寃쎈줈: {driver_path}")
        except Exception as e:
            if log_callback:
                log_callback(f"?좑툘 webdriver-manager ?ㅻ쪟 (濡쒖뺄 ?쒕씪?대쾭 ?ъ슜 ?쒕룄): {str(e)}")
            driver_path = None
        
        options = webdriver.ChromeOptions()
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--disable-gpu")
        options.add_argument("--log-level=3")
        options.add_argument("--silent")

        if driver_path:
            service = Service(driver_path)
        else:
            service = Service()
        
        driver = webdriver.Chrome(service=service, options=options)
        driver.set_page_load_timeout(20)
        driver.implicitly_wait(5)
        
        driver.get("about:blank")
        
        if log_callback:
            log_callback("??Chrome ?쒕씪?대쾭媛 ?깃났?곸쑝濡?珥덇린?붾릺?덉뒿?덈떎.")
            
        return driver

    except Exception as e:
        if log_callback:
            log_callback(f"??Chrome ?쒕씪?대쾭 ?앹꽦 ?ㅽ뙣: {str(e)}")
        raise e


class KeywordExtractorMainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        
        # ?꾩씠肄??ㅼ젙
        icon_path = get_icon_path()
        if icon_path:
            self.setWindowIcon(QIcon(icon_path))
            safe_print(f"아이콘 설정 완료: {icon_path}")
        else:
            safe_print("아이콘 파일을 찾을 수 없습니다.")
        
        self.setWindowTitle("네이버 연관키워드 추출기")
        self.resize(1200, 800) # 湲곕낯 ?ш린 ?ㅼ젙
        self.showMaximized()   # ?꾨줈洹몃옩 ?ㅽ뻾 ???꾩껜 ?붾㈃?쇰줈 ?쒖옉
        
        # ?ㅼ젙 諛??쒕씪?대쾭 珥덇린??
        self.settings = Settings()
        self.driver = None
        # self.search_thread = None  # ?⑥씪 ?ㅻ젅?????由ъ뒪???ъ슜
        self.active_threads = []     # ?? ??? ??
        self.completed_threads = 0   # ??? ??? ?
        self.total_threads = 0       # ?? ??? ?
        self.stop_requested = False
        
        # ?щ옒??蹂댄샇 ?ㅼ젙
        self.setup_crash_protection()
        
        # UI 珥덇린??
        self.init_ui()
        self.setup_chrome_driver()
        
        # ?ㅽ????곸슜
        self.setStyleSheet(STYLESHEET)

    def check_license_info(self):
        """?쇱씠?좎뒪 ?뺣낫 ?뺤씤"""
        # 癒몄떊 ID ?뺤씤
        machine_id = get_machine_id()
        
        # 援ш? ?쒗듃?먯꽌 留뚮즺???뺤씤
        expiration_date = check_license_from_sheet(machine_id)
        
        if expiration_date:
            try:
                # ?좎쭨 鍮꾧탳 (YYYY-MM-DD ?뺤떇 媛??
                exp_date = datetime.strptime(str(expiration_date).strip(), '%Y-%m-%d')
                today = datetime.now()
                
                if exp_date < today:
                    # 留뚮즺??
                    self.show_license_dialog(machine_id, expired=True)
                else:
                    # ?좏슚??
                    self.usage_label.setText(f"사용 기간: {expiration_date}까지")
                    self.usage_label.setStyleSheet(f"font-size: 13px; font-weight: bold; color: {NAVER_GREEN};")
            except:
                # ?좎쭨 ?뺤떇???꾨땲硫??쇰떒 ?듦낵 (?⑥닚 ?띿뒪???? ?먮뒗 留뚮즺???놁쓬?쇰줈 媛꾩＜
                # ?ш린?쒕뒗 ?띿뒪??洹몃?濡??쒖떆 (?? "臾댁젣??)
                self.usage_label.setText(f"사용 기간: {expiration_date}")
                self.usage_label.setStyleSheet(f"font-size: 13px; font-weight: bold; color: {NAVER_GREEN};")
        else:
            # ?깅줉?섏? ?딆쓬
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
        """?щ옒??蹂댄샇 ?ㅼ젙"""
        global _current_window
        _current_window = self
        sys.excepthook = handle_exception
        
        # ?쇱씠?좎뒪 泥댄겕 ?쒖옉
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
        """硫붿씤 ?덈룄?곗슜 Chrome WebDriver ?ㅼ젙 (?꾩슂 ??"""
        # 蹂묐젹 紐⑤뱶?먯꽌??媛쒕퀎 ?ㅻ젅?쒓? ?쒕씪?대쾭瑜??앹꽦?섎?濡??ш린?쒕뒗 ?앹꽦?섏? ?딄굅??
        # ?먮뒗 ?⑥씪 ?ㅽ뻾 ?뚯뒪?몃? ?꾪빐 ?④꺼?????덉쓬. 
        # ?쇰떒 湲곗〈 濡쒖쭅 ?좎?瑜??꾪빐 ?⑥닔 ?몄텧濡?蹂寃쏀븯吏留? ?ㅼ젣濡쒕뒗 start_search?먯꽌 ?앹꽦??
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

        # ?ㅻ뜑 而⑦뀒?대꼫 (?쒕ぉ + ?ъ슜 湲곌컙)
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
        
        # ?쒕ぉ (以묒븰 ?뺣젹???꾪빐 ?묒そ??stretch 異붽?)
        title_label = QLabel("네이버 연관키워드 추출기")
        title_label.setStyleSheet(f"font-size: 16px; font-weight: 800; color: {NAVER_GREEN};")
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        # ?ъ슜 湲곌컙
        self.usage_label = QLabel("사용 기간: 확인 중...")
        self.usage_label.setStyleSheet(f"font-size: 13px; font-weight: bold; color: #555555;")
        
        # ?쇱そ ?щ갚 (titie??以묒븰?쇰줈 諛湲??꾪빐)
        header_layout.addStretch()
        header_layout.addWidget(title_label)
        header_layout.addStretch()
        header_layout.addWidget(self.usage_label)
        
        main_layout.addWidget(header_widget)
        
        # ?곷떒 ?뱀뀡 (寃??+ 吏꾪뻾 ?곹솴) - 媛濡?諛곗튂
        top_section_layout = QHBoxLayout()
        
        # 寃???뱀뀡 (?쇱そ)
        self.setup_search_section(top_section_layout)
        
        # 吏꾪뻾 ?곹솴 ?뱀뀡 (?ㅻⅨ履?
        self.setup_progress_section(top_section_layout)
        
        main_layout.addLayout(top_section_layout)
        
        # ????꾩튂 ?뱀뀡 (?섎떒)
        self.setup_save_section(main_layout)
        self.main_tabs.addTab(extractor_tab, "연관 키워드 추출")
            
        # ?곹깭諛?
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("준비 완료")

    def setup_search_section(self, main_layout):
        """검색 섹션 설정"""
        search_group = QGroupBox("키워드 검색")
        search_layout = QVBoxLayout(search_group)
        
        # 寃???낅젰李?
        self.search_input = MultiKeywordTextEdit()
        self.search_input.setPlaceholderText(
            "사용 방법\n"
            "1. 키워드를 한 줄에 하나씩 입력하세요.\n"
            "2. Enter로 바로 추출 시작, Shift+Enter로 줄바꿈합니다.\n"
            "3. 여러 키워드를 동시에 병렬 처리합니다."
        )
        # ?믪씠 ?쒗븳 ?쒓굅 諛??뺤옣 ?뺤콉 ?ㅼ젙
        self.search_input.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self.search_input.search_requested.connect(self.start_search)
        search_layout.addWidget(self.search_input)
        
        # ?щ갚 理쒖냼??
        search_layout.setContentsMargins(10, 10, 10, 10)
        search_layout.setSpacing(5)
        
        # 踰꾪듉??
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
        
        # 湲곕낯 寃쎈줈 ?ㅼ젙
        saved_dir = self.settings.get_save_dir()
        if saved_dir and os.path.exists(saved_dir):
            self.save_path_input.setText(saved_dir)
        else:
            # ?ъ슜?먮퀎 諛뷀깢?붾㈃ 寃쎈줈 ?숈쟻 ?앹꽦
            desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
            default_dir = os.path.join(desktop_path, "keyword_results")
            
            try:
                os.makedirs(default_dir, exist_ok=True)
            except Exception:
                # 諛뷀깢?붾㈃ ?묎렐 ?ㅽ뙣 ????臾몄꽌濡??泥?
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
        progress_layout.setContentsMargins(10, 10, 10, 10) # ?щ갚 ?쇱튂
        
        # ???꾩젽?쇰줈 蹂寃?
        self.progress_tabs = QTabWidget()
        self.progress_tabs.setStyleSheet("""
            QTabWidget::pane { border: 1px solid #CCCCCC; border-radius: 8px; }
            QTabBar::tab { background: #f0f0f0; padding: 8px 12px; border-top-left-radius: 6px; border-top-right-radius: 6px; margin-right: 2px; }
            QTabBar::tab:selected { background: #E6F0FD; font-weight: bold; color: #1E6ECA; border-bottom: 2px solid #1E6ECA; }
        """)
        
        # '?꾩껜' ??(?쒖뒪??濡쒓렇??
        self.total_log_text = SmartProgressTextEdit(min_height=100, max_height=800)
        self.total_log_text.setReadOnly(True)
        self.total_log_text.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self.total_log_text.setPlaceholderText("여기에 전체 진행 로그가 표시됩니다.")
        self.progress_tabs.addTab(self.total_log_text, "전체 로그")
        
        # ??愿由ъ슜 ?뺤뀛?덈━
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

        
        # UI ?곹깭 蹂寃?
        self.start_button.setEnabled(False)
        self.pause_button.setEnabled(True)
        self.pause_button.setText("일시정지")
        self.stop_button.setEnabled(True)
        self.search_input.setEnabled(False)
        # self.progress_text.clear() ??젣??(??珥덇린?붾줈 ?泥?
        self.status_bar.showMessage("키워드 추출 중...")
        
        self.update_progress(f"키워드 추출 작업 시작 (총 {len(keywords)}개)")
        self.update_progress(f"입력 키워드: {', '.join(keywords)}")
        self.update_progress(f"저장 폴더: {save_dir}")
        self.update_progress("병렬 처리 모드로 실행합니다.")
        
        # ??珥덇린??(湲곗〈 媛쒕퀎 ???쒓굅, '?꾩껜 濡쒓렇'???좎?)
        while self.progress_tabs.count() > 1:
            self.progress_tabs.removeTab(1)
            
        self.log_widgets = {"전체": self.total_log_text}
        self.total_log_text.clear()
        
        # 寃???ㅻ젅???쒖옉 (蹂묐젹 ?ㅽ뻾)
        self.active_threads = []
        self.completed_threads = 0
        self.total_threads = len(keywords)
        self.stop_requested = False
        
        # 媛??ㅼ썙?쒕퀎 ???앹꽦
        for keyword in keywords:
            log_widget = SmartProgressTextEdit(min_height=100, max_height=800)
            log_widget.setReadOnly(True)
            self.progress_tabs.addTab(log_widget, keyword)
            self.log_widgets[keyword] = log_widget
            
            # ???꾪솚 (泥?踰덉㎏ ?ㅼ썙?쒕줈)
            if keywords.index(keyword) == 0:
                self.progress_tabs.setCurrentIndex(1)
        
        for i, keyword in enumerate(keywords):
            thread = ParallelKeywordThread(keyword, save_dir, True)
            thread.finished.connect(self.on_thread_finished)
            thread.error.connect(self.on_thread_error)
            thread.log.connect(self.update_progress) # ?쒓렇??留ㅽ븨 ?먮룞 泥섎━??
            
            self.active_threads.append(thread)
            thread.start()
            self.update_progress(keyword, f"'{keyword}' 작업 시작...")

    def on_thread_finished(self, save_path):
        """?ㅻ젅???묒뾽 ?꾨즺 泥섎━"""
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
        """紐⑤뱺 ?ㅻ젅?쒓? ?꾨즺?섏뿀?붿? ?뺤씤"""
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
        """吏꾪뻾 ?곹솴 ?낅뜲?댄듃 (?ㅼ썙?쒕퀎 ??遺꾨━ 吏??"""
        current_time = datetime.now().strftime('%H:%M:%S')
        
        # ?몄옄 泥섎━ (湲곗〈 ?명솚??+ ?덈줈???쒓렇??
        if message is None:
            # ?⑥씪 ?몄옄 ?몄텧??寃쎌슦 (湲곕낯 ?쒖뒪??硫붿떆吏 ??
            target_keyword = "전체"
            msg_content = keyword_or_msg
        else:
            # (?ㅼ썙?? 硫붿떆吏) ?뺥깭 ?몄텧
            target_keyword = keyword_or_msg
            msg_content = message

        target_keyword = sanitize_display_text(target_keyword)
        msg_content = sanitize_display_text(msg_content)
        formatted_message = f"[{current_time}] {msg_content}"
        
        # 1. ?대떦 ?ㅼ썙?쒖쓽 媛쒕퀎 ??뿉 濡쒓렇 異붽?
        if target_keyword in self.log_widgets:
            widget = self.log_widgets[target_keyword]
            if hasattr(widget, 'append_with_smart_scroll'):
                widget.append_with_smart_scroll(formatted_message)
            else:
                widget.append(formatted_message)
        
        # 2. '?꾩껜 濡쒓렇' ??뿉??紐⑤뱺 濡쒓렇 異붽? (?좏깮?ы빆, 紐⑤땲?곕쭅 ?몄쓽 ?꾪빐)
        if target_keyword != "전체":
            formatted_total_msg = f"[{current_time}] [{target_keyword}] {msg_content}"
            if hasattr(self.total_log_text, 'append_with_smart_scroll'):
                self.total_log_text.append_with_smart_scroll(formatted_total_msg)
            else:
                self.total_log_text.append(formatted_total_msg)

    def pause_resume_search(self):
        """?쇱떆?뺤?/?ш컻 ?좉?"""
        # 蹂묐젹 紐⑤뱶?먯꽌??媛쒕퀎 ?ㅻ젅???쇱떆?뺤? 吏?먯씠 蹂듭옟?섎?濡?
        # ?꾩옱????湲곕뒫??鍮꾪솢?깊솕?섍굅??濡쒓렇留??④? (?먮뒗 異뷀썑 援ы쁽)
        # ?ш린?쒕뒗 ?⑥닚??踰꾪듉 ?곹깭留??좉??섎뒗 寃껋쑝濡??꾩떆 泥섎━
        pass

    def on_search_paused(self):
        """?쇱떆?뺤? ??UI ?낅뜲?댄듃"""
        self.pause_button.setText("재개")
        self.status_bar.showMessage("키워드 추출이 일시정지되었습니다.")
    
    def on_search_resumed(self):
        """?ш컻 ??UI ?낅뜲?댄듃"""
        self.pause_button.setText("일시정지")
        self.status_bar.showMessage("키워드 추출 중...")

    def reset_ui_state(self):
        """UI ?곹깭 由ъ뀑"""
        self.start_button.setEnabled(True)
        self.pause_button.setEnabled(False)
        self.pause_button.setText("일시정지")
        self.stop_button.setEnabled(False)
        self.search_input.setEnabled(True)

    def closeEvent(self, event):
        """Handle window close event and cleanup."""
        global _crash_save_enabled
        
        # ?쒖꽦 ?ㅻ젅?쒓? ?덈뒗吏 ?뺤씤
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
                    
                    # 紐⑤뱺 ?ㅻ젅??以묐떒
                    for thread in running_threads:
                        thread.stop()
                        thread.wait(1000)
                    
                    self.active_threads = []
                else:
                    event.ignore()
                    return

        _crash_save_enabled = False
        
        if self.driver:
            pass  # 硫붿씤 ?쒕씪?대쾭?????댁긽 ?ъ슜?섏? ?딆쓬
        
        event.accept()


def main():
    if sys.platform == 'win32':
        import ctypes
        ctypes.windll.kernel32.SetConsoleOutputCP(65001)
        ctypes.windll.kernel32.SetConsoleCP(65001)

    if not verify_machine_id_guard():
        sys.exit(1)
    
    app = QApplication(sys.argv)
    
    # ?좏뵆由ъ??댁뀡 ?꾩씠肄??ㅼ젙
    icon_path = get_icon_path()
    if icon_path:
        app.setWindowIcon(QIcon(icon_path))
        safe_print(f"아이콘 설정 완료: {icon_path}")
    else:
        safe_print("아이콘 파일을 찾을 수 없습니다.")
    
    app.setApplicationName("네이버 연관키워드 추출기")
    
    # 1. 癒몄떊 ID ?뺤씤
    machine_id = get_machine_id()
    safe_print(f"Machine ID: {machine_id}")
    
    # 2. ?쇱씠?좎뒪 泥댄겕 (?숆린??- ?꾨줈洹몃옩 ?쒖옉 ???꾩닔)
    # 2. ?쇱씠?좎뒪 泥댄겕 (?숆린??- ?꾨줈洹몃옩 ?쒖옉 ???꾩닔)
    expiry_date_str = check_license_from_sheet(machine_id)
    
    if expiry_date_str:
        try:
            # ?좎쭨 鍮꾧탳 濡쒖쭅 (YYYY-MM-DD ?뺤떇 媛??
            expiry_date = datetime.strptime(expiry_date_str, '%Y-%m-%d')
            current_date = datetime.now()
            
            # 留뚮즺?쇱씠 吏??寃쎌슦 (留뚮즺???ㅼ쓬?좊???李⑤떒)
            if current_date > expiry_date + pd.Timedelta(days=1):
                safe_print(f"라이선스 만료: {expiry_date_str}")
                app_dummy = QApplication.instance() or QApplication(sys.argv)
                
                # 留뚮즺 ?ㅼ씠?쇰줈洹??쒖떆
                dialog = ExpiredDialog(expiry_date_str)
                dialog.exec()
                sys.exit(0)
            
            # ?쇱씠?좎뒪 ?좏슚??-> 硫붿씤 ?꾨줈洹몃옩 ?ㅽ뻾
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
            # ?좎쭨 ?뺤떇???섎せ??寃쎌슦?먮룄 ?쇰떒 ?ㅽ뻾? ?쒖폒二쇰릺 寃쎄퀬 (?좎? ?몄쓽)
            safe_print(f"라이선스 날짜 형식 확인 필요: {expiry_date_str}")
            window = KeywordExtractorMainWindow()
            window.usage_label.setText(f"사용 기간: {expiry_date_str}")
            window.show()
            sys.exit(app.exec())
            
    else:
        # ?쇱씠?좎뒪 ?놁쓬 -> ?ㅼ씠?쇰줈洹??쒖떆 ??醫낅즺
        safe_print("미등록 기기 - 실행 차단")
        dialog = UnregisteredDialog(machine_id)
        dialog.exec()
        sys.exit(0)

if __name__ == "__main__":
    main() 







