# 추천 파일명: hwp_batch_password_setter.py

# 필요한 패키지 설치:

# pip install pywin32 pikepdf openpyxl xlrd msoffcrypto-tool tkinterdnd2
# tkinterdnd2는 드래그앤드롭 기능용 선택 패키지입니다. 없어도 프로그램은 실행됩니다.

import csv
import json
import os
import queue
import re
import shutil
import stat
import struct
import subprocess
import tempfile
import threading
import time
import traceback
import unicodedata
import zipfile
from dataclasses import dataclass, field
from datetime import datetime, timedelta
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

try:
    import pythoncom
    import pywintypes
    import win32com.client
    import win32con
    import win32gui
    import win32process
except ImportError:
    pythoncom = None
    pywintypes = None
    win32com = None
    win32con = None
    win32gui = None
    win32process = None

try:
    import pikepdf
except ImportError:
    pikepdf = None

try:
    import openpyxl
except ImportError:
    openpyxl = None

try:
    import xlrd
except ImportError:
    xlrd = None

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
except ImportError:
    DND_FILES = None
    TkinterDnD = None

APP_TITLE = "문서 일괄 암호 설정기 (HWP·PDF·Excel)"

# HWP/HWPX는 COM 자동화, PDF/Excel은 순수 Python 라이브러리로 처리

HWP_EXTENSIONS  = {".hwp", ".hwpx"}
PDF_EXTENSIONS  = {".pdf"}
EXCEL_EXTENSIONS = {".xlsx", ".xlsm", ".xls"}
WORD_EXTENSIONS = {".docx", ".docm", ".doc"}
POWERPOINT_EXTENSIONS = {".pptx", ".pptm", ".ppt"}
IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tif", ".tiff", ".webp"}
SUPPORTED_EXTENSIONS = HWP_EXTENSIONS | PDF_EXTENSIONS | EXCEL_EXTENSIONS | WORD_EXTENSIONS | POWERPOINT_EXTENSIONS
SEARCHABLE_EXTENSIONS = SUPPORTED_EXTENSIONS | IMAGE_EXTENSIONS
OLD_OFFICE_EXTENSIONS = {".xls", ".doc", ".ppt"}
OLD_OFFICE_BACKUP_DIR = "_old_office_backup"
RETENTION_HIGH_RISK_KEYWORDS = [
    "주민등록", "주민번호", "등본", "초본", "복지카드", "장애인등록증",
    "신분증", "운전면허", "여권", "가족관계", "진단서", "소견서",
    "장애정도", "장기요양인정서", "수급자증명", "기초생활수급",
]
RETENTION_GENERAL_KEYWORDS = [
    "이용자명단", "명단", "개인정보", "상담", "사례관리", "평가", "계약서", "동의서",
    "신청서", "입소", "퇴소", "보호자", "수급자", "이용자", "대상자", "회원",
    "연락처", "주소", "생년월일", "급여제공", "제공기록", "사례회의", "면담",
    "방문기록", "모니터링", "욕구사정", "서비스계획", "이용계약", "개별파일",
    "종결", "사례", "건강", "건강관리", "개별상담", "초기상담", "상담일지",
    "개별지원", "지원계획", "개별계획", "관리카드", "카드", "기록지", "일지",
    "접수", "의뢰", "연계", "사정", "재사정", "위기", "위기관리", "사후관리",
    "이용신청", "이용중지", "이용종결", "서비스종결", "퇴소상담", "입소상담",
    "보호자상담", "가정방문", "방문상담", "생활기록", "활동기록", "관찰기록",
    "투약", "복약", "약물", "질병", "병력", "의료", "진료", "검진", "간호",
    "재활", "치료", "심리", "정서", "자립", "직업", "훈련", "교육일지",
]
RETENTION_OLD_DOC_EXTENSIONS = {".hwp", ".doc", ".xls", ".ppt"}
RETENTION_MODERN_DOC_EXTENSIONS = {".hwpx", ".docx", ".xlsx", ".pptx"}

RESULT_PENDING = "대기"
RESULT_SUCCESS = "성공"
RESULT_SKIPPED = "건너뜀"
RESULT_FAILED  = "실패"
RESULT_CANCELLED = "취소"
TREE_CHECKED   = "☑"
TREE_UNCHECKED = "☐"

RESULT_CODE_SUCCESS = "SUCCESS"
RESULT_CODE_FAIL = "FAIL"
RESULT_CODE_SKIPPED = "SKIPPED"
RESULT_CODE_ALREADY_ENCRYPTED = "ALREADY_ENCRYPTED"
RESULT_CODE_NO_MSOFFCRYPTO = "NO_MSOFFCRYPTO"
RESULT_CODE_CANCELLED = "CANCELLED"

DEFAULT_SELECTION_BATCH_SIZE = 100
HWP_COM_RESTART_EVERY = 20
HWP_HEAVY_FILE_RESTART_SECONDS = 45
HWP_COM_COOLDOWN_SECONDS = 0.4

ABOUT_MESSAGE = (
    "비밀번호 설정 도구\n\n"
    "Version 1.4\n\n"
    "[목적]\n"
    "문서 파일(Excel, PDF, HWP 등)에 일괄 암호 설정을 수행합니다.\n\n"
    "[주의사항]\n"
    "- 작업 전 백업을 권장합니다.\n"
    "- 파일 형식에 따라 암호 적용 방식이 다를 수 있습니다.\n"
    "- 작업 중 강제 종료 시 파일 손상이 발생할 수 있습니다.\n\n"
    "Created by JK"
)

HELP_MANUAL_TEXT = """개인정보파일 암호화 도구 사용설명서

1. 프로그램 목적
이 프로그램은 개인정보가 포함된 문서 파일을 일괄로 암호화하기 위한 도구입니다.

지원 대상은 HWP/HWPX, PDF, Excel, Word, PowerPoint 파일입니다. 여러 폴더에 흩어진 파일을 한 번에 검색하거나 드래그앤드롭으로 추가한 뒤, 공통 암호를 적용할 수 있습니다.
JPG, PNG 등 이미지 파일은 암호화 대상은 아니지만 삭제 후보 점검 대상으로 함께 검색됩니다.

개인정보 보호를 위해 암호화가 끝난 파일은 반드시 실제로 열어 암호가 적용되었는지 확인하는 것을 권장합니다.

2. 기본 사용 순서
1. 대상 폴더 선택을 누르거나 파일을 목록에 드래그앤드롭합니다.
2. 필요한 경우 하위 폴더 포함을 체크합니다.
3. 파일 검색을 눌러 암호화 대상 파일을 불러옵니다.
4. 공통 암호와 공통 암호 확인을 입력합니다.
5. 필요하면 실행 전 점검을 눌러 파일 접근 가능 여부를 확인합니다.
6. 암호 설정 실행을 누릅니다.
7. 작업 완료 후 성공/실패 로그와 파일별 결과를 확인합니다.
8. 중요한 파일은 직접 열어 암호 적용 여부를 확인합니다.

3. 주요 옵션 설명
하위 폴더 포함
대상 폴더 안의 하위 폴더까지 검색합니다. 교육자료, 제출자료처럼 폴더가 여러 단계로 나뉜 경우 체크하는 것이 좋습니다.

암호 파일 스킵
이미 암호가 걸린 파일은 다시 처리하지 않고 건너뜁니다. 기존 암호 파일을 덮어 처리하는 실수를 줄이기 위한 옵션입니다.

백업 폴더 생성 후 처리
암호화 전에 원본 파일을 별도 백업 폴더에 복사합니다. 중요한 자료를 처리할 때 권장합니다.
단, 백업 폴더에는 암호화 전 원본 파일이 남을 수 있으므로, 작업 완료 후 보관 필요 여부를 확인해야 합니다.

구형 Office 변환
구형 Office 파일을 최신 형식으로 변환한 뒤 암호화를 시도합니다.

.xls -> .xlsx
.doc -> .docx
.ppt -> .pptx

구형 파일은 직접 열기 암호 설정이 제한되기 때문에, 이 옵션을 켜면 Microsoft Office를 이용해 최신 형식으로 변환한 뒤 암호화합니다.

암호 보기
입력한 공통 암호를 화면에 표시합니다. 주변 사람이 볼 수 있는 환경에서는 사용하지 않는 것이 좋습니다.

삭제 후보 점검
파일명, 수정일, 확장자를 기준으로 오래된 개인정보 문서일 가능성을 점검합니다.
이 기능은 파일을 자동 삭제하지 않으며, 파일 내용도 열지 않습니다.
결과는 참고용 판단 자료이며, 실제 파기 여부는 담당자가 파일 내용과 기관 보존 기준을 확인한 뒤 결정해야 합니다.

4. 구형 Office 변환 사용 시 주의사항
구형 Office 변환 옵션은 편리하지만, 개인정보 보호 관점에서 반드시 확인해야 할 부분이 있습니다.

변환과 암호화가 성공하면 원본 구형 파일은 다음 백업 폴더로 이동됩니다.

대상 폴더\\_old_office_backup\\원래 하위 경로\\원본파일.xls/doc/ppt

예시:

교육자료
├─ 한자\\자료.xls
└─ 와인\\강의.doc

변환 후 원본 백업 위치:

교육자료
└─ _old_office_backup
   ├─ 한자\\자료.xls
   └─ 와인\\강의.doc

중요:
변환된 최신 파일이 정상이고 암호가 적용된 것을 확인한 뒤, _old_office_backup 폴더를 삭제하는 것을 권장합니다.

이 백업 폴더에는 암호화 전 원본 파일이 들어 있을 수 있습니다. 개인정보 파일인 경우 백업 폴더를 그대로 남겨두면 개인정보 유출 위험이 생길 수 있습니다.

5. 파일 형식별 안내
PDF
PDF 파일은 열기 암호를 설정합니다. 작업 후 PDF 뷰어에서 열 때 암호를 요구하는지 확인해 주세요.

Excel
.xlsx, .xlsm 파일은 열기 암호를 설정합니다. .xls 파일은 구형 형식이므로 구형 Office 변환 옵션을 켜야 최신 형식으로 변환 후 암호화할 수 있습니다.

Word
.docx, .docm 파일은 열기 암호를 설정합니다. .doc 파일은 구형 형식이므로 구형 Office 변환 옵션을 켜야 최신 형식으로 변환 후 암호화할 수 있습니다.

PowerPoint
.pptx, .pptm 파일은 열기 암호를 설정합니다. .ppt 파일은 구형 형식이므로 구형 Office 변환 옵션을 켜야 최신 형식으로 변환 후 암호화할 수 있습니다.

이미지
.jpg, .jpeg, .png, .bmp, .gif, .tif, .tiff, .webp 파일은 문서처럼 열기 암호를 설정하지 않습니다.
대신 삭제 후보 점검에서 파일명, 수정일, 확장자를 기준으로 개인정보 이미지 가능성을 확인합니다.

HWP/HWPX
한글 프로그램을 이용해 암호를 설정합니다. 작업 중 한글에서 접근 허용 창이 뜰 수 있습니다. 프로그램 창 뒤에 가려질 수 있으니, 한글 접근 허용 창이 보이면 허용을 눌러 주세요.

HWP/HWPX 파일은 용량이 크거나 이미지, 표, 개체가 많으면 암호화 시간이 조금 더 걸릴 수 있습니다.

HWPX는 자동 검증 신뢰도가 낮을 수 있으므로, 작업 후 한글에서 직접 열어 암호 적용 여부를 확인하는 것을 권장합니다.

6. 한 번에 처리하기 좋은 파일 수
파일 수가 너무 많으면 Office 또는 한글문서 자동화가 느려지거나, 중간에 접근 허용 창을 놓칠 수 있습니다.

권장 기준:
- PDF 위주: 100개 내외 가능
- Excel/Word/PowerPoint 위주: 50개 내외 권장
- HWP/HWPX 포함: 20~30개 내외 권장
- 구형 Office 변환 포함: 10~30개 내외 권장
- 대용량 HWP/HWPX 포함: 10~20개 정도로 나누어 처리 권장

PC 환경에 따라 더 많은 파일도 처리될 수 있지만, 개인정보 파일은 처리 후 확인이 중요하므로 실패 확인과 재처리가 쉬운 단위로 나누어 처리하는 것을 권장합니다.

7. 삭제 후보 점검 안내
삭제 후보 점검은 파일명, 수정일, 확장자를 기준으로 오래된 개인정보 문서일 가능성을 점수로 보여주는 기능입니다.

점검 기준:
- 5년 초과 파일은 점수가 올라갑니다.
- 주민번호, 등본, 복지카드, 신분증, 진단서 등 민감한 단어가 파일명에 있으면 점수가 올라갑니다.
- 명단, 개인정보, 상담, 사례, 사례회의, 종결, 건강관리, 기록지, 계약서, 동의서, 수급자, 보호자 등 개인정보 관련 단어가 파일명에 있으면 점수가 올라갑니다.
- 구형 문서, PDF, 이미지 파일은 점수가 추가됩니다.
- 주민번호.jpg, 신분증.png, 복지카드사진.jpeg처럼 개인정보로 추정되는 이미지 파일은 삭제 권장으로 분류될 수 있습니다.

분류 결과:
- 삭제 권장: 보관 필요 여부 확인 후 파기 검토
- 검토 필요: 담당자 확인 후 보관/파기 결정
- 낮은 위험: 필요 시 보관

주의:
이 결과는 참고용입니다. 실제 삭제 여부는 담당자가 파일 내용과 기관 보존 기준을 확인한 뒤 결정해야 합니다.
본 프로그램은 파일을 자동 삭제하지 않으며, 삭제 버튼도 제공하지 않습니다.

8. 작업 후 확인해야 할 것
작업이 끝나면 아래 항목을 확인해 주세요.

1. 실패 파일이 있는지 로그를 확인합니다.
2. 중요한 파일은 직접 열어 암호 입력창이 뜨는지 확인합니다.
3. 구형 Office 변환을 사용했다면 _old_office_backup 폴더를 확인합니다.
4. 변환된 최신 파일이 정상이고 암호가 적용되었다면, 개인정보 보호를 위해 _old_office_backup 폴더를 삭제합니다.
5. 백업 폴더 생성 후 처리를 사용했다면 백업 폴더에 암호화 전 원본이 남아 있는지 확인합니다.
6. 로그 파일을 저장한 경우, 로그에 파일 경로 등 업무 정보가 포함될 수 있으므로 보관 위치에 주의합니다.

9. 자주 발생하는 상황
Q. 구형 파일이 실패로 표시됩니다.
.xls, .doc, .ppt는 구형 형식입니다. 구형 Office 변환 옵션을 켜고 다시 실행해 주세요.

Q. 한글 작업 중 멈춘 것처럼 보입니다.
한글 접근 허용 창이 프로그램 뒤에 떠 있을 수 있습니다. 작업 표시줄 또는 한글 창을 확인하고 접근을 허용해 주세요. 큰 한글 파일은 시간이 더 걸릴 수 있습니다.

Q. 이미 암호가 있는 파일이 건너뛰어집니다.
암호 파일 스킵 옵션이 켜져 있으면 기존 암호 파일은 처리하지 않습니다. 기존 암호 파일을 다시 처리해야 하는 특수한 경우가 아니라면 켜두는 것을 권장합니다.

Q. 드래그앤드롭이 되지 않습니다.
배포용 실행 파일은 드래그앤드롭 기능을 포함해 제작되어 있습니다.
다만 사용 환경에 따라 드래그앤드롭이 동작하지 않는 경우에는 대상 폴더 선택과 파일 검색 방식으로도 동일하게 사용할 수 있습니다.

Q. 삭제 후보 점검 결과가 실제와 다를 수 있나요?
그럴 수 있습니다. 삭제 후보 점검은 파일명, 수정일, 확장자만 기준으로 판단합니다. 파일 내용은 확인하지 않으므로 최종 판단은 담당자가 해야 합니다.

10. 권장 운영 방식
개인정보 파일을 처리할 때는 아래 방식으로 사용하는 것을 권장합니다.

1. 원본 자료를 바로 처리하지 말고 사본 폴더에서 먼저 테스트합니다.
2. 암호는 5자 이상으로 설정하고, 조직의 암호 규칙이 있다면 그 규칙을 따릅니다.
3. 처리 후 성공 파일 중 일부를 직접 열어 암호 적용 여부를 확인합니다.
4. 구형 Office 변환을 사용했다면 백업 폴더를 확인 후 삭제합니다.
5. 최종 배포 또는 공유 전, 암호가 없는 원본 파일이 남아 있지 않은지 다시 확인합니다.
6. 삭제 후보 점검 결과는 정기적으로 확인하되, 실제 삭제 전에는 기관 보존 기준을 확인합니다.
"""

@dataclass
class FileItem:
    path: str
    selected: bool = True
    status: str = RESULT_PENDING
    detail: str = ""
    extension: str = ""
    accessible: bool = False
    backup_ready: bool = False
    last_result: str = RESULT_PENDING
    timestamp: str = ""

    def __post_init__(self):
        if not self.extension:
            self.extension = Path(self.path).suffix.lower()

class HwpAutomationError(Exception):
    """한글 COM 자동화 관련 사용자 정의 예외."""

class HwpOpenPasswordRequiredError(HwpAutomationError):
    """문서 열기 시 암호가 필요한 것으로 판단되는 경우."""

# —————————————————————————

# PDF 암호 설정 (pikepdf)

# —————————————————————————

def set_pdf_password(file_path, password):
    """
    pikepdf를 이용해 PDF에 열기 암호를 설정한다.
    저장은 임시 파일에 먼저 기록한 뒤 원본을 교체한다.
    반환값: (success: bool, message: str)
    """
    if pikepdf is None:
        return False, "pikepdf 라이브러리가 설치되지 않았습니다. 'pip install pikepdf'를 실행해 주세요."

    temp_path = None
    try:
        with pikepdf.open(file_path) as pdf:
            with tempfile.NamedTemporaryFile(
                delete=False,
                dir=os.path.dirname(file_path),
                prefix=f".{Path(file_path).stem}_",
                suffix=".pdf",
            ) as tmp:
                temp_path = tmp.name

            pdf.save(
                temp_path,
                encryption=pikepdf.Encryption(owner=password, user=password, R=6),
            )

        if not temp_path or not os.path.exists(temp_path):
            return False, "PDF 암호 설정 오류: 임시 저장 파일이 생성되지 않았습니다."
        if os.path.getsize(temp_path) <= 0:
            return False, "PDF 암호 설정 오류: 임시 저장 파일 크기가 0입니다."

        os.replace(temp_path, file_path)
        return True, "PDF 암호 설정 완료"
    except pikepdf.PasswordError:
        return False, "이미 암호가 걸린 PDF입니다. (열기 암호 필요)"
    except Exception as exc:
        return False, f"PDF 암호 설정 오류: {type(exc).__name__}: {exc}"
    finally:
        if temp_path and os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except OSError:
                pass

def detect_pdf_password(file_path):
    """
    PDF가 암호로 잠겨 있는지 확인한다.
    반환값: True(암호 있음) | False(없음) | None(확인 불가)
    """
    if pikepdf is None:
        return None
    try:
        with pikepdf.open(file_path):
            return False
    except pikepdf.PasswordError:
        return True
    except Exception:
        return None

# —————————————————————————

# Excel 암호 설정 (openpyxl / xlrd)

# —————————————————————————

def set_excel_password(file_path, password):
    """
    Excel 파일에 실제 파일 열기 암호를 설정한다.
    시트/통합문서 보호는 열기 암호가 아니므로 성공으로 처리하지 않는다.
    """
    ext = Path(file_path).suffix.lower()

    if ext == ".xls":
        return False, (
            ".xls(구형 Excel) 파일은 Python에서 파일 열기 암호 설정이 지원되지 않습니다. "
            "Excel에서 직접 저장하거나 .xlsx로 변환 후 사용해 주세요."
        )

    try:
        try:
            import io
            import msoffcrypto

            tmp_path = None
            with open(file_path, "rb") as f:
                office_file = msoffcrypto.OfficeFile(f)
                encrypted = io.BytesIO()
                office_file.encrypt(password, encrypted)

            try:
                encrypted_data = encrypted.getvalue()
                if not encrypted_data:
                    return False, "Excel 암호 설정 오류: 암호화 결과가 비어 있습니다."

                tmp_path = make_temp_encryption_path(file_path)
                with open(tmp_path, "wb") as f:
                    f.write(encrypted_data)

                if not os.path.exists(tmp_path):
                    return False, "Excel 암호 설정 오류: 임시 암호화 파일이 생성되지 않았습니다."
                if os.path.getsize(tmp_path) <= 0:
                    return False, "Excel 암호 설정 오류: 임시 암호화 파일 크기가 0입니다."

                replace_with_verified_temp(tmp_path, file_path)
                tmp_path = None
            finally:
                if tmp_path and os.path.exists(tmp_path):
                    try:
                        os.remove(tmp_path)
                    except OSError:
                        pass

            return True, "열기 암호 설정 완료"
        except ImportError:
            return False, (
                "msoffcrypto-tool이 설치되어 있지 않아 Excel 파일 열기 암호를 설정할 수 없습니다. "
                "'pip install msoffcrypto-tool' 설치 후 다시 실행해 주세요. "
                "시트/통합문서 보호는 파일 열기 암호가 아니므로 성공으로 처리하지 않았습니다."
            )
    except Exception as exc:
        return False, f"Excel 암호 설정 오류: {type(exc).__name__}: {exc}"

def detect_excel_password(file_path):
    """
    Excel 파일이 암호로 잠겨 있는지 확인한다.
    반환값: True | False | None
    """
    ext = Path(file_path).suffix.lower()
    try:
        with open(file_path, "rb") as f:
            header = f.read(8)
    except Exception:
        return None

    # OOXML(.xlsx/.xlsm) 파일이 ZIP이 아니라 OLE 컨테이너면
    # 실제 파일 열기 암호화된 경우가 많다.
    if ext in {".xlsx", ".xlsm"}:
        if header.startswith(b"\xD0\xCF\x11\xE0") and not zipfile.is_zipfile(file_path):
            return True

    # msoffcrypto-tool이 있으면 암호화 여부를 우선 확인한다.
    try:
        import msoffcrypto

        with open(file_path, "rb") as f:
            office_file = msoffcrypto.OfficeFile(f)
            if office_file.is_encrypted():
                return True
    except ImportError:
        pass
    except Exception:
        pass

    if ext == ".xls":
        return None
    if openpyxl is None:
        return None
    wb = None
    try:
        wb = openpyxl.load_workbook(file_path, read_only=True, keep_vba=(ext == ".xlsm"))
        return False
    except Exception as exc:
        msg = str(exc).lower()
        if (
            "encrypt" in msg
            or "password" in msg
            or "protected" in msg
            or "file is not a zip file" in msg
            or "ole file" in msg
        ):
            return True
        return None
    finally:
        if wb is not None:
            try:
                wb.close()
            except Exception:
                pass


def set_office_document_password(file_path, password):
    """
    Word/PowerPoint OOXML 파일에 열기 암호를 설정한다.
    msoffcrypto-tool이 필요하며, 구형 바이너리 형식(.doc/.ppt)은 안내 처리한다.
    """
    ext = Path(file_path).suffix.lower()

    if ext in {".doc", ".ppt"}:
        return False, (
            f"{ext}(구형 Office) 파일은 현재 Python에서 자동 파일 열기 암호 설정이 지원되지 않습니다. "
            "Office에서 직접 저장하거나 최신 형식으로 변환 후 사용해 주세요."
        )

    try:
        import io
        import msoffcrypto
    except ImportError:
        return False, (
            "Word/PowerPoint 열기 암호 설정에는 msoffcrypto-tool 라이브러리가 필요합니다. "
            "'pip install msoffcrypto-tool'을 실행해 주세요."
        )

    tmp_path = None
    try:
        with open(file_path, "rb") as f:
            office_file = msoffcrypto.OfficeFile(f)
            encrypted = io.BytesIO()
            office_file.encrypt(password, encrypted)

        encrypted_data = encrypted.getvalue()
        if not encrypted_data:
            return False, "Office 문서 암호 설정 오류: 암호화 결과가 비어 있습니다."

        tmp_path = make_temp_encryption_path(file_path)
        with open(tmp_path, "wb") as f:
            f.write(encrypted_data)

        if not os.path.exists(tmp_path):
            return False, "Office 문서 암호 설정 오류: 임시 암호화 파일이 생성되지 않았습니다."
        if os.path.getsize(tmp_path) <= 0:
            return False, "Office 문서 암호 설정 오류: 임시 암호화 파일 크기가 0입니다."

        replace_with_verified_temp(tmp_path, file_path)
        tmp_path = None

        return True, "열기 암호 설정 완료"
    except Exception as exc:
        return False, f"Office 문서 암호 설정 오류: {type(exc).__name__}: {exc}"
    finally:
        if tmp_path and os.path.exists(tmp_path):
            try:
                os.remove(tmp_path)
            except OSError:
                pass


def detect_office_document_password(file_path):
    """
    Word/PowerPoint 파일의 암호화 여부를 확인한다.
    반환값: True | False | None
    """
    ext = Path(file_path).suffix.lower()
    try:
        with open(file_path, "rb") as f:
            header = f.read(8)
    except Exception:
        return None

    if ext in {".docx", ".docm", ".pptx", ".pptm"}:
        if header.startswith(b"\xD0\xCF\x11\xE0") and not zipfile.is_zipfile(file_path):
            return True
        if zipfile.is_zipfile(file_path):
            return False

    try:
        import msoffcrypto

        with open(file_path, "rb") as f:
            office_file = msoffcrypto.OfficeFile(f)
            if office_file.is_encrypted():
                return True
        return False
    except ImportError:
        return None
    except Exception:
        return None


def make_unique_path(path):
    candidate = path
    base, ext = os.path.splitext(path)
    counter = 1
    while os.path.exists(candidate):
        candidate = f"{base}_{counter}{ext}"
        counter += 1
    return candidate


def make_temp_encryption_path(file_path):
    base_path = f"{file_path}.tmp_enc"
    if not os.path.exists(base_path):
        return base_path

    counter = 1
    while True:
        candidate = f"{base_path}_{counter}"
        if not os.path.exists(candidate):
            return candidate
        counter += 1


def replace_with_verified_temp(tmp_path, file_path):
    """
    검증된 임시 파일을 원본 위치로 교체한다.
    os.replace가 권한 문제로 실패하면 원본을 임시 원본 파일로 옮긴 뒤 교체하고,
    교체 실패 시 가능한 원본을 복구한다.
    """
    try:
        os.chmod(file_path, stat.S_IWRITE)
    except Exception:
        pass

    try:
        os.replace(tmp_path, file_path)
        return None
    except PermissionError:
        rollback_path = make_unique_path(f"{file_path}.tmp_orig")
        try:
            os.replace(file_path, rollback_path)
            os.replace(tmp_path, file_path)
            try:
                os.remove(rollback_path)
            except OSError:
                pass
            return None
        except Exception:
            if os.path.exists(rollback_path) and not os.path.exists(file_path):
                try:
                    os.replace(rollback_path, file_path)
                except Exception:
                    pass
            raise


def normalize_retention_text(text):
    normalized = unicodedata.normalize("NFC", text or "")
    return re.sub(r"\s+", "", normalized).lower()


def score_keyword_group(filename_key, keywords, first_score, additional_score):
    found = []
    score = 0
    details = []
    for keyword in keywords:
        keyword_key = normalize_retention_text(keyword)
        if keyword_key and keyword_key in filename_key:
            found.append(keyword)
            added = first_score if len(found) == 1 else additional_score
            score += added
            details.append(f"{keyword} +{added}")
    return found, score, details


def analyze_retention_risk(file_path):
    path = Path(file_path)
    filename = unicodedata.normalize("NFC", path.name)
    ext = path.suffix.lower()
    stat_result = os.stat(file_path)
    modified_at = datetime.fromtimestamp(stat_result.st_mtime)
    elapsed_days = max((datetime.now() - modified_at).days, 0)
    elapsed_years = elapsed_days / 365

    score = 0
    details = []
    detected_keywords = []

    if datetime.now() - modified_at > timedelta(days=365 * 5):
        score += 3
        details.append("5년 초과 +3")

    filename_key = normalize_retention_text(filename)
    high_found, high_score, high_details = score_keyword_group(
        filename_key, RETENTION_HIGH_RISK_KEYWORDS, 5, 2
    )
    general_found, general_score, general_details = score_keyword_group(
        filename_key, RETENTION_GENERAL_KEYWORDS, 3, 1
    )
    detected_keywords.extend(high_found)
    detected_keywords.extend(general_found)
    score += high_score + general_score
    details.extend(high_details)
    details.extend(general_details)

    if ext in RETENTION_OLD_DOC_EXTENSIONS:
        score += 2
        details.append(f"{ext} 구형 문서 +2")
    elif ext == ".pdf":
        score += 2
        details.append("PDF +2")
    elif ext in IMAGE_EXTENSIONS:
        score += 2
        details.append(f"{ext} 이미지 +2")
    elif ext in RETENTION_MODERN_DOC_EXTENSIONS:
        score += 1
        details.append(f"{ext} 신형 문서 +1")

    if score >= 8:
        classification = "🔴 삭제 권장"
        recommendation = "보관 필요 여부 확인 후 파기 검토"
    elif score >= 5:
        classification = "🟡 검토 필요"
        recommendation = "담당자 확인 후 보관/파기 결정"
    else:
        classification = "🟢 낮은 위험"
        recommendation = "필요 시 보관"

    return {
        "파일명": filename,
        "전체경로": str(path),
        "확장자": ext,
        "수정일": modified_at.strftime("%Y-%m-%d %H:%M:%S"),
        "경과연수": f"{elapsed_years:.1f}",
        "탐지키워드": ", ".join(detected_keywords) if detected_keywords else "-",
        "점수": score,
        "점수근거": " / ".join(details) if details else "특이사항 없음",
        "분류": classification,
        "권장조치": recommendation,
    }


def scan_retention_candidates(file_paths):
    results = []
    errors = []
    for file_path in file_paths:
        try:
            results.append(analyze_retention_risk(file_path))
        except Exception as exc:
            errors.append(f"{file_path}: {type(exc).__name__}: {exc}")
    return results, errors


def export_retention_report(results, csv_path):
    headers = ["파일명", "전체경로", "확장자", "수정일", "경과연수", "탐지키워드", "점수", "점수근거", "분류", "권장조치"]
    with open(csv_path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=headers)
        writer.writeheader()
        for row in results:
            writer.writerow({header: row.get(header, "") for header in headers})


def get_modern_office_path(file_path):
    ext = Path(file_path).suffix.lower()
    new_ext = {".xls": ".xlsx", ".doc": ".docx", ".ppt": ".pptx"}.get(ext)
    if not new_ext:
        return None
    return make_unique_path(str(Path(file_path).with_suffix(new_ext)))


def convert_old_office_to_modern(file_path):
    """
    설치된 Microsoft Office COM을 이용해 구형 Office 파일을 최신 형식으로 저장한다.
    원본 파일은 이 함수에서 이동/삭제하지 않는다.
    """
    if pythoncom is None or win32com is None:
        return False, None, "구형 Office 변환에는 pywin32와 Microsoft Office가 필요합니다."

    ext = Path(file_path).suffix.lower()
    new_path = get_modern_office_path(file_path)
    if not new_path:
        return False, None, f"{ext} 파일은 변환 대상이 아닙니다."

    initialized = False
    app = None
    doc = None
    try:
        pythoncom.CoInitialize()
        initialized = True

        if ext == ".xls":
            app = win32com.client.DispatchEx("Excel.Application")
            app.Visible = False
            app.DisplayAlerts = False
            doc = app.Workbooks.Open(os.path.abspath(file_path))
            doc.SaveAs(os.path.abspath(new_path), FileFormat=51)  # xlOpenXMLWorkbook
            doc.Close(SaveChanges=False)
        elif ext == ".doc":
            app = win32com.client.DispatchEx("Word.Application")
            app.Visible = False
            app.DisplayAlerts = 0
            doc = app.Documents.Open(os.path.abspath(file_path), ReadOnly=False)
            doc.SaveAs2(os.path.abspath(new_path), FileFormat=16)  # wdFormatXMLDocument
            doc.Close(SaveChanges=False)
        elif ext == ".ppt":
            app = win32com.client.DispatchEx("PowerPoint.Application")
            try:
                app.DisplayAlerts = 1
            except Exception:
                try:
                    app.DisplayAlerts = 7
                except Exception:
                    pass
            doc = app.Presentations.Open(os.path.abspath(file_path), WithWindow=False)
            doc.SaveAs(os.path.abspath(new_path), 24)  # ppSaveAsOpenXMLPresentation
            doc.Close()

        if not os.path.exists(new_path) or os.path.getsize(new_path) <= 0:
            return False, None, "최신 형식 변환 파일이 생성되지 않았거나 크기가 0입니다."
        return True, new_path, "최신 형식 변환 완료"
    except Exception as exc:
        if new_path and os.path.exists(new_path):
            try:
                os.remove(new_path)
            except OSError:
                pass
        return False, None, f"구형 Office 변환 오류: {type(exc).__name__}: {exc}"
    finally:
        if doc is not None:
            try:
                doc.Close(SaveChanges=False)
            except Exception:
                pass
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass
        if initialized:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

# —————————————————————————

# HWP COM 자동화

# —————————————————————————

class HwpComManager:
    """
    한글 COM 자동화를 담당하는 래퍼 클래스.

    주의:
    - 한글 버전, 보안 모듈, COM 등록 상태에 따라 일부 메서드/액션 이름이 다를 수 있다.
    - 특히 문서 암호 설정용 액션/옵션 문자열은 환경에 따라 다를 수 있으므로,
      아래 _set_password_* 메서드의 "여기서 수정 필요 가능" 주석을 확인해야 한다.
    """

    def __init__(self):
        self.hwp = None
        self._initialized = False

    def start(self):
        if pythoncom is None or win32com is None:
            raise HwpAutomationError(
                "pywin32가 설치되어 있지 않습니다. 명령 프롬프트에서 'pip install pywin32'를 실행해 주세요."
            )

        try:
            pythoncom.CoInitialize()
            self.hwp = win32com.client.DispatchEx("HWPFrame.HwpObject")
            self.hwp.XHwpWindows.Item(0).Visible = False

            try:
                self.hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
            except Exception:
                pass

            self._initialized = True
        except pywintypes.com_error as exc:
            self._safe_uninit()
            raise HwpAutomationError(
                "한글 COM 객체 생성에 실패했습니다.\n"
                "한컴오피스 한글이 설치되어 있는지, COM 등록이 정상인지 확인해 주세요.\n"
                f"상세 오류: {self._format_exception(exc)}"
            ) from exc
        except Exception as exc:
            self._safe_uninit()
            raise HwpAutomationError(
                "한글 COM 초기화 중 알 수 없는 오류가 발생했습니다.\n"
                f"상세 오류: {self._format_exception(exc)}"
            ) from exc

    def quit(self):
        try:
            if self.hwp is not None:
                try:
                    self.hwp.Quit()
                except Exception:
                    pass
        finally:
            self.hwp = None
            self._safe_uninit()

    def _safe_uninit(self):
        if self._initialized and pythoncom is not None:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
        self._initialized = False

    @staticmethod
    def _format_exception(exc):
        return f"{type(exc).__name__}: {exc}"

    def open_document(self, file_path):
        """
        한글 Open 메서드는 버전별로 받는 인자 개수가 다를 수 있어 여러 시그니처를 순차 시도한다.
        """
        ext = Path(file_path).suffix.lower()
        fmt = "HWPX" if ext == ".hwpx" else "HWP"
        open_options = [
            "",
            "forceopen:true",
            "lock:false",
            "forceopen:true;lock:false",
            "lock:false;forceopen:true",
        ]
        open_candidates = [
            ((file_path,), "Open(path)"),
        ]
        for option in open_options:
            open_candidates.append(((file_path, "", option), f"Open(path, '', '{option}')"))
            open_candidates.append(((file_path, fmt, option), f"Open(path, '{fmt}', '{option}')"))
            if option:
                open_candidates.append(((file_path, option), f"Open(path, '{option}')"))

        call_errors = []

        for args, signature_name in open_candidates:
            try:
                opened = self.hwp.Open(*args)
                if opened:
                    return
                call_errors.append(f"{signature_name}: False 반환")
            except pywintypes.com_error as exc:
                err_text = self._format_exception(exc)
                if self._looks_like_password_required(err_text):
                    raise HwpOpenPasswordRequiredError(
                        "문서를 여는 중 암호 입력이 필요한 것으로 보입니다.\n"
                        f"호출 방식: {signature_name}\n"
                        f"상세 오류: {err_text}"
                    ) from exc
                call_errors.append(f"{signature_name}: {err_text}")
            except Exception as exc:
                call_errors.append(f"{signature_name}: {self._format_exception(exc)}")

        raise HwpAutomationError(
            "문서 열기에 실패했습니다.\n"
            "현재 한글 COM Open 메서드 시그니처가 사용 환경과 다르거나, 파일 손상/권한 문제일 수 있습니다.\n"
            f"시도한 호출: {' | '.join(call_errors)}"
        )

    @staticmethod
    def _looks_like_password_required(message):
        lowered = message.lower()
        keywords = ["password", "암호", "비밀번호", "protected", "encrypt", "security", "보호"]
        wrong_parameter_markers = [
            "매개 변수의 개수가 잘못되었습니다", "wrong number of arguments",
            "type mismatch", "parameter", "argument",
        ]
        if any(marker in lowered for marker in wrong_parameter_markers):
            return False
        return any(keyword in lowered for keyword in keywords)

    def close_document(self):
        if self.hwp is None:
            return
        try:
            self.hwp.Clear(3)
        except Exception:
            try:
                self.hwp.Run("FileClose")
            except Exception:
                pass

    def set_password_and_save(self, file_path, password):
        """
        문서 암호 설정 후 저장한다.
        여러 방식을 순차 시도한다.
        """
        errors = []

        for method in (
            self._set_password_by_dialog_sendkeys,
            self._set_password_by_file_password_action,
            self._set_password_by_security_action,
        ):
            try:
                method(file_path, password)
                return
            except Exception as exc:
                errors.append(f"{method.__name__}: {self._format_exception(exc)}")

        raise HwpAutomationError(
            "문서 암호 설정/저장에 실패했습니다.\n"
            "현재 한글 버전에서 사용 중인 COM 액션명 또는 저장 옵션 문자열이 다를 수 있습니다.\n"
            f"시도한 방법들: {' | '.join(errors)}"
        )

    def _set_password_by_dialog_sendkeys(self, file_path, password):
        """
        한글의 문서 암호 대화상자를 열고 키 입력으로 암호를 넣는다.

        [수정됨] 암호창 대기 시간을 폴링 방식으로 개선:
        - 고정 sleep(1.2초) 대신 AppActivate 성공 여부를 폴링으로 확인해
          느린 환경에서도 타이밍 오류가 발생하지 않도록 했다.
        """
        if win32gui is None or win32con is None or win32process is None:
            raise HwpAutomationError("Win32 GUI 모듈을 사용할 수 없어 암호 대화상자를 자동 입력할 수 없습니다.")

        try:
            self.hwp.XHwpWindows.Item(0).Visible = True
        except Exception:
            pass

        fill_result = {"ok": False, "message": "암호 대화상자 자동 입력이 실행되지 않았습니다."}
        filler = threading.Thread(
            target=self._blind_fill_password_dialog,
            args=(password, fill_result),
            daemon=True,
        )
        filler.start()

        open_errors = []
        dialog_opened = False

        for action_name in ("FilePassword", "DocumentPassword", "SecurityPassword"):
            try:
                self.hwp.HAction.Run(action_name)
                dialog_opened = True
                break
            except Exception as exc:
                open_errors.append(f"HAction.Run({action_name}): {self._format_exception(exc)}")
            try:
                self.hwp.Run(action_name)
                dialog_opened = True
                break
            except Exception as exc:
                open_errors.append(f"Run({action_name}): {self._format_exception(exc)}")

        if not dialog_opened:
            raise HwpAutomationError(f"문서 암호 설정 대화상자를 열지 못했습니다. 상세: {' | '.join(open_errors)}")

        filler.join(timeout=15.0)
        if not fill_result["ok"]:
            raise HwpAutomationError(fill_result["message"])

        self._save_current_document(file_path)

    def _blind_fill_password_dialog(self, password, result_holder):
        """
        [수정됨] 고정 sleep(1.2초) → 폴링 방식으로 변경.
        AppActivate가 성공할 때까지 최대 8초 대기한 뒤 키 입력을 전송한다.
        """
        initialized = False
        try:
            if pythoncom is not None:
                pythoncom.CoInitialize()
                initialized = True
            shell = win32com.client.Dispatch("WScript.Shell")
            title_hints = ("문서 암호", "암호 설정", "비밀번호", "Password")

            # 폴링: 최대 8초(0.2초 간격 × 40회) 동안 암호창 활성화 시도
            activated = False
            for _ in range(40):
                for title_hint in title_hints:
                    try:
                        if shell.AppActivate(title_hint):
                            activated = True
                            break
                    except Exception:
                        pass
                if activated:
                    break
                time.sleep(0.2)

            if not activated:
                result_holder["message"] = (
                    "문서 암호 대화상자를 활성화하지 못해 자동 입력을 중단했습니다. "
                    "창 제목이 예상과 다르거나 대화상자가 열리지 않았을 수 있습니다."
                )
                return

            # 암호 대화상자가 활성화된 뒤 추가 0.2초 대기 후 키 입력
            time.sleep(0.2)
            self._send_password_key_sequence(shell, password, enter_at_end=True)
            result_holder["ok"] = True
            result_holder["message"] = "활성 암호 대화상자에 키 입력 완료"
        except Exception as exc:
            result_holder["message"] = f"활성 암호 대화상자 키 입력 중 오류: {self._format_exception(exc)}"
        finally:
            if initialized:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass

    @staticmethod
    def _send_password_key_sequence(shell, password, enter_at_end):
        for index in range(2):
            shell.SendKeys("^a")
            time.sleep(0.05)
            shell.SendKeys(password)
            time.sleep(0.12)
            if index < 1:
                shell.SendKeys("{TAB}")
                time.sleep(0.1)

        for _ in range(2):
            shell.SendKeys("{TAB}")
            time.sleep(0.08)

        if enter_at_end:
            shell.SendKeys("{ENTER}")
            time.sleep(0.5)

    def _fill_password_dialog(self, password, result_holder, existing_windows):
        dialog_hwnd = None

        for _ in range(120):
            dialog_hwnd = self._find_password_dialog_window(existing_windows)
            if dialog_hwnd:
                break
            time.sleep(0.1)

        if not dialog_hwnd:
            result_holder["message"] = "문서 암호 설정 대화상자 창을 찾지 못했습니다."
            return

        controls = self._get_child_controls(dialog_hwnd)
        edit_controls = [
            hwnd for hwnd, cls, text in controls
            if cls.lower() in ("edit", "richedit20w", "richedit20a")
        ]

        try:
            if len(edit_controls) >= 2:
                for hwnd in edit_controls[:4]:
                    win32gui.SendMessage(hwnd, win32con.WM_SETTEXT, 0, password)
                    time.sleep(0.05)

                button_hwnd = self._find_dialog_button(
                    controls,
                    preferred_texts=("설정", "확인", "OK", "적용"),
                    excluded_texts=("취소", "Cancel"),
                )
                if not button_hwnd:
                    result_holder["message"] = (
                        "문서 암호 설정 대화상자의 설정/확인 버튼을 찾지 못했습니다. "
                        f"감지된 컨트롤: {[(cls, text) for _, cls, text in controls]}"
                    )
                    return

                win32gui.SendMessage(button_hwnd, win32con.BM_CLICK, 0, 0)
            else:
                self._send_keys_to_password_dialog(dialog_hwnd, password)

            time.sleep(0.5)
            result_holder["ok"] = True
            result_holder["message"] = "문서 암호 설정 대화상자 자동 입력 완료"
        except Exception as exc:
            result_holder["message"] = f"문서 암호 설정 대화상자 자동 입력 중 오류: {self._format_exception(exc)}"

    @staticmethod
    def _send_keys_to_password_dialog(dialog_hwnd, password):
        shell = win32com.client.Dispatch("WScript.Shell")

        try:
            win32gui.ShowWindow(dialog_hwnd, win32con.SW_RESTORE)
        except Exception:
            pass
        try:
            win32gui.SetForegroundWindow(dialog_hwnd)
        except Exception:
            pass
        time.sleep(0.3)

        shell.SendKeys("+{TAB}")
        time.sleep(0.08)
        shell.SendKeys("+{TAB}")
        time.sleep(0.08)
        shell.SendKeys("{TAB}")
        time.sleep(0.08)

        for index in range(4):
            shell.SendKeys("^a")
            time.sleep(0.05)
            shell.SendKeys(password)
            time.sleep(0.12)
            if index < 3:
                shell.SendKeys("{TAB}")
                time.sleep(0.1)

        for _ in range(4):
            shell.SendKeys("{TAB}")
            time.sleep(0.08)
        shell.SendKeys("{ENTER}")

    @staticmethod
    def _snapshot_visible_windows():
        windows = set()
        current_pid = os.getpid()

        def callback(hwnd, _):
            try:
                if not win32gui.IsWindowVisible(hwnd):
                    return True
                _, pid = win32process.GetWindowThreadProcessId(hwnd)
                if pid != current_pid:
                    windows.add(hwnd)
            except Exception:
                pass
            return True

        win32gui.EnumWindows(callback, None)
        return windows

    @staticmethod
    def _find_password_dialog_window(existing_windows):
        matches = []
        current_pid = os.getpid()

        def callback(hwnd, _):
            if not win32gui.IsWindowVisible(hwnd):
                return True
            title = win32gui.GetWindowText(hwnd)

            if title and (APP_TITLE in title or "일괄 암호 설정기" in title):
                return True

            try:
                _, pid = win32process.GetWindowThreadProcessId(hwnd)
                if pid == current_pid:
                    return True
            except Exception:
                return True

            class_name = win32gui.GetClassName(hwnd)
            controls = HwpComManager._get_child_controls(hwnd)
            control_classes = {cls for _, cls, _ in controls}

            if not controls:
                return True
            if not ({"TkChild", "Edit", "Button", "HwpApp"} & control_classes):
                return True

            looks_like_password_window = bool(title and "암호" in title)
            looks_like_new_dialog = hwnd not in existing_windows and class_name in ("#32770", "TkTopLevel", "HNC_DIALOG")
            looks_like_hwp_tk_popup = hwnd not in existing_windows and "TkChild" in control_classes

            if looks_like_password_window or looks_like_new_dialog or looks_like_hwp_tk_popup:
                matches.append(hwnd)
            return True

        win32gui.EnumWindows(callback, None)
        return matches[0] if matches else None

    @staticmethod
    def _get_child_controls(parent_hwnd):
        controls = []

        def callback(hwnd, _):
            try:
                controls.append((hwnd, win32gui.GetClassName(hwnd), win32gui.GetWindowText(hwnd)))
            except Exception:
                pass
            return True

        win32gui.EnumChildWindows(parent_hwnd, callback, None)
        return controls

    @staticmethod
    def _find_dialog_button(controls, preferred_texts, excluded_texts):
        buttons = [
            (hwnd, text)
            for hwnd, cls, text in controls
            if cls.lower() == "button" and not any(excluded in text for excluded in excluded_texts)
        ]

        for preferred in preferred_texts:
            for hwnd, text in buttons:
                if preferred.lower() in text.lower():
                    return hwnd

        return buttons[0][0] if buttons else None

    def _set_password_by_file_password_action(self, file_path, password):
        action_names = [
            "FilePassword", "SecurityPassword", "SecurityFilePassword",
            "DocumentPassword", "FileSetPassword",
        ]
        parameter_names = [
            "Password", "Password2", "PasswordConfirm", "ConfirmPassword",
            "OpenPassword", "OpenPasswordConfirm", "OpenPasswd", "OpenPasswdConfirm",
            "DocumentPassword", "DocumentPasswordConfirm", "UserPassword", "UserPasswordConfirm",
        ]

        errors = []

        for action_name in action_names:
            try:
                hset = self.hwp.HParameterSet.HFilePassword.HSet
                self.hwp.HAction.GetDefault(action_name, hset)

                for parameter_name in parameter_names:
                    try:
                        setattr(self.hwp.HParameterSet.HFilePassword, parameter_name, password)
                    except Exception:
                        pass
                    try:
                        hset.SetItem(parameter_name, password)
                    except Exception:
                        pass

                executed = self.hwp.HAction.Execute(action_name, hset)
                if not executed:
                    errors.append(f"{action_name}: Execute가 False를 반환")
                    continue

                self._save_current_document(file_path)
                return
            except Exception as exc:
                errors.append(f"{action_name}: {self._format_exception(exc)}")

        raise HwpAutomationError(
            "문서 암호 설정 액션 실행에 실패했습니다. "
            f"상세: {' | '.join(errors)}"
        )

    def _save_current_document(self, file_path):
        ext = Path(file_path).suffix.lower()
        fmt = "HWPX" if ext == ".hwpx" else "HWP"

        save_errors = []
        for save_call in (
            lambda: self.hwp.HAction.Run("FileSave"),
            lambda: self.hwp.Run("FileSave"),
            lambda: self.hwp.Save(),
            lambda: self.hwp.SaveAs(file_path, fmt, ""),
        ):
            try:
                result = save_call()
                if result is None or result:
                    return
                save_errors.append("저장 호출이 False를 반환")
            except Exception as exc:
                save_errors.append(self._format_exception(exc))

        try:
            shell = win32com.client.Dispatch("WScript.Shell")
            for title_hint in ("한글", "Hwp", "Hancom"):
                try:
                    if shell.AppActivate(title_hint):
                        break
                except Exception:
                    pass
            time.sleep(0.2)
            shell.SendKeys("^s")
            time.sleep(0.8)
            return
        except Exception as exc:
            save_errors.append(f"Ctrl+S 저장 시도: {self._format_exception(exc)}")

        raise HwpAutomationError(f"문서 저장에 실패했습니다. 상세: {' | '.join(save_errors)}")

    def _set_password_by_security_action(self, file_path, password):
        action_names = ["FileSetSecurity", "DocSecurity", "DocumentSecurity"]

        last_error = None
        for action_name in action_names:
            try:
                action = self.hwp.CreateAction(action_name)
                pset = action.CreateSet()
                action.GetDefault(pset)

                for key in ("Password", "OpenPassword", "DocumentPassword"):
                    try:
                        pset.SetItem(key, password)
                    except Exception:
                        pass

                executed = action.Execute(pset)
                if not executed:
                    continue

                ext = Path(file_path).suffix.lower()
                fmt = "HWPX" if ext == ".hwpx" else "HWP"
                if not self.hwp.SaveAs(file_path, fmt, ""):
                    raise HwpAutomationError("보안 액션 후 저장에 실패했습니다.")
                return
            except Exception as exc:
                last_error = exc

        if last_error:
            raise last_error
        raise HwpAutomationError("보안 액션 방식으로 저장하지 못했습니다.")

    # —————————————————————————

    # GUI 애플리케이션

    # —————————————————————————

class HwpBatchPasswordApp:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("1250x860")
        self.root.minsize(1100, 700)

        self.folder_var = tk.StringVar()
        self.include_subfolders_var = tk.BooleanVar(value=True)
        self.skip_encrypted_var = tk.BooleanVar(value=True)
        self.backup_var = tk.BooleanVar(value=False)
        self.convert_old_office_var = tk.BooleanVar(value=False)
        self.password_var = tk.StringVar()
        self.password_confirm_var = tk.StringVar()
        self.show_password_var = tk.BooleanVar(value=False)

        self.total_files_var = tk.StringVar(value="0")
        self.selected_files_var = tk.StringVar(value="0")
        self.success_count_var = tk.StringVar(value="0")
        self.failed_count_var = tk.StringVar(value="0")
        self.skipped_count_var = tk.StringVar(value="0")
        self.progress_text_var = tk.StringVar(value="대기 중")
        self.current_file_var = tk.StringVar(value="-")
        self.sort_state = {"column": None, "descending": False}

        self.file_items = []
        self.file_map = {}
        self.logs = []
        self.failed_paths = set()
        self.current_run_paths = set()
        self.run_started_at = None
        self.retry_counts = {}
        self.dnd_enabled = False
        self.dnd_notice_logged = False
        self.backup_notice_pending = False
        self.retention_processing = False
        self.retention_results = []
        self.selection_filter_extensions = None
        self.selection_filter_label = "암호화 대상"

        self.worker_thread = None
        self.queue = queue.Queue()
        self.hwp_notice_ack = threading.Event()
        self.stop_requested = False
        self.processing = False

        self._build_ui()
        self._poll_queue()

    # ------------------------------------------------------------------
    # UI 구성
    # ------------------------------------------------------------------

    def _build_ui(self):
        self._build_menu()
        main = ttk.Frame(self.root, padding=10)
        main.pack(fill="both", expand=True)

        self._build_top_controls(main)
        self._build_file_tree(main)
        self._build_log_area(main)
        self._build_status_bar(main)
        self._setup_drag_and_drop()

    def _build_menu(self):
        menu_bar = tk.Menu(self.root)
        help_menu = tk.Menu(menu_bar, tearoff=0)
        help_menu.add_command(label="사용설명서", command=self.show_help_manual)
        help_menu.add_command(label="정보", command=self.show_about)
        menu_bar.add_cascade(label="도움말", menu=help_menu)
        self.root.configure(menu=menu_bar)

    def _build_top_controls(self, parent):
        frame = ttk.LabelFrame(parent, text="작업 설정", padding=10)
        frame.pack(fill="x", pady=(0, 10))

        frame.columnconfigure(1, weight=1)
        frame.columnconfigure(3, weight=1)
        frame.columnconfigure(4, weight=1)

        ttk.Label(frame, text="대상 폴더").grid(row=0, column=0, sticky="w", padx=(0, 6), pady=4)
        ttk.Entry(frame, textvariable=self.folder_var).grid(row=0, column=1, sticky="ew", pady=4)
        ttk.Button(frame, text="대상 폴더 선택", command=self.choose_folder).grid(row=0, column=2, padx=6, pady=4)
        ttk.Button(frame, text="파일 검색", command=self.search_files).grid(row=0, column=3, sticky="w", pady=4)

        ttk.Checkbutton(frame, text="하위 폴더 포함", variable=self.include_subfolders_var).grid(
            row=1, column=0, sticky="w", pady=4
        )
        ttk.Checkbutton(frame, text="암호 파일 스킵", variable=self.skip_encrypted_var).grid(
            row=1, column=1, sticky="w", pady=4
        )
        ttk.Checkbutton(frame, text="백업 폴더 생성 후 처리", variable=self.backup_var).grid(
            row=1, column=2, sticky="w", pady=4
        )
        # 지원 형식 안내
        ttk.Label(
            frame,
            text=(
                "지원 형식: HWP·HWPX (한글 COM) │ PDF (pikepdf) │ "
                "XLSX·XLSM (지원) │ DOCX·DOCM │ PPTX·PPTM │ "
                "XLS·DOC·PPT (구형 형식, 자동 암호 설정 제한)"
            ),
            foreground="gray",
        ).grid(row=1, column=3, sticky="w", pady=4)

        ttk.Label(frame, text="공통 암호").grid(row=2, column=0, sticky="w", padx=(0, 6), pady=4)
        self.password_entry = ttk.Entry(frame, textvariable=self.password_var, show="*")
        self.password_entry.grid(row=2, column=1, sticky="ew", pady=4)
        ttk.Label(frame, text="공통 암호 확인").grid(row=2, column=2, sticky="w", padx=(12, 6), pady=4)
        self.password_confirm_entry = ttk.Entry(frame, textvariable=self.password_confirm_var, show="*")
        self.password_confirm_entry.grid(row=2, column=3, sticky="ew", pady=4)

        button_frame = ttk.Frame(frame)
        button_frame.grid(row=3, column=0, columnspan=5, sticky="w", pady=(8, 0))

        ttk.Button(button_frame, text="전체 선택", command=self.select_all).pack(side="left", padx=(0, 6))
        ttk.Button(button_frame, text="전체 해제", command=self.deselect_all).pack(side="left", padx=(0, 6))
        ttk.Button(button_frame, text="실행 전 점검", command=self.preview_run).pack(side="left", padx=(0, 6))
        self.run_button = ttk.Button(button_frame, text="암호 설정 실행", command=self.run_password_setting)
        self.run_button.pack(side="left", padx=(0, 6))
        ttk.Button(button_frame, text="작업 중지", command=self.request_stop).pack(side="left", padx=(0, 6))
        ttk.Button(button_frame, text="암호 파일 목록 제거", command=self.remove_encrypted_files).pack(side="left", padx=(0, 6))
        ttk.Button(button_frame, text="로그 저장", command=self.save_log).pack(side="left", padx=(0, 6))
        ttk.Button(button_frame, text="실패 파일만 재선택", command=self.select_failed_only).pack(side="left", padx=(0, 6))
        ttk.Button(button_frame, text="실패 파일 폴더 열기", command=self.open_failed_folders).pack(side="left")
        self.retention_button = ttk.Button(button_frame, text="삭제 후보 점검", command=self.run_retention_scan)
        self.retention_button.pack(side="left", padx=(6, 0))
        ttk.Checkbutton(
            button_frame, text="구형 Office 변환", variable=self.convert_old_office_var
        ).pack(side="left", padx=(12, 0))
        ttk.Checkbutton(
            button_frame, text="암호 보기", variable=self.show_password_var,
            command=self.toggle_password_visibility
        ).pack(side="left", padx=(12, 0))

        selection_frame = ttk.Frame(frame)
        selection_frame.grid(row=4, column=0, columnspan=5, sticky="w", pady=(6, 0))

        ttk.Button(selection_frame, text="그림파일 목록 제거", command=self.remove_image_files).pack(side="left", padx=(0, 6))
        ttk.Button(selection_frame, text="파일 유형별 선택", command=self.show_file_type_selection).pack(side="left", padx=(0, 6))
        ttk.Button(selection_frame, text="다음 100개 선택", command=self.select_next_batch).pack(side="left", padx=(0, 6))

    def toggle_password_visibility(self):
        show_char = "" if self.show_password_var.get() else "*"
        self.password_entry.configure(show=show_char)
        self.password_confirm_entry.configure(show=show_char)

    def show_about(self):
        messagebox.showinfo(APP_TITLE, ABOUT_MESSAGE)

    def show_help_manual(self):
        window = tk.Toplevel(self.root)
        window.title("사용설명서")
        window.geometry("900x700")
        window.minsize(700, 500)
        window.transient(self.root)

        frame = ttk.Frame(window, padding=10)
        frame.pack(fill="both", expand=True)

        text = tk.Text(frame, wrap="word", font=("맑은 고딕", 10))
        yscroll = ttk.Scrollbar(frame, orient="vertical", command=text.yview)
        text.configure(yscrollcommand=yscroll.set)

        text.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)

        text.insert("1.0", HELP_MANUAL_TEXT)
        text.configure(state="disabled")

        ttk.Button(frame, text="닫기", command=window.destroy).grid(row=1, column=0, columnspan=2, sticky="e", pady=(8, 0))

    def _build_file_tree(self, parent):
        frame = ttk.LabelFrame(parent, text="대상 파일 목록", padding=10)
        frame.pack(fill="both", expand=True, pady=(0, 10))

        columns = ("selected", "status", "ext", "path", "detail")
        self.tree = ttk.Treeview(frame, columns=columns, show="headings", height=18)
        self.update_tree_headings()

        self.tree.column("selected", width=60, anchor="center", stretch=False)
        self.tree.column("status", width=80, anchor="center", stretch=False)
        self.tree.column("ext", width=70, anchor="center", stretch=False)
        self.tree.column("path", width=550, anchor="w")
        self.tree.column("detail", width=360, anchor="w")

        yscroll = ttk.Scrollbar(frame, orient="vertical", command=self.tree.yview)
        xscroll = ttk.Scrollbar(frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll.grid(row=1, column=0, sticky="ew")

        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)

        self.tree.bind("<Button-1>", self.on_tree_click)
        self.tree.bind("<Double-1>", self.on_tree_double_click)

    def _build_log_area(self, parent):
        frame = ttk.LabelFrame(parent, text="처리 로그", padding=10)
        frame.pack(fill="both", expand=False)

        self.log_text = tk.Text(frame, height=12, wrap="none", state="disabled")
        yscroll = ttk.Scrollbar(frame, orient="vertical", command=self.log_text.yview)
        xscroll = ttk.Scrollbar(frame, orient="horizontal", command=self.log_text.xview)
        self.log_text.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)

        self.log_text.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll.grid(row=1, column=0, sticky="ew")
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)

    def _build_status_bar(self, parent):
        frame = ttk.LabelFrame(parent, text="상태", padding=10)
        frame.pack(fill="x", pady=(10, 0))

        info = ttk.Frame(frame)
        info.pack(fill="x")

        status_items = [
            ("총 검색 파일 수", self.total_files_var),
            ("선택 파일 수", self.selected_files_var),
            ("성공 수", self.success_count_var),
            ("실패 수", self.failed_count_var),
            ("건너뜀 수", self.skipped_count_var),
            ("진행 상태", self.progress_text_var),
        ]

        for idx, (label, var) in enumerate(status_items):
            ttk.Label(info, text=f"{label}:").grid(row=0, column=idx * 2, sticky="w", padx=(0, 4))
            ttk.Label(info, textvariable=var).grid(row=0, column=idx * 2 + 1, sticky="w", padx=(0, 12))

        ttk.Label(frame, text="현재 처리 파일:").pack(anchor="w", pady=(8, 0))
        ttk.Label(frame, textvariable=self.current_file_var).pack(anchor="w")

        self.progress = ttk.Progressbar(frame, mode="determinate")
        self.progress.pack(fill="x", pady=(8, 0))

    def _setup_drag_and_drop(self):
        if DND_FILES is None:
            if not self.dnd_notice_logged:
                self.add_log("드래그앤드롭 기능을 사용하려면 tkinterdnd2가 필요합니다.")
                self.dnd_notice_logged = True
            return
        try:
            self.tree.drop_target_register(DND_FILES)
            self.tree.dnd_bind("<<Drop>>", self.on_files_dropped)
            self.dnd_enabled = True
        except Exception:
            self.dnd_enabled = False
            if not self.dnd_notice_logged:
                self.add_log("드래그앤드롭 기능을 사용하려면 tkinterdnd2가 필요합니다.")
                self.dnd_notice_logged = True

    # ------------------------------------------------------------------
    # 폴더 / 파일 조회
    # ------------------------------------------------------------------

    def choose_folder(self):
        folder = filedialog.askdirectory(title="대상 폴더 선택")
        if folder:
            self.folder_var.set(folder)

    def add_log(self, message):
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        line = f"[{timestamp}] {message}\n"
        self.log_text.configure(state="normal")
        self.log_text.insert("end", line)
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def clear_logs(self):
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")
        self.logs.clear()

    def request_stop(self):
        if not self.processing:
            messagebox.showinfo(APP_TITLE, "현재 진행 중인 작업이 없습니다.")
            return
        if self.stop_requested:
            return
        self.stop_requested = True
        self.progress_text_var.set("중지 요청 중")
        self.add_log("사용자가 작업 중지를 요청했습니다. 현재 파일 처리 후 안전하게 종료합니다.")

    def validate_folder(self):
        folder = self.folder_var.get().strip()
        if not folder:
            messagebox.showwarning(APP_TITLE, "대상 폴더를 먼저 선택해 주세요.")
            return None
        if not os.path.isdir(folder):
            messagebox.showerror(APP_TITLE, "선택한 경로가 폴더가 아니거나 존재하지 않습니다.")
            return None
        return os.path.abspath(folder)

    @staticmethod
    def derive_common_parent(file_items):
        parent_paths = [os.path.dirname(item.path) for item in file_items if item.path]
        if not parent_paths:
            return None

        try:
            root_folder = os.path.commonpath(parent_paths)
        except ValueError:
            # 드라이브가 섞인 경우에는 첫 번째 파일의 폴더를 기준으로 삼는다.
            root_folder = parent_paths[0]

        if not os.path.isdir(root_folder):
            root_folder = os.path.dirname(root_folder)
        return os.path.abspath(root_folder) if root_folder else None

    def get_execution_root_folder(self, selected_items):
        folder = self.folder_var.get().strip()
        if folder:
            if not os.path.isdir(folder):
                messagebox.showerror(APP_TITLE, "선택한 경로가 폴더가 아니거나 존재하지 않습니다.")
                return None
            return os.path.abspath(folder)

        root_folder = self.derive_common_parent(selected_items)
        if not root_folder:
            messagebox.showwarning(APP_TITLE, "대상 폴더를 선택하거나 파일을 추가해 주세요.")
            return None

        self.folder_var.set(root_folder)
        self.add_log(f"대상 폴더가 비어 있어 목록 파일 기준 폴더를 사용합니다: {root_folder}")
        return root_folder

    def search_files(self):
        folder = self.validate_folder()
        if not folder:
            return

        self.clear_logs()
        self.file_items.clear()
        self.file_map.clear()
        self.failed_paths.clear()

        include_sub = self.include_subfolders_var.get()
        found_paths = self.collect_target_files(folder, include_sub)

        for path in found_paths:
            item = FileItem(path=os.path.abspath(path))
            self.file_items.append(item)
            self.file_map[item.path] = item

        self.refresh_tree()
        self.update_counters()
        self.progress.configure(value=0, maximum=max(len(self.file_items), 1))

        self.add_log(f"파일 검색 완료: {len(self.file_items)}건")

    def on_files_dropped(self, event):
        try:
            dropped = [os.path.abspath(path) for path in self.root.tk.splitlist(event.data)]
        except Exception as exc:
            self.add_log(f"드래그앤드롭 처리 중 오류: {type(exc).__name__}: {exc}")
            return

        added_count = 0
        for dropped_path in dropped:
            if os.path.isdir(dropped_path):
                for path in self.collect_target_files(dropped_path, True):
                    added_count += self._add_file_item(path)
            elif os.path.isfile(dropped_path) and self.should_include_file(os.path.basename(dropped_path)):
                added_count += self._add_file_item(dropped_path)

        if added_count:
            if not self.folder_var.get().strip():
                root_folder = self.derive_common_parent(self.file_items)
                if root_folder:
                    self.folder_var.set(root_folder)
                    self.add_log(f"대상 폴더를 드래그된 파일 기준으로 자동 설정했습니다: {root_folder}")
            self.refresh_tree()
            self.update_counters()
            self.progress.configure(value=0, maximum=max(len(self.file_items), 1))
            self.add_log(f"드래그앤드롭으로 파일 {added_count}건을 추가했습니다.")
        else:
            self.add_log("드래그앤드롭으로 추가된 지원 대상 파일이 없습니다.")

    def _add_file_item(self, path):
        abs_path = os.path.abspath(path)
        if abs_path in self.file_map:
            return 0
        item = FileItem(path=abs_path)
        self.file_items.append(item)
        self.file_map[item.path] = item
        return 1

    def collect_target_files(self, root_folder, include_subfolders):
        results = []
        if include_subfolders:
            def _on_walk_error(exc):
                if isinstance(exc, PermissionError):
                    self.add_log(f"[건너뜀] 폴더 접근 권한 없음: {exc.filename}")
                else:
                    self.add_log(f"[건너뜀] 폴더 탐색 오류: {type(exc).__name__}: {exc}")

            for current_root, dirs, files in os.walk(root_folder, onerror=_on_walk_error):
                dirs[:] = [
                    d for d in dirs
                    if not self.is_hidden_name(d) and not self.is_backup_dir_name(d)
                ]
                for name in files:
                    if self.should_include_file(name):
                        results.append(os.path.join(current_root, name))
        else:
            try:
                for name in os.listdir(root_folder):
                    full_path = os.path.join(root_folder, name)
                    if os.path.isfile(full_path) and self.should_include_file(name):
                        results.append(full_path)
            except Exception as exc:
                messagebox.showerror(APP_TITLE, f"폴더 조회 중 오류가 발생했습니다.\n{type(exc).__name__}: {exc}")
                return []

        return sorted({os.path.abspath(p) for p in results})

    @staticmethod
    def is_hidden_name(name):
        return name.startswith(".")

    @staticmethod
    def is_backup_dir_name(name):
        return name.startswith("_backup_hwp_password_") or name == OLD_OFFICE_BACKUP_DIR

    @staticmethod
    def should_include_file(filename):
        lower = filename.lower()
        # 임시 파일 제외 (~$로 시작하는 Office 임시 파일)
        if lower.startswith("~$"):
            return False
        return Path(lower).suffix in SEARCHABLE_EXTENSIONS

    # ------------------------------------------------------------------
    # 트리 뷰 조작
    # ------------------------------------------------------------------

    def update_tree_headings(self):
        labels = {
            "selected": "선택",
            "status": "상태",
            "ext": "확장자",
            "path": "파일 경로",
            "detail": "상세 메시지",
        }
        current = self.sort_state.get("column")
        descending = self.sort_state.get("descending", False)

        for column, label in labels.items():
            suffix = ""
            if column == current:
                suffix = " ▼" if descending else " ▲"
            self.tree.heading(column, text=f"{label}{suffix}", command=lambda c=column: self.on_tree_heading_click(c))


    def on_tree_heading_click(self, column):
        previous = self.sort_state.get("column")
        descending = self.sort_state.get("descending", False)

        if previous == column:
            descending = not descending
        else:
            descending = False

        self.sort_state = {"column": column, "descending": descending}
        self.sort_file_items(column, descending)
        self.refresh_tree()


    def sort_file_items(self, column, descending=False):
        def normalize(value):
            if value is None:
                return ""
            if isinstance(value, bool):
                return int(value)
            return str(value).lower()

        key_map = {
            "selected": lambda item: int(item.selected),
            "status": lambda item: item.status,
            "ext": lambda item: item.extension,
            "path": lambda item: item.path,
            "detail": lambda item: item.detail,
        }
        key_func = key_map.get(column, lambda item: item.path)
        self.file_items.sort(
            key=lambda item: (normalize(key_func(item)), normalize(item.path)),
            reverse=descending,
        )
    def refresh_tree(self):
        self.update_tree_headings()
        for item_id in self.tree.get_children():
            self.tree.delete(item_id)

        for item in self.file_items:
            self.tree.insert(
                "", "end", iid=item.path,
                values=(
                    TREE_CHECKED if item.selected else TREE_UNCHECKED,
                    item.status,
                    item.extension,
                    item.path,
                    item.detail,
                ),
            )

    def update_tree_row(self, item: FileItem):
        if self.tree.exists(item.path):
            self.tree.item(
                item.path,
                values=(
                    TREE_CHECKED if item.selected else TREE_UNCHECKED,
                    item.status,
                    item.extension,
                    item.path,
                    item.detail,
                ),
            )

    def on_tree_click(self, event):
        region = self.tree.identify("region", event.x, event.y)
        column = self.tree.identify_column(event.x)
        item_id = self.tree.identify_row(event.y)

        if region == "cell" and column == "#1" and item_id:
            item = self.file_map.get(item_id)
            if item:
                item.selected = not item.selected
                self.update_tree_row(item)
                self.update_counters()
            return "break"
        return None

    def on_tree_double_click(self, event):
        item_id = self.tree.identify_row(event.y)
        if not item_id:
            return None

        item = self.file_map.get(item_id)
        if not item:
            return None

        try:
            os.startfile(item.path)
            self.add_log(f"[열기] {item.path} | 연결된 프로그램으로 파일을 열었습니다.")
        except Exception as exc:
            messagebox.showerror(
                APP_TITLE,
                f"파일을 여는 중 오류가 발생했습니다.\n{type(exc).__name__}: {exc}",
            )
        return "break"

    def set_action_buttons_state(self):
        if hasattr(self, "run_button"):
            self.run_button.configure(state="disabled" if self.retention_processing else "normal")
        if hasattr(self, "retention_button"):
            self.retention_button.configure(state="disabled" if self.processing or self.retention_processing else "normal")

    def run_retention_scan(self):
        if self.processing:
            messagebox.showinfo(APP_TITLE, "암호화 작업 중에는 삭제 후보 점검을 실행할 수 없습니다.")
            return
        if self.retention_processing:
            messagebox.showinfo(APP_TITLE, "이미 삭제 후보 점검이 진행 중입니다.")
            return

        selected_items = self.get_selected_items()
        if not selected_items:
            messagebox.showwarning(APP_TITLE, "점검할 파일을 선택해 주세요.")
            return

        file_paths = [item.path for item in selected_items]
        self.retention_processing = True
        self.retention_results = []
        self.set_action_buttons_state()
        self.progress_text_var.set("삭제 후보 점검 중")
        self.current_file_var.set("-")
        self.add_log(f"삭제 후보 점검을 시작합니다. 대상 파일 수: {len(file_paths)}")

        threading.Thread(
            target=self._worker_retention_scan,
            args=(file_paths,),
            daemon=True,
        ).start()

    def _worker_retention_scan(self, file_paths):
        results, errors = scan_retention_candidates(file_paths)
        self.queue.put(("retention_done", {"results": results, "errors": errors}))

    def show_retention_results_window(self, results):
        window = tk.Toplevel(self.root)
        window.title("삭제 후보 점검 결과")
        window.geometry("1200x620")
        window.minsize(900, 450)
        window.transient(self.root)

        frame = ttk.Frame(window, padding=10)
        frame.pack(fill="both", expand=True)

        ttk.Label(
            frame,
            text=(
                "이 결과는 파일명, 수정일, 확장자 기준의 추정 결과입니다.\n"
                "실제 파기 여부는 담당자가 파일 내용과 기관 보존 기준을 확인한 뒤 결정하십시오.\n"
                "본 프로그램은 파일을 자동 삭제하지 않습니다."
            ),
            foreground="red",
        ).pack(anchor="w", pady=(0, 8))

        table_frame = ttk.Frame(frame)
        table_frame.pack(fill="both", expand=True)

        columns = ("filename", "modified", "years", "keywords", "score", "details", "class", "action")
        tree = ttk.Treeview(table_frame, columns=columns, show="headings", height=18)
        headings = {
            "filename": "파일명",
            "modified": "수정일",
            "years": "경과연수",
            "keywords": "탐지키워드",
            "score": "점수",
            "details": "점수근거",
            "class": "분류",
            "action": "권장조치",
        }
        widths = {
            "filename": 220,
            "modified": 150,
            "years": 80,
            "keywords": 170,
            "score": 60,
            "details": 320,
            "class": 110,
            "action": 220,
        }
        for col in columns:
            tree.heading(col, text=headings[col])
            tree.column(col, width=widths[col], anchor="w", stretch=(col in {"filename", "details", "action"}))

        yscroll = ttk.Scrollbar(table_frame, orient="vertical", command=tree.yview)
        xscroll = ttk.Scrollbar(table_frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)

        tree.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll.grid(row=1, column=0, sticky="ew")
        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)

        path_by_iid = {}
        for index, row in enumerate(sorted(results, key=lambda r: (-int(r.get("점수", 0)), r.get("파일명", "")))):
            iid = str(index)
            path_by_iid[iid] = row.get("전체경로", "")
            tree.insert(
                "",
                "end",
                iid=iid,
                values=(
                    row.get("파일명", ""),
                    row.get("수정일", ""),
                    row.get("경과연수", ""),
                    row.get("탐지키워드", ""),
                    row.get("점수", ""),
                    row.get("점수근거", ""),
                    row.get("분류", ""),
                    row.get("권장조치", ""),
                ),
            )

        def open_selected_location():
            selected = tree.selection()
            if not selected:
                messagebox.showinfo(APP_TITLE, "위치를 열 파일을 선택해 주세요.")
                return
            file_path = path_by_iid.get(selected[0])
            if not file_path:
                return
            try:
                subprocess.run(["explorer", "/select,", os.path.normpath(file_path)], check=False)
            except Exception as exc:
                messagebox.showerror(APP_TITLE, f"파일 위치를 여는 중 오류가 발생했습니다.\n{type(exc).__name__}: {exc}")

        def save_csv():
            csv_path = filedialog.asksaveasfilename(
                title="삭제 후보 점검 CSV 저장",
                defaultextension=".csv",
                filetypes=[("CSV 파일", "*.csv"), ("모든 파일", "*.*")],
            )
            if not csv_path:
                return
            try:
                export_retention_report(results, csv_path)
                messagebox.showinfo(APP_TITLE, "삭제 후보 점검 CSV를 저장했습니다.")
            except Exception as exc:
                messagebox.showerror(APP_TITLE, f"CSV 저장 중 오류가 발생했습니다.\n{type(exc).__name__}: {exc}")

        tree.bind("<Double-1>", lambda _event: open_selected_location())

        button_row = ttk.Frame(frame)
        button_row.pack(fill="x", pady=(8, 0))
        ttk.Button(button_row, text="CSV 저장", command=save_csv).pack(side="left")
        ttk.Button(button_row, text="파일 위치 열기", command=open_selected_location).pack(side="left", padx=(8, 0))
        ttk.Button(button_row, text="닫기", command=window.destroy).pack(side="right")

    def _show_hwp_permission_dialog(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("한글 권한 확인")
        dialog.resizable(False, False)
        dialog.transient(self.root)

        frame = ttk.Frame(dialog, padding=16)
        frame.pack(fill="both", expand=True)

        ttk.Label(
            frame,
            text=(
                "이제부터 HWP/HWPX 파일을 처리합니다.\n\n"
                "한글 문서는 암호 입력 과정에서 자동 키보드 입력이 사용됩니다.\n"
                "한컴에서 접근 권한 확인 창이 뜨면 '모두 허용'을 눌러야 합니다.\n"
                "접근 권한 확인 창이 이 프로그램이나 한글 창 뒤에 가려질 수 있으니,\n"
                "보이지 않으면 뒤쪽 창을 확인해 주세요.\n\n"
                "권한창 확인 준비가 되면 아래 버튼을 눌러 계속 진행하세요."
            ),
            justify="left",
        ).pack(anchor="w")

        button_row = ttk.Frame(frame)
        button_row.pack(fill="x", pady=(14, 0))

        def continue_run():
            self.hwp_notice_ack.set()
            dialog.destroy()

        def stop_run():
            self.request_stop()
            self.hwp_notice_ack.set()
            dialog.destroy()

        ttk.Button(button_row, text="권한창 확인 후 계속", command=continue_run).pack(side="right")
        ttk.Button(button_row, text="작업 중지", command=stop_run).pack(side="right", padx=(0, 8))

        dialog.update_idletasks()
        x = self.root.winfo_rootx() + max((self.root.winfo_width() - dialog.winfo_width()) // 2, 0)
        y = self.root.winfo_rooty() + max((self.root.winfo_height() - dialog.winfo_height()) // 2, 0)
        dialog.geometry(f"+{x}+{y}")
        dialog.grab_set()
        dialog.focus_force()
        dialog.wait_window()

    def select_all(self):
        for item in self.file_items:
            item.selected = True
            self.update_tree_row(item)
        self.update_counters()

    def deselect_all(self):
        for item in self.file_items:
            item.selected = False
            self.update_tree_row(item)
        self.update_counters()

    def select_failed_only(self):
        has_failed = False
        for item in self.file_items:
            item.selected = (item.last_result == RESULT_FAILED) and (item.path in self.failed_paths)
            has_failed = has_failed or item.selected
            self.update_tree_row(item)
        self.update_counters()
        if has_failed:
            self.add_log("이전 실행에서 실패한 파일만 다시 선택했습니다.")
        else:
            self.add_log("선택할 실패 파일이 없습니다.")

    def remove_image_files(self):
        if not self.file_items:
            messagebox.showwarning(APP_TITLE, "먼저 파일 검색을 실행해 주세요.")
            return
        if self.processing or self.retention_processing:
            messagebox.showinfo(APP_TITLE, "작업이 진행 중일 때는 목록을 변경할 수 없습니다.")
            return

        removed_items = [item for item in self.file_items if item.extension.lower() in IMAGE_EXTENSIONS]
        if not removed_items:
            self.add_log("목록에서 제거할 그림파일이 없습니다.")
            return

        removed_paths = {item.path for item in removed_items}
        self.file_items = [item for item in self.file_items if item.path not in removed_paths]
        self.file_map = {item.path: item for item in self.file_items}
        self.failed_paths.intersection_update(self.file_map.keys())
        self.refresh_tree()
        self.update_counters()
        self.progress.configure(maximum=max(len(self.file_items), 1))
        self.add_log(
            f"그림파일 {len(removed_items)}건을 목록에서 제거했습니다. "
            f"현재 목록에는 {len(self.file_items)}건이 남아 있습니다."
        )

    def show_file_type_selection(self):
        if not self.file_items:
            messagebox.showwarning(APP_TITLE, "먼저 파일 검색을 실행해 주세요.")
            return

        window = tk.Toplevel(self.root)
        window.title("파일 유형별 선택")
        window.transient(self.root)
        window.resizable(False, False)

        frame = ttk.Frame(window, padding=12)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text="선택할 파일 유형").grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 8))

        options = [
            ("암호화 대상 전체", SUPPORTED_EXTENSIONS),
            ("HWP/HWPX", HWP_EXTENSIONS),
            ("HWP만", {".hwp"}),
            ("HWPX만", {".hwpx"}),
            ("PDF", PDF_EXTENSIONS),
            ("Excel", EXCEL_EXTENSIONS),
            ("Word/PowerPoint", WORD_EXTENSIONS | POWERPOINT_EXTENSIONS),
            ("그림파일", IMAGE_EXTENSIONS),
        ]

        for index, (label, extensions) in enumerate(options, start=1):
            ttk.Button(
                frame,
                text=label,
                width=18,
                command=lambda l=label, e=extensions: (
                    self.select_by_file_type(l, e),
                    window.destroy(),
                ),
            ).grid(row=(index - 1) // 3 + 1, column=(index - 1) % 3, padx=4, pady=4, sticky="ew")

        ttk.Button(frame, text="닫기", command=window.destroy).grid(
            row=4, column=0, columnspan=3, sticky="e", pady=(8, 0)
        )

        window.update_idletasks()
        x = self.root.winfo_rootx() + max((self.root.winfo_width() - window.winfo_width()) // 2, 0)
        y = self.root.winfo_rooty() + max((self.root.winfo_height() - window.winfo_height()) // 2, 0)
        window.geometry(f"+{x}+{y}")
        window.grab_set()
        window.focus_force()

    def select_by_file_type(self, label, extensions):
        extensions = {ext.lower() for ext in extensions}
        self.selection_filter_extensions = extensions
        self.selection_filter_label = label
        selected_count = 0

        for item in self.file_items:
            item.selected = item.extension.lower() in extensions
            if item.selected:
                selected_count += 1
            self.update_tree_row(item)

        self.update_counters()
        self.add_log(f"파일 유형 선택: {label} {selected_count}건을 선택했습니다.")

    def select_next_batch(self, batch_size=DEFAULT_SELECTION_BATCH_SIZE):
        if not self.file_items:
            messagebox.showwarning(APP_TITLE, "먼저 파일 검색을 실행해 주세요.")
            return

        extensions = self.selection_filter_extensions or SUPPORTED_EXTENSIONS
        extensions = {ext.lower() for ext in extensions}
        eligible_items = [
            item for item in self.file_items
            if item.extension.lower() in extensions
            and item.last_result not in {RESULT_SUCCESS, RESULT_FAILED, RESULT_SKIPPED}
        ]

        if not eligible_items:
            self.deselect_all()
            self.add_log(
                f"{self.selection_filter_label} 중 다음 {batch_size}개로 선택할 대기 파일이 없습니다. "
                "실패 파일은 '실패 파일만 재선택'으로 따로 선택할 수 있습니다."
            )
            return

        selected_paths = {item.path for item in eligible_items[:batch_size]}
        for item in self.file_items:
            item.selected = item.path in selected_paths
            self.update_tree_row(item)

        self.update_counters()
        self.add_log(
            f"{self.selection_filter_label} 대기 파일 중 다음 {min(batch_size, len(eligible_items))}건을 선택했습니다."
        )

    def open_failed_folders(self):
        failed_items = [
            item for item in self.file_items
            if item.last_result == RESULT_FAILED and item.path in self.failed_paths
        ]
        if not failed_items:
            self.add_log("열 수 있는 실패 파일 폴더가 없습니다.")
            return

        opened_dirs = []
        for folder_path in sorted({os.path.dirname(item.path) for item in failed_items}):
            try:
                os.startfile(folder_path)
                opened_dirs.append(folder_path)
            except Exception as exc:
                self.add_log(f"[실패] {folder_path} | 폴더 열기 오류: {type(exc).__name__}: {exc}")

        if opened_dirs:
            self.add_log(f"실패 파일이 있는 폴더 {len(opened_dirs)}곳을 열었습니다.")

    def remove_encrypted_files(self):
        if not self.file_items:
            messagebox.showwarning(APP_TITLE, "먼저 파일 검색을 실행해 주세요.")
            return

        removed_items = []
        kept_items = []
        total = len(self.file_items)

        self.progress.configure(value=0, maximum=max(total, 1))
        self.progress_text_var.set(f"암호 파일 목록 제거 준비 중 (0/{total})")
        self.current_file_var.set("-")
        self.root.update_idletasks()

        for index, item in enumerate(self.file_items, start=1):
            self.progress.configure(value=index)
            self.progress_text_var.set(f"암호 파일 목록 제거 중 ({index}/{total})")
            self.current_file_var.set(item.path)
            if index == 1 or index % 25 == 0 or index == total:
                self.root.update_idletasks()

            encrypted = self.detect_encrypted_file(item.path, item.extension)
            if encrypted is True:
                removed_items.append(item)
            else:
                kept_items.append(item)

        self.file_items = kept_items
        self.file_map = {item.path: item for item in self.file_items}
        self.failed_paths.intersection_update(self.file_map.keys())
        self.refresh_tree()
        self.update_counters()
        self.progress_text_var.set(f"암호 파일 목록 제거 완료 ({len(removed_items)}건 제거 / {len(self.file_items)}건 남음)")
        self.current_file_var.set("-")
        self.progress.configure(value=0, maximum=max(len(self.file_items), 1))
        self.root.update_idletasks()

        if removed_items:
            self.add_log(
                f"암호화된 파일 {len(removed_items)}건을 목록에서 제거했습니다. "
                f"현재 목록에는 {len(self.file_items)}건이 남아 있습니다."
            )
            for item in removed_items:
                self.add_result_log(item.path, "목록제거", "암호화된 파일로 판정되어 목록에서 제거")
        else:
            self.add_log(f"목록에서 제거할 암호화 파일이 없습니다. 현재 목록에는 {len(self.file_items)}건이 남아 있습니다.")

    @staticmethod
    def detect_encrypted_file(file_path, extension=None):
        ext = (extension or Path(file_path).suffix.lower()).lower()

        if ext in HWP_EXTENSIONS:
            return HwpBatchPasswordApp.detect_existing_password(file_path)
        if ext in PDF_EXTENSIONS:
            return detect_pdf_password(file_path)
        if ext in EXCEL_EXTENSIONS:
            return detect_excel_password(file_path)
        if ext in WORD_EXTENSIONS or ext in POWERPOINT_EXTENSIONS:
            return detect_office_document_password(file_path)
        return None

    def update_counters(self):
        total    = len(self.file_items)
        selected = sum(1 for item in self.file_items if item.selected)
        success  = sum(1 for item in self.file_items if item.last_result == RESULT_SUCCESS)
        failed   = sum(1 for item in self.file_items if item.last_result == RESULT_FAILED)
        skipped  = sum(1 for item in self.file_items if item.last_result == RESULT_SKIPPED)

        self.total_files_var.set(str(total))
        self.selected_files_var.set(str(selected))
        self.success_count_var.set(str(success))
        self.failed_count_var.set(str(failed))
        self.skipped_count_var.set(str(skipped))

    # ------------------------------------------------------------------
    # 입력 검증
    # ------------------------------------------------------------------

    def validate_password_inputs(self):
        password = self.password_var.get()
        password_confirm = self.password_confirm_var.get()

        # [수정됨] strip() 후 빈 문자열 확인 + strip된 값이 아닌 원본 그대로 사용
        # 단, 공백만으로 이뤄진 암호는 허용하지 않는다
        if not password or not password.strip():
            messagebox.showwarning(APP_TITLE, "공백 암호는 허용되지 않습니다. 공통 암호를 입력해 주세요.")
            return None

        if password != password_confirm:
            messagebox.showwarning(APP_TITLE, "공통 암호와 공통 암호 확인 값이 일치하지 않습니다.")
            return None

        if len(password) < 5:
            messagebox.showwarning(
                APP_TITLE,
                "한글 문서 암호는 한글 버전에 따라 최소 5타 이상이 필요합니다. 5자 이상의 암호를 입력해 주세요.",
            )
            return None

        return password  # 원본 password 반환 (공백 포함 의도적 암호 허용)

    def get_selected_items(self):
        return [item for item in self.file_items if item.selected]

    @staticmethod
    def get_processing_label(extension):
        ext = extension.lower()
        if ext in PDF_EXTENSIONS:
            return "PDF"
        if ext in EXCEL_EXTENSIONS:
            return "Excel"
        if ext in WORD_EXTENSIONS:
            return "Word"
        if ext in POWERPOINT_EXTENSIONS:
            return "PowerPoint"
        if ext in HWP_EXTENSIONS:
            return "HWP"
        if ext in IMAGE_EXTENSIONS:
            return "이미지"
        return "파일"

    @staticmethod
    def get_processing_order(item):
        ext = item.extension.lower()
        if ext in PDF_EXTENSIONS:
            return (0, item.path.lower())
        if ext in EXCEL_EXTENSIONS:
            return (1, item.path.lower())
        if ext in WORD_EXTENSIONS:
            return (2, item.path.lower())
        if ext in POWERPOINT_EXTENSIONS:
            return (3, item.path.lower())
        if ext == ".hwp":
            return (4, item.path.lower())
        if ext == ".hwpx":
            return (5, item.path.lower())
        if ext in IMAGE_EXTENSIONS:
            return (6, item.path.lower())
        return (9, item.path.lower())

    @staticmethod
    def format_progress_text(index, total, extension, done=False):
        percent = int((index / total) * 100) if total else 0
        eta_text = HwpBatchPasswordApp.calculate_eta_text_static(index, total, getattr(HwpBatchPasswordApp, "_run_started_at_static", None))
        if done:
            return f"진행: {index} / {total} ({percent}%) | 예상 남은 시간: {eta_text}"
        return f"진행: {index} / {total} ({percent}%) - {HwpBatchPasswordApp.get_processing_label(extension)} | 예상 남은 시간: {eta_text}"

    @staticmethod
    def calculate_eta_text_static(completed, total, started_at):
        if not started_at or completed < 2 or total <= completed:
            return "계산 중"
        elapsed = time.time() - started_at
        if elapsed <= 0:
            return "계산 중"
        avg = elapsed / completed
        remaining = max(total - completed, 0)
        return HwpBatchPasswordApp.format_seconds(avg * remaining)

    @staticmethod
    def format_seconds(seconds):
        total_seconds = max(int(seconds), 0)
        minutes, secs = divmod(total_seconds, 60)
        hours, minutes = divmod(minutes, 60)
        if hours:
            return f"{hours}시간 {minutes}분 {secs}초"
        if minutes:
            return f"{minutes}분 {secs}초"
        return f"{secs}초"

    # ------------------------------------------------------------------
    # 실행 전 점검
    # ------------------------------------------------------------------

    def preview_run(self):
        if not self.file_items:
            messagebox.showwarning(APP_TITLE, "먼저 파일 검색을 실행해 주세요.")
            return

        selected_items = self.get_selected_items()
        if not selected_items:
            messagebox.showwarning(APP_TITLE, "선택된 파일이 없습니다.")
            return

        folder = self.get_execution_root_folder(selected_items)
        if not folder:
            return

        self.add_log("실행 전 점검을 시작합니다.")
        backup_root = self.get_backup_root(folder)
        preview_count = 0

        for item in selected_items:
            accessible, access_msg = self.check_file_access(item.path)
            backup_ready, backup_msg = self.check_backup_ready(folder, item.path, backup_root)

            item.accessible = accessible
            item.backup_ready = backup_ready
            item.status = RESULT_PENDING
            item.detail = f"점검: 접근={'가능' if accessible else '불가'}, 백업={'가능' if backup_ready else '불가'}"
            self.update_tree_row(item)

            self.add_result_log(
                path=item.path,
                result="점검",
                detail=(
                    f"확장자={item.extension}, "
                    f"접근 점검={access_msg}, "
                    f"백업 점검={backup_msg}"
                ),
            )
            preview_count += 1

        self.progress_text_var.set(f"점검 완료 ({preview_count}건)")
        self.current_file_var.set("-")
        self.add_log(f"실행 전 점검이 완료되었습니다. 점검 파일 수: {preview_count}")

    # ------------------------------------------------------------------
    # 암호 설정 실행
    # ------------------------------------------------------------------

    def run_password_setting(self):
        if self.processing:
            messagebox.showinfo(APP_TITLE, "이미 작업이 진행 중입니다.")
            return
        if self.retention_processing:
            messagebox.showinfo(APP_TITLE, "삭제 후보 점검 중에는 암호화 작업을 시작할 수 없습니다.")
            return

        selected_items = self.get_selected_items()
        if not selected_items:
            messagebox.showwarning(APP_TITLE, "처리 대상이 0개입니다. 파일을 선택해 주세요.")
            return

        folder = self.get_execution_root_folder(selected_items)
        if not folder:
            return

        password = self.validate_password_inputs()
        if password is None:
            return

        image_items = [item for item in selected_items if item.extension.lower() in IMAGE_EXTENSIONS]
        if image_items:
            selected_items = [item for item in selected_items if item.extension.lower() not in IMAGE_EXTENSIONS]
            self.add_log(
                f"이미지 파일 {len(image_items)}건은 암호화 실행 대상에서 자동 제외했습니다. "
                "이미지 정리는 삭제 후보 점검 또는 '그림파일 목록 제거'를 사용해 주세요."
            )
        if not selected_items:
            messagebox.showwarning(APP_TITLE, "선택된 항목이 이미지 파일뿐이라 암호화할 문서가 없습니다.")
            return

        xls_items = [item for item in selected_items if item.extension.lower() == ".xls"]
        if xls_items:
            if self.convert_old_office_var.get():
                self.add_log(
                    f".xls 파일 {len(xls_items)}건이 포함되어 있습니다. "
                    "최신 .xlsx 형식으로 변환 후 암호화를 시도합니다."
                )
            else:
                self.add_log(
                    f".xls 파일 {len(xls_items)}건이 포함되어 있습니다. "
                    "구형 Excel(.xls)은 자동 파일 열기 암호 설정이 제한되어 건너뛰기 또는 안내 처리될 수 있습니다."
                )
        old_office_items = [item for item in selected_items if item.extension.lower() in {".doc", ".ppt"}]
        if old_office_items:
            if self.convert_old_office_var.get():
                self.add_log(
                    f"구형 Office 파일 {len(old_office_items)}건이 포함되어 있습니다. "
                    "최신 형식으로 변환 후 암호화를 시도합니다."
                )
            else:
                self.add_log(
                    f"구형 Office 파일 {len(old_office_items)}건이 포함되어 있습니다. "
                    "DOC/PPT 구형 형식은 자동 파일 열기 암호 설정이 제한되어 건너뛰기 또는 안내 처리될 수 있습니다."
                )

        self.processing = True
        self.set_action_buttons_state()
        self.stop_requested = False
        self.run_started_at = time.time()
        HwpBatchPasswordApp._run_started_at_static = self.run_started_at
        self.progress.configure(value=0, maximum=max(len(selected_items), 1))
        self.progress_text_var.set("작업 시작 준비 중")
        self.current_file_var.set("-")
        self.current_run_paths = {item.path for item in selected_items}
        self.backup_notice_pending = False

        selected_paths = {item.path for item in selected_items}
        for item in self.file_items:
            if item.path in selected_paths:
                item.status = RESULT_PENDING
                item.detail = ""
                self.update_tree_row(item)

        failed_retry_items = [item for item in selected_items if item.last_result == RESULT_FAILED]
        if failed_retry_items and len(failed_retry_items) == len(selected_items):
            for item in failed_retry_items:
                self.retry_counts[item.path] = self.retry_counts.get(item.path, 0) + 1
            self.add_log(f"실패 파일 재실행 시작: {len(failed_retry_items)}건")

        self.add_log(f"암호 설정 작업을 시작합니다. 대상 파일 수: {len(selected_items)}")

        skip_encrypted = bool(self.skip_encrypted_var.get())
        backup_enabled = bool(self.backup_var.get())
        convert_old_office = bool(self.convert_old_office_var.get())
        pdf_items = [item for item in selected_items if item.extension.lower() in PDF_EXTENSIONS]
        excel_items = [item for item in selected_items if item.extension.lower() in EXCEL_EXTENSIONS]
        word_items = [item for item in selected_items if item.extension.lower() in WORD_EXTENSIONS]
        powerpoint_items = [item for item in selected_items if item.extension.lower() in POWERPOINT_EXTENSIONS]
        hwp_items = [item for item in selected_items if item.extension.lower() == ".hwp"]
        hwpx_items = [item for item in selected_items if item.extension.lower() == ".hwpx"]
        ordered_items = pdf_items + excel_items + word_items + powerpoint_items + hwp_items + hwpx_items
        self.hwp_notice_ack.clear()

        self.worker_thread = threading.Thread(
            target=self._worker_run,
            args=(folder, password, ordered_items, skip_encrypted, backup_enabled, convert_old_office),
            daemon=True,
        )
        self.worker_thread.start()

    def _worker_run(self, folder, password, selected_items, skip_encrypted, backup_enabled, convert_old_office):
        # HWP 파일이 포함된 경우에만 COM 인스턴스를 초기화한다.
        hwp_items = [i for i in selected_items if i.extension in HWP_EXTENSIONS]
        hwp_notice_shown = False
        hwp_focus_enabled = False

        manager = None
        hwp_start_error = None
        hwp_processed_since_restart = 0

        try:
            if hwp_items:
                try:
                    manager = HwpComManager()
                    manager.start()
                    self.queue.put(("com_ready", None))
                except Exception as exc:
                    hwp_start_error = exc
                    manager = None
                    self.queue.put((
                        "log",
                        "한글 COM 자동화 초기화에 실패하여 HWP/HWPX 파일은 실패 처리하고 PDF/Excel 처리는 계속 진행합니다.\n"
                        f"{type(exc).__name__}: {exc}",
                    ))

            total = len(selected_items)
            backup_root = self.get_backup_root(folder)

            for idx, item in enumerate(selected_items, start=1):
                if self.stop_requested:
                    self.queue.put(("log", "중지 요청이 감지되어 남은 파일을 취소 처리합니다."))
                    for cancel_idx, pending_item in enumerate(selected_items[idx - 1:], start=idx):
                        self.queue.put((
                            "file_done",
                            {
                                "path": pending_item.path,
                                "extension": pending_item.extension,
                                "result": RESULT_CANCELLED,
                                "detail": "사용자 중지 요청으로 처리되지 않음",
                                "index": cancel_idx,
                                "total": total,
                            },
                        ))
                    break

                if item.extension in HWP_EXTENSIONS and hwp_items and not hwp_notice_shown:
                    self.queue.put(("hwp_focus_on", None))
                    hwp_focus_enabled = True
                    self.queue.put(("log", "[주의] HWP 처리 중 - 마우스/키보드 조작 금지"))
                    self.queue.put(("hwp_notice", None))
                    self.hwp_notice_ack.wait(timeout=60)
                    hwp_notice_shown = True

                self.queue.put(("progress_text", self.format_progress_text(idx, total, item.extension)))
                self.queue.put(("current_file", item.path))

                original_path = item.path
                item_started_at = time.time()
                result, detail = self.process_single_file(
                    manager=manager,
                    root_folder=folder,
                    file_item=item,
                    password=password,
                    backup_root=backup_root,
                    skip_encrypted=skip_encrypted,
                    backup_enabled=backup_enabled,
                    convert_old_office=convert_old_office,
                    hwp_start_error=hwp_start_error,
                    progress_callback=lambda text: self.queue.put(("progress_text", text)),
                    current_file_callback=lambda text: self.queue.put(("current_file", text)),
                    log_callback=lambda text: self.queue.put(("log", text)),
                )

                elapsed_seconds = time.time() - item_started_at
                if item.extension in HWP_EXTENSIONS:
                    hwp_processed_since_restart += 1

                should_restart = False
                restart_message = ""
                if item.extension in HWP_EXTENSIONS:
                    if self.should_restart_hwp_com(result, detail, elapsed_seconds):
                        should_restart = True
                        if elapsed_seconds >= HWP_HEAVY_FILE_RESTART_SECONDS and result != RESULT_FAILED:
                            restart_message = "큰/느린 HWP 처리 후 안정화를 위해 한글 COM 인스턴스를 다시 초기화합니다."
                        else:
                            restart_message = "한글 COM 연결이 불안정해져 인스턴스를 다시 초기화합니다."
                    elif hwp_processed_since_restart >= HWP_COM_RESTART_EVERY:
                        should_restart = True
                        restart_message = (
                            f"HWP/HWPX {hwp_processed_since_restart}건 처리 후 안정화를 위해 "
                            "한글 COM 인스턴스를 다시 초기화합니다."
                        )

                if should_restart:
                    self.queue.put(("log", restart_message))
                    if manager:
                        try:
                            manager.quit()
                        except Exception:
                            pass
                    time.sleep(HWP_COM_COOLDOWN_SECONDS)
                    manager = None
                    hwp_start_error = None
                    hwp_processed_since_restart = 0
                    try:
                        manager = HwpComManager()
                        manager.start()
                        self.queue.put(("com_ready", None))
                    except Exception as restart_exc:
                        hwp_start_error = restart_exc
                        manager = None
                        self.queue.put((
                            "log",
                            "한글 COM 자동화 재초기화에 실패했습니다.\n"
                            f"{type(restart_exc).__name__}: {restart_exc}",
                        ))

                self.queue.put((
                    "file_done",
                    {
                        "path": original_path,
                        "extension": item.extension,
                        "new_path": item.path if item.path != original_path else None,
                        "new_extension": item.extension if item.path != original_path else None,
                        "result": result,
                        "detail": detail,
                        "index": idx,
                        "total": total,
                    },
                ))
        except Exception as exc:
            self.queue.put((
                "fatal_error",
                "작업 시작에 실패했습니다.\n"
                f"{type(exc).__name__}: {exc}\n\n"
                f"{traceback.format_exc()}",
            ))
        finally:
            if manager:
                try:
                    manager.quit()
                except Exception:
                    pass
            if hwp_focus_enabled:
                self.queue.put(("hwp_focus_off", None))
            self.queue.put(("finished", None))

    def process_single_file(
        self,
        manager,
        root_folder,
        file_item,
        password,
        backup_root,
        skip_encrypted,
        backup_enabled,
        convert_old_office,
        hwp_start_error,
        progress_callback=None,
        current_file_callback=None,
        log_callback=None,
    ):
        """
        파일 형식별로 적절한 암호 설정 루틴을 호출한다.
        - HWP/HWPX: 한글 COM 자동화
        - PDF: pikepdf
        - XLSX/XLSM/XLS: openpyxl (+ msoffcrypto-tool if available)
        """
        file_path = file_item.path
        ext = file_item.extension

        if ext in HWP_EXTENSIONS and manager is None:
            detail = "한글 COM 자동화 초기화 실패로 HWP/HWPX 처리를 진행할 수 없습니다."
            if hwp_start_error is not None:
                detail += f" {type(hwp_start_error).__name__}: {hwp_start_error}"
            return RESULT_FAILED, detail

        try:
            if self.stop_requested:
                return RESULT_CANCELLED, "사용자 중지 요청으로 처리되지 않음"

            if ext in IMAGE_EXTENSIONS:
                return RESULT_SKIPPED, "이미지 파일은 암호화 대상이 아니며 삭제 후보 점검에서만 사용됩니다."

            accessible, access_msg = self.check_file_access(file_path)
            if not accessible:
                return RESULT_FAILED, f"접근 불가: {access_msg}"

            if ext in OLD_OFFICE_EXTENSIONS:
                if not convert_old_office:
                    if ext == ".xls":
                        return RESULT_FAILED, (
                            ".xls(구형 Excel) 파일은 Python에서 파일 열기 암호 설정이 지원되지 않습니다. "
                            "Excel에서 직접 저장하거나 .xlsx로 변환 후 사용해 주세요."
                        )
                    return RESULT_FAILED, (
                        f"{ext}(구형 Office) 파일은 현재 Python에서 자동 파일 열기 암호 설정이 지원되지 않습니다. "
                        "Office에서 직접 저장하거나 최신 형식으로 변환 후 사용해 주세요."
                    )

                if progress_callback:
                    progress_callback(f"{self.get_processing_label(ext)} 구형 파일 최신 형식 변환 중")
                if current_file_callback:
                    current_file_callback(f"{file_path} -> 최신 형식 변환 중")
                if log_callback:
                    log_callback(f"[변환중] {file_path} | 구형 Office 파일을 최신 형식으로 변환 중입니다.")

                ok, new_path, msg = convert_old_office_to_modern(file_path)
                if not ok:
                    return RESULT_FAILED, msg

                new_ext = Path(new_path).suffix.lower()
                try:
                    if progress_callback:
                        progress_callback(f"{self.get_processing_label(new_ext)} 변환 완료, 암호화 중")
                    if current_file_callback:
                        current_file_callback(new_path)
                    if new_ext in EXCEL_EXTENSIONS:
                        ok, enc_msg = set_excel_password(new_path, password)
                    elif new_ext in WORD_EXTENSIONS or new_ext in POWERPOINT_EXTENSIONS:
                        ok, enc_msg = set_office_document_password(new_path, password)
                    else:
                        ok, enc_msg = False, f"변환된 파일 형식을 처리할 수 없습니다: {new_ext}"

                    if not ok:
                        try:
                            os.remove(new_path)
                        except OSError:
                            pass
                        return RESULT_FAILED, f"{msg}, 암호화 실패: {enc_msg}"

                    backup_ok, backup_msg = self.move_old_office_to_backup(root_folder, file_path)
                    if not backup_ok:
                        return RESULT_FAILED, (
                            f"{msg}, {enc_msg}. 단, 원본 구형 파일 백업 이동 실패: {backup_msg} "
                            "변환된 암호 파일은 생성되었으므로 원본 파일을 직접 확인해 주세요."
                        )

                    file_item.path = new_path
                    file_item.extension = new_ext
                    return RESULT_SUCCESS, (
                        f"{msg} 및 {enc_msg}. 원본 백업: {backup_msg}. "
                        "개인정보 보호를 위해 변환된 신형 파일 확인 후 백업 폴더의 구형 원본 삭제를 권장합니다."
                    )
                except Exception as exc:
                    try:
                        if new_path and os.path.exists(new_path):
                            os.remove(new_path)
                    except OSError:
                        pass
                    return RESULT_FAILED, f"구형 Office 변환 후 암호화 오류: {type(exc).__name__}: {exc}"

            # ── PDF ──────────────────────────────────────────────────
            if ext in PDF_EXTENSIONS:
                if skip_encrypted:
                    existing = self.detect_encrypted_file(file_path, ext)
                    if existing is True:
                        return RESULT_SKIPPED, "이미 암호가 걸린 PDF - 건너뜀"

                if backup_enabled:
                    ok, msg = self.create_backup(root_folder, file_path, backup_root)
                    if not ok:
                        return RESULT_FAILED, f"백업 실패: {msg}"

                ok, msg = set_pdf_password(file_path, password)
                if ok:
                    return RESULT_SUCCESS, msg
                if "이미 암호가 걸린 PDF" in msg:
                    return RESULT_SKIPPED, msg
                return RESULT_FAILED, msg

            # ── Excel ─────────────────────────────────────────────────
            if ext in EXCEL_EXTENSIONS:
                if skip_encrypted:
                    existing = self.detect_encrypted_file(file_path, ext)
                    if existing is True:
                        return RESULT_SKIPPED, "이미 암호가 걸린 Excel - 건너뜀"

                if backup_enabled:
                    ok, msg = self.create_backup(root_folder, file_path, backup_root)
                    if not ok:
                        return RESULT_FAILED, f"백업 실패: {msg}"

                ok, msg = set_excel_password(file_path, password)
                if ok:
                    return RESULT_SUCCESS, msg
                if "이미 암호" in msg or "열기 암호 필요" in msg:
                    return RESULT_SKIPPED, msg
                return RESULT_FAILED, msg

            # ── Word / PowerPoint ────────────────────────────────────
            if ext in WORD_EXTENSIONS or ext in POWERPOINT_EXTENSIONS:
                doc_label = "Word" if ext in WORD_EXTENSIONS else "PowerPoint"
                if skip_encrypted:
                    existing = self.detect_encrypted_file(file_path, ext)
                    if existing is True:
                        return RESULT_SKIPPED, f"이미 암호가 걸린 {doc_label} - 건너뜀"

                if backup_enabled:
                    ok, msg = self.create_backup(root_folder, file_path, backup_root)
                    if not ok:
                        return RESULT_FAILED, f"백업 실패: {msg}"

                ok, msg = set_office_document_password(file_path, password)
                if ok:
                    return RESULT_SUCCESS, msg
                if "이미 암호" in msg or "열기 암호 필요" in msg:
                    return RESULT_SKIPPED, msg
                return RESULT_FAILED, msg

            # ── HWP / HWPX ───────────────────────────────────────────
            existing_password = self.detect_existing_password(file_path)
            if skip_encrypted and existing_password is True:
                return RESULT_SKIPPED, "파일 헤더에서 기존 문서 암호가 감지되어 건너뜀"
            if skip_encrypted and existing_password is None and ext == ".hwp":
                return RESULT_FAILED, (
                    "기존 암호 여부를 판독하지 못해 안전을 위해 처리하지 않았습니다. "
                    "파일이 표준 HWP 5.x 형식이 아니거나 손상/특수 형식일 수 있습니다."
                )

            if backup_enabled:
                backup_ok, backup_msg = self.create_backup(root_folder, file_path, backup_root)
                if not backup_ok:
                    return RESULT_FAILED, f"백업 실패: {backup_msg}"

            try:
                manager.open_document(file_path)
            except Exception as exc:
                if skip_encrypted and isinstance(exc, HwpOpenPasswordRequiredError):
                    return RESULT_SKIPPED, f"이미 암호가 있거나 암호 입력이 필요한 파일로 판단되어 건너뜀: {exc}"
                return RESULT_FAILED, f"문서 열기 실패: {exc}"

            try:
                manager.set_password_and_save(file_path, password)
            finally:
                manager.close_document()

            # 암호 플래그 재검증
            password_flag = self.detect_existing_password(file_path)
            if password_flag is True:
                return RESULT_SUCCESS, "암호 설정 및 저장 완료, 파일 헤더 암호 플래그 확인됨"
            if ext == ".hwpx":
                return RESULT_SUCCESS, (
                    "암호 설정 및 저장 완료 (HWPX는 자동 검증 신뢰도가 낮아 암호 플래그 확인이 제한됩니다. "
                    "한글에서 직접 다시 열어 암호 확인을 권장합니다)"
                )
            if password_flag is False:
                return RESULT_FAILED, (
                    "암호 설정 액션과 저장은 실행됐지만 파일 헤더에서 암호 플래그가 확인되지 않았습니다. "
                    "현재 한글 버전의 문서 암호 COM 액션명/파라미터명이 코드와 다를 가능성이 큽니다."
                )

            return RESULT_FAILED, (
                "암호 설정 액션과 저장은 실행됐지만 이 파일 형식은 자동 검증이 제한됩니다."
            )

        except Exception as exc:
            return RESULT_FAILED, (
                f"예외 발생: {type(exc).__name__}: {exc}\n"
                f"상세 추적:\n{traceback.format_exc()}"
            )

    @staticmethod
    def should_restart_hwp_com(result, detail, elapsed_seconds=None):
        if elapsed_seconds is not None and elapsed_seconds >= HWP_HEAVY_FILE_RESTART_SECONDS:
            return True
        if result != RESULT_FAILED:
            return False
        lowered = str(detail).lower()
        restart_markers = [
            "rpc 서버를 사용할 수 없습니다",
            "rpc server is unavailable",
            "서버에서 예외 오류가 발생했습니다",
            "server threw an exception",
            "hwpframe.hwpobject.run",
            "hwpframe.hwpobject.hparameterset",
            "hwpframe.hwpobject.createaction",
            "매개 변수의 개수가 잘못되었습니다",
            "wrong number of arguments",
        ]
        return any(marker in lowered for marker in restart_markers)

    # ------------------------------------------------------------------
    # 유틸리티 — HwpBatchPasswordApp에 올바르게 배치
    # (기존 코드에서 HwpComManager에 잘못 배치되어 있던 메서드들을 이동)
    # ------------------------------------------------------------------

    def get_backup_root(self, root_folder):
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        return os.path.join(root_folder, f"_backup_hwp_password_{timestamp}")

    def check_backup_ready(self, root_folder, file_path, backup_root):
        if not self.backup_var.get():
            return True, "백업 옵션 미사용"

        try:
            rel_path = os.path.relpath(file_path, root_folder)
            target_path = os.path.join(backup_root, rel_path)
            target_dir = os.path.dirname(target_path)
            os.makedirs(target_dir, exist_ok=True)
            return True, f"백업 경로 준비 가능: {target_path}"
        except Exception as exc:
            return False, f"백업 경로 준비 실패: {type(exc).__name__}: {exc}"

    def create_backup(self, root_folder, file_path, backup_root):
        try:
            rel_path = os.path.relpath(file_path, root_folder)
            backup_path = os.path.join(backup_root, rel_path)
            os.makedirs(os.path.dirname(backup_path), exist_ok=True)
            shutil.copy2(file_path, backup_path)
            self.backup_notice_pending = True
            return True, f"백업 완료: {backup_path}"
        except Exception as exc:
            return False, f"{type(exc).__name__}: {exc}"

    def move_old_office_to_backup(self, root_folder, file_path):
        try:
            rel_path = os.path.relpath(file_path, root_folder)
            if rel_path.startswith("..") or os.path.isabs(rel_path):
                rel_path = os.path.basename(file_path)
            backup_path = os.path.join(root_folder, OLD_OFFICE_BACKUP_DIR, rel_path)
            backup_path = make_unique_path(backup_path)
            os.makedirs(os.path.dirname(backup_path), exist_ok=True)
            shutil.move(file_path, backup_path)
            return True, backup_path
        except Exception as exc:
            return False, f"{type(exc).__name__}: {exc}"

    def add_result_log(self, path, result, detail, result_code=None):
        record = {
            "처리시각": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "파일경로": path,
            "확장자": Path(path).suffix.lower(),
            "결과": result,
            "RESULT_CODE": result_code or self.get_result_code(result, detail),
            "상세메시지": detail,
        }
        self.logs.append(record)
        self.add_log(f"[{result}] {path} | {detail}")

    @staticmethod
    def get_result_code(result, detail):
        if result == RESULT_SUCCESS:
            if "msoffcrypto-tool 미설치" in detail:
                return RESULT_CODE_NO_MSOFFCRYPTO
            return RESULT_CODE_SUCCESS
        if result == RESULT_CANCELLED:
            return RESULT_CODE_CANCELLED
        if result == RESULT_SKIPPED:
            if "이미 암호" in detail or "기존 문서 암호" in detail:
                return RESULT_CODE_ALREADY_ENCRYPTED
            return RESULT_CODE_SKIPPED
        if result == RESULT_FAILED:
            return RESULT_CODE_FAIL
        return result.upper().replace(" ", "_")

    @staticmethod
    def check_file_access(file_path):
        try:
            if not os.path.exists(file_path):
                return False, "파일이 존재하지 않습니다."
            if not os.path.isfile(file_path):
                return False, "일반 파일이 아닙니다."
            with open(file_path, "rb+"):
                pass
            return True, "읽기/쓰기 접근 가능"
        except PermissionError as exc:
            return False, f"권한 문제: {exc}"
        except OSError as exc:
            return False, f"파일 접근 오류: {exc}"
        except Exception as exc:
            return False, f"알 수 없는 접근 오류: {type(exc).__name__}: {exc}"

    # ------------------------------------------------------------------
    # HWP 암호 감지 (OLE 파일 헤더 직접 파싱)
    # ------------------------------------------------------------------

    @staticmethod
    def detect_existing_password(file_path):
        ext = Path(file_path).suffix.lower()

        if ext == ".hwp":
            return HwpBatchPasswordApp.detect_hwp_password_flag(file_path)

        if ext == ".hwpx":
            return HwpBatchPasswordApp.detect_hwpx_password_flag(file_path)

        return None

    @staticmethod
    def detect_hwp_password_flag(file_path):
        """
        HWP 5.x OLE 파일의 FileHeader 스트림 속성 비트를 확인한다.
        FileHeader의 36~39바이트가 속성 플래그이며, 0x02가 암호 설정 여부.
        """
        header = HwpBatchPasswordApp.read_hwp_file_header_stream(file_path)
        if header is not None:
            if len(header) < 40:
                return None
            signature = header[:32]
            if b"HWP Document File" not in signature:
                return None
            flags = int.from_bytes(header[36:40], byteorder="little", signed=False)
            return bool(flags & 0x02)

        if pythoncom is None:
            return None

        try:
            try:
                pythoncom.CoInitialize()
            except Exception:
                pass
            mode = getattr(pythoncom, "STGM_READ", 0) | getattr(pythoncom, "STGM_SHARE_DENY_NONE", 0x40)
            storage = pythoncom.StgOpenStorage(file_path, None, mode, None, 0)
            stream = storage.OpenStream("FileHeader", None, mode, 0)
            header = stream.Read(256)

            if len(header) < 40:
                return None

            signature = header[:32]
            if b"HWP Document File" not in signature:
                return None

            flags = int.from_bytes(header[36:40], byteorder="little", signed=False)
            return bool(flags & 0x02)
        except Exception:
            return None

    @staticmethod
    def detect_hwpx_password_flag(file_path):
        """
        HWPX는 ZIP 기반이라 표준 ZIP 암호화 플래그와 패키지 내부 암호화 흔적을 함께 확인한다.
        한글 버전에 따라 ZIP 플래그 없이 XML/메타 파일에 보안 정보가 남는 경우가 있어
        제한된 범위만 읽어 재암호화 방지용으로 보수적으로 판정한다.
        """
        try:
            if not zipfile.is_zipfile(file_path):
                return None

            with zipfile.ZipFile(file_path, "r") as zf:
                infos = zf.infolist()
                if not infos:
                    return None
                if any(info.flag_bits & 0x1 for info in infos):
                    return True

                strong_name_markers = (
                    "encryption", "encrypted", "encrypt", "password",
                    "security", "drm", "cipher", "certificate",
                )
                for info in infos:
                    name = info.filename.lower()
                    if any(marker in name for marker in strong_name_markers):
                        return True

                scanned_bytes = 0
                max_total_scan = 5 * 1024 * 1024
                max_entry_scan = 1024 * 1024
                marker_pattern = re.compile(
                    rb"(password|encrypt|encrypted|encryption|drm|cipher|salt)",
                    re.IGNORECASE,
                )

                for info in infos:
                    if info.is_dir() or info.file_size <= 0:
                        continue
                    name = info.filename.lower()
                    if not (
                        name.endswith((".xml", ".rels", ".txt", ".dat"))
                        or "settings" in name
                        or "header" in name
                        or "manifest" in name
                    ):
                        continue
                    if scanned_bytes >= max_total_scan:
                        break

                    try:
                        with zf.open(info) as stream:
                            data = stream.read(min(info.file_size, max_entry_scan))
                    except RuntimeError:
                        return True
                    except Exception:
                        continue

                    scanned_bytes += len(data)
                    if marker_pattern.search(data):
                        return True

                return False
        except Exception:
            return None

    @staticmethod
    def read_hwp_file_header_stream(file_path):
        """
        pywin32 OLE 열기가 실패하는 환경을 대비한 최소 CFB(OLE Compound File) 판독기.
        HWP의 FileHeader 스트림만 읽기 위한 용도.
        """
        FREESECT = 0xFFFFFFFF
        ENDOFCHAIN = 0xFFFFFFFE
        FATSECT = 0xFFFFFFFD

        try:
            with open(file_path, "rb") as f:
                data = f.read()

            if len(data) < 512 or data[:8] != b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
                return None

            sector_size = 1 << struct.unpack_from("<H", data, 30)[0]
            mini_sector_size = 1 << struct.unpack_from("<H", data, 32)[0]
            first_dir_sector = struct.unpack_from("<I", data, 48)[0]
            mini_cutoff = struct.unpack_from("<I", data, 56)[0]
            first_mini_fat_sector = struct.unpack_from("<I", data, 60)[0]
            num_mini_fat_sectors = struct.unpack_from("<I", data, 64)[0]
            first_difat_sector = struct.unpack_from("<I", data, 68)[0]
            num_difat_sectors = struct.unpack_from("<I", data, 72)[0]

            def sector_bytes(sector_id):
                if sector_id in (FREESECT, ENDOFCHAIN, FATSECT):
                    return b""
                start = (sector_id + 1) * sector_size
                end = start + sector_size
                return data[start:end]

            difat = list(struct.unpack_from("<109I", data, 76))
            next_difat = first_difat_sector
            for _ in range(num_difat_sectors):
                if next_difat in (FREESECT, ENDOFCHAIN):
                    break
                sec = sector_bytes(next_difat)
                if len(sec) < sector_size:
                    break
                entries = struct.unpack("<" + "I" * (sector_size // 4), sec)
                difat.extend(entries[:-1])
                next_difat = entries[-1]

            fat = []
            for fat_sector in difat:
                if fat_sector in (FREESECT, ENDOFCHAIN, FATSECT):
                    continue
                sec = sector_bytes(fat_sector)
                if len(sec) == sector_size:
                    fat.extend(struct.unpack("<" + "I" * (sector_size // 4), sec))

            def read_chain(start_sector):
                chunks = []
                sector = start_sector
                visited = set()
                while sector not in (FREESECT, ENDOFCHAIN) and sector < len(fat) and sector not in visited:
                    visited.add(sector)
                    chunks.append(sector_bytes(sector))
                    sector = fat[sector]
                return b"".join(chunks)

            directory_data = read_chain(first_dir_sector)
            entries = []
            root_entry = None
            file_header_entry = None

            for offset in range(0, len(directory_data), 128):
                entry = directory_data[offset:offset + 128]
                if len(entry) < 128:
                    continue

                name_len = struct.unpack_from("<H", entry, 64)[0]
                if name_len < 2:
                    continue

                name_raw = entry[:name_len - 2]
                try:
                    name = name_raw.decode("utf-16le")
                except UnicodeDecodeError:
                    continue

                entry_type = entry[66]
                start_sector = struct.unpack_from("<I", entry, 116)[0]
                size = struct.unpack_from("<Q", entry, 120)[0]
                parsed = {"name": name, "type": entry_type, "start": start_sector, "size": size}
                entries.append(parsed)

                if entry_type == 5:
                    root_entry = parsed
                elif entry_type == 2 and name == "FileHeader":
                    file_header_entry = parsed

            if not file_header_entry:
                return None

            if file_header_entry["size"] < mini_cutoff and root_entry:
                root_stream = read_chain(root_entry["start"])
                mini_fat = []
                sector = first_mini_fat_sector
                visited = set()
                for _ in range(num_mini_fat_sectors):
                    if sector in (FREESECT, ENDOFCHAIN) or sector >= len(fat) or sector in visited:
                        break
                    visited.add(sector)
                    sec = sector_bytes(sector)
                    if len(sec) == sector_size:
                        mini_fat.extend(struct.unpack("<" + "I" * (sector_size // 4), sec))
                    sector = fat[sector]

                chunks = []
                mini_sector = file_header_entry["start"]
                visited.clear()
                while (
                    mini_sector not in (FREESECT, ENDOFCHAIN)
                    and mini_sector < len(mini_fat)
                    and mini_sector not in visited
                ):
                    visited.add(mini_sector)
                    start = mini_sector * mini_sector_size
                    chunks.append(root_stream[start:start + mini_sector_size])
                    mini_sector = mini_fat[mini_sector]
                return b"".join(chunks)[:file_header_entry["size"]]

            return read_chain(file_header_entry["start"])[:file_header_entry["size"]]
        except Exception:
            return None

    # ------------------------------------------------------------------
    # Queue 폴링 (백그라운드 스레드 → UI 업데이트)
    # ------------------------------------------------------------------

    def _poll_queue(self):
        try:
            while True:
                action, payload = self.queue.get_nowait()
                if action == "progress_text":
                    self.progress_text_var.set(payload)
                elif action == "hwp_focus_on":
                    try:
                        self.root.grab_release()
                    except Exception:
                        pass
                    try:
                        self.root.attributes("-topmost", False)
                        self.root.lower()
                    except Exception:
                        pass
                elif action == "hwp_notice":
                    self._show_hwp_permission_dialog()
                    try:
                        self.root.attributes("-topmost", False)
                        self.root.lower()
                    except Exception:
                        pass
                    self.hwp_notice_ack.set()
                elif action == "hwp_focus_off":
                    try:
                        self.root.grab_release()
                    except Exception:
                        pass
                    finally:
                        self.root.attributes("-topmost", False)
                    try:
                        self.root.lift()
                    except Exception:
                        pass
                elif action == "current_file":
                    self.current_file_var.set(payload)
                elif action == "file_done":
                    self._handle_file_done(payload)
                elif action == "fatal_error":
                    try:
                        self.root.grab_release()
                    except Exception:
                        pass
                    finally:
                        self.root.attributes("-topmost", False)
                    self.progress_text_var.set("실패")
                    self.add_log(payload)
                    messagebox.showerror(APP_TITLE, payload)
                elif action == "com_ready":
                    self.add_log("한글 COM 자동화 초기화 완료")
                elif action == "log":
                    self.add_log(payload)
                elif action == "retention_done":
                    self.retention_processing = False
                    self.retention_results = payload.get("results", [])
                    self.set_action_buttons_state()
                    for error in payload.get("errors", []):
                        self.add_log(f"[점검오류] {error}")
                    self.progress_text_var.set("삭제 후보 점검 완료")
                    self.current_file_var.set("-")
                    self.add_log(f"삭제 후보 점검 완료: {len(self.retention_results)}건")
                    self.show_retention_results_window(self.retention_results)
                elif action == "finished":
                    self.processing = False
                    self.set_action_buttons_state()
                    try:
                        self.root.grab_release()
                    except Exception:
                        pass
                    finally:
                        self.root.attributes("-topmost", False)
                    self.current_file_var.set("-")
                    if self.progress["maximum"] > 0 and float(self.progress["value"]) >= float(self.progress["maximum"]):
                        self.progress_text_var.set("작업 완료")
                    elif self.progress_text_var.get() != "실패":
                        self.progress_text_var.set("작업 종료")
                    self.log_processing_summary()
                    self.show_backup_privacy_notice_if_needed()
        except queue.Empty:
            pass
        finally:
            self.root.after(150, self._poll_queue)

    def _handle_file_done(self, payload):
        path   = payload["path"]
        result = payload["result"]
        detail = payload["detail"]
        index  = payload["index"]
        total  = payload["total"]
        new_path = payload.get("new_path")
        new_extension = payload.get("new_extension")
        payload_extension = payload.get("extension") or Path(path).suffix.lower()

        item = self.file_map.get(path)
        if item:
            if new_path:
                self.file_map.pop(path, None)
                item.path = new_path
                item.extension = new_extension or Path(new_path).suffix.lower()
                self.file_map[item.path] = item
                if path in self.current_run_paths:
                    self.current_run_paths.discard(path)
                    self.current_run_paths.add(item.path)
            item.status = result
            item.detail = detail
            item.last_result = result
            item.timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            self.update_tree_row(item)

        if OLD_OFFICE_BACKUP_DIR in detail:
            self.backup_notice_pending = True

        if result == RESULT_FAILED:
            self.failed_paths.add(item.path if item else path)
        else:
            self.failed_paths.discard(path)
            if item:
                self.failed_paths.discard(item.path)

        self.progress.configure(value=index)
        progress_extension = item.extension if item else payload_extension
        self.progress_text_var.set(self.format_progress_text(index, total, progress_extension, done=True))
        self.add_result_log(item.path if item else path, result, detail)
        self.update_counters()

    # ------------------------------------------------------------------
    # 로그 저장
    # ------------------------------------------------------------------

    def save_log(self):
        if not self.logs:
            messagebox.showwarning(APP_TITLE, "저장할 로그가 없습니다.")
            return

        file_path = filedialog.asksaveasfilename(
            title="로그 저장",
            defaultextension=".csv",
            filetypes=[("CSV 파일", "*.csv"), ("Excel 파일", "*.xlsx"), ("텍스트 파일", "*.txt"), ("모든 파일", "*.*")],
        )
        if not file_path:
            return

        try:
            lower = file_path.lower()
            if lower.endswith(".txt"):
                self._save_log_txt(file_path)
            elif lower.endswith(".xlsx"):
                self._save_log_xlsx(file_path)
            else:
                self._save_log_csv(file_path)
            messagebox.showinfo(APP_TITLE, f"로그를 저장했습니다.\n{file_path}")
        except Exception as exc:
            messagebox.showerror(APP_TITLE, f"로그 저장 중 오류가 발생했습니다.\n{type(exc).__name__}: {exc}")

    def _save_log_csv(self, file_path):
        with open(file_path, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.DictWriter(
                f,
                fieldnames=["처리시각", "파일경로", "확장자", "결과", "RESULT_CODE", "상세메시지"],
            )
            writer.writeheader()
            writer.writerows(self.logs)

    def _save_log_txt(self, file_path):
        with open(file_path, "w", encoding="utf-8-sig") as f:
            for row in self.logs:
                f.write(
                    f"처리시각: {row['처리시각']}\n"
                    f"파일경로: {row['파일경로']}\n"
                    f"확장자: {row['확장자']}\n"
                    f"결과: {row['결과']}\n"
                    f"RESULT_CODE: {row.get('RESULT_CODE', '')}\n"
                    f"상세메시지: {row['상세메시지']}\n"
                    f"{'-' * 80}\n"
                )

    def _save_log_xlsx(self, file_path):
        if openpyxl is None:
            raise RuntimeError("Excel 로그 저장에는 openpyxl 라이브러리가 필요합니다.")

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "로그"
        headers = ["처리시각", "파일경로", "확장자", "결과", "RESULT_CODE", "상세메시지"]
        ws.append(headers)
        for row in self.logs:
            ws.append([row.get(header, "") for header in headers])
        wb.save(file_path)

    def log_processing_summary(self):
        if not self.current_run_paths:
            return

        summary = {}
        total = 0
        for item in self.file_items:
            if item.path not in self.current_run_paths:
                continue
            category = self.get_summary_category(item.extension)
            bucket = summary.setdefault(
                category,
                {"total": 0, RESULT_SUCCESS: 0, RESULT_FAILED: 0, RESULT_SKIPPED: 0, RESULT_CANCELLED: 0},
            )
            bucket["total"] += 1
            total += 1
            if item.last_result in bucket:
                bucket[item.last_result] += 1

        self.add_log(f"총 {total}개 처리 완료")
        for category in ("Excel", "Word", "PowerPoint", "PDF", "HWP/HWPX", "이미지", "기타"):
            bucket = summary.get(category)
            if not bucket:
                continue
            self.add_log(
                f"- {category}: {bucket['total']}개 "
                f"(성공 {bucket[RESULT_SUCCESS]} / 실패 {bucket[RESULT_FAILED]} / "
                f"스킵 {bucket[RESULT_SKIPPED]} / 취소 {bucket[RESULT_CANCELLED]})"
            )
        if any(OLD_OFFICE_BACKUP_DIR in row.get("상세메시지", "") for row in self.logs):
            self.add_log(
                f"[개인정보 보호 안내] 구형 Office 원본은 {OLD_OFFICE_BACKUP_DIR} 폴더에 보관되어 있습니다. "
                "변환된 신형 파일이 정상이고 암호가 적용된 것을 확인한 뒤 백업 폴더 삭제를 권장합니다."
            )
        self.current_run_paths = set()

    def show_backup_privacy_notice_if_needed(self):
        if not self.backup_notice_pending:
            return
        self.backup_notice_pending = False
        messagebox.showwarning(
            APP_TITLE,
            "백업 파일 개인정보 보호 안내\n\n"
            "이번 작업에서 백업 파일이 생성되었을 수 있습니다.\n"
            "백업 폴더에는 암호화 전 원본 파일 또는 구형 Office 원본 파일이 남아 있을 수 있습니다.\n\n"
            "변환된 신형 파일과 암호 적용 상태를 확인한 뒤, 개인정보 보호를 위해 불필요한 백업 폴더를 삭제해 주세요.",
        )

    @staticmethod
    def get_summary_category(extension):
        ext = extension.lower()
        if ext in EXCEL_EXTENSIONS:
            return "Excel"
        if ext in WORD_EXTENSIONS:
            return "Word"
        if ext in POWERPOINT_EXTENSIONS:
            return "PowerPoint"
        if ext in PDF_EXTENSIONS:
            return "PDF"
        if ext in HWP_EXTENSIONS:
            return "HWP/HWPX"
        if ext in IMAGE_EXTENSIONS:
            return "이미지"
        return "기타"

    # —————————————————————————

    # 진입점

    # —————————————————————————

def main():
    root = TkinterDnD.Tk() if TkinterDnD is not None else tk.Tk()

    try:
        style = ttk.Style(root)
        if "vista" in style.theme_names():
            style.theme_use("vista")
    except Exception:
        pass

    app = HwpBatchPasswordApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
