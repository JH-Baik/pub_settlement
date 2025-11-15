from __future__ import annotations
import os, re
from datetime import datetime
import pandas as pd

VERSION = "1.0.0"
SUPPORTED_EXTS = {".xlsx", ".xls"}

_CURRENCY = re.compile(r"[₩$,]")
_WS = re.compile(r"\s+")
def _normalize_number_text(s: str) -> str:
    s = s.strip()
    s = s.replace("−", "-")              # 유니코드 마이너스
    s = _CURRENCY.sub("", s)             # 통화 기호 제거
    s = s.replace("%", "")               # 퍼센트 제거
    s = _WS.sub("", s)                   # 공백 제거
    # (1,234) → -1234
    if s.startswith("(") and s.endswith(")"):
        s = "-" + s[1:-1]
    s = s.replace(",", "")
    return s

def safe_int(v, default: int = 0) -> int:
    try:
        if pd.isna(v): return default
        if isinstance(v, str): v = _normalize_number_text(v)
        return int(float(v))
    except Exception:
        return default

def safe_float(v, default: float = 0.0) -> float:
    try:
        if pd.isna(v): return default
        if isinstance(v, str): v = _normalize_number_text(v)
        return float(v)
    except Exception:
        return default
    
def timestamp(fmt: str = "%Y%m%d_%H%M%S") -> str:
    return datetime.now().strftime(fmt)

def pick_engine(filepath: str) -> str | None:
    ext = os.path.splitext(filepath)[1].lower()
    if ext == ".xlsx":
        return "openpyxl"
    if ext == ".xls":
        # 주의: xlrd 2.x는 .xls 미지원. 1.2.0 사용 권장
        return "xlrd"
    return None

def detect_bookstore(filepath: str) -> str | None:
    name = os.path.basename(filepath).lower()
    if ("예스24" in name) or ("yes24" in name):
        return "yes24"
    if ("교보" in name) or ("kyobo" in name):
        return "kyobo"
    if ("알라딘" in name) or ("aladin" in name) or ("aladdin" in name):
        return "aladin"
    return None
