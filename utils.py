from __future__ import annotations
import os
from datetime import datetime
import pandas as pd

VERSION = "1.0.0"
SUPPORTED_EXTS = {".xlsx", ".xls"}

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

def safe_int(v, default: int = 0) -> int:
    try:
        if pd.isna(v):
            return default
        if isinstance(v, str):
            v = v.replace(",", "").strip()
            if v == "":
                return default
        return int(float(v))
    except Exception:
        return default

def safe_float(v, default: float = 0.0) -> float:
    try:
        if pd.isna(v):
            return default
        if isinstance(v, str):
            v = v.replace(",", "").strip()
            if v == "":
                return default
        return float(v)
    except Exception:
        return default

def detect_bookstore(filepath: str) -> str | None:
    name = os.path.basename(filepath).lower()
    if ("예스24" in name) or ("yes24" in name):
        return "yes24"
    if ("교보" in name) or ("kyobo" in name):
        return "kyobo"
    if ("알라딘" in name) or ("aladin" in name) or ("aladdin" in name):
        return "aladin"
    return None
