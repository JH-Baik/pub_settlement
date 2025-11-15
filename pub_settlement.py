from __future__ import annotations
import pandas as pd
from utils import pick_engine, safe_int, safe_float, detect_bookstore

class BookstoreSettlementProcessor:
    """서점별 정산서 처리/통합"""

    def __init__(self):
        self.unified_data: list[dict] = []

    # ----------------- YES24 ----------------- #
    def process_yes24(self, filepath: str) -> tuple[int, str | None]:
        try:
            engine = pick_engine(filepath)
            if engine is None:
                return 0, "지원하지 않는 파일 형식입니다(.xls/.xlsx)"

            df = pd.read_excel(filepath, engine=engine, dtype={"ISBN13": str})
            required = {"상품명", "입고번호"}
            miss = [c for c in required if c not in df.columns]
            if miss:
                return 0, f"예스24 형식 누락 컬럼: {', '.join(miss)}"

            df = df[df["상품명"].notna()]
            df = df[df["입고번호"].notna()]

            for _, row in df.iterrows():
                record = {
                    "도서명": str(row.get("상품명", "")).strip(),
                    "저자명": "",
                    "ISBN": ("" if pd.isna(row.get("ISBN13")) else str(row.get("ISBN13")).strip()),
                    "서점명": "예스24",
                    "입고수량": safe_int(row.get("입고수량", 0)),
                    "단가": safe_float(row.get("원가", 0)),
                    "정산액": safe_float(row.get("조정입고금액", 0)),
                    "정가": safe_int(row.get("정가", 0)),
                    "입고율": safe_int(row.get("입고율", 0)),
                }
                self.unified_data.append(record)
            return len(df), None
        except Exception as e:
            return 0, f"예스24 처리 오류: {e}"

    # ----------------- 교보문고 ----------------- #
    def process_kyobo(self, filepath: str) -> tuple[int, str | None]:
        try:
            engine = pick_engine(filepath)
            if engine is None:
                return 0, "지원하지 않는 파일 형식입니다(.xls/.xlsx)"

            header_candidates = [3, 2, 0]
            last_err = None
            df = None
            for h in header_candidates:
                try:
                    tmp = pd.read_excel(filepath, engine=engine, header=h)
                    if tmp is not None and len(tmp.columns) > 3:
                        df = tmp
                        break
                except Exception as he:
                    last_err = he
            if df is None:
                return 0, f"교보 형식 해석 실패: {last_err}"

            # 컬럼 정규화(개행/공백 제거)
            norm_cols = [str(c).replace("\n", "").replace(" ", "") for c in df.columns]
            df.columns = norm_cols

            key_map = {
                "상품명": None,
                "상품코드": None,
                "수량": None,
                "합계금액": None,
                "정가": None,
                "공급율": None,
            }
            for c in df.columns:
                if c in key_map:
                    key_map[c] = c
                elif c in ("공급률",):
                    key_map["공급율"] = c
                elif c in ("합계",):
                    key_map["합계금액"] = c

            required = ["상품명", "수량", "합계금액"]
            miss = [k for k in required if not key_map.get(k)]
            if miss:
                return 0, f"교보 형식 컬럼 누락: {', '.join(miss)}"

            # 합계/NaN 행 제거
            df = df[pd.to_numeric(df[key_map["수량"]], errors="coerce").notna()]
            df = df[pd.to_numeric(df[key_map["합계금액"]], errors="coerce").notna()]

            cnt = 0
            for _, row in df.iterrows():
                q = safe_int(row.get(key_map["수량"], 0))
                total_amt = safe_int(row.get(key_map["합계금액"], 0))
                unit = round((total_amt / q), 2) if q > 0 else 0.0  # 소수점 2자리
                product_code = ""
                if key_map.get("상품코드"):
                    product_code = str(row.get(key_map["상품코드"]))
                    if pd.isna(product_code):
                        product_code = ""
                    else:
                        product_code = product_code.strip()
                record = {
                    "도서명": str(row.get(key_map["상품명"], "")).strip(),
                    "저자명": "",
                    "ISBN": product_code,  # 문자열로 유지
                    "서점명": "교보문고",
                    "입고수량": q,
                    "단가": unit,
                    "정산액": total_amt,
                    "정가": safe_int(row.get(key_map.get("정가", ""), 0)) if key_map.get("정가") else 0,
                    "입고율": safe_int(row.get(key_map.get("공급율", ""), 0)) if key_map.get("공급율") else 0,
                }
                self.unified_data.append(record)
                cnt += 1
            return cnt, None
        except Exception as e:
            return 0, f"교보 처리 오류: {e}"

    # ----------------- 알라딘 ----------------- #
    def process_aladin(self, filepath: str) -> tuple[int, str | None]:
        return 0, "알라딘 파일은 일자별 집계만 있어 도서별 처리 불가.\n도서별 상세 정산서가 필요합니다."

    # ----------------- 자동 라우팅 ----------------- #
    def process_file(self, filepath: str) -> tuple[int, str | None]:
        bs = detect_bookstore(filepath)
        if bs == "yes24":
            return self.process_yes24(filepath)
        if bs == "kyobo":
            return self.process_kyobo(filepath)
        if bs == "aladin":
            return self.process_aladin(filepath)
        return 0, "서점 자동 감지 실패(파일명에 '예스24'/'교보'/'알라딘' 포함 권장)."

    # ----------------- 결과/저장 ----------------- #
    def get_unified_dataframe(self) -> pd.DataFrame:
        if not self.unified_data:
            return pd.DataFrame()
        df = pd.DataFrame(self.unified_data)
        cols = ["도서명", "저자명", "ISBN", "서점명", "입고수량", "단가", "정산액", "정가", "입고율"]
        for c in cols:
            if c not in df.columns:
                df[c] = None
        return df[cols]

    def save_to_csv(self, path: str) -> bool:
        df = self.get_unified_dataframe()
        if df.empty:
            return False
        df.to_csv(path, index=False, encoding="utf-8-sig")
        return True

    def save_to_excel(self, path: str) -> bool:
        df = self.get_unified_dataframe()
        if df.empty:
            return False
        df.to_excel(path, index=False, engine="openpyxl")
        return True
