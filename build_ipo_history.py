import re
from pathlib import Path

import pandas as pd


BASE_DIR = Path(__file__).resolve().parent
DEFAULT_SOURCE = Path(
    "/Users/1112658/Library/CloudStorage/OneDrive-개인/doc/_헌규 재테크/주식분석/공모주 투자 결과.xlsx"
)
OUTPUT_FILE = BASE_DIR / "ipo_history.csv"

STANDARD_COLUMNS = [
    "연도",
    "번호",
    "구분",
    "종목",
    "업종",
    "증권사",
    "수요예측_경쟁률",
    "밴드상단_초과비율",
    "일반청약_경쟁률",
    "의무확약비율_전",
    "의무확약비율_후",
    "유통주식수_비율",
    "유통주식_비율_후",
    "시가총액_억원",
    "환매청구권",
    "특이점",
    "공모가",
    "시초가",
    "시초가_수익률",
    "첫날_종가",
    "종가_수익률",
    "평균_매도가",
    "수익률",
    "수익금",
    "세금_고려",
]


HEADER_ALIASES = {
    "번호": ["번호", "Unnamed: 0"],
    "종목": ["종목"],
    "업종": ["업종"],
    "증권사": ["증권사"],
    "수요예측_경쟁률": ["수요예측\n경쟁률", "수요예측 경쟁률"],
    "밴드상단_초과비율": ["밴드상단\n초과비율", "밴드상단 초과비율"],
    "일반청약_경쟁률": ["일반청약\n경쟁률", "일반청약 경쟁률"],
    "의무확약비율_전": ["의무확약비율(전)", "의무확약비율"],
    "의무확약비율_후": ["의무확약비율(후)"],
    "유통주식수_비율": ["유통주식수 비율"],
    "유통주식_비율_후": ["유통주식 비율(후)"],
    "시가총액_억원": ["공모가 기준 \n시가총액", "공모가 기준 시가총액", "시가총액"],
    "환매청구권": ["환매청구권"],
    "특이점": ["특이점"],
    "공모가": ["공모가"],
    "시초가": ["시초가"],
    "시초가_수익률": ["시초가 수익률"],
    "첫날_종가": ["첫날 종가"],
    "종가_수익률": ["종가 수익률"],
    "평균_매도가": ["평균 매도가"],
    "수익률": ["수익률"],
    "수익금": ["수익금"],
    "세금_고려": ["세금 고려"],
}


def compact(value):
    return re.sub(r"\s+", " ", str(value)).strip()


def normalize_header(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def find_header_row(raw):
    for idx, row in raw.iterrows():
        values = {compact(value) for value in row.tolist() if not pd.isna(value)}
        if "종목" in values and ("업종" in values or "수요예측 경쟁률" in values):
            return idx
    raise RuntimeError("헤더 행을 찾지 못했습니다.")


def build_column_map(headers):
    mapping = {}
    normalized = {normalize_header(header): i for i, header in enumerate(headers)}
    compacted = {compact(header): i for i, header in enumerate(headers)}
    for std_col, aliases in HEADER_ALIASES.items():
        for alias in aliases:
            if alias in normalized:
                mapping[std_col] = normalized[alias]
                break
            if compact(alias) in compacted:
                mapping[std_col] = compacted[compact(alias)]
                break
    return mapping


def parse_number(value):
    if pd.isna(value):
        return pd.NA
    text = compact(value)
    if not text or text.lower() == "nan" or text == "미참여":
        return pd.NA
    multiplier = 1.0
    if "조" in text:
        multiplier = 10000.0
    text = re.sub(r"[^0-9.\-]", "", text)
    if not text:
        return pd.NA
    try:
        return float(text) * multiplier
    except ValueError:
        return pd.NA


def parse_ratio(value):
    number = parse_number(value)
    if pd.isna(number):
        return pd.NA
    number = float(number)
    if number > 10:
        return number / 100
    return number


def classify_type(name):
    name = str(name).lower()
    return "스펙" if re.search(r"스팩|스펙|spec|spac|히어로", name) else "일반"


def normalize_ipo_name(name):
    text = compact(name)
    if text == "키움히어로2호":
        return "키움히어로스펙2호"
    return name


def normalize_putback(value):
    text = compact(value)
    if not text:
        return ""
    if "환매" in text:
        return "유"
    if text in {"유", "무"}:
        return text
    return ""


def parse_sheet(path, sheet_name):
    raw = pd.read_excel(path, sheet_name=sheet_name, header=None)
    header_idx = find_header_row(raw)
    headers = raw.iloc[header_idx].tolist()
    data = raw.iloc[header_idx + 1 :].copy()
    column_map = build_column_map(headers)
    rows = []
    for _, row in data.iterrows():
        item = {"연도": int(sheet_name)}
        for col in STANDARD_COLUMNS:
            item.setdefault(col, pd.NA)
        for std_col, col_idx in column_map.items():
            item[std_col] = row.iloc[col_idx]
        if pd.isna(item.get("종목")) or not compact(item.get("종목")):
            continue
        if compact(item.get("종목")) in {"평균", "종목"}:
            continue
        item["종목"] = normalize_ipo_name(item.get("종목"))
        item["번호"] = parse_number(item.get("번호"))
        item["구분"] = classify_type(item.get("종목"))
        if item["구분"] == "스펙":
            item["업종"] = "스펙"
        for col in ["수요예측_경쟁률", "일반청약_경쟁률", "시가총액_억원", "공모가", "시초가", "첫날_종가", "평균_매도가", "수익금"]:
            item[col] = parse_number(item.get(col))
        for col in [
            "밴드상단_초과비율",
            "의무확약비율_전",
            "의무확약비율_후",
            "유통주식수_비율",
            "유통주식_비율_후",
            "시초가_수익률",
            "종가_수익률",
            "수익률",
        ]:
            item[col] = parse_ratio(item.get(col))
        item["환매청구권"] = normalize_putback(item.get("환매청구권")) or normalize_putback(item.get("특이점"))
        rows.append(item)
    return rows


def build_history(source=DEFAULT_SOURCE):
    xls = pd.ExcelFile(source)
    rows = []
    for sheet in xls.sheet_names:
        if not re.fullmatch(r"\d{4}", str(sheet)):
            continue
        rows.extend(parse_sheet(source, sheet))
    df = pd.DataFrame(rows)
    for col in STANDARD_COLUMNS:
        if col not in df.columns:
            df[col] = pd.NA
    df = df[STANDARD_COLUMNS]
    df = df[df["연도"].ne(2020)].copy()
    df = df.sort_values(["연도", "번호"], ascending=[False, True])
    df["번호"] = df.groupby("연도").cumcount() + 1
    return df


def main():
    df = build_history()
    OUTPUT_FILE.parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(OUTPUT_FILE, index=False, encoding="utf-8-sig")
    print(f"저장 완료: {OUTPUT_FILE} ({len(df)}행)")


if __name__ == "__main__":
    main()
