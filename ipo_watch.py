import argparse
import io
import os
import re
import time
import zipfile
from datetime import date, datetime, timedelta
from html.parser import HTMLParser
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode
from urllib.request import Request, urlopen
from xml.etree import ElementTree as ET

import pandas as pd


BASE_DIR = Path(__file__).resolve().parent
PROJECT_DIR = BASE_DIR.parent
OUTPUT_CSV = BASE_DIR / "ipo_watch_results.csv"
OUTPUT_XLSX = BASE_DIR / "ipo_watch_results.xlsx"
RAW_OUTPUT_CSV = BASE_DIR / "ipo_watch_raw.csv"
CORP_CODE_CACHE = BASE_DIR / "corp_codes.csv"

NAVER_IPO_URL = "https://finance.naver.com/sise/ipo.naver"
DART_API_BASE = "https://opendart.fss.or.kr/api"
DART_VIEWER_BASE = "https://dart.fss.or.kr/dsaf001/main.do?rcpNo="
INTERNAL_COLUMNS = [
    "작업일",
    "회사",
    "업종",
    "개인청약_시작일",
    "상장일",
    "네이버_일정열",
    "DART회사명",
    "공시일",
    "공시명",
    "접수번호",
    "DART_URL",
    "공모가",
    "증권사",
    "수요예측_경쟁률",
    "일반청약_경쟁률",
    "의무확약비율_전",
    "의무확약비율_후",
    "의무보유확약_계",
    "의무보유확약_미확약",
    "의무보유확약_후_계",
    "의무보유확약_후_미확약",
    "시가총액_억원",
    "유통가능주식수",
    "유통가능비율",
    "유통가능비율_후",
    "시초가",
    "첫날_종가",
    "시초가_수익률",
    "종가_수익률",
    "환매청구권",
    "공모가_근거",
    "수요예측_근거",
    "의무확약_근거",
    "발행실적_접수번호",
    "발행실적_URL",
    "발행실적_근거",
    "시가총액_근거",
    "유통가능_근거",
    "가격_근거",
    "네이버_원자료",
    "오류",
]
EXCEL_COLUMNS = [
    "번호",
    "구분",
    "종목",
    "업종",
    "증권사",
    "상장일",
    "수요예측\n경쟁률",
    "일반청약\n경쟁률",
    "의무확약비율(전)",
    "의무확약비율(후)",
    "유통주식수 비율",
    "유통주식 비율(후)",
    "시가총액",
    "환매청구권",
    "공모가",
    "시초가",
    "시초가 수익률",
    "첫날 종가",
    "종가 수익률",
    "평균 매도가",
    "수익률",
    "수익금",
    "세금 고려",
]


class TextExtractor(HTMLParser):
    def __init__(self):
        super().__init__()
        self.parts = []

    def handle_data(self, data):
        text = str(data).strip()
        if text:
            self.parts.append(text)

    def handle_starttag(self, tag, attrs):
        if tag.lower() in ("br", "p", "tr", "td", "th", "div"):
            self.parts.append("\n")

    def handle_endtag(self, tag):
        if tag.lower() in ("p", "tr", "td", "th", "div", "table"):
            self.parts.append("\n")

    def text(self):
        text = " ".join(self.parts)
        text = re.sub(r"[ \t\r\f\v]+", " ", text)
        text = re.sub(r" *\n+ *", "\n", text)
        return text.strip()


def compact_text(value):
    return re.sub(r"\s+", " ", str(value)).strip()


def normalize_name(value):
    text = str(value)
    text = re.sub(r"\(.*?\)", "", text)
    text = re.sub(r"[\s㈜주식회사·\-_]", "", text)
    return text.lower()


def decode_bytes(raw):
    for encoding in ("cp949", "euc-kr", "utf-8"):
        try:
            return raw.decode(encoding)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="replace")


def fetch_url(url, params=None, timeout=40, retries=3, retry_delay=3):
    request_url = url if not params else f"{url}?{urlencode(params)}"
    request = Request(
        request_url,
        headers={
            "User-Agent": (
                "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126 Safari/537.36"
            )
        },
    )
    last_error = None
    for attempt in range(1, retries + 1):
        try:
            with urlopen(request, timeout=timeout) as response:
                return response.read()
        except HTTPError as error:
            last_error = error
            if 400 <= error.code < 500 and error.code not in {408, 429}:
                raise RuntimeError(f"HTTP 오류: {error.code} {error.reason} ({request_url})") from error
        except URLError as error:
            last_error = error
        if attempt < retries:
            time.sleep(retry_delay * attempt)
    if isinstance(last_error, HTTPError):
        raise RuntimeError(f"HTTP 오류: {last_error.code} {last_error.reason} ({request_url})") from last_error
    raise RuntimeError(f"네트워크 오류: {getattr(last_error, 'reason', last_error)} ({request_url})") from last_error


def load_dart_api_key():
    key = os.getenv("DART_API_KEY", "").strip()
    if key:
        return key

    # 로컬 전용 fallback. 공개 저장소에 키를 새로 복제하지 않기 위해 기존 로컬 파일에서만 읽는다.
    for path in [PROJECT_DIR / "ship" / "ship_order.py", PROJECT_DIR / "증권사" / "securities_valueup.py"]:
        if not path.exists():
            continue
        text = path.read_text(encoding="utf-8", errors="ignore")
        match = re.search(r'["\']([A-Za-z0-9]{40})["\']', text)
        if match:
            return match.group(1)
    return ""


def parse_dart_error(raw):
    try:
        root = ET.fromstring(raw.decode("utf-8", errors="ignore"))
    except ET.ParseError:
        return None
    status = root.findtext("status")
    message = root.findtext("message")
    if status in ("000", "013"):
        return None
    if status or message:
        return f"DART API 오류(status={status or 'unknown'}): {message or '알 수 없는 오류'}"
    return None


def dart_fetch(api_key, endpoint, params):
    full_params = {"crtfc_key": api_key, **params}
    raw = fetch_url(f"{DART_API_BASE}/{endpoint}", full_params)
    error = parse_dart_error(raw)
    if error:
        raise RuntimeError(error)
    return raw


def load_corp_codes(api_key, refresh=False):
    if CORP_CODE_CACHE.exists() and not refresh:
        return pd.read_csv(CORP_CODE_CACHE, dtype=str).fillna("")

    raw = dart_fetch(api_key, "corpCode.xml", {})
    if not zipfile.is_zipfile(io.BytesIO(raw)):
        raise RuntimeError(f"corpCode.xml 응답이 ZIP이 아닙니다: {raw[:200]!r}")

    with zipfile.ZipFile(io.BytesIO(raw)) as zipped:
        xml = zipped.read(zipped.namelist()[0])
    root = ET.fromstring(xml)
    rows = []
    for item in root.findall("list"):
        rows.append({
            "corp_code": item.findtext("corp_code", ""),
            "corp_name": item.findtext("corp_name", ""),
            "stock_code": item.findtext("stock_code", ""),
        })
    df = pd.DataFrame(rows)
    df.to_csv(CORP_CODE_CACHE, index=False, encoding="utf-8-sig")
    return df


def resolve_corp(corp_codes, company_name):
    wanted = normalize_name(company_name)
    if not wanted:
        return "", ""

    work = corp_codes.copy()
    work["normalized"] = work["corp_name"].map(normalize_name)

    exact = work[work["normalized"].eq(wanted)]
    if not exact.empty:
        row = exact.iloc[0]
        return row["corp_code"], row["corp_name"]

    contains = work[
        work["normalized"].map(lambda name: wanted in name or name in wanted if name else False)
    ]
    if not contains.empty:
        row = contains.iloc[0]
        return row["corp_code"], row["corp_name"]

    return "", ""


def stock_code_for_corp(corp_codes, corp_code):
    if not corp_code:
        return ""
    matched = corp_codes[corp_codes["corp_code"].astype(str).eq(str(corp_code))]
    if matched.empty:
        return ""
    return str(matched.iloc[0].get("stock_code", "")).strip()


def normalize_column_name(value):
    if isinstance(value, tuple):
        value = " ".join([str(v) for v in value if str(v) != "nan"])
    return compact_text(value)


def clean_naver_company_name(value):
    text = compact_text(value)
    text = re.sub(r"^(코스닥|코스피|유가증권|코넥스)", "", text)
    text = re.split(r"\s+공모가\s+", text, maxsplit=1)[0]
    return compact_text(text)


NAVER_INFO_LABELS = [
    "공모가",
    "업종",
    "주관사",
    "개인청약경쟁률",
    "진행상태",
    "개인청약",
    "상장일",
    "PDF",
    "팁",
    "STEP1",
]


def extract_naver_field(info, label):
    text = compact_text(info)
    labels = [re.escape(item) for item in NAVER_INFO_LABELS if item != label]
    pattern = rf"{re.escape(label)}\s+(.+?)(?=\s+(?:{'|'.join(labels)})\b|$)"
    match = re.search(pattern, text)
    if not match:
        return ""
    return compact_text(match.group(1))


def parse_short_date(value, base_year):
    value = compact_text(value)
    if not value or value == "미정":
        return None
    match = re.search(r"(\d{4})[./-](\d{1,2})[./-](\d{1,2})", value)
    if match:
        return date(int(match.group(1)), int(match.group(2)), int(match.group(3)))
    match = re.search(r"(\d{2})[./-](\d{1,2})[./-](\d{1,2})", value)
    if match:
        year = 2000 + int(match.group(1))
        return date(year, int(match.group(2)), int(match.group(3)))
    match = re.search(r"(\d{1,2})[./-](\d{1,2})", value)
    if match:
        return date(base_year, int(match.group(1)), int(match.group(2)))
    return None


def parse_naver_date_range(value, base_year):
    value = compact_text(value)
    if "~" not in value:
        parsed = parse_short_date(value, base_year)
        return parsed, parsed
    left, right = value.split("~", 1)
    start = parse_short_date(left, base_year)
    if not start:
        return None, None
    right = compact_text(right)
    if re.match(r"^\d{1,2}[./-]\d{1,2}$", right):
        end = parse_short_date(f"{start.year}.{right}", base_year)
    else:
        end = parse_short_date(right, base_year)
    return start, end


def parse_naver_competition(value):
    match = re.search(r"([\d,.]+)\s*:?\s*1?", str(value))
    return parse_money_number(match.group(1)) if match else pd.NA


def parse_naver_ipo_row(row, run_date):
    raw_company = row.get("종목", "") or row.get("회사", "")
    info = compact_text(f"{row.get('종목', '')} {row.get('투자정보', '')}")
    company = clean_naver_company_name(raw_company)
    subscription_start, subscription_end = parse_naver_date_range(extract_naver_field(info, "개인청약"), run_date.year)
    listing_date = parse_short_date(extract_naver_field(info, "상장일"), run_date.year)
    return {
        "회사": company,
        "업종": extract_naver_field(info, "업종"),
        "증권사": extract_naver_field(info, "주관사"),
        "개인청약_시작일": subscription_start.isoformat() if subscription_start else "",
        "개인청약_종료일": subscription_end.isoformat() if subscription_end else "",
        "상장일": listing_date.isoformat() if listing_date else "",
        "일반청약_경쟁률": parse_naver_competition(extract_naver_field(info, "개인청약경쟁률")),
        "네이버_원자료": " | ".join(f"{k}: {v}" for k, v in row.items() if v and v != "nan"),
    }


def flatten_naver_ipo_tables(tables):
    rows = []
    for table in tables:
        if table.empty:
            continue
        table = table.copy()
        table.columns = [normalize_column_name(col) for col in table.columns]
        table = table.dropna(how="all")
        for _, row in table.iterrows():
            item = {col: compact_text(row.get(col, "")) for col in table.columns}
            joined = " ".join(item.values())
            if not joined or "종목" in joined and len(item) <= 2:
                continue
            rows.append(item)
    return rows


def find_key_by_keywords(row, keywords):
    for key in row.keys():
        key_norm = re.sub(r"\s+", "", key)
        if any(keyword in key_norm for keyword in keywords):
            return key
    return ""


def extract_dates(text, base_year):
    text = str(text)
    found = []
    for year, month, day in re.findall(r"(\d{4})[./-]\s*(\d{1,2})[./-]\s*(\d{1,2})", text):
        found.append(date(int(year), int(month), int(day)))

    text_without_year_dates = re.sub(r"\d{4}[./-]\s*\d{1,2}[./-]\s*\d{1,2}", " ", text)
    for month, day in re.findall(r"(?<!\d)(\d{1,2})[./-]\s*(\d{1,2})(?!\d)", text_without_year_dates):
        month_i = int(month)
        day_i = int(day)
        if 1 <= month_i <= 12 and 1 <= day_i <= 31:
            found.append(date(base_year, month_i, day_i))
    return found


def find_subscription_start(row, run_date):
    personal_col = find_key_by_keywords(row, ["개인청약", "청약일정", "청약일"])
    if personal_col:
        dates = extract_dates(row.get(personal_col, ""), run_date.year)
        if dates:
            return dates[0], personal_col

    joined = " ".join(row.values())
    dates = extract_dates(joined, run_date.year)
    return (dates[0], "전체행") if dates else (None, "")


def naver_ipo_candidates(run_date):
    html = decode_bytes(fetch_url(NAVER_IPO_URL))
    tables = pd.read_html(io.StringIO(html))
    rows = flatten_naver_ipo_tables(tables)
    target_date = run_date + timedelta(days=1)
    candidates = []
    for row in rows:
        parsed = parse_naver_ipo_row(row, run_date)
        if not parsed["회사"]:
            continue
        subscription_start = datetime.strptime(parsed["개인청약_시작일"], "%Y-%m-%d").date() if parsed["개인청약_시작일"] else None
        if subscription_start != target_date:
            continue
        parsed["네이버_일정열"] = "개인청약"
        candidates.append(parsed)
    return candidates


def naver_ipo_lookup(run_date):
    html = decode_bytes(fetch_url(NAVER_IPO_URL))
    tables = pd.read_html(io.StringIO(html))
    lookup = {}
    for row in flatten_naver_ipo_tables(tables):
        parsed = parse_naver_ipo_row(row, run_date)
        if parsed["회사"]:
            lookup[normalize_name(parsed["회사"])] = parsed
    return lookup


def search_investment_prospectus(api_key, corp_code, start_date, end_date):
    rows = []
    page = 1
    while True:
        raw = dart_fetch(
            api_key,
            "list.xml",
            {
                "corp_code": corp_code,
                "bgn_de": start_date.strftime("%Y%m%d"),
                "end_de": end_date.strftime("%Y%m%d"),
                "page_no": page,
                "page_count": 100,
                "sort": "date",
                "sort_mth": "desc",
            },
        )
        root = ET.fromstring(raw)
        status = root.findtext("status")
        if status == "013":
            return []
        if status != "000":
            raise RuntimeError(f"DART list 오류(status={status}): {root.findtext('message')}")
        for item in root.findall("list"):
            report_name = item.findtext("report_nm", "")
            if "투자설명서" in report_name:
                rows.append({
                    "공시일": item.findtext("rcept_dt", ""),
                    "공시명": report_name,
                    "접수번호": item.findtext("rcept_no", ""),
                    "DART_URL": f"{DART_VIEWER_BASE}{item.findtext('rcept_no', '')}",
                })
        total_page = int(root.findtext("total_page", "1") or "1")
        if page >= total_page:
            break
        page += 1
    return rows


def search_filings_by_keyword(api_key, corp_code, start_date, end_date, keyword):
    rows = []
    page = 1
    while True:
        raw = dart_fetch(
            api_key,
            "list.xml",
            {
                "corp_code": corp_code,
                "bgn_de": start_date.strftime("%Y%m%d"),
                "end_de": end_date.strftime("%Y%m%d"),
                "page_no": page,
                "page_count": 100,
                "sort": "date",
                "sort_mth": "desc",
            },
        )
        root = ET.fromstring(raw)
        status = root.findtext("status")
        if status == "013":
            return []
        if status != "000":
            raise RuntimeError(f"DART list 오류(status={status}): {root.findtext('message')}")
        for item in root.findall("list"):
            report_name = item.findtext("report_nm", "")
            if keyword in report_name:
                rcept_no = item.findtext("rcept_no", "")
                rows.append({
                    "공시일": item.findtext("rcept_dt", ""),
                    "공시명": report_name,
                    "접수번호": rcept_no,
                    "DART_URL": f"{DART_VIEWER_BASE}{rcept_no}",
                })
        total_page = int(root.findtext("total_page", "1") or "1")
        if page >= total_page:
            break
        page += 1
    return rows


def download_document_text(api_key, rcept_no):
    raw = dart_fetch(api_key, "document.xml", {"rcept_no": rcept_no})
    texts = []
    if zipfile.is_zipfile(io.BytesIO(raw)):
        with zipfile.ZipFile(io.BytesIO(raw)) as zipped:
            for name in zipped.namelist():
                parser = TextExtractor()
                parser.feed(decode_bytes(zipped.read(name)))
                texts.append(parser.text())
    else:
        parser = TextExtractor()
        parser.feed(decode_bytes(raw))
        texts.append(parser.text())
    return "\n".join(texts)


def extract_general_subscription_competition(text):
    context = find_forward_context(text, ["청약 및 배정현황"], 6000)
    if not context:
        return pd.NA, ""
    match = re.search(r"일반투자자\s+(.{0,800}?)(?:\s+계\s+|\s+주1\)|\s+4\.\s*일반투자자|$)", context)
    if not match:
        return pd.NA, ""
    row_text = compact_text(match.group(1))
    values = numbers_from_line(row_text)
    if len(values) >= 4 and values[0]:
        competition = values[3] / values[0]
        return round(competition, 2), compact_text(f"일반투자자 {row_text}")
    return pd.NA, compact_text(f"일반투자자 {row_text}")


def extract_final_lockup_ratio(text):
    context = find_forward_context(text, ["의무보유확약기간별 배정현황"], 5000)
    if not context:
        return pd.NA, pd.NA, pd.NA, ""
    section_end = context.find("Ⅲ.")
    if section_end > 0:
        context = context[:section_end]
    uncommitted_rows = list(re.finditer(r"(미확약\s+.*?)(?:\s+(?:계|합계)\s+)", context))
    total_rows = list(re.finditer(r"((?:계|합계)\s+.*?)(?:\s+Ⅲ\.|\s+주\d+\)|$)", context))
    if not uncommitted_rows or not total_rows:
        return pd.NA, pd.NA, pd.NA, ""
    uncommitted_line = compact_text(uncommitted_rows[-1].group(1))
    total_line = compact_text(total_rows[-1].group(1))
    uncommitted = aggregate_quantity_from_row(uncommitted_line)
    total = aggregate_quantity_from_row(total_line)
    if pd.isna(total) or not total:
        return pd.NA, total, uncommitted, compact_text(f"{total_line} / {uncommitted_line}")
    ratio = (float(total) - (float(uncommitted) if pd.notna(uncommitted) else 0.0)) / float(total)
    return ratio, total, uncommitted, compact_text(f"{total_line} / {uncommitted_line}")


def parse_issuance_report_metrics(text, existing_metrics):
    general_competition, general_source = extract_general_subscription_competition(text)
    final_lockup_ratio, final_total, final_uncommitted, final_source = extract_final_lockup_ratio(text)
    metrics = {
        "일반청약_경쟁률": general_competition,
        "의무확약비율_후": final_lockup_ratio,
        "의무보유확약_후_계": final_total,
        "의무보유확약_후_미확약": final_uncommitted,
        "발행실적_근거": compact_text(f"{general_source} / {final_source}"),
    }

    float_shares = safe_float(existing_metrics.get("유통가능주식수", pd.NA))
    market_cap = safe_float(existing_metrics.get("시가총액_억원", pd.NA))
    offer_price = safe_float(existing_metrics.get("공모가", pd.NA))
    if pd.notna(float_shares) and pd.notna(market_cap) and pd.notna(offer_price) and pd.notna(final_total):
        total_shares = market_cap * 100_000_000 / offer_price
        final_locked_shares = float(final_total) - (float(final_uncommitted) if pd.notna(final_uncommitted) else 0.0)
        final_float_shares = max(0.0, float(float_shares) - final_locked_shares)
        metrics["유통가능비율_후"] = final_float_shares / total_shares if total_shares else pd.NA
    return metrics


def apply_issuance_report(api_key, corp_code, row, run_date):
    listing_date = parse_short_date(row.get("상장일", ""), run_date.year)
    search_end = min(run_date, listing_date) if listing_date else run_date
    search_start = search_end - timedelta(days=365)
    reports = search_filings_by_keyword(api_key, corp_code, search_start, search_end, "증권발행실적보고서")
    if not reports:
        return row
    report = reports[0]
    text = download_document_text(api_key, report["접수번호"])
    row.update(parse_issuance_report_metrics(text, row))
    row["발행실적_접수번호"] = report["접수번호"]
    row["발행실적_URL"] = report["DART_URL"]
    return row


def fetch_listing_price(stock_code, listing_date):
    if not stock_code or not listing_date:
        return {}
    ymd = listing_date.strftime("%Y%m%d")
    raw = fetch_url(
        "https://api.finance.naver.com/siseJson.naver",
        {
            "symbol": stock_code,
            "requestType": "1",
            "startTime": ymd,
            "endTime": ymd,
            "timeframe": "day",
        },
    )
    text = decode_bytes(raw)
    numbers = re.findall(
        r"\[(?:'|\")?(\d{8})(?:'|\")?,\s*([\d.,]+),\s*([\d.,]+),\s*([\d.,]+),\s*([\d.,]+),",
        text,
    )
    if not numbers:
        return {}
    _, open_price, high_price, low_price, close_price = numbers[-1]
    return {
        "시초가": float(str(open_price).replace(",", "")),
        "첫날_종가": float(str(close_price).replace(",", "")),
        "가격_근거": f"네이버 일별시세 {stock_code} {ymd}",
    }


def apply_listing_price(row, stock_code, run_date, force=False):
    listing_date = parse_short_date(row.get("상장일", ""), run_date.year)
    if not listing_date or run_date <= listing_date:
        return row
    if not force and has_value(row.get("시초가")) and has_value(row.get("첫날_종가")):
        return row
    prices = fetch_listing_price(stock_code, listing_date)
    if not prices:
        return row
    row.update(prices)
    offer_price = safe_float(row.get("공모가", pd.NA))
    open_price = safe_float(row.get("시초가", pd.NA))
    close_price = safe_float(row.get("첫날_종가", pd.NA))
    if pd.notna(offer_price) and offer_price:
        if pd.notna(open_price):
            row["시초가_수익률"] = open_price / offer_price - 1
        if pd.notna(close_price) and pd.notna(open_price) and open_price:
            row["종가_수익률"] = close_price / open_price - 1
    return row


def update_existing_listing_prices(df, run_date, refresh_corp_codes=False):
    if df.empty:
        return df

    api_key = load_dart_api_key()
    if not api_key:
        raise RuntimeError("DART API 키가 없습니다. DART_API_KEY 환경변수를 설정하세요.")

    corp_codes = load_corp_codes(api_key, refresh=refresh_corp_codes)
    refreshed_corp_codes = False
    rows = []

    for source_row in df.to_dict("records"):
        row = dict(source_row)
        listing_date = parse_short_date(row.get("상장일", ""), run_date.year)
        if not listing_date or run_date <= listing_date:
            rows.append(row)
            continue
        company_name = row.get("DART회사명") or row.get("회사")
        corp_code, dart_name = resolve_corp(corp_codes, company_name)
        stock_code = stock_code_for_corp(corp_codes, corp_code)
        if not stock_code and not refreshed_corp_codes:
            corp_codes = load_corp_codes(api_key, refresh=True)
            refreshed_corp_codes = True
            corp_code, dart_name = resolve_corp(corp_codes, company_name)
            stock_code = stock_code_for_corp(corp_codes, corp_code)

        if not stock_code:
            row["오류"] = compact_text(
                f"{row.get('오류', '')} / 상장일 시세 조회용 종목코드를 찾지 못했습니다."
            ).strip(" /")
            rows.append(row)
            continue

        try:
            row = apply_listing_price(row, stock_code, run_date, force=True)
            if dart_name and not has_value(row.get("DART회사명")):
                row["DART회사명"] = dart_name
        except Exception as error:
            row["오류"] = compact_text(f"{row.get('오류', '')} / 상장일 시세 조회 실패: {error}").strip(" /")
        rows.append(row)

    updated = pd.DataFrame(rows)
    for col in INTERNAL_COLUMNS:
        if col not in updated.columns:
            updated[col] = ""
    return updated[INTERNAL_COLUMNS]


def find_context(text, keywords, window=3500):
    compact = re.sub(r"\r\n?", "\n", text)
    flat = compact_text(compact)
    positions = []
    for keyword in keywords:
        idx = flat.find(keyword)
        if idx >= 0:
            positions.append(idx)
    if not positions:
        return ""
    idx = min(positions)
    return flat[max(0, idx - 400): idx + window]


def find_forward_context(text, keywords, window=3500):
    compact = re.sub(r"\r\n?", "\n", text)
    flat = compact_text(compact)
    positions = []
    for keyword in keywords:
        idx = flat.find(keyword)
        if idx >= 0:
            positions.append(idx)
    if not positions:
        return ""
    idx = min(positions)
    return flat[idx: idx + window]


def parse_money_number(value):
    if value is None:
        return pd.NA
    cleaned = re.sub(r"[^\d.]", "", str(value))
    if not cleaned:
        return pd.NA
    try:
        return float(cleaned)
    except ValueError:
        return pd.NA


def safe_float(value):
    if pd.isna(value):
        return pd.NA
    try:
        return float(str(value).replace(",", ""))
    except ValueError:
        return pd.NA


def has_value(value):
    if pd.isna(value):
        return False
    text = str(value).strip()
    return bool(text) and text.lower() not in {"nan", "none", "<na>"}


def round_to_integer(value):
    value = safe_float(value)
    if pd.isna(value):
        return pd.NA
    return int(round(float(value), 0))


def extract_first(patterns, text):
    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            value = match.group(1)
            return value, compact_text(text[max(0, match.start() - 120): match.end() + 180])
    return pd.NA, ""


def extract_last(patterns, text):
    best = None
    best_pattern = ""
    for pattern in patterns:
        matches = list(re.finditer(pattern, text, re.IGNORECASE))
        if matches:
            best = matches[-1]
            best_pattern = pattern
    if not best:
        return pd.NA, ""
    value = best.group(1)
    return value, compact_text(text[max(0, best.start() - 120): best.end() + 180])


def extract_offer_price(text):
    context = find_context(text, ["주당공모가액", "공모가액", "공모가격"], 2200)
    value, source = extract_last(
        [
            r"확정\s*공모\s*가액.{0,120}?([\d,]+)\s*원\s*으로",
            r"확정\s*공모\s*가액\s*인\s*([\d,]+)\s*원",
            r"확정공모가액.{0,120}?([\d,]+)\s*원",
        ],
        text,
    )
    if pd.notna(value):
        return parse_money_number(value), source

    value, source = extract_first(
        [
            r"주당\s*공모\s*가액.{0,120}?([\d,]+)\s*원",
            r"공모\s*가액.{0,120}?([\d,]+)\s*원",
            r"확정\s*공모\s*가.{0,120}?([\d,]+)\s*원",
            r"공모\s*가격.{0,120}?([\d,]+)\s*원",
        ],
        context or text,
    )
    return parse_money_number(value), source


def extract_demand_competition(text):
    context = find_context(text, ["수요예측 참여 내역", "수요예측참여내역"], 6500)
    source_text = context or text
    match = re.search(r"경쟁률(?:주\d+\))?\s+(.{0,700}?)(?:주\d+\)|\([나-힣]\)|수요예측 신청가격|$)", source_text)
    if match:
        values = numbers_from_line(match.group(1))
        if values:
            return values[-1], compact_text(source_text[max(0, match.start() - 120): match.end() + 180])
    value, source = extract_first([r"경쟁률.{0,500}?([\d,.]+)\s*(?:대\s*)?[:：]\s*1"], source_text)
    return parse_money_number(value), source


def numbers_from_line(line):
    values = []
    for token in re.findall(r"[-+]?\d{1,3}(?:,\d{3})+(?:\.\d+)?|[-+]?\d+(?:\.\d+)?", line):
        try:
            values.append(float(token.replace(",", "")))
        except ValueError:
            pass
    return values


def pick_largest_count(line):
    values = [value for value in numbers_from_line(line) if value >= 1]
    return max(values) if values else pd.NA


def aggregate_quantity_from_row(row_text):
    values = numbers_from_line(row_text)
    if len(values) >= 2:
        return values[-2]
    return pd.NA


def extract_lockup_ratio(text):
    context = find_forward_context(
        text,
        ["의무보유확약기간별 수요예측 참여내역", "의무보유확약 기간별 수요예측 참여내역"],
        9000,
    )
    if not context:
        context = find_forward_context(text, ["의무보유 확약"], 9000)
    if not context:
        return pd.NA, pd.NA, pd.NA, ""
    section_end = context.find("(라)")
    if section_end > 0:
        context = context[:section_end]

    uncommitted_rows = list(re.finditer(r"(미확약\s+.*?)(?:\s+(?:계|합계)\s+)", context))
    total_rows = list(re.finditer(r"((?:계|합계)\s+.*?)(?:\s+주\d+\)|\s+\([라-힣]\)|$)", context))

    if uncommitted_rows:
        uncommitted_line = compact_text(uncommitted_rows[-1].group(1))
        uncommitted = aggregate_quantity_from_row(uncommitted_line)
        source2 = uncommitted_line
    else:
        value, source2 = extract_first([r"미확약.{0,500}?([\d,]+)\s*(?:주|건|%)"], context)
        uncommitted = parse_money_number(value)

    if total_rows:
        total_line = compact_text(total_rows[-1].group(1))
        total = aggregate_quantity_from_row(total_line)
        source = total_line
    else:
        value, source = extract_first([r"(?:계|합계).{0,500}?([\d,]+)\s*(?:주|건|%)"], context)
        total = parse_money_number(value)

    ratio = pd.NA
    if pd.notna(total) and total:
        uncommitted_value = float(uncommitted) if pd.notna(uncommitted) else 0.0
        ratio = (float(total) - uncommitted_value) / float(total)
    return ratio, total, uncommitted, compact_text(f"{source} / {source2}")


def extract_market_cap(text, offer_price):
    context = find_context(text, ["시가총액", "상장예정주식수", "상장 예정주식수"], 4500)
    value, source = extract_first(
        [
            r"시가\s*총액.{0,160}?([\d,]+(?:\.\d+)?)\s*억원",
            r"시가\s*총액.{0,160}?([\d,]+(?:\.\d+)?)\s*백만원",
        ],
        context or text,
    )
    if pd.notna(value):
        amount = parse_money_number(value)
        if "백만원" in source and pd.notna(amount):
            amount = amount / 100
        return amount, source

    shares, shares_source = extract_first(
        [
            r"상장\s*예정\s*주식\s*수.{0,120}?([\d,]+)\s*주",
            r"상장예정주식수.{0,120}?([\d,]+)\s*주",
        ],
        context or text,
    )
    shares = parse_money_number(shares)
    if pd.notna(shares) and pd.notna(offer_price):
        return float(shares) * float(offer_price) / 100_000_000, f"{shares_source} * 주당공모가액"
    return pd.NA, ""


def extract_float_shares(text):
    context = find_context(text, ["유통가능", "상장직후 유통", "상장 후 유통"], 7000)
    special = re.search(
        r"의무보유\s*수량을\s*제외한\s*주식수\s*([\d,]+)\s*주\s*\(\s*([\d.]+)\s*%\s*\).{0,120}?유통가능",
        context or text,
    )
    if special:
        source = compact_text((context or text)[max(0, special.start() - 120): special.end() + 180])
        return parse_money_number(special.group(1)), parse_money_number(special.group(2)), source

    shares, source = extract_first(
        [
            r"유통가능(?:주식수|물량|주식).{0,180}?([\d,]+)\s*주",
            r"상장\s*(?:직후|후).{0,80}?유통가능.{0,180}?([\d,]+)\s*주",
        ],
        context or text,
    )
    ratio, ratio_source = extract_first(
        [
            r"유통가능.{0,220}?([\d.]+)\s*%",
            r"상장\s*(?:직후|후).{0,80}?유통가능.{0,220}?([\d.]+)\s*%",
        ],
        context or text,
    )
    source = compact_text(f"{source} {ratio_source}")
    return parse_money_number(shares), parse_money_number(ratio), source


def extract_putback_option(text):
    context = find_context(text, ["환매청구권"], 1800)
    if not context:
        return ""
    if re.search(r"환매청구권.{0,220}부여하지\s*않|환매청구권.{0,220}해당하는\s*사항이\s*존재하지\s*않", context):
        return "무"
    if re.search(r"환매청구권.{0,300}(부여|행사|권리)", context):
        return "유"
    return ""


def extract_underwriter(text):
    context = find_context(text, ["대표주관회사", "인수인", "인수회사"], 3500)
    value, _ = extract_last(
        [
            r"대표\s*주관\s*회사\s+([가-힣A-Za-z0-9&().㈜\s]+?)(?:\s+보통주|\s+기명식|\s+증권의종류|\s+인수수량|\s+주관|\s+공동|\s|$)",
            r"대표주관회사\s+([가-힣A-Za-z0-9&().㈜\s]+?)(?:\s+보통주|\s+기명식|\s+증권의종류|\s+인수수량|\s|$)",
        ],
        context or text,
    )
    if pd.isna(value):
        return ""
    value = compact_text(value)
    value = re.sub(r"㈜|주식회사|\(주\)", "", value)
    value = re.split(r"\s{2,}|[,/]", value)[0]
    return compact_text(value)


def parse_prospectus_metrics(text):
    offer_price, offer_source = extract_offer_price(text)
    competition, competition_source = extract_demand_competition(text)
    lockup_ratio, lockup_total, lockup_uncommitted, lockup_source = extract_lockup_ratio(text)
    market_cap, market_cap_source = extract_market_cap(text, offer_price)
    float_shares, float_ratio, float_source = extract_float_shares(text)
    underwriter = extract_underwriter(text)
    return {
        "공모가": offer_price,
        "증권사": underwriter,
        "수요예측_경쟁률": competition,
        "의무확약비율_전": lockup_ratio,
        "의무보유확약_계": lockup_total,
        "의무보유확약_미확약": lockup_uncommitted,
        "시가총액_억원": market_cap,
        "유통가능주식수": float_shares,
        "유통가능비율": float_ratio,
        "환매청구권": extract_putback_option(text),
        "공모가_근거": offer_source,
        "수요예측_근거": competition_source,
        "의무확약_근거": lockup_source,
        "시가총액_근거": market_cap_source,
        "유통가능_근거": float_source,
    }


def collect_ipo_watch(run_date, refresh_corp_codes=False):
    api_key = load_dart_api_key()
    if not api_key:
        raise RuntimeError("DART API 키가 없습니다. DART_API_KEY 환경변수를 설정하세요.")

    candidates = naver_ipo_candidates(run_date)
    corp_codes = load_corp_codes(api_key, refresh=refresh_corp_codes)
    rows = []
    for candidate in candidates:
        row = {
            "작업일": run_date.isoformat(),
            **candidate,
            "DART회사명": "",
            "접수번호": "",
            "공시일": "",
            "공시명": "",
            "DART_URL": "",
            "오류": "",
        }
        try:
            corp_code, dart_name = resolve_corp(corp_codes, candidate["회사"])
            row["DART회사명"] = dart_name
            if not corp_code:
                raise RuntimeError("DART corp_code를 찾지 못했습니다.")

            filings = search_investment_prospectus(
                api_key,
                corp_code,
                run_date - timedelta(days=365),
                run_date,
            )
            if not filings:
                raise RuntimeError("최근 1년 투자설명서 공시를 찾지 못했습니다.")

            filing = filings[0]
            row.update(filing)
            text = download_document_text(api_key, filing["접수번호"])
            prospectus_metrics = parse_prospectus_metrics(text)
            if not prospectus_metrics.get("증권사") and row.get("증권사"):
                prospectus_metrics["증권사"] = row["증권사"]
            row.update(prospectus_metrics)
            row = apply_issuance_report(api_key, corp_code, row, run_date)
            row = apply_listing_price(row, stock_code_for_corp(corp_codes, corp_code), run_date)
        except Exception as error:
            row["오류"] = str(error)
        rows.append(row)
    return pd.DataFrame(rows, columns=INTERNAL_COLUMNS)


def collect_company_prospectuses(company_names, run_date, search_start, search_end, refresh_corp_codes=False):
    api_key = load_dart_api_key()
    if not api_key:
        raise RuntimeError("DART API 키가 없습니다. DART_API_KEY 환경변수를 설정하세요.")

    corp_codes = load_corp_codes(api_key, refresh=refresh_corp_codes)
    naver_lookup = naver_ipo_lookup(run_date)
    rows = []
    for company_name in company_names:
        company_name = compact_text(company_name)
        naver_info = naver_lookup.get(normalize_name(company_name), {})
        row = {
            "작업일": run_date.isoformat(),
            "회사": company_name,
            "업종": naver_info.get("업종", ""),
            "개인청약_시작일": naver_info.get("개인청약_시작일", ""),
            "상장일": naver_info.get("상장일", ""),
            "네이버_일정열": "회사명 직접입력",
            "네이버_원자료": naver_info.get("네이버_원자료", ""),
            "증권사": naver_info.get("증권사", ""),
            "일반청약_경쟁률": naver_info.get("일반청약_경쟁률", pd.NA),
            "DART회사명": "",
            "접수번호": "",
            "공시일": "",
            "공시명": "",
            "DART_URL": "",
            "오류": "",
        }
        try:
            corp_code, dart_name = resolve_corp(corp_codes, company_name)
            row["DART회사명"] = dart_name
            if not corp_code:
                raise RuntimeError("DART corp_code를 찾지 못했습니다.")

            filings = search_investment_prospectus(api_key, corp_code, search_start, search_end)
            if not filings:
                raise RuntimeError(f"{search_start:%Y-%m-%d}~{search_end:%Y-%m-%d} 투자설명서 공시를 찾지 못했습니다.")

            filing = filings[0]
            row.update(filing)
            text = download_document_text(api_key, filing["접수번호"])
            prospectus_metrics = parse_prospectus_metrics(text)
            if not prospectus_metrics.get("증권사") and row.get("증권사"):
                prospectus_metrics["증권사"] = row["증권사"]
            row.update(prospectus_metrics)
            row = apply_issuance_report(api_key, corp_code, row, run_date)
            row = apply_listing_price(row, stock_code_for_corp(corp_codes, corp_code), run_date)
        except Exception as error:
            row["오류"] = str(error)
        rows.append(row)
    return pd.DataFrame(rows, columns=INTERNAL_COLUMNS)


def merge_with_existing(new_df, output_csv):
    key_cols = ["작업일", "회사", "개인청약_시작일", "접수번호"]
    if output_csv.exists():
        old_df = pd.read_csv(output_csv, dtype=str).fillna("")
        combined = pd.concat([old_df, new_df.astype(str).fillna("")], ignore_index=True)
    else:
        combined = new_df.astype(str).fillna("")
    for col in key_cols:
        if col not in combined.columns:
            combined[col] = ""
    combined = combined.drop_duplicates(subset=key_cols, keep="last")
    for col in INTERNAL_COLUMNS:
        if col not in combined.columns:
            combined[col] = ""
    return combined[INTERNAL_COLUMNS]


def normalize_ratio_for_excel(value):
    value = safe_float(value)
    if pd.isna(value):
        return pd.NA
    return float(value) / 100 if float(value) > 1 else float(value)


def format_percent(value):
    value = normalize_ratio_for_excel(value)
    if pd.isna(value):
        return ""
    return f"{float(value) * 100:.2f}%"


def format_return_percent(value):
    value = safe_float(value)
    if pd.isna(value):
        return ""
    return f"{float(value) * 100:.2f}%"


def format_market_cap(value):
    value = round_to_integer(value)
    if pd.isna(value):
        return ""
    return f"{value}억"


def classify_ipo_type(company_name):
    normalized = str(company_name).lower()
    if re.search(r"스팩|스펙|spec|spac", normalized):
        return "스펙"
    return "일반"


def to_excel_shape(df):
    for col in INTERNAL_COLUMNS:
        if col not in df.columns:
            df[col] = ""
    shaped = pd.DataFrame()
    shaped["번호"] = range(1, len(df) + 1)
    shaped["구분"] = df["회사"].map(classify_ipo_type)
    shaped["종목"] = df["회사"]
    shaped["업종"] = df["업종"]
    shaped["증권사"] = df["증권사"]
    shaped["상장일"] = df["상장일"]
    shaped["수요예측\n경쟁률"] = df["수요예측_경쟁률"].map(safe_float)
    shaped["일반청약\n경쟁률"] = df["일반청약_경쟁률"].map(safe_float)
    shaped["의무확약비율(전)"] = df["의무확약비율_전"].map(format_percent)
    shaped["의무확약비율(후)"] = df["의무확약비율_후"].map(format_percent)
    shaped["유통주식수 비율"] = df["유통가능비율"].map(format_percent)
    shaped["유통주식 비율(후)"] = df["유통가능비율_후"].map(format_percent)
    shaped["시가총액"] = df["시가총액_억원"].map(format_market_cap)
    shaped["환매청구권"] = df["환매청구권"]
    shaped["공모가"] = df["공모가"].map(round_to_integer)
    shaped["시초가"] = df["시초가"].map(round_to_integer)
    shaped["시초가 수익률"] = df["시초가_수익률"].map(format_return_percent)
    shaped["첫날 종가"] = df["첫날_종가"].map(round_to_integer)
    shaped["종가 수익률"] = df["종가_수익률"].map(format_return_percent)
    shaped["평균 매도가"] = ""
    shaped["수익률"] = ""
    shaped["수익금"] = ""
    shaped["세금 고려"] = ""
    return shaped[EXCEL_COLUMNS]


def save_outputs(df, output_csv=OUTPUT_CSV, output_xlsx=OUTPUT_XLSX, raw_output_csv=RAW_OUTPUT_CSV):
    output_csv.parent.mkdir(parents=True, exist_ok=True)
    for col in INTERNAL_COLUMNS:
        if col not in df.columns:
            df[col] = ""
    raw_df = df[INTERNAL_COLUMNS]
    raw_df.to_csv(raw_output_csv, index=False, encoding="utf-8-sig")
    excel_df = to_excel_shape(raw_df)
    excel_df.to_csv(output_csv, index=False, encoding="utf-8-sig")
    with pd.ExcelWriter(output_xlsx, engine="openpyxl") as writer:
        excel_df.to_excel(writer, index=False, sheet_name="IPO자동수집")


def parse_args():
    parser = argparse.ArgumentParser(description="네이버 IPO 일정과 DART 투자설명서 기반 공모주 핵심지표 자동수집")
    parser.add_argument("--date", default=date.today().isoformat(), help="작업일 YYYY-MM-DD. 기본값은 오늘")
    parser.add_argument("--companies", nargs="*", default=[], help="네이버 일정 대신 회사명을 직접 입력해 투자설명서를 검색합니다.")
    parser.add_argument("--search-start", default="2020-01-01", help="--companies 검색 시작일 YYYY-MM-DD")
    parser.add_argument("--search-end", default=date.today().isoformat(), help="--companies 검색 종료일 YYYY-MM-DD")
    parser.add_argument("--refresh-corp-codes", action="store_true", help="DART corpCode 캐시를 새로 받습니다.")
    parser.add_argument("--no-merge", action="store_true", help="기존 결과와 병합하지 않고 이번 실행 결과만 저장합니다.")
    parser.add_argument("--dry-run", action="store_true", help="파일 저장 없이 결과를 출력합니다.")
    return parser.parse_args()


def main():
    args = parse_args()
    run_date = datetime.strptime(args.date, "%Y-%m-%d").date()
    if args.companies:
        search_start = datetime.strptime(args.search_start, "%Y-%m-%d").date()
        search_end = datetime.strptime(args.search_end, "%Y-%m-%d").date()
        result = collect_company_prospectuses(
            args.companies,
            run_date,
            search_start,
            search_end,
            refresh_corp_codes=args.refresh_corp_codes,
        )
    else:
        result = collect_ipo_watch(run_date, refresh_corp_codes=args.refresh_corp_codes)
    output = result if args.no_merge else merge_with_existing(result, RAW_OUTPUT_CSV)
    output = update_existing_listing_prices(output, run_date, refresh_corp_codes=args.refresh_corp_codes)

    if args.dry_run:
        if output.empty:
            print("내일 개인청약 시작 종목이 없습니다.")
        else:
            print(to_excel_shape(output).to_string(index=False))
        return

    save_outputs(output)
    print(f"저장 완료: {OUTPUT_CSV} ({len(output)}행)")
    print(f"저장 완료: {OUTPUT_XLSX}")


if __name__ == "__main__":
    main()
