from pathlib import Path
import os
import re
from datetime import date

import altair as alt
import pandas as pd
import streamlit as st


BASE_DIR = Path(__file__).resolve().parent
HISTORY_FILE = BASE_DIR / "ipo_history.csv"
WATCH_FILE = BASE_DIR / "ipo_watch_results.csv"
MANUAL_INPUT_FILE = BASE_DIR / "ipo_manual_inputs.csv"


st.set_page_config(page_title="공모주 투자 대시보드", layout="wide")


WATCH_COLUMNS = [
    "번호",
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
    "종가",
    "종가 수익률",
    "평균 매도가",
    "수익률",
    "수익금",
]
WATCH_AMOUNT_COLUMNS = ["공모가", "시초가", "종가", "평균 매도가", "수익금"]
MANUAL_INPUT_COLUMNS = ["_row_id", "평균 매도가", "수익금"]
WATCH_STATS_COLUMNS = [
    "번호",
    "구분",
    "종목",
    "수요예측\n경쟁률",
    "일반청약\n경쟁률",
    "의무확약비율(전)",
    "의무확약비율(후)",
    "유통주식수 비율",
    "유통주식 비율(후)",
    "시가총액",
    "시초가 수익률",
    "종가 수익률",
    "수익률",
    "수익금",
]
PROTECTED_COLUMNS = ["수익금"]
WATCH_HIGHER_BETTER_COLUMNS = ["수요예측\n경쟁률", "일반청약\n경쟁률", "의무확약비율(전)", "의무확약비율(후)"]
WATCH_LOWER_BETTER_COLUMNS = ["유통주식수 비율", "유통주식 비율(후)"]
WATCH_COMPETITION_COLUMNS = ["수요예측\n경쟁률", "일반청약\n경쟁률"]
WATCH_PERCENT_DISPLAY_COLUMNS = [
    "의무확약비율(전)",
    "의무확약비율(후)",
    "유통주식수 비율",
    "유통주식 비율(후)",
    "시초가 수익률",
    "종가 수익률",
    "수익률",
]


st.markdown(
    """
    <style>
    .blur-value {
        filter: blur(7px);
        user-select: none;
        display: inline-block;
    }
    .protected-metric-label {
        color: #6b7280;
        font-size: 0.9rem;
        margin-bottom: 0.25rem;
    }
    .protected-metric-value {
        font-size: 2.2rem;
        line-height: 1.2;
        font-weight: 500;
    }
    .protected-panel {
        min-height: 330px;
        border: 1px solid #e5e7eb;
        border-radius: 8px;
        display: flex;
        align-items: center;
        justify-content: center;
        color: #6b7280;
        background: #f9fafb;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


def to_number(series):
    return pd.to_numeric(series, errors="coerce")


def pct(value, digits=1):
    if pd.isna(value):
        return "-"
    return f"{value * 100:.{digits}f}%"


def money(value, unit="원", digits=0):
    if pd.isna(value):
        return "-"
    return f"{value:,.{digits}f}{unit}"


def admin_password():
    secret_password = ""
    try:
        secret_password = st.secrets.get("IPO_ADMIN_PASSWORD", "")
    except Exception:
        secret_password = ""
    return os.getenv("IPO_ADMIN_PASSWORD") or secret_password or "admin"


def protected_metric_card(label, value, is_admin):
    if is_admin:
        metric_card(label, value)
    else:
        st.markdown(
            f"""
            <div class="protected-metric-label">{label}</div>
            <div class="protected-metric-value"><span class="blur-value">{value}</span></div>
            """,
            unsafe_allow_html=True,
        )


def mask_public_columns(df, columns=PROTECTED_COLUMNS):
    masked = df.copy()
    for col in columns:
        if col in masked.columns:
            masked[col] = "••••"
    return masked


def value_or_blank(value):
    if pd.isna(value):
        return ""
    return value


def rounded_number_or_blank(value):
    if pd.isna(value):
        return ""
    try:
        number = float(value)
    except (TypeError, ValueError):
        return value
    return int(round(number)) if number.is_integer() else round(number, 2)


def parse_number(value):
    if pd.isna(value):
        return pd.NA
    text = str(value).strip()
    if not text:
        return pd.NA
    text = re.sub(r"[^0-9.\-]", "", text)
    if not text:
        return pd.NA
    return pd.to_numeric(text, errors="coerce")


def comma_number_string(value, digits=0):
    number = parse_number(value)
    if pd.isna(number):
        return ""
    return f"{float(number):,.{digits}f}"


def percent_string(value, digits=2):
    if pd.isna(value):
        return ""
    try:
        return f"{float(value) * 100:.{digits}f}%"
    except (TypeError, ValueError):
        return value


def market_cap_string(value):
    if pd.isna(value):
        return ""
    try:
        return f"{round(float(value)):,}억"
    except (TypeError, ValueError):
        return value


def percent_display_from_display_values(values, digits=1):
    numbers = values.map(parse_number).dropna()
    if numbers.empty:
        return ""
    return f"{numbers.mean():.{digits}f}%"


def average_display_number(values, digits=0):
    numbers = values.map(parse_number).dropna()
    if numbers.empty:
        return ""
    return comma_number_string(numbers.mean(), digits=digits)


def sum_display_number(values):
    numbers = values.map(parse_number).dropna()
    if numbers.empty:
        return ""
    return comma_number_string(numbers.sum())


def rows_with_numeric_value(table, column):
    if column not in table.columns:
        return table.iloc[0:0]
    return table[table[column].map(parse_number).notna()]


def build_watch_stats(table, group_label="전체"):
    stats_rows = []
    current_year = date.today().year
    for label, year in [("올해", current_year), ("작년", current_year - 1)]:
        year_table = table[pd.to_numeric(table["_연도"], errors="coerce").eq(year)]
        row = {col: "" for col in WATCH_STATS_COLUMNS}
        row["번호"] = label
        row["구분"] = group_label
        row["종목"] = f"{len(year_table):,}건"
        if not year_table.empty:
            for col in ["수요예측\n경쟁률", "일반청약\n경쟁률"]:
                row[col] = average_display_number(year_table[col], digits=2)
            row["시가총액"] = market_cap_string(year_table["시가총액"].map(parse_number).dropna().mean())
            for col in [
                "의무확약비율(전)",
                "의무확약비율(후)",
                "유통주식수 비율",
                "유통주식 비율(후)",
            ]:
                row[col] = percent_display_from_display_values(year_table[col])
            open_price_table = rows_with_numeric_value(year_table, "시초가")
            close_price_table = rows_with_numeric_value(year_table, "종가")
            avg_sell_price_table = rows_with_numeric_value(year_table, "평균 매도가")
            row["시초가 수익률"] = percent_display_from_display_values(open_price_table["시초가 수익률"])
            row["종가 수익률"] = percent_display_from_display_values(close_price_table["종가 수익률"])
            row["수익률"] = percent_display_from_display_values(avg_sell_price_table["수익률"])
            row["수익금"] = sum_display_number(avg_sell_price_table["수익금"])
        stats_rows.append(row)
    return pd.DataFrame(stats_rows, columns=WATCH_STATS_COLUMNS)


def current_year_watch_averages(table):
    current_year = date.today().year
    current_table = table[pd.to_numeric(table["_연도"], errors="coerce").eq(current_year)]
    averages = {}
    for col in [*WATCH_HIGHER_BETTER_COLUMNS, *WATCH_LOWER_BETTER_COLUMNS]:
        if col in current_table.columns:
            values = current_table[col].map(parse_number).dropna()
            if not values.empty:
                averages[col] = values.mean()
    return averages


def compare_to_average_style(value, average, higher_is_better=True):
    number = parse_number(value)
    if pd.isna(number) or pd.isna(average) or average == 0:
        return ""

    ratio = float(number) / float(average)
    if higher_is_better:
        if ratio >= 1.2:
            return "color: #B91C1C; font-weight: 800;"
        if ratio >= 1.05:
            return "color: #EF4444; font-weight: 600;"
        if ratio <= 0.8:
            return "color: #1D4ED8; font-weight: 800;"
        if ratio <= 0.95:
            return "color: #60A5FA; font-weight: 600;"
    else:
        if ratio <= 0.8:
            return "color: #B91C1C; font-weight: 800;"
        if ratio <= 0.95:
            return "color: #EF4444; font-weight: 600;"
        if ratio >= 1.2:
            return "color: #1D4ED8; font-weight: 800;"
        if ratio >= 1.05:
            return "color: #60A5FA; font-weight: 600;"
    return ""


def market_cap_style(value):
    number = parse_number(value)
    if pd.isna(number):
        return ""
    if number <= 1000:
        return "color: #B91C1C; font-weight: 800;"
    if number >= 1500:
        return "font-weight: 700;"
    return ""


def style_watch_table(display_table, full_table):
    averages = current_year_watch_averages(full_table)

    def style_cell(value, col):
        if col in WATCH_HIGHER_BETTER_COLUMNS and col in averages:
            return compare_to_average_style(value, averages[col], higher_is_better=True)
        if col in WATCH_LOWER_BETTER_COLUMNS and col in averages:
            return compare_to_average_style(value, averages[col], higher_is_better=False)
        if col == "시가총액":
            return market_cap_style(value)
        return ""

    styles = pd.DataFrame("", index=display_table.index, columns=display_table.columns)
    for col in display_table.columns:
        if col in [*WATCH_HIGHER_BETTER_COLUMNS, *WATCH_LOWER_BETTER_COLUMNS, "시가총액"]:
            styles[col] = display_table[col].map(lambda value, col=col: style_cell(value, col))
    return display_table.style.apply(lambda _: styles, axis=None)


def prepare_watch_display_table(table):
    display_table = table.copy()
    for col in WATCH_COMPETITION_COLUMNS:
        if col in display_table.columns:
            display_table[col] = display_table[col].map(parse_number)
    for col in WATCH_PERCENT_DISPLAY_COLUMNS:
        if col in display_table.columns:
            display_table[col] = display_table[col].map(parse_number)
    return display_table


def prepare_watch_editor_table(table):
    editor_table = table[["_row_id", "종목", "상장일", "공모가", "평균 매도가", "수익금"]].copy()
    for col in ["공모가", "평균 매도가", "수익금"]:
        editor_table[col] = pd.to_numeric(editor_table[col].map(parse_number), errors="coerce")
    return editor_table


def watch_column_config():
    config = {"_row_id": None}
    for col in WATCH_COMPETITION_COLUMNS:
        config[col] = st.column_config.NumberColumn(col, format="%,.2f")
    for col in WATCH_PERCENT_DISPLAY_COLUMNS:
        config[col] = st.column_config.NumberColumn(col, format="%.2f%%")
    return config


PERCENT_TABLE_COLUMNS = [
    "밴드상단_초과비율",
    "의무확약비율_전",
    "의무확약비율_후",
    "유통주식수_비율",
    "유통주식_비율_후",
    "시초가_수익률",
    "종가_수익률",
    "수익률",
]


def format_ratio_cell(value, digits=1):
    if pd.isna(value):
        return ""
    try:
        return f"{float(value) * 100:.{digits}f}%"
    except (TypeError, ValueError):
        return value


def display_table(df):
    table = df.copy()
    for col in PERCENT_TABLE_COLUMNS:
        if col in table.columns:
            table[col] = table[col].map(format_ratio_cell)
    return table


def normalize_watch_columns(df):
    table = df.copy()
    if "종가" not in table.columns and "첫날 종가" in table.columns:
        table["종가"] = table["첫날 종가"]
    for col in WATCH_COLUMNS:
        if col not in table.columns:
            table[col] = ""
    return table[WATCH_COLUMNS]


def column_or_blank(df, col):
    if col in df.columns:
        return df[col]
    return pd.Series([""] * len(df), index=df.index)


def history_to_watch_table(df):
    df = apply_spec_normalization(df)
    table = pd.DataFrame(index=df.index)
    table["_연도"] = column_or_blank(df, "연도")
    table["_row_id"] = make_row_id(
        column_or_blank(df, "종목"),
        pd.Series([""] * len(df), index=df.index),
        column_or_blank(df, "공모가"),
        column_or_blank(df, "증권사"),
    )
    table["번호"] = column_or_blank(df, "번호")
    table["종목"] = column_or_blank(df, "종목")
    table["업종"] = column_or_blank(df, "업종")
    spec_mask = is_spec_row(column_or_blank(df, "구분"), table["종목"])
    table.loc[spec_mask, "업종"] = "스펙"
    table["증권사"] = column_or_blank(df, "증권사")
    table["상장일"] = ""
    table["수요예측\n경쟁률"] = column_or_blank(df, "수요예측_경쟁률").map(rounded_number_or_blank)
    table["일반청약\n경쟁률"] = column_or_blank(df, "일반청약_경쟁률").map(rounded_number_or_blank)
    table["의무확약비율(전)"] = column_or_blank(df, "의무확약비율_전").map(percent_string)
    table["의무확약비율(후)"] = column_or_blank(df, "의무확약비율_후").map(percent_string)
    table["유통주식수 비율"] = column_or_blank(df, "유통주식수_비율").map(percent_string)
    table["유통주식 비율(후)"] = column_or_blank(df, "유통주식_비율_후").map(percent_string)
    table["시가총액"] = column_or_blank(df, "시가총액_억원").map(market_cap_string)
    table["환매청구권"] = column_or_blank(df, "환매청구권").map(value_or_blank)
    table["공모가"] = column_or_blank(df, "공모가").map(rounded_number_or_blank)
    table["시초가"] = column_or_blank(df, "시초가").map(rounded_number_or_blank)
    table["시초가 수익률"] = column_or_blank(df, "시초가_수익률").map(percent_string)
    table["종가"] = column_or_blank(df, "첫날_종가").map(rounded_number_or_blank)
    table["종가 수익률"] = column_or_blank(df, "종가_수익률").map(percent_string)
    table["평균 매도가"] = column_or_blank(df, "평균_매도가").map(rounded_number_or_blank)
    table["수익률"] = column_or_blank(df, "수익률").map(percent_string)
    table["수익금"] = column_or_blank(df, "수익금").map(rounded_number_or_blank)
    return table[["_연도", "_row_id", *WATCH_COLUMNS]]


def watch_file_to_table(df):
    df = apply_spec_normalization(df)
    raw_type = column_or_blank(df, "구분")
    raw_name = column_or_blank(df, "종목")
    table = normalize_watch_columns(df)
    spec_mask = is_spec_row(raw_type, raw_name)
    table.loc[spec_mask, "업종"] = "스펙"
    listed_dates = pd.to_datetime(table["상장일"], errors="coerce")
    table.insert(0, "_연도", listed_dates.dt.year)
    table.insert(
        1,
        "_row_id",
        make_row_id(table["종목"], table["상장일"], table["공모가"], table["증권사"]),
    )
    return table


def merge_watch_and_history(watch_df, history_df):
    history_table = history_to_watch_table(history_df.sort_values(["연도", "번호"], ascending=[False, True]))
    if watch_df is None or watch_df.empty:
        return history_table

    watch_table = watch_file_to_table(watch_df)
    history_number_by_name = (
        history_table[["종목", "번호"]]
        .dropna(subset=["종목"])
        .drop_duplicates("종목", keep="first")
        .set_index("종목")["번호"]
        .to_dict()
    )
    watch_table["번호"] = watch_table.apply(
        lambda row: history_number_by_name.get(row["종목"], row["번호"]),
        axis=1,
    )
    duplicate_names = set(watch_table["종목"].dropna().astype(str))
    history_table = history_table[~history_table["종목"].astype(str).isin(duplicate_names)]
    return pd.concat([watch_table, history_table], ignore_index=True)


def sort_watch_table(table):
    sorted_table = table.copy()
    sorted_table["_상장일_정렬"] = pd.to_datetime(sorted_table["상장일"], errors="coerce")
    sorted_table["_연도_정렬"] = pd.to_numeric(sorted_table["_연도"], errors="coerce")
    sorted_table["_번호_정렬"] = pd.to_numeric(sorted_table["번호"].map(parse_number), errors="coerce")
    return (
        sorted_table.sort_values(
            ["_상장일_정렬", "_연도_정렬", "_번호_정렬"],
            ascending=[False, False, True],
            na_position="last",
        )
        .drop(columns=["_상장일_정렬", "_연도_정렬", "_번호_정렬"])
        .reset_index(drop=True)
    )


def is_spec_row(type_series, name_series):
    text = type_series.fillna("").astype(str) + " " + name_series.fillna("").astype(str)
    return text.str.contains("스팩|스펙|spec|히어로", case=False, regex=True, na=False)


def apply_spec_normalization(df):
    normalized = df.copy()
    if "종목" in normalized.columns:
        normalized["종목"] = normalized["종목"].replace({"키움히어로2호": "키움히어로스펙2호"})
    if "구분" in normalized.columns and "종목" in normalized.columns:
        spec_mask = is_spec_row(normalized["구분"], normalized["종목"])
        normalized.loc[spec_mask, "구분"] = "스펙"
        if "업종" in normalized.columns:
            normalized.loc[spec_mask, "업종"] = "스펙"
    return normalized


def filter_by_industry_group(df, selected_group):
    if selected_group == "전체" or "업종" not in df.columns:
        return df
    if selected_group == "일반":
        return df[df["업종"].ne("스펙")]
    if selected_group == "스펙":
        return df[df["업종"].eq("스펙")]
    return df


def make_row_id(name_series, listed_date_series, offer_price_series, broker_series):
    parts = [
        name_series.fillna("").astype(str).str.strip(),
        listed_date_series.fillna("").astype(str).str.strip(),
        offer_price_series.fillna("").astype(str).str.strip(),
        broker_series.fillna("").astype(str).str.strip(),
    ]
    return parts[0] + "|" + parts[1] + "|" + parts[2] + "|" + parts[3]


def load_manual_inputs():
    if not MANUAL_INPUT_FILE.exists():
        return pd.DataFrame(columns=MANUAL_INPUT_COLUMNS)
    manual = pd.read_csv(MANUAL_INPUT_FILE, dtype=str).fillna("")
    for col in MANUAL_INPUT_COLUMNS:
        if col not in manual.columns:
            manual[col] = ""
    return manual[MANUAL_INPUT_COLUMNS]


def apply_manual_inputs(table):
    manual = load_manual_inputs()
    if manual.empty:
        return table
    merged = table.merge(manual, on="_row_id", how="left", suffixes=("", "_manual"))
    for col in ["평균 매도가", "수익금"]:
        manual_col = f"{col}_manual"
        if manual_col in merged.columns:
            has_manual = merged[manual_col].fillna("").astype(str).str.strip() != ""
            merged.loc[has_manual, col] = merged.loc[has_manual, manual_col]
            merged = merged.drop(columns=[manual_col])
    return merged


def save_manual_inputs_from_editor(edited_table):
    manual = load_manual_inputs()
    manual = manual[~manual["_row_id"].isin(edited_table["_row_id"])]
    updates = edited_table[MANUAL_INPUT_COLUMNS].copy()
    has_value = (
        updates["평균 매도가"].fillna("").astype(str).str.strip().ne("")
        | updates["수익금"].fillna("").astype(str).str.strip().ne("")
    )
    updates = updates[has_value]
    manual = pd.concat([manual, updates], ignore_index=True)
    manual.to_csv(MANUAL_INPUT_FILE, index=False, encoding="utf-8-sig")


def format_watch_table(table):
    formatted = table.copy()
    for col in WATCH_AMOUNT_COLUMNS:
        if col in formatted.columns:
            formatted[col] = formatted[col].map(comma_number_string)
    offer_price = formatted["공모가"].map(parse_number)
    avg_price = formatted["평균 매도가"].map(parse_number)
    has_avg = offer_price.notna() & avg_price.notna() & offer_price.ne(0)
    formatted.loc[has_avg, "수익률"] = ((avg_price[has_avg] - offer_price[has_avg]) / offer_price[has_avg]).map(percent_string)
    return formatted


def load_history():
    if not HISTORY_FILE.exists():
        st.error("ipo_history.csv가 없습니다. 먼저 `.venv/bin/python ipo/build_ipo_history.py`를 실행하세요.")
        st.stop()
    df = pd.read_csv(HISTORY_FILE)
    if "연도" in df.columns:
        df = df[pd.to_numeric(df["연도"], errors="coerce").ne(2020)].copy()
    numeric_cols = [
        "연도",
        "번호",
        "수요예측_경쟁률",
        "밴드상단_초과비율",
        "일반청약_경쟁률",
        "의무확약비율_전",
        "의무확약비율_후",
        "유통주식수_비율",
        "유통주식_비율_후",
        "시가총액_억원",
        "공모가",
        "시초가",
        "시초가_수익률",
        "첫날_종가",
        "종가_수익률",
        "평균_매도가",
        "수익률",
        "수익금",
    ]
    for col in numeric_cols:
        if col in df.columns:
            df[col] = to_number(df[col])
    for col in ["구분", "종목", "업종", "증권사", "환매청구권", "특이점"]:
        if col in df.columns:
            df[col] = df[col].fillna("").astype(str)
    if "구분" in df.columns:
        df["구분"] = df["구분"].str.strip().replace("", "일반")
    if "증권사" in df.columns:
        df["증권사"] = df["증권사"].str.strip().replace("", "미기재")
    if "업종" in df.columns:
        df["업종"] = df["업종"].str.strip().replace("", "미기재")
    df = apply_spec_normalization(df)
    return df


def metric_card(label, value, delta=None):
    st.metric(label, value, delta)


def bar_chart(df, x, y, title, color="#2D9CDB", height=330):
    return (
        alt.Chart(df)
        .mark_bar(color=color, opacity=0.82)
        .encode(
            x=alt.X(x, title=""),
            y=alt.Y(y, title=title),
            tooltip=list(df.columns),
        )
        .properties(height=height)
    )


def line_chart(df, x, y, title, color="#EB5757", height=330):
    return (
        alt.Chart(df)
        .mark_line(point=True, color=color, strokeWidth=3)
        .encode(
            x=alt.X(x, title=""),
            y=alt.Y(y, title=title),
            tooltip=list(df.columns),
        )
        .properties(height=height)
    )


def scatter_chart(df, x, y, color_field, x_title, y_title, x_is_percent=False, show_profit=True):
    base = df.dropna(subset=[x, y]).copy()
    x_axis = alt.Axis(format=".1%") if x_is_percent else alt.Axis()
    x_tooltip_format = ".1%" if x_is_percent else ",.2f"
    tooltip = [
        alt.Tooltip("연도:O"),
        alt.Tooltip("종목:N"),
        alt.Tooltip("업종:N"),
        alt.Tooltip("증권사:N"),
        alt.Tooltip(f"{x}:Q", title=x_title, format=x_tooltip_format),
        alt.Tooltip(f"{y}:Q", title=y_title, format=".1%"),
    ]
    if show_profit:
        tooltip.append(alt.Tooltip("수익금:Q", title="수익금", format=",.0f"))
    return (
        alt.Chart(base)
        .mark_circle(size=80, opacity=0.72)
        .encode(
            x=alt.X(f"{x}:Q", title=x_title, axis=x_axis),
            y=alt.Y(f"{y}:Q", title=y_title, axis=alt.Axis(format="%")),
            color=alt.Color(f"{color_field}:N", title=color_field),
            tooltip=tooltip,
        )
        .properties(height=390)
    )


RETURN_METRIC_OPTIONS = {
    "시초가 수익률": ("시초가_수익률", "시초가"),
    "종가 수익률": ("종가_수익률", "첫날_종가"),
    "수익률": ("수익률", "평균_매도가"),
}


def build_return_metric_data(df, selected_metrics, required_columns=None):
    required_columns = required_columns or []
    frames = []
    for metric_name in selected_metrics:
        value_col, price_col = RETURN_METRIC_OPTIONS[metric_name]
        if value_col not in df.columns or price_col not in df.columns:
            continue
        metric_df = df.dropna(subset=[*required_columns, value_col, price_col]).copy()
        if metric_df.empty:
            continue
        metric_df["수익률구분"] = metric_name
        metric_df["표시수익률"] = metric_df[value_col]
        frames.append(metric_df)
    if not frames:
        return pd.DataFrame()
    return pd.concat(frames, ignore_index=True)


def metric_dropdown_selector(label, options, default, key):
    dropdown_options = ["전체", *options]
    selected = st.selectbox(
        label,
        dropdown_options,
        index=dropdown_options.index(default) if default in dropdown_options else 0,
        key=key,
    )
    if selected == "전체":
        return list(options)
    return [selected]


history = load_history()

st.title("공모주 투자 대시보드")
st.caption("기존 공모주 투자 엑셀을 표준화해 만든 분석 화면입니다.")

with st.sidebar:
    st.header("접속 모드")
    access_mode = st.radio("모드", ["일반", "관리자"], horizontal=True)
    is_admin = False
    if access_mode == "관리자":
        password = st.text_input("관리자 비밀번호", type="password")
        is_admin = password == admin_password()
        if not is_admin:
            st.warning("비밀번호가 맞아야 관리자 데이터가 표시됩니다.")
    st.divider()
    st.header("필터")
    years = sorted(history["연도"].dropna().astype(int).unique().tolist())
    selected_years = st.slider("연도", min_value=min(years), max_value=max(years), value=(min(years), max(years)))
    selected_industry_group = st.selectbox("업종 구분", ["전체", "일반", "스펙"], index=0)
    broker_options = sorted([v for v in history["증권사"].unique() if v])
    selected_brokers = st.multiselect("증권사", broker_options, default=broker_options)

view = history[history["연도"].between(selected_years[0], selected_years[1])].copy()
view = filter_by_industry_group(view, selected_industry_group).copy()
if selected_brokers:
    view = view[view["증권사"].isin(selected_brokers)]

result_view = view[view["수익률"].notna() & view["평균_매도가"].notna()].copy()

total_count = len(view)
profit_count = len(result_view)
avg_return = result_view["수익률"].mean()
win_rate = (result_view["수익률"] > 0).mean() if not result_view.empty else pd.NA
total_profit = result_view["수익금"].sum(min_count=1)

cols = st.columns(4)
with cols[0]:
    metric_card("공모주 수", f"{total_count:,}건")
with cols[1]:
    metric_card("성과 입력 종목", f"{profit_count:,}건")
with cols[2]:
    metric_card("평균 수익률", pct(avg_return, 1))
with cols[3]:
    protected_metric_card("누적 수익금", money(total_profit, "원", 0), is_admin)

tab_overview, tab_factor, tab_watch = st.tabs(
    ["성과 요약", "투자 지표", "자동수집"]
)

with tab_overview:
    yearly = (
        result_view.groupby("연도", as_index=False)
        .agg(
            종목수=("종목", "count"),
            평균수익률=("수익률", "mean"),
            승률=("수익률", lambda s: (s > 0).mean()),
            수익금=("수익금", "sum"),
            평균수요예측=("수요예측_경쟁률", "mean"),
            평균일반청약=("일반청약_경쟁률", "mean"),
        )
        .sort_values("연도")
    )
    yearly["평균수익률_pct"] = yearly["평균수익률"] * 100
    yearly["승률_pct"] = yearly["승률"] * 100
    yearly["누적수익금"] = yearly["수익금"].cumsum()
    c1, c2 = st.columns(2)
    with c1:
        st.subheader("연도별 평균 수익률")
        yearly_return_tooltip = [
            alt.Tooltip("연도:O"),
            alt.Tooltip("종목수:Q", title="종목 수"),
            alt.Tooltip("평균수익률_pct:Q", title="평균 수익률(%)", format=".0f"),
        ]
        if is_admin:
            yearly_return_tooltip.append(alt.Tooltip("수익금:Q", title="수익금", format=",.0f"))
        st.altair_chart(
            alt.Chart(yearly)
            .mark_line(point=True, color="#EB5757", strokeWidth=3)
            .encode(
                x=alt.X("연도:O", title=""),
                y=alt.Y(
                    "평균수익률_pct:Q",
                    title="평균 수익률(%)",
                    axis=alt.Axis(labelExpr="format(datum.value, '.0f') + '%'"),
                ),
                tooltip=yearly_return_tooltip,
            )
            .properties(height=330),
            width="stretch",
        )
    with c2:
        st.subheader("연도별 수익금")
        if is_admin:
            profit_bar = (
                alt.Chart(yearly)
                .mark_bar(color="#27AE60", opacity=0.78)
                .encode(
                    x=alt.X("연도:O", title=""),
                    y=alt.Y("수익금:Q", title="수익금(원)", axis=alt.Axis(format=",.0f")),
                    tooltip=[
                        alt.Tooltip("연도:O"),
                        alt.Tooltip("수익금:Q", title="연도별 수익금", format=",.0f"),
                        alt.Tooltip("누적수익금:Q", title="누적 수익금", format=",.0f"),
                    ],
                )
            )
            cumulative_line = (
                alt.Chart(yearly)
                .mark_line(point=True, color="#F2994A", strokeWidth=3)
                .encode(
                    x=alt.X("연도:O", title=""),
                    y=alt.Y(
                        "누적수익금:Q",
                        title="누적 수익금(원)",
                        axis=alt.Axis(format=",.0f", orient="right"),
                    ),
                    tooltip=[
                        alt.Tooltip("연도:O"),
                        alt.Tooltip("누적수익금:Q", title="누적 수익금", format=",.0f"),
                    ],
                )
            )
            st.altair_chart(
                (profit_bar + cumulative_line).resolve_scale(y="independent").properties(height=330),
                width="stretch",
            )
        else:
            st.markdown(
                '<div class="protected-panel"><span class="blur-value">관리자 전용 수익금 차트</span></div>',
                unsafe_allow_html=True,
            )

    c3, c4 = st.columns(2)
    with c3:
        st.subheader("수익률 분포")
        st.altair_chart(
            alt.Chart(result_view.dropna(subset=["수익률"]))
            .mark_bar(color="#56CCF2", opacity=0.82)
            .encode(
                x=alt.X("수익률:Q", bin=alt.Bin(maxbins=24), title="수익률", axis=alt.Axis(format="%")),
                y=alt.Y("count():Q", title="종목 수"),
                tooltip=[alt.Tooltip("count():Q", title="종목 수")],
            )
            .properties(height=320),
            width="stretch",
        )
    with c4:
        st.subheader("연도별 승률")
        st.altair_chart(
            alt.Chart(yearly)
            .mark_line(point=True, color="#9B51E0", strokeWidth=3)
            .encode(
                x=alt.X("연도:O", title=""),
                y=alt.Y(
                    "승률_pct:Q",
                    title="승률(%)",
                    axis=alt.Axis(labelExpr="format(datum.value, '.0f') + '%'"),
                    scale=alt.Scale(zero=False),
                ),
                tooltip=[
                    alt.Tooltip("연도:O"),
                    alt.Tooltip("종목수:Q", title="종목 수"),
                    alt.Tooltip("승률_pct:Q", title="승률(%)", format=".0f"),
                ],
            )
            .properties(height=320),
            width="stretch",
        )

with tab_factor:
    factor_years = sorted(view["연도"].dropna().astype(int).unique().tolist())
    if factor_years:
        selected_factor_year = st.selectbox(
            "투자지표 연도",
            ["전체", *sorted(factor_years, reverse=True)],
            key="factor_year",
        )
        factor_view = view.copy() if selected_factor_year == "전체" else view[view["연도"].eq(selected_factor_year)].copy()
    else:
        factor_view = view.copy()

    st.subheader("수요예측 경쟁률과 수익률")
    selected_return_metrics = metric_dropdown_selector(
        "표시할 수익률",
        list(RETURN_METRIC_OPTIONS.keys()),
        "전체",
        "return_metric_selector",
    )
    return_metric_data = build_return_metric_data(
        factor_view,
        selected_return_metrics,
        required_columns=["수요예측_경쟁률"],
    )
    if return_metric_data.empty:
        st.info("선택한 수익률을 표시할 수 있는 가격 데이터가 없습니다.")
    else:
        st.altair_chart(
            alt.Chart(return_metric_data)
            .mark_circle(size=78, opacity=0.72)
            .encode(
                x=alt.X("수요예측_경쟁률:Q", title="수요예측 경쟁률"),
                y=alt.Y("표시수익률:Q", title="수익률", axis=alt.Axis(format="%")),
                color=alt.Color(
                    "수익률구분:N",
                    title="수익률 구분",
                    scale=alt.Scale(
                        domain=["시초가 수익률", "종가 수익률", "수익률"],
                        range=["#2F80ED", "#F2994A", "#EB5757"],
                    ),
                ),
                tooltip=[
                    alt.Tooltip("연도:O"),
                    alt.Tooltip("종목:N"),
                    alt.Tooltip("업종:N"),
                    alt.Tooltip("증권사:N"),
                    alt.Tooltip("수익률구분:N", title="수익률 구분"),
                    alt.Tooltip("수요예측_경쟁률:Q", title="수요예측 경쟁률", format=",.2f"),
                    alt.Tooltip("표시수익률:Q", title="수익률", format=".1%"),
                    alt.Tooltip("시초가:Q", title="시초가", format=",.0f"),
                    alt.Tooltip("첫날_종가:Q", title="첫날 종가", format=",.0f"),
                    alt.Tooltip("평균_매도가:Q", title="평균 매도가", format=",.0f"),
                ],
            )
            .properties(height=430),
            width="stretch",
        )

    c1, c2 = st.columns(2)
    with c1:
        st.subheader("의무확약비율(후)과 수익률")
        st.altair_chart(
            scatter_chart(
                factor_view.dropna(subset=["평균_매도가"]),
                "의무확약비율_후",
                "수익률",
                "구분",
                "의무확약비율(후)",
                "수익률",
                x_is_percent=True,
                show_profit=is_admin,
            ),
            width="stretch",
        )
    with c2:
        st.subheader("유통주식수 비율과 수익률")
        st.altair_chart(
            scatter_chart(
                factor_view.dropna(subset=["평균_매도가"]),
                "유통주식수_비율",
                "수익률",
                "구분",
                "유통주식수 비율",
                "수익률",
                x_is_percent=True,
                show_profit=is_admin,
            ),
            width="stretch",
        )

    st.subheader("업종별 평균 수익률")
    selected_industry_metrics = metric_dropdown_selector(
        "업종별 수익률 기준",
        list(RETURN_METRIC_OPTIONS.keys()),
        "수익률",
        "industry_return_metric_selector",
    )
    industry_metric_data = build_return_metric_data(factor_view, selected_industry_metrics)
    if industry_metric_data.empty:
        st.info("선택한 수익률을 집계할 수 있는 가격 데이터가 없습니다.")
    else:
        industry = (
            industry_metric_data.groupby(["업종", "수익률구분"], as_index=False)
            .agg(종목수=("종목", "count"), 평균수익률=("표시수익률", "mean"), 수익금=("수익금", "sum"))
            .query("종목수 >= 2")
        )
        top_industries = (
            industry.groupby("업종")["평균수익률"]
            .max()
            .sort_values(ascending=False)
            .head(20)
            .index
        )
        industry = industry[industry["업종"].isin(top_industries)].copy()
        industry_tooltip = [
            alt.Tooltip("업종:N"),
            alt.Tooltip("수익률구분:N", title="수익률 구분"),
            alt.Tooltip("종목수:Q"),
            alt.Tooltip("평균수익률:Q", title="평균 수익률", format=".1%"),
        ]
        if is_admin:
            industry_tooltip.append(alt.Tooltip("수익금:Q", format=",.0f"))
        st.altair_chart(
            alt.Chart(industry)
            .mark_bar(opacity=0.82)
            .encode(
                x=alt.X("평균수익률:Q", title="평균 수익률", axis=alt.Axis(format="%")),
                y=alt.Y("업종:N", sort=alt.SortField("평균수익률", order="descending"), title=""),
                color=alt.Color(
                    "수익률구분:N",
                    title="수익률 구분",
                    scale=alt.Scale(
                        domain=["시초가 수익률", "종가 수익률", "수익률"],
                        range=["#2F80ED", "#F2994A", "#EB5757"],
                    ),
                ),
                tooltip=industry_tooltip,
            )
            .properties(height=520),
            width="stretch",
        )

with tab_watch:
    st.subheader("자동수집 후보")
    watch_df = pd.read_csv(WATCH_FILE) if WATCH_FILE.exists() else pd.DataFrame()
    watch_table = merge_watch_and_history(watch_df, history).fillna("")
    watch_table = apply_manual_inputs(watch_table)
    watch_table = format_watch_table(watch_table)
    watch_table = filter_by_industry_group(watch_table, selected_industry_group).copy()
    watch_table_for_style = watch_table.copy()
    watch_stats = build_watch_stats(watch_table, selected_industry_group).astype(str)
    watch_years = sorted(
        pd.to_numeric(watch_table["_연도"], errors="coerce").dropna().astype(int).unique().tolist(),
        reverse=True,
    )
    watch_year_options = ["전체", *watch_years]
    default_watch_index = watch_year_options.index(date.today().year) if date.today().year in watch_year_options else 0
    selected_watch_year = st.selectbox("자동수집 연도", watch_year_options, index=default_watch_index, key="watch_year")
    if selected_watch_year != "전체":
        watch_table = watch_table[pd.to_numeric(watch_table["_연도"], errors="coerce").eq(selected_watch_year)]
    watch_table = sort_watch_table(watch_table)
    watch_editor_table = prepare_watch_editor_table(watch_table)
    watch_table = prepare_watch_display_table(watch_table[["_row_id", *WATCH_COLUMNS]])
    st.caption("신규 자동수집 결과와 첨부 엑셀 기반 과거 데이터를 같은 컬럼 형식으로 합쳐서 표시합니다.")
    watch_stats_display = watch_stats if is_admin else mask_public_columns(watch_stats, ["수익금"])
    st.dataframe(watch_stats_display, width="stretch", hide_index=True)
    visible_watch_table = watch_table if is_admin else mask_public_columns(watch_table)
    st.dataframe(
        style_watch_table(visible_watch_table, watch_table_for_style),
        width="stretch",
        hide_index=True,
        column_config=watch_column_config(),
    )
    if is_admin:
        with st.expander("평균 매도가/수익금 입력"):
            edited_watch_table = st.data_editor(
                watch_editor_table,
                width="stretch",
                hide_index=True,
                disabled=["종목", "상장일", "공모가"],
                column_config={
                    "_row_id": None,
                    "공모가": st.column_config.NumberColumn("공모가", format="%,.0f"),
                    "평균 매도가": st.column_config.NumberColumn("평균 매도가", format="%,.0f"),
                    "수익금": st.column_config.NumberColumn("수익금", format="%,.0f"),
                },
                num_rows="fixed",
                key="watch_manual_input_editor",
            )
            if st.button("평균 매도가/수익금 저장"):
                save_manual_inputs_from_editor(edited_watch_table)
                st.success("입력값을 저장했습니다.")
                st.rerun()
