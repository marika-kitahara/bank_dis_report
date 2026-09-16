import io
import re
import time
from collections import defaultdict

import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(page_title="後方数値分析用(Display)", layout="wide")
st.title("後方数値分析用(Display) 作成")
st.caption("ファイル1 + 複数のファイル2をアップロードして、集計済みExcelを出力します。")

OUTPUT_COLUMNS = ["日", "キャンペーン", "表示回数", "クリック数", "コンバージョン", "通貨コード", "費用"]
MEDIA_ORDER = ["YDN", "Pmax", "LINE", "Youtube", "Criteo", "Meta", "X"]
RAW_SHEETS = {
    "YDN": "【YDN】ローデータ",
    "Pmax": "【Pmax】ローデータ",
    "LINE": "【LINE】ローデータ",
    "Youtube": "【Youtube】ローデータ",
    "Criteo": "【Criteo】ローデータ",
    "Meta": ["【Facebook】ローデータ", "【Facabook】ローデータ"],  # 正式表記＋旧仕様の誤記も許容
    "X": "【X】ローデータ",
}

COLUMN_MAP = {
    "YDN": {
        "日": "日", "キャンペーン名": "キャンペーン", "インプレッション数": "表示回数",
        "クリック数": "クリック数", "コンバージョン数": "コンバージョン", "コスト": "費用",
    },
    "Pmax": {
        "日": "日", "キャンペーン": "キャンペーン", "表示回数": "表示回数",
        "クリック数": "クリック数", "コンバージョン": "コンバージョン",
        "通貨コード": "通貨コード", "費用": "費用",
    },
    "LINE": {
        "日": "日", "キャンペーン名": "キャンペーン", "インプレッション数": "表示回数",
        "クリック数": "クリック数", "CV（LINE Tagクリック）": "コンバージョン", "ご利用金額": "費用",
    },
    "Youtube": {
        "日": "日", "キャンペーン": "キャンペーン", "表示回数": "表示回数",
        "クリック数": "クリック数", "コンバージョン（現在のモデル）": "コンバージョン", "費用": "費用",
    },
    "Criteo": {
        "日": "日", "キャンペーン": "キャンペーン", "表示回数": "表示回数",
        "クリック数": "クリック数", "コンバージョン（現在のモデル）": "コンバージョン", "費用": "費用",
    },
    "Meta": {
        "日": "日", "キャンペーン名": "キャンペーン", "インプレッション": "表示回数",
        "クリック(すべて)": "クリック数", "結果": "コンバージョン", "消化金額 (JPY)": "費用",
    },
    "X": {
        "Time period": "日", "Ad Group name": "キャンペーン", "Impressions": "表示回数",
        "Clicks": "クリック数", "Leads": "コンバージョン", "Spend": "費用",
    },
}


def clean_col(v):
    return str(v).replace("\u3000", " ").strip() if v is not None else ""


def normalize_text(v):
    if pd.isna(v):
        return ""
    return str(v).strip()


def read_sheet(file_obj, sheet_name):
    """存在しない/空のシートはNone。候補リストなら最初に存在するシートを読む。"""
    file_obj.seek(0)
    xls = pd.ExcelFile(file_obj, engine="openpyxl")
    candidates = sheet_name if isinstance(sheet_name, (list, tuple)) else [sheet_name]
    actual = next((name for name in candidates if name in xls.sheet_names), None)
    if actual is None:
        return None
    df = pd.read_excel(xls, sheet_name=actual, dtype=object)
    df.columns = [clean_col(c) for c in df.columns]
    if df.empty or df.dropna(how="all").empty:
        return None
    return df.dropna(how="all").reset_index(drop=True)


def find_first_sheet(file_obj, candidates):
    file_obj.seek(0)
    xls = pd.ExcelFile(file_obj, engine="openpyxl")
    for name in candidates:
        if name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=name, dtype=object)
            df.columns = [clean_col(c) for c in df.columns]
            return df.dropna(how="all").reset_index(drop=True)
    return None


def coerce_date(s):
    return pd.to_datetime(s, errors="coerce").dt.normalize()


def coerce_num(s):
    if s is None:
        return pd.Series(dtype=float)
    if pd.api.types.is_numeric_dtype(s):
        return pd.to_numeric(s, errors="coerce")
    return pd.to_numeric(
        s.astype(str).str.replace(",", "", regex=False).str.replace("¥", "", regex=False).str.strip(),
        errors="coerce",
    )


def standardize_media(df, media):
    mapping = COLUMN_MAP[media]
    missing = [c for c in mapping if c not in df.columns]
    # Youtube/Criteoは仕様上、通貨コード列が空見出しの可能性があるため要求しない
    if missing:
        raise ValueError(f"{media}: 必要列が見つかりません → {', '.join(missing)}")

    out = pd.DataFrame(index=df.index)
    for src, dst in mapping.items():
        out[dst] = df[src]
    for c in OUTPUT_COLUMNS:
        if c not in out.columns:
            out[c] = "" if c == "通貨コード" else np.nan

    out = out[OUTPUT_COLUMNS]
    out["日"] = coerce_date(out["日"])
    out["キャンペーン"] = out["キャンペーン"].map(normalize_text)
    for c in ["表示回数", "クリック数", "コンバージョン", "費用"]:
        out[c] = coerce_num(out[c]).fillna(0)
    out["通貨コード"] = out["通貨コード"].fillna("").astype(str).replace("nan", "")
    out = out[out["日"].notna()].reset_index(drop=True)
    return out


def prepare_campaign(df):
    required = ["テンプレコード", "開始日", "終了日"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError("キャンペーン情報: 必要列が見つかりません → " + ", ".join(missing))
    x = df.copy()
    x["テンプレコード"] = x["テンプレコード"].map(normalize_text)
    x["開始日"] = coerce_date(x["開始日"])
    x["終了日"] = coerce_date(x["終了日"])
    return x


def prepare_media_master(df):
    required = ["テンプレコード", "大項目", "メニュー名", "媒体コード"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError("媒体コードマスタ: 必要列が見つかりません → " + ", ".join(missing))
    x = df.copy()
    for c in required:
        x[c] = x[c].map(normalize_text)
    return x


def build_campaign_date_map(campaign_df):
    """日付→テンプレコードを一度だけ作る。重複日は元仕様どおり先頭行を優先。"""
    date_map = {}
    valid = campaign_df.dropna(subset=["開始日", "終了日"])
    for row in valid.itertuples(index=False):
        code = normalize_text(getattr(row, "テンプレコード"))
        if not code:
            continue
        for dt in pd.date_range(getattr(row, "開始日"), getattr(row, "終了日"), freq="D"):
            date_map.setdefault(pd.Timestamp(dt).normalize(), code)
    return date_map


def build_master_index(master):
    """(テンプレコード, 大項目)→[(メニュー名小文字, 媒体コード)] に索引化。"""
    index = defaultdict(list)
    for row in master[["テンプレコード", "大項目", "メニュー名", "媒体コード"]].itertuples(index=False, name=None):
        period, media, menu, code = row
        period = normalize_text(period)
        media = normalize_text(media)
        menu = normalize_text(menu)
        code = normalize_text(code)
        if period and media and menu and code:
            index[(period, media)].append((menu.lower(), code))
    return index


def build_backward_index(backward):
    """後方数値を媒体・日付単位で索引化。"""
    valid = backward[(backward["__media"] != "") & backward["__date"].notna()].copy()
    media_total = valid.groupby("__media").size().to_dict()
    day_counts = valid.groupby(["__media", "__date"]).size().to_dict()
    codes_by_media_day = valid.groupby(["__media", "__date"])["__code"].apply(list).to_dict()
    rows_by_media = defaultdict(list)
    for idx, media, dt, code in valid[["__media", "__date", "__code"]].itertuples(index=True, name=None):
        rows_by_media[media].append((idx, dt, code))
    return media_total, day_counts, codes_by_media_day, rows_by_media


def media_codes_for_row_fast(campaign_name, period, media, master_index):
    if not period:
        return []
    parts = [p.lower() for p in str(campaign_name).split("_") if p]
    candidates = master_index.get((str(period), media), [])
    seen = set()
    codes = []
    for menu_lower, code in candidates:
        if (not parts or all(p in menu_lower for p in parts)) and code not in seen:
            seen.add(code)
            codes.append(code)
    return codes


def add_matching_columns_fast(df, media, campaign_date_map, master_index, backward_index):
    x = df.copy()
    media_total, day_counts, codes_by_media_day, _ = backward_index

    periods = [campaign_date_map.get(pd.Timestamp(dt).normalize(), "") for dt in x["日"]]
    codes_list = [
        media_codes_for_row_fast(cn, p, media, master_index)
        for cn, p in zip(x["キャンペーン"], periods)
    ]
    code_texts = ["該当なし" if not cs else f"{len(cs)}種類 ({'、'.join(cs)})" for cs in codes_list]
    code_texts_lower = [t.lower() for t in code_texts]

    x["期間"] = periods
    x["媒体コード"] = code_texts
    x["種類数"] = [len(cs) for cs in codes_list]
    x["転記用コスト(net)"] = np.where(x["種類数"].to_numpy() > 0, x["費用"].to_numpy() / x["種類数"].replace(0, np.nan).to_numpy(), np.nan)
    x["転記用コスト(gross)"] = x["転記用コスト(net)"] / 0.9
    x[media] = ""

    total_count = int(media_total.get(media, 0))
    match_counts = []
    date_counts = []
    alloc_units = []
    for dt, code_text_lower, cost in zip(x["日"], code_texts_lower, x["費用"]):
        key = (media, dt)
        day_count = int(day_counts.get(key, 0))
        if code_text_lower != "該当なし":
            count = sum(1 for c in codes_by_media_day.get(key, []) if c and c.lower() in code_text_lower)
        else:
            count = 0
        denom = count if count > 0 else (day_count if day_count > 0 else total_count)
        match_counts.append(count)
        date_counts.append(day_count)
        alloc_units.append(float(cost) / denom if denom else 0.0)

    x.insert(7, "媒体一致件数", match_counts)
    x.insert(8, "日付件数", date_counts)
    x.insert(9, "按分単価", alloc_units)
    return x


def calculate_backward_cost_fast(backward, media_frames, backward_index):
    """媒体ローデータ×後方データの全件二重ループをやめ、必要な対象行だけに加算。"""
    result = backward.copy()
    costs = np.zeros(len(result), dtype=float)
    _, _, _, rows_by_media = backward_index

    for media, df in media_frames.items():
        target_rows = rows_by_media.get(media, [])
        if not target_rows or df.empty:
            continue

        all_indices = [idx for idx, _, _ in target_rows]
        indices_by_day = defaultdict(list)
        code_rows_by_day = defaultdict(list)
        for idx, dt, code in target_rows:
            indices_by_day[dt].append(idx)
            code_rows_by_day[dt].append((idx, code.lower() if code else ""))

        for r in df[["日", "媒体一致件数", "日付件数", "按分単価", "媒体コード"]].itertuples(index=False, name=None):
            dt, match_count, day_count, unit, code_text = r
            unit = 0.0 if pd.isna(unit) else float(unit)
            match_count = int(match_count or 0)
            day_count = int(day_count or 0)

            if match_count > 0:
                code_text_lower = str(code_text).lower()
                for idx, code_lower in code_rows_by_day.get(dt, []):
                    if code_lower and code_lower in code_text_lower:
                        costs[idx] += unit
            elif day_count > 0:
                for idx in indices_by_day.get(dt, []):
                    costs[idx] += unit
            else:
                # 元式の「媒体一致件数=0 かつ 日付件数=0」は媒体内の全行へ配賦
                for idx in all_indices:
                    costs[idx] += unit

    result["集計コスト"] = costs
    return result

def _excel_safe_df(df, null_as_text=False):
    """Excel出力用に日時を日付へ落とし、必要なら欠損を文字列NULLへ変換。"""
    out = df.copy()
    for col in out.columns:
        if pd.api.types.is_datetime64_any_dtype(out[col]):
            out[col] = pd.to_datetime(out[col], errors="coerce").dt.date
    if null_as_text:
        out = out.astype(object).where(pd.notna(out), "NULL")
    return out


def _write_df_fast(writer, sheet_name, df, null_as_text=False):
    """XlsxWriterで高速出力。全使用セルはMeiryo UI。"""
    wb = writer.book
    df = _excel_safe_df(df, null_as_text=null_as_text)
    df.to_excel(writer, sheet_name=sheet_name, index=False, header=False, startrow=1,
                na_rep=("NULL" if null_as_text else ""))
    ws = writer.sheets[sheet_name]

    base_fmt = wb.add_format({"font_name": "Meiryo UI"})
    header_fmt = wb.add_format({
        "font_name": "Meiryo UI", "bold": True,
        "bg_color": "#D9EAF7", "align": "center"
    })
    date_fmt = wb.add_format({"font_name": "Meiryo UI", "num_format": "yyyy/m/d"})

    # ヘッダーは自前で書く（Pandas既定フォントを使わせない）
    for j, col in enumerate(df.columns):
        ws.write(0, j, str(col), header_fmt)

    # 列単位でフォント/日付書式を適用。全セルを後から走査しないので高速。
    for j, col in enumerate(df.columns):
        series = df[col]
        # NULL文字列や空欄が先頭にあっても、列内に実日付が1件でもあれば日付列として扱う。
        # 後方数値データのように日付とNULLが混在する列でも、日付セルを必ず
        # yyyy/m/d + Meiryo UI で再書込する。
        def _is_real_date(v):
            import datetime as _dt
            return isinstance(v, (pd.Timestamp, _dt.datetime, _dt.date)) and not isinstance(v, str)

        is_date = any(_is_real_date(v) for v in series.head(1000).tolist()) if len(series) else False
        fmt = date_fmt if is_date else base_fmt

        # pandas/XlsxWriter は日付セルに独自書式を付けるため、列書式だけでは
        # yyyy/m/d・Meiryo UI が反映されないことがある。日付列だけ明示的に再書込する。
        if is_date:
            for i, value in enumerate(series, start=1):
                if pd.isna(value) or value == "NULL":
                    if null_as_text and value == "NULL":
                        ws.write(i, j, "NULL", base_fmt)
                    else:
                        ws.write_blank(i, j, None, date_fmt)
                else:
                    # datetime/date のどちらでも 00:00:00 を表示させず日付だけに固定
                    if isinstance(value, pd.Timestamp):
                        value = value.to_pydatetime()
                    elif hasattr(value, "year") and hasattr(value, "month") and hasattr(value, "day"):
                        import datetime as _dt
                        value = _dt.datetime(value.year, value.month, value.day)
                    ws.write_datetime(i, j, value, date_fmt)

        # 自動幅は全行スキャンせず、先頭200行だけで概算。
        # 値型に関係なく必ず文字列化してから幅を測る。
        # Excel由来の列では float / int / Timestamp 等が混在することがあるため、
        # len(value) を直接呼ばない。
        sample_vals = [
            "" if pd.isna(v) else str(v)
            for v in series.head(200).tolist()
        ] if len(series) else []
        max_len = max([len(str(col))] + [len(str(v)) for v in sample_vals]) + 2
        ws.set_column(j, j, max(10, min(max_len, 35)), fmt)

    ws.freeze_panes(1, 0)
    if len(df.columns):
        ws.autofilter(0, 0, max(len(df), 1), len(df.columns) - 1)


def _write_cost_diff_sheet(writer):
    """添付の『コスト差分』シートを、出力ブック内参照に直して再現。"""
    wb = writer.book
    ws = wb.add_worksheet("コスト差分")
    writer.sheets["コスト差分"] = ws

    text_fmt = wb.add_format({"font_name": "Meiryo UI"})
    num_fmt = wb.add_format({"font_name": "Meiryo UI", "num_format": '#,##0_);[Red](#,##0)'})
    hair_fmt = wb.add_format({"font_name": "Meiryo UI", "border": 7, "num_format": '#,##0_);[Red](#,##0)'})
    hair_text_fmt = wb.add_format({"font_name": "Meiryo UI", "border": 7})

    # コスト差分シートは全列150px固定・目盛線なし
    ws.hide_gridlines(2)
    for col_idx in range(0, 9):  # A:I
        fmt = num_fmt if col_idx >= 2 else text_fmt
        ws.set_column_pixels(col_idx, col_idx, 150, fmt)
    ws.write("B2", "コスト差分確認", text_fmt)

    media = MEDIA_ORDER
    for j, m in enumerate(media, start=2):
        col_letter = chr(ord("A") + j)
        ws.write(2, j, m, hair_text_fmt)
        # 添付元の [1] 外部参照を、今回作るブック内のシート参照へ変更。
        ws.write_formula(3, j, f'=SUMIF(\'後方数値データ(加工版)\'!$AE:$AE,{col_letter}$3,\'後方数値データ(加工版)\'!$AF:$AF)', hair_fmt)
        ws.write_formula(4, j, f'=SUM(\'{m}\'!G:G)', hair_fmt)
        ws.write_formula(5, j, f'={col_letter}4-{col_letter}5', num_fmt)

    ws.write("B4", "後方数値", text_fmt)
    ws.write("B5", "ローデータ", text_fmt)
    ws.write("B6", "差分", text_fmt)


def to_excel_bytes(backward_out, campaign_df_original, master_original, media_frames, progress=None):
    """XlsxWriterベースの高速Excel出力。"""
    bio = io.BytesIO()

    helper_cols = [c for c in backward_out.columns if c.startswith("__")]
    source_cols = [c for c in backward_out.columns if c not in helper_cols and c != "集計コスト"]
    base_cols = source_cols[:31]  # A:AE
    b_out = backward_out[base_cols].copy()
    b_out["コスト"] = backward_out["集計コスト"].values  # AF

    with pd.ExcelWriter(
        bio,
        engine="xlsxwriter",
        date_format="yyyy/m/d",
        datetime_format="yyyy/m/d",
        engine_kwargs={"options": {"strings_to_formulas": False}},
    ) as writer:
        writer.book.set_calc_mode("auto")

        # 媒体コードマスタは出力ブックの先頭（一番左）に配置
        if progress: progress.progress(0.74, text="Excel：媒体コードマスタを書き出し中…")
        _write_df_fast(writer, "媒体コードマスタ", master_original)

        if progress: progress.progress(0.76, text="Excel：後方数値データを書き出し中…")
        _write_df_fast(writer, "後方数値データ(加工版)", b_out, null_as_text=True)

        # 添付シートは後方数値データの直後に配置
        if progress: progress.progress(0.80, text="Excel：コスト差分シートを作成中…")
        _write_cost_diff_sheet(writer)

        for i, media in enumerate(MEDIA_ORDER):
            if progress:
                progress.progress(0.82 + i * 0.015, text=f"Excel：{media}を書き出し中…")
            df = media_frames.get(media, pd.DataFrame(columns=OUTPUT_COLUMNS + ["媒体一致件数", "日付件数", "按分単価", "期間", "媒体コード", "種類数", "転記用コスト(net)", "転記用コスト(gross)", media]))
            _write_df_fast(writer, media, df)

        if progress: progress.progress(0.94, text="Excel：キャンペーン情報を書き出し中…")
        _write_df_fast(writer, "キャンペーン情報", campaign_df_original)

        if progress: progress.progress(0.99, text="Excelファイルを仕上げ中…")

    if progress: progress.progress(1.0, text="完了！")
    return bio.getvalue()


st.subheader("1. ファイルをアップロード")
file1 = st.file_uploader("ファイル1（後方数値データ(加工版)・キャンペーン情報・媒体コードマスタ）", type=["xlsx", "xlsm"], accept_multiple_files=False)
file2s = st.file_uploader("ファイル2（媒体ローデータ）※複数可", type=["xlsx", "xlsm"], accept_multiple_files=True)
st.info("媒体コードマスタはファイル1内の「媒体コードマスタ」「MediaMaster」「媒体コードマスタver3」から自動取得します。")

if st.button("レポートを作成", type="primary", disabled=not (file1 and file2s)):
    try:
        started_at = time.perf_counter()
        with st.spinner("読込・集計中…"):
            backward = find_first_sheet(file1, ["後方数値データ(加工版)"])
            campaign_original = find_first_sheet(file1, ["キャンペーン情報", "Campaign"])
            if backward is None:
                raise ValueError("ファイル1に「後方数値データ(加工版)」シートがありません。")
            if campaign_original is None:
                raise ValueError("ファイル1に「キャンペーン情報」シートがありません。")

            # 媒体コードマスタを探索
            master_original = find_first_sheet(file1, ["媒体コードマスタ", "MediaMaster", "媒体コードマスタver3"])
            if master_original is None:
                raise ValueError("ファイル1に媒体コードマスタが見つかりません。")

            campaign = prepare_campaign(campaign_original)
            master = prepare_media_master(master_original)

            # 後方数値データの参照列 B/C/AE を列位置で取得（Excel基準）
            if backward.shape[1] < 31:
                raise ValueError("後方数値データ(加工版)にAE列まで存在しません。")
            backward = backward.copy()
            backward["__date"] = coerce_date(backward.iloc[:, 1])   # B列
            backward["__code"] = backward.iloc[:, 2].map(normalize_text)  # C列
            backward["__media"] = backward.iloc[:, 30].map(normalize_text)  # AE列

            # 高速化用インデックスを最初に一度だけ作成
            campaign_date_map = build_campaign_date_map(campaign)
            master_index = build_master_index(master)
            backward_index = build_backward_index(backward)

            # 複数ファイル2を媒体別に縦結合。シート欠落/空はスキップ。
            raw_by_media = defaultdict(list)
            skipped = []
            for f in file2s:
                for media, sheet in RAW_SHEETS.items():
                    try:
                        df = read_sheet(f, sheet)
                        if df is not None:
                            raw_by_media[media].append(df)
                        else:
                            label = " / ".join(sheet) if isinstance(sheet, list) else sheet
                            skipped.append(f"{f.name} / {label}: 空またはシートなし")
                    except Exception as e:
                        label = " / ".join(sheet) if isinstance(sheet, list) else sheet
                        raise ValueError(f"{f.name} / {label} の読込でエラー: {e}") from e

            media_frames = {}
            processed = []
            progress = st.progress(0, text="媒体データを処理中…")
            for media_no, media in enumerate(MEDIA_ORDER, start=1):
                if raw_by_media[media]:
                    raw = pd.concat(raw_by_media[media], ignore_index=True)
                    std = standardize_media(raw, media)
                    calc = add_matching_columns_fast(std, media, campaign_date_map, master_index, backward_index)
                    media_frames[media] = calc
                    processed.append(f"{media}: {len(calc):,}行")
                else:
                    media_frames[media] = pd.DataFrame()
                progress.progress(0.08 + media_no * 0.07, text=f"集計：{media} を処理しました")

            progress.progress(0.62, text="集計：後方数値へコストを反映中…")
            backward_out = calculate_backward_cost_fast(backward, media_frames, backward_index)
            progress.progress(0.74, text="集計完了。Excel出力を開始します…")
            excel_bytes = to_excel_bytes(backward_out, campaign_original, master_original, media_frames, progress=progress)
            elapsed = time.perf_counter() - started_at
            progress.empty()

        st.success(f"完成しました！ 処理時間：{elapsed:.1f}秒")
        st.balloons()
        st.write(" / ".join(processed) if processed else "有効な媒体ローデータは0行でした。")
        if skipped:
            with st.expander("スキップしたシート"):
                st.write("\n".join(skipped))
        st.download_button(
            "後方数値分析用(Display).xlsx をダウンロード",
            data=excel_bytes,
            file_name="後方数値分析用(Display).xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )
    except Exception as e:
        st.error(str(e))
        st.exception(e)
