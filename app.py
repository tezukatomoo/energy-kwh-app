import io
import re
import unicodedata
from typing import Dict, Optional

import streamlit as st
import pdfplumber
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side


# =========================================================
# 簡易パスワード認証（試用用）
# =========================================================
PASSWORD = "energy2025"  # 必要に応じて変更

if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    st.title("ログイン")
    pw = st.text_input("パスワード", type="password")
    if st.button("ログイン"):
        if pw == PASSWORD:
            st.session_state.authenticated = True
            st.rerun()
        else:
            st.error("パスワードが違います")
    st.stop()


# =========================================================
# タイプキー抽出（★潰さない版）
# =========================================================
def extract_type_key_from_filename(name: str) -> str:
    """
    PDFファイル名からタイプキーを抽出
    例:
      A(2F).pdf   → A(2F)
      A'(1F).pdf  → A'(1F)
    ※ 先頭1文字に潰さない
    """
    s = unicodedata.normalize("NFKC", name).strip()
    s = s.replace("／", "/")

    if "/" in s:
        s = s.split("/")[-1]

    if s.lower().endswith(".pdf"):
        s = s[:-4]

    return s.strip()


def extract_type_key_from_label(label: str) -> str:
    """
    住戸リストの「住宅タイプの名称」からタイプキーを抽出
    例:
      （仮称〇〇）/A(2F) → A(2F)
      A(3F)             → A(3F)
    """
    s = unicodedata.normalize("NFKC", str(label)).strip()
    s = s.replace("／", "/")

    if "/" in s:
        s = s.split("/")[-1]

    return s.strip()


# =========================================================
# PDFから消費電力量[kWh] *1 を抽出
# =========================================================
def extract_kwh_from_pdf_bytes(pdf_bytes: bytes) -> Optional[int]:
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            page = pdf.pages[-1]  # 6/6ページ想定
            raw = page.extract_text() or ""
    except Exception:
        return None

    raw = unicodedata.normalize("NFKC", raw).replace("ｋＷｈ", "kWh")
    lines = [ln.strip() for ln in raw.splitlines() if ln.strip()]

    for i, ln in enumerate(lines):
        if "消費電力量" in ln and "kWh" in ln:
            for j in range(1, 4):
                if i + j < len(lines):
                    m = re.search(r"([0-9]{3,}(?:,[0-9]{3})*)", lines[i + j])
                    if m:
                        return int(m.group(1).replace(",", ""))
            m = re.search(r"([0-9]{3,}(?:,[0-9]{3})*)", ln)
            if m:
                return int(m.group(1).replace(",", ""))

    return None


# =========================================================
# 住戸リストCSVの列検出
# =========================================================
def detect_unitlist_columns(df: pd.DataFrame):
    col_row = next(c for c in df.columns if "行" in c)
    col_num = next(c for c in df.columns if ("住戸" in c and "番号" in c))
    candidates = [
        c for c in df.columns
        if ("住宅タイプ" in c) or ("タイプ" in c and "名称" in c)
    ]
    if not candidates:
        raise RuntimeError("『住宅タイプの名称』列が見つかりません")
    return col_row, col_num, candidates[0]


# =========================================================
# Excel（標準形）作成
# =========================================================
def build_standard_excel(unit_list: pd.DataFrame, project_name: str) -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "集計"

    thin = Side(border_style="thin", color="999999")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    header_fill = PatternFill("solid", fgColor="E6F2FF")
    total_fill = PatternFill("solid", fgColor="FFF2CC")
    title_fill = PatternFill("solid", fgColor="D9EAD3")
    bold = Font(bold=True)
    center = Alignment(horizontal="center")
    right = Alignment(horizontal="right")

    # 物件名
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=10)
    t = ws.cell(row=1, column=1)
    t.value = project_name
    t.font = Font(bold=True, size=14)
    t.alignment = center
    t.fill = title_fill

    # 左ヘッダ
    left_headers = ["行番号", "住戸の番号", "タイプ", "消費電力量[kWh]"]
    for c, h in enumerate(left_headers, start=1):
        cell = ws.cell(row=2, column=c, value=h)
        cell.font = bold
        cell.fill = header_fill
        cell.alignment = center
        cell.border = border

    # 左データ
    for i, row in unit_list.iterrows():
        r = i + 3
        ws.cell(row=r, column=1, value=row["行番号"]).border = border
        ws.cell(row=r, column=2, value=row["住戸の番号"]).border = border
        ws.cell(row=r, column=3, value=row["タイプ"]).border = border
        ws.cell(row=r, column=4, value=row["消費電力量[kWh]"]).border = border

        ws.cell(row=r, column=1).alignment = center
        ws.cell(row=r, column=2).alignment = right
        ws.cell(row=r, column=3).alignment = center
        ws.cell(row=r, column=4).alignment = right

    # 左合計
    total_units = int(unit_list["住戸の番号"].nunique())
    total_kwh = int(unit_list["消費電力量[kWh]"].sum())
    sum_row = len(unit_list) + 3

    ws.cell(row=sum_row, column=1, value="合計住戸数").fill = total_fill
    ws.cell(row=sum_row, column=2, value=total_units).fill = total_fill
    ws.cell(row=sum_row, column=3, value="合計消費電力量[kWh]").fill = total_fill
    ws.cell(row=sum_row, column=4, value=total_kwh).fill = total_fill

    for c in range(1, 5):
        ws.cell(row=sum_row, column=c).font = bold
        ws.cell(row=sum_row, column=c).border = border

    # タイプ別集計
    ts = (
        unit_list
        .groupby("タイプ", as_index=False)
        .agg(
            戸数=("住戸の番号", "count"),
            合計消費電力量_kWh=("消費電力量[kWh]", "sum"),
        )
    )
    ts["kwh_per_unit"] = (ts["合計消費電力量_kWh"] / ts["戸数"]).round(0).astype(int)

    # 右ヘッダ
    right_headers = [
        "タイプ", "戸数",
        "1住戸あたり消費電力量[kWh]",
        "合計消費電力量[kWh]"
    ]
    for c, h in enumerate(right_headers, start=6):
        cell = ws.cell(row=2, column=c, value=h)
        cell.font = bold
        cell.fill = header_fill
        cell.alignment = center
        cell.border = border

    # 右データ
    r0 = 3
    for i, row in ts.sort_values("タイプ").iterrows():
        ws.cell(row=r0, column=6, value=row["タイプ"]).border = border
        ws.cell(row=r0, column=7, value=int(row["戸数"])).border = border
        ws.cell(row=r0, column=8, value=int(row["kwh_per_unit"])).border = border
        ws.cell(row=r0, column=9, value=int(row["合計消費電力量_kWh"])).border = border

        for c in range(6, 10):
            ws.cell(row=r0, column=c).alignment = right if c >= 7 else center
        r0 += 1
   # 右：合計（タイプ別集計の下に表示）
    sum_units = int(ts["戸数"].sum())
    sum_kwh = int(ts["合計消費電力量_kWh"].sum())

    # 1行空けて見やすくする
    r0 += 1

    # 合計住戸数
    ws.cell(row=r0, column=6, value="合計住戸数").fill = total_fill
    ws.cell(row=r0, column=7, value=sum_units).fill = total_fill
    ws.cell(row=r0, column=6).font = bold
    ws.cell(row=r0, column=7).font = bold
    ws.cell(row=r0, column=6).border = border
    ws.cell(row=r0, column=7).border = border
    ws.cell(row=r0, column=6).alignment = center
    ws.cell(row=r0, column=7).alignment = right

    # 合計消費電力量
    r0 += 1
    ws.cell(row=r0, column=6, value="合計消費電力量[kWh]").fill = total_fill
    ws.cell(row=r0, column=7, value=sum_kwh).fill = total_fill
    ws.cell(row=r0, column=6).font = bold
    ws.cell(row=r0, column=7).font = bold
    ws.cell(row=r0, column=6).border = border
    ws.cell(row=r0, column=7).border = border
    ws.cell(row=r0, column=6).alignment = center
    ws.cell(row=r0, column=7).alignment = right

    # 列幅
    ws.column_dimensions["A"].width = 10
    ws.column_dimensions["B"].width = 15
    ws.column_dimensions["C"].width = 12
    ws.column_dimensions["D"].width = 20
    ws.column_dimensions["F"].width = 12
    ws.column_dimensions["G"].width = 10
    ws.column_dimensions["H"].width = 26
    ws.column_dimensions["I"].width = 22

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


# =========================================================
# Streamlit UI
# =========================================================
def main():
    st.title("東京都環境計画書　専用部 消費電力量集計ツール")

    project_name = st.text_input(
        "物件名",
        value="（仮称）〇〇計画 新築工事"
    )

    csv_file = st.file_uploader(
        "住戸リストCSV",
        type=["csv"]
    )

    pdf_files = st.file_uploader(
        "タイプ別PDF（複数選択）",
        type=["pdf"],
        accept_multiple_files=True
    )

    if st.button("集計実行"):
        if not csv_file or not pdf_files:
            st.error("CSVとPDFを両方アップロードしてください")
            return

        # PDF → タイプ別kWh
        type_kwh: Dict[str, Optional[int]] = {}
        rows = []

        for f in pdf_files:
            kwh = extract_kwh_from_pdf_bytes(f.read())
            tkey = extract_type_key_from_filename(f.name)
            rows.append({"PDF名": f.name, "タイプ": tkey, "kWh": kwh})
            type_kwh[tkey] = kwh

        st.subheader("PDF抽出結果")
        st.dataframe(pd.DataFrame(rows))

        # CSV読み込み
        for enc in ("utf-8-sig", "cp932", "utf-8"):
            try:
                units = pd.read_csv(csv_file, encoding=enc)
                break
            except Exception:
                units = None

        if units is None:
            st.error("CSVを読み込めませんでした")
            return

        col_row, col_num, col_type = detect_unitlist_columns(units)

        units["タイプ"] = units[col_type].apply(extract_type_key_from_label)
        units["消費電力量[kWh]"] = units["タイプ"].map(type_kwh)

        unit_list = units[[col_row, col_num, "タイプ", "消費電力量[kWh]"]]
        unit_list.columns = ["行番号", "住戸の番号", "タイプ", "消費電力量[kWh]"]

        st.subheader("住戸別マッピング（先頭50行）")
        st.dataframe(unit_list.head(50))

        missing = unit_list[unit_list["消費電力量[kWh]"].isna()]
        if not missing.empty:
            st.warning("kWhが取得できていないタイプがあります")
            st.dataframe(missing["タイプ"].value_counts())

        excel = build_standard_excel(unit_list, project_name)
        st.download_button(
            "Excelダウンロード",
            data=excel,
            file_name=f"{project_name}_消費電力量集計.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


if __name__ == "__main__":
    main()
