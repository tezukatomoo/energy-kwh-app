import io
import re
import unicodedata
from typing import Dict, Optional

import pdfplumber
import pandas as pd
import streamlit as st
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side


# ---------- PDFから「消費電力量[kWh] *1」抽出 ----------

def extract_kwh_from_pdf_bytes(pdf_bytes: bytes) -> Optional[int]:
    """
    PDF（バイト列）の最終ページから
    「(1) 設計二次エネルギー消費量等（参考値）」中の
    「消費電力量[kWh] *1」のすぐ下の数値を抜き出す
    """
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            page = pdf.pages[-1]
            raw = page.extract_text() or ""
    except Exception as e:
        st.warning(f"PDF読み込み失敗: {e}")
        return None

    # 全角→半角、ｋＷｈ→kWh 正規化
    raw = unicodedata.normalize("NFKC", raw).replace("ｋＷｈ", "kWh")
    lines = [ln.strip() for ln in raw.splitlines() if ln.strip()]

    for i, ln in enumerate(lines):
        if "消費電力量" in ln and "kWh" in ln:
            # 「すぐ下」の数行を優先
            for j in range(1, 4):
                if i + j < len(lines):
                    m = re.search(r"([0-9]{3,}(?:,[0-9]{3})*)", lines[i + j])
                    if m:
                        return int(m.group(1).replace(",", ""))
            # 見つからなければ同じ行も見る
            m = re.search(r"([0-9]{3,}(?:,[0-9]{3})*)", ln)
            if m:
                return int(m.group(1).replace(",", ""))

    return None


# ---------- タイプキー抽出 ----------

def extract_type_key_from_filename(name: str) -> str:
    """
    'A(2F).pdf', "A'(1F).pdf" などのファイル名からタイプキー 'A', "A'" を取り出す
    """
    s = unicodedata.normalize("NFKC", name).strip()
    s = s.replace("／", "/")
    if "/" in s:
        s = s.split("/")[-1]  # 末尾側だけ

    # 先頭のアルファベット＋任意の ' をタイプキーとする
    m = re.match(r"([A-Za-zＡ-Ｚ]+'?)", s)
    if m:
        return unicodedata.normalize("NFKC", m.group(1))
    return s


def extract_type_key_from_label(label: str) -> str:
    """
    住戸リストの「住宅タイプの名称」からタイプキーを抽出
    例： '（仮称〇〇）/A(2F)' → 'A'
    """
    s = unicodedata.normalize("NFKC", str(label)).strip()
    s = s.replace("／", "/")
    if "/" in s:
        s = s.split("/")[-1]
    m = re.match(r"([A-Za-zＡ-Ｚ]+'?)", s)
    if m:
        return unicodedata.normalize("NFKC", m.group(1))
    return s


# ---------- 住戸リストCSVの列検出 ----------

def detect_unitlist_columns(df: pd.DataFrame):
    col_row = next(c for c in df.columns if "行" in c)
    col_num = next(c for c in df.columns if ("住戸" in c and "番号" in c))
    candidates = [
        c for c in df.columns
        if ("住宅タイプ" in c) or ("タイプ" in c and "名称" in c)
    ]
    if not candidates:
        raise RuntimeError("『住宅タイプの名称』に相当する列が見つかりません。")
    col_type = candidates[0]
    return col_row, col_num, col_type


# ---------- Excel作成（標準形レイアウト） ----------

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

    # タイトル行（1行目）
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=10)
    tcell = ws.cell(row=1, column=1)
    tcell.value = project_name
    tcell.font = Font(bold=True, size=14)
    tcell.alignment = center
    tcell.fill = title_fill

    # 左：住戸別ヘッダ
    left_headers = ["行番号", "住戸の番号", "タイプ", "消費電力量[kWh]"]
    for c, h in enumerate(left_headers, start=1):
        cell = ws.cell(row=2, column=c, value=h)
        cell.font = bold
        cell.fill = header_fill
        cell.alignment = center
        cell.border = border

    # 左：データ
    for i, row in unit_list.iterrows():
        r = i + 3
        ws.cell(row=r, column=1, value=row["行番号"]).border = border
        ws.cell(row=r, column=2, value=row["住戸の番号"]).border = border
        ws.cell(row=r, column=3, value=row["タイプ"]).border = border
        v = row["消費電力量[kWh]"]
        ws.cell(row=r, column=4, value=int(v) if pd.notnull(v) else None).border = border

        ws.cell(row=r, column=1).alignment = center
        ws.cell(row=r, column=2).alignment = right
        ws.cell(row=r, column=3).alignment = center
        ws.cell(row=r, column=4).alignment = right

    # 左：合計行
    total_units = int(unit_list["住戸の番号"].nunique())
    total_kwh = int(unit_list["消費電力量[kWh]"].sum())
    sum_row = len(unit_list) + 3
    labels = ["合計住戸数", total_units, "合計消費電力量[kWh]", total_kwh]
    for c, val in enumerate(labels, start=1):
        cell = ws.cell(row=sum_row, column=c, value=val)
        cell.fill = total_fill
        cell.font = bold
        cell.border = border
        cell.alignment = right if c in (2, 4) else center

    # タイプ別集計
    type_summary = (
        unit_list
        .groupby("タイプ", as_index=False)
        .agg(戸数=("住戸の番号", "count"),
             合計消費電力量_kWh=("消費電力量[kWh]", "sum"))
    )
    type_summary["1住戸あたり消費電力量_kWh"] = (
        type_summary["合計消費電力量_kWh"] / type_summary["戸数"]
    ).round(0).astype(int)

    # Python内部では英字の列名に変換しておく（属性アクセス用）
    type_summary = type_summary.rename(
        columns={"1住戸あたり消費電力量_kWh": "kwh_per_unit"}
    )

    # 右：ヘッダ
    right_headers = ["タイプ", "戸数",
                     "1住戸あたり消費電力量[kWh]",
                     "合計消費電力量[kWh]"]
    for c, h in enumerate(right_headers, start=6):
        cell = ws.cell(row=2, column=c, value=h)
        cell.font = bold
        cell.fill = header_fill
        cell.alignment = center
        cell.border = border

    # 右：データ
    ts_sorted = type_summary.sort_values("タイプ")
    start_row = 3
    for idx, row in enumerate(ts_sorted.itertuples(index=False), start=start_row):
        # row は (タイプ, 戸数, 合計消費電力量_kWh, kwh_per_unit) というタプル
        ws.cell(row=idx, column=6, value=row.タイプ).border = border
        ws.cell(row=idx, column=7, value=int(row.戸数)).border = border
        ws.cell(row=idx, column=8, value=int(row.kwh_per_unit)).border = border
        ws.cell(row=idx, column=9, value=int(row.合計消費電力量_kWh)).border = border
        for c in range(6, 10):
            ws.cell(row=idx, column=c).alignment = right if c >= 7 else center

    summary_row = len(ts_sorted) + start_row + 1
    cell1 = ws.cell(row=summary_row, column=6, value="合計住戸数")
    cell2 = ws.cell(row=summary_row, column=7, value=total_units)
    cell3 = ws.cell(row=summary_row + 1, column=6, value="合計消費電力量[kWh]")
    cell4 = ws.cell(row=summary_row + 1, column=7, value=total_kwh)
    for cell in (cell1, cell2, cell3, cell4):
        cell.fill = total_fill
        cell.font = bold
        cell.border = border
    cell1.alignment = center
    cell3.alignment = center
    cell2.alignment = right
    cell4.alignment = right

    # 列幅
    ws.column_dimensions["A"].width = 10
    ws.column_dimensions["B"].width = 15
    ws.column_dimensions["C"].width = 10
    ws.column_dimensions["D"].width = 20
    ws.column_dimensions["F"].width = 10
    ws.column_dimensions["G"].width = 10
    ws.column_dimensions["H"].width = 26
    ws.column_dimensions["I"].width = 22

    # バイト列として返す
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ---------- Streamlit UI ----------

def main():
    st.title("東京都環境計画書：専用部 消費電力量集計ツール")
    st.caption("6/6ページ『設計二次エネルギー消費量等（参考値）』の『消費電力量[kWh] *1』のみを使って集計します。")

    project_name = st.text_input(
        "物件名（Excelタイトル行に表示）",
        value="（仮称）〇〇計画 新築工事",
    )

    st.subheader("① 住戸リストCSVをアップロード")
    csv_file = st.file_uploader(
        "住戸リストCSVファイル（例：〇〇_住戸リスト.csv）",
        type=["csv"],
        key="unit_csv",
    )

    st.subheader("② タイプ別PDFをアップロード")
    pdf_files = st.file_uploader(
        "タイプ別PDF（A(2F).pdf, B(2F).pdf など）を複数選択してください",
        type=["pdf"],
        accept_multiple_files=True,
        key="pdfs",
    )

    if st.button("③ 集計実行"):
        if not csv_file or not pdf_files:
            st.error("CSV と PDF を両方アップロードしてください。")
            return

        # --- PDF→タイプ別kWh ---
        st.write("PDFからタイプ別の『消費電力量[kWh] *1』を抽出中…")
        type_kwh: Dict[str, Optional[int]] = {}
        pdf_summary_rows = []

        for f in pdf_files:
            pdf_bytes = f.read()
            fname = f.name
            type_key = extract_type_key_from_filename(fname)
            kwh_val = extract_kwh_from_pdf_bytes(pdf_bytes)
            pdf_summary_rows.append(
                {"PDF名": fname, "タイプキー": type_key, "抽出kWh": kwh_val}
            )
            if type_key not in type_kwh or type_kwh[type_key] is None:
                type_kwh[type_key] = kwh_val

        pdf_summary_df = pd.DataFrame(pdf_summary_rows).sort_values("PDF名")
        st.write("PDFごとの抽出結果：")
        st.dataframe(pdf_summary_df)

        st.write("タイプ別kWhマップ：")
        st.dataframe(
            pd.DataFrame(
                [{"タイプキー": k, "kWh": v} for k, v in sorted(type_kwh.items())]
            )
        )

        # --- CSV読み込み ---
        st.write("住戸リストCSVを読み込み中…")
        units = None
        for enc in ("utf-8-sig", "cp932", "utf-8"):
            try:
                units = pd.read_csv(csv_file, encoding=enc)
                st.write(f"CSV読み込み成功（encoding={enc}）")
                break
            except Exception:
                continue
        if units is None:
            st.error("住戸リストCSVの読み込みに失敗しました。")
            return

        try:
            col_row, col_num, col_type = detect_unitlist_columns(units)
        except Exception as e:
            st.error(f"住戸リストの列特定に失敗しました: {e}")
            st.write("CSVの列名：", list(units.columns))
            return

        st.write(f"列対応: 行番号 = {col_row}, 住戸の番号 = {col_num}, タイプ名称 = {col_type}")

        # タイプキー列を追加
        units["タイプキー"] = units[col_type].apply(extract_type_key_from_label)
        units["消費電力量[kWh]"] = units["タイプキー"].map(type_kwh)

        unit_list = units[[col_row, col_num, "タイプキー", "消費電力量[kWh]"]].copy()
        unit_list.columns = ["行番号", "住戸の番号", "タイプ", "消費電力量[kWh]"]

        st.write("住戸別マッピング結果（先頭50行）：")
        st.dataframe(unit_list.head(50))

        # マッピング漏れチェック
        missing = unit_list[unit_list["消費電力量[kWh]"].isna()]["タイプ"].value_counts()
        if not missing.empty:
            st.warning("消費電力量が付与できていないタイプがあります：")
            st.dataframe(missing.rename("件数").to_frame())
        else:
            st.success("全ての住戸に消費電力量[kWh]がマッピングされました。")

        # --- Excel生成 ---
        st.write("Excel（標準形レイアウト）を生成中…")
        excel_bytes = build_standard_excel(unit_list, project_name)

        out_name = f"{project_name}_消費電力量集計.xlsx"
        st.success("Excel生成完了。以下からダウンロードできます。")
        st.download_button(
            label="Excelダウンロード",
            data=excel_bytes,
            file_name=out_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


if __name__ == "__main__":
    main()
