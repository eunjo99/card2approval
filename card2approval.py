# -*- coding: utf-8 -*-
"""
카드내역 원본 엑셀 -> 회사 포맷(멀티 시트) + 보고용 테이블 생성 스크립트
--------------------------------------------------------------------
필요 패키지: pandas, openpyxl
설치:
    pip install pandas openpyxl

실행 예시:
    python card2approval_multisheet.py --raw "원본.xlsx" --sheet "Sheet1" \
      --mapping card_employee_mapping.csv --out "결과.xlsx" --month "7월"

생성되는 시트:
- 시트1: '전체'  -> 요구 컬럼만 추출 + 직원/사업장 매핑 포함
- 시트2: '판교'  -> '전체' 중 사업장=판교
- 시트3: '대전'  -> '전체' 중 사업장=대전
- 시트4: '보고용' -> 직원(카드번호)별 테이블(헤더+내역+합계), 블록 간 1행 공백,
                    테두리, 합계(A:C 병합+가운데), 금액 콤마 서식
주의:
- 원본이 .xls 인 경우 일부 환경에서 읽기 오류가 날 수 있으니, 안될 때는 엑셀에서 .xlsx로 저장하여 사용하세요.
"""
import argparse
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, Border, Side, PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows

SOURCE_COLS = [
    "카드번호","승인일자","승인시간","승인금액(원화)","승인금액(외화)",
    "공급가액(원화)","부가세","외화거래일환율","외화거래국가코드",
    "가맹점사업자번호","가맹점명","가맹점업종명"
]

OUTPUT_ORDER = [
    "카드번호","승인일자","승인시간","승인금액(원화)","승인금액(외화)",
    "공급가액(원화)","부가세","외화거래일환율","외화거래국가코드",
    "가맹점사업자번호","가맹점명","가맹점업종명","적요","비용 구분"
]

CURRENCY_COLS = ["승인금액(원화)","승인금액(외화)","공급가액(원화)","부가세"]

def read_raw(path, sheet):
    df = pd.read_excel(path, sheet_name=sheet, dtype=str)
    cols_present = [c for c in SOURCE_COLS if c in df.columns]
    df = df[cols_present].copy()

    # 날짜 포맷
    if "승인일자" in df.columns:
        try:
            df["승인일자"] = pd.to_datetime(df["승인일자"], errors="coerce").dt.strftime("%Y.%m.%d")
        except Exception:
            pass

    # 금액류 숫자화(콤마 제거)
    for col in ["승인금액(원화)","공급가액(원화)","부가세"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col].astype(str).str.replace(",",""), errors="coerce").fillna(0).astype(int)

    return df

def load_mapping(mapping_csv):
    m = pd.read_csv(mapping_csv, dtype=str).fillna("")
    expected = {"card_number_masked","employee_name","title","site"}
    missing = expected - set(m.columns)
    if missing:
        raise ValueError(f"매핑 CSV에 누락된 컬럼: {missing}")
    return m

def build_multi_sheet(df_raw, mapping, month_label, out_path):
    # 매핑 병합
    df = df_raw.merge(mapping, left_on="카드번호", right_on="card_number_masked", how="left")

    # 시트1: 전체
    whole_cols = SOURCE_COLS + ["employee_name","title","site"]
    whole_cols = [c for c in whole_cols if c in df.columns]
    df_whole = df[whole_cols].copy()

    wb = Workbook()

    # 전체
    ws_all = wb.active
    ws_all.title = "전체"
    for r in dataframe_to_rows(df_whole, index=False, header=True):
        ws_all.append(r)

    # 판교/대전 필터 시트
    def add_filtered(name):
        ws = wb.create_sheet(title=name)
        dfx = df_whole[df_whole["site"] == name].copy() if "site" in df_whole.columns else df_whole.iloc[0:0].copy()
        for r in dataframe_to_rows(dfx, index=False, header=True):
            ws.append(r)
        return ws
    ws_pg = add_filtered("판교")
    ws_dj = add_filtered("대전")

     # --- 금액 서식 적용 (전체/판교/대전 시트) ---
    def apply_currency_format(ws):
        if ws.max_row < 2:  # 데이터 없는 경우
            return
        # 헤더(1행) 제외, 2행부터 끝까지
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=ws.max_column):
            for cell in row:
                if cell.value is None:
                    continue
                if isinstance(cell.value, (int, float)):
                    cell.number_format = "#,##0"

    apply_currency_format(ws_all)
    apply_currency_format(ws_pg)
    apply_currency_format(ws_dj)

    # 보고용
    ws = wb.create_sheet(title="보고용")
    title = f"{month_label} 법인카드 이용내역" if month_label else "법인카드 이용내역"
    ws.append([title])
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=14)
    ws["A1"].font = Font(size=14, bold=True)
    ws["A1"].alignment = Alignment(horizontal="center")

    thin = Side(style="thin")
    border_all = Border(left=thin, right=thin, top=thin, bottom=thin)
    header_fill = PatternFill("solid", fgColor="F2F2F2")

    row_cursor = 3
    grand_total = 0

    # 정렬
    sort_cols = [c for c in ["site","employee_name","카드번호","승인일자","승인시간"] if c in df.columns]
    df_sorted = df.sort_values(sort_cols).reset_index(drop=True)

    for site in [s for s in ["판교","대전"] if ("site" in df_sorted.columns and s in df_sorted["site"].unique().tolist())]:
        site_df = df_sorted[df_sorted["site"] == site].copy()
        site_total = 0

        # 사이트 헤더
        ws.cell(row=row_cursor, column=1, value=f"{site} 내역")
        ws.cell(row=row_cursor, column=1).font = Font(bold=True)
        row_cursor += 1

        for (emp, title_k, card_no), g in site_df.groupby(["employee_name","title","카드번호"], dropna=False):
            # 헤더 행
            header = OUTPUT_ORDER
            for ci, val in enumerate(header, start=1):
                c = ws.cell(row=row_cursor, column=ci, value=val)
                c.font = Font(bold=True)
                c.fill = header_fill
                c.border = border_all
                c.alignment = Alignment(horizontal="center")
            header_row = row_cursor
            row_cursor += 1

            # 데이터 행(적요/비용 구분 비움)
            g = g.copy()
            for extra in ["적요","비용 구분"]:
                g[extra] = ""
            g = g[[c for c in OUTPUT_ORDER if c in g.columns]]

            start_data_row = row_cursor
            for _, r in g.iterrows():
                for ci, col in enumerate(OUTPUT_ORDER, start=1):
                    val = r.get(col, "")
                    cell = ws.cell(row=row_cursor, column=ci, value=val)
                    cell.border = border_all
                row_cursor += 1
            end_data_row = row_cursor - 1

            # 숫자 서식 적용(내역)
            for col_letter, col_name in zip(list("ABCDEFGHIJKLMN"), OUTPUT_ORDER):
                if col_name in CURRENCY_COLS:
                    for rr in range(start_data_row, end_data_row+1):
                        ws[f"{col_letter}{rr}"].number_format = "#,##0"

            # 직원 합계
            emp_sum = int(g["승인금액(원화)"].sum()) if "승인금액(원화)" in g.columns else 0
            ws.merge_cells(start_row=row_cursor, start_column=1, end_row=row_cursor, end_column=3)
            total_label = ws.cell(row=row_cursor, column=1, value=f"{emp} {title_k} 합계")
            total_label.font = Font(bold=True)
            total_label.alignment = Alignment(horizontal="center")
            total_amt = ws.cell(row=row_cursor, column=4, value=emp_sum)
            total_amt.font = Font(bold=True)
            total_amt.number_format = "#,##0"

            # 테두리(헤더~합계까지)
            for cc in range(1, 15):
                ws.cell(row=header_row, column=cc).border = border_all
            for rr in range(start_data_row, row_cursor+1):
                for cc in range(1, 15):
                    ws.cell(row=rr, column=cc).border = border_all

            row_cursor += 1  # 합계 줄 끝
            row_cursor += 1  # 블록 간 공백 1행

            site_total += emp_sum

        # 사이트 합계
        ws.merge_cells(start_row=row_cursor, start_column=1, end_row=row_cursor, end_column=3)
        lab = ws.cell(row=row_cursor, column=1, value=f"{site} 합계")
        lab.font = Font(bold=True)
        lab.alignment = Alignment(horizontal="center")
        amt = ws.cell(row=row_cursor, column=4, value=site_total)
        amt.font = Font(bold=True)
        amt.number_format = "#,##0"
        for cc in range(1, 5):
            ws.cell(row=row_cursor, column=cc).border = border_all

        row_cursor += 2  # 사이트 블록 뒤 공백 1행 포함
        grand_total += site_total

    # 전체 합계
    ws.merge_cells(start_row=row_cursor, start_column=1, end_row=row_cursor, end_column=3)
    lab = ws.cell(row=row_cursor, column=1, value="전체 합계")
    lab.font = Font(bold=True)
    lab.alignment = Alignment(horizontal="center")
    amt = ws.cell(row=row_cursor, column=4, value=grand_total)
    amt.font = Font(bold=True)
    amt.number_format = "#,##0"
    for cc in range(1, 5):
        ws.cell(row=row_cursor, column=cc).border = border_all

    # 열 너비
    widths = [18,12,10,14,14,14,10,16,16,18,24,18,18,12]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[chr(64+i)].width = w

    wb.save(out_path)

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--raw", required=True, help="원본 엑셀 경로(.xlsx 권장)")
    ap.add_argument("--sheet", default=0, help="시트명 또는 인덱스")
    ap.add_argument("--mapping", required=True, help="card_employee_mapping.csv 경로")
    ap.add_argument("--out", required=True, help="출력 엑셀 경로(.xlsx)")
    ap.add_argument("--month", default="", help='예: "7월" (제목에 사용)')
    args = ap.parse_args()

    df_raw = read_raw(args.raw, args.sheet)
    mapping = load_mapping(args.mapping)
    build_multi_sheet(df_raw, mapping, args.month, args.out)
    print(f"[OK] 생성: {args.out}")

if __name__ == "__main__":
    main()