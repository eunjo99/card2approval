# app.py 예시
import streamlit as st
from card2approval import read_raw, load_mapping_from_filelike, build_excel_bytes

st.title("법인카드 결재용 포맷 변환기")

raw_file = st.file_uploader("원본 업로드 (.xlsx/.xls)", type=["xlsx","xls"])
map_file = st.file_uploader("매핑 CSV 업로드", type=["csv"])
month = st.text_input("월(예: 7월)", "")

if raw_file and map_file and st.button("변환 실행"):
    df_raw = read_raw(raw_file, 0)
    mapping = load_mapping_from_filelike(map_file)
    out_bytes = build_excel_bytes(df_raw, mapping, month)
    st.download_button("결과 파일 다운로드", out_bytes, "결과.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")