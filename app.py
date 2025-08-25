import streamlit as st
import pandas as pd
import tempfile
import os

# 여기 함수들은 card2approval_multisheet.py에서 그대로 가져오세요
from card2approval import read_raw, load_mapping, build_multi_sheet

MAPPING_FILE = "card_employee_mapping.csv"  # 고정된 매핑 파일

st.title("법인카드 결재용 포맷 변환기")

uploaded_file = st.file_uploader("원본 카드내역 파일 업로드 (.xlsx/.xls)")

month = st.text_input("월(예: 7월)", "")

if uploaded_file is not None:
    if st.button("변환 실행"):
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_in:
            tmp_in.write(uploaded_file.getbuffer())
            tmp_in_path = tmp_in.name

        out_path = "결과.xlsx"

        # 변환 실행
        df_raw = read_raw(tmp_in_path, 0)  # 0 = 첫 번째 시트
        mapping = load_mapping(MAPPING_FILE)
        build_multi_sheet(df_raw, mapping, month, out_path)

        # 결과 다운로드 버튼
        with open(out_path, "rb") as f:
            st.success("변환 완료!")
            st.download_button(
                label="결과 파일 다운로드",
                data=f,
                file_name=out_path,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        # 임시파일 정리
        os.remove(tmp_in_path)