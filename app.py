import streamlit as st
from card2approval import read_raw, load_mapping_from_upload, build_excel_bytes

st.title("법인카드 결재용 포맷 변환기 (EnF)")

# ✅ 원본 업로드는 xlsx만
raw_file = st.file_uploader("원본 카드내역 업로드 (.xlsx)", type=["xlsx"])

# ✅ 매핑 업로드는 CSV 또는 XLSX
map_file = st.file_uploader("임직원 카드번호 내역 업로드 (CSV 또는 XLSX)", type=["csv","xlsx"])

month = st.text_input("월(예: 7월)", "")

if raw_file and map_file:
    if st.button("변환 실행"):
        try:
            df_raw = read_raw(raw_file, 0)
            mapping = load_mapping_from_upload(map_file)

            out_bytes = build_excel_bytes(df_raw, mapping, month)

            # ✅ 결과 파일명: "{월} 법인카드 이용내역.xlsx"
            month_clean = (month or "").strip()
            if month_clean == "":
                download_name = "법인카드 이용내역.xlsx"
            else:
                download_name = f"{month_clean} 법인카드 이용내역.xlsx"

            st.success("변환 완료! 아래 버튼으로 다운로드하세요.")
            st.download_button(
                "결과 파일 다운로드",
                data=out_bytes,
                file_name=download_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        except Exception as e:
            st.error(f"에러: {e}")
else:
    st.info("원본(.xlsx)과 매핑(CSV/XLSX)을 모두 업로드해 주세요.")
