import streamlit as st
import pandas as pd

st.title("헌금 이름 정렬 프로그램")

uploaded_file = st.file_uploader("엑셀 파일 업로드", type=["xlsx", "xls"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    output_text = ""

    for column in df.columns:
        title = str(column)
        col_data = df[column].dropna().astype(str)

        sorted_col = sorted(col_data)

        output_text += f"[{title}]\n"

        for i in range(0, len(sorted_col), 17):
            line = " ".join(sorted_col[i:i+17])
            output_text += line + "\n"

        output_text += "\n"

    st.text_area("결과", output_text, height=400)

    st.download_button(
        label="텍스트 파일 다운로드",
        data=output_text,
        file_name="정리결과.txt"
    )