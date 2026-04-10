import streamlit as st
import pandas as pd
from io import BytesIO

# 워드 라이브러리
try:
    from docx import Document
    DOCX_AVAILABLE = True
except:
    DOCX_AVAILABLE = False

# 페이지 설정
st.set_page_config(
    page_title="헌금 이름 정렬",
    page_icon="📋",
    layout="wide"
)

# 스타일
st.markdown("""
<style>
.main {
    background: linear-gradient(135deg, #f8f6f2 0%, #fdfcf9 100%);
}
h1 {
    color: #4b3f35;
}
section[data-testid="stSidebar"] {
    background: #f3efe8;
}
.stDownloadButton button {
    border-radius: 10px;
}
</style>
""", unsafe_allow_html=True)

# 사이드바
st.sidebar.image("logo.png", use_container_width=100)
st.sidebar.markdown("---")
user_name = st.sidebar.text_input("사용자", placeholder="대망교회")
st.sidebar.info("엑셀 업로드 후 자동 정렬됩니다")

# 메인
st.title("📋 헌금 이름 정렬 프로그램")

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

    if user_name:
        st.success(f"{user_name}님 결과입니다")

    st.text_area("결과", output_text, height=400)

    # 텍스트 파일
    txt_data = output_text.encode("utf-8-sig")

    col1, col2 = st.columns(2)

    with col1:
        st.download_button(
            label="📄 텍스트 다운로드 (.txt)",
            data=txt_data,
            file_name="헌금정리결과.txt",
            mime="text/plain"
        )

    if DOCX_AVAILABLE:
        doc = Document()

        if user_name:
            doc.add_heading(f"{user_name}님 결과", 0)
        else:
            doc.add_heading("헌금 이름 정렬 결과", 0)

        for line in output_text.split("\n"):
            doc.add_paragraph(line)

        buffer = BytesIO()
        doc.save(buffer)
        buffer.seek(0)

        with col2:
            st.download_button(
                label="📘 워드 다운로드 (.docx)",
                data=buffer,
                file_name="헌금정리결과.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
    else:
        st.warning("⚠️ 워드 다운로드를 사용하려면 python-docx 설치 필요")
