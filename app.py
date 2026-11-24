import os
import tempfile
import streamlit as st
import whisper
from docx import Document

st.set_page_config(page_title="MP3 語音轉 Word 逐字稿", page_icon="🎙️", layout="centered")
st.title("🎙️ MP3 語音轉 Word 逐字稿")
st.write("上傳 MP3 語音檔，使用 Whisper 轉換為文字並匯出 `.docx` 下載。")


@st.cache_resource(show_spinner="正在載入 Whisper 模型，首次載入需稍候...")
def load_whisper_model(model_size: str):
    return whisper.load_model(model_size)


def build_docx(result: dict, source_name: str, include_timestamps: bool) -> bytes:
    doc = Document()
    doc.add_heading("語音轉文字逐字稿", level=1)
    doc.add_paragraph(f"來源檔案：{source_name}")
    doc.add_paragraph("")

    if include_timestamps and "segments" in result:
        for seg in result["segments"]:
            start_sec = int(seg["start"])
            end_sec = int(seg["end"])
            start_str = f"{start_sec // 60:02d}:{start_sec % 60:02d}"
            end_str = f"{end_sec // 60:02d}:{end_sec % 60:02d}"
            text = seg["text"].strip()
            if text:
                doc.add_paragraph(f"[{start_str} - {end_str}] {text}")
    else:
        full_text = result.get("text", "").strip()
        doc.add_paragraph(full_text)

    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_docx:
        doc.save(tmp_docx.name)
        tmp_docx_path = tmp_docx.name

    try:
        with open(tmp_docx_path, "rb") as f:
            return f.read()
    finally:
        if os.path.exists(tmp_docx_path):
            os.remove(tmp_docx_path)


uploaded_file = st.file_uploader("選擇 MP3 檔案", type=["mp3"])

col1, col2 = st.columns(2)
with col1:
    # 免費版 Streamlit Cloud 資源有限，建議優先選 tiny / base / small
    model_size = st.selectbox("Whisper 模型大小", ["tiny", "base", "small", "medium"], index=1)
with col2:
    language_option = st.selectbox("語音語言", ["中文 (zh)", "自動偵測", "英文 (en)", "日文 (ja)"])

include_timestamps = st.checkbox("在 Word 中附上時間戳記", value=True)

language_map = {
    "中文 (zh)": "zh",
    "自動偵測": None,
    "英文 (en)": "en",
    "日文 (ja)": "ja",
}
selected_language = language_map[language_option]

if uploaded_file is not None:
    st.audio(uploaded_file, format="audio/mp3")

    if st.button("開始轉換為 Word", type="primary"):
        tmp_mp3_path = None
        try:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".mp3") as tmp_mp3:
                tmp_mp3.write(uploaded_file.getvalue())
                tmp_mp3_path = tmp_mp3.name

            model = load_whisper_model(model_size)

            with st.spinner("正在進行語音轉文字，請稍候..."):
                result = model.transcribe(tmp_mp3_path, language=selected_language)

            docx_bytes = build_docx(result, uploaded_file.name, include_timestamps)
            base_name = os.path.splitext(uploaded_file.name)[0]
            output_filename = f"{base_name}_逐字稿.docx"

            st.success("轉換完成！")

            preview_text = "\n".join(
                f"[{int(s['start']) // 60:02d}:{int(s['start']) % 60:02d} - {int(s['end']) // 60:02d}:{int(s['end']) % 60:02d}] {s['text'].strip()}"
                for s in result.get("segments", [])
                if s.get("text", "").strip()
            ) if include_timestamps else result.get("text", "").strip()

            st.text_area("逐字稿預覽", value=preview_text, height=220)

            st.download_button(
                label="⬇️ 下載 Word 檔 (.docx)",
                data=docx_bytes,
                file_name=output_filename,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )
        except Exception as e:
            st.error(f"轉換過程發生錯誤：{e}")
        finally:
            if tmp_mp3_path and os.path.exists(tmp_mp3_path):
                os.remove(tmp_mp3_path)
