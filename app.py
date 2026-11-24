import os
import re
import tempfile

import streamlit as st
import whisper
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

# ==================== 頁面設定 ====================

st.set_page_config(
    page_title="MP3 語音轉 Word 逐字稿",
    page_icon="🎧",
    layout="centered"
)

st.title("🎧 MP3 語音轉 Word 逐字稿")
st.write(
    "上傳 MP3 檔案，使用 OpenAI Whisper 轉錄為文字，"
    "並自動排版後下載 Word (.docx) 檔。"
)

# ==================== 工具函式 ====================

def clean_text(text: str) -> str:
    """清理 Whisper 常見的多餘空白與重複標點。"""
    text = text.replace("\n", " ")
    text = re.sub(r"\s+", " ", text).strip()

    # 移除標點前多餘空格
    text = re.sub(r"\s+([，。、！？；：」』）])", r"\1", text)

    # 移除重複標點
    text = re.sub(r"([。！？])\1+", r"\1", text)
    text = re.sub(r"([，、；：])\1+", r"\1", text)

    return text.strip()


def split_into_sentences(text: str) -> list[str]:
    """依中文標點切句。"""
    text = clean_text(text)

    sentences = re.split(r"(?<=[。！？；])\s*", text)
    sentences = [s.strip() for s in sentences if s.strip()]

    return sentences


def group_sentences(
    sentences: list[str],
    sentences_per_paragraph: int = 3
) -> list[str]:
    """將句子合併為段落。"""
    paragraphs = []

    for i in range(0, len(sentences), sentences_per_paragraph):
        group = sentences[i:i + sentences_per_paragraph]
        paragraphs.append("".join(group))

    return paragraphs


def create_word_document(
    audio_filename: str,
    language: str,
    model_name: str,
    transcript: str,
    sentences_per_paragraph: int
) -> str:
    """建立排版後的 Word 檔，並回傳檔案路徑。"""

    transcript = clean_text(transcript)
    sentences = split_into_sentences(transcript)
    formatted_paragraphs = group_sentences(
        sentences,
        sentences_per_paragraph=sentences_per_paragraph
    )

    base_name = os.path.splitext(audio_filename)[0]
    word_filename = f"{base_name}_逐字稿_排版.docx"

    document = Document()

    # 頁面邊界
    for section in document.sections:
        section.top_margin = Cm(2.0)
        section.bottom_margin = Cm(2.0)
        section.left_margin = Cm(2.5)
        section.right_margin = Cm(2.5)

    # 內文字型與行距
    normal_style = document.styles["Normal"]
    normal_style.font.name = "Calibri"
    normal_style.font.size = Pt(12)

    paragraph_format = normal_style.paragraph_format
    paragraph_format.line_spacing = 1.5
    paragraph_format.space_after = Pt(8)

    # 主標題
    title = document.add_heading("語音逐字稿", level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # 基本資料
    info = document.add_paragraph()

    info_run = info.add_run(f"來源音檔：{audio_filename}\n")
    info_run.font.size = Pt(10)
    info_run.font.color.rgb = RGBColor(89, 89, 89)

    info_run = info.add_run(
        f"辨識語言：{language if language else '自動偵測'}\n"
    )
    info_run.font.size = Pt(10)
    info_run.font.color.rgb = RGBColor(89, 89, 89)

    info_run = info.add_run(f"使用模型：{model_name}")
    info_run.font.size = Pt(10)
    info_run.font.color.rgb = RGBColor(89, 89, 89)

    # 完整文字
    document.add_heading("完整文字", level=1)

    if formatted_paragraphs:
        for paragraph_text in formatted_paragraphs:
            p = document.add_paragraph(paragraph_text)
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    else:
        document.add_paragraph("沒有偵測到可轉錄的文字。")

    document.save(word_filename)

    return word_filename

# ==================== 介面設定 ====================

with st.sidebar:
    st.header("轉錄設定")

    language = st.selectbox(
        "辨識語言",
        options=["zh", "en", "自動偵測"],
        index=0,
        help="選擇 MP3 的主要語言；若不確定，可選擇「自動偵測」。"
    )

    if language == "自動偵測":
        language = None

    model_name = st.selectbox(
        "Whisper 模型",
        options=["tiny", "base", "small", "medium", "large-v3", "turbo"],
        index=2,
        help="模型越大通常越準確，但需要更多記憶體與時間。"
    )

    sentences_per_paragraph = st.slider(
        "每段包含句數",
        min_value=1,
        max_value=8,
        value=3,
        help="設定 Word 檔中每個段落包含多少句子。"
    )

st.divider()

# ==================== 上傳檔案 ====================

uploaded_file = st.file_uploader(
    "請上傳 MP3 檔案",
    type=["mp3"],
    accept_multiple_files=False
)

if uploaded_file is None:
    st.info("請先上傳 MP3 檔案。")
    st.stop()

st.success(f"已選擇檔案：{uploaded_file.name}")
st.write(f"檔案大小：{uploaded_file.size / 1024 / 1024:.2f} MB")

# ==================== 開始轉錄 ====================

if st.button("開始轉錄並產生 Word 檔", type="primary"):
    if not uploaded_file.name.lower().endswith(".mp3"):
        st.error("請上傳 .mp3 格式的音訊檔。")
        st.stop()

    # 將上傳檔案儲存為暫存檔
    with tempfile.NamedTemporaryFile(
        delete=False,
        suffix=".mp3"
    ) as temp_audio:
        temp_audio.write(uploaded_file.getvalue())
        audio_path = temp_audio.name

    try:
        with st.status("正在處理語音檔 ...", expanded=True) as status:

            st.write(f"正在載入 Whisper 模型：{model_name}")
            model = whisper.load_model(model_name)

            st.write("正在轉錄語音，請稍候 ...")

            transcribe_kwargs = {
                "verbose": False,
                "fp16": False
            }

            if language:
                transcribe_kwargs["language"] = language

            result = model.transcribe(audio_path, **transcribe_kwargs)

            transcript = result.get("text", "").strip()

            if not transcript:
                status.update(
                    label="轉錄失敗：沒有偵測到文字。",
                    state="error"
                )
                st.error("沒有偵測到可轉錄的文字，請確認音檔內容或改用其他模型。")
                st.stop()

            st.write("正在建立排版後的 Word 檔 ...")

            word_path = create_word_document(
                audio_filename=uploaded_file.name,
                language=language,
                model_name=model_name,
                transcript=transcript,
                sentences_per_paragraph=sentences_per_paragraph
            )

            status.update(
                label="轉錄與 Word 檔建立完成！",
                state="complete"
            )

        # 顯示轉錄結果
        st.subheader("轉錄結果")

        with st.expander("查看完整文字", expanded=True):
            st.write(transcript)

        # 提供下載
        with open(word_path, "rb") as word_file:
            word_bytes = word_file.read()

        st.download_button(
            label="⬇️ 下載 Word 檔",
            data=word_bytes,
            file_name=os.path.basename(word_path),
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "wordprocessingml.document"
            ),
            type="primary"
        )

    except Exception as e:
        st.error(f"處理時發生錯誤：{e}")

    finally:
        # 刪除暫存 MP3 與 Word 檔
        if os.path.exists(audio_path):
            os.remove(audio_path)

        if "word_path" in locals() and os.path.exists(word_path):
            os.remove(word_path)
