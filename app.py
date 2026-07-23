"""
术语表辅助翻译工具 — Streamlit Web App

Upload a document (.docx / .pdf) and a bilingual glossary (.xlsx),
get an AI translation that strictly follows your glossary.

Translation is powered by the DeepSeek API.
Your glossary stays on the server — users only see the translated output.
"""

import streamlit as st
import tempfile
import os
from pathlib import Path
from openai import OpenAI
from docx import Document
import pdfplumber

# Import the glossary helper (must be in the same directory)
from glossary_helper import load_glossary, protect_terms, restore_terms

# ---------------------------------------------------------------------------
# Page config
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="术语表辅助翻译工具",
    page_icon="🌐",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ---------------------------------------------------------------------------
# Sidebar — settings
# ---------------------------------------------------------------------------
with st.sidebar:
    st.header("⚙️ 设置")

    st.subheader("🔑 API Key")
    api_key = st.text_input(
        "DeepSeek API Key",
        type="password",
        value=os.environ.get("DEEPSEEK_API_KEY", ""),
        placeholder="sk-...",
        help="从环境变量自动读取。也可手动填入自己的 Key。",
    )

    st.divider()

    st.subheader("🌍 翻译方向")
    direction_label = st.radio(
        "选择翻译方向",
        options=["自动检测", "中 → 英", "英 → 中"],
        index=0,
    )
    direction_map = {"自动检测": "auto", "中 → 英": "cn2en", "英 → 中": "en2cn"}
    direction = direction_map[direction_label]

    st.divider()

    st.subheader("🤖 翻译模型")
    model = st.selectbox(
        "选择模型",
        options=[
            "deepseek-v4-flash",
            "deepseek-v4-pro",
        ],
        index=0,
        format_func=lambda m: {
            "deepseek-v4-flash": "V4 Flash（推荐 · 快速便宜）",
            "deepseek-v4-pro": "V4 Pro（旗舰 · 翻译质量最高）",
        }.get(m, m),
        help="Flash 适合日常翻译，Pro 适合正式文档/合同等高质量场景。",
    )

    st.divider()
    st.caption("💡 术语表文件不上传时，将作为普通 AI 翻译运行。")

# ---------------------------------------------------------------------------
# Main area
# ---------------------------------------------------------------------------
st.title("🌐 术语表辅助翻译工具")
st.caption(
    "上传文档和术语表 → AI 在翻译时自动识别并替换术语 → 下载术语一致的译文"
)

# --- File upload area ---
col1, col2 = st.columns(2)

with col1:
    st.subheader("📄 1. 上传文档")
    doc_file = st.file_uploader(
        "支持 .docx 和 .pdf 格式",
        type=["docx", "pdf"],
        help="上传待翻译的文档",
        label_visibility="collapsed",
    )

with col2:
    st.subheader("📖 2. 上传术语表（可选）")
    glossary_file = st.file_uploader(
        "两列 Excel：源语言 | 目标语言（首行为表头）",
        type=["xlsx"],
        help="格式：第一列 = 源语言术语，第二列 = 目标语言翻译。支持 .xlsx。",
        label_visibility="collapsed",
    )

# --- Default glossary ---
DEFAULT_GLOSSARY_PATH = Path(__file__).parent / "glossary.xlsx"
has_default_glossary = DEFAULT_GLOSSARY_PATH.exists()

if has_default_glossary and not glossary_file:
    st.info(f"📌 检测到内置术语表（{DEFAULT_GLOSSARY_PATH.name}），将自动使用。")

# ---------------------------------------------------------------------------
# Translation logic
# ---------------------------------------------------------------------------
def extract_text(file_bytes: bytes, filename: str) -> str:
    """Extract text from .docx or .pdf file bytes."""
    suffix = Path(filename).suffix.lower()
    with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp:
        tmp.write(file_bytes)
        tmp_path = tmp.name

    try:
        if suffix == ".docx":
            doc = Document(tmp_path)
            text = "\n".join([p.text for p in doc.paragraphs])
        elif suffix == ".pdf":
            with pdfplumber.open(tmp_path) as pdf:
                text = "\n".join([page.extract_text() or "" for page in pdf.pages])
        else:
            raise ValueError(f"不支持的文件格式: {suffix}")
    finally:
        os.unlink(tmp_path)

    return text


def chunk_text(text: str, max_chars: int = 6000) -> list[str]:
    """Split text into chunks at paragraph boundaries, each ≤ max_chars."""
    paragraphs = text.split("\n")
    chunks = []
    current = ""

    for para in paragraphs:
        if len(current) + len(para) + 1 <= max_chars:
            current = (current + "\n" + para) if current else para
        else:
            if current:
                chunks.append(current)
            # If a single paragraph exceeds max_chars, split it
            if len(para) > max_chars:
                for i in range(0, len(para), max_chars):
                    chunks.append(para[i:i + max_chars])
                current = ""
            else:
                current = para

    if current:
        chunks.append(current)

    return chunks if chunks else [""]


def translate_text(
    text: str,
    direction: str,
    api_key: str,
    model: str,
    progress_bar,
    status_text,
) -> str:
    """Call DeepSeek API to translate text, preserving ⟨Tn⟩ placeholders."""
    client = OpenAI(
        api_key=api_key,
        base_url="https://api.deepseek.com",
    )

    if direction == "cn2en":
        lang_instruction = "Translate the following Chinese text to English."
    elif direction == "en2cn":
        lang_instruction = "Translate the following English text to Chinese."
    else:
        lang_instruction = (
            "Translate the following text to the other language "
            "(Chinese → English or English → Chinese, whichever applies)."
        )

    system_prompt = (
        f"{lang_instruction}\n\n"
        "CRITICAL RULES:\n"
        "1. Preserve ALL placeholders like ⟨T0⟩, ⟨T1⟩, ⟨T123⟩ — "
        "these are special tokens. Keep them EXACTLY as-is at the position "
        "where they appear. Do NOT translate, modify, split, or remove them.\n"
        "2. The surrounding text should read as natural, fluent prose in the "
        "target language. Adjust grammar, word order, prepositions, and "
        "articles as needed to accommodate where placeholders sit.\n"
        "3. Preserve paragraph breaks from the original.\n"
        "4. Output ONLY the translated text. No markdown, no HTML, no "
        "commentary, no notes, no explanations."
    )

    # Use chunking for long texts
    chunks = chunk_text(text)
    total_chunks = len(chunks)

    translated_chunks = []
    for i, chunk in enumerate(chunks):
        progress_bar.progress(
            (i / total_chunks) * 0.9,  # Reserve 10% for restore step
            text=f"翻译中… ({i + 1}/{total_chunks})",
        )
        status_text.text(f"正在翻译第 {i + 1}/{total_chunks} 段…")

        response = client.chat.completions.create(
            model=model,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": chunk},
            ],
            max_tokens=8192,
            temperature=0.3,  # Lower temperature for consistent translation
        )
        translated_chunks.append(response.choices[0].message.content)

    return "\n".join(translated_chunks)


# ---------------------------------------------------------------------------
# Translate button
# ---------------------------------------------------------------------------
can_translate = doc_file is not None and api_key.strip() != ""

if not can_translate:
    if doc_file is None:
        st.warning("👆 请先上传文档")
    elif not api_key.strip():
        st.warning("👆 请在侧边栏填入 API Key")

if st.button(
    "🚀 开始翻译",
    type="primary",
    disabled=not can_translate,
    use_container_width=True,
):
    # --- Step 1: Extract text ---
    with st.status("📄 提取文档文本…", expanded=True) as status:
        st.write("正在读取文档…")
        source_text = extract_text(doc_file.getvalue(), doc_file.name)
        st.write(f"✅ 提取完成，共 {len(source_text):,} 字符")

        if not source_text.strip():
            st.error("文档内容为空，请检查文件。")
            st.stop()

        # --- Step 2: Load glossary & protect terms ---
        use_glossary = False
        if glossary_file:
            st.write("📖 加载术语表…")
            with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
                tmp.write(glossary_file.getvalue())
                glossary_path = tmp.name
            glossary_data = load_glossary(glossary_path)
            os.unlink(glossary_path)
            use_glossary = True
        elif has_default_glossary:
            st.write("📖 加载内置术语表…")
            glossary_data = load_glossary(str(DEFAULT_GLOSSARY_PATH))
            use_glossary = True
        else:
            st.info("ℹ️ 未提供术语表，将作为普通 AI 翻译运行。")

        protected_data = None
        text_to_translate = source_text

        if use_glossary:
            st.write(f"🔍 术语表共 {glossary_data['count']} 条术语，正在匹配…")
            protected_data = protect_terms(source_text, glossary_data, direction)
            text_to_translate = protected_data["protected_text"]
            st.write(
                f"✅ 匹配到 **{protected_data['matched_count']}** 条术语，"
                f"已替换为占位符保护"
            )

        # --- Step 3: Translate ---
        status.update(label="🤖 AI 翻译中…", state="running")
        progress_bar = st.progress(0, text="准备翻译…")
        status_text = st.empty()

        translated = translate_text(
            text_to_translate,
            direction,
            api_key.strip(),
            model,
            progress_bar,
            status_text,
        )
        progress_bar.progress(0.95, text="翻译完成，正在恢复术语…")
        status_text.text("翻译完成！")

        # --- Step 4: Restore terms ---
        if protected_data:
            st.write("🔄 恢复术语…")
            final_text = restore_terms(translated, protected_data)
            st.write(f"✅ 已恢复 {protected_data['matched_count']} 条术语")
        else:
            final_text = translated

        progress_bar.progress(1.0, text="✅ 完成！")
        progress_bar.empty()
        status_text.empty()

        status.update(label="✅ 翻译完成！", state="complete")

        # --- Step 5: Show result & download ---
        st.divider()
        st.subheader("📋 翻译结果")

        # Stats row
        stat_col1, stat_col2, stat_col3, stat_col4 = st.columns(4)
        with stat_col1:
            st.metric("源文字符", f"{len(source_text):,}")
        with stat_col2:
            st.metric("译文字符", f"{len(final_text):,}")
        with stat_col3:
            st.metric(
                "翻译方向",
                {"cn2en": "中→英", "en2cn": "英→中"}.get(direction, direction),
            )
        with stat_col4:
            matched = protected_data["matched_count"] if protected_data else 0
            st.metric("匹配术语数", f"{matched} 条")

        # Text preview with tabs
        tab1, tab2, tab3 = st.tabs(["📝 译文预览", "📄 原文对照", "📊 术语详情"])

        with tab1:
            st.text_area(
                "翻译结果",
                value=final_text,
                height=400,
                label_visibility="collapsed",
            )

        with tab2:
            col_a, col_b = st.columns(2)
            with col_a:
                st.caption("**原文**")
                st.text_area(
                    "原文",
                    value=source_text[:5000]
                    + ("…" if len(source_text) > 5000 else ""),
                    height=400,
                    label_visibility="collapsed",
                )
            with col_b:
                st.caption("**译文**")
                st.text_area(
                    "译文",
                    value=final_text[:5000]
                    + ("…" if len(final_text) > 5000 else ""),
                    height=400,
                    label_visibility="collapsed",
                )

        with tab3:
            if protected_data and protected_data["matched_terms"]:
                st.dataframe(
                    [
                        {"源术语": t["source"], "目标翻译": t["target"]}
                        for t in protected_data["matched_terms"]
                    ],
                    use_container_width=True,
                    hide_index=True,
                )
            else:
                st.info("未使用术语表，无法显示术语详情。")

        st.divider()

        # Download button
        output_filename = Path(doc_file.name).stem + "_translated.txt"
        st.download_button(
            label="📥 下载译文 (.txt)",
            data=final_text.encode("utf-8"),
            file_name=output_filename,
            mime="text/plain",
            type="primary",
            use_container_width=True,
        )

        # --- Matched terms log (collapsed) ---
        if protected_data:
            with st.expander("🔍 术语匹配日志"):
                st.json(
                    {
                        "direction": direction,
                        "total_glossary_terms": glossary_data["count"],
                        "matched_count": protected_data["matched_count"],
                        "matched_terms": [
                            {"source": t["source"], "target": t["target"]}
                            for t in protected_data["matched_terms"]
                        ],
                    }
                )
