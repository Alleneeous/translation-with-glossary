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
import base64
import zlib
import json
from pathlib import Path
from openai import OpenAI
from docx import Document
import pdfplumber

# Import the glossary helper (must be in the same directory)
from glossary_helper import load_glossary, protect_terms, restore_terms

# ---------------------------------------------------------------------------
# Embedded glossary (compressed & base64-encoded — not readable by humans)
# Generated from glossary.xlsx at build time.
# ---------------------------------------------------------------------------
_EMBEDDED_GLOSSARY_B64 = """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"""


def _decode_embedded_glossary():
    """Decode the embedded glossary from compressed base64."""
    compressed = base64.b64decode(_EMBEDDED_GLOSSARY_B64)
    json_str = zlib.decompress(compressed).decode("utf-8")
    return json.loads(json_str)

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

# --- Default glossary (embedded, not visible as a file) ---
has_default_glossary = True  # Glossary is baked into the app code

if has_default_glossary and not glossary_file:
    st.info("📌 已加载内置术语表（210条），将自动使用。可上传自己的术语表覆盖。")

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
            st.write("📖 加载用户术语表…")
            with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
                tmp.write(glossary_file.getvalue())
                glossary_path = tmp.name
            glossary_data = load_glossary(glossary_path)
            os.unlink(glossary_path)
            use_glossary = True
        elif has_default_glossary:
            st.write("📖 加载内置术语表…")
            glossary_data = _decode_embedded_glossary()
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
