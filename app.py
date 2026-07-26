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
import re
import base64
import zlib
import json
from pathlib import Path
from openai import OpenAI
from docx import Document
import pdfplumber

# Import the glossary helper (must be in the same directory)
from glossary_helper import (
    load_glossary,
    protect_terms,
    restore_terms,
    sanitize_glossary_dict,
)

# ---------------------------------------------------------------------------
# Embedded glossary (compressed & base64-encoded — not readable by humans)
# Generated from glossary.xlsx at build time.
# ---------------------------------------------------------------------------
_EMBEDDED_GLOSSARY_B64 = """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"""


def _decode_embedded_glossary():
    """Decode the embedded glossary from compressed base64.

    Applies the same sanitization as fresh xlsx loads so both paths behave
    identically (strips X/Y editorial alternatives, trailing (s) notes, and
    enumerated placeholder rows).
    """
    compressed = base64.b64decode(_EMBEDDED_GLOSSARY_B64)
    json_str = zlib.decompress(compressed).decode("utf-8")
    data = json.loads(json_str)
    return sanitize_glossary_dict(data)

# ---------------------------------------------------------------------------
# Page config
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="TransLegal",
    page_icon="⚖️",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ---------------------------------------------------------------------------
# Custom CSS
# ---------------------------------------------------------------------------
st.markdown("""
<style>
/* ===== Fonts (Claude uses Copernicus for headings; Instrument Serif is the closest free match) ===== */
@import url('https://fonts.googleapis.com/css2?family=Instrument+Serif:ital@0;1&family=Inter:wght@400;500;600&display=swap');

/* ===== Base ===== */
html, body, [class*="css"] {
    font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', 'PingFang SC', 'Hiragino Sans GB', 'Microsoft YaHei', sans-serif;
    color: #3D3929;
}

.stApp {
    background: #F5F4EE;
}

/* ===== Center content, cap width like claude.ai ===== */
.main .block-container {
    max-width: 880px;
    padding-top: 2.5rem;
    padding-bottom: 3rem;
}

/* ===== Hero (serif title, no card chrome) ===== */
.hero-wrap {
    text-align: center;
    padding: 1rem 0 0.5rem 0;
    margin-bottom: 2rem;
}
.hero-title {
    font-family: 'Instrument Serif', 'Iowan Old Style', Georgia, serif;
    font-size: 3.4rem;
    color: #3D3929;
    margin: 0;
    padding: 0;
    letter-spacing: -0.01em;
    font-weight: 400;
    line-height: 1.05;
}
.hero-subtitle {
    color: #87857A;
    font-size: 1rem;
    margin: 0.75rem 0 0 0;
    font-weight: 400;
}
.steps-row {
    display: flex;
    justify-content: center;
    gap: 2rem;
    margin: 1.5rem 0 0 0;
    font-size: 0.85rem;
    color: #87857A;
    flex-wrap: wrap;
}
.step-item {
    display: inline-flex;
    align-items: baseline;
    gap: 0.4rem;
}
.step-item .step-num {
    font-family: 'Instrument Serif', Georgia, serif;
    color: #D97757;
    font-size: 1.05rem;
    font-weight: 500;
    font-style: italic;
}

/* ===== Section cards (borderless-feel, subtle) ===== */
.settings-card {
    background: #FAF9F5;
    border: 1px solid #EDE9DE;
    border-radius: 12px;
    padding: 1.5rem;
    margin-bottom: 1rem;
}
.settings-card h3 {
    margin: 0 0 1rem 0 !important;
    font-family: 'Inter', sans-serif !important;
    font-size: 0.75rem !important;
    color: #87857A !important;
    text-transform: uppercase !important;
    letter-spacing: 0.1em !important;
    font-weight: 600 !important;
    border-bottom: none !important;
    padding-bottom: 0 !important;
}

/* ===== File uploader ===== */
[data-testid="stFileUploader"] {
    border: 1.5px dashed #D9D3C4 !important;
    border-radius: 10px !important;
    padding: 1rem !important;
    background: #FDFCF9 !important;
    transition: border-color 0.2s, background 0.2s;
}
[data-testid="stFileUploader"]:hover {
    border-color: #D97757 !important;
    background: #FBF7F2 !important;
}
[data-testid="stFileUploader"] section > button {
    background: #FFFFFF !important;
    border: 1px solid #E5E1D5 !important;
    color: #3D3929 !important;
    border-radius: 8px !important;
    font-weight: 500 !important;
}
[data-testid="stFileUploader"] section > button:hover {
    background: #F5F4EE !important;
    border-color: #D97757 !important;
}

/* ===== Buttons — secondary/default ===== */
.stButton > button {
    border-radius: 10px !important;
    font-weight: 500 !important;
    transition: background 0.15s ease, border-color 0.15s ease !important;
    padding: 0.55rem 1.4rem !important;
    border: 1px solid #E5E1D5 !important;
    background: #FFFFFF !important;
    color: #3D3929 !important;
    box-shadow: none !important;
}
.stButton > button:hover {
    background: #F5F4EE !important;
    transform: none !important;
    box-shadow: none !important;
}

/* Primary button — Claude copper */
.stButton > button[kind="primary"] {
    background: #D97757 !important;
    border: 1px solid #C7683F !important;
    color: #FFFFFF !important;
    font-size: 0.95rem !important;
}
.stButton > button[kind="primary"]:hover {
    background: #C7683F !important;
    border-color: #B85C34 !important;
}

/* Disabled */
.stButton > button:disabled {
    background: #EDE9DE !important;
    border-color: #EDE9DE !important;
    color: #B5B0A3 !important;
    cursor: not-allowed !important;
}

/* Download button */
.stDownloadButton > button {
    border-radius: 10px !important;
    font-weight: 500 !important;
    background: #D97757 !important;
    border: 1px solid #C7683F !important;
    color: #FFFFFF !important;
    transition: background 0.15s ease !important;
    padding: 0.55rem 1.4rem !important;
}
.stDownloadButton > button:hover {
    background: #C7683F !important;
    border-color: #B85C34 !important;
    box-shadow: none !important;
    transform: none !important;
}

/* ===== Metrics — serif value, all-caps small label ===== */
[data-testid="stMetric"] {
    background: #FAF9F5;
    border-radius: 10px;
    padding: 1rem 1.25rem;
    border: 1px solid #EDE9DE;
}
[data-testid="stMetricValue"] {
    color: #D97757 !important;
    font-family: 'Instrument Serif', Georgia, serif !important;
    font-size: 2rem !important;
    font-weight: 400 !important;
    line-height: 1.1 !important;
}
[data-testid="stMetricLabel"] {
    color: #87857A !important;
    font-size: 0.72rem !important;
    text-transform: uppercase !important;
    letter-spacing: 0.08em !important;
    font-weight: 500 !important;
}

/* ===== Text areas ===== */
.stTextArea textarea {
    border-radius: 10px !important;
    border: 1px solid #EDE9DE !important;
    background: #FFFFFF !important;
    font-family: 'Inter', sans-serif !important;
    padding: 0.75rem !important;
    color: #3D3929 !important;
    box-shadow: none !important;
}
.stTextArea textarea:focus {
    border-color: #D97757 !important;
    box-shadow: 0 0 0 3px rgba(217, 119, 87, 0.15) !important;
}

/* ===== Text input ===== */
.stTextInput input {
    border-radius: 10px !important;
    border: 1px solid #EDE9DE !important;
    background: #FFFFFF !important;
    padding: 0.55rem 0.85rem !important;
    color: #3D3929 !important;
    box-shadow: none !important;
}
.stTextInput input:focus {
    border-color: #D97757 !important;
    box-shadow: 0 0 0 3px rgba(217, 119, 87, 0.15) !important;
}

/* ===== Select ===== */
.stSelectbox [data-baseweb="select"] > div {
    border-radius: 10px !important;
    border-color: #EDE9DE !important;
    background: #FFFFFF !important;
}
.stSelectbox [data-baseweb="select"] > div:hover {
    border-color: #D97757 !important;
}

/* ===== Radio ===== */
.stRadio > div {
    gap: 0.5rem !important;
}
.stRadio [data-testid="stMarkdownContainer"] p {
    font-weight: 500;
    color: #3D3929;
}

/* ===== Progress bar ===== */
.stProgress > div > div {
    background: #D97757 !important;
    border-radius: 4px;
}
.stProgress > div {
    background: #EDE9DE !important;
    border-radius: 4px;
}

/* ===== Expander ===== */
[data-testid="stExpander"] {
    border-radius: 10px !important;
    border: 1px solid #EDE9DE !important;
    background: #FAF9F5 !important;
    box-shadow: none !important;
}
[data-testid="stExpander"] summary {
    color: #3D3929 !important;
    font-weight: 500 !important;
}
[data-testid="stExpander"] summary:hover {
    color: #D97757 !important;
}

/* ===== Info / warning callouts ===== */
div[data-testid="stNotification"] {
    border-radius: 10px !important;
    border: 1px solid #EDE9DE !important;
    background: #FAF9F5 !important;
    color: #3D3929 !important;
}

/* ===== Status widget ===== */
[data-testid="stStatusWidget"], .stStatus {
    border-radius: 10px !important;
    border: 1px solid #EDE9DE !important;
    background: #FAF9F5 !important;
}

/* ===== Divider ===== */
hr {
    border-color: #EDE9DE !important;
    margin: 2rem 0 !important;
}

/* ===== Tabs — underline style like claude.ai ===== */
.stTabs [data-baseweb="tab-list"] {
    gap: 0.5rem;
    border-bottom: 1px solid #EDE9DE;
    background: transparent !important;
}
.stTabs [data-baseweb="tab"] {
    border-radius: 0 !important;
    padding: 0.6rem 0.75rem !important;
    font-weight: 500;
    color: #87857A !important;
    background: transparent !important;
    border-bottom: 2px solid transparent !important;
    margin-bottom: -1px !important;
}
.stTabs [aria-selected="true"] {
    color: #D97757 !important;
    border-bottom: 2px solid #D97757 !important;
    background: transparent !important;
}

/* ===== Dataframe ===== */
[data-testid="stDataFrame"] {
    border-radius: 10px;
    overflow: hidden;
    border: 1px solid #EDE9DE;
}

/* ===== Caption ===== */
[data-testid="stCaptionContainer"], .stCaption {
    color: #87857A !important;
    font-size: 0.85rem !important;
}

/* ===== Headings ===== */
h2, h3, h4 {
    color: #3D3929 !important;
    font-weight: 500 !important;
}
h2 {
    font-family: 'Instrument Serif', 'Iowan Old Style', Georgia, serif !important;
    font-weight: 400 !important;
    font-size: 1.9rem !important;
    letter-spacing: -0.01em !important;
    margin-top: 2rem !important;
}

/* ===== Hide sidebar default ===== */
[data-testid="stSidebar"] {
    display: none;
}

/* ===== Footer ===== */
.app-footer {
    text-align: center;
    color: #A39B8B;
    font-size: 0.8rem;
    padding: 2.5rem 0 1rem 0;
    margin-top: 3rem;
    font-family: 'Instrument Serif', Georgia, serif;
    font-style: italic;
}

/* Remove default streamlit deploy button and menu clutter */
.stDeployButton { display: none !important; }
#MainMenu { visibility: hidden; }
header { background: transparent !important; }
</style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# Hero Section
# ---------------------------------------------------------------------------
st.markdown("""
<div class="hero-wrap">
    <h1 class="hero-title">⚖️ TransLegal</h1>
    <p class="hero-subtitle">上传文档 · AI 匹配术语并翻译 · 下载术语一致的译文</p>
    <div class="steps-row">
        <div class="step-item"><span class="step-num">01</span> 上传文档 & 术语表</div>
        <div class="step-item"><span class="step-num">02</span> 选择翻译选项</div>
        <div class="step-item"><span class="step-num">03</span> 下载译文</div>
    </div>
</div>
""", unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# Main Setup Area — two columns
# ---------------------------------------------------------------------------
col_left, col_right = st.columns([5, 4], gap="large")

with col_right:
    st.markdown('<div class="settings-card">', unsafe_allow_html=True)
    st.markdown("### 翻译设置")

    _server_key = os.environ.get("DEEPSEEK_API_KEY", "")
    api_key_input = st.text_input(
        "DeepSeek API Key",
        type="password",
        placeholder="已预填，无需修改" if _server_key else "sk-...",
        help="已从服务器环境变量自动加载。也可填入自己的 Key 覆盖。",
    )
    effective_key = api_key_input.strip() if api_key_input.strip() else _server_key

    direction_label = st.radio(
        "翻译方向",
        options=["自动检测", "中 → 英", "英 → 中"],
        index=0,
        horizontal=True,
    )
    direction_map = {"自动检测": "auto", "中 → 英": "cn2en", "英 → 中": "en2cn"}
    direction = direction_map[direction_label]

    model = st.selectbox(
        "翻译模型",
        options=["deepseek-v4-flash", "deepseek-v4-pro"],
        index=0,
        format_func=lambda m: {
            "deepseek-v4-flash": "V4 Flash — 快速便宜，适合日常翻译",
            "deepseek-v4-pro": "V4 Pro — 旗舰质量，适合正式文档",
        }.get(m, m),
    )
    st.markdown('</div>', unsafe_allow_html=True)

    # Hidden tip
    _embedded_preview = _decode_embedded_glossary()
    st.caption(
        f"术语表文件不上传时，内置术语表（{_embedded_preview['count']} 条）将自动使用。"
    )

with col_left:
    st.markdown('<div class="settings-card">', unsafe_allow_html=True)
    st.markdown("### 上传文件")

    doc_file = st.file_uploader(
        "文档（.docx / .pdf）",
        type=["docx", "pdf"],
        help="上传待翻译的文档",
    )

    glossary_file = st.file_uploader(
        "术语表（可选 .xlsx）",
        type=["xlsx"],
        help="格式：第一列 = 源语言术语，第二列 = 目标语言翻译。不上传则自动使用内置术语表。",
    )

    # Show file info after upload
    if doc_file:
        file_size_kb = len(doc_file.getvalue()) / 1024
        st.caption(f"已选择：**{doc_file.name}** ({file_size_kb:.1f} KB)")

    if glossary_file:
        st.caption(f"自定义术语表：**{glossary_file.name}**")

    st.markdown('</div>', unsafe_allow_html=True)

    # --- Translate button ---
    can_translate = doc_file is not None and effective_key.strip() != ""

    if not can_translate:
        if doc_file is None:
            st.info("请上传待翻译的文档开始")
        elif not effective_key.strip():
            st.warning("请在右侧填入 DeepSeek API Key")

    col_btn1, col_btn2, col_btn3 = st.columns([1, 2, 1])
    with col_btn2:
        translate_clicked = st.button(
            "→ 开始翻译",
            type="primary",
            disabled=not can_translate,
            use_container_width=True,
        )

# ---------------------------------------------------------------------------
# Translation logic
# ---------------------------------------------------------------------------
_PAGE_FOOTER_RE = re.compile(
    r"^\s*(?:[-–—]\s*\d+\s*[-–—]|第\s*\d+\s*页(?:\s*(?:共|/)\s*\d+\s*页)?)\s*$",
    re.MULTILINE,
)
_MULTI_BLANK_RE = re.compile(r"\n\s*\n\s*\n+")

# Punctuation that legitimately ends a line in CJK/Latin. If a PDF line ends
# with one of these, we treat the line break as real; otherwise it's likely
# a mid-sentence wrap and we join with the next line.
_LINE_END_PUNCT = set("。！？；：.!?;:）)】」』】\"”’…—")
_LINE_START_MARKERS = re.compile(
    r"^\s*(?:[一二三四五六七八九十]+[、.．]|\d+[、.．)]|（\d+）|\(\d+\)|[①②③④⑤⑥⑦⑧⑨⑩]|[IVX]+[\.\)]|第[一二三四五六七八九十百千0-9]+[条章节款项])",
)


def clean_extracted_text(text: str, suffix: str) -> str:
    """Post-process raw extracted text to remove page-level artifacts.

    - Strip standalone page footers like "- 1 -" and "第 3 页 共 7 页".
    - For PDFs, join mid-sentence line wraps back into paragraphs.
    - Collapse runs of 3+ blank lines down to a single blank line.
    """
    text = _PAGE_FOOTER_RE.sub("", text)

    if suffix == ".pdf":
        merged_lines = []
        for line in text.split("\n"):
            stripped = line.rstrip()
            if (
                merged_lines
                and merged_lines[-1]
                and merged_lines[-1][-1] not in _LINE_END_PUNCT
                and stripped
                and not _LINE_START_MARKERS.match(stripped)
            ):
                merged_lines[-1] = merged_lines[-1] + stripped.lstrip()
            else:
                merged_lines.append(stripped)
        text = "\n".join(merged_lines)

    text = _MULTI_BLANK_RE.sub("\n\n", text)
    return text.strip()


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

    return clean_extracted_text(text, suffix)


def chunk_text(text: str, max_chars: int = 4000) -> list[str]:
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
            if len(para) > max_chars:
                for i in range(0, len(para), max_chars):
                    chunks.append(para[i:i + max_chars])
                current = ""
            else:
                current = para

    if current:
        chunks.append(current)

    return chunks if chunks else [""]


# ---------------------------------------------------------------------------
# Output post-processing: style / format fixes applied after model returns
# ---------------------------------------------------------------------------
_ROLE_ALT = r"Plaintiff|Defendant|Appellant|Appellee|Petitioner|Respondent|Applicant"

_MD_BOLD_STAR_RE = re.compile(r"\*\*([^*\n]+?)\*\*")
_MD_BOLD_UNDER_RE = re.compile(r"__([^_\n]+?)__")
_MD_ITAL_UNDER_RE = re.compile(r"(?<![A-Za-z0-9_])_([^_\n]+?)_(?![A-Za-z0-9_])")
_MD_ITAL_STAR_RE = re.compile(r"(?<![*A-Za-z0-9])\*([^*\n]+?)\*(?![*A-Za-z0-9])")

_THE_ROLE_RE = re.compile(r"\b[Tt]he (" + _ROLE_ALT + r")(s?)\b")
_ROLE_SPACE_S_RE = re.compile(r"\b(" + _ROLE_ALT + r") s\b")

# Handles all four surface variants the model emits:
#   (hereinafter referred to as "Apple")           → ("Apple")
#   (hereinafter referred to as the "CNIPA")       → (the "CNIPA")
#   (hereinafter referred to as Sentecke Company)  → ("Sentecke Company")
#   (hereinafter referred to as the Company)       → (the "Company")
_HEREIN_RE = re.compile(
    r'\(hereinafter referred to as (the )?"?([^)"]+?)"?\s*\)'
)

# Article citation normalisation: convert model's verbose forms to compact
#   Article 130, Paragraph 1           → Article 130.1
#   Article 19, Item (3)               → Article 19(3)
#   Article 157, Paragraph 1, Item (2) → Article 157.1(2)
_ARTICLE_ITEM_RE = re.compile(
    r"Article (\d+), (Paragraph (\d+), )?Item \(([^)]+)\)"
)
_ARTICLE_PARA_RE = re.compile(
    r"Article (\d+), Paragraph (\d+)"
)

# Redundant "patent number No." → "patent number"
_PATENT_NUMBER_NO_RE = re.compile(r"patent number No\.\s*")

# "offering for sale" after a past-participle verb in a list → "offered for sale"
# Matches: sold and offering for sale  /  manufactured, sold, offering for sale
_OFFERING_PAST_RE = re.compile(
    r"\b(sold|manufactured|imported|produced|distributed|offered)"
    r"((?:,?\s+(?:and\s+)?))offering for sale\b"
)

# Glossary-collision dedup: model writes "the place where the ⟨T⟩" but the
# glossary target already starts with "place where" → "the place where the
# place where X" → "the place where X"
_PLACE_WHERE_DUP_RE = re.compile(r"\bthe place where the place where\b")
# "the place *of* place where" — model writes "the place of" + glossary target
# "place where" → "the place of place where" → "the place where"
_PLACE_OF_PLACE_RE = re.compile(r"\bthe place of place where\b")

# Glossary collision: model writes "where the ⟨T⟩ is domiciled" but the
# glossary target is "defendant's domicile" → "defendant's domicile is
# domiciled" → "defendant's domicile"
_DOMICILE_DUP_RE = re.compile(r"\bdefendant's domicile is domiciled\b")

# "is committed occurred" → "is committed"
_COMMITTED_OCCURRED_RE = re.compile(r"\bis committed occurred\b")

# Grammar: "should be transfer to" → "should be transferred to"
_TRANSFER_TO_RE = re.compile(r"\bshould be transfer to\b")

# Grammar: "a [vowel]" → "an [vowel]"
# Note: U/u is deliberately excluded — in legal documents "a United States",
# "a user", "a USB" are far more common than "an umbrella" or "an undertaking".
# Missing "an" before rare U-words is preferable to the glaring "an United".
_A_AN_RE = re.compile(r"\ba ([AEIOaeio])")


def _herein_sub(m):
    prefix = m.group(1) or ""
    return f'({prefix}"{m.group(2).strip()}")'


def _article_item_sub(m):
    """Article N, [Paragraph P, ]Item (X) → Article N[.P](X)."""
    if m.group(2):  # has Paragraph
        return f"Article {m.group(1)}.{m.group(3)}({m.group(4)})"
    return f"Article {m.group(1)}({m.group(4)})"


def _article_para_sub(m):
    """Article N, Paragraph P → Article N.P."""
    return f"Article {m.group(1)}.{m.group(2)}"


def _offering_past_sub(m):
    """sold and offering for sale → sold and offered for sale."""
    return f"{m.group(1)}{m.group(2)}offered for sale"


def _to_curly_quotes(text: str) -> str:
    """Convert ASCII straight quotes to curly quotes.

    Double quotes alternate open/close per line. For single quotes, the
    typographic rule is: preceded by an alphanumeric char → right single
    (’) — this covers both apostrophes (Plaintiff's) and closing quotes
    after a word (no'); preceded by whitespace/punctuation → left single
    (‘), i.e. an opening quote. No state needed for singles.
    """
    out = []
    double_open = True
    for i, ch in enumerate(text):
        if ch == "\n":
            double_open = True
            out.append(ch)
        elif ch == '"':
            out.append("“" if double_open else "”")
            double_open = not double_open
        elif ch == "'":
            prev = text[i - 1] if i > 0 else ""
            out.append("’" if prev.isalnum() else "‘")
        else:
            out.append(ch)
    return "".join(out)


def postprocess_translation(text: str) -> str:
    """Apply style/format fixes to raw model output.

    Order matters:
    1. Strip markdown emphasis (** _ * markers) — must run before quote
       conversion, otherwise stray asterisks get quoted in.
    2. Simplify 'hereinafter referred to as' — while ASCII quotes are intact.
    3. Fix 'Defendant s' space-plural artifact.
    4. Drop leading 'the' before role labels (after step 3 so 'the Defendant s'
       first becomes 'the Defendants' then 'Defendants').
    5. Normalise article citations (Article 130, Paragraph 1 → Article 130.1).
    6. Fix "patent number No." redundancy.
    7. Fix "offering for sale" after past participles in lists.
    8. Convert ASCII straight quotes to curly quotes — must run last so all
       previous regexes see plain ASCII quotes.
    """
    text = _MD_BOLD_STAR_RE.sub(r"\1", text)
    text = _MD_BOLD_UNDER_RE.sub(r"\1", text)
    text = _MD_ITAL_UNDER_RE.sub(r"\1", text)
    text = _MD_ITAL_STAR_RE.sub(r"\1", text)

    text = _HEREIN_RE.sub(_herein_sub, text)

    text = _ROLE_SPACE_S_RE.sub(r"\1s", text)
    text = _THE_ROLE_RE.sub(r"\1\2", text)

    # Article citations: order matters — item+paragraph before plain paragraph
    text = _ARTICLE_ITEM_RE.sub(_article_item_sub, text)
    text = _ARTICLE_PARA_RE.sub(_article_para_sub, text)
    # Fix "patent number No." → "patent number"
    text = _PATENT_NUMBER_NO_RE.sub("patent number ", text)
    # Fix "sold and offering for sale" → "sold and offered for sale"
    text = _OFFERING_PAST_RE.sub(_offering_past_sub, text)
    # Fix glossary collisions
    text = _PLACE_WHERE_DUP_RE.sub("the place where", text)
    text = _PLACE_OF_PLACE_RE.sub("the place where", text)
    text = _DOMICILE_DUP_RE.sub("defendant's domicile", text)
    text = _COMMITTED_OCCURRED_RE.sub("is committed", text)
    # Fix common grammar errors
    text = _TRANSFER_TO_RE.sub("should be transferred to", text)
    text = _A_AN_RE.sub(r"an \1", text)

    text = _to_curly_quotes(text)
    return text


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
        lang_instruction = "Translate the following Chinese legal text into English."
    elif direction == "en2cn":
        lang_instruction = "Translate the following English legal text into Chinese."
    else:
        lang_instruction = (
            "Translate the following legal text to the other language "
            "(Chinese → English or English → Chinese, whichever applies)."
        )

    system_prompt = (
        f"{lang_instruction}\n\n"
        "This is a formal legal document (court ruling, contract, brief, or "
        "notarial certificate). Translate with the register and conventions of "
        "professional legal drafting.\n\n"
        "FORMATTING PROHIBITIONS (strict — the output is plain text):\n"
        "- Do NOT use markdown bold (**...** or __...__).\n"
        "- Do NOT use markdown italics (*...* or _..._).\n"
        "- Do NOT use markdown headings (#, ##, ###).\n"
        "- Do NOT use markdown code fences or inline code (```, `).\n"
        "- Do NOT use markdown list bullets (-, *, +). Preserve the source's "
        "own numbering (1., I., (1), ①, 一、) verbatim as ordinary text.\n"
        "- If the source uses visual emphasis, express it through word choice "
        "or CAPITALIZATION, never through markdown symbols.\n\n"
        "REGISTER & STYLE:\n"
        "- Use formal legal English: 'shall' for obligations, 'hereby', "
        "'notwithstanding', 'pursuant to', 'in accordance with', etc.\n"
        "- Be precise. Do not paraphrase, soften, or add explanatory hedges. "
        "Keep the sentence structure close to the source when possible.\n"
        "- For Chinese → English, use standard PRC-legal-English conventions: "
        "'the People's Republic of China', 'People's Court', 'Civil Procedure Law'.\n"
        "- Use curly quotes (“”‘’) rather than straight quotes.\n\n"
        "NAMES & PARTY LABELS:\n"
        "- When transliterating Chinese personal names to Pinyin, UPPERCASE "
        "the surname (usually the first token). Examples: '张乃倩' → "
        "'ZHANG Naiqian'; '陈丽美' → 'CHEN Limei'; '蓝世文' → 'LAN Shiwen'. "
        "Apply this to parties, legal representatives, attorneys, judges, and "
        "court clerks. Foreign personal names stay as-is.\n"
        "- Party-role labels (Plaintiff, Defendant, Appellant, Appellee, "
        "Petitioner, Respondent, Applicant) are used WITHOUT a leading 'the'. "
        "Write 'Defendant argues that...', 'submitted by Plaintiff', NOT "
        "'the Defendant argues...' or 'submitted by the Plaintiff'. Capitalize "
        "them when referring to parties in this case.\n"
        "- Company abbreviations use the format 'Full Name (\"Abbrev\")' — "
        "do NOT write 'hereinafter referred to as'. Example: 'Apple Electronics "
        "Products Commerce (Beijing) Co., Ltd. (\"Apple\")'. Company abbreviations "
        "take no leading 'the'.\n"
        "- Non-company term abbreviations use the format 'Full Name (the \"Abbrev\")'. "
        "Example: 'The China National Intellectual Property Administration "
        "(the \"CNIPA\")'.\n"
        "- Translate '本院' and '贵院' (this court / your court) uniformly as "
        "'the Court' (capital C).\n\n"
        "HEADINGS:\n"
        "- Top-level document title: Title Case, on its own line, no trailing period.\n"
        "- Second-level headings: start with 'I. ', 'II. ', 'III. '; heading text "
        "in Sentence case (first letter capitalized, rest lowercase except proper "
        "nouns); NO trailing period.\n"
        "- Third-level headings: start with '(i). ', '(ii). ', '(iii). '; also "
        "Sentence case, NO trailing period.\n"
        "- Never decorate any heading with markdown symbols.\n\n"
        "CITATIONS:\n"
        "- '第 X 条第 Y 款第 Z 项' → 'Article X.Y(Z)' (e.g. 'Article 157.1(2)').\n"
        "- '第 X 条第 Y 款' → 'Article X.Y'.\n"
        "- '第 X 条' → 'Article X'.\n"
        "- Preserve the original article/section numbers exactly; do not "
        "re-number or convert Chinese numerals into anything other than the "
        "equivalent Arabic numeral in the citation form above.\n\n"
        "PLACEHOLDERS:\n"
        "- Special tokens like ⟨T0⟩, ⟨T1⟩, ⟨T123⟩ are pre-translated glossary "
        "terms. Keep them EXACTLY as written — do not translate, split, alter, "
        "or drop them.\n"
        "- Treat each ⟨Tn⟩ as a noun phrase. When a ⟨Tn⟩ sits next to a Latin "
        "word, insert whitespace so the output reads naturally (write "
        "'⟨T0⟩ Company', not '⟨T0⟩Company').\n"
        "- If your own natural translation would already contain the same word "
        "that a nearby ⟨Tn⟩ expands to (e.g. you're about to write 'jurisdiction' "
        "and a ⟨Tn⟩ token also carries 'jurisdiction'), drop your redundant word "
        "so the sentence reads once, not twice.\n\n"
        "STRUCTURE:\n"
        "- Preserve paragraph breaks and blank lines.\n"
        "- Preserve inline tables, dates, currency amounts, and identifiers "
        "(patent numbers, case numbers, credit codes, addresses) verbatim.\n\n"
        "OUTPUT: Only the translated text. No markdown, no HTML, no commentary, "
        "no notes, no explanations, no XML tags."
    )

    chunks = chunk_text(text)
    total_chunks = len(chunks)

    translated_chunks = []
    for i, chunk in enumerate(chunks):
        progress_bar.progress(
            (i / total_chunks) * 0.9,
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
            temperature=0.3,
        )
        translated_chunks.append(response.choices[0].message.content)

    return postprocess_translation("\n".join(translated_chunks))


# ---------------------------------------------------------------------------
# Translation execution
# ---------------------------------------------------------------------------
if translate_clicked:
    if doc_file is None:
        st.error("文档已丢失，请重新上传后再点开始翻译。")
        st.stop()

    # Capture bytes and name eagerly. On some Streamlit rerun paths the
    # uploader reference can go stale mid-execution — reading once up front
    # avoids the "'NoneType' object has no attribute 'getvalue'" trap.
    doc_bytes = doc_file.getvalue()
    doc_name = doc_file.name
    glossary_bytes = glossary_file.getvalue() if glossary_file else None

    # --- Step 1: Extract text ---
    with st.status("提取文档文本…", expanded=True) as status:
        st.write("正在读取文档…")
        source_text = extract_text(doc_bytes, doc_name)
        st.write(f"提取完成，共 **{len(source_text):,}** 字符")

        if not source_text.strip():
            st.error("文档内容为空，请检查文件。")
            st.stop()

        # --- Step 2: Load glossary & protect terms ---
        use_glossary = False
        glossary_data = None
        if glossary_bytes is not None:
            st.write("加载用户术语表…")
            with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
                tmp.write(glossary_bytes)
                glossary_path = tmp.name
            glossary_data = load_glossary(glossary_path)
            os.unlink(glossary_path)
            use_glossary = True
        else:
            st.write("加载内置术语表…")
            glossary_data = _decode_embedded_glossary()
            use_glossary = True

        protected_data = None
        text_to_translate = source_text

        if use_glossary:
            cleaned = glossary_data.get("cleaned_count", 0)
            skipped = glossary_data.get("skipped_count", 0)
            hygiene_note = (
                f"（已清洗 {cleaned} 条，跳过 {skipped} 条枚举模板）"
                if (cleaned or skipped)
                else ""
            )
            st.write(
                f"术语表共 **{glossary_data['count']}** 条术语{hygiene_note}，正在匹配…"
            )
            protected_data = protect_terms(source_text, glossary_data, direction)
            text_to_translate = protected_data["protected_text"]
            st.write(
                f"匹配到 **{protected_data['matched_count']}** 条术语，已替换为占位符"
            )

        # --- Step 3: Translate ---
        status.update(label="AI 翻译中…", state="running")
        progress_bar = st.progress(0, text="准备翻译…")
        status_text = st.empty()

        translated = translate_text(
            text_to_translate,
            direction,
            effective_key.strip(),
            model,
            progress_bar,
            status_text,
        )
        progress_bar.progress(0.95, text="翻译完成，正在恢复术语…")
        status_text.text("翻译完成！")

        # --- Step 4: Restore terms ---
        if protected_data:
            st.write("恢复术语…")
            final_text = restore_terms(translated, protected_data)
            st.write(f"已恢复 **{protected_data['matched_count']}** 条术语")
        else:
            final_text = translated

        progress_bar.progress(1.0, text="完成！")
        progress_bar.empty()
        status_text.empty()

        status.update(label="翻译完成！", state="complete")

        st.toast("翻译完成！请查看结果并下载译文。", icon=":material/check_circle:")

        # --- Step 5: Store in session state for persistent display ---
        st.session_state.translation_result = {
            "source_text": source_text,
            "final_text": final_text,
            "direction": direction,
            "protected_data": protected_data,
            "glossary_data": glossary_data,
            "doc_filename": doc_name,
        }

# ---------------------------------------------------------------------------
# Display results (persisted in session state)
# ---------------------------------------------------------------------------
if "translation_result" in st.session_state:
    result = st.session_state.translation_result
    source_text = result["source_text"]
    final_text = result["final_text"]
    direction = result["direction"]
    protected_data = result["protected_data"]
    glossary_data = result["glossary_data"]
    doc_filename = result["doc_filename"]

    st.divider()
    st.markdown("## 翻译结果")

    # --- Stats row in cards ---
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
        st.metric("匹配术语", f"{matched} 条")

    # --- Tabs ---
    tab1, tab2, tab3 = st.tabs(["译文预览", "原文对照", "术语详情"])

    with tab1:
        st.text_area(
            "翻译结果",
            value=final_text,
            height=420,
            label_visibility="collapsed",
            key="result_preview",
        )

    with tab2:
        col_a, col_b = st.columns(2)
        with col_a:
            st.caption("**原文**")
            st.text_area(
                "原文",
                value=source_text[:5000]
                + ("…" if len(source_text) > 5000 else ""),
                height=420,
                label_visibility="collapsed",
                key="source_preview",
            )
        with col_b:
            st.caption("**译文**")
            st.text_area(
                "译文",
                value=final_text[:5000]
                + ("…" if len(final_text) > 5000 else ""),
                height=420,
                label_visibility="collapsed",
                key="translated_preview",
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

    # --- Download button ---
    st.divider()
    col_dl1, col_dl2, col_dl3 = st.columns([1, 3, 1])
    with col_dl2:
        output_filename = Path(doc_filename).stem + "_translated.txt"
        st.download_button(
            label="↓ 下载译文 (.txt)",
            data=final_text.encode("utf-8"),
            file_name=output_filename,
            mime="text/plain",
            type="primary",
            use_container_width=True,
        )

    # --- Matched terms log (collapsed) ---
    if protected_data:
        with st.expander("术语匹配详情"):
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

# ---------------------------------------------------------------------------
# Footer
# ---------------------------------------------------------------------------
st.markdown("""
<div class="app-footer">
    TransLegal · Powered by DeepSeek
</div>
""", unsafe_allow_html=True)
