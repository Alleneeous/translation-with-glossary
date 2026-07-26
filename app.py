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

# Glossary collision: model writes "where the ⟨T⟩ is domiciled" but the
# glossary target is "defendant's domicile" → "defendant's domicile is
# domiciled" → "defendant's domicile"
_DOMICILE_DUP_RE = re.compile(r"\bdefendant's domicile is domiciled\b")

# "is committed occurred" → "is committed"
_COMMITTED_OCCURRED_RE = re.compile(r"\bis committed occurred\b")

# Grammar: "should be transfer to" → "should be transferred to"
_TRANSFER_TO_RE = re.compile(r"\bshould be transfer to\b")

# Grammar: "a [vowel]" → "an [vowel]"
_A_AN_RE = re.compile(r"\ba ([AEIOUaeiou])")


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
