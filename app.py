"""
Canva Presentation Automation
  Tab 1 — Canva PPTX  →  Google Doc (formatted outline)
  Tab 2 — Google Doc  →  Canva Bulk Create CSV
"""

import csv
import io
import json
import re
import time
from typing import Optional

import streamlit as st
from pptx import Presentation
from pptx.enum.shapes import PP_PLACEHOLDER

# ══════════════════════════════════════════════════════════════════════════════
# PAGE CONFIG & CUSTOM CSS
# ══════════════════════════════════════════════════════════════════════════════

st.set_page_config(
    page_title="Canva Automation",
    page_icon="🎨",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown("""
<style>
  @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600&family=DM+Mono:wght@400;500&display=swap');

  html, body, [class*="css"] { font-family: 'DM Sans', sans-serif; }

  /* ── Top header bar ── */
  .top-header {
    background: linear-gradient(135deg, #0f0f23 0%, #1a1a3e 60%, #0d1b2a 100%);
    padding: 2rem 2.5rem 1.5rem;
    border-radius: 16px;
    margin-bottom: 1.5rem;
    border: 1px solid rgba(255,255,255,0.08);
  }
  .top-header h1 {
    color: #ffffff;
    font-size: 1.9rem;
    font-weight: 600;
    margin: 0 0 .35rem 0;
    letter-spacing: -.02em;
  }
  .top-header p {
    color: rgba(255,255,255,0.55);
    font-size: .9rem;
    margin: 0;
  }
  .badge {
    display: inline-block;
    background: rgba(99,102,241,0.25);
    color: #a5b4fc;
    font-size: .7rem;
    font-weight: 600;
    letter-spacing: .08em;
    text-transform: uppercase;
    padding: .2rem .6rem;
    border-radius: 100px;
    border: 1px solid rgba(99,102,241,0.4);
    margin-right: .5rem;
    margin-bottom: .6rem;
  }

  /* ── Cards ── */
  .card {
    background: #fafafa;
    border: 1px solid #e8e8ee;
    border-radius: 12px;
    padding: 1.4rem 1.6rem;
    margin-bottom: 1rem;
  }
  .card-title {
    font-size: .8rem;
    font-weight: 600;
    letter-spacing: .07em;
    text-transform: uppercase;
    color: #6366f1;
    margin-bottom: .5rem;
  }

  /* ── Outline preview ── */
  .outline-box {
    background: #0f0f1a;
    color: #e2e8f0;
    font-family: 'DM Mono', monospace;
    font-size: .8rem;
    line-height: 1.8;
    padding: 1.2rem 1.4rem;
    border-radius: 10px;
    white-space: pre;
    overflow-x: auto;
    max-height: 380px;
    overflow-y: auto;
    border: 1px solid rgba(99,102,241,0.3);
  }
  .outline-box .lvl0 { color: #93c5fd; font-weight: 600; }
  .outline-box .lvl1 { color: #c4b5fd; }
  .outline-box .lvl2 { color: #6ee7b7; }
  .outline-box .lvl3 { color: #fde68a; }
  .outline-box .lvl4 { color: #fca5a5; }

  /* ── Step pills ── */
  .step-pill {
    display: inline-flex;
    align-items: center;
    gap: .4rem;
    background: #f0f0ff;
    color: #4338ca;
    font-size: .78rem;
    font-weight: 600;
    padding: .3rem .8rem;
    border-radius: 100px;
    border: 1px solid #c7d2fe;
    margin-bottom: 1rem;
  }

  /* ── Info/success boxes ── */
  .info-box {
    background: #eff6ff;
    border: 1px solid #bfdbfe;
    border-radius: 10px;
    padding: .9rem 1.1rem;
    color: #1e40af;
    font-size: .85rem;
    margin: .8rem 0;
  }
  .success-box {
    background: #f0fdf4;
    border: 1px solid #86efac;
    border-radius: 10px;
    padding: .9rem 1.1rem;
    color: #15803d;
    font-size: .85rem;
    margin: .8rem 0;
  }
  .warn-box {
    background: #fffbeb;
    border: 1px solid #fcd34d;
    border-radius: 10px;
    padding: .9rem 1.1rem;
    color: #92400e;
    font-size: .85rem;
    margin: .8rem 0;
  }

  /* ── Canva field table ── */
  .field-table { width: 100%; border-collapse: collapse; font-size: .85rem; }
  .field-table th {
    background: #f5f3ff;
    color: #4338ca;
    font-weight: 600;
    padding: .5rem .8rem;
    text-align: left;
    border-bottom: 2px solid #e0d9ff;
  }
  .field-table td {
    padding: .45rem .8rem;
    border-bottom: 1px solid #f0f0f0;
    color: #374151;
  }
  .field-table tr:last-child td { border-bottom: none; }
  .code-chip {
    background: #1e1e2e;
    color: #a5b4fc;
    font-family: 'DM Mono', monospace;
    font-size: .78rem;
    padding: .15rem .5rem;
    border-radius: 5px;
  }

  /* ── Streamlit overrides ── */
  div[data-testid="stFileUploader"] { border-radius: 10px; }
  .stButton > button {
    font-family: 'DM Sans', sans-serif;
    font-weight: 500;
    border-radius: 8px;
  }
  div[data-baseweb="tab"] button { font-family: 'DM Sans', sans-serif; font-weight: 500; }
  div[data-testid="stDataFrame"] { border-radius: 10px; overflow: hidden; }
</style>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# OUTLINE LOGIC (shared between tabs)
# ══════════════════════════════════════════════════════════════════════════════

def to_roman(n: int) -> str:
    val  = [1000,900,500,400,100,90,50,40,10,9,5,4,1]
    syms = ["M","CM","D","CD","C","XC","L","XL","X","IX","V","IV","I"]
    result = ""
    for v, s in zip(val, syms):
        while n >= v:
            result += s; n -= v
    return result

def to_alpha(n: int, upper: bool = True) -> str:
    return chr((64 if upper else 96) + n)

OUTLINE_PREFIX = {
    0: lambda n: f"{to_roman(n)}.",
    1: lambda n: f"{to_alpha(n)}.",
    2: lambda n: f"{n}.",
    3: lambda n: f"{to_alpha(n, False)})",
    4: lambda n: f"({n})",
}

INDENT = "    "   # 4 spaces per level


def _is_title_ph(shape) -> bool:
    if not shape.is_placeholder:
        return False
    try:
        ph = shape.placeholder_format.type
        return ph in (PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE)
    except Exception:
        return False

def _avg_size(paras: list) -> float:
    sizes = [s for p in paras for s in p.get("_sizes", [])]
    return sum(sizes) / len(sizes) if sizes else 0


def extract_slides(pptx_bytes: bytes) -> list[dict]:
    prs = Presentation(io.BytesIO(pptx_bytes))
    slides_out = []
    for slide in prs.slides:
        shapes_info = []
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            paras = []
            for p in shape.text_frame.paragraphs:
                text = p.text.strip()
                if not text:
                    continue
                sizes = [r.font.size.pt for r in p.runs if r.font and r.font.size]
                paras.append({"text": text, "level": p.level or 0, "_sizes": sizes})
            if not paras:
                continue
            shapes_info.append({
                "is_title": _is_title_ph(shape),
                "paras": paras,
                "top": shape.top or 0,
                "left": shape.left or 0,
                "avg_size": _avg_size(paras),
            })
        shapes_info.sort(key=lambda s: (s["top"], s["left"]))

        title_shape = next((s for s in shapes_info if s["is_title"]), None)
        if title_shape is None and shapes_info:
            title_shape = max(shapes_info, key=lambda s: s["avg_size"])

        title_text = ""
        body_shapes = []
        for s in shapes_info:
            if s is title_shape:
                title_text = " ".join(p["text"] for p in s["paras"])
            else:
                body_shapes.append(s)

        body = [{"text": p["text"], "level": p["level"]} for s in body_shapes for p in s["paras"]]
        slides_out.append({"title": title_text, "body": body})
    return slides_out


def build_numbered_outline(slides: list[dict]) -> list[dict]:
    """Returns list of {level, prefix, text, formatted} dicts."""
    raw = []
    for slide in slides:
        if slide["title"]:
            raw.append({"level": 0, "text": slide["title"]})
        for p in slide["body"]:
            raw.append({"level": p["level"] + 1, "text": p["text"]})

    counters: dict[int, int] = {}
    result = []
    for item in raw:
        lvl = item["level"]
        for k in list(counters):
            if k > lvl: del counters[k]
        counters[lvl] = counters.get(lvl, 0) + 1
        prefix = OUTLINE_PREFIX.get(lvl, lambda n: "–")(counters[lvl])
        result.append({
            "level":     lvl,
            "prefix":    prefix,
            "text":      item["text"],
            "formatted": f"{INDENT * lvl}{prefix} {item['text']}",
        })
    return result


def outline_to_plain_text(outline: list[dict]) -> str:
    return "\n".join(item["formatted"] for item in outline)


def outline_to_html_preview(outline: list[dict]) -> str:
    css_cls = ["lvl0", "lvl1", "lvl2", "lvl3", "lvl4"]
    lines = []
    for item in outline:
        cls = css_cls[min(item["level"], 4)]
        escaped = item["formatted"].replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        lines.append(f'<span class="{cls}">{escaped}</span>')
    return '<div class="outline-box">' + "\n".join(lines) + "</div>"


# ══════════════════════════════════════════════════════════════════════════════
# GOOGLE DOCS HELPERS
# ══════════════════════════════════════════════════════════════════════════════

DOCS_SCOPE  = "https://www.googleapis.com/auth/documents"
DRIVE_SCOPE = "https://www.googleapis.com/auth/drive"
INDENT_PT   = 18  # points per level

def _get_service_account_creds(scopes: list[str]):
    """Build credentials from Streamlit secrets (service account JSON)."""
    try:
        from google.oauth2 import service_account as sa
        info = dict(st.secrets["gcp_service_account"])
        # private_key newlines may be escaped in TOML
        if "\\n" in info.get("private_key", ""):
            info["private_key"] = info["private_key"].replace("\\n", "\n")
        return sa.Credentials.from_service_account_info(info, scopes=scopes)
    except Exception as e:
        st.error(f"❌ Could not load Google credentials from secrets: {e}")
        return None


def gdoc_write_outline(outline: list[dict], title: str, share_email: Optional[str] = None) -> Optional[str]:
    """Create a Google Doc with the formatted outline. Returns doc URL or None."""
    from googleapiclient.discovery import build as gbuild

    scopes = [DOCS_SCOPE, DRIVE_SCOPE] if share_email else [DOCS_SCOPE]
    creds = _get_service_account_creds(scopes)
    if not creds:
        return None

    docs_svc  = gbuild("docs",  "v1", credentials=creds)
    drive_svc = gbuild("drive", "v3", credentials=creds) if share_email else None

    # Create doc
    doc    = docs_svc.documents().create(body={"title": title}).execute()
    doc_id = doc["documentId"]

    # Insert all text
    full_text = outline_to_plain_text(outline)
    docs_svc.documents().batchUpdate(
        documentId=doc_id,
        body={"requests": [{"insertText": {"location": {"index": 1}, "text": full_text}}]},
    ).execute()

    # Apply formatting per line
    fmt_requests = []
    char_idx = 1
    for i, item in enumerate(outline):
        ln  = len(item["formatted"])
        lvl = item["level"]
        start, end = char_idx, char_idx + ln

        fmt_requests.append({
            "updateParagraphStyle": {
                "range": {"startIndex": start, "endIndex": end},
                "paragraphStyle": {
                    "indentStart": {"magnitude": lvl * INDENT_PT, "unit": "PT"},
                    "spaceAbove":  {"magnitude": 4 if lvl == 0 else 0, "unit": "PT"},
                },
                "fields": "indentStart,spaceAbove",
            }
        })
        if lvl == 0:
            fmt_requests.append({
                "updateTextStyle": {
                    "range": {"startIndex": start, "endIndex": end},
                    "textStyle": {"bold": True, "fontSize": {"magnitude": 13, "unit": "PT"}},
                    "fields": "bold,fontSize",
                }
            })
        elif lvl == 1:
            fmt_requests.append({
                "updateTextStyle": {
                    "range": {"startIndex": start, "endIndex": end},
                    "textStyle": {"bold": True},
                    "fields": "bold",
                }
            })
        char_idx = end + (1 if i < len(outline) - 1 else 0)

    if fmt_requests:
        docs_svc.documents().batchUpdate(
            documentId=doc_id, body={"requests": fmt_requests}
        ).execute()

    # Share with user's email if provided
    if drive_svc and share_email:
        drive_svc.permissions().create(
            fileId=doc_id,
            body={"type": "user", "role": "writer", "emailAddress": share_email},
            sendNotificationEmail=False,
        ).execute()

    return f"https://docs.google.com/document/d/{doc_id}/edit"


def gdoc_read_lines(doc_id: str) -> list[str]:
    """Read non-empty lines from a Google Doc."""
    from googleapiclient.discovery import build as gbuild
    creds = _get_service_account_creds([DOCS_SCOPE])
    if not creds:
        return []
    svc = gbuild("docs", "v1", credentials=creds)
    doc = svc.documents().get(documentId=doc_id).execute()
    lines = []
    for element in doc.get("body", {}).get("content", []):
        para = element.get("paragraph")
        if not para:
            continue
        text = "".join(
            el["textRun"]["content"]
            for el in para.get("elements", [])
            if "textRun" in el
        ).strip()
        if text:
            lines.append(text)
    return lines


# ══════════════════════════════════════════════════════════════════════════════
# CSV GENERATION
# ══════════════════════════════════════════════════════════════════════════════

_ROMAN_RE = re.compile(r"^(M{0,4}(?:CM|CD|D?C{0,3})(?:XC|XL|L?X{0,3})(?:IX|IV|V?I{0,3}))\.\s+", re.I)
_ALPHA_RE = re.compile(r"^[A-Z]\.\s+")
_NUM_RE   = re.compile(r"^\d+\.\s+")
_LOWER_RE = re.compile(r"^[a-z]\)\s+")
_PAREN_RE = re.compile(r"^\(\d+\)\s+")

def _parse_level(line: str) -> tuple[int, str]:
    s = line.lstrip()
    if _ROMAN_RE.match(s): return 0, _ROMAN_RE.sub("", s).strip()
    if _ALPHA_RE.match(s): return 1, _ALPHA_RE.sub("", s).strip()
    if _NUM_RE.match(s):   return 2, _NUM_RE.sub("",   s).strip()
    if _LOWER_RE.match(s): return 3, _LOWER_RE.sub("", s).strip()
    if _PAREN_RE.match(s): return 4, _PAREN_RE.sub("", s).strip()
    return -1, s

def parse_slides_from_lines(lines: list[str]) -> list[dict]:
    slides, current = [], None
    for line in lines:
        lvl, text = _parse_level(line)
        if lvl < 0:
            continue
        if lvl == 0:
            current = {"title": text, "points": [], "subpoints": [], "details": [], "notes": []}
            slides.append(current)
        elif current is None:
            continue
        elif lvl == 1: current["points"].append(text)
        elif lvl == 2: current["subpoints"].append(text)
        elif lvl == 3: current["details"].append(text)
        elif lvl == 4: current["notes"].append(text)
    return slides

def generate_csv_bytes(slides: list[dict]) -> tuple[bytes, list[str]]:
    max_pts  = max((len(s["points"])    for s in slides), default=0)
    max_sub  = max((len(s["subpoints"]) for s in slides), default=0)
    max_det  = max((len(s["details"])   for s in slides), default=0)
    max_note = max((len(s["notes"])     for s in slides), default=0)

    headers = ["slide_title"]
    headers += [f"point_{i}"    for i in range(1, max_pts  + 1)]
    headers += [f"subpoint_{i}" for i in range(1, max_sub  + 1)]
    headers += [f"detail_{i}"   for i in range(1, max_det  + 1)]
    headers += [f"note_{i}"     for i in range(1, max_note + 1)]

    buf = io.StringIO()
    writer = csv.DictWriter(buf, fieldnames=headers, extrasaction="ignore")
    writer.writeheader()
    for slide in slides:
        row: dict = {"slide_title": slide["title"]}
        for i, t in enumerate(slide["points"],    1): row[f"point_{i}"]    = t
        for i, t in enumerate(slide["subpoints"], 1): row[f"subpoint_{i}"] = t
        for i, t in enumerate(slide["details"],   1): row[f"detail_{i}"]   = t
        for i, t in enumerate(slide["notes"],     1): row[f"note_{i}"]     = t
        writer.writerow(row)
    return buf.getvalue().encode("utf-8"), headers


# ══════════════════════════════════════════════════════════════════════════════
# HEADER
# ══════════════════════════════════════════════════════════════════════════════

st.markdown("""
<div class="top-header">
  <h1>🎨 Canva Presentation Automation</h1>
  <p>Extract slide text → format into an outline → push to Google Docs → generate Canva Bulk Create CSV</p>
  <br>
  <span class="badge">100% Free</span>
  <span class="badge">No Canva API needed</span>
  <span class="badge">Google Docs</span>
  <span class="badge">Bulk Create</span>
</div>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# TABS
# ══════════════════════════════════════════════════════════════════════════════

tab1, tab2 = st.tabs(["📄  Step 1 — PPTX → Google Doc", "📊  Step 2 — Google Doc → Canva CSV"])


# ─────────────────────────────────────────────────────────────────────────────
#  TAB 1
# ─────────────────────────────────────────────────────────────────────────────
with tab1:
    st.markdown('<div class="step-pill">① Upload your Canva export and build the outline</div>', unsafe_allow_html=True)

    col_up, col_opts = st.columns([1.4, 1], gap="large")

    with col_up:
        st.markdown("#### Upload Canva PPTX")
        pptx_file = st.file_uploader(
            "Export from Canva → Share → Download → PowerPoint (.pptx)",
            type=["pptx"],
            key="pptx_upload",
        )

    with col_opts:
        st.markdown("#### Options")
        doc_title = st.text_input("Google Doc title", value="Presentation Outline", key="doc_title")
        google_email = st.text_input(
            "Share doc with (your Gmail)",
            placeholder="you@gmail.com",
            help="The created doc will be shared with this address so it appears in your Drive.",
            key="google_email",
        )

        gdocs_enabled = "gcp_service_account" in st.secrets
        if gdocs_enabled:
            st.markdown('<div class="success-box">✅ Google service account detected — Google Docs push is enabled.</div>', unsafe_allow_html=True)
        else:
            st.markdown('<div class="warn-box">⚠️ No Google credentials found in secrets. You can still download the outline as a text file.</div>', unsafe_allow_html=True)

    # ── Process PPTX ─────────────────────────────────────────────────────────
    if pptx_file:
        with st.spinner("Extracting slides …"):
            pptx_bytes = pptx_file.read()
            slides  = extract_slides(pptx_bytes)
            outline = build_numbered_outline(slides)
            # Save to session state so Tab 2 can pick it up
            st.session_state["outline"] = outline
            st.session_state["outline_lines"] = outline_to_plain_text(outline).split("\n")

        st.markdown(f"**{len(slides)} slides → {len(outline)} outline items**")
        st.markdown("#### Outline Preview")
        st.markdown(outline_to_html_preview(outline), unsafe_allow_html=True)

        st.markdown("---")
        dl_col, gdoc_col = st.columns(2, gap="medium")

        # ── Download as text ─────────────────────────────────────────────────
        with dl_col:
            st.markdown("##### 📥 Option A — Download as text")
            plain_text = outline_to_plain_text(outline)
            st.download_button(
                label="⬇️  Download outline.txt",
                data=plain_text.encode("utf-8"),
                file_name="outline.txt",
                mime="text/plain",
                use_container_width=True,
            )
            st.caption("Use this in Tab 2 by uploading the .txt file, no Google account needed.")

        # ── Push to Google Doc ────────────────────────────────────────────────
        with gdoc_col:
            st.markdown("##### ☁️  Option B — Push to Google Doc")
            if gdocs_enabled:
                if st.button("🚀  Create Google Doc", use_container_width=True, type="primary"):
                    if not google_email:
                        st.warning("Enter your Gmail address above to share the doc with yourself.")
                    else:
                        with st.spinner("Creating Google Doc …"):
                            url = gdoc_write_outline(outline, doc_title, google_email)
                        if url:
                            doc_id = url.split("/d/")[1].split("/")[0]
                            st.session_state["last_doc_id"] = doc_id
                            st.markdown(f'<div class="success-box">✅ Doc created! <a href="{url}" target="_blank">Open in Google Docs ↗</a><br><small>Doc ID: <code>{doc_id}</code> (saved for Tab 2)</small></div>', unsafe_allow_html=True)
            else:
                st.button("🚀  Create Google Doc", use_container_width=True, disabled=True)
                st.caption("Add a service account to Streamlit secrets to enable this. See README.")
    else:
        st.markdown("""
<div class="info-box">
  👆 Upload a <strong>.pptx</strong> file exported from Canva to get started.<br><br>
  In Canva: <strong>Share → Download → Microsoft PowerPoint (.pptx)</strong>
</div>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
#  TAB 2
# ─────────────────────────────────────────────────────────────────────────────
with tab2:
    st.markdown('<div class="step-pill">② Generate the Canva Bulk Create CSV from your outline</div>', unsafe_allow_html=True)

    # ── Source selection ──────────────────────────────────────────────────────
    has_session_outline = "outline_lines" in st.session_state
    gdocs_enabled2      = "gcp_service_account" in st.secrets

    source_options = ["From Step 1 (this session)"] if has_session_outline else []
    source_options += ["Upload outline .txt file", "Google Doc ID"]
    source = st.radio("Outline source", source_options, horizontal=True)

    lines: list[str] = []

    if source == "From Step 1 (this session)" and has_session_outline:
        lines = st.session_state["outline_lines"]
        st.markdown(f'<div class="success-box">✅ Using outline from Step 1 — {len(lines)} lines loaded.</div>', unsafe_allow_html=True)

    elif source == "Upload outline .txt file":
        txt_file = st.file_uploader("Upload the outline.txt downloaded in Step 1", type=["txt"], key="txt_upload")
        if txt_file:
            lines = txt_file.read().decode("utf-8").split("\n")
            st.success(f"Loaded {len(lines)} lines from file.")

    elif source == "Google Doc ID":
        default_id = st.session_state.get("last_doc_id", "")
        doc_id_input = st.text_input(
            "Google Doc ID",
            value=default_id,
            placeholder="Paste the document ID from the URL",
            help="Found in the URL: docs.google.com/document/d/**THIS_PART**/edit",
        )
        if doc_id_input:
            if not gdocs_enabled2:
                st.markdown('<div class="warn-box">⚠️ Google credentials not configured — cannot read Google Docs directly. Download the outline.txt from Step 1 and upload it instead.</div>', unsafe_allow_html=True)
            else:
                if st.button("📥 Load from Google Doc"):
                    with st.spinner("Reading Google Doc …"):
                        lines = gdoc_read_lines(doc_id_input.strip())
                    if lines:
                        st.session_state["loaded_lines"] = lines
                        st.success(f"Loaded {len(lines)} lines from Google Doc.")
                    else:
                        st.error("Could not read the doc. Make sure the service account has access.")
                if "loaded_lines" in st.session_state:
                    lines = st.session_state["loaded_lines"]

    # ── Generate CSV ──────────────────────────────────────────────────────────
    if lines:
        slides = parse_slides_from_lines(lines)

        if not slides:
            st.warning("No slides detected. Make sure the outline uses the Roman numeral format from Step 1.")
        else:
            st.markdown(f"**{len(slides)} slides parsed**")

            csv_bytes, headers = generate_csv_bytes(slides)

            # Preview as dataframe
            import csv as csvlib
            import pandas as pd
            df = pd.read_csv(io.BytesIO(csv_bytes))
            st.markdown("#### CSV Preview")
            st.dataframe(df, use_container_width=True, height=240)

            st.download_button(
                label="⬇️  Download canva_bulk_create.csv",
                data=csv_bytes,
                file_name="canva_bulk_create.csv",
                mime="text/csv",
                use_container_width=False,
                type="primary",
            )

            # ── Canva instructions ────────────────────────────────────────────
            st.markdown("---")
            st.markdown("#### 🎨 How to use this CSV in Canva Bulk Create")

            st.markdown("""
<div class="info-box">
  Canva Bulk Create is a <strong>free built-in feature</strong> that fills a template from a CSV — one page per row.
  Follow the steps below to auto-generate your presentation.
</div>
""", unsafe_allow_html=True)

            with st.expander("📋 Step-by-step Canva setup (click to expand)", expanded=True):
                st.markdown("""
1. **Open your destination Canva template** (the design you want to fill).
2. **Click a textbox** you want to auto-fill (e.g. the slide title box).
3. In the top toolbar, click **"Connect data"** *(or go to Apps → Bulk Create → Get started)*.
4. **Name the field exactly** as shown in the table below — then press Enter.
5. Repeat for every textbox on the template (title, bullet 1, bullet 2, …).
6. Go to **Apps → Bulk Create → Get started → Upload CSV**.
7. Select `canva_bulk_create.csv`.
8. Click **Continue** → Canva generates one page per row. ✅
""")

                # Field name table
                PURPOSE = {
                    "slide_title": ("🔤 Slide / page title",     "The main heading textbox"),
                    "point_":      ("• Main bullet (A. level)",  "Primary body bullet points"),
                    "subpoint_":   ("  · Sub-bullet (1. level)", "Second-level bullet points"),
                    "detail_":     ("    › Detail (a) level)",   "Third-level detail text"),
                    "note_":       ("      ◦ Note ((1) level)",  "Deepest level notes"),
                }
                rows_html = ""
                for h in headers:
                    for prefix, (purpose, desc) in PURPOSE.items():
                        if h.startswith(prefix):
                            num = h.split("_")[-1] if h.split("_")[-1].isdigit() else ""
                            label = f"{purpose} #{num}" if num else purpose
                            rows_html += f"<tr><td>{label}</td><td><span class='code-chip'>{h}</span></td><td>{desc}</td></tr>"
                            break

                st.markdown(f"""
<table class="field-table">
  <thead><tr><th>Textbox purpose</th><th>Field name to enter in Canva</th><th>Description</th></tr></thead>
  <tbody>{rows_html}</tbody>
</table>
""", unsafe_allow_html=True)

    else:
        if not has_session_outline:
            st.markdown("""
<div class="info-box">
  Complete <strong>Step 1</strong> first, or upload an <code>outline.txt</code> file here directly.
</div>
""", unsafe_allow_html=True)
