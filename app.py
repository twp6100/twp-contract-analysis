import streamlit as st

# ======================================================================
#  AUTHENTICATION GATE
#  When deployed to Streamlit Community Cloud with an [auth] section
#  in secrets.toml, this requires the user to log in with Google.
#  When run locally without an [auth] section, the gate is skipped.
# ======================================================================
def _auth_is_configured():
    try:
        return "auth" in st.secrets and st.secrets["auth"].get("client_id")
    except Exception:
        return False


def _render_login_screen():
    st.markdown(
        '<h1 style="color:#1a1a2e;text-align:center;margin-top:4rem;">'
        "🔒 TWP Contract Analysis</h1>",
        unsafe_allow_html=True,
    )
    st.markdown(
        '<p style="color:#666;text-align:center;font-size:1.1rem;">'
        "This application is private. Please sign in to continue."
        "</p>",
        unsafe_allow_html=True,
    )
    _, mid, _ = st.columns([1, 1, 1])
    with mid:
        st.button("Sign in with Google", on_click=st.login, use_container_width=True)


if _auth_is_configured():
    if not st.user.is_logged_in:
        _render_login_screen()
        st.stop()
# ======================================================================
#  END AUTHENTICATION GATE
# ======================================================================
import io
import os
import re
import sys
import json
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st
import dateparser
import pdfplumber
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


# ======================================================================
#  CONFIGURATION  —  edit these two values before using the app
# ======================================================================

# Your DeepSeek API key (starts with "sk-")
DEEPSEEK_API_KEY = st.secrets.get("deepseek_api_key", "")

# Path to your Google Cloud Vision service account JSON key file
GOOGLE_VISION_KEY_PATH = r"C:\TWP_Build\contract_date_tool\google_credentials\google_vision_key.json"

# Path to Poppler (needed only for OCR of scanned PDFs)
POPPLER_PATH = None if os.name != "nt" else r"C:\poppler\Library\bin"


# ======================================================================
#  Patterns
# ======================================================================
DATE_PATTERNS = [
    r"\b\d{1,2}[/-]\d{1,2}[/-]\d{2,4}\b",
    r"\b\d{4}[/-]\d{1,2}[/-]\d{1,2}\b",
    r"\b(?:January|February|March|April|May|June|July|August|"
    r"September|October|November|December)\s+\d{1,2},?\s+\d{4}\b",
    r"\b\d{1,2}\s+(?:January|February|March|April|May|June|July|August|"
    r"September|October|November|December)\s+\d{4}\b",
]

DATE_KEYWORDS = {
    "Effective Date":        ["effective date", "commencement date", "starts on", "beginning on"],
    "Expiration Date":       ["expiration date", "expires on", "end date", "termination date", "conclude on"],
    "Renewal Date":          ["renewal date", "renew", "extension term", "automatically renew"],
    "Payment Due Date":      ["payment due", "due date", "invoice date", "payable on"],
    "Notice Deadline":       ["notice period", "written notice", "days' notice", "days notice"],
    "Termination Deadline":  ["termination", "cancel", "terminate"],
    "Signature Date":        ["signed on", "executed on", "date of signature", "entered into as of"],
    "Delivery Date":         ["delivery date", "deliver by", "delivery on"],
}

FREQUENCY_WORDS = {
    "weekly":       {"label": "Weekly",       "multiplier": 52},
    "bi-weekly":    {"label": "Bi-weekly",    "multiplier": 26},
    "biweekly":     {"label": "Bi-weekly",    "multiplier": 26},
    "semi-monthly": {"label": "Semi-monthly", "multiplier": 24},
    "monthly":      {"label": "Monthly",      "multiplier": 12},
    "quarterly":    {"label": "Quarterly",    "multiplier": 4},
    "semi-annually":{"label": "Semi-annually","multiplier": 2},
    "semiannually": {"label": "Semi-annually","multiplier": 2},
    "annually":     {"label": "Annually",     "multiplier": 1},
    "yearly":       {"label": "Annually",     "multiplier": 1},
    "daily":        {"label": "Daily",        "multiplier": 365},
    "hourly":       {"label": "Hourly",       "multiplier": 2080},
}

MONEY_FREQ_PATTERNS = [
    (r"\$\s?([\d,]+(?:\.\d{1,2})?)\s*/\s*(weekly|bi-?weekly|monthly|quarterly|"
     r"semi-?annually|annually|yearly|daily|hourly)", "slash"),
    (r"\$\s?([\d,]+(?:\.\d{1,2})?)\s+(?:per|a|an|each)\s+"
     r"(week|bi-?week|month|quarter|semi-?annual|year|annum|day|hour)\b", "per"),
    (r"\$\s?([\d,]+(?:\.\d{1,2})?)\s+(weekly|bi-?weekly|monthly|quarterly|"
     r"semi-?annually|annually|yearly|daily|hourly)\b", "after"),
]

DURATION_PATTERNS = [
    (r"\b(?:within|after|before|no later than|not later than|less than|"
     r"more than|at least)\s+(\d{1,4})\s+(day|week|month|year)s?\b", "relative"),
    (r"\b(\d{1,4})[-\s](day|week|month|year|hour)s?\s+"
     r"(?:period|term|notice|cure|window|deadline)\b", "compound"),
    (r"\b(?:for\s+a\s+period\s+of|for\s+a\s+term\s+of|for\s+the\s+term\s+of|"
     r"term\s+of)\s+(\d{1,4})\s+(day|week|month|year)s?\b", "duration"),
]

TIME_OF_DAY_PATTERN = re.compile(
    r"\b(\d{1,2}):(\d{2})\s*(AM|PM|am|pm)\b(?:\s+([A-Z][a-z]+))?",
    re.IGNORECASE
)

TERM_LANGUAGE = [
    "initial term", "renewal term", "option period", "notice period",
    "cure period", "grace period", "effective period", "term of the agreement",
]

BARE_AMOUNT_PATTERN = re.compile(r"\$\s?([\d,]+(?:\.\d{1,2})?)")


# ======================================================================
#  OCR  —  Tesseract first, Google Vision fallback
# ======================================================================
def tesseract_ocr(pdf_path):
    try:
        from pdf2image import convert_from_path
        import pytesseract
        import shutil

        tesseract_path = (
            shutil.which("tesseract")
            or r"C:\Program Files\Tesseract-OCR\tesseract.exe"
        )
        pytesseract.pytesseract.tesseract_cmd = tesseract_path

        images = convert_from_path(
            str(pdf_path), dpi=200, poppler_path=POPPLER_PATH
        )
        return "\n".join(pytesseract.image_to_string(img) for img in images)
    except Exception as e:
        st.warning(f"Tesseract OCR failed: {e}")
        return ""


def tesseract_output_is_poor(text):
    """Decide if Tesseract's output looks too poor to use."""
    if not text or len(text.strip()) < 100:
        return True

    stripped = text.strip()
    total_chars = len(stripped)
    if total_chars == 0:
        return True

    weird = sum(1 for c in stripped if not (c.isalnum() or c.isspace()
                                             or c in ".,;:!?()[]{}\"'/-&$%#@*+=<>"))
    if weird / total_chars > 0.15:
        return True

    words = re.findall(r"\b[A-Za-z]{2,}\b", stripped)
    if len(words) < 20:
        return True

    short_words = sum(1 for w in words if len(w) < 3)
    if len(words) > 0 and short_words / len(words) > 0.6:
        return True

    return False


def google_vision_ocr(pdf_path):
    """Run Google Cloud Vision OCR on a PDF (page by page)."""
    try:
        from google.cloud import vision
        from pdf2image import convert_from_path
        import io as _io

        os.environ["GOOGLE_APPLICATION_CREDENTIALS"] = GOOGLE_VISION_KEY_PATH
        client = vision.ImageAnnotatorClient()

        images = convert_from_path(
            str(pdf_path), dpi=200, poppler_path=POPPLER_PATH
        )

        all_text = []
        for img in images:
            buf = _io.BytesIO()
            img.save(buf, format="PNG")
            content = buf.getvalue()

            image = vision.Image(content=content)
            response = client.document_text_detection(image=image)

            if response.error.message:
                raise RuntimeError(response.error.message)

            all_text.append(response.full_text_annotation.text or "")

        return "\n".join(all_text)
    except Exception as e:
        st.error(f"Google Vision OCR failed: {e}")
        return ""


def ocr_pdf(pdf_path):
    """OCR a PDF: Tesseract first, Google Vision fallback if poor."""
    text = tesseract_ocr(pdf_path)

    if tesseract_output_is_poor(text):
        st.info("Tesseract output was low quality — falling back to Google Vision…")
        vision_text = google_vision_ocr(pdf_path)
        if vision_text.strip():
            return vision_text
        return text

    return text


# ======================================================================
#  Text extraction
# ======================================================================
def extract_text(file_bytes, filename):
    ext = Path(filename).suffix.lower()
    tmp = Path.cwd() / f"_uploaded{ext}"
    tmp.write_bytes(file_bytes)
    try:
        if ext == ".pdf":
            with pdfplumber.open(tmp) as pdf:
                text = "\n".join((p.extract_text() or "") for p in pdf.pages)
            if not text.strip():
                text = ocr_pdf(tmp)
            return text

        if ext == ".docx":
            return "\n".join(p.text for p in Document(tmp).paragraphs)

        return file_bytes.decode("utf-8", errors="ignore")
    finally:
        tmp.unlink(missing_ok=True)


# ======================================================================
#  DeepSeek summary
# ======================================================================
SUMMARY_SYSTEM_PROMPT = (
    "You are a document analyst. Read the provided document and return a "
    "JSON object with exactly these keys:\n"
    '  "document_type": a short string (e.g., "Contract", '
    '"Invoice", "Receipt", "Promissory Note", "NDA") — include a subtype '
    "in parentheses when clearly identifiable, e.g. \"Contract (Master "
    "Services Agreement)\".\n"
    '  "parties": a list of objects, each with "label" (e.g., "Party A") '
    'and "name" (the name and any role in parentheses, e.g. '
    '"Acme Corp (Client)"). List every distinct party you can identify.\n'
    '  "gist": a 2-3 sentence plain-English summary of the key terms.\n\n'
    "If a field cannot be determined from the document, set its value to "
    'the exact string "Not found in document". Return only valid JSON, '
    "no markdown fences, no commentary."
)


def deepseek_summarize(text, timeout=30):
    """Call DeepSeek to summarize the document."""
    if not DEEPSEEK_API_KEY or DEEPSEEK_API_KEY.startswith("sk-PASTE"):
        return None, "API key not configured"

    try:
        from openai import OpenAI
    except ImportError:
        return None, "openai package not installed"

    doc_text = text[:30000] if len(text) > 30000 else text

    try:
        client = OpenAI(
            api_key=DEEPSEEK_API_KEY,
            base_url="https://api.deepseek.com",
            timeout=timeout,
        )
        response = client.chat.completions.create(
            model="deepseek-chat",
            messages=[
                {"role": "system", "content": SUMMARY_SYSTEM_PROMPT},
                {"role": "user", "content": doc_text},
            ],
            temperature=0.1,
            response_format={"type": "json_object"},
        )
        raw = response.choices[0].message.content
        data = json.loads(raw)
        return data, None
    except Exception as e:
        msg = str(e)
        if "401" in msg or "Unauthorized" in msg:
            msg = "API key invalid or expired"
        elif "timeout" in msg.lower():
            msg = "request timed out"
        elif "Connection" in msg or "network" in msg.lower():
            msg = "no internet connection"
        elif "429" in msg or "rate" in msg.lower():
            msg = "rate limit exceeded"
        return None, msg


NOT_FOUND = "Not found in document"
UNAVAILABLE = "Summary unavailable"
UNABLE = "Unable to summarize this document"


def build_summary_section(text):
    """Return a dict with: status, doc_type, parties, gist, ai_status_line."""
    result = {
        "status": "ok",
        "doc_type": NOT_FOUND,
        "parties": [],
        "gist": NOT_FOUND,
        "ai_status_line": "",
    }

    summary, error = deepseek_summarize(text)

    if error:
        result["status"] = "unavailable"
        reason = f"{UNAVAILABLE} — {error}"
        result["doc_type"] = reason
        result["gist"] = reason
        result["ai_status_line"] = f"DeepSeek call failed: {error}"
        return result

    if not summary:
        result["status"] = "unable"
        result["doc_type"] = UNABLE
        result["gist"] = UNABLE
        result["ai_status_line"] = "DeepSeek returned no usable summary."
        return result

    result["doc_type"] = summary.get("document_type") or NOT_FOUND

    parties = summary.get("parties") or []
    if isinstance(parties, list):
        result["parties"] = [
            {"label": p.get("label", "Party"), "name": p.get("name", NOT_FOUND)}
            for p in parties if isinstance(p, dict)
        ]
    if not result["parties"]:
        result["parties"] = [{"label": "Party A", "name": NOT_FOUND}]

    result["gist"] = summary.get("gist") or NOT_FOUND
    result["ai_status_line"] = "Summary generated by DeepSeek V4.1-Flash."
    return result


# ======================================================================
#  Date / time detection
# ======================================================================
def _context(text, start, end, window=120):
    s = max(0, start - window)
    e = min(len(text), end + window)
    return text[s:e].replace("\n", " ").strip()


def _money_to_float(s):
    try:
        return float(s.replace(",", ""))
    except (ValueError, AttributeError):
        return None


def _format_money(value):
    if value is None:
        return ""
    return f"${value:,.2f}"


def _annualize(amount, freq_key):
    if amount is None or freq_key is None:
        return None
    spec = FREQUENCY_WORDS.get(freq_key.lower())
    if not spec:
        return None
    return amount * spec["multiplier"]


def _normalize_freq(word):
    w = word.lower().strip()
    aliases = {
        "week": "weekly", "biweek": "bi-weekly", "bi-week": "bi-weekly",
        "month": "monthly", "quarter": "quarterly",
        "semi-annual": "semi-annually", "semiannual": "semi-annually",
        "year": "annually", "annum": "annually",
        "day": "daily", "hour": "hourly",
    }
    if w in FREQUENCY_WORDS:
        return w
    return aliases.get(w)


def classify(context):
    ctx = context.lower()
    for category, keywords in DATE_KEYWORDS.items():
        if any(kw in ctx for kw in keywords):
            return category
    return "Other Date"


def find_dates(text, window=120):
    results = []
    for pattern in DATE_PATTERNS:
        for m in re.finditer(pattern, text, flags=re.IGNORECASE):
            raw = m.group(0)
            context = _context(text, m.start(), m.end(), window)
            parsed = dateparser.parse(raw, settings={"PREFER_DATES_FROM": "future"})
            results.append({
                "category": classify(context),
                "date": parsed.strftime("%Y-%m-%d") if parsed else "Unparsed",
                "amount": "", "frequency": "", "annualized": "",
                "raw": raw, "context": context,
            })
    return results


def find_time_references(text, window=120):
    results = []

    for pattern, _kind in MONEY_FREQ_PATTERNS:
        for m in re.finditer(pattern, text, flags=re.IGNORECASE):
            amount = _money_to_float(m.group(1))
            freq_key = _normalize_freq(m.group(2))
            freq_label = FREQUENCY_WORDS.get(freq_key, {}).get("label", m.group(2)) if freq_key else m.group(2)
            annual = _annualize(amount, freq_key)
            results.append({
                "category": "Payment Frequency", "date": "",
                "amount": _format_money(amount), "frequency": freq_label,
                "annualized": _format_money(annual) if annual else "",
                "raw": m.group(0), "context": _context(text, m.start(), m.end(), window),
            })

    for pattern, kind in DURATION_PATTERNS:
        for m in re.finditer(pattern, text, flags=re.IGNORECASE):
            num, unit = m.group(1), m.group(2).lower()
            cat = {"relative": "Relative Period", "compound": "Duration", "duration": "Term Length"}[kind]
            results.append({
                "category": cat, "date": "", "amount": "",
                "frequency": f"{num} {unit}{'s' if int(num) != 1 else ''}",
                "annualized": "", "raw": m.group(0),
                "context": _context(text, m.start(), m.end(), window),
            })

    for m in TIME_OF_DAY_PATTERN.finditer(text):
        hour, minute = m.group(1), m.group(2)
        meridiem = m.group(3).upper()
        tz = m.group(4) or ""
        results.append({
            "category": "Time of Day", "date": "", "amount": "",
            "frequency": f"{hour}:{minute} {meridiem}" + (f" {tz}" if tz else ""),
            "annualized": "", "raw": m.group(0),
            "context": _context(text, m.start(), m.end(), window),
        })

    for phrase in TERM_LANGUAGE:
        for m in re.finditer(re.escape(phrase), text, flags=re.IGNORECASE):
            results.append({
                "category": "Term Language", "date": "", "amount": "",
                "frequency": phrase, "annualized": "",
                "raw": m.group(0),
                "context": _context(text, m.start(), m.end(), window),
            })

    for m in BARE_AMOUNT_PATTERN.finditer(text):
        nearby = text[max(0, m.start() - 20):m.end() + 40].lower()
        if any(fw in nearby for fw in FREQUENCY_WORDS):
            continue
        if any(kw in nearby for kw in ["per ", "/week", "/month", "/year", "/hour", "/day"]):
            continue
        amount = _money_to_float(m.group(1))
        results.append({
            "category": "One-time Amount", "date": "",
            "amount": _format_money(amount), "frequency": "one-time",
            "annualized": "", "raw": m.group(0),
            "context": _context(text, m.start(), m.end(), window),
        })

    return results


def find_all_references(text, window=120):
    results = find_dates(text, window) + find_time_references(text, window)
    seen, unique = set(), []
    for r in results:
        key = (r["category"], r["raw"].lower())
        if key not in seen:
            seen.add(key)
            unique.append(r)
    return unique


def to_rows(results):
    def sort_key(r):
        has_date = r["date"] not in ("", "Unparsed")
        return (0 if has_date else 1, r["date"] or "zzz", r["category"])
    ordered = sorted(results, key=sort_key)
    rows = []
    for r in ordered:
        ctx = r["context"]
        if len(ctx) > 200:
            ctx = ctx[:200] + "…"
        rows.append({
            "Category": r["category"], "Date": r["date"],
            "Amount": r["amount"], "Frequency": r["frequency"],
            "Annualized": r["annualized"], "Raw Text": r["raw"],
            "Context": ctx,
        })
    return rows


# ======================================================================
#  Report builders
# ======================================================================
LABEL_COLOR = RGBColor(0x1F, 0x3A, 0x8A)
AMBER = RGBColor(0xB8, 0x6E, 0x00)
GRAY = RGBColor(0x80, 0x80, 0x80)
EM_DASH = "\u2014"


def _add_label_and_value(paragraph, label, value, value_color=None, italic=False):
    run_label = paragraph.add_run(f"{label}: ")
    run_label.bold = True
    run_label.font.color.rgb = LABEL_COLOR
    run_label.font.size = Pt(10)

    text = value if value not in ("", None) else EM_DASH
    run_value = paragraph.add_run(str(text))
    run_value.font.size = Pt(10)
    if value_color is not None:
        run_value.font.color.rgb = value_color
    if italic:
        run_value.italic = True


def _color_for_summary_value(value):
    if value is None:
        return GRAY, True
    v = str(value)
    if v.startswith(UNAVAILABLE):
        return AMBER, False
    if v == UNABLE or v == NOT_FOUND:
        return GRAY, True
    return None, False


def _add_horizontal_line(doc, color="BFBFBF"):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "6")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), color)
    pBdr.append(bottom)
    pPr.append(pBdr)
    return p


def build_docx(summary, rows, path, contract_name):
    doc = Document()

    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    doc.add_heading("TWP Contract Date and Time Analysis", level=0)

    p = doc.add_paragraph()
    p.add_run("Source contract: ").bold = True
    p.add_run(contract_name)
    doc.add_paragraph(f"Generated: {datetime.now():%Y-%m-%d %H:%M}")

    _add_horizontal_line(doc)

    # ---- Summary section ----
    doc.add_heading("Document Summary", level=1)

    dt_color, dt_italic = _color_for_summary_value(summary["doc_type"])
    dt_para = doc.add_paragraph()
    _add_label_and_value(dt_para, "Document Type", summary["doc_type"], dt_color, dt_italic)

    parties_para = doc.add_paragraph()
    label_run = parties_para.add_run("Parties:")
    label_run.bold = True
    label_run.font.color.rgb = LABEL_COLOR
    label_run.font.size = Pt(10)

    for p in summary["parties"]:
        pp = doc.add_paragraph()
        pp.paragraph_format.left_indent = Inches(0.25)
        name = p["name"]
        color, italic = _color_for_summary_value(name)
        label_run = pp.add_run(f"{p['label']}: ")
        label_run.bold = True
        label_run.font.color.rgb = LABEL_COLOR
        label_run.font.size = Pt(10)
        value_run = pp.add_run(str(name))
        value_run.font.size = Pt(10)
        if color is not None:
            value_run.font.color.rgb = color
        if italic:
            value_run.italic = True

    gist_color, gist_italic = _color_for_summary_value(summary["gist"])
    gist_para = doc.add_paragraph()
    _add_label_and_value(gist_para, "Gist of Terms", summary["gist"], gist_color, gist_italic)

    if summary.get("ai_status_line"):
        status_para = doc.add_paragraph()
        status_para.paragraph_format.space_before = Pt(2)
        status_run = status_para.add_run(summary["ai_status_line"])
        status_run.italic = True
        status_run.font.size = Pt(8)
        status_run.font.color.rgb = GRAY

    _add_horizontal_line(doc)

    # ---- Dates / terms section (card layout) ----
    doc.add_heading("Dates and Terms", level=1)

    for idx, row in enumerate(rows):
        table = doc.add_table(rows=3, cols=2)
        table.autofit = False
        for r in table.rows:
            r.cells[0].width = Inches(3.0)
            r.cells[1].width = Inches(3.5)

        cell_map = [("Category", "Date"), ("Amount", "Frequency"), ("Annualized", "Raw Text")]
        for row_idx, (left_label, right_label) in enumerate(cell_map):
            lc = table.rows[row_idx].cells[0]
            rc = table.rows[row_idx].cells[1]
            lc.paragraphs[0].text = ""
            rc.paragraphs[0].text = ""
            _add_label_and_value(lc.paragraphs[0], left_label, row.get(left_label, ""))
            _add_label_and_value(rc.paragraphs[0], right_label, row.get(right_label, ""))

        ctx_para = doc.add_paragraph()
        ctx_para.paragraph_format.space_before = Pt(6)
        ctx_para.paragraph_format.space_after = Pt(6)
        cl = ctx_para.add_run("Context: ")
        cl.bold = True
        cl.font.color.rgb = LABEL_COLOR
        cl.font.size = Pt(10)
        cv = ctx_para.add_run(row.get("Context", "") or EM_DASH)
        cv.font.size = Pt(10)

        if idx < len(rows) - 1:
            _add_horizontal_line(doc)
            doc.add_paragraph().paragraph_format.space_after = Pt(4)

    doc.save(str(path))


def build_xlsx(summary, rows, path):
    """Excel: Summary sheet + Data sheet."""
    from openpyxl import load_workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    df = pd.DataFrame(rows)
    with pd.ExcelWriter(str(path), engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Dates and Terms", index=False)

    wb = load_workbook(str(path))

    summary_ws = wb.create_sheet("Summary", 0)

    header_font = Font(bold=True, size=14, color="1F3A8A")
    label_font = Font(bold=True, size=11, color="1F3A8A")
    gray_italic = Font(size=10, italic=True, color="808080")
    amber = Font(size=10, color="B86E00")

    def style_value_font(value):
        if value is None:
            return gray_italic
        v = str(value)
        if v.startswith(UNAVAILABLE):
            return amber
        if v == UNABLE or v == NOT_FOUND:
            return gray_italic
        return Font(size=10)

    summary_ws["A1"] = "Document Summary"
    summary_ws["A1"].font = header_font

    summary_ws["A3"] = "Document Type"
    summary_ws["A3"].font = label_font
    summary_ws["B3"] = summary["doc_type"]
    summary_ws["B3"].font = style_value_font(summary["doc_type"])

    summary_ws["A5"] = "Parties"
    summary_ws["A5"].font = label_font
    row_cursor = 6
    for p in summary["parties"]:
        summary_ws.cell(row=row_cursor, column=1, value=p["label"]).font = label_font
        cell = summary_ws.cell(row=row_cursor, column=2, value=p["name"])
        cell.font = style_value_font(p["name"])
        row_cursor += 1

    row_cursor += 1
    summary_ws.cell(row=row_cursor, column=1, value="Gist of Terms").font = label_font
    gist_cell = summary_ws.cell(row=row_cursor, column=2, value=summary["gist"])
    gist_cell.font = style_value_font(summary["gist"])
    gist_cell.alignment = Alignment(wrap_text=True, vertical="top")
    summary_ws.column_dimensions["A"].width = 20
    summary_ws.column_dimensions["B"].width = 100

    if summary.get("ai_status_line"):
        row_cursor += 2
        status_cell = summary_ws.cell(row=row_cursor, column=1, value=summary["ai_status_line"])
        status_cell.font = Font(size=9, italic=True, color="808080")

    # Style the Data sheet
    data_ws = wb["Dates and Terms"]
    column_widths = {
        "Category": 22, "Date": 12, "Amount": 12,
        "Frequency": 16, "Annualized": 14,
        "Raw Text": 20, "Context": 80,
    }
    for col_idx, header in enumerate(df.columns, start=1):
        letter = get_column_letter(col_idx)
        data_ws.column_dimensions[letter].width = column_widths.get(header, 18)

    header_fill = PatternFill(start_color="1F3A8A", end_color="1F3A8A", fill_type="solid")
    for cell in data_ws[1]:
        cell.font = Font(bold=True, color="FFFFFF", size=11)
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    data_ws.row_dimensions[1].height = 28

    wrap_align = Alignment(wrap_text=True, vertical="top")
    plain_align = Alignment(vertical="top")
    thin = Side(border_style="thin", color="D9D9D9")
    light_border = Border(left=thin, right=thin, top=thin, bottom=thin)

    headers = list(df.columns)
    ctx_col_idx = headers.index("Context") + 1 if "Context" in headers else None
    raw_col_idx = headers.index("Raw Text") + 1 if "Raw Text" in headers else None

    for row_idx in range(2, data_ws.max_row + 1):
        for col_idx in range(1, len(headers) + 1):
            cell = data_ws.cell(row=row_idx, column=col_idx)
            cell.border = light_border
            if col_idx == ctx_col_idx or col_idx == raw_col_idx:
                cell.alignment = wrap_align
            else:
                cell.alignment = plain_align

    data_ws.freeze_panes = "A2"

    wb.save(str(path))


def build_markdown(summary, rows, contract_name):
    lines = [
        "# TWP Contract Date and Time Analysis", "",
        f"**Source contract:** {contract_name}  ",
        f"**Generated:** {datetime.now():%Y-%m-%d %H:%M}", "",
        "## Document Summary", "",
        f"- **Document Type:** {summary['doc_type']}",
    ]
    for p in summary["parties"]:
        lines.append(f"- **{p['label']}:** {p['name']}")
    lines.append(f"- **Gist of Terms:** {summary['gist']}")
    if summary.get("ai_status_line"):
        lines.append("")
        lines.append(f"*{summary['ai_status_line']}*")

    lines.extend(["", "## Dates and Terms", ""])

    headers = ["Category", "Date", "Amount", "Frequency", "Annualized", "Raw Text", "Context"]
    lines.append("| " + " | ".join(headers) + " |")
    lines.append("|" + "|".join(["---"] * len(headers)) + "|")
    for r in rows:
        cells = [str(r.get(h, "")).replace("|", "\\|") for h in headers]
        lines.append("| " + " | ".join(cells) + " |")
    return "\n".join(lines)


# ======================================================================
#  Streamlit UI
# ======================================================================
st.set_page_config(page_title="TWP Contract Date and Time Analysis", page_icon="📅", layout="wide")
st.markdown('<h1 style="color:#1a1a2e;">📅 TWP Contract Date and Time Analysis</h1>', unsafe_allow_html=True)
st.markdown(
    '<p style="color:#666;font-size:1.05rem;">'
    "Upload a contract and get a document summary plus a clean table of every "
    "important date and time reference — ready to download as Word, Excel, or Markdown."
    "</p>",
    unsafe_allow_html=True,
)

if "rows" not in st.session_state:
    st.session_state.rows = None
if "summary" not in st.session_state:
    st.session_state.summary = None
if "contract_name" not in st.session_state:
    st.session_state.contract_name = None

with st.sidebar:
    st.header("⚙️ Options")
    window = st.slider("Context window (characters)", 60, 300, 120, 20)
    show_unparsed = st.checkbox("Show rows with no date value", value=True)
    st.divider()
    st.markdown("**Supported files**")
    st.markdown("- PDF (`.pdf`)\n- Word (`.docx`)\n- Plain text (`.txt`, `.md`)")
    st.divider()
    st.caption(
        "⚠️ This tool is a helper, not legal advice. "
        "Always verify dates and amounts against the original contract."
    )

uploaded = st.file_uploader("Drop your contract here", type=["pdf", "docx", "txt", "md"])
_, col, _ = st.columns([1, 1, 1])
with col:
    run = st.button(
        "🔍 Analyze Document",
        type="primary",
        use_container_width=True,
        disabled=uploaded is None,
    )

if run and uploaded is not None:
    with st.spinner("Reading the document…"):
        try:
            text = extract_text(uploaded.getvalue(), uploaded.name)
            if not text.strip():
                st.error(
                    "Couldn't read any text from this file. "
                    "It may be a scanned PDF and OCR failed — check the message below."
                )
                st.stop()

            with st.spinner("Requesting document summary from DeepSeek…"):
                summary = build_summary_section(text)

            with st.spinner("Hunting for dates and terms…"):
                results = find_all_references(text, window=window)
                rows = to_rows(results)
                if not show_unparsed:
                    rows = [r for r in rows if r["Date"] not in ("", "Unparsed")]

            st.session_state.rows = rows
            st.session_state.summary = summary
            st.session_state.contract_name = uploaded.name
        except Exception as e:
            st.error(f"Something went wrong: {e}")
            st.stop()

if st.session_state.rows is not None and st.session_state.summary is not None:
    rows = st.session_state.rows
    summary = st.session_state.summary
    contract_name = st.session_state.contract_name

    # ---- Summary panel ----
    st.subheader("📄 Document Summary")
    with st.container(border=True):
        col1, col2 = st.columns([3, 2])

        with col1:
            st.markdown(f"**Document Type:** {summary['doc_type']}")
            st.markdown("**Parties:**")
            for p in summary["parties"]:
                st.markdown(f"&nbsp;&nbsp;&nbsp;&nbsp;**{p['label']}:** {p['name']}")

        with col2:
            st.markdown("**Gist of Terms:**")
            st.markdown(summary["gist"])

        if summary.get("ai_status_line"):
            st.caption(f"_{summary['ai_status_line']}_")

    st.divider()

    # ---- Data table ----
    st.subheader("📋 Dates and Terms")
    st.success(f"Found **{len(rows)}** reference(s) in **{contract_name}**.")

    df = pd.DataFrame(rows)
    dated = df.loc[df["Date"].notna() & (df["Date"] != "") & (df["Date"] != "Unparsed"), "Date"]
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total references", len(df))
    c2.metric("Categories", df["Category"].nunique())
    c3.metric("Earliest date", dated.min() if not dated.empty else "—")
    c4.metric("Latest date", dated.max() if not dated.empty else "—")

    cats = sorted(df["Category"].unique().tolist())
    selected = st.multiselect("Filter by category", options=cats, default=cats)
    filtered = df[df["Category"].isin(selected)] if selected else df.iloc[0:0]
    st.dataframe(filtered, use_container_width=True, hide_index=True)

    st.divider()
    st.subheader("⬇️ Download the report")
    rows_filtered = filtered.to_dict("records")
    stem = Path(contract_name).stem
    d1, d2, d3 = st.columns(3)

    docx_path = Path.cwd() / "_out.docx"
    build_docx(summary, rows_filtered, docx_path, contract_name)
    d1.download_button(
        "📄 Word (.docx)",
        data=docx_path.read_bytes(),
        file_name=f"{stem}_analysis.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True,
    )
    docx_path.unlink(missing_ok=True)

    xlsx_path = Path.cwd() / "_out.xlsx"
    build_xlsx(summary, rows_filtered, xlsx_path)
    d2.download_button(
        "📊 Excel (.xlsx)",
        data=xlsx_path.read_bytes(),
        file_name=f"{stem}_analysis.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
    xlsx_path.unlink(missing_ok=True)

    d3.download_button(
        "📝 Markdown (.md)",
        data=build_markdown(summary, rows_filtered, contract_name).encode("utf-8"),
        file_name=f"{stem}_analysis.md",
        mime="text/markdown",
        use_container_width=True,
    )

    st.divider()
    if st.button("🔄 Analyze another document"):
        st.session_state.rows = None
        st.session_state.summary = None
        st.session_state.contract_name = None
        st.rerun()

st.divider()
st.caption("Built with Streamlit • Files are processed locally except the summary, which is sent to DeepSeek.")
