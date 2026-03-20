def pdf_safe(text):
    """Replace unicode chars that crash fpdf2 helvetica font."""
    if not text:
        return ""
    replacements = {
        "\u2014": "--",   # em dash
        "\u2013": "-",    # en dash
        "\u2018": "'",    # left single quote
        "\u2019": "'",    # right single quote
        "\u201c": '"',    # left double quote
        "\u201d": '"',    # right double quote
        "\u2022": "-",    # bullet
        "\u2026": "...",  # ellipsis
        "\u25b8": ">",    # triangle
        "\u2713": "OK",   # checkmark
        "\u00a0": " ",    # non-breaking space
        "\u2192": "->",   # arrow
    }
    for char, rep in replacements.items():
        text = text.replace(char, rep)
    return text.encode("latin-1", errors="replace").decode("latin-1")


"""
report_gen.py — Generate formatted PDF and Word reports.
"""
from fpdf import FPDF
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from datetime import datetime
from io import BytesIO


LEVEL_COLORS = {
    "critical": (239, 68,  68),   # red
    "warning":  (245, 158, 11),   # amber
    "info":     (99,  102, 241),  # purple
    "pass":     (16,  185, 129),  # green
}
LEVEL_LABELS = {
    "critical": "CRITICAL",
    "warning":  "WARNING",
    "info":     "INFO",
    "pass":     "PASS",
}


# ── PDF ────────────────────────────────────────────────────────────────────────
class ReviewPDF(FPDF):
    def __init__(self, client_name, year_end):
        super().__init__()
        self.client_name = client_name
        self.year_end    = str(year_end)[:10] if year_end else ""
        self.set_auto_page_break(auto=True, margin=18)

    def header(self):
        self.set_font("Helvetica", "B", 9)
        self.set_fill_color(30, 27, 75)
        self.set_text_color(255, 255, 255)
        self.cell(0, 8, pdf_safe(f"  BOOKKEEPING REVIEW -- {self.client_name}  |  Year ended {self.year_end}"), fill=True, ln=True)
        self.set_text_color(0, 0, 0)
        self.ln(2)

    def footer(self):
        self.set_y(-12)
        self.set_font("Helvetica", "I", 7)
        self.set_text_color(150, 150, 150)
        self.cell(0, 6, f"Generated {datetime.now().strftime('%Y-%m-%d %H:%M')}  |  Page {self.page_no()}", align="C")

    def section_title(self, title):
        self.ln(3)
        self.set_font("Helvetica", "B", 11)
        self.set_fill_color(237, 237, 254)
        self.set_text_color(30, 27, 75)
        self.cell(0, 7, f"  {title}", fill=True, ln=True)
        self.set_text_color(0, 0, 0)
        self.ln(1)

    def check_row(self, level, title, detail):
        r, g, b = LEVEL_COLORS.get(level, (100, 100, 100))
        label    = LEVEL_LABELS.get(level, level.upper())

        # Badge
        self.set_font("Helvetica", "B", 7)
        self.set_fill_color(r, g, b)
        self.set_text_color(255, 255, 255)
        self.cell(18, 5, f" {label}", fill=True)
        self.set_text_color(0, 0, 0)

        # Title
        self.set_font("Helvetica", "B", 9)
        self.set_fill_color(250, 250, 252)
        remaining = self.w - self.l_margin - self.r_margin - 18
        self.cell(remaining, 5, pdf_safe(f"  {title[:95]}"), fill=True, ln=True)

        # Detail
        if detail and level != "pass":
            self.set_font("Helvetica", "", 8)
            self.set_text_color(80, 80, 80)
            self.set_x(self.l_margin + 20)
            self.multi_cell(remaining - 2, 4, pdf_safe(detail[:300]))
            self.set_text_color(0, 0, 0)
        self.ln(1)

    def ai_section(self, title, content):
        self.section_title(title)
        self.set_font("Helvetica", "", 9)
        self.set_text_color(30, 30, 30)
        # Split into lines and handle headers
        for line in content.split("\n"):
            line = line.strip()
            if not line:
                self.ln(2); continue
            if line.startswith("##") or (line.isupper() and len(line) > 3):
                self.set_font("Helvetica", "B", 9)
                self.cell(0, 5, pdf_safe(line.replace("#", "").strip()), ln=True)
                self.set_font("Helvetica", "", 9)
            elif line.startswith("**") and line.endswith("**"):
                self.set_font("Helvetica", "B", 9)
                self.cell(0, 5, pdf_safe(line.replace("**", "")), ln=True)
                self.set_font("Helvetica", "", 9)
            else:
                self.multi_cell(0, 4.5, pdf_safe(line[:200]))
        self.ln(3)


def generate_pdf(data, checks, ai_results):
    """Generate a complete PDF review report. Returns bytes."""
    pdf = ReviewPDF(data.get("client_name", ""), data.get("year_end", ""))
    pdf.add_page()

    # ── Cover info ─────────────────────────────────────────────────────────
    pdf.set_font("Helvetica", "B", 15)
    pdf.set_text_color(30, 27, 75)
    pdf.cell(0, 10, "BOOKKEEPING REVIEW REPORT", ln=True, align="C")
    pdf.set_font("Helvetica", "", 10)
    pdf.set_text_color(80, 80, 80)
    pdf.cell(0, 6, data.get("client_name", ""), ln=True, align="C")
    pdf.cell(0, 6, f"Year Ended: {str(data.get('year_end',''))[:10]}", ln=True, align="C")
    pdf.cell(0, 6, f"Prepared by: {data.get('prepared_by','')}  |  Signer: {data.get('signer','')}  |  Version: {data.get('version','')}", ln=True, align="C")
    pdf.ln(4)

    # ── Summary counts ──────────────────────────────────────────────────────
    from auto_checks import summarize_checks
    counts = summarize_checks(checks)
    pdf.set_font("Helvetica", "B", 10)
    pdf.set_text_color(0, 0, 0)
    summary = (
        f"  🔴 Critical: {counts['critical']}   "
        f"🟡 Warnings: {counts['warning']}   "
        f"🔵 Info: {counts['info']}   "
        f"✅ Passed: {counts['pass']}"
    )
    pdf.set_fill_color(245, 245, 255)
    pdf.cell(0, 8, summary, fill=True, ln=True)
    pdf.ln(3)

    # ── Automatic checks by category ───────────────────────────────────────
    pdf.section_title("AUTOMATIC CHECKS")
    from itertools import groupby
    # Show non-pass items first, then pass
    sorted_checks = (
        [c for c in checks if c["level"] != "pass"] +
        [c for c in checks if c["level"] == "pass"]
    )
    current_cat = None
    for c in sorted_checks:
        cat = c.get("category", "General")
        if cat != current_cat:
            pdf.set_font("Helvetica", "BI", 8)
            pdf.set_text_color(100, 100, 150)
            pdf.cell(0, 5, f"  ▸ {cat}", ln=True)
            pdf.set_text_color(0, 0, 0)
            current_cat = cat
        pdf.check_row(c["level"], c["title"], c.get("detail", ""))

    # ── AI sections ─────────────────────────────────────────────────────────
    prompt_titles = {
        "full_review":        "FULL FILE REVIEW",
        "tax_planning":       "TAX PLANNING OPPORTUNITIES",
        "missing_expenses":   "MISSING EXPENSES ANALYSIS",
        "staff_queries":      "STAFF QUERIES & CORRECTIONS",
        "management_summary": "MANAGEMENT SUMMARY (INTERNAL)",
        "client_summary":     "CLIENT EXECUTIVE SUMMARY",
        "engagement_notes":   "ENGAGEMENT FILE NOTES",
        "unusual_items":      "UNUSUAL ITEMS — CRA RISK FLAGS",
    }
    for key, content in ai_results.items():
        if content:
            pdf.add_page()
            pdf.ai_section(prompt_titles.get(key, key.upper()), content)

    buf = BytesIO()
    pdf.output(buf)
    buf.seek(0)
    return buf.read()


# ── Word ───────────────────────────────────────────────────────────────────────
def add_heading(doc, text, level=1):
    p = doc.add_heading(text, level=level)
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    return p


def add_check_para(doc, check):
    level  = check["level"]
    colors = {"critical": RGBColor(220, 38, 38), "warning": RGBColor(217, 119, 6),
              "info": RGBColor(79, 70, 229), "pass": RGBColor(5, 150, 105)}
    labels = LEVEL_LABELS
    p = doc.add_paragraph()
    run = p.add_run(f"[{labels.get(level, level.upper())}] ")
    run.bold = True
    run.font.color.rgb = colors.get(level, RGBColor(100, 100, 100))
    run.font.size = Pt(9)
    run2 = p.add_run(check["title"])
    run2.bold = True
    run2.font.size = Pt(9)
    if check.get("detail") and level != "pass":
        p2 = doc.add_paragraph(check["detail"])
        p2.runs[0].font.size = Pt(8)
        p2.runs[0].font.color.rgb = RGBColor(80, 80, 80)
        p2.paragraph_format.left_indent = Inches(0.3)
    p.paragraph_format.space_after = Pt(3)


def generate_word(data, checks, ai_results):
    """Generate a formatted Word document. Returns bytes."""
    doc = Document()

    # Margins
    for section in doc.sections:
        section.top_margin    = Inches(0.75)
        section.bottom_margin = Inches(0.75)
        section.left_margin   = Inches(0.9)
        section.right_margin  = Inches(0.9)

    # Title
    title = doc.add_heading("BOOKKEEPING REVIEW REPORT", 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for text in [
        data.get("client_name", ""),
        f"Year Ended: {str(data.get('year_end',''))[:10]}",
        f"Prepared by: {data.get('prepared_by','')}  |  Signer: {data.get('signer','')}",
        f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}",
    ]:
        run = p.add_run(text + "\n")
        run.font.size = Pt(10)

    doc.add_paragraph()

    # Summary counts
    from auto_checks import summarize_checks
    counts = summarize_checks(checks)
    p = doc.add_paragraph()
    for label, key, color in [
        ("Critical", "critical", RGBColor(220,38,38)),
        ("Warnings", "warning",  RGBColor(217,119,6)),
        ("Info",     "info",     RGBColor(79,70,229)),
        ("Passed",   "pass",     RGBColor(5,150,105)),
    ]:
        run = p.add_run(f"{label}: {counts[key]}    ")
        run.bold = True
        run.font.color.rgb = color
        run.font.size = Pt(11)
    doc.add_paragraph()

    # Automatic checks
    add_heading(doc, "AUTOMATIC CHECKS", 1)
    sorted_checks = (
        [c for c in checks if c["level"] != "pass"] +
        [c for c in checks if c["level"] == "pass"]
    )
    current_cat = None
    for c in sorted_checks:
        cat = c.get("category", "General")
        if cat != current_cat:
            p = doc.add_paragraph(f"▸ {cat}")
            p.runs[0].bold = True
            p.runs[0].font.size = Pt(9)
            p.runs[0].font.color.rgb = RGBColor(80, 80, 140)
            current_cat = cat
        add_check_para(doc, c)

    # AI sections
    prompt_titles = {
        "full_review":        "FULL FILE REVIEW",
        "tax_planning":       "TAX PLANNING OPPORTUNITIES",
        "missing_expenses":   "MISSING EXPENSES ANALYSIS",
        "staff_queries":      "STAFF QUERIES & CORRECTIONS",
        "management_summary": "MANAGEMENT SUMMARY (INTERNAL)",
        "client_summary":     "CLIENT EXECUTIVE SUMMARY",
        "engagement_notes":   "ENGAGEMENT FILE NOTES",
        "unusual_items":      "UNUSUAL ITEMS — CRA RISK FLAGS",
    }
    for key, content in ai_results.items():
        if content:
            doc.add_page_break()
            add_heading(doc, prompt_titles.get(key, key.upper()), 1)
            for line in content.split("\n"):
                line = line.strip()
                if not line: continue
                if line.startswith("**") and line.endswith("**"):
                    p = doc.add_paragraph(line.replace("**", ""))
                    p.runs[0].bold = True
                    p.runs[0].font.size = Pt(9)
                elif line.startswith("#"):
                    p = doc.add_paragraph(line.replace("#", "").strip())
                    p.runs[0].bold = True
                    p.runs[0].font.size = Pt(10)
                else:
                    p = doc.add_paragraph(line)
                    p.runs[0].font.size = Pt(9)

    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.read()
