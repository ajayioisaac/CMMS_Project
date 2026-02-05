import os
import re
import pandas as pd
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib.enums import TA_JUSTIFY


# ===================== CONFIG =====================
EXCEL_FILE = r"D:/CESL/CMMS Issues/Refined/Preventive_Mainenance_data_2026-01-31T08_29_13.606Z.xlsx"
SHEET_NAME = "Preventive_Mainenance_data_2026"
OUTPUT_DIR = r"D:/CESL/CMMS Issues/Refined/Procedure_PDFs"

START_STEP_COL = 12   # M
END_STEP_COL = 71     # BT
STEP_GROUP_SIZE = 3
# =================================================

os.makedirs(OUTPUT_DIR, exist_ok=True)


def safe_str(value):
    """Convert NaN / None / float safely to string"""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    return str(value).strip()


def split_itemized_text(text):
    """Split numbered descriptions into readable sub-steps"""
    text = safe_str(text)
    if not text:
        return []

    text = " ".join(text.split())
    parts = re.split(r'(?<=\.)\s*(?=\d+\.)', text)
    return [p.strip() for p in parts if p.strip()]


def header_footer(canvas, doc, procedure_title):
    canvas.saveState()
    canvas.setFont("Helvetica-Bold", 10)
    canvas.drawString(20 * mm, 285 * mm, f"Procedure: {procedure_title}")
    canvas.restoreState()


def create_procedure_pdf(proc_title, proc_desc, total_duration, steps):
    safe_title = re.sub(r'[\\/*?:"<>|]', "_", proc_title)
    pdf_path = os.path.join(OUTPUT_DIR, f"{safe_title}.pdf")

    doc = SimpleDocTemplate(
        pdf_path,
        pagesize=A4,
        rightMargin=20 * mm,
        leftMargin=20 * mm,
        topMargin=25 * mm,
        bottomMargin=20 * mm
    )

    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        "Title",
        parent=styles["Title"],
        spaceAfter=12
    )

    section_style = ParagraphStyle(
        "Section",
        parent=styles["Heading2"],
        spaceBefore=12
    )

    body_style = ParagraphStyle(
        "Body",
        parent=styles["BodyText"],
        spaceAfter=6,
        alignment=TA_JUSTIFY
    )

    step_title_style = ParagraphStyle(
        "StepTitle",
        parent=styles["Heading4"],
        spaceBefore=10
    )

    substep_style = ParagraphStyle(
        "SubStep",
        parent=styles["BodyText"],
        leftIndent=15,
        spaceAfter=4,
        alignment=TA_JUSTIFY
    )

    story = []

    # ---- Header ----
    story.append(Paragraph(safe_str(proc_title), title_style))
    story.append(Paragraph(f"<b>Description:</b> {safe_str(proc_desc)}", body_style))
    story.append(Paragraph(
        f"<b>Total Procedure Duration:</b> {safe_str(total_duration)}",
        body_style
    ))
    story.append(Spacer(1, 12))

    # ---- Steps ----
    story.append(Paragraph("Procedure Steps", section_style))

    for idx, step in enumerate(steps, start=1):
        story.append(
            Paragraph(f"{idx}. {safe_str(step['title'])}", step_title_style)
        )

        sub_steps = split_itemized_text(step["description"])
        if sub_steps:
            for sub in sub_steps:
                story.append(Paragraph(sub, substep_style))
        else:
            story.append(Paragraph(safe_str(step["description"]), substep_style))

        story.append(
            Paragraph(
                f"<b>Duration:</b> {safe_str(step['duration'])}",
                substep_style
            )
        )

    doc.build(
        story,
        onFirstPage=lambda c, d: header_footer(c, d, proc_title),
        onLaterPages=lambda c, d: header_footer(c, d, proc_title)
    )


def generate_procedure_pdfs():
    df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME)

    created = set()

    for _, row in df.iterrows():
        proc_title = safe_str(row.iloc[10])   # K
        proc_desc = safe_str(row.iloc[11])    # L
        total_duration = safe_str(row.iloc[72])  # BU

        if not proc_title or proc_title in created:
            continue

        steps = []

        for col in range(START_STEP_COL, END_STEP_COL + 1, STEP_GROUP_SIZE):
            step_title = safe_str(row.iloc[col])
            step_desc = safe_str(row.iloc[col + 1])
            step_duration = safe_str(row.iloc[col + 2])

            if not step_title and not step_desc:
                continue

            steps.append({
                "title": step_title,
                "description": step_desc,
                "duration": step_duration
            })

        if steps:
            create_procedure_pdf(proc_title, proc_desc, total_duration, steps)
            created.add(proc_title)


if __name__ == "__main__":
    generate_procedure_pdfs()
