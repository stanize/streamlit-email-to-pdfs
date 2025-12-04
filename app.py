import streamlit as st
import io
import zipfile
import os
from datetime import datetime
import extract_msg

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont


# -----------------------------------------
# Register full unicode fonts
# -----------------------------------------
pdfmetrics.registerFont(TTFont('DejaVu', 'DejaVuSans.ttf'))
pdfmetrics.registerFont(TTFont('DejaVu-Bold', 'DejaVuSans-Bold.ttf'))


# -----------------------------------------
# Clean invisible / unsupported characters
# -----------------------------------------
def clean_text(t):
    if not t:
        return ""

    INVISIBLE_CHARS = [
        "\u200b",  # zero width space
        "\u200c",  # zero width non-joiner
        "\u200d",  # zero width joiner
        "\ufeff",  # zero width no-break space
        "\u202a", "\u202b", "\u202c", "\u202d", "\u202e",  # directional marks
        "\u2060",  # word joiner
        "\u000b",  # vertical tab
        "\u000c",  # form feed
        "\u00ad",  # soft hyphen
    ]

    for ch in INVISIBLE_CHARS:
        t = t.replace(ch, "")

    return t


# -----------------------------------------
# Convert MSG → PDF using Platypus
# -----------------------------------------
def msg_to_pdf_bytes(msg_bytes: bytes, filename: str) -> bytes:
    temp_file = io.BytesIO(msg_bytes)
    temp_file.name = filename

    msg = extract_msg.Message(temp_file)

    sender = clean_text(msg.sender or "")
    to = clean_text(msg.to or "")
    cc = clean_text(msg.cc or "")
    subject = clean_text(msg.subject or "")
    date = clean_text(msg.date or "")
    body = clean_text(msg.body or "")

    pdf_buffer = io.BytesIO()
    doc = SimpleDocTemplate(pdf_buffer, pagesize=A4, leftMargin=40, rightMargin=40, topMargin=40, bottomMargin=40)

    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="NormalUnicode", fontName="DejaVu", fontSize=10, leading=14))
    styles.add(ParagraphStyle(name="HeaderUnicode", fontName="DejaVu-Bold", fontSize=12, leading=16))

    elements = []

    # Header
    elements.append(Paragraph("<b>Email Message</b>", styles["HeaderUnicode"]))
    elements.append(Spacer(1 * inch, 0.15 * inch))

    info = f"""
        <b>Subject:</b> {subject}<br/>
        <b>From:</b> {sender}<br/>
        <b>To:</b> {to}<br/>
    """
    if cc:
        info += f"<b>CC:</b> {cc}<br/>"

    if date:
        info += f"<b>Date:</b> {date}<br/>"

    elements.append(Paragraph(info, styles["NormalUnicode"]))
    elements.append(Spacer(1 * inch, 0.25 * inch))

    # Body
    elements.append(Paragraph("<b>Body:</b>", styles["HeaderUnicode"]))
    elements.append(Spacer(1 * inch, 0.15 * inch))

    # Convert \n to <br/> for Paragraph
    body_html = body.replace("\n", "<br/>")
    elements.append(Paragraph(body_html, styles["NormalUnicode"]))

    doc.build(elements)

    pdf_buffer.seek(0)
    return pdf_buffer.getvalue()


# -----------------------------------------
# Streamlit App
# -----------------------------------------
st.title("MSG → PDF Converter (ZIP → ZIP) — Perfect Formatting Version")

uploaded_zip = st.file_uploader("Upload ZIP with MSG files:", type=["zip"])

if uploaded_zip is not None and st.button("Convert to PDFs"):
    try:
        zip_bytes = uploaded_zip.read()
        input_zip = zipfile.ZipFile(io.BytesIO(zip_bytes))

        output_buffer = io.BytesIO()
        msg_count = 0

        with zipfile.ZipFile(output_buffer, "w", zipfile.ZIP_DEFLATED) as outzip:
            for f in input_zip.infolist():
                if f.is_dir() or not f.filename.lower().endswith(".msg"):
                    continue

                msg_count += 1
                msg_data = input_zip.read(f.filename)
                pdf_data = msg_to_pdf_bytes(msg_data, f.filename)

                base = os.path.splitext(os.path.basename(f.filename))[0]
                outzip.writestr(f"{base}.pdf", pdf_data)

        if msg_count == 0:
            st.warning("No .msg files found in ZIP.")
        else:
            output_buffer.seek(0)
            outname = f"converted_pdfs_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"

            st.success(f"Converted {msg_count} emails successfully!")
            st.download_button(
                "Download ZIP of PDFs",
                data=output_buffer,
                file_name=outname,
                mime="application/zip"
            )

    except Exception as e:
        st.error(f"Error: {e}")
