import streamlit as st
import io
import zipfile
import os
from datetime import datetime

import extract_msg

from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm

# ==========================================
# Register Unicode Font (Fixes ■ black boxes)
# ==========================================
# Make sure DejaVuSans.ttf is in the same folder as app.py
pdfmetrics.registerFont(TTFont('DejaVu', 'DejaVuSans.ttf'))
pdfmetrics.registerFont(TTFont('DejaVu-Bold', 'DejaVuSans-Bold.ttf'))


# ==========================================
# Convert MSG → PDF (Unicode Safe)
# ==========================================
def msg_to_pdf_bytes(msg_bytes: bytes, filename: str) -> bytes:
    temp_msg_file = io.BytesIO(msg_bytes)
    temp_msg_file.name = filename

    msg = extract_msg.Message(temp_msg_file)

    sender = msg.sender or ""
    to = msg.to or ""
    cc = msg.cc or ""
    subject = msg.subject or ""
    date = msg.date or ""

    body = msg.body or ""
    if isinstance(body, bytes):
        body = body.decode("utf-8", errors="replace")

    # PDF buffer
    pdf_buffer = io.BytesIO()
    c = canvas.Canvas(pdf_buffer, pagesize=A4)
    width, height = A4

    margin_x = 20 * mm
    margin_y = 20 * mm

    y = height - margin_y

    # Header Title
    c.setFont("DejaVu-Bold", 14)
    c.drawString(margin_x, y, "Email Message")
    y -= 20

    # Email metadata
    c.setFont("DejaVu", 10)
    header_lines = [
        f"Subject: {subject}",
        f"From: {sender}",
        f"To: {to}",
    ]
    if cc:
        header_lines.append(f"CC: {cc}")
    if date:
        header_lines.append(f"Date: {date}")

    for line in header_lines:
        c.drawString(margin_x, y, line)
        y -= 14

    # Divider
    y -= 8
    c.line(margin_x, y, width - margin_x, y)
    y -= 20

    # Body Title
    c.setFont("DejaVu-Bold", 12)
    c.drawString(margin_x, y, "Body:")
    y -= 16

    # Multi-line body with Unicode wrapping
    c.setFont("DejaVu", 10)
    max_chars = 110
    line_height = 13

    for paragraph in body.splitlines():
        while len(paragraph) > max_chars:
            c.drawString(margin_x, y, paragraph[:max_chars])
            paragraph = paragraph[max_chars:]
            y -= line_height
            if y < margin_y:
                c.showPage()
                c.setFont("DejaVu", 10)
                y = height - margin_y

        c.drawString(margin_x, y, paragraph)
        y -= line_height

        if y < margin_y:
            c.showPage()
            c.setFont("DejaVu", 10)
            y = height - margin_y

    c.showPage()
    c.save()

    pdf_buffer.seek(0)
    return pdf_buffer.getvalue()


# ==========================================
# Streamlit App
# ==========================================
st.title("MSG → PDF Converter (ZIP → ZIP)")

st.write(
    "Upload a `.zip` containing `.msg` files. "
    "The app converts each email into a clean Unicode PDF and returns a ZIP of PDFs."
)

uploaded_zip = st.file_uploader("Upload ZIP with MSG files", type=["zip"])

if uploaded_zip is not None:
    if st.button("Convert to PDFs"):
        try:
            zip_bytes = uploaded_zip.read()
            input_zip = zipfile.ZipFile(io.BytesIO(zip_bytes))

            output_buffer = io.BytesIO()
            msg_count = 0

            with zipfile.ZipFile(output_buffer, "w", zipfile.ZIP_DEFLATED) as outzip:
                for item in input_zip.infolist():
                    if item.is_dir():
                        continue
                    if not item.filename.lower().endswith(".msg"):
                        continue

                    msg_count += 1
                    msg_data = input_zip.read(item.filename)

                    pdf_bytes = msg_to_pdf_bytes(msg_data, os.path.basename(item.filename))

                    base = os.path.splitext(os.path.basename(item.filename))[0]
                    outzip.writestr(f"{base}.pdf", pdf_bytes)

            if msg_count == 0:
                st.warning("No .msg files found in the ZIP.")
            else:
                output_buffer.seek(0)
                outname = f"converted_pdfs_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"

                st.success(f"Done! Converted {msg_count} email(s).")
                st.download_button(
                    "Download ZIP with PDFs",
                    data=output_buffer,
                    file_name=outname,
                    mime="application/zip"
                )

        except Exception as e:
            st.error(f"Error: {e}")
