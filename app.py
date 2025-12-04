import streamlit as st
import io
import zipfile
import os
from datetime import datetime

import extract_msg
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm


# ============================
# Utility: Convert one MSG to PDF (in memory)
# ============================

def msg_to_pdf_bytes(msg_bytes: bytes, filename: str) -> bytes:
    """
    Convert a single .msg file (as bytes) to a PDF (returned as bytes).
    Layout: header info + plain text body.
    """
    # MSG parser
    temp_msg_file = io.BytesIO(msg_bytes)
    temp_msg_file.name = filename  # extract_msg likes a .name

    msg = extract_msg.Message(temp_msg_file)
    msg_sender = msg.sender or ""
    msg_to = msg.to or ""
    msg_cc = msg.cc or ""
    msg_subject = msg.subject or ""
    msg_date = msg.date or ""
    body = msg.body or ""

    if isinstance(body, bytes):
        body = body.decode("utf-8", errors="ignore")

    # Create PDF in memory
    pdf_buffer = io.BytesIO()
    c = canvas.Canvas(pdf_buffer, pagesize=A4)
    width, height = A4

    x_margin = 20 * mm
    y_margin = 20 * mm
    max_width = width - 2 * x_margin

    def draw_multiline_text(text, start_y, font_name="Helvetica", font_size=10, leading=12):
        c.setFont(font_name, font_size)
        y = start_y
        for line in text.splitlines():
            while len(line) > 0:
                max_chars = 100  # simple wrapping by characters
                segment = line[:max_chars]
                line = line[max_chars:]
                c.drawString(x_margin, y, segment)
                y -= leading
                if y < y_margin:
                    c.showPage()
                    c.setFont(font_name, font_size)
                    y = height - y_margin
        return y

    y = height - y_margin

    # Header
    c.setFont("Helvetica-Bold", 14)
    c.drawString(x_margin, y, "Email Message")
    y -= 18

    c.setFont("Helvetica", 10)
    header_lines = [
        f"Subject: {msg_subject}",
        f"From: {msg_sender}",
        f"To: {msg_to}",
    ]
    if msg_cc:
        header_lines.append(f"CC: {msg_cc}")
    if msg_date:
        header_lines.append(f"Date: {msg_date}")

    for hl in header_lines:
        c.drawString(x_margin, y, hl)
        y -= 12

    # Divider
    y -= 10
    c.line(x_margin, y, x_margin + max_width, y)
    y -= 20

    # Body
    c.setFont("Helvetica-Bold", 12)
    c.drawString(x_margin, y, "Body:")
    y -= 16

    y = draw_multiline_text(body, y)

    c.showPage()
    c.save()

    pdf_buffer.seek(0)
    return pdf_buffer.getvalue()


# ============================
# Streamlit App
# ============================

st.title("MSG → PDF Converter (ZIP → ZIP)")

st.write(
    "Upload a `.zip` file containing Outlook `.msg` files. "
    "I'll convert each email to a PDF and give you back a `.zip` of PDFs."
)

uploaded_zip = st.file_uploader("Upload ZIP with .msg files", type=["zip"])

if uploaded_zip is not None:
    if st.button("Convert to PDFs"):
        try:
            # Read uploaded ZIP into memory
            zip_bytes = uploaded_zip.read()
            input_zip = zipfile.ZipFile(io.BytesIO(zip_bytes))

            # Prepare output ZIP in memory
            output_buffer = io.BytesIO()
            msg_count = 0

            with zipfile.ZipFile(output_buffer, "w", zipfile.ZIP_DEFLATED) as output_zip:
                for info in input_zip.infolist():
                    if info.is_dir():
                        continue

                    if not info.filename.lower().endswith(".msg"):
                        continue

                    msg_count += 1
                    msg_data = input_zip.read(info.filename)

                    pdf_bytes = msg_to_pdf_bytes(msg_data, os.path.basename(info.filename))

                    base_name = os.path.splitext(os.path.basename(info.filename))[0]
                    pdf_name = f"{base_name}.pdf"

                    output_zip.writestr(pdf_name, pdf_bytes)

            if msg_count == 0:
                st.warning("No .msg files found in the uploaded ZIP.")
            else:
                output_buffer.seek(0)
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                out_name = f"converted_pdfs_{timestamp}.zip"

                st.success(f"Conversion complete! Processed {msg_count} .msg files.")
                st.download_button(
                    label="Download ZIP with PDFs",
                    data=output_buffer,
                    file_name=out_name,
                    mime="application/zip",
                )

        except zipfile.BadZipFile:
            st.error("The uploaded file is not a valid ZIP archive.")
        except Exception as e:
            st.error(f"An error occurred: {e}")
