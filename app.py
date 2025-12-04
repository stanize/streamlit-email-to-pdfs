import streamlit as st
import io
import zipfile
import os
from datetime import datetime
import extract_msg
import re
from html.parser import HTMLParser

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont


# -----------------------------------------
# HTML to ReportLab converter
# -----------------------------------------
class HTMLToReportLab(HTMLParser):
    def __init__(self):
        super().__init__()
        self.text_parts = []
        self.in_table = False
        self.table_data = []
        self.current_row = []
        self.current_cell = []
        self.in_bold = False
        self.in_italic = False
        
    def handle_starttag(self, tag, attrs):
        if tag == 'table':
            self.in_table = True
            self.table_data = []
        elif tag == 'tr' and self.in_table:
            self.current_row = []
        elif tag in ['td', 'th'] and self.in_table:
            self.current_cell = []
        elif tag == 'br':
            self.text_parts.append('<br/>')
        elif tag == 'b' or tag == 'strong':
            self.in_bold = True
            self.text_parts.append('<b>')
        elif tag == 'i' or tag == 'em':
            self.in_italic = True
            self.text_parts.append('<i>')
        elif tag == 'p':
            self.text_parts.append('<br/>')
            
    def handle_endtag(self, tag):
        if tag == 'table':
            self.in_table = False
        elif tag == 'tr' and self.in_table:
            if self.current_row:
                self.table_data.append(self.current_row)
            self.current_row = []
        elif tag in ['td', 'th'] and self.in_table:
            cell_text = ''.join(self.current_cell).strip()
            self.current_row.append(cell_text)
            self.current_cell = []
        elif tag == 'b' or tag == 'strong':
            self.in_bold = False
            self.text_parts.append('</b>')
        elif tag == 'i' or tag == 'em':
            self.in_italic = False
            self.text_parts.append('</i>')
        elif tag == 'p':
            self.text_parts.append('<br/>')
            
    def handle_data(self, data):
        if self.in_table and self.current_cell is not None:
            self.current_cell.append(data)
        else:
            self.text_parts.append(data)
    
    def get_content(self):
        return ''.join(self.text_parts)
    
    def get_tables(self):
        return self.table_data


# -----------------------------------------
# Clean invisible / unsupported characters
# -----------------------------------------
def clean_text(t):
    if not t:
        return ""

    INVISIBLE_CHARS = [
        "\u200b", "\u200c", "\u200d", "\ufeff",
        "\u202a", "\u202b", "\u202c", "\u202d", "\u202e",
        "\u2060", "\u000b", "\u000c", "\u00ad",
    ]

    for ch in INVISIBLE_CHARS:
        t = t.replace(ch, "")

    return t


# -----------------------------------------
# Parse HTML and extract tables
# -----------------------------------------
def parse_html_body(html_body):
    parser = HTMLToReportLab()
    try:
        parser.feed(html_body)
    except:
        pass
    return parser.get_content(), parser.get_tables()


# -----------------------------------------
# Convert MSG → PDF using Platypus
# -----------------------------------------
def msg_to_pdf_bytes(msg_bytes: bytes, filename: str) -> bytes:
    temp_file = io.BytesIO(msg_bytes)
    temp_file.name = filename

    msg = extract_msg.Message(temp_file)

    # Convert all fields to strings explicitly
    sender = clean_text(str(msg.sender) if msg.sender else "")
    to = clean_text(str(msg.to) if msg.to else "")
    cc = clean_text(str(msg.cc) if msg.cc else "")
    subject = clean_text(str(msg.subject) if msg.subject else "")
    date = clean_text(str(msg.date) if msg.date else "")
    
    # Try to get HTML body first, fallback to plain text
    body_html = msg.htmlBody if hasattr(msg, 'htmlBody') and msg.htmlBody else None
    body_text = clean_text(str(msg.body) if msg.body else "")

    pdf_buffer = io.BytesIO()
    doc = SimpleDocTemplate(pdf_buffer, pagesize=A4, leftMargin=40, rightMargin=40, topMargin=40, bottomMargin=40)

    styles = getSampleStyleSheet()
    
    elements = []

    # Header
    elements.append(Paragraph("<b>Email Message</b>", styles["Heading1"]))
    elements.append(Spacer(width=1, height=12))

    info = f"""
        <b>Subject:</b> {subject}<br/>
        <b>From:</b> {sender}<br/>
        <b>To:</b> {to}<br/>
    """
    if cc:
        info += f"<b>CC:</b> {cc}<br/>"
    if date:
        info += f"<b>Date:</b> {date}<br/>"

    elements.append(Paragraph(info, styles["Normal"]))
    elements.append(Spacer(width=1, height=18))

    # Body
    elements.append(Paragraph("<b>Body:</b>", styles["Heading2"]))
    elements.append(Spacer(width=1, height=12))

    if body_html:
        # Parse HTML and extract tables
        parsed_text, tables = parse_html_body(body_html)
        
        # Add parsed text
        if parsed_text.strip():
            elements.append(Paragraph(parsed_text, styles["Normal"]))
            elements.append(Spacer(width=1, height=12))
        
        # Add tables
        for table_data in tables:
            if table_data and len(table_data) > 0:
                try:
                    t = Table(table_data)
                    t.setStyle(TableStyle([
                        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                        ('FONTSIZE', (0, 0), (-1, 0), 10),
                        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                        ('GRID', (0, 0), (-1, -1), 1, colors.black)
                    ]))
                    elements.append(t)
                    elements.append(Spacer(width=1, height=12))
                except:
                    pass
    else:
        # Convert plain text to HTML
        body_html_converted = body_text.replace('\n', '<br/>')
        elements.append(Paragraph(body_html_converted, styles["Normal"]))

    doc.build(elements)

    pdf_buffer.seek(0)
    return pdf_buffer.getvalue()


# -----------------------------------------
# Streamlit App
# -----------------------------------------
st.title("MSG → PDF Converter (ZIP → ZIP)")

st.markdown("""
Upload a ZIP file containing .msg email files, and this tool will convert them to PDFs 
with preserved formatting including tables, lists, and basic styling.
""")

uploaded_zip = st.file_uploader("Upload ZIP with MSG files:", type=["zip"])

if uploaded_zip is not None and st.button("Convert to PDFs"):
    with st.spinner("Converting emails to PDFs..."):
        try:
            zip_bytes = uploaded_zip.read()
            input_zip = zipfile.ZipFile(io.BytesIO(zip_bytes))

            output_buffer = io.BytesIO()
            msg_count = 0
            errors = []

            with zipfile.ZipFile(output_buffer, "w", zipfile.ZIP_DEFLATED) as outzip:
                for f in input_zip.infolist():
                    if f.is_dir() or not f.filename.lower().endswith(".msg"):
                        continue

                    try:
                        msg_count += 1
                        msg_data = input_zip.read(f.filename)
                        pdf_data = msg_to_pdf_bytes(msg_data, f.filename)

                        base = os.path.splitext(os.path.basename(f.filename))[0]
                        outzip.writestr(f"{base}.pdf", pdf_data)
                    except Exception as e:
                        errors.append(f"{f.filename}: {str(e)}")

            if msg_count == 0:
                st.warning("No .msg files found in ZIP.")
            else:
                output_buffer.seek(0)
                outname = f"converted_pdfs_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"

                st.success(f"✅ Converted {msg_count} emails successfully!")
                
                if errors:
                    st.warning(f"⚠️ {len(errors)} file(s) had errors:")
                    for error in errors:
                        st.text(error)
                
                st.download_button(
                    "⬇️ Download ZIP of PDFs",
                    data=output_buffer,
                    file_name=outname,
                    mime="application/zip"
                )

        except Exception as e:
            import traceback
            st.error(f"❌ Error: {e}")
            with st.expander("Show full error details"):
                st.code(traceback.format_exc())
