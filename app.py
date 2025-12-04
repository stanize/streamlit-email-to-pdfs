import streamlit as st
import io
import zipfile
import os
from datetime import datetime
import olefile
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib import colors
import pdfkit
import tempfile


# -----------------------------------------
# MSG Property Tags
# -----------------------------------------
PR_SUBJECT = 0x0037001F
PR_SENDER_NAME = 0x0C1A001F
PR_SENT_REPRESENTING_NAME = 0x0042001F
PR_DISPLAY_TO = 0x0E04001F
PR_DISPLAY_CC = 0x0E03001F
PR_CLIENT_SUBMIT_TIME = 0x00390040
PR_BODY = 0x1000001F
PR_HTML = 0x10130102


# -----------------------------------------
# Parse MSG file manually
# -----------------------------------------
def parse_msg_file(msg_bytes):
    """Parse MSG file using olefile directly"""
    ole = olefile.OleFileIO(msg_bytes)
    
    def get_property(prop_id):
        """Get property from MSG file"""
        prop_tag = f"{prop_id:08X}"
        
        possible_names = [
            f"__substg1.0_{prop_tag}",
            f"__properties_version1.0/{prop_tag}",
        ]
        
        for name in possible_names:
            if ole.exists(name):
                try:
                    data = ole.openstream(name).read()
                    if prop_tag.endswith("001F"):
                        try:
                            return data.decode('utf-16-le').rstrip('\x00')
                        except:
                            return data.decode('utf-8', errors='ignore').rstrip('\x00')
                    elif prop_tag.endswith("0102"):
                        return data
                    elif prop_tag.endswith("0040"):
                        return data.decode('utf-8', errors='ignore')
                    return data.decode('utf-8', errors='ignore')
                except:
                    pass
        return ""
    
    msg_data = {
        'subject': get_property(PR_SUBJECT),
        'sender': get_property(PR_SENDER_NAME) or get_property(PR_SENT_REPRESENTING_NAME),
        'to': get_property(PR_DISPLAY_TO),
        'cc': get_property(PR_DISPLAY_CC),
        'date': get_property(PR_CLIENT_SUBMIT_TIME),
        'body': get_property(PR_BODY),
        'html': get_property(PR_HTML),
    }
    
    ole.close()
    return msg_data


# -----------------------------------------
# Clean text
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
# Convert MSG → PDF using wkhtmltopdf
# -----------------------------------------
def msg_to_pdf_bytes(msg_bytes: bytes, filename: str) -> bytes:
    # Parse MSG file
    msg_data = parse_msg_file(msg_bytes)
    
    sender = clean_text(str(msg_data.get('sender', '')))
    to = clean_text(str(msg_data.get('to', '')))
    cc = clean_text(str(msg_data.get('cc', '')))
    subject = clean_text(str(msg_data.get('subject', '')))
    date = clean_text(str(msg_data.get('date', '')))
    body_html = msg_data.get('html')
    body_text = clean_text(str(msg_data.get('body', '')))

    # Create complete HTML document with email header
    html_content = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8">
        <style>
            body {{
                font-family: Arial, sans-serif;
                margin: 20px;
            }}
            .email-header {{
                background-color: #f5f5f5;
                padding: 15px;
                border: 1px solid #ddd;
                margin-bottom: 20px;
                font-size: 11pt;
            }}
            .email-header p {{
                margin: 5px 0;
            }}
            .email-header strong {{
                display: inline-block;
                width: 80px;
            }}
            .email-body {{
                margin-top: 20px;
            }}
        </style>
    </head>
    <body>
        <div class="email-header">
            <p><strong>Subject:</strong> {subject}</p>
            <p><strong>From:</strong> {sender}</p>
            <p><strong>To:</strong> {to}</p>
    """
    
    if cc:
        html_content += f'<p><strong>CC:</strong> {cc}</p>'
    if date:
        html_content += f'<p><strong>Date:</strong> {date}</p>'
    
    html_content += '</div><div class="email-body">'
    
    if body_html:
        # Use the HTML body directly
        if isinstance(body_html, bytes):
            body_html = body_html.decode('utf-8', errors='ignore')
        html_content += body_html
    else:
        # Fallback to plain text
        body_html_converted = body_text.replace('\n', '<br/>')
        html_content += f'<pre style="font-family: Arial; white-space: pre-wrap;">{body_html_converted}</pre>'
    
    html_content += '</div></body></html>'
    
    # Convert HTML to PDF using pdfkit
    options = {
        'page-size': 'A4',
        'margin-top': '0.75in',
        'margin-right': '0.75in',
        'margin-bottom': '0.75in',
        'margin-left': '0.75in',
        'encoding': "UTF-8",
        'no-outline': None,
        'enable-local-file-access': None
    }
    
    try:
        pdf_bytes = pdfkit.from_string(html_content, False, options=options)
        return pdf_bytes
    except Exception as e:
        st.error(f"Error converting {filename}: {str(e)}")
        raise


# -----------------------------------------
# Streamlit App
# -----------------------------------------
st.title("MSG → PDF Converter (ZIP → ZIP)")

st.markdown("""
Upload a ZIP file containing .msg email files, and this tool will convert them to PDFs 
with **full HTML formatting preserved** including colors, tables, fonts, and layout.
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
