import streamlit as st
import io
import zipfile
import os
from datetime import datetime
import extract_msg
from weasyprint import HTML, CSS
from weasyprint.text.fonts import FontConfiguration


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
# Convert MSG → PDF with HTML rendering
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

    # Create HTML document
    html_content = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8">
        <style>
            @page {{
                size: A4;
                margin: 2cm;
            }}
            body {{
                font-family: Arial, sans-serif;
                font-size: 10pt;
                line-height: 1.4;
                color: #000;
            }}
            .email-header {{
                background-color: #f5f5f5;
                padding: 15px;
                border: 1px solid #ddd;
                margin-bottom: 20px;
                border-radius: 5px;
            }}
            .email-header p {{
                margin: 5px 0;
            }}
            .email-header strong {{
                display: inline-block;
                width: 80px;
                color: #333;
            }}
            .email-body {{
                padding: 10px;
                border: 1px solid #eee;
                background-color: #fff;
            }}
            table {{
                border-collapse: collapse;
                width: 100%;
                margin: 10px 0;
            }}
            table td, table th {{
                border: 1px solid #ddd;
                padding: 8px;
                text-align: left;
            }}
            table th {{
                background-color: #f2f2f2;
                font-weight: bold;
            }}
            img {{
                max-width: 100%;
                height: auto;
            }}
        </style>
    </head>
    <body>
        <div class="email-header">
            <p><strong>Subject:</strong> {subject}</p>
            <p><strong>From:</strong> {sender}</p>
            <p><strong>To:</strong> {to}</p>
            {f'<p><strong>CC:</strong> {cc}</p>' if cc else ''}
            {f'<p><strong>Date:</strong> {date}</p>' if date else ''}
        </div>
        <div class="email-body">
    """

    if body_html:
        # Use HTML body if available
        html_content += body_html
    else:
        # Convert plain text to HTML
        body_html_converted = body_text.replace('\n', '<br/>')
        html_content += f"<div>{body_html_converted}</div>"

    html_content += """
        </div>
    </body>
    </html>
    """

    # Convert HTML to PDF using WeasyPrint
    font_config = FontConfiguration()
    pdf_bytes = HTML(string=html_content).write_pdf(font_config=font_config)

    return pdf_bytes


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
        import traceback
        st.error(f"Error: {e}")
        st.code(traceback.format_exc())
