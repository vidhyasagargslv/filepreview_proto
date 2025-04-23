import os
import mammoth 
from docx import Document as DocxDocument
import base64
import tempfile
import pdfplumber
import mammoth
import pdfplumber
from bs4 import BeautifulSoup
import io
from pdf2image import convert_from_path

def convert_docx_to_html(file_path):
    """Convert DOCX to HTML using Mammoth with enhanced styling options"""
    try:
        # Custom style map for better formatting preservation
        style_map = """
        p[style-name='Heading 1'] => h1:fresh
        p[style-name='Heading 2'] => h2:fresh
        p[style-name='Heading 3'] => h3:fresh
        p[style-name='Heading 4'] => h4:fresh
        r[style-name='Strong'] => strong
        table => table.docx-table
        """
        
        with open(file_path, 'rb') as docx_file:
            # Convert with custom style map
            result = mammoth.convert_to_html(docx_file, style_map=style_map)
            html = result.value
            
            # Add CSS for better table styling
            css = """
            <style>
            .docx-table {
                border-collapse: collapse;
                width: 100%;
                margin-bottom: 1em;
            }
            .docx-table td, .docx-table th {
                border: 1px solid #ddd;
                padding: 8px;
            }
            pre {
                white-space: pre-wrap;
                font-family: inherit;
            }
            </style>
            """
            
            # Process the HTML with BeautifulSoup for additional cleanup
            soup = BeautifulSoup(html, 'html.parser')
            
            # Wrap the content in proper HTML structure
            full_html = f"<!DOCTYPE html><html><head>{css}</head><body>{soup}</body></html>"
            
            return full_html
    except Exception as e:
        print(f"Error converting DOCX to HTML: {e}")
        return None

def convert_pdf_to_html(file_path):
    """Convert PDF to HTML with better formatting preservation, including line spaces."""
    try:
        html = ["<!DOCTYPE html><html><head>",
                "<style>",
                ".pdf-page { margin-bottom: 20px; position: relative; }",
                ".pdf-text { position: relative; z-index: 1; }",
                ".pdf-table { border-collapse: collapse; width: 100%; margin: 10px 0; }",
                ".pdf-table td, .pdf-table th { border: 1px solid #ddd; padding: 8px; }",
                "</style>",
                "</head><body>"]
        
        with pdfplumber.open(file_path) as pdf:
            for page_num, page in enumerate(pdf.pages):
                html.append(f'<div class="pdf-page" id="page-{page_num+1}">')
                
                # Extract text with positioning information
                text = page.extract_text(x_tolerance=3, y_tolerance=3)
                if text:
                    # Convert newlines to proper HTML breaks
                    paragraphs = text.split('\n\n')  # Split by double newlines for paragraphs
                    for para in paragraphs:
                        if para.strip():
                            # Convert remaining single newlines to <br> tags
                            para_html = para.replace('\n', '<br>')
                            html.append(f'<div class="pdf-text"><p>{para_html}</p></div>')
                
                # Extract tables
                tables = page.extract_tables()
                for table in tables:
                    if table:
                        html.append('<table class="pdf-table">')
                        for row in table:
                            html.append('<tr>')
                            for cell in row:
                                cell_content = "" if cell is None else cell.replace('\n', '<br>')
                                html.append(f'<td>{cell_content}</td>')
                            html.append('</tr>')
                        html.append('</table>')
                
                # Handle images - extract images using pdf2image
                if page_num == 0:  # For demonstration, process only first page images
                    try:
                        images = convert_from_path(file_path, first_page=page_num+1, last_page=page_num+1)
                        for img in images:
                            # Save image to buffer
                            img_buffer = io.BytesIO()
                            img.save(img_buffer, format="PNG")
                            img_str = base64.b64encode(img_buffer.getvalue()).decode('utf-8')
                            # Add image to HTML with proper positioning
                            html.append(f'<div class="pdf-image"><img src="data:image/png;base64,{img_str}" alt="Page image"></div>')
                    except Exception as e:
                        print(f"Error extracting images: {e}")
                
                html.append('</div>')  # Close page div
        
        html.append("</body></html>")
        return "".join(html)
    except Exception as e:
        print(f"Error converting PDF to HTML: {e}")
        return None


def convert_html_from_string(html_content):
    """Process HTML content directly (for HTML file uploads or text edits)"""
    return html_content

def update_document_html(document_content, file_type=None):
    """Process document content and return HTML"""
    if file_type == "docx":
        with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as temp_file:
            temp_file.write(document_content)
            temp_path = temp_file.name
        html_content = convert_docx_to_html(temp_path)
        os.unlink(temp_path)
        return html_content
    elif file_type == "pdf":
        with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as temp_file:
            temp_file.write(document_content)
            temp_path = temp_file.name
        html_content = convert_pdf_to_html(temp_path)
        os.unlink(temp_path)
        return html_content
    else:
        return convert_html_from_string(document_content)