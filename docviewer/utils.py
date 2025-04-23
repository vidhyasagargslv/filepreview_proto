import os
import mammoth 
from docx import Document as DocxDocument
import base64
import tempfile

def convert_docx_to_html(file_path):
    """Convert DOCX to HTML using Mammoth"""
    try:
        with open(file_path, 'rb') as docx_file:
            result = mammoth.convert_to_html(docx_file)
            return result.value
    except Exception as e:
        print(f"Error converting DOCX to HTML: {e}")
        return None

def convert_html_from_string(html_content):
    """Process HTML content directly (for HTML file uploads or text edits)"""
    # You could add additional processing here if needed
    return html_content

def update_document_html(document_content, is_docx=False):
    """Process document content and return HTML"""
    if is_docx:
        # Create a temporary file for the DOCX content
        with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as temp_file:
            temp_file.write(document_content)
            temp_path = temp_file.name
        
        # Convert the temporary DOCX file to HTML
        html_content = convert_docx_to_html(temp_path)
        
        # Clean up the temporary file
        os.unlink(temp_path)
        return html_content
    else:
        # Process HTML content directly
        return convert_html_from_string(document_content)
