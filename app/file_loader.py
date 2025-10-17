# file_loader.py
import io
import zipfile
import pandas as pd
from PyPDF2 import PdfReader
from docx import Document

def extract_text_from_pdf_bytes(b: bytes) -> str:
    text = ""
    pdf = PdfReader(io.BytesIO(b))
    for page in pdf.pages:
        text += page.extract_text() or ""
    return text

def extract_text_from_docx_bytes(b: bytes) -> str:
    doc = Document(io.BytesIO(b))
    return "\n".join([p.text for p in doc.paragraphs])

def extract_text_from_excel_bytes(b: bytes) -> str:
    df = pd.read_excel(io.BytesIO(b))
    text = ""
    for i, row in df.iterrows():
        row_text = " | ".join(f"{col}: {row[col]}" for col in df.columns)
        text += f"Row {i+1}: {row_text}\n"
    return text

def extract_text_from_zip_bytes(b: bytes) -> str:
    text = ""
    with zipfile.ZipFile(io.BytesIO(b)) as z:
        for filename in z.namelist():
            with z.open(filename) as f:
                inner = f.read()
                text += get_raw_text(inner, filename) + "\n\n"
    return text

def get_raw_text(file_bytes: bytes, filename: str) -> str:
    filename = filename.lower()
    if filename.endswith(".pdf"):
        return extract_text_from_pdf_bytes(file_bytes)
    elif filename.endswith(".docx"):
        return extract_text_from_docx_bytes(file_bytes)
    elif filename.endswith((".xlsx", ".xls")):
        return extract_text_from_excel_bytes(file_bytes)
    elif filename.endswith(".zip"):
        return extract_text_from_zip_bytes(file_bytes)
    else:
        # fallback: try to decode as text
        try:
            return file_bytes.decode("utf-8", errors="ignore")
        except Exception:
            return ""