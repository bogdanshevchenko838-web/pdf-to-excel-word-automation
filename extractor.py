import fitz  # PyMuPDF
from pathlib import Path
from typing import List


def extract_text_from_pdf(pdf_path: Path | str) -> List[str]:
    pdf_path = Path(pdf_path)
    doc = fitz.open(pdf_path)
    pages_text: List[str] = []

    for page in doc:
        text = page.get_text("text")
        pages_text.append(text)

    doc.close()
    return pages_text
