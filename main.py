import sys
import json
from pathlib import Path

from extractor import extract_text_from_pdf
from parsers.demo_parser import parse_invoice_demo
from normalize_invoice import normalize_invoice
from writers.excel_writer import save_to_excel
from writers.word_writer import generate_word
from config import OUTPUT_DIR


def process_single_pdf(pdf_path: Path):
    print(f"\n[+] Processing: {pdf_path}")

    # 1. PDF -> list of pages
    pages = extract_text_from_pdf(pdf_path)
    full_text = "\n".join(pages)

    # 2. DEMO parser
    parsed_demo = parse_invoice_demo(full_text)

    # 3. basic + raw_text 
    basic = {
        "invoice_number": parsed_demo.get("invoice_number"),
        "order_number": parsed_demo.get("order_number"),
        "invoice_date": parsed_demo.get("invoice_date"),
        "due_date": parsed_demo.get("due_date"),
        "raw_text": full_text,
    }

    data = normalize_invoice(basic, parsed_demo)

    invoice_id = data.get("invoice_number") or pdf_path.stem or "output"

    # 4. JSON
    output_dir = Path(OUTPUT_DIR)
    output_dir.mkdir(parents=True, exist_ok=True)

    json_path = output_dir / f"invoice_{invoice_id}.json"
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=4)

    print(f"[+] JSON saved: {json_path}")

    # 5. Excel + Word
    save_to_excel(data)
    generate_word(data)

    print("[+] Completed.")


def main():
    if len(sys.argv) < 2:
        print("Usage:")
        print("  python main.py input/sample.pdf")
        print("  python main.py input/   # папка с PDF")
        return

    target = Path(sys.argv[1])

    if not target.exists():
        print(f"[ERR] Path not found: {target}")
        return

    if target.is_file() and target.suffix.lower() == ".pdf":
        process_single_pdf(target)
        return

    if target.is_dir():
        pdf_files = list(target.glob("*.pdf"))
        if not pdf_files:
            print("[ERR] No PDF files found in folder.")
            return

        for pdf in pdf_files:
            process_single_pdf(pdf)

        print("\n[+] All PDFs in folder processed.")
        return

    print("[ERR] Invalid input. Provide a PDF file or folder path.")


if __name__ == "__main__":
    main()
