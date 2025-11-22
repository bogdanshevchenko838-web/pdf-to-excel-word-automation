import re
from typing import Dict, Any


def _extract_block(label_start: str, label_end: str, text: str) -> str:
    pattern = rf"{label_start}\s*(.*?){label_end}"
    m = re.search(pattern, text, re.S)
    if not m:
        return ""
    lines = [line.strip() for line in m.group(1).splitlines() if line.strip()]
    return "\n".join(lines)


def _fmt_num(value) -> str:
    if value is None or value == "":
        return ""
    try:
        return f"{float(value):.2f}"
    except Exception:
        return str(value)


def normalize_invoice(basic: Dict[str, Any], parsed: Dict[str, Any]) -> Dict[str, str]:
    text = basic.get("raw_text", "") or ""

    from_block = _extract_block("From:", "To:", text)
    to_block = _extract_block("To:", "Hrs/Qty", text)

    cut_labels = ["Invoice Number", "Order Number", "Invoice Date", "Due Date", "Total Due"]
    for label in cut_labels:
        idx = from_block.find(label)
        if idx != -1:
            from_block = from_block[:idx].rstrip()
            break

    items = parsed.get("items") or []
    first_item = items[0] if items else {}

    totals = parsed.get("totals") or {}

    return {
        "invoice_number": (parsed.get("invoice_number") or basic.get("invoice_number") or ""),
        "order_number": (parsed.get("order_number") or basic.get("order_number") or ""),
        "invoice_date": (parsed.get("invoice_date") or basic.get("invoice_date") or ""),
        "due_date": (parsed.get("due_date") or basic.get("due_date") or ""),

        "from": from_block,
        "to": to_block,

        "hrs_qty": str(first_item.get("qty", "")) if first_item else "",
        "service": first_item.get("description", "") if first_item else "",
        "rate_price": _fmt_num(first_item.get("unit_price")),
        "adjust": "0.00%",
        "sub_total": _fmt_num(totals.get("subtotal")),
        "tax": _fmt_num(totals.get("tax")),
        "total": _fmt_num(totals.get("total")),
    }
