import re

def _clean(value):
    if not value:
        return ""
    return value.strip()

def parse_invoice_demo(text: str) -> dict:

    # ---------- HEADER TABLE ----------
    m_inv = re.search(r"Invoice Number\s+([A-Z0-9\-]+)", text)
    m_ord = re.search(r"Order Number\s+([0-9]+)", text)
    m_invdate = re.search(r"Invoice Date\s+([A-Za-z0-9, ]+)", text)
    m_duedate = re.search(r"Due Date\s+([A-Za-z0-9, ]+)", text)
    m_total_due = re.search(r"Total Due\s+\$?([0-9]+\.[0-9]+)", text)

    invoice_number = _clean(m_inv.group(1)) if m_inv else ""
    order_number = _clean(m_ord.group(1)) if m_ord else ""
    invoice_date = _clean(m_invdate.group(1)) if m_invdate else ""
    due_date = _clean(m_duedate.group(1)) if m_duedate else ""
    total_due = float(m_total_due.group(1)) if m_total_due else None

    # ---------- SELLER / BUYER ----------
    seller = {"name": "", "email": ""}
    buyer = {"name": "", "email": ""}

    m_from = re.search(r"From:\s*(.*?)To:", text, re.S)
    if m_from:
        block = [b.strip() for b in m_from.group(1).splitlines() if b.strip()]
        if block:
            seller["name"] = block[0]
        seller["email"] = next((line for line in block if "@" in line), "")

    m_to = re.search(r"To:\s*(.*?)Invoice Date", text, re.S)
    if m_to:
        block = [b.strip() for b in m_to.group(1).splitlines() if b.strip()]
        if block:
            buyer["name"] = block[0]
        buyer["email"] = next((line for line in block if "@" in line), "")

    # ---------- ITEMS ----------
    items = []

    m_items_block = re.search(
        r"Hrs/Qty\s*Service\s*Rate/Price\s*Adjust\s*Sub Total\s*(.*?)Sub Total\s*\$[0-9]+\.[0-9]+",
        text,
        re.S,
    )
    if m_items_block:
        block = m_items_block.group(1)

        m_qty = re.search(r"(?<!\$)(\d+(?:\.\d+)?)", block)
        qty = float(m_qty.group(1)) if m_qty else None

        lines = [l.strip() for l in block.splitlines() if l.strip()]
        service = ""
        if m_qty:
            for i, line in enumerate(lines):
                if re.fullmatch(r"\d+(?:\.\d+)?", line):
                    if i + 1 < len(lines):
                        service = lines[i + 1]
                    break

        money = re.findall(r"\$([0-9]+\.[0-9]+)", block)
        rate_price = float(money[0]) if len(money) >= 1 else None
        line_total = float(money[1]) if len(money) >= 2 else None

        m_adj = re.search(r"([0-9]+\.[0-9]+%)", block)
        adjust = m_adj.group(1) if m_adj else ""

        if qty is not None:
            items.append(
                {
                    "qty": qty,
                    "description": service,
                    "unit_price": rate_price,
                    "line_total": line_total,
                    "adjust": adjust,
                }
            )

    # ---------- TOTALS (строки исправлены в raw-string) ----------
    def grab_money(label: str):
        vals = re.findall(fr"{label}\s*\n\s*\$([0-9]+\.[0-9]+)", text, re.I)
        if not vals:
            vals = re.findall(fr"{label}\s*\$?([0-9]+\.[0-9]+)", text, re.I)
        return float(vals[-1]) if vals else None

    subtotal = grab_money(r"Sub[\s\-]?Total")
    tax = grab_money("Tax")
    total = grab_money(r"Total(?!\s*Due)")

    return {
        "invoice_number": invoice_number,
        "order_number": order_number,
        "invoice_date": invoice_date,
        "due_date": due_date,
        "seller": seller,
        "buyer": buyer,
        "items": items,
        "totals": {
            "subtotal": subtotal,
            "tax": tax,
            "total": total,
            "total_due": total_due,
        },
    }
