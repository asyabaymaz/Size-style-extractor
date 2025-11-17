"""
Style–Size Extractor (Desktop Drag-and-Drop App – Tkinter)

This version is cleaned up for packaging as a **Windows EXE** (no Python needed on the target machine).

Features
- Click-to-browse PDF (or select .txt) and extract
- Parses labels case-insensitively: Size/Style:, Shirt Size/Style:, Shirt Size:, Product Size -Style:, Product Size:, Style:
- Quantity detection near the label (searches a few lines above/below)
- Aggregates identical Style/Size text as written
- Hoodie summary: counts only lines containing "Hooded Sweatshirt" or "Unisex Hoodie"
- Export CSV/XLSX (with Sweatshirts and Hoodies sheets)
"""

from __future__ import annotations
import os
import re
import sys
from collections import defaultdict
from dataclasses import dataclass
from typing import Dict, Iterable, List, Sequence, Tuple

# GUI
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# Third-party deps
try:
    from pdfminer_high_level import extract_text as pdf_extract_text  # type: ignore
except Exception:
    try:
        from pdfminer.high_level import extract_text as pdf_extract_text  # type: ignore
    except Exception:
        pdf_extract_text = None

try:
    import pandas as pd
except Exception as e:
    print("This app requires pandas. Try: pip install pandas", file=sys.stderr)
    raise

def detect_max_blank(lines):
    blank_counts = []
    for i, line in enumerate(lines):
        if "style" in line.lower() or "size" in line.lower():
            count = 0
            j = i + 1
            while j < len(lines) and lines[j].strip() == "":
                count += 1
                j += 1
            blank_counts.append(count)

    if not blank_counts:
        return 1

    avg = sum(blank_counts) / len(blank_counts)
    return max(1, int(round(avg)) + 1)


# ------------------------- Quantity helpers -------------------------
def _parse_quantity_from_line(line: str) -> int | None:
    """
    Try to parse quantity from a single line.
    Supports formats like:
      - Quantity: 3
      - Qty: 2
      - Qty 5
      - Quantity x 4
    """
    m = re.search(r"(?:qty|quantity)\s*[:x]*\s*(\d+)", line, re.IGNORECASE)
    if m:
        return int(m.group(1))
    return None


def find_quantity_near(lines: List[str], idx: int, radius: int = 5) -> int:
    """
    Look for a Quantity / Qty line a few lines ABOVE or BELOW the label line.
    If none found, default to 1.
    """
    # Look upwards
    for offset in range(1, radius + 1):
        j = idx - offset
        if j < 0:
            break
        q = _parse_quantity_from_line(lines[j])
        if q is not None:
            return q

    # Look downwards
    for offset in range(1, radius + 1):
        j = idx + offset
        if j >= len(lines):
            break
        q = _parse_quantity_from_line(lines[j])
        if q is not None:
            return q

    # Default
    return 1


# ------------------------- Config -------------------------
LABEL_PATTERNS: Sequence[str] = (
    r"Size/Style:",
    r"Shirt Size/Style:",
    r"Shirt Size:",
    r"Product Size\s*-\s*Style:",
    r"Product Size:",
    r"Style:",
    r"Size:",
)

LABEL_RE = re.compile(
    r'^\s*(?:' + '|'.join(p.replace(":", "") for p in LABEL_PATTERNS) + r')\s*:\s*(.*)\s*$',
    re.IGNORECASE,
)

# eski QTY_RE / nearest_non_empty aslında kullanılmıyor ama kalsın dursun istersek ileride
QTY_RE = re.compile(r"^\s*(?:qty|quantity)\s*[:x]*\s*(\d+)\s*$", re.IGNORECASE)

HOODIE_MARKERS = ("hooded sweatshirt", "unisex hoodie")


# ------------------------- Core logic -------------------------
def normalize_key(s: str) -> str:
    s2 = re.sub(r"\s+", " ", s.strip())
    s2 = s2.replace("–", "-").replace("—", "-")
    return s2


def is_hoodie(txt: str) -> bool:
    return any(marker in txt.lower() for marker in HOODIE_MARKERS)


def is_sweatshirt_nonhoodie(txt: str) -> bool:
    t = txt.lower()
    return ("sweatshirt" in t) and (not is_hoodie(txt))


def parse_lines(lines: Sequence[str], max_blank: int = 2) -> List[Tuple[str, int]]:
    """
    Parse raw text lines into (label_value, quantity) pairs.

    - Fixes non-breaking spaces
    - Detects label lines via LABEL_RE
    - For each label line, searches a few lines above/below for Qty/Quantity
    """
    entries: List[Tuple[str, int]] = []
    lines_list = [ln.replace("\xa0", " ") for ln in lines]

    for idx, ln in enumerate(lines_list):
        m = LABEL_RE.match(ln)
        if not m:
            continue

        label_value = m.group(1).strip()
        qty = find_quantity_near(lines_list, idx, radius=5)
        entries.append((label_value, qty))

    return entries


@dataclass
class ExtractionResult:
    agg: Dict[str, int]
    unique_count: int
    sweatshirt_nonhoodie_total: int
    hoodie_total: int

    @property
    def sweatshirt_total(self) -> int:
        return self.sweatshirt_nonhoodie_total


def summarize(entries: Iterable[Tuple[str, int]]) -> ExtractionResult:
    agg: Dict[str, int] = defaultdict(int)
    for raw, q in entries:
        agg[normalize_key(raw)] += int(q)
    hoodie_total = sum(q for k, q in agg.items() if is_hoodie(k))
    sweat_total = sum(q for k, q in agg.items() if is_sweatshirt_nonhoodie(k))
    return ExtractionResult(
        agg=dict(agg),
        unique_count=len(agg),
        sweatshirt_nonhoodie_total=sweat_total,
        hoodie_total=hoodie_total,
    )


# ------------------------- I/O helpers -------------------------
def read_pdf_text(path: str) -> str:
    if pdf_extract_text is None:
        raise RuntimeError("pdfminer.six is not installed. Install it or provide a .txt file.")
    return pdf_extract_text(path) or ""


def extract_from_path(path: str, max_blank: int = 2) -> ExtractionResult:
    ext = os.path.splitext(path)[1].lower()
    if ext == ".pdf":
        text = read_pdf_text(path)
    else:
        with open(path, "r", encoding="utf-8", errors="ignore") as f:
            text = f.read()
    lines = [ln.rstrip("\n") for ln in text.splitlines()]
    entries = parse_lines(lines, max_blank=max_blank)
    return summarize(entries)


# ------------------------- Category logic -------------------------
def category_rank(name: str) -> int:
    n = name.lower()

    short_sleeve_pat = r"(shortsleeve|short\s*-?\s*sleeve|t\s*-?\s*shirt|tshirt|\btee\b)"
    toddler_pat = r"\b[2-5]t\b"
    size_pat = r"\b(xs|s|m|l|xl|[2-6]xl)\b"

    # 0 = Sweatshirts (non-hoodie)
    if "sweatshirt" in n and "hooded" not in n and "hoodie" not in n:
        return 0

    # 1 = Long sleeve
    if "long sleeve" in n or "long-sleeve" in n or "longsleeve" in n:
        return 1

    # 2 = Hoodies
    if ("hooded sweatshirt" in n) or ("unisex hoodie" in n) or re.search(r"\bhoodie\b", n):
        return 2

    # 10 = V-necks (Unisex V-neck dahil)
    if ("v-neck" in n) or ("v neck" in n) or ("vneck" in n):
        return 10

    # 3 = Adult/Unisex short-sleeve / tee (default if not youth/toddler)
    if (
        re.search(short_sleeve_pat, n)
        and not ("youth" in n or re.search(toddler_pat, n) or "toddler" in n)
    ):
        return 3

    # 3b = Standalone Unisex with size treated as tee (e.g., "Unisex XL")
    if (
        "unisex" in n
        and not any(k in n for k in ["hoodie", "hooded", "sweatshirt"])
        and not any(v in n for v in ["v-neck", "v neck", "vneck"])
        and not ("youth" in n or "toddler" in n or re.search(toddler_pat, n))
        and re.search(size_pat, n)
    ):
        return 3

    # 4 = Youth short sleeve / tee
    if re.search(short_sleeve_pat, n) and "youth" in n:
        return 4

    # 5 = Other youth
    if "youth" in n:
        return 5

    # 6 = Toddler short-sleeve (2T–5T)
    if re.search(toddler_pat, n) and re.search(r"(short|tee|shirt)", n):
        return 6

    # 7 = Other toddler
    if re.search(toddler_pat, n) or "toddler" in n:
        return 7

    # 8 = Onesie / baby bodysuit
    if ("onesie" in n) or ("baby bodysuit" in n) or ("bodysuit" in n):
        return 8

    # 11 = everything else
    return 11


def to_dataframe(agg: Dict[str, int]) -> "pd.DataFrame":
    """Build a DataFrame sorted by custom category order, then by size within each category."""

    size_order = {s: i for i, s in enumerate(["xs", "s", "m", "l", "xl", "2xl", "3xl", "4xl", "5xl", "6xl"])}

    def size_rank(name: str) -> int:
        n = name.lower()
        m = re.search(r"\b(xs|s|m|l|xl|[2-6]xl)\b", n)
        if not m:
            return len(size_order)
        return size_order.get(m.group(1), len(size_order))

    def sort_key(item):
        name = item[0]
        return (category_rank(name), size_rank(name), name.lower())

    data = sorted(agg.items(), key=sort_key)
    return pd.DataFrame(data, columns=["Style / Size", "Total Quantity Ordered"])


def save_outputs(df: "pd.DataFrame", csv_path: str | None, xlsx_path: str | None) -> None:
    if csv_path:
        df.to_csv(csv_path, index=False)
    if xlsx_path:
        with pd.ExcelWriter(xlsx_path, engine="openpyxl") as xlw:
            df.to_excel(xlw, sheet_name="All", index=False)
            df_sweat = df[
                df["Style / Size"].str.contains(r"sweatshirt", case=False, regex=True)
                & ~df["Style / Size"].str.contains(r"hooded sweatshirt|unisex hoodie", case=False, regex=True)
            ]
            df_hood = df[
                df["Style / Size"].str.contains(r"hooded sweatshirt|unisex hoodie", case=False, regex=True)
            ]
            df_sweat.to_excel(xlw, sheet_name="Sweatshirts", index=False)
            df_hood.to_excel(xlw, sheet_name="Hoodies", index=False)


# ------------------------- GUI -------------------------
class App:
    def __init__(self) -> None:
        self.root = tk.Tk()
        self.root.title("Style–Size Extractor")
        self.pdf_path: str | None = None
        self.df: "pd.DataFrame | None" = None

        frm = ttk.Frame(self.root, padding=10)
        frm.pack(fill="both", expand=True)

        self.lbl = ttk.Label(frm, text="Click Browse to select a PDF (or .txt)")
        self.lbl.pack(pady=10)

        self.btn_browse = ttk.Button(frm, text="Browse PDF", command=self.browse)
        self.btn_browse.pack(pady=5)

        self.btn_extract = ttk.Button(frm, text="Extract", command=self.extract)
        self.btn_extract.pack(pady=5)

        self.text = tk.Text(frm, height=20)
        self.text.pack(fill="both", expand=True, pady=10)

        self.btn_save_csv = ttk.Button(frm, text="Save CSV", command=self.save_csv, state="disabled")
        self.btn_save_xlsx = ttk.Button(frm, text="Save XLSX", command=self.save_xlsx, state="disabled")
        self.btn_save_csv.pack(pady=3)
        self.btn_save_xlsx.pack(pady=3)

    def browse(self) -> None:
        path = filedialog.askopenfilename(
            filetypes=[("PDF files", "*.pdf"), ("Text files", "*.txt"), ("All files", "*")]
        )
        if path:
            self.pdf_path = path
            self.lbl.config(text=os.path.basename(path))

    def extract(self) -> None:
        if not self.pdf_path:
            messagebox.showerror("Error", "No file selected!")
            return
        try:
            result = extract_from_path(self.pdf_path, max_blank=2)
            self.df = to_dataframe(result.agg)
            self.text.delete("1.0", tk.END)

            # --- SUMMARY COUNTING ---
            total_items = int(sum(result.agg.values()))

            sweatshirt_total = result.sweatshirt_nonhoodie_total
            hoodie_total = result.hoodie_total
            vneck_total = 0
            longsleeve_total = 0
            adult_tee_total = 0
            youth_total = 0
            toddler_total = 0
            onesie_total = 0
            apron_total = 0
            tote_total = 0

            for style, q in result.agg.items():
                cat = category_rank(style)
                s_lower = style.lower()

                if cat == 1:
                    longsleeve_total += q
                elif cat == 3:
                    adult_tee_total += q
                elif cat in (4, 5):
                    youth_total += q
                elif cat in (6, 7):
                    toddler_total += q
                elif cat == 8:
                    onesie_total += q
                elif cat == 10:
                    vneck_total += q

                if "apron" in s_lower:
                    apron_total += q
                if "tote" in s_lower:
                    tote_total += q

            # --- SUMMARY OUTPUT ---
            self.text.insert(tk.END, "--- SUMMARY ---\n")
            self.text.insert(tk.END, f"Total unique combos: {result.unique_count}\n")
            self.text.insert(tk.END, f"Total items: {total_items}\n")
            self.text.insert(tk.END, "By Category:\n")
            self.text.insert(tk.END, f"  • Sweatshirts (non-hoodie): {sweatshirt_total}\n")
            self.text.insert(tk.END, f"  • Hoodies: {hoodie_total}\n")
            self.text.insert(tk.END, f"  • Adult Short Sleeve / Tees: {adult_tee_total}\n")
            self.text.insert(tk.END, f"  • V-Neck: {vneck_total}\n")
            self.text.insert(tk.END, f"  • Long Sleeve: {longsleeve_total}\n")
            self.text.insert(tk.END, f"  • Youth (all): {youth_total}\n")
            self.text.insert(tk.END, f"  • Toddler (all): {toddler_total}\n")
            self.text.insert(tk.END, f"  • Onesie / Baby Bodysuit: {onesie_total}\n")
            self.text.insert(tk.END, f"  • Apron: {apron_total}\n")
            self.text.insert(tk.END, f"  • Tote Bag: {tote_total}\n")

            # --- DETAILED LIST ---
            self.text.insert(tk.END, "\n--- DETAILS (Style / Size) ---\n")

            size_order = {s: i for i, s in enumerate(
                ["xs", "s", "m", "l", "xl", "2xl", "3xl", "4xl", "5xl", "6xl"]
            )}

            def size_rank_local(name: str) -> int:
                n = name.lower()
                m = re.search(r"\b(xs|s|m|l|xl|[2-6]xl)\b", n)
                if not m:
                    return len(size_order)
                return size_order.get(m.group(1), len(size_order))

            sorted_items = sorted(
                result.agg.items(),
                key=lambda item: (category_rank(item[0]), size_rank_local(item[0]), item[0].lower()),
            )

            for style, qty in sorted_items:
                self.text.insert(tk.END, f"{style}: {qty}\n")

            self.btn_save_csv.config(state="normal")
            self.btn_save_xlsx.config(state="normal")

        except Exception as e:
            messagebox.showerror("Error", str(e))

    def save_csv(self) -> None:
        if self.df is None:
            messagebox.showerror("Error", "No data to save.")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if path:
            save_outputs(self.df, csv_path=path, xlsx_path=None)
            messagebox.showinfo("Saved", f"CSV saved to:\n{path}")

    def save_xlsx(self) -> None:
        if self.df is None:
            messagebox.showerror("Error", "No data to save.")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if path:
            save_outputs(self.df, csv_path=None, xlsx_path=path)
            messagebox.showinfo("Saved", f"XLSX saved to:\n{path}")

    def run(self) -> None:
        self.root.mainloop()


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--run-tests":
        print("No tests wired yet.")
    else:
        app = App()
        app.run()
