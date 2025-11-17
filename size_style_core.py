"""
Style–Size Extractor (no GUI).

This module contains only the parsing, summarizing and category logic,
so it can be reused from both Tkinter (desktop) and Streamlit (web) apps.
"""

from __future__ import annotations
import os
import re
import sys
from collections import defaultdict
from dataclasses import dataclass
from typing import Dict, Iterable, List, Sequence, Tuple

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
   For each label line, searches a few lines above/below for Qty/Quantity
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
