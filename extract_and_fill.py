# extract_and_fill.py — fix6
import re
import unicodedata
from io import BytesIO
from typing import Dict, List, Tuple, Optional
from datetime import datetime
from zoneinfo import ZoneInfo

import pdfplumber
import pandas as pd
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

DATE_RE = re.compile(r"\b([0-3]?\d)[./-]([01]?\d)[./-]([12]\d{3})\b")

def today_ch() -> str:
    """Return today's date in Europe/Zurich as dd.mm.yyyy."""
    return datetime.now(ZoneInfo("Europe/Zurich")).strftime("%d.%m.%Y")

def _strip_accents(s: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c))

def _insert_missing_spaces(text: str) -> str:
    text = re.sub(r"(\d)[A-Za-z]", lambda m: m.group(0)[0] + " " + m.group(0)[1:], text)
    text = re.sub(r"(\d)(PC\d{3,})", r"\1 \2", text)
    return text

def extract_text_and_tables_from_pdf(file_like) -> Tuple[str, List[pd.DataFrame]]:
    texts = []
    tables: List[pd.DataFrame] = []
    with pdfplumber.open(file_like) as pdf:
        for page in pdf.pages:
            raw_text = page.extract_text() or ""
            raw_text = _insert_missing_spaces(raw_text)
            texts.append(raw_text)
            # try multi-tables
            try:
                for raw in page.extract_tables() or []:
                    if not raw or len(raw) < 1:
                        continue
                    header = raw[0]
                    rows = raw[1:] if len(raw) > 1 else []
                    if not header or all(h is None or str(h).strip()=="" for h in header):
                        ncols = max(len(r) for r in raw if r) if raw else 0
                        header = [f"Col{i+1}" for i in range(ncols)]
                    df = pd.DataFrame(rows, columns=[(h or "").strip() for h in header])
                    df = _clean_df(df)
                    if df.shape[1] >= 2 and df.shape[0] >= 1:
                        tables.append(df)
            except Exception:
                pass
            # try single table
            try:
                raw_single = page.extract_table()
                if raw_single and len(raw_single) > 1:
                    header = raw_single[0]
                    rows = raw_single[1:]
                    if not header or all(h is None or str(h).strip()=="" for h in header):
                        ncols = max(len(r) for r in raw_single if r) if raw_single else 0
                        header = [f"Col{i+1}" for i in range(ncols)]
                    df = pd.DataFrame(rows, columns=[(h or "").strip() for h in header])
                    df = _clean_df(df)
                    if df.shape[1] >= 2 and df.shape[0] >= 1:
                        tables.append(df)
            except Exception:
                pass
    # de-duplicate by (cols,shape)
    unique = []
    sigs = set()
    for df in tables:
        sig = (tuple(df.columns), df.shape)
        if sig not in sigs:
            sigs.add(sig)
            unique.append(df)
    full_text = "\n".join(texts)
    return full_text, unique

def _clean_df(df: pd.DataFrame) -> pd.DataFrame:
    df.columns = [str(c).strip() for c in df.columns]
    df = df.loc[:, ~(df.columns.str.strip()=="")]
    df = df.applymap(lambda x: (str(x).strip() if x is not None else ""))
    df = df.loc[~(df.apply(lambda r: all((not str(v).strip()) for v in r), axis=1))]
    return df.reset_index(drop=True)

def parse_fields_from_text(text: str) -> Dict[str, str]:
    fields: Dict[str, str] = {}
    t1 = _strip_accents(text).lower().replace("\xa0", " ")
    # N° commande
    m = re.search(r"commande fournisseur n[°o]\s*([A-Z0-9\-_]+)", t1, flags=re.IGNORECASE)
    if m:
        fields["N°commande fournisseur"] = m.group(1).strip()
        fields["Commande fournisseur"] = fields["N°commande fournisseur"]
    # Total TTC CHF (either order)
    m = re.search(r"(montant total ttc chf|total ttc chf)\s*([0-9'’.,]+)", t1, flags=re.IGNORECASE)
    if m:
        fields["Total TTC CHF"] = m.group(2).strip()
    else:
        m = re.search(r"([0-9'’.,]+)\s*(montant total ttc chf|total ttc chf)", t1, flags=re.IGNORECASE)
        if m:
            fields["Total TTC CHF"] = m.group(1).strip()
    # Date du jour: always override later with today_ch()
    m = re.search(r"\bdate\s+([0-3]?\d[./-][01]?\d[./-][12]\d{3})", t1)
    if m:
        fields["date du jour"] = m.group(1).strip()
    return fields

def _parse_date_str(s: str):
    m = DATE_RE.search(s)
    if not m:
        return None
    d, mth, y = m.groups()
    try:
        return datetime(int(y), int(mth), int(d))
    except ValueError:
        return None

def extract_latest_receipt_deadline_from_tables(tables: List[pd.DataFrame]) -> Optional[str]:
    """Look for rows containing 'Délai de réception' (accent-insensitive), gather dates in the same row, return max."""
    dates = []
    for df in tables:
        for _, row in df.iterrows():
            cells = [str(v) if v is not None else "" for v in row.tolist()]
            norm_cells = [_strip_accents(c.lower()) for c in cells]
            if any("delai de reception" in c for c in norm_cells):
                for c in cells:
                    dt = _parse_date_str(c)
                    if dt: dates.append(dt)
    if not dates:
        return None
    return max(dates).strftime("%d.%m.%Y")

def extract_latest_receipt_deadline_from_text(text: str) -> Optional[str]:
    """Fallback: scan plain text lines with 'Délai de réception' and return the max date."""
    dates = []
    norm = _strip_accents(text).lower().replace("\xa0", " ")
    for ln in [l.strip() for l in norm.splitlines() if l.strip()]:
        if "delai de reception" in ln:
            dt = _parse_date_str(ln)
            if dt: dates.append(dt)
    if not dates:
        # pattern within 30 chars
        pattern = re.compile(r"delai de reception.{0,30}?([0-3]?\d[./-][01]?\d[./-][12]\d{3})")
        for m in pattern.finditer(norm):
            dt = _parse_date_str(m.group(0))
            if dt: dates.append(dt)
    if not dates:
        return None
    return max(dates).strftime("%d.%m.%Y")

def clean_items_df_keep_full(df: pd.DataFrame) -> pd.DataFrame:
    """Keep the PDF table as-is, but:
       - drop any row whose first non-empty cell starts with 'Indice :' or 'Délai de réception :'
       - drop the column named 'TVA' (accent-insensitive) if present
    """
    # drop rows
    rows = []
    for _, r in df.iterrows():
        first_non_empty = ""
        for v in r.tolist():
            s = str(v).strip() if v is not None else ""
            if s:
                first_non_empty = s
                break
        norm_first = _strip_accents(first_non_empty.lower())
        if norm_first.startswith("indice :") or norm_first.startswith("delai de reception :"):
            continue
        rows.append(r)
    new_df = pd.DataFrame(rows, columns=df.columns) if rows else df.iloc[0:0].copy()
    # drop TVA column
    keep_cols = []
    for c in new_df.columns:
        if _strip_accents(str(c)).strip().lower() == "tva":
            continue
        keep_cols.append(c)
    new_df = new_df[keep_cols] if keep_cols else new_df
    return new_df.reset_index(drop=True)

def replace_placeholders_everywhere(doc: Document, mapping: Dict[str, str]):
    def _replace_in_paragraph(paragraph, target: str, replacement: str):
        full_text = "".join(run.text for run in paragraph.runs)
        if target not in full_text:
            return
        new_text = full_text.replace(target, replacement)
        for idx in range(len(paragraph.runs)-1, -1, -1):
            r = paragraph.runs[idx]
            r.clear(); r.text = ""
        if not paragraph.runs:
            paragraph.add_run(new_text)
        else:
            paragraph.runs[0].text = new_text

    for p in doc.paragraphs:
        for key, val in mapping.items():
            for variant in (f"« {key} »", f"«\xa0{key}\xa0»"):
                _replace_in_paragraph(p, variant, str(val))

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for key, val in mapping.items():
                        for variant in (f"« {key} »", f"«\xa0{key}\xa0»"):
                            _replace_in_paragraph(p, variant, str(val))

    for section in doc.sections:
        for hdr in (section.header, section.footer):
            for p in hdr.paragraphs:
                for key, val in mapping.items():
                    for variant in (f"« {key} »", f"«\xa0{key}\xa0»"):
                        _replace_in_paragraph(p, variant, str(val))

def find_first_table(doc: Document):
    try:
        return doc.tables[0]
    except IndexError:
        return None

def clear_table_rows_but_header(table):
    while len(table.rows) > 1:
        table._element.remove(table.rows[1]._element)

def insert_any_df_into_doc(doc: Document, df: pd.DataFrame):
    """Insert df into the first table if column count matches; otherwise create a new table at the end."""
    if df is None or df.empty:
        return
    df = df.copy()
    df.columns = [str(c) for c in df.columns]

    target = find_first_table(doc)
    if target is not None and len(target.columns) == len(df.columns):
        # reuse
        # set headers
        for i, c in enumerate(df.columns):
            target.rows[0].cells[i].text = str(c)
        clear_table_rows_but_header(target)
        for _, row in df.iterrows():
            cells = target.add_row().cells
            for i, col in enumerate(df.columns):
                val = "" if pd.isna(row[col]) else str(row[col])
                cells[i].text = val
                for para in cells[i].paragraphs:
                    if para.runs:
                        para.runs[0].font.size = Pt(10)
                    # align numbers to right
                    if re.match(r"^\s*[0-9'’.,]+\s*$", val):
                        para.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
                    else:
                        para.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
    else:
        # create new table at end
        tbl = doc.add_table(rows=1, cols=len(df.columns))
        for i, c in enumerate(df.columns):
            tbl.rows[0].cells[i].text = str(c)
        for _, row in df.iterrows():
            cells = tbl.add_row().cells
            for i, col in enumerate(df.columns):
                val = "" if pd.isna(row[col]) else str(row[col])
                cells[i].text = val
                for para in cells[i].paragraphs:
                    if para.runs:
                        para.runs[0].font.size = Pt(10)
                    if re.match(r"^\s*[0-9'’.,]+\s*$", val):
                        para.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
                    else:
                        para.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

def process_pdf_to_docx(pdf_bytes: bytes, template_docx_bytes: bytes) -> Tuple[bytes, Dict[str, str], pd.DataFrame]:
    # Extract
    text, tables = extract_text_and_tables_from_pdf(BytesIO(pdf_bytes))

    # Fields
    fields = parse_fields_from_text(text)
    # Force today's date (CH timezone)
    fields["date du jour"] = today_ch()

    # Délai de réception
    from_candidates = []
    # try row-based detection
    from_candidates.append(extract_latest_receipt_deadline_from_tables(tables))
    # fallback on text
    from_candidates.append(extract_latest_receipt_deadline_from_text(text))
    for cand in from_candidates:
        if cand:
            fields["Délai de réception"] = cand
            break

    # Items table: take the largest table as-is and clean per rules
    if tables:
        base_df = max(tables, key=lambda d: (d.shape[0], d.shape[1]))
        items_df = clean_items_df_keep_full(base_df)
    else:
        items_df = pd.DataFrame()

    # Build partial doc (placeholders only; caller inserts table & total)
    doc = Document(BytesIO(template_docx_bytes))
    replace_placeholders_everywhere(doc, fields)
    out = BytesIO(); doc.save(out); out.seek(0)
    return out.getvalue(), fields, items_df
