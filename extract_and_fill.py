# extract_and_fill.py
import re
from io import BytesIO
from decimal import Decimal, InvalidOperation
from typing import Dict, List, Tuple, Optional

import pdfplumber
import pandas as pd
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

EXPECTED_COLUMNS = ["Pos.", "Référence", "Désignation", "Qté", "Prix unit.", "Total CHF"]

# ------------------------------
# PDF extraction + cleaning
# ------------------------------

def _insert_missing_spaces(text: str) -> str:
    # Insert a missing space before codes like 'PC123456' if glued to a number
    # e.g., "23.40PC235490" -> "23.40 PC235490"
    return re.sub(r"(\d)(PC\d{3,})", r"\1 \2", text)

def extract_text_and_tables_from_pdf(file_like) -> Tuple[str, List[pd.DataFrame]]:
    """
    Extract full text + candidate tables using pdfplumber.
    Returns (full_text, list_of_dataframes)
    """
    texts = []
    tables: List[pd.DataFrame] = []
    with pdfplumber.open(file_like) as pdf:
        for page in pdf.pages:
            raw_text = page.extract_text() or ""
            texts.append(_insert_missing_spaces(raw_text))

            # Try multiple table extraction strategies
            try:
                for raw in page.extract_tables() or []:
                    if not raw or len(raw) < 1:
                        continue
                    header = raw[0]
                    rows = raw[1:] if len(raw) > 1 else []
                    if not header or all(h is None or str(h).strip()=="" for h in header):
                        # create generic headers
                        ncols = max(len(r) for r in raw if r) if raw else 0
                        header = [f"Col{i+1}" for i in range(ncols)]
                    df = pd.DataFrame(rows, columns=[(h or "").strip() for h in header])
                    df = _clean_df(df)
                    if df.shape[1] >= 3 and df.shape[0] >= 1:
                        tables.append(df)
            except Exception:
                pass

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
                    if df.shape[1] >= 3 and df.shape[0] >= 1:
                        tables.append(df)
            except Exception:
                pass

    # Deduplicate by signature
    unique: List[pd.DataFrame] = []
    sigs = set()
    for df in tables:
        sig = (tuple(df.columns), df.shape)
        if sig not in sigs:
            sigs.add(sig)
            unique.append(df)

    full_text = "\n".join(texts)
    return full_text, unique

def _clean_df(df: pd.DataFrame) -> pd.DataFrame:
    # drop all-empty columns, strip headers and cells
    df.columns = [str(c).strip() for c in df.columns]
    df = df.loc[:, ~(df.columns.str.strip()=="")]
    df = df.applymap(lambda x: (str(x).strip() if x is not None else ""))
    # remove empty rows
    df = df.loc[~(df.apply(lambda r: all((not str(v).strip()) for v in r), axis=1))]
    return df.reset_index(drop=True)

# ------------------------------
# Field parsing (header/footer info)
# ------------------------------

def parse_fields_from_text(text: str) -> Dict[str, str]:
    """
    Extracts common fields from the free text using regexes.
    """
    fields: Dict[str, str] = {}

    m = re.search(r"Commande fournisseur N[°º]\s*([A-Z0-9\-_]+)", text, flags=re.IGNORECASE)
    if m:
        fields["N°commande fournisseur"] = m.group(1).strip()
        fields["Commande fournisseur"] = fields["N°commande fournisseur"]

    m = re.search(r"Date\s*([0-3]?\d\.[01]?\d\.[12]\d{3})", text)
    if m:
        fields["date du jour"] = m.group(1).strip()

    m = re.search(r"Délai de réception\s*:\s*([0-3]?\d\.[01]?\d\.[12]\d{3})", text)
    if m:
        fields["date Délai de livraison"] = m.group(1).strip()

    m = re.search(r"Condition de paiement\s*([A-Za-z0-9\s]+)", text)
    if m:
        fields["Cond. de paiement"] = m.group(1).strip()

    m = re.search(r"(Montant Total TTC CHF|Total TTC CHF)\s*([0-9'’.,]+)", text, flags=re.IGNORECASE)
    if m:
        fields["Total TTC CHF"] = m.group(2).strip()

    return fields

# ------------------------------
# Column auto-mapping
# ------------------------------

def suggest_column_mapping(df: pd.DataFrame) -> Dict[str, Optional[str]]:
    """
    Suggest mapping from df columns to EXPECTED_COLUMNS keys.
    Returns dict: {expected_col_name: df_column_name or None}
    """
    cols = list(df.columns)
    mapping: Dict[str, Optional[str]] = {k: None for k in EXPECTED_COLUMNS}

    # Simple heuristics
    for c in cols:
        lc = c.lower()
        if mapping["Référence"] is None and re.search(r"réf|ref|pc\d+|code", lc):
            mapping["Référence"] = c
        if mapping["Désignation"] is None and re.search(r"désignation|designation|descr|désign", lc):
            mapping["Désignation"] = c
        if mapping["Qté"] is None and re.search(r"qt|quant|qte|qty", lc):
            mapping["Qté"] = c
        if mapping["Prix unit."] is None and re.search(r"prix.*unit|unit.*price|p\.?u", lc):
            mapping["Prix unit."] = c
        if mapping["Total CHF"] is None and re.search(r"total|montant", lc):
            mapping["Total CHF"] = c
        if mapping["Pos."] is None and re.search(r"pos|position", lc):
            mapping["Pos."] = c

    # Fallbacks by content
    for c in cols:
        sample = " ".join(df[c].astype(str).head(10).tolist()).lower()
        if mapping["Référence"] is None and re.search(r"\bpc\d{3,}\b", sample):
            mapping["Référence"] = c

    return mapping

def apply_mapping(df: pd.DataFrame, mapping: Dict[str, Optional[str]]) -> pd.DataFrame:
    """
    Return a new DataFrame with EXPECTED_COLUMNS using the mapping.
    Missing columns are filled with blanks.
    """
    out = pd.DataFrame()
    for k in EXPECTED_COLUMNS:
        src = mapping.get(k)
        out[k] = df[src] if src in df.columns else ""
    # compute totals if missing and possible
    if out["Total CHF"].eq("").all() and not out["Prix unit."].eq("").all() and not out["Qté"].eq("").all():
        try:
            out["Total CHF"] = [
                _fmt_money(_to_decimal(q) * _to_decimal(pu)) if q and pu else ""
                for q, pu in zip(out["Qté"], out["Prix unit."])
            ]
        except Exception:
            pass
    # normalize numbers alignment later in DOCX
    return out

def _to_decimal(s: str) -> Decimal:
    s = str(s).strip().replace("'", "").replace("’", "").replace(" ", "").replace(",", ".")
    try:
        return Decimal(s)
    except (InvalidOperation, ValueError):
        return Decimal(0)

def _fmt_money(d: Decimal) -> str:
    # Format like 1'234.56 (CH)
    q = d.quantize(Decimal("0.01"))
    s = f"{q:.2f}"
    whole, dot, frac = s.partition(".")
    rev = whole[::-1]
    grouped = "'".join(rev[i:i+3] for i in range(0, len(rev), 3))[::-1]
    return grouped + (dot + frac if frac else "")

# ------------------------------
# DOCX utilities
# ------------------------------

def _replace_in_paragraph(paragraph, target: str, replacement: str):
    if not target:
        return
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

def replace_placeholders_everywhere(doc: Document, mapping: Dict[str, str]):
    # paragraphs
    for p in doc.paragraphs:
        for key, val in mapping.items():
            for variant in (f"« {key} »", f"«\xa0{key}\xa0»"):
                _replace_in_paragraph(p, variant, str(val))

    # tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for key, val in mapping.items():
                        for variant in (f"« {key} »", f"«\xa0{key}\xa0»"):
                            _replace_in_paragraph(p, variant, str(val))

    # headers/footers
    for section in doc.sections:
        for hdr in (section.header, section.footer):
            for p in hdr.paragraphs:
                for key, val in mapping.items():
                    for variant in (f"« {key} »", f"«\xa0{key}\xa0»"):
                        _replace_in_paragraph(p, variant, str(val))

def find_or_create_items_table(doc: Document, expected_cols: List[str]) -> 'docx.table.Table':
    # find a table whose first row matches (subset) of expected headers
    for table in doc.tables:
        if len(table.rows) >= 1:
            headers = [c.text.strip() for c in table.rows[0].cells]
            if all(any(h.lower()==eh.lower() for h in headers) for eh in expected_cols if eh in headers):
                _clear_table_rows_but_header(table)
                return table
    # else create new table at end
    table = doc.add_table(rows=1, cols=len(expected_cols))
    for i, h in enumerate(expected_cols):
        table.rows[0].cells[i].text = h
    return table

def _clear_table_rows_but_header(table):
    while len(table.rows) > 1:
        table._element.remove(table.rows[1]._element)

def insert_items_into_table(table, items_df: pd.DataFrame):
    for _, row in items_df.iterrows():
        cells = table.add_row().cells
        for i, col in enumerate(items_df.columns):
            val = "" if pd.isna(row[col]) else str(row[col])
            cells[i].text = val
            for para in cells[i].paragraphs:
                if para.runs:
                    para.runs[0].font.size = Pt(10)
                if _looks_numeric(val):
                    para.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
                else:
                    para.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

def _looks_numeric(s: str) -> bool:
    try:
        _ = _to_decimal(s)
        return True
    except Exception:
        return False

def add_total_bottom_right(doc: Document, total: Optional[str]):
    if not total:
        return
    p = doc.add_paragraph()
    p.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
    run = p.add_run(f"Total TTC CHF {total}")
    run.bold = True
    run.font.size = Pt(11)

# ------------------------------
# High-level API
# ------------------------------

def process_pdf_to_docx(
    pdf_bytes: bytes,
    template_docx_bytes: bytes,
    placeholder_overrides: Optional[Dict[str, str]] = None,
    custom_mapping: Optional[Dict[str, Optional[str]]] = None
) -> Tuple[bytes, Dict[str, str], pd.DataFrame]:
    """
    Returns (generated_docx_bytes, detected_fields, mapped_items_df)
    """
    text, tables = extract_text_and_tables_from_pdf(BytesIO(pdf_bytes))
    fields = parse_fields_from_text(text)

    if placeholder_overrides:
        fields.update({k: v for k, v in placeholder_overrides.items() if v})

    # Choose largest table and build mapping
    if tables:
        base_df = max(tables, key=lambda d: (d.shape[0], d.shape[1]))
    else:
        base_df = pd.DataFrame()

    if not base_df.empty:
        auto_map = suggest_column_mapping(base_df)
        if custom_mapping:
            auto_map.update({k: v for k, v in custom_mapping.items()})
        items_df = apply_mapping(base_df, auto_map)
    else:
        items_df = pd.DataFrame(columns=EXPECTED_COLUMNS)

    # Build document
    doc = Document(BytesIO(template_docx_bytes))
    replace_placeholders_everywhere(doc, fields)

    table = find_or_create_items_table(doc, EXPECTED_COLUMNS)
    insert_items_into_table(table, items_df)

    add_total_bottom_right(doc, fields.get("Total TTC CHF"))

    out = BytesIO()
    doc.save(out)
    out.seek(0)
    return out.read(), fields, items_df
