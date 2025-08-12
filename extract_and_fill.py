# extract_and_fill.py — fix3: "Délai de livraison" = max date found in PDF tables
import re
import unicodedata
from io import BytesIO
from decimal import Decimal, InvalidOperation
from typing import Dict, List, Tuple, Optional
from datetime import datetime

import pdfplumber
import pandas as pd
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

EXPECTED_COLUMNS = ["Pos.", "Référence", "Désignation", "Qté", "Prix unit.", "Total CHF"]

DATE_RE = re.compile(r"\b([0-3]?\d)[./-]([01]?\d)[./-]([12]\d{3})\b")

def _strip_accents(s: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c))

def _norm_spaces(s: str) -> str:
    return re.sub(r"\s+", " ", s).strip()

def _insert_missing_spaces(text: str) -> str:
    text = re.sub(r"(\d)[A-Za-z]", lambda m: m.group(0)[0] + " " + m.group(0)[1:], text)
    text = re.sub(r"(\d)(PC\d{3,})", r"\1 \2", text)
    return text

def _parse_date_str(s: str) -> Optional[datetime]:
    m = DATE_RE.search(s)
    if not m:
        return None
    d, mth, y = m.groups()
    try:
        return datetime(int(y), int(mth), int(d))
    except ValueError:
        return None

def extract_text_and_tables_from_pdf(file_like) -> Tuple[str, List[pd.DataFrame]]:
    texts = []
    tables: List[pd.DataFrame] = []
    with pdfplumber.open(file_like) as pdf:
        for page in pdf.pages:
            raw_text = page.extract_text() or ""
            raw_text = _insert_missing_spaces(raw_text)
            texts.append(raw_text)
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
    # de-duplicate
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
    t1 = _strip_accents(_norm_spaces(text))

    m = re.search(r"commande fournisseur n[°o]\s*([A-Z0-9\-_]+)", t1, flags=re.IGNORECASE)
    if m:
        fields["N°commande fournisseur"] = m.group(1).strip()
        fields["Commande fournisseur"] = fields["N°commande fournisseur"]

    m = re.search(r"\bdate\s+([0-3]?\d[./-][01]?\d[./-][12]\d{3})", t1)
    if m:
        fields["date du jour"] = m.group(1).strip()

    # Condition de paiement
    m = re.search(r"condition de paiement[: ]+([0-9a-z ]{4,})", t1)
    if m:
        fields["Cond. de paiement"] = m.group(1).strip()

    # Total TTC (both orders)
    m = re.search(r"(montant total ttc chf|total ttc chf)\s*([0-9'’.,]+)", t1, flags=re.IGNORECASE)
    if m:
        fields["Total TTC CHF"] = m.group(2).strip()
    else:
        m = re.search(r"([0-9'’.,]+)\s*(montant total ttc chf|total ttc chf)", t1, flags=re.IGNORECASE)
        if m:
            fields["Total TTC CHF"] = m.group(1).strip()

    return fields

def extract_latest_delivery_date_from_tables(tables: List[pd.DataFrame]) -> Optional[str]:
    """Scan all table cells, collect dd.mm.yyyy and return the max date as 'dd.mm.yyyy'."""
    dates: List[datetime] = []
    for df in tables:
        for val in df.astype(str).values.ravel():
            dt = _parse_date_str(val)
            if dt:
                dates.append(dt)
    if not dates:
        return None
    latest = max(dates)
    return latest.strftime("%d.%m.%Y")

# ---- Column mapping helpers ----
def suggest_column_mapping(df: pd.DataFrame) -> Dict[str, Optional[str]]:
    cols = list(df.columns)
    mapping: Dict[str, Optional[str]] = {k: None for k in EXPECTED_COLUMNS}
    for c in cols:
        lc = _strip_accents(c.lower())
        if mapping["Référence"] is None and re.search(r"\b(ref|ref\.|pc\d+|code)\b", lc):
            mapping["Référence"] = c
        if mapping["Désignation"] is None and re.search(r"designation|description|design", lc):
            mapping["Désignation"] = c
        if mapping["Qté"] is None and re.search(r"\b(qte|qt|qty|quant)\b", lc):
            mapping["Qté"] = c
        if mapping["Prix unit."] is None and re.search(r"(prix.*unit|unit.*price|p\.?u)", lc):
            mapping["Prix unit."] = c
        if mapping["Total CHF"] is None and re.search(r"\b(total|montant)\b", lc):
            mapping["Total CHF"] = c
        if mapping["Pos."] is None and re.search(r"\b(pos|position)\b", lc):
            mapping["Pos."] = c
    for c in cols:
        sample = _strip_accents(" ".join(df[c].astype(str).head(10).tolist()).lower())
        if mapping["Référence"] is None and re.search(r"\bpc\d{3,}\b", sample):
            mapping["Référence"] = c
    return mapping

def apply_mapping(df: pd.DataFrame, mapping: Dict[str, Optional[str]]) -> pd.DataFrame:
    out = pd.DataFrame()
    for k in EXPECTED_COLUMNS:
        src = mapping.get(k)
        out[k] = df[src] if src in df.columns else ""
    if out["Total CHF"].eq("").all() and not out["Prix unit."].eq("").all() and not out["Qté"].eq("").all():
        try:
            out["Total CHF"] = [
                _fmt_money(_to_decimal(q) * _to_decimal(pu)) if q and pu else ""
                for q, pu in zip(out["Qté"], out["Prix unit."])
            ]
        except Exception:
            pass
    return out

def _to_decimal(s: str) -> Decimal:
    s = str(s).strip().replace("'", "").replace("’", "").replace(" ", "").replace(",", ".")
    try:
        return Decimal(s)
    except (InvalidOperation, ValueError):
        return Decimal(0)

def _fmt_money(d: Decimal) -> str:
    q = d.quantize(Decimal("0.01"))
    s = f"{q:.2f}"
    whole, dot, frac = s.partition(".")
    rev = whole[::-1]
    grouped = "'".join(rev[i:i+3] for i in range(0, len(rev), 3))[::-1]
    return grouped + (dot + frac if frac else "")

# ---- DOCX helpers ----
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

def find_or_create_items_table(doc: Document, expected_cols: List[str]):
    for table in doc.tables:
        if len(table.rows) >= 1:
            headers = [c.text.strip() for c in table.rows[0].cells]
            if any(h in headers for h in expected_cols):
                _clear_table_rows_but_header(table)
                return table
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

def process_pdf_to_docx(
    pdf_bytes: bytes,
    template_docx_bytes: bytes,
    placeholder_overrides: Optional[Dict[str, str]] = None,
    custom_mapping: Optional[Dict[str, Optional[str]]] = None
):
    # Extract
    text, tables = extract_text_and_tables_from_pdf(BytesIO(pdf_bytes))
    fields = parse_fields_from_text(text)
    # NEW: compute delivery deadline from tables (max date)
    latest_date = extract_latest_delivery_date_from_tables(tables)
    if latest_date:
        fields["Délai de livraison"] = latest_date

    # Apply overrides (user wins)
    if placeholder_overrides:
        fields.update({k: v for k, v in placeholder_overrides.items() if v})

    # Build items table
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

    # Prepare a minimal doc with placeholders replaced (table insert later)
    doc = Document(BytesIO(template_docx_bytes))
    replace_placeholders_everywhere(doc, fields)
    out = BytesIO()
    doc.save(out); out.seek(0)
    return out.getvalue(), fields, items_df
