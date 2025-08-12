# extract_and_fill.py
import re
from io import BytesIO
from decimal import Decimal
from typing import Dict, List, Tuple, Optional

import pdfplumber
import pandas as pd
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

def extract_text_and_tables_from_pdf(file_like) -> Tuple[str, List[pd.DataFrame]]:
    tables: List[pd.DataFrame] = []
    texts = []
    with pdfplumber.open(file_like) as pdf:
        for page in pdf.pages:
            txt = page.extract_text() or ""
            texts.append(txt)
            # Try table extraction
            try:
                for raw in page.extract_tables() or []:
                    if not raw or len(raw) < 1: 
                        continue
                    header = raw[0]
                    rows = raw[1:] if len(raw) > 1 else []
                    if any(h is None for h in header):
                        header = [f"Col{i+1}" for i in range(len(header))]
                    df = pd.DataFrame(rows, columns=[(h or "").strip() for h in header])
                    df = df.loc[:, ~(df.columns.str.strip() == "")]
                    if df.shape[0] > 0 and df.shape[1] > 1:
                        tables.append(df)
            except Exception:
                pass
            try:
                raw_single = page.extract_table()
                if raw_single and len(raw_single) > 1:
                    header = raw_single[0]
                    rows = raw_single[1:]
                    if any(h is None for h in header):
                        header = [f"Col{i+1}" for i in range(len(header))]
                    df = pd.DataFrame(rows, columns=[(h or "").strip() for h in header])
                    df = df.loc[:, ~(df.columns.str.strip() == "")]
                    if df.shape[0] > 0 and df.shape[1] > 1:
                        tables.append(df)
            except Exception:
                pass
    # Deduplicate by signature
    dedup, unique = [], []
    for df in tables:
        sig = (tuple(df.columns), df.shape)
        if sig not in dedup:
            dedup.append(sig)
            unique.append(df)
    full_text = "\n".join(texts)
    return full_text, unique

def parse_fields_from_text(text: str) -> Dict[str, str]:
    fields = {}
    m = re.search(r"Commande fournisseur N[°º]\s*([A-Z0-9\-\_]+)", text, flags=re.IGNORECASE)
    if m:
        fields["N°commande fournisseur"] = m.group(1).strip()
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
    if "N°commande fournisseur" in fields:
        fields["Commande fournisseur"] = fields["N°commande fournisseur"]
    return fields

def _replace_in_paragraph(paragraph, target: str, replacement: str):
    if not target:
        return
    full_text = "".join(run.text for run in paragraph.runs)
    if target not in full_text:
        return
    new_text = full_text.replace(target, replacement)
    # clear runs
    for idx in range(len(paragraph.runs)-1, -1, -1):
        r = paragraph.runs[idx]
        r.clear(); r.text = ""
    # set single run
    if not paragraph.runs:
        paragraph.add_run(new_text)
    else:
        paragraph.runs[0].text = new_text

def replace_placeholders(doc: Document, mapping: Dict[str, str]):
    for p in doc.paragraphs:
        for key, val in mapping.items():
            target = f"« {key} »"
            target_nbsp = f"«\xa0{key}\xa0»"
            _replace_in_paragraph(p, target, str(val))
            _replace_in_paragraph(p, target_nbsp, str(val))
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for key, val in mapping.items():
                    target = f"« {key} »"
                    target_nbsp = f"«\xa0{key}\xa0»"
                    for p in cell.paragraphs:
                        _replace_in_paragraph(p, target, str(val))
                        _replace_in_paragraph(p, target_nbsp, str(val))

def find_first_suitable_table(doc: Document, min_cols: int = 4):
    for table in doc.tables:
        try:
            if len(table.columns) >= min_cols:
                return table
        except Exception:
            pass
    return None

def clear_table_rows_but_header(table):
    while len(table.rows) > 1:
        table._element.remove(table.rows[1]._element)

def _is_numeric_string(s: str) -> bool:
    if s is None:
        return False
    s = str(s).strip()
    if not s:
        return False
    try:
        s2 = s.replace("'", "").replace("’", "").replace(" ", "").replace(",", ".")
        Decimal(s2)
        return True
    except Exception:
        return False

def insert_df_into_table(doc: Document, df: pd.DataFrame):
    table = find_first_suitable_table(doc, min_cols=max(4, df.shape[1]))
    if table is None:
        table = doc.add_table(rows=1, cols=len(df.columns))
        hdr = table.rows[0].cells
        for i, c in enumerate(df.columns):
            hdr[i].text = str(c)
    clear_table_rows_but_header(table)
    for _, row in df.iterrows():
        cells = table.add_row().cells
        for i, col in enumerate(df.columns):
            text = "" if pd.isna(row[col]) else str(row[col])
            cells[i].text = text
            for para in cells[i].paragraphs:
                if para.runs:
                    para.runs[0].font.size = Pt(10)
                para.alignment = (
                    WD_PARAGRAPH_ALIGNMENT.RIGHT if _is_numeric_string(text)
                    else WD_PARAGRAPH_ALIGNMENT.LEFT
                )

def add_total_bottom_right(doc: Document, total: Optional[str]):
    if not total:
        return
    p = doc.add_paragraph()
    p.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
    run = p.add_run(f"Total TTC CHF {total}")
    run.bold = True
    run.font.size = Pt(11)

def process_pdf_to_docx(pdf_file_like, template_docx_bytes: bytes, placeholder_overrides: Optional[Dict[str, str]] = None) -> bytes:
    text, tables = extract_text_and_tables_from_pdf(pdf_file_like)
    fields = parse_fields_from_text(text)
    if placeholder_overrides:
        fields.update({k: v for k, v in placeholder_overrides.items() if v is not None})
    df = None
    if tables:
        df = max(tables, key=lambda d: (d.shape[0], d.shape[1]))
        df.columns = [str(c).strip() for c in df.columns]
    doc = Document(BytesIO(template_docx_bytes))
    replace_placeholders(doc, fields)
    if df is not None and df.shape[0] > 0:
        insert_df_into_table(doc, df)
    add_total_bottom_right(doc, fields.get("Total TTC CHF"))
    out = BytesIO()
    doc.save(out)
    out.seek(0)
    return out.read()
