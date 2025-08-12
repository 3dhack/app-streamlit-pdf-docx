# extract_and_fill.py — fix11: safe table insertion (build with add_table then move), keep fix10 features
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

from docx.oxml import OxmlElement
from docx.oxml.ns import qn

DATE_RE = re.compile(r"\b([0-3]?\d)[./-]([01]?\d)[./-]([12]\d{3})\b")

COLUMNS_TARGET = ["Pos", "Référence", "Désignation", "Unité", "Qté", "Prix unit.", "Px u. Net", "Total CHF", "TVA"]

def today_ch() -> str:
    return datetime.now(ZoneInfo("Europe/Zurich")).strftime("%d.%m.%Y")

def _strip_accents(s: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c))

def _insert_missing_spaces(text: str) -> str:
    text = re.sub(r"(\d)(PC|PCE|KG|M|MM|CM|L)\b", r"\1 \2", text)
    text = re.sub(r"(\d)[A-Za-z]", lambda m: m.group(0)[0] + " " + m.group(0)[1:], text)
    return text

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
                    if not raw or len(raw) < 2:
                        continue
                    header = raw[0]
                    rows = raw[1:]
                    if not any(x for x in header):
                        continue
                    df = pd.DataFrame(rows, columns=[(h or "").strip() for h in header])
                    if df.shape[1] >= 3 and df.shape[0] >= 1:
                        tables.append(_clean_df(df))
            except Exception:
                pass
            try:
                raw_single = page.extract_table()
                if raw_single and len(raw_single) > 1:
                    header = raw_single[0]
                    rows = raw_single[1:]
                    if any(x for x in header):
                        df = pd.DataFrame(rows, columns=[(h or "").strip() for h in header])
                        if df.shape[1] >= 3 and df.shape[0] >= 1:
                            tables.append(_clean_df(df))
            except Exception:
                pass
    unique, sigs = [], set()
    for df in tables:
        sig = (tuple(df.columns), df.shape)
        if sig not in sigs:
            sigs.add(sig)
            unique.append(df)
    return "\n".join(texts), unique

def _clean_df(df: pd.DataFrame) -> pd.DataFrame:
    df.columns = [str(c).strip() for c in df.columns]
    df = df.applymap(lambda x: (str(x).strip() if x is not None else ""))
    df = df.loc[~(df.apply(lambda r: all((not str(v).strip()) for v in r), axis=1))]
    return df.reset_index(drop=True)

def parse_fields_from_text(text: str) -> Dict[str, str]:
    """Parse key header fields from raw text; keep original accents/case when possible."""
    fields: Dict[str, str] = {}

    raw = text.replace("\xa0", " ")
    norm = _strip_accents(raw).lower()

    # N° commande fournisseur (CF) — force uppercase
    m_norm = re.search(r"commande fournisseur n[°o]\s*([A-Za-z0-9\-_]+)", norm, flags=re.IGNORECASE)
    if m_norm:
        cf = m_norm.group(1).strip()
        fields["N°commande fournisseur"] = cf.upper()
        fields["Commande fournisseur"] = cf.upper()

    # Notre référence : take from raw, but trim before "No TVA"/"N° TVA"/"TVA"
    m_line = re.search(r"(?i)(Notre\s+référence\s*:\s*)(.*)", raw)
    if m_line:
        after = m_line.group(2).strip()
        after_norm = _strip_accents(after).lower()
        cut_tokens = ["no tva", "n° tva", "n o tva", "tva", "no  tva"]
        cut_idx = None
        for tok in cut_tokens:
            idx = after_norm.find(tok)
            if idx != -1:
                cut_idx = idx
                break
        if cut_idx is not None:
            value = after[:cut_idx].strip(" -–—\t·:;")
        else:
            value = after.strip(" -–—\t·:;")
        fields["Notre référence"] = value[:60]

    # Total TTC CHF
    m = re.search(r"(?i)(Montant\s+Total\s+TTC\s+CHF|Total\s+TTC\s+CHF)\s*([0-9'’.,]+)", raw)
    if m:
        fields["Total TTC CHF"] = m.group(2).strip()
    else:
        m = re.search(r"([0-9'’.,]+)\s*(?i)(Montant\s+Total\s+TTC\s+CHF|Total\s+TTC\s+CHF)", raw)
        if m:
            fields["Total TTC CHF"] = m.group(1).strip()

    return fields

def _parse_date_str(s: str) -> Optional[datetime]:
    m = DATE_RE.search(s)
    if not m:
        return None
    d, mth, y = m.groups()
    try:
        return datetime(int(y), int(mth), int(d))
    except ValueError:
        return None

def latest_receipt_date_from_text(text: str) -> Optional[str]:
    dates = []
    norm = _strip_accents(text).lower().replace("\xa0", " ")
    lines = [l.strip() for l in norm.splitlines() if l.strip()]
    for ln in lines:
        if "delai de reception" in ln:
            dt = _parse_date_str(ln)
            if dt:
                dates.append(dt)
    if not dates:
        for m in re.finditer(r"delai de reception.{0,30}?([0-3]?\d[./-][01]?\d[./-][12]\d{3})", norm):
            dt = _parse_date_str(m.group(0))
            if dt:
                dates.append(dt)
    if not dates:
        return None
    return max(dates).strftime("%d.%m.%Y")

def reconstruct_items_from_text(text: str) -> pd.DataFrame:
    unit_words = r"(PC|PCE|PCS|PIECE|PIECES|UN|UNITES?|KG|G|MG|L|ML|M|MM|CM)"
    money = r"[0-9'’.,]+"
    item_re = re.compile(
        rf"^\s*(?P<pos>\d{{1,4}})\s+"
        rf"(?P<ref>\d{{3,}})\s+"
        rf"(?P<designation>.+?)\s+"
        rf"(?P<unite>{unit_words})\s+"
        rf"(?P<qte>\d+)\s+"
        rf"(?P<pu>{money})\s+"
        rf"(?P<pxu>{money})\s+"
        rf"(?P<total>{money})"
        rf"(?:\s+(?P<tva>\d{{2,3}}))?\s*$",
        re.IGNORECASE
    )
    stop_cues = ("total", "recapitulation", "code tva", "montant", "taux")

    rows: List[Dict[str, str]] = []
    current = None
    started = False
    for raw_ln in text.splitlines():
        ln = raw_ln.strip()
        if not ln:
            continue

        m = item_re.match(ln)
        if m:
            pos = m.group("pos")
            try:
                if int(pos) % 10 != 0:
                    if started:
                        break
                    else:
                        continue
            except Exception:
                if started:
                    break
                else:
                    continue

            started = True
            if current:
                rows.append(current)
            gd = m.groupdict()
            current = {
                "Pos": gd.get("pos",""),
                "Référence": gd.get("ref",""),
                "Désignation": (gd.get("designation","") or "").strip(),
                "Unité": gd.get("unite","").upper(),
                "Qté": gd.get("qte",""),
                "Prix unit.": gd.get("pu",""),
                "Px u. Net": gd.get("pxu",""),
                "Total CHF": gd.get("total",""),
                "TVA": gd.get("tva","") or "",
            }
            continue

        if started and re.match(r"^\d", ln):
            break

        low = _strip_accents(ln).lower()
        if started and any(low.startswith(c) for c in stop_cues):
            break

        if any(low.startswith(pref) for pref in ("tarif douanier", "pays d'origine", "indice :", "delai de reception :")):
            continue

        if started and current:
            if not any(low.startswith(c) for c in stop_cues):
                if current["Désignation"]:
                    current["Désignation"] += " " + ln.strip()
                else:
                    current["Désignation"] = ln.strip()

    if current:
        rows.append(current)

    df = pd.DataFrame(rows, columns=COLUMNS_TARGET)
    return df

def truncate_after_items_block(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return df
    keep = []
    started = False
    for _, r in df.iterrows():
        first = ""
        for v in r.tolist():
            s = str(v).strip() if v is not None else ""
            if s:
                first = s
                break
        if re.fullmatch(r"\d{1,4}", first):
            try:
                if int(first) % 10 == 0:
                    keep.append(r)
                    started = True
                    continue
                else:
                    if started:
                        break
                    else:
                        continue
            except Exception:
                if started:
                    break
                else:
                    continue
        else:
            if started:
                break
            else:
                continue
    if not keep:
        return df
    out = pd.DataFrame(keep, columns=df.columns).reset_index(drop=True)
    return out

def clean_items_df_keep_full(df: pd.DataFrame) -> pd.DataFrame:
    df2 = truncate_after_items_block(df.copy())
    filtered = []
    for _, r in df2.iterrows():
        first_non_empty = ""
        for v in r.tolist():
            s = str(v).strip() if v is not None else ""
            if s:
                first_non_empty = _strip_accents(s.lower())
                break
        if first_non_empty.startswith("indice :") or first_non_empty.startswith("delai de reception :"):
            continue
        filtered.append(r)
    new_df = pd.DataFrame(filtered, columns=df2.columns) if filtered else df2.iloc[0:0].copy()
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

def find_paragraph_anchor(doc: Document) -> Optional[object]:
    """Find the paragraph containing 'Cond. de paiement' (accent-insensitive, plural accepted)."""
    target_re = re.compile(r"cond\.\s*de\s*paiement[s]?", re.IGNORECASE)
    for p in doc.paragraphs:
        txt = _strip_accents(p.text).lower()
        if target_re.search(txt):
            return p
    return None

def insert_paragraph_after(paragraph, text=""):
    """Insert a new paragraph after given paragraph and return it."""
    new_p = OxmlElement('w:p')
    paragraph._p.addnext(new_p)
    from docx.text.paragraph import Paragraph
    para = Paragraph(new_p, paragraph._parent)
    if text:
        para.add_run(text)
    return para

def set_table_borders(table):
    try:
        table.style = "Table Grid"
    except Exception:
        pass
    tbl = table._element
    tblPr = tbl.tblPr if tbl.tblPr is not None else OxmlElement('w:tblPr')
    tblBorders = OxmlElement('w:tblBorders')
    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        element = OxmlElement(f'w:{edge}')
        element.set(qn('w:val'), 'single')
        element.set(qn('w:sz'), '8')
        element.set(qn('w:space'), '0')
        element.set(qn('w:color'), 'auto')
        tblBorders.append(element)
    tblPr.append(tblBorders)
    tbl.append(tblPr)

def insert_df_two_lines_below_anchor(doc: Document, df: pd.DataFrame):
    """Insert df as table exactly two blank lines below the 'Cond. de paiement' paragraph.
       Build table with document.add_table() to ensure tblGrid exists, then move it.
    """
    if df is None or df.empty:
        return
    df = df.copy()
    df.columns = [str(c) for c in df.columns]

    anchor = find_paragraph_anchor(doc)

    # Create table safely at end
    tbl = doc.add_table(rows=1, cols=len(df.columns))
    # header
    for i, c in enumerate(df.columns):
        tbl.rows[0].cells[i].text = str(c)
    # rows
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
    set_table_borders(tbl)

    # Move table after anchor + two blank lines (or leave at end if no anchor)
    if anchor is not None:
        p1 = insert_paragraph_after(anchor, "")
        p2 = insert_paragraph_after(p1, "")
        # move underlying element
        p2._p.addnext(tbl._tbl)

def process_pdf_to_docx(pdf_bytes: bytes, template_docx_bytes: bytes) -> Tuple[bytes, Dict[str, str], pd.DataFrame]:
    text, tables = extract_text_and_tables_from_pdf(BytesIO(pdf_bytes))
    fields = parse_fields_from_text(text)
    fields["date du jour"] = today_ch()

    # Délai de réception (max) from text
    from_text = []
    norm = _strip_accents(text).lower().replace("\xa0"," ")
    for ln in [l.strip() for l in norm.splitlines() if l.strip()]:
        if "delai de reception" in ln:
            m = DATE_RE.search(ln)
            if m:
                d, mth, y = m.groups()
                try:
                    from_text.append(datetime(int(y), int(mth), int(d)))
                except Exception:
                    pass
    if from_text:
        fields["Délai de réception"] = max(from_text).strftime("%d.%m.%Y")

    # Items table: prefer pdf tables; else reconstruct from text; then truncate and clean
    if tables:
        base_df = max(tables, key=lambda d: (d.shape[0], d.shape[1]))
        items_df = clean_items_df_keep_full(base_df)
    else:
        items_df = reconstruct_items_from_text(text)
        items_df = clean_items_df_keep_full(items_df)

    # Build doc with placeholders replaced
    doc = Document(BytesIO(template_docx_bytes))
    replace_placeholders_everywhere(doc, fields)
    out = BytesIO(); doc.save(out); out.seek(0)
    return out.getvalue(), fields, items_df

# Public helpers re-exported for app
def insert_items_table_at_position(doc_bytes: bytes, df: pd.DataFrame) -> bytes:
    doc = Document(BytesIO(doc_bytes))
    insert_df_two_lines_below_anchor(doc, df)
    out = BytesIO(); doc.save(out); out.seek(0)
    return out.getvalue()
