# streamlit_app.py — fix24b (clean)
import streamlit as st
from pathlib import Path
from extract_and_fill import process_pdf_to_docx, build_final_doc

import tempfile, subprocess, sys, platform

def convert_docx_bytes_to_pdf_bytes(docx_bytes: bytes) -> bytes | None:
    """
    Try multiple strategies to convert DOCX -> PDF on Linux/Cloud:
    1) LibreOffice ('soffice --headless') if available
    2) unoconv if available
    3) docx2pdf (works mainly on Windows/Mac)
    Returns PDF bytes or None if all strategies fail.
    """
    try:
        with tempfile.TemporaryDirectory() as tmpdir:
            docx_path = Path(tmpdir) / "out.docx"
            pdf_path  = Path(tmpdir) / "out.pdf"
            with open(docx_path, "wb") as f:
                f.write(docx_bytes)

            # Strategy 1: soffice (LibreOffice)
            try:
                subprocess.run(
                    ["soffice", "--headless", "--convert-to", "pdf", "--outdir", str(Path(tmpdir)), str(docx_path)],
                    check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=60
                )
                if pdf_path.exists():
                    return pdf_path.read_bytes()
            except Exception:
                pass

            # Strategy 2: unoconv
            try:
                subprocess.run(
                    ["unoconv", "-f", "pdf", str(docx_path)],
                    check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=60
                )
                if pdf_path.exists():
                    return pdf_path.read_bytes()
                # unoconv may name the file with same base
                for p in Path(tmpdir).glob("*.pdf"):
                    return p.read_bytes()
            except Exception:
                pass

            # Strategy 3: docx2pdf (best on Windows/Mac)
            try:
                from docx2pdf import convert as docx2pdf_convert
                # On non-Windows/macOS this may fail if MS Word is not present
                out_tmp_pdf = Path(tmpdir) / "out_docx2pdf.pdf"
                docx2pdf_convert(str(docx_path), str(out_tmp_pdf))
                if out_tmp_pdf.exists():
                    return out_tmp_pdf.read_bytes()
            except Exception:
                pass

            return None
    except Exception:
        return None


st.set_page_config(page_title="PDF → DOCX (Commande fournisseur)", layout="wide")
st.title("PDF → DOCX : Remplissage automatique")


TEMPLATE_PATH = Path(__file__).parent / "template.docx"
tmpl_bytes = TEMPLATE_PATH.read_bytes() if TEMPLATE_PATH.exists() else None

# Uploaders
pdf_file = st.file_uploader("PDF de la commande", type=["pdf"])
if tmpl_bytes is None:
    up = st.file_uploader("Modèle Word (.docx)", type=["docx"])
    if up:
        tmpl_bytes = up.read()

# State
for k in ["fields", "items_df", "doc_with_placeholders"]:
    if k not in st.session_state:
        st.session_state[k] = None

def _analyze(pdf_bytes, tmpl_bytes):
    out_doc_bytes, fields, items_df = process_pdf_to_docx(pdf_bytes, tmpl_bytes)
    st.session_state["fields"] = fields
    st.session_state["items_df"] = items_df
    st.session_state["doc_with_placeholders"] = out_doc_bytes

# Auto analyze
if pdf_file and tmpl_bytes and st.session_state["fields"] is None:
    with st.spinner("Analyse du PDF..."):
        _analyze(pdf_file.read(), tmpl_bytes)

if st.button("🔁 Réanalyser"):
    if pdf_file and tmpl_bytes:
        with st.spinner("Ré-analyse du PDF..."):
            _analyze(pdf_file.read(), tmpl_bytes)
    else:
        st.warning("Fournis le PDF.")

fields = st.session_state.get("fields") or {}
if fields:
    st.subheader("Champs détectés")
    st.write(fields)

st.subheader("Aperçu du tableau")
items_df = st.session_state.get("items_df")
if items_df is not None and not items_df.empty:
    st.dataframe(items_df, use_container_width=True)
else:
    st.info("Le tableau sera reconstruit si aucune table fiable n'est détectée.")

if st.button("🧾 Générer le DOCX/PDF", disabled=not (tmpl_bytes and st.session_state.get("doc_with_placeholders"))):
    base_doc_bytes = st.session_state["doc_with_placeholders"]
    total_ttc = (st.session_state["fields"] or {}).get("Total TTC CHF", "")
    final_doc = build_final_doc(base_doc_bytes, st.session_state["items_df"], total_ttc)

    commande = (st.session_state["fields"] or {}).get("Commande fournisseur", "").strip()
    filename = f"Facture {commande}.docx" if commande else "Facture.docx"

    st.success("DOCX généré !")
    st.download_button("Télécharger le DOCX généré", data=final_doc, file_name=filename, mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
