# streamlit_app.py — fix21
import streamlit as st
from pathlib import Path
from extract_and_fill import process_pdf_to_docx, build_final_doc

st.set_page_config(page_title="PDF → DOCX (Commande fournisseur)", layout="wide")
st.title("PDF → DOCX : Remplissage automatique")
st.caption("Fix21 : reconstruction des lignes stoppée au 'Total CHF' de chaque article (anti-spillover).")

TEMPLATE_PATH = Path(__file__).parent / "template.docx"
tmpl_bytes = TEMPLATE_PATH.read_bytes() if TEMPLATE_PATH.exists() else None

with st.sidebar:
    st.header("Étapes")
    st.markdown("1. Uploader le **PDF**")
    if tmpl_bytes is None:
        st.markdown("2. Uploader le **modèle .docx** (ou place `template.docx` dans le repo)")
    st.markdown("3. Générer le **DOCX**")

if tmpl_bytes is None:
    up = st.file_uploader("Modèle Word (.docx)", type=["docx"])
    if up: tmpl_bytes = up.read()

pdf_file = st.file_uploader("PDF de la commande", type=["pdf"])

for k in ["fields", "items_df", "doc_with_placeholders"]:
    if k not in st.session_state:
        st.session_state[k] = None

def _analyze(pdf_bytes, tmpl_bytes):
    out_doc_bytes, fields, items_df = process_pdf_to_docx(pdf_bytes, tmpl_bytes)
    st.session_state["fields"] = fields
    st.session_state["items_df"] = items_df
    st.session_state["doc_with_placeholders"] = out_doc_bytes

if pdf_file and tmpl_bytes and st.session_state["fields"] is None:
    with st.spinner("Analyse du PDF..."):
        _analyze(pdf_file.read(), tmpl_bytes)

if st.button("🔁 Réanalyser"):
    if pdf_file and tmpl_bytes:
        with st.spinner("Ré-analyse du PDF..."):
            _analyze(pdf_file.read(), tmpl_bytes)
    else:
        st.warning("Fournis le PDF et un modèle (ou `template.docx`).")

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

if st.button("🧾 Générer le DOCX final", disabled=not (tmpl_bytes and st.session_state.get("doc_with_placeholders"))):
    base_doc_bytes = st.session_state["doc_with_placeholders"]
    total_ttc = (st.session_state["fields"] or {}).get("Total TTC CHF", "")
    final_doc = build_final_doc(base_doc_bytes, st.session_state["items_df"], total_ttc)

    commande = (st.session_state["fields"] or {}).get("Commande fournisseur", "").strip()
    filename = f"Facture {commande}.docx" if commande else "Facture.docx"

    st.success("DOCX généré !")
    st.download_button("Télécharger le DOCX généré", data=final_doc, file_name=filename, mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
