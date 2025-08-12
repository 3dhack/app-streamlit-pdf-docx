# streamlit_app.py — fix6 UI (no Cond. de paiement, no table editor, auto today's date)
import streamlit as st
import pandas as pd
from io import BytesIO
from docx import Document

from extract_and_fill import (
    process_pdf_to_docx,
    insert_any_df_into_doc,
)

st.set_page_config(page_title="PDF → DOCX (Commande fournisseur)", layout="wide")

st.title("PDF → DOCX : Remplissage automatique")
st.caption("Charge un PDF + un modèle DOCX. La date du jour est forcée (Europe/Zurich). Le tableau complet est repris (sans 'Indice :' / 'Délai de réception :' et sans colonne TVA).")

with st.sidebar:
    st.header("Étapes")
    st.markdown("1. Uploader le **PDF**")
    st.markdown("2. Uploader le **modèle .docx**")
    st.markdown("3. Vérifier l'aperçu")
    st.markdown("4. **Générer** et télécharger le `.docx`")

pdf_file = st.file_uploader("PDF de la commande", type=["pdf"])
docx_template = st.file_uploader("Modèle Word (.docx)", type=["docx"])

for k in ["fields", "items_df", "tmpl_bytes", "pdf_bytes", "doc_with_placeholders"]:
    if k not in st.session_state:
        st.session_state[k] = None

def _analyze():
    try:
        pdf_bytes = pdf_file.read()
        tmpl_bytes = docx_template.read()
        st.session_state["pdf_bytes"] = pdf_bytes
        st.session_state["tmpl_bytes"] = tmpl_bytes

        out_doc_bytes, fields, items_df = process_pdf_to_docx(pdf_bytes, tmpl_bytes)
        st.session_state["fields"] = fields
        st.session_state["items_df"] = items_df
        st.session_state["doc_with_placeholders"] = out_doc_bytes
    except Exception as e:
        st.exception(e)

if pdf_file and docx_template and (st.session_state["fields"] is None):
    with st.spinner("Analyse du PDF..."):
        _analyze()

if pdf_file and docx_template and st.session_state["fields"] is not None:
    if st.button("🔁 Réanalyser à partir des fichiers envoyés"):
        with st.spinner("Ré-analyse du PDF..."):
            _analyze()

fields = st.session_state["fields"] or {}
if fields:
    st.subheader("Champs utilisés")
    st.write({
        "N°commande fournisseur": fields.get("N°commande fournisseur", ""),
        "Commande fournisseur": fields.get("Commande fournisseur", ""),
        "date du jour": fields.get("date du jour", ""),
        "Délai de réception": fields.get("Délai de réception", ""),
        "Total TTC CHF": fields.get("Total TTC CHF", ""),
    })

st.subheader("Aperçu du tableau extrait (après nettoyage)")
items_df = st.session_state["items_df"]
if items_df is not None and not items_df.empty:
    st.dataframe(items_df, use_container_width=True)
else:
    st.info("Aucun tableau détecté dans le PDF.")

disabled = not (st.session_state["tmpl_bytes"] and st.session_state["doc_with_placeholders"])
if st.button("🧾 Générer le DOCX final", disabled=disabled):
    try:
        # Recharger le doc (avec placeholders déjà remplacés) et insérer le tableau complet
        base_doc_bytes = st.session_state["doc_with_placeholders"]
        doc = Document(BytesIO(base_doc_bytes))
        insert_any_df_into_doc(doc, st.session_state["items_df"])

        out = BytesIO()
        doc.save(out); out.seek(0)
        st.success("DOCX généré !")
        st.download_button(
            "Télécharger le DOCX généré",
            data=out.getvalue(),
            file_name="commande_remplie.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    except Exception as e:
        st.exception(e)
