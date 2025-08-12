# streamlit_app.py — fix11 UI: same as fix10, now safe table creation
import streamlit as st
import pandas as pd
from io import BytesIO

from extract_and_fill import (
    process_pdf_to_docx,
    insert_items_table_at_position,
)

st.set_page_config(page_title="PDF → DOCX (Commande fournisseur)", layout="wide")

st.title("PDF → DOCX : Remplissage automatique")
st.caption("CF en majuscule, « Notre référence » coupée avant « No TVA ». Le tableau est inséré deux lignes sous « Cond. de paiement » (création sûre).")

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
        "Notre référence": fields.get("Notre référence", ""),
        "date du jour": fields.get("date du jour", ""),
        "Délai de réception": fields.get("Délai de réception", ""),
        "Total TTC CHF": fields.get("Total TTC CHF", ""),
    })

st.subheader("Aperçu du tableau (tronqué & nettoyé)")
items_df = st.session_state["items_df"]
if items_df is not None and not items_df.empty:
    st.dataframe(items_df, use_container_width=True)
else:
    st.warning("Aucun tableau détecté ou reconstruit.")

if st.button("🧾 Générer le DOCX final", disabled=not (st.session_state.get('tmpl_bytes') and st.session_state.get('doc_with_placeholders'))):
    try:
        base_doc_bytes = st.session_state["doc_with_placeholders"]
        final_doc = insert_items_table_at_position(base_doc_bytes, st.session_state["items_df"])

        st.success("DOCX généré !")
        st.download_button(
            "Télécharger le DOCX généré",
            data=final_doc,
            file_name="commande_remplie.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    except Exception as e:
        st.exception(e)
