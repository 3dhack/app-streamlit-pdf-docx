# streamlit_app.py — fix28
import streamlit as st
from pathlib import Path
from extract_and_fill import process_pdf_to_docx, build_final_doc

st.set_page_config(page_title="PDF → DOCX (Commande fournisseur)", layout="wide")
st.title("PDF → DOCX : Remplissage automatique")

# --- Init dynamic uploader keys ---
if "reset_count" not in st.session_state:
    st.session_state.reset_count = 0
if "pdf_uploader_key" not in st.session_state:
    st.session_state.pdf_uploader_key = f"pdf_uploader_{st.session_state.reset_count}"
if "docx_uploader_key" not in st.session_state:
    st.session_state.docx_uploader_key = f"docx_uploader_{st.session_state.reset_count}"
# --- End init ---


# --- Bouton Réinitialiser ---
if st.button("🔄 Réinitialiser"):
    # Clear working state
    for key in ["fields", "items_df", "doc_with_placeholders"]:
        if key in st.session_state:
            del st.session_state[key]
    # Bump keys so uploaders visually reset
    st.session_state.reset_count += 1
    st.session_state.pdf_uploader_key = f"pdf_uploader_{st.session_state.reset_count}"
    st.session_state.docx_uploader_key = f"docx_uploader_{st.session_state.reset_count}"
    st.rerun()

# --- Fin Réinitialiser ---


TEMPLATE_PATH = Path(__file__).parent / "template.docx"
tmpl_bytes = TEMPLATE_PATH.read_bytes() if TEMPLATE_PATH.exists() else None

pdf_file = st.file_uploader("PDF de la commande", type=["pdf"], key=st.session_state.pdf_uploader_key)
if tmpl_bytes is None:
    up = st.file_uploader("Modèle Word (.docx)", type=["docx"], key=st.session_state.docx_uploader_key)
    if up:
        tmpl_bytes = up.read()

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

if pdf_file and st.button("🔁 Réanalyser"):
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


ready_to_generate = bool(tmpl_bytes and st.session_state.get("doc_with_placeholders"))
if ready_to_generate:
    if st.button("🧾 Générer le DOCX"):
        base_doc_bytes = st.session_state["doc_with_placeholders"]
        total_ttc = (st.session_state["fields"] or {}).get("Total TTC CHF", "")
        final_doc = build_final_doc(base_doc_bytes, st.session_state["items_df"], total_ttc)

        commande = (st.session_state["fields"] or {}).get("Commande fournisseur", "").strip()
        filename = f"Facture {commande}.docx" if commande else "Facture.docx"

        st.success("DOCX généré !")
        st.download_button("🟦 Télécharger le DOCX", data=final_doc, file_name=filename, mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
else:
    st.info("Importe un PDF (et un modèle si nécessaire) puis lance l'analyse pour afficher le bouton de génération.")
