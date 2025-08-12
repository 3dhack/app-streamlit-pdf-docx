# streamlit_app.py — fix19 (static template support)
import streamlit as st
from pathlib import Path
from io import BytesIO

from extract_and_fill import process_pdf_to_docx, build_final_doc

st.set_page_config(page_title="PDF → DOCX (Commande fournisseur)", layout="wide")
st.title("PDF → DOCX : Remplissage automatique")
st.caption("Modèle Word chargé automatiquement depuis le dépôt (template.docx). Téléversement facultatif.")

# --- Template handling ---
TEMPLATE_PATH = Path(__file__).parent / "template.docx"
tmpl_bytes = None
template_source = ""

if TEMPLATE_PATH.exists():
    tmpl_bytes = TEMPLATE_PATH.read_bytes()
    template_source = "from_repo"

# Allow optional override via sidebar
with st.sidebar:
    st.header("Options")
    override = st.checkbox("Remplacer le modèle embarqué (uploader un .docx)", value=False, help="Coche si tu veux tester un autre modèle que template.docx")
    st.markdown("---")
    st.header("Étapes")
    st.markdown("1. Uploader le **PDF**")
    if not TEMPLATE_PATH.exists():
        st.markdown("2. Uploader le **modèle .docx** *(le dépôt ne contient pas `template.docx`)*")
    st.markdown("3. Générer le **DOCX**")

if override or not TEMPLATE_PATH.exists():
    up = st.file_uploader("Modèle Word (.docx)", type=["docx"], help="Sinon, le fichier `template.docx` du repo sera utilisé")
    if up:
        tmpl_bytes = up.read()
        template_source = "uploaded"

pdf_file = st.file_uploader("PDF de la commande", type=["pdf"])

# State
for k in ["fields", "items_df", "doc_with_placeholders"]:
    if k not in st.session_state:
        st.session_state[k] = None

def _analyze(pdf_bytes: bytes, tmpl_bytes: bytes):
    out_doc_bytes, fields, items_df = process_pdf_to_docx(pdf_bytes, tmpl_bytes)
    st.session_state["fields"] = fields
    st.session_state["items_df"] = items_df
    st.session_state["doc_with_placeholders"] = out_doc_bytes

# Run analysis automatically when both inputs are present
if pdf_file and tmpl_bytes and st.session_state["fields"] is None:
    with st.spinner("Analyse du PDF..."):
        _analyze(pdf_file.read(), tmpl_bytes)

# Allow re-analyze
col1, col2 = st.columns(2)
with col1:
    if st.button("🔁 Réanalyser"):
        if pdf_file and tmpl_bytes:
            with st.spinner("Ré-analyse du PDF..."):
                _analyze(pdf_file.read(), tmpl_bytes)
        else:
            st.warning("Merci de fournir le PDF et un modèle (ou `template.docx` dans le repo).")
with col2:
    st.write(("Modèle utilisé : **repo/template.docx**" if template_source=="from_repo" else ("Modèle utilisé : **upload**" if template_source=="uploaded" else "Aucun modèle trouvé")))

# Show fields preview
fields = st.session_state.get("fields") or {}
if fields:
    st.subheader("Champs détectés")
    st.write({
        "N°commande fournisseur": fields.get("N°commande fournisseur", ""),
        "Commande fournisseur": fields.get("Commande fournisseur", ""),
        "Notre référence": fields.get("Notre référence", ""),
        "date du jour": fields.get("date du jour", ""),
        "Délai de réception": fields.get("Délai de réception", ""),
        "Délai de livraison": fields.get("Délai de livraison", ""),
        "Total TTC CHF": fields.get("Total TTC CHF", ""),
    })

# Dataframe preview
st.subheader("Aperçu du tableau")
items_df = st.session_state.get("items_df")
if items_df is not None and not items_df.empty:
    st.dataframe(items_df, use_container_width=True)
else:
    st.info("Le tableau sera reconstruit automatiquement si indisponible.")

# Generate final doc
if st.button("🧾 Générer le DOCX final", disabled=not (tmpl_bytes and st.session_state.get("doc_with_placeholders"))):
    try:
        base_doc_bytes = st.session_state["doc_with_placeholders"]
        total_ttc = (st.session_state["fields"] or {}).get("Total TTC CHF", "")
        final_doc = build_final_doc(base_doc_bytes, st.session_state["items_df"], total_ttc)

        # File name: "Facture {Commande fournisseur}.docx"
        commande = (st.session_state["fields"] or {}).get("Commande fournisseur", "").strip()
        filename = f"Facture {commande}.docx" if commande else "Facture.docx"

        st.success("DOCX généré !")
        st.download_button(
            "Télécharger le DOCX généré",
            data=final_doc,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    except Exception as e:
        st.exception(e)
