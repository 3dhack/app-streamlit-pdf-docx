# streamlit_app.py
import streamlit as st
import pandas as pd
from io import BytesIO
from docx import Document

from extract_and_fill import (
    process_pdf_to_docx,
    EXPECTED_COLUMNS
)

st.set_page_config(page_title="PDF → DOCX (Commande fournisseur)", layout="wide")

st.title("PDF → DOCX : Remplissage automatique")
st.caption("Charge un PDF + un modèle DOCX. Corrige la détection si besoin, puis génère le document final.")

with st.sidebar:
    st.header("Étapes")
    st.markdown("1. Uploader le **PDF**")
    st.markdown("2. Uploader le **modèle .docx** avec placeholders « … »")
    st.markdown("3. Vérifier/éditer le **tableau** et les **champs**")
    st.markdown("4. **Générer** et télécharger le `.docx`")

pdf_file = st.file_uploader("PDF de la commande", type=["pdf"])
docx_template = st.file_uploader("Modèle Word (.docx)", type=["docx"])

if pdf_file and docx_template:
    # First pass to get auto mapping and fields
    pdf_bytes = pdf_file.read()
    tmpl_bytes = docx_template.read()

    # Optional: detect placeholders present in the DOCX
    with st.expander("Placeholders détectés dans le modèle"):
        doc = Document(BytesIO(tmpl_bytes))
        ph = []
        for p in doc.paragraphs:
            txt = p.text
            ph += [seg.strip("« »\xa0 ") for seg in txt.split("«") if "»" in seg]
        for t in doc.tables:
            for row in t.rows:
                for cell in row.cells:
                    txt = cell.text
                    ph += [seg.strip("« »\xa0 ") for seg in txt.split("«") if "»" in seg]
        st.write(sorted(set([p for p in ph if p])))

    st.subheader("Champs (auto-détectés, modifiables)")
    col1, col2, col3 = st.columns(3)
    with col1:
        f_cmd = st.text_input("N°commande fournisseur", value="")
        f_cmd2 = st.text_input("Commande fournisseur", value="")
        f_date = st.text_input("date du jour", value="")
    with col2:
        f_delai = st.text_input("date Délai de livraison", value="")
        f_cond = st.text_input("Cond. de paiement", value="")
        f_total = st.text_input("Total TTC CHF", value="")
    fields_over = {
        "N°commande fournisseur": f_cmd,
        "Commande fournisseur": f_cmd2,
        "date du jour": f_date,
        "date Délai de livraison": f_delai,
        "Cond. de paiement": f_cond,
        "Total TTC CHF": f_total,
    }

    st.subheader("Génération initiale")
    gen = st.button("Analyser le PDF et proposer le tableau")
    if gen:
        out_bytes, detected_fields, items_df = process_pdf_to_docx(
            pdf_bytes, tmpl_bytes, placeholder_overrides=fields_over, custom_mapping=None
        )

        # Merge: prefer user overrides if provided
        for k, v in detected_fields.items():
            if not fields_over.get(k):
                fields_over[k] = v

        st.markdown("**Champs détectés (après fusion)**")
        st.json(fields_over)

        st.markdown("**Tableau proposé (éditable)**")
        if items_df.empty:
            items_df = pd.DataFrame(columns=EXPECTED_COLUMNS)
        edit_df = st.data_editor(items_df, num_rows="dynamic", use_container_width=True)

        st.markdown("---")
        if st.button("Générer le DOCX final"):
            # Rebuild the doc with edited table by injecting mapping via columns kept identical
            # We reconstruct using the same columns expected by the writer
            from extract_and_fill import replace_placeholders_everywhere, find_or_create_items_table, insert_items_into_table, add_total_bottom_right
            from io import BytesIO as _BytesIO

            doc = Document(_BytesIO(tmpl_bytes))
            replace_placeholders_everywhere(doc, fields_over)
            table = find_or_create_items_table(doc, EXPECTED_COLUMNS)
            # clear and insert
            while len(table.rows) > 1:
                table._element.remove(table.rows[1]._element)
            insert_items_into_table(table, edit_df)
            add_total_bottom_right(doc, fields_over.get("Total TTC CHF", ""))

            out = _BytesIO()
            doc.save(out); out.seek(0)
            st.success("DOCX généré avec la table éditée !")
            st.download_button(
                "Télécharger le DOCX généré",
                data=out.getvalue(),
                file_name="commande_remplie.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
else:
    st.info("Téléverse un PDF **et** un modèle .docx pour commencer.")
