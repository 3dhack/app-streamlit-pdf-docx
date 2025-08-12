# streamlit_app.py
import streamlit as st
from io import BytesIO
from extract_and_fill import extract_text_and_tables_from_pdf, parse_fields_from_text, process_pdf_to_docx

st.set_page_config(page_title="PDF → DOCX (Commande fournisseur)", layout="wide")

st.title("PDF → DOCX : Remplissage automatique de commande fournisseur")
st.caption("Upload le PDF de commande et ton modèle .docx, ajuste les champs si besoin, puis télécharge le document généré.")

with st.sidebar:
    st.header("Étapes")
    st.markdown("1. Uploader le **PDF** de la commande")
    st.markdown("2. Uploader le **modèle .docx** avec les placeholders « … »")
    st.markdown("3. Vérifier les champs détectés et le tableau")
    st.markdown("4. Cliquer **Générer le DOCX**")
    st.markdown("---")
    st.markdown("**Astuce :** Si le tableau ne s'extrait pas correctement, envoie-moi un échantillon.")

pdf_file = st.file_uploader("PDF de la commande", type=["pdf"])
docx_template = st.file_uploader("Modèle Word (.docx) avec placeholders", type=["docx"])

if pdf_file and docx_template:
    with st.spinner("Analyse du PDF..."):
        pdf_bytes = pdf_file.read()
        text, tables = extract_text_and_tables_from_pdf(BytesIO(pdf_bytes))
        fields = parse_fields_from_text(text)

    st.subheader("Champs détectés")
    st.json(fields or {"info": "Aucun champ détecté automatiquement."})

    st.subheader("Tableau détecté (plus grand)")
    if tables:
        df = max(tables, key=lambda d: (d.shape[0], d.shape[1]))
        st.dataframe(df, use_container_width=True)
    else:
        st.warning("Aucun tableau détecté dans le PDF.")

    st.subheader("Ajuster / compléter les placeholders")
    col1, col2 = st.columns(2)
    with col1:
        p_cmd = st.text_input("« N°commande fournisseur »", fields.get("N°commande fournisseur", ""))
        p_cmd2 = st.text_input("« Commande fournisseur » (si présent dans le modèle)", fields.get("Commande fournisseur", p_cmd))
        p_date = st.text_input("« date du jour »", fields.get("date du jour", ""))
    with col2:
        p_delai = st.text_input("« date Délai de livraison »", fields.get("date Délai de livraison", ""))
        p_total = st.text_input("« Total TTC CHF »", fields.get("Total TTC CHF", ""))
        p_cond = st.text_input("« Cond. de paiement »", fields.get("Cond. de paiement", ""))

    if st.button("Générer le DOCX"):
        overrides = {
            "N°commande fournisseur": p_cmd,
            "Commande fournisseur": p_cmd2,
            "date du jour": p_date,
            "date Délai de livraison": p_delai,
            "Total TTC CHF": p_total,
            "Cond. de paiement": p_cond,
        }
        docx_template.seek(0)
        out_bytes = process_pdf_to_docx(BytesIO(pdf_bytes), docx_template.read(), placeholder_overrides=overrides)
        st.success("Document généré !")
        st.download_button(
            "Télécharger le DOCX généré",
            data=out_bytes,
            file_name="commande_remplie.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
else:
    st.info("Téléverse un PDF **et** un modèle .docx pour commencer.")
