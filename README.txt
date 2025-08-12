# PDF → DOCX (Streamlit) — fix 2 (parsing robuste)

- Parsing **insensible aux accents** (commande fournisseur, délai, etc.)
- Gestion du **Total TTC** même si le **nombre précède le libellé**
- Normalisation espaces (cas comme `1'347.36Montant`)
- Auto-préremplissage, éditeur de tableau, génération robuste

## Déploiement
1) Pousse ces fichiers dans ton repo GitHub public
2) Streamlit Cloud redéploie
3) Test en direct puis intégration WordPress (?embed=true)
