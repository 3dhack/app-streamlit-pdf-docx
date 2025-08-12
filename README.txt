# PDF → DOCX (Streamlit) — fix14
Corrections & ajouts :
- Remplacement de `insert_paragraph_after` natif (qui n’existe pas) par une version custom (XML), pour insérer le tableau **2 lignes** sous « Cond. de paiement ».
- **Total TTC CHF** est inséré **immédiatement** après le tableau (aligné à droite).
- **Deux** retours à la ligne **exactement** sont insérés après le total; les paragraphes vides supplémentaires juste après sont **supprimés**.
- Bordures visibles, Pos/Désignation à gauche, largeurs préférées conservées.

Déploiement : push → Streamlit Cloud redéploie.
