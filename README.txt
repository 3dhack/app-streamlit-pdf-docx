# PDF → DOCX (Streamlit) — fix6
Modifs demandées :
1) **Date du jour** : forcée à la date actuelle (Europe/Zurich) — format dd.mm.yyyy.
2) **Cond. de paiement** : supprimé de l'UI et des remplacements.
3) **Tableau** : reprise du **tableau complet** du PDF dans le Word, **sans édition**.
   - Suppression des **lignes** dont la première cellule non vide commence par "Indice :" ou "Délai de réception :".
   - Suppression de la **colonne "TVA"** (insensible aux accents).
   - Insertion dans la **première table** si le nombre de colonnes correspond, sinon création d'une table à la fin.

## Déploiement
- Remplace les fichiers dans ton repo GitHub (commit + push)
- Streamlit Cloud redéploie automatiquement
- Intègre dans WordPress avec `?embed=true` si besoin
