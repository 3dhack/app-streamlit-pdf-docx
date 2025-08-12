# PDF → DOCX (Streamlit) — fix5
- **Délai de réception** : détecté d'abord dans les **lignes du tableau** (libellé + date), puis **fallback** sur le **texte brut** (lignes contenant "Délai de réception : dd.mm.yyyy"). On prend la **date maximale**.
- Mapping compatibilité: remplit aussi « date Délai de livraison » si présent dans l'ancien modèle.
- Auto-préremplissage, éditeur de tableau, génération robuste.

## Déploiement
1) Remplace tes fichiers dans le repo GitHub (commit + push)
2) Streamlit Cloud redéploie
3) Vérifie que « Délai de réception » est prérempli
