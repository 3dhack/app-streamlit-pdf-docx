# PDF → DOCX (Streamlit) — fix4
- **Délai de réception** : extrait en cherchant les lignes du tableau contenant 'Délai de réception' (accent-insensible), puis prend **la date la plus grande de la ligne** et enfin **la plus grande de toutes les lignes**.
- Remplacement compatible avec anciens modèles via la clé **« date Délai de livraison »** qui reçoit la même valeur.
- Préremplissage auto, édition du tableau, génération robuste.

## Déploiement
1) Remplace tes fichiers dans le repo GitHub (commit + push)
2) Streamlit Cloud redéploie
3) Vérifie que « Délai de réception » est prérempli avec la **date maximale** trouvée dans le tableau
