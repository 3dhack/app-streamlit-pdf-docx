# PDF → DOCX (Streamlit) — fix9
Nouveautés :
1) **Tableau avec bordures** (style *Table Grid* + bordures forcées en oxml).
2) **CF en majuscule** : 'Commande fournisseur' et 'N°commande fournisseur' sont écrits en UPPERCASE.
3) **Notre référence** : extraction automatique depuis 'Notre référence : xxxxx' dans le PDF, champ Word « Notre référence » rempli.

Toujours en place :
- Date du jour forcée (Europe/Zurich).
- Délai de réception = plus grande date trouvée dans le PDF.
- Reconstruction du tableau si nécessaire + arrêt au dernier item (10/20/30...).

## Déploiement
1) Remplace tes fichiers dans le repo GitHub (commit + push)
2) Streamlit Cloud redéploie
3) Intègre dans WordPress (`?embed=true`) si besoin
