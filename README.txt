# PDF → DOCX (Streamlit) — fix7
- Si `pdfplumber` n'extrait aucun tableau, on reconstruit le tableau d'articles à partir du texte brut (pattern compatible avec ton exemple).
- Désignations multilignes gérées (ex. "G3/4").
- Lignes "Indice :" / "Délai de réception :" ignorées et colonne "TVA" supprimée.
- Date du jour forcée (Europe/Zurich).

## Déploiement
1) Pousse ces fichiers dans un repo GitHub public
2) Streamlit Cloud: Deploy public app → `streamlit_app.py`
3) Teste l'URL directe, puis intègre dans WordPress avec `?embed=true`
