# PDF → DOCX (Streamlit) — fix10
Nouveautés :
1) **« Notre référence »** : on capture après `Notre référence :` mais **on coupe avant** `No TVA` / `N° TVA` / `TVA`. Résultat = seulement *nom prénom*.
2) **Position du tableau** : insertion **exactement deux lignes sous** le paragraphe « Cond. de paiement » (insensible aux accents/pluriel). Si l'ancre n'est pas trouvée, le tableau est ajouté en bas du document.
3) On conserve toutes les features précédentes : CF en majuscule, date du jour (Europe/Zurich), Délai de réception = date max, reconstruction et arrêt des items (10/20/30...), bordures de tableau, suppression des lignes 'Indice :'/'Délai de réception :' et de la colonne 'TVA'.

## Déploiement
1) Pousse ces fichiers dans ton repo GitHub public
2) Streamlit Cloud: Deploy public app → `streamlit_app.py`
3) Teste l'URL directe, puis intègre dans WordPress avec `?embed=true`
