# PDF → DOCX (Streamlit) — version PRO (mappage + édition)

Améliorations incluses :
- Détection + **mappage des colonnes** vers Pos./Référence/Désignation/Qté/Prix/Total
- **Insertion robuste** dans la première table du modèle (ou création)
- Remplacement de placeholders « … » **partout** (paragraphes, tables, en-têtes/pieds)
- Ajout du **Total TTC CHF** en bas à droite
- **Édition manuelle** du tableau avant génération (Data Editor Streamlit)

## Déploiement
1. Push dans un repo GitHub **public**
2. Streamlit Cloud → **Deploy a public app from GitHub**
3. Main file: `streamlit_app.py`

## Intégration WordPress
```html
<iframe src="https://xxx.streamlit.app?embed=true" width="100%" height="900" style="border:none;"></iframe>
```

## Conseils
- Si certaines colonnes ne sont pas reconnues, édite le tableau dans l’interface avant de générer.
- Pour une extraction plus fiable sur des PDF "difficiles", on peut ajouter un fallback Camelot/Tabula (nécessite deps système, pas dispo sur Streamlit Cloud gratuit).
