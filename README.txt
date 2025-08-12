# PDF → DOCX (Streamlit)

Application Streamlit prête à déployer sur **Streamlit Community Cloud**.
- Upload d'un **PDF de commande**
- Extraction du texte + **tableau** avec `pdfplumber`
- Parsing des **champs clés**
- Remplissage d'un **modèle Word (.docx)** via `python-docx`
- **Téléchargement** du `.docx` généré

## Démarrage local
```bash
python -m venv venv
source venv/bin/activate   # Windows: venv\Scripts\activate
pip install -r requirements.txt
streamlit run streamlit_app.py
```

## Déploiement Streamlit Community Cloud
1. Crée un repo GitHub **public** et pousse ces fichiers.
2. Va sur https://share.streamlit.io → **New app** → sélectionne le repo et `streamlit_app.py`.
3. Déploie → récupère l'URL `https://xxx.streamlit.app`.

## Intégration WordPress
Bloc HTML personnalisé :
```html
<iframe src="https://xxx.streamlit.app?embed=true" width="100%" height="900" style="border:none;"></iframe>
```

Plugin Advanced iFrame :
```
[advanced_iframe use_shortcode_attributes="true" src="https://xxx.streamlit.app?embed=true" width="100%" height="900"]
```

## Notes
- Les placeholders au format `« clef »` sont remplacés (même si découpés en runs).
- La première table du modèle est utilisée pour l'insertion du tableau (ou une nouvelle est créée).
- Le **Total TTC CHF** est ajouté en bas à droite.
