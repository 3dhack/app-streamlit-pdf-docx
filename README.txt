# PDF → DOCX (Streamlit) — fix19
**Objectif :** Ne plus téléverser le modèle Word à chaque fois.

## Ce qui change
- L'app charge automatiquement `template.docx` depuis le **répertoire de l'app**.
- Si `template.docx` est absent, un **uploader** s'affiche.
- Tu peux forcer un autre modèle via la case **“Remplacer le modèle embarqué”**.

## Comment utiliser
1. Place ton modèle Word dans le dépôt Streamlit sous le nom **`template.docx`** (c'est lui qui sera utilisé par défaut).
2. Déploie ou redéploie sur Streamlit Cloud.
3. Dans l'app, uploade uniquement le **PDF** — le modèle est pris automatiquement.
4. Si besoin d'un autre modèle ponctuellement, coche **“Remplacer le modèle embarqué”** et upload un `.docx`.

> L'ancrage du tableau est basé sur la ligne **“Cond. de paiement”** du modèle. Garde ce libellé pour le positionnement à 2 lignes en-dessous.

Le parsing/rendu reste identique au **fix18** (titre *Facture xxx* gras 12pt, en-tête coloré léger, pas de lignes verticales, Total TTC sous le tableau + 2 retours).