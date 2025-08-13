# PDF → DOCX — fix27
- Identique à fix25 (DOCX + PDF) mais **un seul retour à la ligne** après la ligne de total intégrée au tableau.
- Total intégré au tableau : dernière ligne, cellules fusionnées (libellé), montant dans la dernière colonne, aligné à droite, **gras**, **doublement souligné**, **bordure supérieure double**.
- Parsing multi-pages robuste (Pos multiples de 10 ; stop sur Récapitulation/Montant Total/Total TTC/Code TVA/Taux ; ignore lignes meta).


Fix28:
- **Supprime l'export PDF** et le code associé (plus de docx2pdf/soffice/unoconv).
- **Nettoyage des paragraphes vides après le tableau** : on garde **exactement 1 seul** retour à la ligne, les autres sont supprimés automatiquement.


Fix29:
- **Centrage** des colonnes **Unité**, **Prix unit.**, **Px u. Net** (en-tête et valeurs).


Fix30:
- **Suppression du gras** sur toutes les **lignes de données** du tableau (bas du tableau), toutes colonnes.
- En-têtes **restent en gras** ; ligne de **Total** reste **gras + souligné double**.


Fix31:
- Ajoute un bouton **"🔄 Réinitialiser"** qui efface l'état (fields/items_df/doc temp) et relance l'app proprement.


Fix32:
- Remplace `st.experimental_rerun()` par `st.rerun()` (évite l'`AttributeError` avec les versions récentes de Streamlit).


Fix33:
- Le bouton **"Réinitialiser"** réinitialise désormais **visuellement** les uploaders (PDF et DOCX) :
  - utilisation de clés dynamiques (`pdf_uploader_<n>`, `docx_uploader_<n>`),
  - incrément des clés à chaque reset pour vider/rafraîchir les widgets.


Fix34:
- Le bouton **"🔁 Réanalyser"** est désormais **affiché uniquement** lorsqu'un **PDF est chargé** (sinon il est caché).


Fix35:
- Le bouton **"🧾 Générer le DOCX"** est **caché** tant que l'analyse n'est pas prête (PDF + modèle disponibles).
- Un message d'info s'affiche à la place pour guider l'utilisateur.


Fix36:
- Masque **toutes les lignes d'information** tant qu'aucun PDF n'est uploadé :
  - plus d'"Aperçu du tableau" ni "Le tableau sera reconstruit..." avant upload,
  - le message d'aide pour la génération n'apparaît que si un PDF est présent.


Fix37:
- **Centrage vertical** du texte dans **toutes les cellules** du tableau (en-tête, lignes de données et total).


Fix38:
- **Centrage vertical renforcé** (ajout d'un setter XML `w:vAlign` = `center`) sur toutes les cellules, y compris après fusion.
- **Suppression du gras** sur la **ligne de total** (conservée en **double souligné** uniquement), pour éviter la “ligne en gras” en bas du tableau.


Fix39:
- **Retire toute double ligne** au niveau global du tableau (les bordures standard restent simples).
- **Conserve une ligne double uniquement sous la ligne de total** (bordure supérieure double du dernier rang).
- **Remet le texte en gras** sur la ligne **Total TTC CHF** (libellé et montant), en plus du **double souligné**.


Fix41:
- **Corrige le nettoyage au bon endroit** : supprime la **bordure basse de la dernière ligne de données** (celle avec les Pos), pas celle du total.
- Conserve la **double ligne/underline** sous **Total TTC CHF** et son texte **en gras**.
- Garde **1 seul retour à la ligne** après le tableau.


Fix42:
- Améliore l'affichage **"Champs détectés"** :
  - tableau propre (2 colonnes **Champ / Valeur**), ordre des infos pertinent,
  - expander pour afficher le **JSON brut** si besoin,
  - note informative si le **Total TTC CHF** utilisé diffère du **Montant Total TTC CHF (PDF)**.
