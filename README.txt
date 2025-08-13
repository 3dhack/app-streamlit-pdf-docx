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
