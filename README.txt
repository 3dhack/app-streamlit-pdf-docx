# PDF → DOCX — fix27
- Identique à fix25 (DOCX + PDF) mais **un seul retour à la ligne** après la ligne de total intégrée au tableau.
- Total intégré au tableau : dernière ligne, cellules fusionnées (libellé), montant dans la dernière colonne, aligné à droite, **gras**, **doublement souligné**, **bordure supérieure double**.
- Parsing multi-pages robuste (Pos multiples de 10 ; stop sur Récapitulation/Montant Total/Total TTC/Code TVA/Taux ; ignore lignes meta).
