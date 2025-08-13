# PDF → DOCX — fix21
- **Reconstruction des articles**: commence sur Pos multiples de 10 (10/20/30/...) et **s'arrête au 'Total CHF' de la ligne**.
  -> Empêche que la dernière ligne d'une page capture du texte de la page suivante.
- Multi-pages: on fusionne les lignes Pos depuis toutes les pages si des tables sont détectées; sinon on reconstruit via le texte.
- Total affiché sous le tableau = montant de **« Total CHF »**.
- Conserve: titre **Facture xxx** (gras 12pt), en-tête coloré léger, pas de lignes verticales intérieures, Total sous tableau + 2 retours, insertion 2 lignes sous « Cond. de paiement ».

Fix22:
- Corrige l'arrêt prématuré après la Pos 10 : les lignes 'Tarif douanier', 'Pays d'origine', 'Indice :', 'Délai de réception :' sont ignorées (et ne stoppent plus l'analyse).
- S'arrête uniquement sur 'Récapitulation / Montant total / Total TTC / Code TVA / Taux'.
