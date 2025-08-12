# PDF → DOCX (Streamlit) — fix17
Nouveautés :
1) **Thème de tableau** appliqué + **première colonne colorée** (teinte légère). Bordures toujours forcées.
2) Titre en haut : **« Facture xxx »** où **xxx** = nombre **après `CF-25-`** dans « Commande fournisseur » (ex. `CF-25-05259` → `Facture 05259`).
3) Champ **« Livré le »** alimenté via « Délai de livraison » (date max).

Toujours :
- Alignements/largeurs (Pos, Référence, Qté, Désignation), Total TTC sous le tableau (+2 lignes), insertion 2 lignes sous « Cond. de paiement ».
- CF en MAJ, Notre référence tronquée avant TVA, date du jour, etc.
