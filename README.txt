# PDF → DOCX (Streamlit) — fix11
Correction d'erreur: **InvalidXmlError: required <w:tblGrid> child element not present**

Changements:
- Le tableau Word est désormais créé via `doc.add_table()` puis **déplacé** sous « Cond. de paiement » (+2 lignes). Ainsi, la structure XML (tblGrid) est toujours valide.
- On conserve toutes les fonctionnalités du fix10: CF uppercase, « Notre référence » coupée avant « No TVA », date du jour (Europe/Zurich), Délai de réception max, reconstruction du tableau si besoin, arrêt au dernier item (10/20/30…), bordures, suppression des lignes Indice/Délai de réception et de la colonne TVA.

Déploiement:
1) Pousse ces fichiers dans ton repo GitHub
2) Streamlit Cloud redéploie
3) Teste la génération: l'erreur `tblGrid` ne doit plus apparaître
