# PDF → DOCX (Streamlit) — fix13
Modifs demandées :
1) **Bordures** visibles autour/dans le tableau (style Table Grid + oxml).
2) **Pos** alignée à gauche et **colonne étroite** (~0.5").
3) **Désignation** alignée à gauche et **plus large** (~3.5").
4) **Total TTC CHF** est affiché **sous le tableau**, aligné à droite.
5) **Deux** retours à la ligne sont ajoutés **après** la ligne Total (pas plus).

Toujours :
- CF en MAJUSCULE, « Notre référence » coupée avant TVA, date du jour (Europe/Zurich),
- Délai de réception = date max,
- Reconstruction si pas de tableau + arrêt 10/20/30…,
- Suppression lignes 'Indice :' / 'Délai de réception :' et colonne 'TVA'.

Déploiement
- Pousse sur GitHub → Streamlit Cloud redéploie.
