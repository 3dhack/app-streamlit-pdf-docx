# PDF → DOCX (Streamlit) — fix12
Modifs demandées :
1) Bordures visibles autour et à l’intérieur du tableau.
2) Colonne **Pos** : alignée à gauche et largeur réduite (~1.2 cm).
3) Colonne **Désignation** : alignée à gauche et largeur accrue (~9 cm).
4) Affichage du champ **Total TTC CHF** sous forme de ligne **en bas à droite** sous le tableau.
5) Ajout de **2 retours à la ligne** après cette ligne Total.

Toujours inclus : CF en majuscules, “Notre référence” tronqué avant “No TVA”, date du jour (Europe/Zurich), “Délai de réception” = date max, reconstruction du tableau si besoin, arrêt sur positions 10/20/30…, suppression “Indice :”/“Délai de réception :” et colonne TVA.

Déploiement : push → Streamlit Cloud redéploie → tester → intégrer WordPress.
