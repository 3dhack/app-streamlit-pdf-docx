# PDF → DOCX — fix24b
- Corrige un possible souci d'indentation/patch précédent.
- Intègre le **Total TTC CHF** *dans* le tableau : ligne finale fusionnée, **alignée à droite**, **gras**, **double souligné**, **bordure supérieure double**.
- Conserve les améliorations parsing (multi-pages, Pos 10/20/30… robustes, Total sous tableau basé sur « Total CHF », etc.).


Fix25:
- Ajoute l'**export PDF** en plus du DOCX.
- Conversion multi-stratégies : d'abord **LibreOffice/soffice**, puis **unoconv**, puis **docx2pdf** (plutôt Windows/Mac).
- Si aucune stratégie n'est dispo sur l'hébergement, l'app affiche une info et propose au moins le DOCX.
- Pour Streamlit Cloud (Linux), il est probable que ni soffice ni unoconv ne soient présents : pour l'export PDF garanti, déploie sur un **VPS** et installe `libreoffice` (ou `unoconv`).
