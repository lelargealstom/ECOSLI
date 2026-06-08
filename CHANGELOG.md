# Change Log — ECOSLI.xlam

Suivi des modifications apportées au complément Excel `ECOSLI.xlam` au fil des sessions de développement.

Format : `[Date] — Auteur — Périmètre`, puis liste des changements par module.

---

## 2026-05-18 — Session Claude Code (avec Thomas LE LARGE)

### Périmètre
Évolution **SIRUS** : ajout de l'affichage de la PBS Projet dans la `ListView_Generation` du UserForm ECOSLI à partir de la feuille `DATA_ASL` du classeur actif, plus calcul de la qualité des prix Spare / Repair.

### `G_SIRUS_PBS.bas` (modifié — modifs non commitées)
- Ajout du `Type DATA_ASL_Col` (40 colonnes) pour le cadrage de `DATA_ASL`.
- Variables globales : `DATA_ASL_type`, `DATA_ASL_HeaderRow`, `SIRUS_LookupDTR_Active`.
- Caches Dictionary privés pour les DTR BdD (Projet_ASL, Info_ATM, Taches_ATM, EMMST, Produit).
- `VerifierFeuilleDATA_ASL(wb, ws)` : vérifie présence de la feuille + au moins une ligne sous l'entête. Message clair à l'utilisateur sinon.
- `MapColonnesDATA_ASL(ws)` : cadrage `Find` xlWhole + fallback xlPart sur la ligne d'entête. Auto-détection de la ligne d'entête via le libellé ancre `DTR Reference (issue de PBS)`. Distinction colonnes critiques / optionnelles, MsgBox récapitulatif si critiques manquantes, `Debug.Print` pour les optionnelles.
- `SIRUS_RemplirListViewGeneration` : population de la `ListView_Generation` avec, par DTR : présence en BdD (Projet_ASL en OUI/NON, ATM-Info / ATM-Taches / EMMST / Produit en OK / vide) + indicateurs qualité prix Spare / Repair.
- `SIRUS_RafraichirListViewGeneration` : alias public à brancher sur un bouton de rafraîchissement.
- `ChargerSetDTR_BdD(table, colonne)` : charge la liste des DTR d'une table BdD dans un Dictionary via `GetRows` (plus robuste qu'une boucle MoveNext sur source Excel ODBC).
- `EvaluerQualitePrix(vManuel, vFournisseur, vCatalogue, [seuil=0.25])` : règle 0/3, 1/3, 2/3, 3/3 avec calcul d'écart à la médiane.
- Helpers privés : `TrouverColonne`, `DetecterLigneEntete`, `SetContient`, `LireCellule`, `EstNumeriqueConnu`, `MedianeConnus`.
- Procédures de test exécutables : `Test_MapColonnesDATA_ASL`, `Test_EvaluerQualitePrix`.

### `G_SIRUS_UserForm.bas` (modifié)
- `SIRUS_ConfigListViewGeneration` : ajout d'une colonne `Projet_ASL` en position 2 (largeur 70). La ListView passe à 12 colonnes : DTR, Projet_ASL, ATM-Info, ATM-Taches, EMMST, Produit, Qualité Prix Spare, Source Ecart Spare, Source Manque Spare, Qualité Prix Repair, Source Ecart Repair, Source Manque Repair.

### `UserForm.frm` (modifié)
- `UserForm_Initialize` : activation du lookup BdD via `G_SIRUS_PBS.SIRUS_LookupDTR_Active = True` avant l'appel à `SIRUS_RemplirListViewGeneration`.

### Documentation
- Création de `CHANGELOG.md` (ce fichier).
- Création de `ECOSLI_SYNTHESE.md` (présentation générale du complément).

### Reste à faire
- Validation par Thomas des 40 libellés `DATA_ASL` actuellement cadrés.
- Test `Test_MapColonnesDATA_ASL` sur un classeur projet réel.
- Branchement d'un bouton `CommandButton_RafraichirPBS` dans la page Generation, sur `G_SIRUS_PBS.SIRUS_RafraichirListViewGeneration`.
- Recette sur plusieurs classeurs aux structures `DATA_ASL` légèrement différentes.

---

## a9e737f — 1st send (initial commit)
Commit initial du complément avec tous les modules historiques :
- `A_*` : événements / mise à jour
- `B_ATM_*` : génération livrables ATM V1 / V2
- `C_EMMST_*` : génération EMMST
- `D_Providor_*` : génération Providor (BreakDownStructure, Finalisation, Maintenance, TNR, Utilitaires)
- `E_Extraction_*` : extraction ATM / EMMST
- `F_FonctionsGlobales` : utilitaires globaux
- `G_SIRUS_*` : pipeline SIRUS (Saisie, Analyse, ATMAP, PBS, UserForm) — modules en cours de développement
- `O_Projet`, `P_Nettoyage`, `Q_TestComparaison`, `U_Mapping`, `V_Nom_ADODB_Type`, `W_AppelUsf`, `X_Test_Appel_BDD`, `Y_Cadrage_SLI`, `Ybis_CadrageTemplate`, `Z_*` : modules transverses
- `clsAppEvents`, `EMMST_Ranges` : classes
- `UserForm.frm`, `Barre_Progression.frm` : formulaires
