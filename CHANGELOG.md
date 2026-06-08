# Change Log — ECOSLI.xlam

Suivi des modifications apportées au complément Excel `ECOSLI.xlam` au fil des sessions de développement.

Format : `[Date] — Auteur — Périmètre`, puis liste des changements par module.

---

## 2026-06-08 — Session Claude Code (avec Thomas LE LARGE)

### Périmètre
Cadrage v0.2 de **SIRUS** (renommage du module **SIRIUS → SIRUS**), corrections sur l'existant, et **Phase 0** de la roadmap SIRUS : socle de configuration piloté par feuille.

### Corrections sur l'existant
- `C_EMMST_Main.bas` (`CopierTemplateEMMST`) : correction d'un `Else` collé en bout de ligne après le `templatePath` FR — le bloc `If langue = "FR" Then ... Else ... End If` ne compilait plus.
- `C_EMMST_Main.bas` / `Z_Variables.bas` : mise à jour des chemins SharePoint vers la nouvelle organisation **bibliothèques DIS** (`DIS Service VBN Pole ISR - Qualité`, `DIS Service VBN Pole ISR - Outils Métier ISR`, `DIS Service VBN Pole ISR - Operations`). Encodage UTF-8 propre rétabli.
- `Z_Variables.bas` (`Type ECOSLI_Col`) : renommage `Designation_BDD` → `Designation_BdD` et `MTBF_gen` → `MTBF_Gen` (cohérence de casse).
- `.gitignore` créé : exclusion d'`ECOSLI.xlsm` (template projet, 1.4 MB binaire), des fichiers temporaires Office (`~$*`, `*.tmp`) et du dossier local Claude (`.claude/`).

### Renommage SIRIUS → SIRUS
Le module est renommé **SIRUS** (cadrage v0.2). 105 occurrences remplacées sur 7 fichiers ; aucun impact sur les modules historiques `B_*` / `C_*` / `D_*` / `E_*` / `Z_*`.
- Renommage des 4 fichiers : `G_SIRIUS_Analyse.bas` → `G_SIRUS_Analyse.bas`, `G_SIRIUS_ATMAP.bas` → `G_SIRUS_ATMAP.bas`, `G_SIRIUS_PBS.bas` → `G_SIRUS_PBS.bas`, `G_SIRIUS_UserForm.bas` → `G_SIRUS_UserForm.bas` (+ `Attribute VB_Name`).
- Renommage de toutes les références internes : `SIRIUS_*` → `SIRUS_*` (fonctions, variables, types, commentaires) dans les 4 modules ci-dessus + `UserForm.frm` + `CHANGELOG.md` + `ECOSLI_SYNTHESE.md`.

### Phase 0 SIRUS — `G_SIRUS_Config.bas` (nouveau module)
Socle de configuration piloté par une feuille `SIRUS_Config` (veryHidden, ListObject `tblSIRUSConfig`) — cf. brief v0.2 section 8.
- `Type Champ_Config` : 13 colonnes (Table, Champ, ColonneBdD, Ordre, TypeChamp, ListeSource, Obligatoire, ConditionObligatoire, Source, SourceHeritage, PoidsMeQ, RegleAlerte, Commentaire).
- Constantes : `SIRUS_CONFIG_SHEET`, `SIRUS_CONFIG_LISTOBJ`, et noms de table `SIRUS_TBL_PRODUIT` / `INFOATM` / `TACHESATM` / `EMMST` / `PROJETSLI` / `PROJETASL`.
- `SIRUS_InitConfigSheet([bForcer])` : crée la feuille + le `ListObject` dans `ThisWorkbook` (le `.xlam`) si absente. Idempotente, ré-init possible via `bForcer:=True`.
- `PeuplerProduit(lo)` (privée) : peuple le config pour la table **Produit** (pilote Phase 0) — 14 champs conformes à la section 6.1 du brief (DTR, Désignation manuelle, Node Type, Fournisseur, Réf. fournisseur, N° plan, MoQ Spare, MTBF intrinsèque/Motif NA conditionnels, Prix/Date validité Spare/Repair conditionnels selon `NodeType`, Désignation BdD en lecture seule).
- `SIRUS_LireConfig(table)` : renvoie une `Collection` de `Champ_Config` triée par `Ordre` croissant. Lecture diagnostique non bloquante (Debug.Print en cas d'absence).
- Helpers privés : `AjouterChamp`, `NzLng`, `NzDbl`.
- Procédures de test exécutables : `Test_InitConfigSheet`, `Test_LireConfig`.

### Reste à faire
- Importer `G_SIRUS_Config.bas` dans le `.xlam` via VBE, exécuter `Test_InitConfigSheet` puis `Test_LireConfig` pour valider le pilote Produit.
- Peupler le config pour `InfoATM` et `TachesATM` (suite Phase 0, brief sections 6.2).
- **Phase 1** : `SIRUS_GenererFeuilleAnalyse(table, operation, dtrs)` — générateur de feuille (création vierge ou pré-remplie via ADODB pour Modifier).
- **Phase 2** : généraliser `SIRUS_AnalyserProduit` en `SIRUS_AnalyserFeuille(table)` piloté par config.

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
