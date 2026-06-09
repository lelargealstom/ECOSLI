# Change Log — ECOSLI.xlam

Suivi des modifications apportées au complément Excel `ECOSLI.xlam` au fil des sessions de développement.

Format : `[Date] — Auteur — Périmètre`, puis liste des changements par module.

---

## 2026-06-09 — Session Claude Code (avec Thomas LE LARGE)

### Périmètre
Phase 0 ter de SIRUS : décisions structurantes sur l'évaluation des règles et la gestion des listes, création du module `G_SIRUS_Listes`, et renommage de la colonne `ConditionObligatoire` → `ConditionDescription` (devient descriptive, plus évaluée).

### Décisions verrouillées
- **Règles d'analyse en dur (« D1 »)** : pas de mini-langage évaluable. Les conditions stockées dans le config (`NodeType IN (LRU,SRU)`, `Description CONTAINS 'Torque Tighten'`, etc.) sont désormais **purement descriptives**. Les vraies règles vivront dans `G_SIRUS_Regles` (à créer Phase 2), `Select Case` par (Table, ColonneBdD) qui prend la ligne courante (Dictionary) et renvoie `Boolean`.
- **Listes en code** (pas en feuille `Listes`). Module `G_SIRUS_Listes` avec `SIRUS_GetListe(cle)` qui renvoie un tableau 1D.
- **ATMAP fonctionnel** : Thomas confirme que `G_SIRUS_ATMAP` est opérationnel (le brief disait « ~80 % », c'était sous-estimé). Phase 5 = figer le contrat PJ + brancher l'appel.

### `G_SIRUS_Listes.bas` (nouveau module)
- Constantes des clés : `SIRUS_LIST_NODETYPE`, `HARDSOFT`, `LANGUE`, `BOOLEEN`, `TYPETRAIN`, `DEVISE`, `NIVMAINTENANCE`.
- `SIRUS_GetListe(cle) As Variant` : tableau 1D ; clé inconnue → `Array()` vide + `Debug.Print`.
- `SIRUS_ListeEnChaine(cle, [sep=","])` : version chaîne pour `Validation.Add Formula1:=...` (Excel limite à 255 caractères).
- 7 listes peuplées avec les valeurs validées par Thomas :
  - `NodeType` : LRU, SRU, LRUNR, SRUNR, KIT, VISSERIE, ASSEMBLY (du brief).
  - `HardSoft` : HARD, SOFT.
  - `Langue` : FR, EN.
  - `Booleen` : Oui, Non.
  - `TypeTrain` : CORADIA, AVELIA, METROPOLIS, XTRAPOLIS, TRAX, INNOVIA, CITADIS, ADESSIA, Voitures, SIG.
  - `Devise` : EUR, USD, CAD, JPY, CHF, GBP.
  - `NiveauMaintenance` : 3, 3/3+, 3+.
- `Test_GetListe` : trace les 7 listes dans la fenêtre Immediate.

### `SIRUS_Champ_Config.cls` / `G_SIRUS_Config.bas`
- Renommage `ConditionObligatoire` → `ConditionDescription` (propriété de la classe + colonne du config + paramètre `sConditionObligatoire` → `sConditionDescription` de `AjouterChamp` + accès dans `SIRUS_LireConfig` et `Test_LireConfig`). Commentaire reformulé pour préciser que c'est descriptif (pas évalué).
- `PeuplerProjetASL` — champ Testabilité : passé de `TypeChamp="Liste", ListeSource="Testabilite"` à `TypeChamp="Texte", ListeSource=""` (n'est pas une liste, validation par Thomas).

### Reste à faire
- Réimporter `G_SIRUS_Config.bas` + `SIRUS_Champ_Config.cls` + nouveau `G_SIRUS_Listes.bas` dans le `.xlam`. Exécuter `Test_InitConfigSheet`, `Test_LireConfig_Toutes`, `Test_GetListe`.
- **Phase 1** : `SIRUS_GenererFeuilleAnalyse(table, operation, dtrs)` — générateur de feuille (lit le config, pose les `Validation.Add` via `SIRUS_ListeEnChaine`, grise les champs `Auto`/`Herite`, pré-remplit via ADODB pour `Modifier`).
- **Phase 2** : créer `G_SIRUS_Regles.bas` (`SIRUS_RegleObligatoire(table, colonneBdD, ligne) As Boolean`, `SIRUS_RegleAlerte(...)`) + généraliser `SIRUS_AnalyserProduit` en `SIRUS_AnalyserFeuille(table)`.

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

### Correction Phase 0 — passage du `Type` en classe
VBA interdit le stockage d'un `Type` (UDT) défini dans un module standard dans une `Collection` (erreur *« Seuls les types définis par l'utilisateur et qui sont définis dans les modules d'objets publics peuvent être passés à des fonctions à liaison tardive »*). Le `Type Champ_Config` est donc remplacé par une **classe** `SIRUS_Champ_Config.cls` (13 champs `Public`, pattern minimal aligné sur `EMMST_Ranges.cls`).
- Nouveau fichier `SIRUS_Champ_Config.cls`.
- `G_SIRUS_Config.bas` : suppression du `Type Champ_Config`, `SIRUS_LireConfig` instancie `New SIRUS_Champ_Config` à chaque ligne (sinon toutes les entrées de la `Collection` pointent vers le même objet), `Test_LireConfig` retrouve un `For Each def In col` natif.

### Phase 0 bis — peuplement des 5 autres tables
Extension de `G_SIRUS_Config.bas` au reste du brief v0.2 sections 6.2 → 6.5 (les modules restent inchangés, seules les procs internes s'ajoutent).
- `PeuplerInfoATM` (section 6.2) — 8 champs : DTR + HARD/SOFT obligatoires ; Descriptif FR/EN, Plan (hérité `Produit.NumeroPlan`), Réparabilité, Poids (hérité `EMMST.Poids` + alerte si absent), Temps générique en lecture seule.
- `PeuplerTachesATM` (section 6.2) — 7 champs : DTR, Rang, Langue, HARD/SOFT, Description (conditionnel rang 3.*/4.*), Critère Nm (conditionnel `Description CONTAINS 'Torque Tighten'`), Banc test (conditionnel `Rang IN (2.1,5.1) AND NodeType>=SRU`).
- `PeuplerEMMST` (section 6.3) — 24 champs principaux : identification, dimensions bare/packaged, conditionnement (Péremption + date conditionnelle), stockage (intérieur/extérieur/hors poussière + commentaire), transport, marquage/manutention, commentaire EN. **À compléter Phase 1+** avec la grille EMMST complète (le brief dit « tous champs attendus » sans lister exhaustivement).
- `PeuplerProjetSLI` (section 6.4) — 15 champs : Nom projet, Type de train, Durée de circulation, Nombre de trains, Profil de mission (Servicing/day, PH/year, RH/year, Km/year), Durée garantie, Devise, Man hour cost, Niveau de maintenance, Réf PBS/EAP, Réf RAM report, LCC scope total. Formation/BTE/MCO/OBSO exclues (PATH).
- `PeuplerProjetASL` (section 6.5) — 11 champs : Nom projet (avec règle de cohérence `NomProjet IN ProjetSLI.NomProjet`), DTR, Désignation PBS, Qté, hérités Produit grisés (Node Type, MTBF intrinsèque, DVL générique), saisies projet (MTBF fonctionnel, Lead time Spare/Repair, Testabilité).
- `SIRUS_InitConfigSheet` enchaîne les 6 `Peupler*`.
- Nouveau `Test_LireConfig_Toutes` : itère sur les 6 tables et trace, par table, le compte total + répartition (Obligatoire Oui/Conditionnel/Non, Source Saisie/Auto/Hérite).
- ~79 champs au total dans le config après init.

### Reste à faire
- Réimporter `G_SIRUS_Config.bas` dans le `.xlam`, exécuter `Test_InitConfigSheet` (qui forcera la recréation) puis `Test_LireConfig_Toutes` pour valider les 6 tables.
- Validation par Thomas des libellés, `ColonneBdD`, listes (`HardSoft`, `Langue`, `TypeTrain`, `Devise`, `NiveauMaintenance`, `Testabilite`, `Booleen`) et expressions `ConditionObligatoire` / `RegleAlerte` (syntaxe à figer avant le moteur d'analyse Phase 2).
- Compléter EMMST avec la grille complète (sections, libellés EN, booléens fins) si la couverture Phase 0 est jugée insuffisante.
- Créer la feuille `Listes` (sources des listes déroulantes) — Phase 0 ter.
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
