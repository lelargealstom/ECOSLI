# ECOSLI.xlam — Synthèse du complément

## Identité

`ECOSLI.xlam` est un complément Excel VBA (Add-in) développé pour **Alstom Services**, site de **Villeurbanne**, à destination de l'équipe **LGM / ISR**. Il automatise la production des livrables SLI (Support Logistique Intégré) à partir d'une base de données Excel partagée, et propose un pipeline de demande d'import des données dans cette BdD.

**Auteur principal :** Thomas LE LARGE (LGM / Alstom).

## Fonctionnement général

1. L'utilisateur ouvre un classeur projet Excel qui contient la feuille `DATA_ASL` (la PBS — *Product Breakdown Structure* — du projet).
2. Il appuie sur **Ctrl + W** depuis ce classeur.
3. Le UserForm `ECOSLI` s'affiche en `vbModeless` (cf. `W_AppelUsf.Appel_Usf`).
4. Il peut alors :
   - **Générer des livrables SLI** à partir de la PBS du classeur actif : ATM V1, ATM V2, EMMST, Providor.
   - **Faire une demande d'import / modification / suppression** de DTR dans la BdD via le pipeline **SIRUS**.
   - **Visualiser** sa PBS et l'état de cohérence avec la BdD dans une `ListView_Generation` (présence dans Projet_ASL, ATM-Info, ATM-Taches, EMMST, Produit + qualité des prix Spare / Repair).

L'add-in est installé via `A_1_ModEvents.Install`, qui le copie dans `Application.UserLibraryPath` et l'active.

## Architecture des modules

Les préfixes des modules reflètent une organisation par domaine :

| Préfixe | Domaine | Modules clés |
|---|---|---|
| `A_*` | Cycle de vie de l'add-in | `A_1_ModEvents` (Auto_Open / Auto_Close, Install), `A_2_ModUpdate` (check de version) |
| `B_ATM_*` | Génération livrable ATM | `B_ATM_Commun`, `B_ATM_Main`, `B_ATM_V1`, `B_ATM_V2` |
| `C_EMMST_*` | Génération livrable EMMST | `C_EMMST_Main`, `C_EMMST_Remplissage` |
| `D_Providor_*` | Génération livrable Providor (PBS provisionnée) | `D_Providor_Main`, `D_Providor_BreakDownStructure`, `D_Providor_Finalisation`, `D_Providor_Maintenance`, `D_Providor_TNR`, `D_Providor_Utilitaires` |
| `E_Extraction_*` | Extraction des données existantes | `E_Extraction_Main`, `E_Extraction_ATM`, `E_Extraction_EMMST` |
| `F_*` | Fonctions globales | `F_FonctionsGlobales` |
| `G_SIRUS_*` | Pipeline SIRUS (demande d'import en BdD) | `G_SIRUS_Saisie` *(à créer si besoin)*, `G_SIRUS_Analyse`, `G_SIRUS_ATMAP`, `G_SIRUS_PBS`, `G_SIRUS_UserForm` |
| `O_*` `P_*` `Q_*` | Métier transverse | `O_Projet`, `P_Nettoyage`, `Q_TestComparaison` |
| `U_*` `V_*` | Mapping / typage | `U_Mapping`, `V_Nom_ADODB_Type` |
| `W_*` `X_*` `Y_*` | UserForm + tests + cadrage | `W_AppelUsf` (Ctrl+W → Show UserForm), `X_Test_Appel_BDD`, `Y_Cadrage_SLI`, `Ybis_CadrageTemplate` |
| `Z_*` | Variables globales & extraction de code | `Z_Variables`, `Z_1_ExtracCode` |
| Classes | | `clsAppEvents` (events Excel), `EMMST_Ranges` |
| Forms | | `UserForm` (interface principale), `Barre_Progression` |

## Connexion à la BdD SLI

La BdD SLI est elle-même un classeur Excel partagé, consulté par ADO :
- Variables globales : `CONNEXION_EXCEL_BDD_PROD`, `CONNEXION_EXCEL_BDD_TRAVAIL` (cf. `Z_Variables` / `V_Nom_ADODB_Type`).
- Procédures pivot :
  - `Open_CONN` — ouvre la connexion
  - `open_REQSEL(STRSEL)` — exécute la requête SELECT, met le RecordSet dans `REQSEL`
  - `Close_CONN` — ferme la connexion
- Tables BdD adressées en SQL Excel `[NomTable$]` : `Projet_ASL$`, `Info_ATM$`, `Taches_ATM$`, `EMMST$`, `Produit$`.

## Conventions techniques

| Convention | Détail |
|---|---|
| Source PBS | `ActiveWorkbook` (pas `ThisWorkbook`) — la PBS est dans le classeur projet de l'utilisateur, pas dans le `.xlam`. |
| Cadrage colonnes | Toujours via `Range.Find(What:=..., LookIn:=xlValues, LookAt:=xlWhole)` puis fallback `xlPart`. Jamais de magic number. Stockage dans un `Type` dédié (cf. `DATA_ASL_Col` dans `G_SIRUS_PBS`). |
| Détection ligne d'entête | `DetecterLigneEntete(ws, libelleAncre)` : recherche d'un libellé ancre sur toute la feuille, robuste aux lignes de titre / filtres au-dessus de l'entête. |
| Gestion d'erreur | **Pas de handler défensif** (règle inversée 2026-07-20 par Thomas pour faciliter le debug avec les ingénieurs). Ne pas écrire `On Error GoTo <label>` + bloc `MsgBox Err.Description` : l'objectif est que VBA s'arrête sur la ligne exacte qui plante (surlignage jaune) pour inspecter les variables. À conserver : (a) `On Error Resume Next` **fonctionnels** (test d'existence de feuille / clé Dictionary, patterns UI tolerants, `CopyFile` idempotent) avec `On Error GoTo 0` juste après ; (b) labels `Sortie:` utilisés en early-exit (ex. `If REQSEL Is Nothing Then GoTo Sortie` → `Close_CONN`) — le label n'est plus un handler d'erreur, juste un point de nettoyage normal ; (c) `Err.Raise` explicite en pré-condition (`If Base Is Nothing Then Err.Raise ...`). |
| Affichage tabulaire | ListView ActiveX (`MSComctlLib.ListView`, *Microsoft Windows Common Controls*) en mode `lvwReport`, `FullRowSelect = True`, `Gridlines = True`, `LabelEdit = lvwManual`. Pas de ListBox pour les tableaux. |
| Pas de référence VBA nouvelle | Toutes les références sont déjà déclarées dans le projet ; on n'en ajoute pas. |
| Commentaires | En français, sans accents pour rester compatible cp1252 / UTF-8 à l'import VBA. Les libellés avec accents utilisent `ChrW(...)` côté code. |
| Encodage des libellés | Convention : `"DESIGNATION (issue de PBS)"` reste en clair, mais `"r" & ChrW(233) & "f du plan d'ensemble"` pour les accents — concaténation avec `ChrW` pour garantir la compatibilité multi-éditeurs. |

## Pipeline SIRUS (focus du chantier en cours)

SIRUS est le sous-système qui permet à l'utilisateur de pousser des demandes d'import / modification / suppression de données dans la BdD SLI depuis le UserForm ECOSLI.

### Flux

1. **Saisie** (page principale du UserForm) — l'utilisateur sélectionne un type de demande (Création / Modification / Suppression) via `Type_Import_Tab` et remplit les champs Produit ou ATM dans `ListView_Produit` / `ListView_Taches`.
2. **Analyse** (`G_SIRUS_Analyse.SIRUS_AnalyserProduit`) — contrôles bloquants et alertants :
   - Champs obligatoires
   - Doublon DTR interne à la saisie
   - DTR déjà présent en BdD (warn pour distinguer Création / Modification)
   - Types de données (MTBF décimal > 0, DVL ≤ 50 ans)
   - Calcul du taux KO ; bloquant si > 20 %
3. **Page Generation** (`G_SIRUS_PBS.SIRUS_RemplirListViewGeneration`) — vue PBS du classeur actif avec, par DTR :
   - Présence en BdD (Projet_ASL / Info_ATM / Taches_ATM / EMMST / Produit)
   - Qualité Spare / Repair calculée à partir de 3 sources de prix : Manuel, Fournisseur Alstom, Catalogue ISP. Règles dans `G_SIRUS_PBS.EvaluerQualitePrix` :
     - 0–1 source connue → `NOK - Manque`
     - 2 sources, écart > 25 % à la médiane → `NOK - Manque & Ecart`
     - 2 sources, écart ≤ 25 % → `NOK - Manque`
     - 3 sources, écart > 25 % → `NOK - Ecart`
     - 3 sources, écart ≤ 25 % → `OK`

### Cadrage `DATA_ASL`

Le module `G_SIRUS_PBS` cadre la feuille `DATA_ASL` du classeur actif via `MapColonnesDATA_ASL` :
- Auto-détection de la ligne d'entête (libellé ancre : `DTR Reference (issue de PBS)`)
- 40 colonnes mappées dans `DATA_ASL_type` (Identification, RAM, DVL/Préventif, Coûts/Prix)
- Colonnes critiques (7) : `DTR Reference (issue de PBS)`, les 4 prix Spare / Repair (Manuel + Alstom FR + Catalogue ISP) + `ISP TIS FRANCE (FCS)`.
- Si une critique est introuvable → MsgBox récapitulatif + abandon du remplissage.
- Si une optionnelle est introuvable → `Debug.Print` seulement, le remplissage continue.

## Pipeline SIRUS complet (mis à jour 2026-06-25)

Le pipeline SIRUS est désormais **end-to-end** : Génération → Saisie → Analyse → Envoi ATMAP. Recette Sprint 1 validée pour Produit Créer + ProjetASL Modifier (Extraction ATM reste à tester).

### Modules SIRUS

| Module | Rôle |
|---|---|
| `G_SIRUS_Config.bas` + `SIRUS_Champ_Config.cls` | Socle config : feuille `SIRUS_Config` (veryHidden) avec ~158 champs sur 6 tables (Produit, InfoATM, TachesATM, EMMST, ProjetSLI, ProjetASL). |
| `G_SIRUS_Listes.bas` | 8 listes en code (NodeType, HardSoft, Langue, Booleen, TypeTrain, Devise, NiveauMaintenance, MotifNA). |
| `G_SIRUS_Generation.bas` | Phase 1 : générateur de feuille `SIRUS_<Table>_<Op>` avec layout Statut/Champs/Commentaires, listes déroulantes, grisage Auto/Herite, verrouillage cellules + protection feuille (Phase 3.1), pré-remplissage BdD pour Modifier, freeze panes. |
| `G_SIRUS_Regles.bas` | Phase 2 : `SIRUS_RegleObligatoire` + `SIRUS_RegleAlerte` (Select Case par Table\|ColonneBdD). 10 règles MVP (Produit/TachesATM/ProjetASL). |
| `G_SIRUS_Analyse.bas` | Phase 2 : `SIRUS_AnalyserFeuille(table, op)` parcourt la feuille, applique les règles, écrit Statut col 1 + commentaire de cellule. Renvoie taux MeQ. |
| `G_SIRUS_Extraction.bas` | Phase 4 entrée A : `SIRUS_DeverserExtractionDansSIRUS` orchestre Main_Extraction_Params + mapping colonnes brutes → SIRUS via ligne 2 (ColonneBdD). |
| `E_Extraction_ATM_Tolerant.bas` | Extraction ATM par détection libellés (cartouche + headers tableau) avec fallback positions. Logique métier conservée (DTR=Left(13), HARD/SOFT=Right(4), prefix REPLACE EN, Version="A1"). |
| `G_SIRUS_ATMAP.bas` | Phase 5 : API REST SharePoint pour ATMAP_TICKETS. Split numérotation : `SIRUS_NumeroVBNSuivant` (`VBN_NN` auto-incrémenté → `Numero_ticket`) + `SIRUS_NumeroTicketSuivant` (`IMP-YYYY-NNN` → `Numero_Import`). `SIRUS_DateEstimeeISO(priorite)` + helper `SIRUS_JoursTraitementSelonPriorite` (1/3/5 j). `SIRUS_ConstruirePJ`, `SIRUS_AttacherFichierAuTicket`. |
| `G_SIRUS_UserForm.bas` | UI Workflow 3 zones (Génération / Analyse / Envoi). Frames Analyse + Envoi visibles ensemble dès la Génération (décision 2026-06-11 : gating MeQ retiré pour phase de test). Handlers `SIRUS_UI_*`. Garde-fou envoi : `SIRUS_ProjetsASL_Inconnus` bloque si projet ProjetASL absent de ProjetSLI. |
| `G_SIRUS_PBS.bas` | Cadrage `DATA_ASL` + ListView Generation avec **% MeQ par table** (Produit/InfoATM/TachesATM/EMMST) calculé via `SIRUS_LireConfig` + `SIRUS_RegleObligatoire`. Col ProjetASL = **nombre de projets distincts** + tooltip de ligne avec liste des projets. SELECT ciblé sur colonnes Saisie + cache config 5× (au lieu 5×N). Helper `NzSafeStr` null-safe pour ADO. |

### Workflow utilisateur

```
1. UserForm → Onglet SIRUS
2. Choisir Op (Créer/Modifier/Extraction) + Table (combo dynamique selon Op)
   - Modifier ProjetSLI/ProjetASL → Combo Projet auto-alimenté depuis BdD
   - Extraction ATM → Version V1/V2 + Fichier livrable via Parcourir
3. Cliquer "Générer" → feuille SIRUS_<Table>_<Op> créée dans le classeur projet
   → Frame Analyse devient visible
4. Saisir / corriger les valeurs dans la feuille (champs marqués * rouges = obligatoires)
5. Cliquer "Analyser" → Statut OK/NOK + commentaire détaillé sur cellule Statut
   → ListView NOK avec lignes rouges (clé = DTR, ou DTR pour ProjetASL)
   → Frame Envoi devient visible si MeQ ≥ 70%, pré-remplissage auto
6. (Modifier) Sujet/Description/Priorité/Attentes pré-remplis → ajuster si besoin
7. Cliquer "Envoyer la demande ATMAP" → confirmation → ticket créé + PJ uploadée
```

### Layout des feuilles SIRUS_*

- **Col 1** = Statut (verrouillé, écrit par moteur d'analyse, fond vert OK / rouge NOK + commentaire de cellule avec détail)
- **Col 2..N+1** = Champs config (selon `Ordre` croissant, marqueurs ` *` rouge si Obligatoire, ` (*)` rouge si Conditionnel)
- **Col N+2** = Commentaires (auto-rempli par l'analyse : 1 alerte = texte simple, N alertes = liste à puces `- alerte` + vbLf, WrapText)
- **Ligne 1** = Libellé humain (gras, fond bleu clair, hauteur 45 avec wrap)
- **Ligne 2** = ColonneBdD technique (masquée, sert d'index pour le moteur)
- **Lignes 3+** = Données
- **Freeze panes** (table-aware via `SIRUS_ColonneClePourTable`) :
  - ProjetASL : col 1-4 (Statut + Projet + Sous Système + DTR)
  - Produit : col 1-3 (Statut + Famille Générique + DTR)
  - Autres : col 1-2 (Statut + DTR / Nom projet)
- **VerticalAlignment** : `xlCenter` sur toute la zone de données

### UX UserForm (mise à jour 2026-06-25)

- **Boutons minimize + maximize natifs** sur la barre de titre du UserForm (API `user32.SetWindowLong` dans `Forms/UserForm.frm`, compat Excel 32/64 bit via `#If VBA7`). Permet le va-et-vient PBS ↔ UserForm comme une fenêtre Windows standard.
- **`Project_List_Change`** : guard `Base Is Nothing` → re-`Open_CONN` à la volée si la connexion ADO globale a été fermée par un flow SIRUS.

## Évolutions à venir (non figées)

D'autres règles métier et nouvelles tables sont attendues côté SIRUS. Le code doit rester :
- **Modulaire** — pas de hardcoding de libellés ou d'index ; tout passe par le `Type` et le `Find`.
- **Documenté** — chaque fonction commentée en français, avec entrées / sorties claires.
- **Extensible** — ajouter une colonne à `DATA_ASL_Col` ou une règle de validation ne doit pas casser l'existant.

## Lancement & dépannage

- **Ctrl + W** : ouvre `UserForm` en `vbModeless`.
- **Tests unitaires** (depuis VBE) :
  - `Test_MapColonnesDATA_ASL` — vérifie le cadrage `DATA_ASL` du classeur actif et affiche les numéros de colonne dans la fenêtre Immediate.
  - `Test_EvaluerQualitePrix` — couvre les 6 cas de la matrice qualité prix.
- **Désactiver le lookup BdD** : passer `G_SIRUS_PBS.SIRUS_LookupDTR_Active = False` avant `SIRUS_RemplirListViewGeneration` pour afficher la PBS sans interroger la BdD (les 5 colonnes Projet_ASL / ATM / EMMST / Produit restent vides). Utile pour tester la vue hors-ligne.
- **Mise à jour du complément** : `A_2_ModUpdate.CheckForUpdatesOnStartup` lit `VersionComplement.txt` sur le SharePoint Alstom (`...\007 - Outils métier ISR\01. BDD SLI\11_Livraison de l'OUTIL excel BDD\`).
