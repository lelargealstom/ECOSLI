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
   - **Faire une demande d'import / modification / suppression** de DTR dans la BdD via le pipeline **SIRIUS**.
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
| `G_SIRIUS_*` | Pipeline SIRIUS (demande d'import en BdD) | `G_SIRIUS_Saisie` *(à créer si besoin)*, `G_SIRIUS_Analyse`, `G_SIRIUS_ATMAP`, `G_SIRIUS_PBS`, `G_SIRIUS_UserForm` |
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
| Cadrage colonnes | Toujours via `Range.Find(What:=..., LookIn:=xlValues, LookAt:=xlWhole)` puis fallback `xlPart`. Jamais de magic number. Stockage dans un `Type` dédié (cf. `DATA_ASL_Col` dans `G_SIRIUS_PBS`). |
| Détection ligne d'entête | `DetecterLigneEntete(ws, libelleAncre)` : recherche d'un libellé ancre sur toute la feuille, robuste aux lignes de titre / filtres au-dessus de l'entête. |
| Gestion d'erreur | `On Error GoTo Sortie` + label de sortie. MsgBox informatif (`vbCritical` / `vbExclamation`) pour les erreurs métier visibles utilisateur. `Debug.Print` pour les diagnostics. |
| Affichage tabulaire | ListView ActiveX (`MSComctlLib.ListView`, *Microsoft Windows Common Controls*) en mode `lvwReport`, `FullRowSelect = True`, `Gridlines = True`, `LabelEdit = lvwManual`. Pas de ListBox pour les tableaux. |
| Pas de référence VBA nouvelle | Toutes les références sont déjà déclarées dans le projet ; on n'en ajoute pas. |
| Commentaires | En français, sans accents pour rester compatible cp1252 / UTF-8 à l'import VBA. Les libellés avec accents utilisent `ChrW(...)` côté code. |
| Encodage des libellés | Convention : `"DESIGNATION (issue de PBS)"` reste en clair, mais `"r" & ChrW(233) & "f du plan d'ensemble"` pour les accents — concaténation avec `ChrW` pour garantir la compatibilité multi-éditeurs. |

## Pipeline SIRIUS (focus du chantier en cours)

SIRIUS est le sous-système qui permet à l'utilisateur de pousser des demandes d'import / modification / suppression de données dans la BdD SLI depuis le UserForm ECOSLI.

### Flux

1. **Saisie** (page principale du UserForm) — l'utilisateur sélectionne un type de demande (Création / Modification / Suppression) via `Type_Import_Tab` et remplit les champs Produit ou ATM dans `ListView_Produit` / `ListView_Taches`.
2. **Analyse** (`G_SIRIUS_Analyse.SIRIUS_AnalyserProduit`) — contrôles bloquants et alertants :
   - Champs obligatoires
   - Doublon DTR interne à la saisie
   - DTR déjà présent en BdD (warn pour distinguer Création / Modification)
   - Types de données (MTBF décimal > 0, DVL ≤ 50 ans)
   - Calcul du taux KO ; bloquant si > 20 %
3. **Page Generation** (`G_SIRIUS_PBS.SIRIUS_RemplirListViewGeneration`) — vue PBS du classeur actif avec, par DTR :
   - Présence en BdD (Projet_ASL / Info_ATM / Taches_ATM / EMMST / Produit)
   - Qualité Spare / Repair calculée à partir de 3 sources de prix : Manuel, Fournisseur Alstom, Catalogue ISP. Règles dans `G_SIRIUS_PBS.EvaluerQualitePrix` :
     - 0–1 source connue → `NOK - Manque`
     - 2 sources, écart > 25 % à la médiane → `NOK - Manque & Ecart`
     - 2 sources, écart ≤ 25 % → `NOK - Manque`
     - 3 sources, écart > 25 % → `NOK - Ecart`
     - 3 sources, écart ≤ 25 % → `OK`

### Cadrage `DATA_ASL`

Le module `G_SIRIUS_PBS` cadre la feuille `DATA_ASL` du classeur actif via `MapColonnesDATA_ASL` :
- Auto-détection de la ligne d'entête (libellé ancre : `DTR Reference (issue de PBS)`)
- 40 colonnes mappées dans `DATA_ASL_type` (Identification, RAM, DVL/Préventif, Coûts/Prix)
- Colonnes critiques (7) : `DTR Reference (issue de PBS)`, les 4 prix Spare / Repair (Manuel + Alstom FR + Catalogue ISP) + `ISP TIS FRANCE (FCS)`.
- Si une critique est introuvable → MsgBox récapitulatif + abandon du remplissage.
- Si une optionnelle est introuvable → `Debug.Print` seulement, le remplissage continue.

## Évolutions à venir (non figées)

D'autres règles métier et nouvelles tables sont attendues côté SIRIUS. Le code doit rester :
- **Modulaire** — pas de hardcoding de libellés ou d'index ; tout passe par le `Type` et le `Find`.
- **Documenté** — chaque fonction commentée en français, avec entrées / sorties claires.
- **Extensible** — ajouter une colonne à `DATA_ASL_Col` ou une règle de validation ne doit pas casser l'existant.

## Lancement & dépannage

- **Ctrl + W** : ouvre `UserForm` en `vbModeless`.
- **Tests unitaires** (depuis VBE) :
  - `Test_MapColonnesDATA_ASL` — vérifie le cadrage `DATA_ASL` du classeur actif et affiche les numéros de colonne dans la fenêtre Immediate.
  - `Test_EvaluerQualitePrix` — couvre les 6 cas de la matrice qualité prix.
- **Désactiver le lookup BdD** : passer `G_SIRIUS_PBS.SIRIUS_LookupDTR_Active = False` avant `SIRIUS_RemplirListViewGeneration` pour afficher la PBS sans interroger la BdD (les 5 colonnes Projet_ASL / ATM / EMMST / Produit restent vides). Utile pour tester la vue hors-ligne.
- **Mise à jour du complément** : `A_2_ModUpdate.CheckForUpdatesOnStartup` lit `VersionComplement.txt` sur le SharePoint Alstom (`...\007 - Outils métier ISR\01. BDD SLI\11_Livraison de l'OUTIL excel BDD\`).
