ISP TIS HORS FRANCE (archive)

# ECOSLI.xlam

Complément Excel VBA (add-in) développé pour **Alstom Services Villeurbanne**, équipe **LGM / ISR**. Il automatise la production des livrables SLI (Support Logistique Intégré) à partir d'une PBS projet (`DATA_ASL`) et propose un pipeline (**SIRUS**) de demande d'import / modification / suppression de DTR dans la BdD SLI Alstom.

**Auteur principal :** Thomas LE LARGE (LGM / Alstom).

---

## Contexte métier

L'ingénieur SLI doit produire, pour chaque projet, quatre livrables réglementaires à partir d'une même PBS et d'une BdD de référence :

- **ATM V1** — Approvisionnement Technique Maintenance (format historique)
- **ATM V2** — Approvisionnement Technique Maintenance (nouveau format)
- **EMMST** — Éléments de Maintenance, Manutention, Stockage et Transport
- **Providor** — PBS provisionnée (BreakDownStructure + Maintenance + TNR)

Avant ECOSLI ces livrables étaient reconstruits à la main. ECOSLI centralise la génération, sécurise la cohérence PBS ↔ BdD, et remonte les écarts au sein d'un même UserForm.

En complément de la génération, **SIRUS** permet à l'ingénieur de pousser dans la BdD SLI Alstom les nouveaux DTR (ou les modifs) directement depuis le complément — sans passer par des ressaisies manuelles ni des tickets ATMAP construits à la main.

---

## Lancement

1. Installer l'add-in via `A_1_ModEvents.Install` (copie dans `Application.UserLibraryPath` + activation).
2. Ouvrir le classeur projet contenant la feuille `DATA_ASL` (la PBS).
3. `Ctrl + W` → le UserForm `ECOSLI` s'affiche en `vbModeless` (cf. `W_AppelUsf.Appel_Usf`).

De là, l'utilisateur peut :

- **Générer** les livrables SLI (ATM V1, ATM V2, EMMST, Providor) depuis la PBS active.
- **Rafraîchir la PBS** contre la BdD (cascade prix Spare / Repair, présence multi-tables, calculs locaux).
- **Faire une demande BdD** via SIRUS (Création / Modification / Suppression Produit ou ATM).
- **Visualiser** la cohérence PBS ↔ BdD dans la `ListView_Generation` (10 colonnes : DTR, Nb projets, ATM-Info, ATM-Taches, EMMST, Produit, Prix Spare, Source Spare, Prix Repair, Source Repair).

---

## Structure du dépôt

```
ECOSLI/
├── ECOSLI.xlsm              # Classeur source (VBA + Forms)
├── Champs_BDD.xlsx          # Référentiel des champs BdD
├── Modules/                 # Modules VBA (.bas)
├── Classes/                 # Classes VBA (.cls)
├── Forms/                   # UserForms (.frm / .frx)
├── Templates/               # Templates livrables + template DATA_ASL
└── Docs/                    # Documentation projet
```

### Modules — organisation par préfixe

| Préfixe | Domaine |
|---|---|
| `A_*` | Cycle de vie de l'add-in (Auto_Open/Close, Install, Update) |
| `B_ATM_*` | Génération ATM V1 / V2 |
| `C_EMMST_*` | Génération EMMST |
| `D_Providor_*` | Génération Providor (BreakDown + Maintenance + TNR + Finalisation) |
| `E_Extraction_*` | Extraction feuilles ATM / EMMST existantes (parcours entrée SIRUS A) |
| `F_*` | Fonctions globales / barre de progression |
| `G_PBS_Refresh` | Rafraîchissement PBS ↔ BdD (cascade prix, présence multi-tables) |
| `G_SIRUS_*` | Pipeline SIRUS (Analyse, Config, Extraction, Génération, Listes, PBS, Règles, ATMAP, UserForm) |
| `O_*` `P_*` `Q_*` | Métier transverse (Projet, Nettoyage, Comparaison) |
| `U_*` `V_*` | Mapping / typage ADODB |
| `W_*` `X_*` `Y_*` | UserForm entry + tests + cadrage SLI / templates |
| `Z_*` | Variables globales + extraction de code |

---

## Connexion à la BdD SLI

La BdD SLI est un classeur Excel partagé, consulté par ADO.

- Variables globales : `CONNEXION_EXCEL_BDD_PROD`, `CONNEXION_EXCEL_BDD_TRAVAIL` (cf. `Z_Variables` / `V_Nom_ADODB_Type`).
- Procédures pivot : `Open_CONN` / `open_REQSEL(STRSEL)` / `Close_CONN`.
- Tables adressées en SQL Excel `[NomTable$]` : `Projet_ASL$`, `Info_ATM$`, `Taches_ATM$`, `EMMST$`, `Produit$`.

---

## Pipeline SIRUS — vue d'ensemble

Décisions verrouillées (brief v0.2) :

- **Prix HORS scope** du moteur d'analyse SIRUS (historisation prix = script séparé, `SnapShot_Prix`).
- **L'analyse ne bloque jamais** : statuts OK / Alerte uniquement (taux de complétude MeQ pondéré + cas "vide justifié").
- **Une feuille par (table × opération)** — Créer et Modifier sont distinctes.
- **Listes déroulantes en code** (`G_SIRUS_Listes.SIRUS_GetListe`), pas en feuille.
- **Règles d'analyse en dur** dans `G_SIRUS_Regles` (Select Case par `(Table, ColonneBdD)`), pas de mini-langage.
- Deux parcours d'entrée : (A) réception ATM fournisseur → `E_Extraction_ATM`, (B) mise à jour PBS `DATA_ASL` → détection proactive.
- Deux surfaces : **UserForm = radar/triage**, **Feuille = établi/correction**.
- Cas pilote : **Produit (Créer + Modifier)**, puis généralisation.

Roadmap complète : voir [`Docs/SIRUS_ROADMAP.md`](Docs/SIRUS_ROADMAP.md).

---

## Conventions techniques

| Convention | Détail |
|---|---|
| Source PBS | `ActiveWorkbook` (jamais `ThisWorkbook`) |
| Cadrage colonnes | `Range.Find(LookIn:=xlValues, LookAt:=xlWhole)` puis fallback `xlPart`. Stockage dans un `Type` dédié (ex : `DATA_ASL_Col`). Pas de magic number. |
| Détection ligne d'entête | `DetecterLigneEntete(ws, libelleAncre)` — robuste aux lignes de titre au-dessus |
| Gestion d'erreur | **Pas de handler défensif** (règle inversée 2026-07-20 pour faciliter le debug avec les ingénieurs). Pas de `On Error GoTo <label>` + `MsgBox Err.Description` : les erreurs remontent naturellement à la ligne exacte (surlignage jaune VBE). Seuls sont conservés : les `On Error Resume Next` **fonctionnels** (test d'existence de feuille / clé Dictionary, `CopyFile` idempotent, patterns UI tolerants), les labels `Sortie:` utilisés en early-exit pour du nettoyage (ex. `Close_CONN`), et les `Err.Raise` explicites en pré-condition. |
| Tableaux à l'écran | ListView ActiveX (`MSComctlLib`) en `lvwReport`, `FullRowSelect = True`, `Gridlines = True`. Pas de ListBox. |
| Références VBA | Ne pas en ajouter — toutes déjà déclarées dans le projet |
| Commentaires | Français sans accents (compat cp1252 / UTF-8 à l'import VBA) |
| Libellés accentués | `ChrW(...)` en concaténation (ex : `"r" & ChrW(233) & "f"`) |

---

## Documentation

Toute la doc projet est dans [`Docs/`](Docs/) :

| Fichier | Contenu |
|---|---|
| [`ECOSLI_SYNTHESE.md`](Docs/ECOSLI_SYNTHESE.md) | Synthèse détaillée du complément (architecture, conventions, pipeline SIRUS) |
| [`CHANGELOG.md`](Docs/CHANGELOG.md) | Journal des sessions de développement |
| [`SIRUS_ROADMAP.md`](Docs/SIRUS_ROADMAP.md) | Roadmap SIRUS v0.2 (phases 0 → 5) |
| [`SIRUS_WORKFLOW.md`](Docs/SIRUS_WORKFLOW.md) | Workflow des deux parcours (A / B) et surfaces radar/établi |
| [`SIRUS_REGLES.md`](Docs/SIRUS_REGLES.md) | Règles métier d'analyse SIRUS |
| [`DATA_ASL_TEMPLATE.md`](Docs/DATA_ASL_TEMPLATE.md) | Cadrage du template PBS `DATA_ASL` |
| [`EXTRACTION_LOGIQUE_METIER.md`](Docs/EXTRACTION_LOGIQUE_METIER.md) | Logique métier de l'extraction ATM / EMMST |
| [`TEMPLATES.md`](Docs/TEMPLATES.md) | Templates livrables (ATM FR/EN, EMMST, PHLST, Providor) |
| [`UI_SIRUS_CONTROLES.md`](Docs/UI_SIRUS_CONTROLES.md) | Contrôles UserForm SIRUS |
| [`ATM_V2_PISTES_OPTIMISATION.md`](Docs/ATM_V2_PISTES_OPTIMISATION.md) | Pistes d'optimisation ATM V2 |
| `Spec_Recettes_ECOSLI.xlsx` | Spécifications de recette |
