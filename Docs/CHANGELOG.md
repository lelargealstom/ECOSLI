# Change Log — ECOSLI.xlam

Suivi des modifications apportées au complément Excel `ECOSLI.xlam` au fil des sessions de développement.

Format : `[Date] — Auteur — Périmètre`, puis liste des changements par module.

---

## 2026-07-07 → 2026-07-08 — Sprint 2 : Règle MeQ Produit v0.2 + SnapShotPrix + doc

Refonte complète des règles MeQ Produit selon le brief SIRUS v0.2. Cadrage des 19 champs obligatoires avec logique conditionnelle factorisée (Motif NA / Motif Preventif / NodeType). Câblage du chargement `Produit_SnapShotPrix` (7 champs virtuels prix/leadtime/MoQ/dates) contre le row `Produit$` pour un MeQ complet. Ajout d'un README à la racine + doc des règles Produit dans `SIRUS_REGLES.md`.

### Règles MeQ Produit (matrice v0.2)

Les 19 champs se répartissent en :
- **Base 3 champs** toujours obligatoires : `DTR`, `Designation_PRODUIT`, `Node type Produit (LRU/SRU…)_PRODUIT` (⚠️ ellipse Unicode U+2026 construite via `ChrW`).
- Si `MotifNonApplicabilite_PRODUIT = "NA"` → uniquement les 3 base sont obligatoires.
- Sinon (Motif=OK) → 16 champs restants obligatoires avec 3 exemptions indépendantes :
  - `Duree de vie generique Produit HARDWARE (annee)_PRODUIT` obligatoire ssi `Motif_NonApplicabilitePreventif_PRODUIT ∈ {HARD, BOTH}`
  - `Duree de vie generique Produit (annee) SOFTWARE _PRODUIT` obligatoire ssi `Motif_NonApplicabilitePreventif ∈ {SOFT, BOTH}`
  - 3 champs REPAIR (Prix, Dates validité, Lead time) non obligatoires si `NodeType ∈ {LRUNR, SRUNR}` (règle non-réparable transverse)

Matrice validée par tests unitaires (9 cas) : `NA/*/* → 3`, `HARD/non-NR → 18`, `HARD/NR → 15`, `SOFT/non-NR → 18`, `SOFT/NR → 15`, `BOTH/non-NR → 19`, `BOTH/NR → 16`, `NA(prev)/non-NR → 17`, `NA(prev)/NR → 14`.

### Champs virtuels alimentés depuis `Produit_SnapShotPrix`

Les 7 champs prix/leadtime/MoQ/dates ne vivent plus dans `Produit$` mais dans la table `Produit_HistoPrix` (feuille `Produit_SnapShotPrix`). Le config Produit les garde comme *champs virtuels* (`Source = "SnapShotPrix"`) : ignorés du SELECT `Produit$` (évite ACE errors) mais comptés dans le MeQ. Injection au moment du chargement par `EnrichirProduitAvecSnapShot`.

Règle de retention : **ligne candidate** = `Prix` non nul OU `OnRequestPrix` renseigné. Sur les candidates, on garde la ligne avec `MAX Date Source` par (DTR, Type). Sur cette ligne :
- `Prix` virtuel écrit uniquement si numérique non nul (OnRequest seul → Prix reste vide, MeQ perd 1 point Prix).
- `Ref Catalogue`, `MoQ`, `LeadTime` toujours copiés (nul → non compté par MeQ).

Mapping :
| Champ virtuel Produit | Source SnapShotPrix |
|---|---|
| `Prix SPARE catalogue a date_PRODUIT` | `Prix` (Type=Spare) |
| `Dates de validite des prix SPARE connus sur Projet (ASL)_PRODUIT` | `Ref Catalogue` |
| `MoQ spare catalogue_PRODUIT` | `MoQ` |
| `Lead time SPARE catalogue (jours) a date_PRODUIT` | `LeadTime` |
| `Prix REPAIR catalogue a date_PRODUIT` | `Prix` (Type=Repair) |
| `Dates de validite des prix REPAIR connus sur Projet (ASL)_PRODUIT` | `Ref Catalogue` |
| `Lead time REPAIR (jours) manuel_PRODUIT` | `LeadTime` |

### Bug ACE OLE DB sur DTR mixed-type

`Produit_SnapShotPrix` a la colonne `DTR` mixed-type (numérique 8600xxx + alphanumérique DTR0000xxx / TRVPxxx). ACE OLE DB scanne par défaut les 8 premières lignes pour typer les colonnes → si numériques, type Numeric → les valeurs alphanumériques reviennent NULL même avec `IMEX=1`.

**Workaround retenu** : ajout de `TypeGuessRows=0` dans les Extended Properties de la connection dédiée (force ACE à scanner toutes les lignes). Combiné à `IMEX=1`, les colonnes mixed-type reviennent en TEXT — les DTR alphanumériques sont lisibles. Coût : quelques secondes au 1er accès (scan complet 29789 lignes) → mise en cache session `m_dictSnapShotPrixCache`, reset via `SIRUS_ResetCacheSnapShotPrix`. Fallback direct-read de feuille (open .xlsb) laissé en place pour dépannage.

### Detail des modifications

**`Modules/G_SIRUS_Config.bas`** :
- Bump `SIRUS_CONFIG_VERSION` → `2026-07-07-v3` (force rebuild auto au Ctrl+W).
- `PeuplerProduit` refondu : 19 champs MeQ + 6 champs hors 19 (info libre, `Obligatoire=Non`).
- Rename `Designation manuelle_PRODUIT` → `Designation_PRODUIT` (nouveau nom BdD).
- Ajout `Motif_NonApplicabilitePreventif_PRODUIT` (⚠️ underscore après Motif — nom BdD) et `MoQ spare catalogue_PRODUIT`.
- Rename `Lead time SPARE (jours) manuel_PRODUIT` → `Lead time SPARE catalogue (jours) a date_PRODUIT`.
- 7 champs virtuels prix passés en `Source = "SnapShotPrix"` (skippés du SELECT `Produit$`, comptés en MeQ).
- `Motif NA` passe de `Oui` à `Conditionnel` (honore les "3 obligatoires" quand valeur = "NA").
- NodeType : nom construit via `ChrW(&H2026)` pour l'ellipse Unicode (BdD utilise `…`, code source utilisait `...`).

**`Modules/G_SIRUS_Regles.bas`** :
- Suppression des 2 anciens Case Produit (règle MTBF conditionnelle + règle MotifNA conditionnelle — obsolètes).
- Ajout `EstObligatoire_Produit(colonne, ligne)` : helper privé qui applique la matrice v0.2 en 4 étages (base 3 → court-circuit MotifNA=NA → DVL HW/SW → 3 REPAIR non-NR → 11 mid).
- `SIRUS_RegleObligatoire` fait un early return `If sTable = "Produit"` vers le helper.
- Test unitaire `Test_Regle_Produit_MatriceMeQ` couvrant les 9 cas de la matrice.
- Anciens tests Produit remplacés (MTBFConditionnel obsolète).

**`Modules/G_SIRUS_PBS.bas`** :
- Cache module `m_dictSnapShotPrixCache` + sub publique `SIRUS_ResetCacheSnapShotPrix`.
- `ChargerSnapShotPrixDernierParDTR()` : ADODB pur avec `TypeGuessRows=0` injecté dans la connection string. Filtre "Prix non nul OU OnRequestPrix" pour la retention. Mise en cache session au 1er appel.
- `EnrichirProduitAvecSnapShot(dictProduit, dictSnap)` : injecte les 7 champs virtuels dans chaque row Produit avant `CalculerMeQDTR_Optim`.
- `SIRUS_RemplirListViewGeneration` : appel des 2 fonctions après le chargement de `dictProduit_Rows`.
- Sub debug `Test_SIRUS_DebugMeQProduit_PremierDTR` : dump complet pour le 1er DTR de DATA_ASL (lignes brutes SnapShotPrix + ligne retenue + row Produit + les 19 champs MeQ avec obligation/rempli/valeur).
- Silence du `Debug.Print` "Colonnes optionnelles introuvables" (non bloquant, laisse les vrais warnings).
- Helpers direct-read `OuvrirFeuilleSnapShotPrix` / `FermerBdDSiOuverte` / `ExtraireCheminBdDSirus` / `TrouverColParLibelleSirus` conservés en fallback (utilisent le match par nom de fichier pour supporter le redirect OneDrive/SharePoint).

**`Modules/G_SIRUS_Listes.bas`** :
- `MotifNA` simplifié : `{NA, NA MECA, NA KIT}` → `{OK, NA}`.
- Ajout liste `MotifNAPreventif` : `{HARD, SOFT, BOTH, NA}` (calculée automatiquement depuis `Info_ATM[HARD/SOFT]` par la formule `NB.SI.ENS` — session 2026-07-07).

### Racine du dépôt

- Ajout `README.md` : contexte projet ECOSLI (livrables SLI + pipeline SIRUS), structure, conventions, pointeurs vers `Docs/`.

### Nettoyage `Modules/*.bas`

- Suppression de la ligne `Attribute VB_Name = "..."` en tête de chacun des 41 modules (facilite le copier-coller manuel dans VBE — VBE gère le nom du module automatiquement).

### Formule MotifNonApplicabilitePreventif (colonne BdD calculée)

Ajoutée dans `Produit$` :

```excel
=SI(NB.SI.ENS(Info_ATM[DTR_INFOATM];[@DTR];Info_ATM[HARD/SOFT];"HARD")+NB.SI.ENS(Info_ATM[DTR_INFOATM];[@DTR];Info_ATM[HARD/SOFT];"SOFT")=0;"NA";SI(ET(NB.SI.ENS(Info_ATM[DTR_INFOATM];[@DTR];Info_ATM[HARD/SOFT];"HARD")>0;NB.SI.ENS(Info_ATM[DTR_INFOATM];[@DTR];Info_ATM[HARD/SOFT];"SOFT")>0);"BOTH";SI(NB.SI.ENS(Info_ATM[DTR_INFOATM];[@DTR];Info_ATM[HARD/SOFT];"HARD")>0;"HARD";"SOFT")))
```

Alimente le champ `Motif_NonApplicabilitePreventif_PRODUIT` de `Produit$` : `HARD` si Info_ATM ne contient que des lignes HARD pour ce DTR, `SOFT` si seulement SOFT, `BOTH` si les 2, `NA` si aucune ligne Info_ATM.

---

## 2026-06-27 → 2026-06-28 — Sprint 1 ter : Alignement DATA_ASL / cascade prix / UX

Session post-Sprint 1 bis. Le bouton Rafraichir étant en place, on aligne `MapColonnesDATA_ASL` sur les nouveaux libellés du template, on refond la logique cascade prix (Manuel sorti, sous-priorité D&IS/FCS), on supprime toutes les formules Excel locales au profit d'écritures VBA en valeurs pures, et on ajoute 2 améliorations UX (pré-sélection projet au Ctrl+W + ListView auto-rafraîchie après Refresh).

### Refonte cascade prix (decision Thomas 2026-06-28)

**Nouvelle cascade Spare** : `Fournisseur` → `Cat D&IS` → `Cat FCS calculé` (col AB) → `Archive`.
**Nouvelle cascade Repair** : `Fournisseur` → `Cat D&IS` → `Archive` (pas de FCS Repair en BdD).

- Suppression des champs `PrixManuelSpare` / `PrixManuelRepair` du `Type DATA_ASL_Col` (colonnes retirées du template).
- Ajout de `PrixSpareCalcule`, `RefSourceSpare`, `RefSourceRepair` au `Type`.
- Renames : `CoutTotal` → `PrixFCS`, `ISP_TIS_FCS` → `PrixSpareCatalogue`, `ISP_TIS_Calc` → `PrixSpareCalcule`, `PrixRepa_ISP` → `PrixRepairCatalogue`.
- `EvaluerQualitePrix` / `MedianeConnus` / `CouleurQualite` **supprimées** (logique trio Manuel/Frs/Cat avec médiane abandonnée).
- Nouvelles fonctions `CalculerPrixSpare` / `CalculerPrixRepair` / `ProjetArchiveAlphaPremier` / `AjouterSubItemPrix` / `AjouterSubItemSource` dans `G_SIRUS_PBS.bas`.
- `G_SIRUS_Generation.CascadePrixDATA_ASL` (paramètre Manuel) remplacée par `CascadePrixSpareDATA_ASL` et `CascadePrixRepairDATA_ASL` (utilisées pour pré-remplir `SIRUS_ProjetASL_Modifier`).

### Refonte ListView Generation : 12 → 10 colonnes

Avant : `DTR | ProjetASL | InfoATM | Taches | EMMST | Produit | QualSpare | EcartSpare | ManqueSpare | QualRepair | EcartRepair | ManqueRepair`
Après : `DTR | Nb projets | ATM-Info | ATM-Taches | EMMST | Produit | Prix Spare | Source Spare | Prix Repair | Source Repair`

- `SIRUS_ConfigListViewGeneration` mis à jour dans `G_SIRUS_UserForm.bas` (4 ColumnHeaders en moins, 4 en plus).
- Affichage Source coloré selon origine :
  - 🟢 vert = Fournisseur
  - 🔵 bleu = D&IS catalogue (libellé = `Ref Source Spare/Repair` lu en col AC/AF, ex `D&IS 2026-2027_v03`)
  - ⚪ gris = FCS calculé (Spare uniquement, valeur lue en col AB = Prix FCS × 1,0557)
  - 🟠 orange = `Projet : <nom>` (archive, projet alphabétiquement 1er pour ce DTR)
  - 🔴 rouge gras = aucun prix dispo

### Refresh : changements logique

- **`RafraichirPrixCatalogue`** : suppression de la cascade D&IS→FCS→AUTRE pour les prix. Désormais col AA (`Prix Spare Catalogue`) = **D&IS strict** (MAX Date), col AE (`Prix Repair Catalogue`) = D&IS strict. Col Z (`Prix FCS`) reste FCS strict (inchangé). Col AC (`Ref Source Spare`) garde une mini-cascade D&IS → FCS pour le **label uniquement** (la ListView en a besoin quand la cascade tombe sur FCS via col AB). Col AF (`Ref Source Repair`) reste D&IS strict (pas de FCS Repair). Helper `ChoisirCascade` supprimé.
- **`RafraichirPrixArchive`** : suppression du calcul AVG sur multi-projets (pas de colonne Date dans `Projet_ASL$`). Pour chaque DTR, on retient désormais le **projet alphabétiquement 1er** (StrComp ascendant) parmi ceux ayant un prix > 0, projet courant exclu. Cols AN/AO contiennent le prix de ce projet retenu. La ListView Generation re-lit `dictProjetASL_Rows` pour reconstruire le label `"Projet : <nom>"`.
- **`RafraichirCalculs`** (nouveau) : reconstruit en valeurs pures les 6 colonnes "calcul local" (zéro formule dans le fichier) :
  - Col I (`NTI`) = Select Case NodeType (LRU/LRUNR → "3", SRU/SRUNR → "3+", sinon vide)
  - Col N (`Tx déf fonct`) = 1 / MTBF fonctionnel (col L) si > 0
  - Col O (`Tx déf gen`) = 1 / MTBF générique (col M) si > 0
  - Col AB (`Prix Spare Calculé`) = Prix FCS × cellule AB2
  - Col AL (`Cout Spare Alstom`) = Cout Spare FRS Not ISP × cellule AL2
  - Col AM (`Prix Repair Alstom`) = Cout Repair FRS Not ISP × cellule AM2
- Helpers `EstNumeriqueNonNul` (filtre Empty/Null/0/non-numérique pour ne pas écrire "0,00 €" quand la source est vide) et `ArrondiPrix` (arrondi commercial 2 décimales via `Application.WorksheetFunction.Round`).
- **`AppliquerFormatPrix`** : applique le format Comptabilité € (`_-* #,##0.00\ "€"_-;...`) à 10 colonnes prix (Z/AA/AB/AE/AH/AJ/AL/AM/AN/AO) en fin de Refresh. Évite l'affichage `373,179.39` (long décimal avec thousand-sep) sur les multiplications par coeff.
- **Appel automatique `SIRUS_RemplirListViewGeneration`** en fin de Refresh : la ListView Generation se recalcule avec les nouvelles données sans nécessiter d'action utilisateur.

### Alignement `MapColonnesDATA_ASL` sur les libellés du nouveau template

41 libellés réécrits dans `G_SIRUS_PBS.bas` pour matcher exactement les en-têtes du template `Templates/ECOSLI_DATA_ASL.xlsx`. Corrections principales :
- `SOUS SYSTÈME (copiés de la PBS...)` → `SOUS SYSTÈME`
- `Tx déffaillance ...` (typo historique) → `Tx défaillance ...`
- `MTBF intrinsèque (RAM report specifique projet (CGR))` → `MTBF intrinsèque RAM Report spécifique`
- `DUREE DE VIE LIMITEE générique (offre et PGR) HARDWARE` → `DVL HARD générique`
- `DUREE DE VIE LIMITEE (RAM report specifique projet (CGR)) HARDWARE` → `DVL HARD spécifique`
- `réf du plan d'ensemble (issue de la BdD)` → `Réf plan ensemble`
- `UNITE de la maintenance corrective/préventive` → `UNITÉ maintenance corrective/préventive`
- Les 4 renames prix (Cout total → Prix FCS, etc.).

Critères critiques rationalisés : DTR + PrixSpareCatalogue + PrixSpareCalcule + PrixRepairCatalogue + CoutSpare_AlstomFR_FRS + PrixRepair_Alstom_FRS.

### UX UserForm

- **`UserForm_Initialize`** : si `ActiveWorkbook` a une cellule nommée `DATA_ASL_PROJET` non vide (template ECOSLI_DATA_ASL déjà rempli par un Refresh précédent), `Project_List.Value` est pré-positionné dessus → l'ingé n'a plus besoin de re-sélectionner le projet à chaque Ctrl+W. Wrap `On Error Resume Next` pour le cas classeur sans cellule nommée.
- **Garde-fou connexion BdD** dans `SIRUS_RemplirListViewGeneration` : si `Base Is Nothing` ou `Base.State <> 1` (cas après `RafraichirPrixArchive` qui fait `Close_CONN`), réouvre la connexion à la volée. Évite l'affichage "Absent" / 0 partout sur les colonnes MeQ après un Refresh.

### Bug fix divers

- `Test_EvaluerQualitePrix` remplacé par `Test_CalculerPrix` (mock feuille temporaire avec 4 cas de cascade).
- `Test_MapColonnesDATA_ASL` met à jour les références aux champs renommés.
- Le diagnostic `Test_RafraichirPrixDryRun` est déjà compatible avec la nouvelle logique D&IS/FCS strict.

### Fichiers modifiés

- `Modules/G_SIRUS_PBS.bas` (gros : Type + Map + ListView + cascades + connexion guard)
- `Modules/G_SIRUS_Generation.bas` (refonte CascadePrix*)
- `Modules/G_PBS_Refresh.bas` (cascades simplifiées + RafraichirCalculs + format € + call ListView)
- `Modules/G_SIRUS_UserForm.bas` (ListView Generation 10 cols)
- `Forms/UserForm.frm` (auto-select projet au Ctrl+W)

### Reste à faire (prochaines sessions)

- Mettre à jour les PowerShell `Templates/_build_template.ps1` + `_patch_template_v2.ps1` pour ne plus poser les formules dans F/I/M/N/W/AB/AL/AM (cohérence "zéro formule"). Pas bloquant car le Refresh écrase, mais évite la confusion sur un template fraîchement généré.
- Phase 2.1 — Enrichir règles InfoATM / EMMST / ProjetSLI.
- Phase 3.2 — Snapshot Modifier.
- Phase 6 — Industrialisation (Historique, Paramètres, Recette).

---

## 2026-06-25 → 2026-06-26 — Sprint 1 bis : Bouton Rafraichir DATA_ASL

Sprint 1 bis livré : refonte complète du flow d'alimentation de DATA_ASL. Plus aucun lien externe SharePoint (les formules `=XLOOKUP(https://...)` fragiles aux migrations sont remplacées par un bouton **"Rafraichir Data PBS"** dans le UserForm qui pull les données via ADODB / lecture Excel directe).

### Template figé `Templates/ECOSLI_DATA_ASL.xlsx`

- **41 colonnes** dans l'ordre verrouillé Thomas (5 blocs : Identification PBS / Générique / RAM / Maintenance / Prix).
- **Logo Alstom** + titre fusionné + bandeau identité (lignes 1-2 zone réservée).
- **Cellules nommées** `DATA_ASL_PROJET` (I1) et `DATA_ASL_DERNIERE_MAJ` (I2) auto-remplies au Refresh.
- **Légende code couleur** : vert saisie ASL / rouge Refresh BdD / bleu calcul local / gris constante.
- **Coefficients ligne 2** : `AB2 = 1,0557` (Prix Spare calculé), `AK2 = 2`, `AL2 = 1,17` (Cout Alstom Spare), `AM2 = 1,75` (Prix Alstom Repair).
- **Labels coefficients ligne 1** (italique gris pour clarté).
- **ListObject `Providor`** conservé (couvre A4:AO + dernière ligne, style `TableStyleLight2`). Nom intact pour compat downstream `D_Providor_Main` / `B_ATM_Main` / `C_EMMST_Main`.
- **Freeze panes** : ligne 4 + cols A-D (4 premières colonnes visibles au scroll horizontal).
- **41 formats cellules** par type (DTR=Texte, MTBF/DVL/LeadTime=Entier, Tx défaillance=Scientifique, Durées tâches=0.000, Prix=#,##0.00).
- **15 commentaires** sur les entêtes clés (tooltip Excel au hover).
- **Mise en forme conditionnelle** : 18 cellules Refresh vides → fond orange clair "À rafraichir".
- **Protection feuille + workbook** : mdp `LGM69000`, `UserInterfaceOnly=True` (VBA peut écrire en Refresh, l'user uniquement dans les 12 cellules vertes saisie ASL). `Structure=True` côté workbook (pas de suppression/ajout/renommage de feuilles).
- **Bouton du template retiré** : déplacé dans le UserForm pour une UX unifiée et éviter les soucis de résolution macro `'ECOSLI.xlam'!Sub`.
- Script PowerShell `Templates/_build_template.ps1` (génération initiale) + `Templates/_patch_template_v2.ps1` (ajout formules locales NTI col I + Durée SOFT gen col W, suppression bouton template).

### Module `Modules/G_PBS_Refresh.bas` (nouveau, ~1200 lignes)

Point d'entrée public : `SIRUS_RafraichirDATA_ASL([wbParam])` câblé au bouton **`Btn_Rafraichir_DataPBS`** ajouté dans le UserForm (à côté de `Project_List`).

**6 sub-routines par bloc / source** :
1. `RafraichirDepuisProduit` — Désignation BdD, Node typed, Réf plan, MTBF générique, DVL HARD/SOFT générique (cols E, G, J, M, S, U) via `[Produit$]`.
2. `RafraichirHARDSOFT` — agrégat `[Taches_ATM$]` → BOTH/HARD/SOFT/"" (col H).
3. `RafraichirDepuisInfoATM` — Temps générique HARD (col P) + SUM Temps analytique HARD/SOFT (cols Q et X) via `[Info_ATM$]`.
4. `RafraichirConstantes` — "PH" (col R) et "Y" (col Y) en dur.
5. `RafraichirPrixCatalogue` — **lecture Excel directe** (bypass ADO, cf. ci-dessous) sur feuille `Produit_SnapShotPrix` avec cascade DIS → FCS → AUTRE.
6. `RafraichirPrixArchive` — AVG `[Projet_ASL$]` avec exclusion projet courant (cols AN, AO).

**Cascade prix Spare/Repair** (décision Thomas 2026-06-26) :
- Priorité 1 : `Ref Catalogue` contient `D&IS Catalogue` (préfix UCase + InStr)
- Priorité 2 : `Ref Catalogue` contient `FCS`
- Priorité 3 : tout autre catalogue (PROVIDOR, AGATE MEDIA, etc.)
- Pour chaque combinaison `(DTR × Type × Catégorie)`, on retient la ligne avec MAX(Date Source).
- À l'écriture, on prend DIS si dispo, sinon FCS, sinon AUTRE. La colonne **`Ref Source Spare/Repair`** affiche le `Ref Catalogue` du choix retenu (traçabilité).
- Exception : `Prix FCS` (col Z) reste **strictement** FCS (sinon le nom n'a plus de sens).

**Bypass ADO pour SnapShotPrix** (point dur résolu en session) :
- **Bug Microsoft ACE OLE DB sur .xlsb** : quand une colonne contient des types mixtes (numérique + alphanumérique), ACE scanne les 8 premières lignes pour deviner le type. Si elles sont toutes Double, ACE catégorise la colonne entière comme Double → toutes les valeurs String plus loin sont retournées **NULL silencieusement** (pas d'erreur). Le cast `[DTR] & ''` ne suffit pas (scan TypeGuessRows précède l'évaluation des expressions). `MaxScanRows=0` déprécié sur ACE 16.0+, registry `TypeGuessRows=0` nécessite admin.
- **Solution** : pour `Produit_SnapShotPrix` uniquement, on ouvre le fichier `.xlsb` directement dans Excel via `Application.Workbooks.Open(... ReadOnly=True)` et on lit la feuille en bulk array (`Range.Value`). Lecture native Excel = tous les types préservés, plus rapide qu'ADO ligne par ligne.
- Helpers ajoutés : `ExtraireCheminBdD()` (parse `CONNEXION_EXCEL_BDD_TRAVAIL`) + `TrouverColParLibelle(ws, sLibelle)` (lookup `Range.Find` case-insensitive).
- Détection si la BdD est déjà ouverte chez l'utilisateur (`For Each wb In Application.Workbooks`) → ne pas la refermer si oui.
- Pour les autres tables (Produit, Info_ATM, Taches_ATM, Projet_ASL), ADO continue de marcher avec le cast `[col_DTR] & '' AS DTR_str` qui prévient le problème dans la majorité des cas (les 8 premières lignes contiennent assez de variété).

**Helpers transverses** :
- `ChargerDtrsPBS(lo)` : dict `{ DTR_UC → row_index DataBodyRange }`.
- `LireProjetCourant()` : depuis `UserForm.Project_List.Value`.
- `EcrireCellulesProjetDate(wb, sProjet)` : remplit `DATA_ASL_PROJET` et `DATA_ASL_DERNIERE_MAJ`.
- `TracerRefreshLog(wb, sProjet, nbDtrs, duree)` : feuille `_RefreshLog` veryHidden, append Date / Projet / nbDTR / Durée / Utilisateur.
- `DeProtegerFeuille` / `ProtegerFeuille` (mdp `LGM69000`, `UserInterfaceOnly=True`).
- `GetFieldSafe` / `NzS` / `NzD` / `ChoisirCascade` (cascade DIS → FCS → AUTRE).

**5 Subs de diagnostic** ajoutées (à exécuter dans le VBE pour debug futur) :
- `Test_DumpFeuillesBdD` : liste toutes les tables/feuilles visibles par ADO dans la BdD TRAVAIL.
- `Test_DumpSnapShotPrix` : schéma de la table + 5 premières lignes.
- `Test_DistinctRefCatalogue` : couples `(Type, Ref Catalogue)` distincts avec nb d'occurrences.
- `Test_CompterPatternsDTR` : classification DTRs par pattern (numérique pur / DTRxxxxx / ATXxxxxx / autre alpha).
- `Test_RafraichirPrixDryRun` : dry-run de la logique prix sur un DTR donné (modifier la constante `TEST_DTR`).

### `Forms/UserForm.frm` — handler bouton

Nouveau handler ajouté :
```vba
Private Sub Btn_Rafraichir_DataPBS_Click()
    Call G_PBS_Refresh.SIRUS_RafraichirDATA_ASL
End Sub
```

Le bouton physique `Btn_Rafraichir_DataPBS` est à ajouter manuellement dans le VBE Designer (à côté de `Project_List`) — les contrôles ActiveX/MSForms vivent dans le binaire `.frx`, pas dans le `.frm`.

### Décisions verrouillées

- **Le bouton de Refresh vit dans le UserForm**, plus dans le template Excel. Plus simple à gérer (pas de résolution macro `'ECOSLI.xlam'!Sub` cross-workbook), UX cohérente avec le reste (Project_List + boutons SIRUS).
- **DATA_ASL conserve sa structure 40 colonnes** (renommée 41 avec Ref Source Spare/Repair). Le ListObject `Providor` est préservé pour la compat 100% Providor/ATM/EMMST.
- **Cascade DIS → FCS → AUTRE** pour les prix Catalogue (Spare et Repair). Strict FCS pour Prix FCS.
- **Bypass ADO** pour SnapShotPrix uniquement (les autres tables continuent en ADO).
- **Projet courant exclu de la moyenne archive** (col AN/AO) : `WHERE [PROJET_PROJETASL] <> '<projet>'`.
- **Constantes** "PH" / "Y" écrites en dur par le Refresh (cols R/Y).
- **Cellules nommées** `DATA_ASL_PROJET` / `DATA_ASL_DERNIERE_MAJ` : préférables aux coordonnées en dur côté VBA.

### Workflow utilisateur (Sprint 1 bis end-to-end)

```
1. L'ingé ASL ouvre son fichier projet ECOSLI_DATA_ASL.xlsx (template)
2. Saisit les colonnes vertes : DTR, Designation PBS, Sous Système, Qte,
   MTBF intr/fct, DVL HARD/SOFT spec, Cout/LeadTime FOURNISSEUR
3. Ctrl+W → UserForm ECOSLI
4. Sélectionne son projet dans Project_List
5. Clic bouton "Rafraichir Data PBS"
   → Lecture des DTRs de DATA_ASL
   → 4 requêtes ADODB (Produit, Taches_ATM, Info_ATM, Projet_ASL)
   → 1 lecture Excel directe (Produit_SnapShotPrix)
   → Écriture en valeurs pures des 20 colonnes Refresh
   → Constantes PH/Y posées en R/Y
   → Cellules I1/I2 auto-remplies (projet + date)
   → Trace _RefreshLog
   → MsgBox final "N DTR(s) traité(s) en X.X s"
6. L'ingé peut ensuite générer ATM / EMMST / Providor (qui lisent DATA_ASL
   inchangée — zéro impact downstream)
```

### Docs ajoutées / mises à jour

- `Docs/DATA_ASL_TEMPLATE.md` (créé session précédente) : spec figée du template, conservé pour référence.
- `Docs/CHANGELOG.md` : cette entrée.

### Reste à faire (prochaines sessions)

- Aligner `MapColonnesDATA_ASL` dans `G_SIRUS_PBS.bas` avec les renommages du template (`Cout total` → `Prix FCS`, `ISP TIS FRANCE (calculé)` → `Prix Spare Calculé`, `ISP TIS FRANCE (FCS)` → `Prix Spare Catalogue`, `Prix repa (ISP Catalogue Répa)` → `Prix Repair Catalogue`). Sinon la ListView_Generation cherche encore les anciens noms.
- Migration legacy : routine `PurgerLiensExternesDATA_ASL` distribuée via `A_2_ModUpdate` au démarrage si version PBS < seuil (pour les classeurs projets anciens déployés chez les autres ingés ASL).
- Retirer les Subs `Test_*` du module `G_PBS_Refresh` une fois le code stabilisé (ou les déplacer dans un module séparé `_Tests.bas`).
- Sprint 1 — Test Extraction ATM BaWu POOL (toujours non terminé).
- Phase 2.1 — Enrichir règles InfoATM / EMMST / ProjetSLI.
- Phase 3.2 — Snapshot Modifier.
- Phase 6 — Industrialisation (Historique, Paramètres, Recette).

---

## 2026-06-11 → 2026-06-25 — Sessions Claude Code (avec Thomas LE LARGE)

Sprint 1 SIRUS — recette end-to-end Produit Créer + ProjetASL Modifier validés. Refonte complète des règles métier ProjetASL, mapping VBN/IMP ATMAP, MeQ Generation par table avec affichage % puis nombre de projets, croisement DATA_ASL pour ProjetASL Modifier, cadrage de la suite (industrialisation + refonte DATA_ASL).

### Phase 5 ATMAP — Split numérotation + date estimée + workflow documenté

- **Numérotation séparée SP** : `Numero_ticket` = `VBN_NN` auto-incrémenté (parse `VBN_<n>` sur les 500 derniers tickets, max+1) ; `Numero_Import` = `IMP-YYYY-NNN` (extrait du sujet). Régression `ChercherMaxNumeroImport` corrigée : le parsing tapait encore sur `Numero_ticket` au lieu de `Numero_Import` après le split.
- `CreateTicketREST` accepte 2 paramètres optionnels supplémentaires : `numeroImport` (champ SP `Numero_Import`) et `dateEstimee` (format ISO `yyyy-mm-ddT00:00:00Z`, champ SP `Date_estimee`).
- **Date estimée** = aujourd'hui + N jours selon priorité (1/3/5 pour High/Medium/Low). Helper centralisé `SIRUS_JoursTraitementSelonPriorite(priorite)` + `SIRUS_DateEstimeeISO(priorite)` dans `G_SIRUS_ATMAP`. Une seule source de vérité utilisée à la fois par `SIRUS_UI_PrioriteChange` (texte Attentes UI) et `SIRUS_UI_BtnEnvoyerClick` (champ SP).
- `Docs/SIRUS_WORKFLOW.md` (nouveau) : 4 statuts (Nouveau / En cours / Refusé / Résolu terminal), transitions autorisées, mapping création ticket SIRUS → champs SharePoint, 3 flows Power Automate proposés (Refus, Résolu, Retard sur Date_estimee).

### Fix Project_List_Change : guard Base = Nothing

`Project_List_Change` plantait régulièrement avec "Object variable not set" parce que les flows SIRUS (`ChargerProjetsSLI`, `PreRemplirDepuisBdD`...) appellent `Close_CONN` après leurs requêtes, ce qui met la connexion globale `Base` à `Nothing`. Le handler tape ensuite sur une connexion fermée.

- Guard `If Base Is Nothing Or Base.State <> 1` → re-`Open_CONN(CONNEXION_EXCEL_BDD_TRAVAIL)` à la volée
- `On Error GoTo Sortie` + MsgBox propre au lieu de crash UserForm
- Sync auto `SIRUS_UI_ProjetListChange` si Zone Envoi visible
- Escape des apostrophes dans la valeur projet

### Refonte ListView_SIRUS_Resultats : 1 ligne par alerte

- Avant : 1 ligne par DTR avec toutes les alertes concaténées par `|` dans la colonne Alertes.
- Après : 1 ligne par alerte avec 3 colonnes **Cle / Champ / Alertes**, rouge gras.
- `sirusAlertes` (collection module-level dans `G_SIRUS_Analyse`) alimentée par `EvaluerChampLigne` lors de l'analyse (`Array(cle, champ, alerte, ligne, colChamp)`).
- `SIRUS_UI_LvResultatsDblClick` parse la `key` ListView pour `Goto` directement sur la cellule fautive (ligne + colonne).
- Cle table-aware : DTR pour Produit/InfoATM/TachesATM/EMMST/ProjetASL, Nom projet pour ProjetSLI.

### Colonne Commentaires auto-remplie par l'analyse

`SIRUS_AnalyserFeuille` écrit désormais le détail des alertes dans la colonne `_commentaires` (col N+2 de la feuille SIRUS) :
- 1 alerte → texte simple
- N alertes → liste à puces `- alerte1` + vbLf + `- alerte2` + ... avec `WrapText = True`
- Localisation dynamique via `TrouverColonneCommentaires` (scan ligne 2 pour le marqueur `_commentaires`)
- Helper `FormaterAlertesCommentaire(sStatutLigne, nAlertes)` + helper `NzSafeStr` pour gérer les `Null` ADO.

### UserForm_Initialize : SIRUS_PhaseInit_v2 wiré

L'appel manquait : `Opt_SIRUS_Creer.Value = True`, frames cachés et ComboBox Priorité (High/Medium/Low en `fmStyleDropDownList`) n'étaient pas initialisés au démarrage.

### Helper centralisé SIRUS_ColonneClePourTable

Nouvelle Public Function dans `G_SIRUS_Config` qui renvoie le numéro de colonne contenant la clé métier d'une feuille SIRUS selon la table :
- ProjetASL → col 4 (DTR ; col 2 = Projet, col 3 = Sous Système)
- Produit → col 3 (DTR ; col 2 = Famille Générique)
- Autres → col 2 (DTR ou Nom projet)

Utilisé par `SIRUS_AnalyserFeuille` (ligneMax + sCle), `SIRUS_CompterLignesFeuille`, `SIRUS_GenererFeuilleAnalyse` (colFreeze = colCle + 1), `SIRUS_RemplirListViewResultats`.

### Refonte complète ProjetASL (config + règles + layout)

- **Ordre verrouillé Thomas 2026-06-11** : 21 colonnes dans l'ordre Statut / Nom Projet / Sous Système / DTR / Designation PBS / Qte / Node Type / MTBF Intr / MTBF Fct / DVL fonctionnel HARD / DVL fonctionnel SOFT / Prix Spare archive / Lead Time Spare archive / Prix Repair archive / Lead Time Repair archive / Testabilite / Prix Preventif / Prix After Market Spare / Prix After Market Repair / STBI/DAT / LCS / Commentaires.
- **Renommages** : DVL générique HARD/SOFT → DVL fonctionnel HARD/SOFT. Lead time Spare → Lead Time Spare archive (idem Repair). Prix Spare (devise projet) → Prix Spare archive.
- **Obligatoires Oui** : Nom Projet, Sous Système, DTR, Designation PBS, Qte, MTBF fct, Prix Spare archive, Lead Time Spare archive (Testabilité passée à Non le 2026-06-15).
- **Obligatoires Conditionnels (NodeType LRU/SRU)** : Prix Repair archive, Lead Time Repair archive (logique métier validée — KIT/VISSERIE/ASSEMBLY n'ont pas de cycle réparation).
- **Testabilité** : `TypeChamp = "Liste"`, `ListeSource = "Booleen"` (dropdown Oui/Non). Puis Obligatoire passée à "Non" le 2026-06-15 (n'est plus critère bloquant).
- **Ajout `Prix Spare archive`** (Ordre 110) : nouveau champ obligatoire qui n'existait pas auparavant. Nom BdD supposé `Prix spare (DEVISE du projet (SLI))_PROJETASL` (à vérifier en BdD).
- **Layout** : col 1-4 figées (Statut + Projet + Sous Système + DTR), `VerticalAlignment = xlCenter` sur toute la zone.
- **Cle métier (ListView Resultats, colCle, freeze, scan ligneMax)** : col 4 (DTR) pour ProjetASL.

### Phase 2.1 — Règles métier ProjetASL

`G_SIRUS_Regles.bas` étendu :
- Cases `Prix repair (DEVISE du projet (SLI))_PROJETASL` + `Lead time REPAIR (jours)_PROJETASL` → obligatoires si NodeType IN (LRU, SRU).
- `MTBF fonctionnel` : règle "toujours True" (validation post-CGR).
- `PROJET_PROJETASL` : nouvelle règle alerte "Projet inconnu dans ProjetSLI : import impossible". Lookup via cache module-level `projetsSLI_Cache` (Dictionary chargé une fois depuis BdD TRAVAIL), API publique `SIRUS_ProjetsSLI()` + `SIRUS_ResetCacheProjetsSLI()`.

### Garde-fou envoi ATMAP : projet inconnu en ProjetSLI

`SIRUS_UI_BtnEnvoyerClick` bloque l'envoi si **au moins un projet** des feuilles `SIRUS_ProjetASL_*` n'existe pas dans `Projet_SLI$`. MsgBox vbCritical avec la liste des projets inconnus + Exit Sub. Helper `SIRUS_ProjetsASL_Inconnus(feuilles)` scanne la col 2 (Nom projet) de chaque feuille ProjetASL et croise avec le cache `SIRUS_ProjetsSLI`. Critère d'intégrité référentielle Table-mère ProjetSLI → Table-fille ProjetASL.

### PreRemplirDepuisBdD robuste

Modification de la stratégie SELECT dans `G_SIRUS_Generation.PreRemplirDepuisBdD` :
- Avant : `SELECT [col1], [col2], ..., [colN]` — plantait si une seule colonne n'existait pas en BdD (cas des `??_*_a_definir` et de nouveaux champs comme Prix Spare).
- Après : `SELECT * FROM table WHERE [cleCol] = ...` puis lecture champ-par-champ tolérante. Colonnes manquantes tracées en `Debug.Print` avec compteur. MsgBox d'erreur enrichi avec le SQL en cas de plantage.

### Affichage progressif assoupli — Frame Analyse + Envoi ensemble

Décision 2026-06-11 : pour la phase de test, on permet l'envoi tout le temps (plus de gating MeQ ≥ 70%). Le blocage MeQ reviendra plus tard.
- `SIRUS_UI_BtnGenererClick` : Frame Analyse + Frame Envoi visibles ensemble dès la Generation, `Btn_SIRUS_Envoyer.Enabled = True`, `SIRUS_UI_PreRemplirEnvoi` appelée à la Generation (plus à l'Analyse).
- `SIRUS_UI_BtnAnalyserClick` : ne gate plus la visibilité de l'envoi.

### Croisement DATA_ASL pour ProjetASL Modifier

Pour ProjetASL Modifier : la source du projet n'est plus `Cmb_SIRUS_Projet` (combo dans Frame_Modifier) mais **`Project_List.Value`** (combo principale en haut du UserForm — celle où l'ingénieur ASL choisit son projet ECOSLI). Après le pré-remplissage BdD, `SIRUS_EnrichirProjetASL_DepuisDATA_ASL` overwrite les champs PBS depuis DATA_ASL :

| Champ SIRUS | Source DATA_ASL |
|---|---|
| Nom Projet | `Project_List.Value` (constante pour toutes lignes) |
| Sous Système | `SousSysteme` |
| DTR | `DTR_PBS` |
| Designation PBS | `Designation_PBS` |
| Qte | `Qte` |
| Node Type | `NodeType` |
| MTBF intrinsèque Projet | `MTBF_Intr` |
| MTBF fonctionnel | `MTBF_Fonct` |
| DVL fonctionnel HARD | `DureeVie_Ana_HARD` |
| DVL fonctionnel SOFT | `DureeVie_Ana_SOFT` |
| Lead Time Spare archive | `LeadTimeSpareCalend` |
| Lead Time Repair archive | `LeadTimeRepairCalend` |
| Prix Spare archive | cascade Fournisseur → Catalogue → Manuel → Archive (helper `CascadePrixDATA_ASL` reprend la règle Providor `RemplirPrixSpare`) |
| Prix Repair archive | idem côté Repair |

DTRs présents en DATA_ASL mais absents de la BdD ProjetASL → nouvelles lignes ajoutées en fin (utile pour compléter la BdD au fur et à mesure).

Helpers privés `TrouverColParColonneBdD(ws, sColBdD)` (lookup col SIRUS par ColonneBdD en ligne 2) + `CascadePrixDATA_ASL` (1re valeur `IsNumeric AND <> 0` parmi 4 colonnes priorisées).

### Phase 6.2 amorcée — Refonte ListView_Generation : MeQ par table BdD

Refonte complète de `SIRUS_RemplirListViewGeneration` (G_SIRUS_PBS) :
- Remplacement de l'affichage OUI/NON par **pourcentage MeQ** par DTR pour chacune des 4 tables BdD (Produit, InfoATM, TachesATM, EMMST). ProjetASL traité différemment (voir ci-dessous). ProjetSLI hors scope.
- **Couleurs** : 100% vert, < 100% rouge gras, "Absent" rouge gras si DTR pas en BdD.
- Calcul basé sur `SIRUS_LireConfig(table)` + `SIRUS_RegleObligatoire(table, colBdD, ligne)` : num = champs Oui remplis + Conditionnels actifs remplis, den = tous les Oui + Conditionnels évalués `True`.
- **Multi-rows par DTR** (TachesATM) : moyenne pondérée (sum num / sum den) sur toutes les rows.
- **Col ProjetASL = nombre de projets distincts** où ce DTR existe en BdD `Projet_ASL$` (au lieu d'un MeQ). Vert si ≥ 1, rouge gras si 0.
- **Tooltip de ligne** (`ListItem.ToolTipText`) : `Projets connus (N) :` + liste avec `- ProjA` + vbLf + `- ProjB` + ... (MSComctlLib v6 ne supporte pas le tooltip par sub-item, donc tooltip de ligne entière, hover sur la cellule DTR).
- Refactor dicts module-level `dictXxx_Rows` (au lieu de `dictXxx` = set de DTRs) : `Dictionary { DTR_UCase → Collection of Dictionary { ColonneBdD → val } }`.

### Optimisations performance Generation

- **`SIRUS_LireConfig` cachée** par table en début de remplissage : 5 appels au lieu de 5×N (économie majeure pour N>50 DTRs).
- **SELECT ciblé** : ne charge que les colonnes `Source = "Saisie"` du config + clé DTR + extra (PROJET_PROJETASL pour ProjetASL). EMMST passe de ~50 colonnes à ~15. Exclut les placeholders `??_*_a_definir`.
- **Fallback SELECT \*** si la requête ciblée échoue (col config non alignée avec BdD réelle) — tolérance + trace `Debug.Print`.
- Fonctions renommées `ChargerRowsTableBdD_Optim` / `CalculerMeQDTR_Optim` reçoivent la `configCol` pré-chargée.

### Null safety dans ChargerRowsTableBdD

Symptôme : pop-up `[EMMST$] ChargerRowsTableBdD : Utilisation incorrecte de Null`, idem `[Produit$]`. Cause : `NzStr` (F_FonctionsGlobales) fait `CStr(v)` sans gérer Null → crash sur les cellules vides ADO. Fix : helper local `NzSafeStr` (guard IsNull/IsEmpty/IsError) + ne jamais stocker Null dans le rowDict (remplacé par "").

### UserForm : bouton minimize + maximize natifs (API user32)

`Forms/UserForm.frm` : bloc API ajouté en tête du code-behind.
- Declare `FindWindow_API`, `GetWindowLong_API`, `SetWindowLong_API`, `DrawMenuBar_API` avec double définition `#If VBA7` pour compat Excel 32 / 64 bit.
- Sub `UI_ActiverBoutonMinimize` ajoute les flags `WS_MINIMIZEBOX` + `WS_MAXIMIZEBOX` au style de la fenêtre (classe `ThunderDFrame`).
- Appel dans `UserForm_Initialize` après `SIRUS_PhaseInit_v2`.
- Tolérant : `On Error Resume Next` autour de l'API — si `FindWindow` échoue, les boutons ne sont pas ajoutés mais le UserForm fonctionne.
- Note : pas de `UserForm_Resize` adaptatif → maximisation laisse du vide en bas/droite. Refonte layout adaptatif à programmer plus tard si besoin.

### Cadrage technique — Refonte DATA_ASL (audit + décisions)

Audit complet via Explore agent sur l'usage des 40 colonnes `DATA_ASL_Col` à travers l'add-in :
- **20 colonnes safe à retirer** (jamais consommées hors PBS) : Designation_BdD, HardSoft (obsolète depuis B_ATM_V1 calcul BdD), NTI, RefPlan, MTBF_Gen, TxDef_Gen, DureeTacheCorrPrev_Gen_HARD, TxDef_Fonct, DureeTacheCorrPrev_Ana_HARD, UniteMaintCorrective, DureeVie_Gen_HARD, DureeVie_Gen_SOFT, DureeTachePrev_Gen_SOFT, DureeTachePrev_Ana_SOFT, UniteMaintPreventive, CoutTotal, ISP_TIS_Calc, CoutSpare_FRS_NotISP, LeadTimeSpare_FRS, CoutRepair_FRS_NotISP, LeadTimeRepair_FRS.
- **12 colonnes critiques à re-router** avant retrait : tDTR, tDES_PBS, tDES_MEF, tSS, tQTE, tNodeType, tMTBF_Intr, tMTBF_fct, tDureevie_Ana_HARD/SOFT, tPrix_FRS_Spare/Repair, tPrix_Cat_Spare/Repair, tPrix_Manuel_Spare/Repair, tPrix_SpareProjet_Archive, tPrix_RepairProjet_Archive (lues par D_Providor_Main + B_ATM_Main + C_EMMST_Main).
- **Challenge UI** : MSFlexGrid écarté définitivement (référence ActiveX `mshflxgd.ocx` introuvable sur poste Alstom verrouillé + retiré officiellement Microsoft 2008 + pas de download propre). Worksheet `SIRUS_Generation_Vue` retenue (zero dépendance, pattern déjà adopté avec `Btn_SIRUS_EditerFeuille`).
- **Cible validée** : bouton "Rafraichir DATA_ASL" dans le UserForm qui pull depuis BdD via ADODB et écrit en valeurs pures les 24 colonnes "issues BdD" (plus de formules `=RECHERCHEX` à liens externes SharePoint fragiles). Compatibilité descendante 100% Providor/ATM/EMMST (les modules continuent à lire DATA_ASL inchangée). À implémenter en prochaine session.

### Docs créées / mises à jour

- **`Docs/SIRUS_WORKFLOW.md`** (nouveau) : statuts ATMAP, transitions, mapping création ticket, flows Power Automate à mettre en place.
- **`Docs/SIRUS_REGLES.md`** (nouveau) : règles ProjetASL (obligatoires, conditionnels, alertes, garde-fou envoi, layout, mapping DATA_ASL → ProjetASL avec cascade prix Providor, section MeQ Generation avec tableau colonnes ListView + couleurs + optim perf documentées). Template à étendre pour les autres tables.

### Décisions verrouillées

- **Testabilité non obligatoire** (2026-06-15) : `PoidsMeQ = 0`, `Obligatoire = "Non"`. N'entre plus dans le MeQ Generation ni dans le MeQ envoi SIRUS.
- **ProjetASL Modifier** : source projet = `Project_List.Value` (combo principale UserForm), plus `Cmb_SIRUS_Projet`.
- **Cascade prix Spare/Repair** (DATA_ASL → ProjetASL archive) : Fournisseur → Catalogue → Manuel → Archive (reprend exactement la règle `D_Providor_BreakDownStructure.RemplirPrixSpare/Repair`).
- **Affichage progressif** : Frame Analyse + Envoi visibles ensemble dès la Generation (gating MeQ retiré temporairement).
- **Cle ProjetASL** : col 4 (DTR), pas col 3 (Sous Système).
- **Refonte DATA_ASL** : bouton Rafraichir (valeurs ADODB pures, garde la structure 40 col) plutôt que retrait physique (qui nécessiterait re-routage Providor/ATM/EMMST). Phase B "retrait physique" reportée.

### Reste à faire (prochaine session)

- **Bouton Rafraichir DATA_ASL** (Étape 1) : module `G_PBS_Refresh.bas` + `Btn_Rafraichir_DATA_ASL` dans la page Generation + feuille trace `_RefreshLog`. Bloquants à clarifier : structure de `SnapShotPrix` (clé composite ? Date max ?) + mapping libellé DATA_ASL → table BdD pour prix Catalogue/Fournisseur. Possibilité : passe diagnostique auto qui scanne les formules existantes pour inférer le mapping.
- Sprint 1 — Test Extraction ATM BaWu POOL (non terminé).
- Phase 2.1 — Enrichir règles métier InfoATM / EMMST / ProjetSLI.
- Phase 3.2 — Snapshot Modifier (feuille `_Snap` veryHidden + diff dans PJ ATMAP).
- Phase 6.1 — Journal historique (feuille `SIRUS_Historique`).
- Phase 6.2 — Paramètres globaux (`G_SIRUS_Parametres.bas`).
- Phase 6.3 — `Docs/SIRUS_RECETTE.md` distribuable aux ingés ASL.
- Vue Worksheet `SIRUS_Generation_Vue` (long terme, après stabilisation).

---

## 2026-06-10 — Session Claude Code (avec Thomas LE LARGE)

Session très chargée — refonte complète du pipeline SIRUS (Phases 2 → 5), refonte de la génération ATM V1 (HARD/SOFT pilotée par BdD, sommaire consolidé, remapping SOFT), refonte tolérante de l'extraction ATM, et bouclage du workflow UI end-to-end.

### Phase 2 SIRUS — Moteur d'analyse (`G_SIRUS_Regles` + `SIRUS_AnalyserFeuille`)

- **`G_SIRUS_Regles.bas`** (nouveau) : API publique `SIRUS_RegleObligatoire(table, colonneBdD, ligne)` + `SIRUS_RegleAlerte(...)`. Pattern D1 validé : `Select Case Table|ColonneBdD`, pas de mini-langage. 5 règles Produit + 4 TachesATM + 1 ProjetASL + 3 alertes (prix sans date, Nb opérateurs > 4, MTBF fct < intrinsèque).
- **`G_SIRUS_Analyse.bas`** étendu : nouvelle `SIRUS_AnalyserFeuille(sTable, [op]) As Double`. Pour chaque ligne, construit Dictionary `{ColonneBdD → valeur}`, applique les règles (Obligatoire config + RegleObligatoire conditionnelle + RegleAlerte), écrit le Statut col 1 (OK vert / NOK (n) rouge) + commentaire de cellule avec détail des alertes. Renvoie le taux MeQ global (somme poids remplis obligatoires / somme poids obligatoires).
- Helpers privés : `EvaluerChampLigne`, `LireLigneEnDictionary`, `EstLigneVide`, `EcrireStatutCellule`, `AjouterAlerte`.
- L'ancien `SIRUS_AnalyserProduit` (ListView du UserForm legacy) reste fonctionnel.

### Refonte layout SIRUS (toutes tables)

Avant : champs en col 1..N + Statut en col N+1. Après : **Statut col 1 + champs col 2..N+1 + Commentaires col N+2**.
- `G_SIRUS_Generation.PoserEntetes` : Statut en col 1 (`ColonneBdD = "_statut"`), Commentaires en bout (`"_commentaires"`).
- `AppliquerListes` / `GriserChampsAutoHerite` / `VerrouillerEtProteger` / `PreRemplirDepuisBdD` : tous décalés de +1 pour placer les champs en col 2..N+1.
- Focus initial sur `(3, 2)` (1re cellule de saisie).
- `PoserLibelleEntete` : suffixe rouge gras ` *` (Oui) ou ` (*)` (Conditionnel) sur l'entête pour signaler visuellement les champs obligatoires. `Characters(...).Font` pour colorer uniquement le suffixe.
- AutoFit avec bornes `14..35` + hauteur entête 45 pour le wrap.

### Phase 3.1 SIRUS — Verrouillage cellules + protection feuille

- Constante module `SIRUS_PROTECTION_PWD = "SIRUS"` (à centraliser Phase 6.2).
- Nouveau helper `VerrouillerEtProteger(ws, col)` :
  - Déverrouille toutes les cellules (`.Cells.Locked = False`)
  - Re-verrouille entêtes (lignes 1+2), col 1 (Statut), colonnes Auto/Herite. Col N+2 (Commentaires) reste **déverrouillée** (saisie libre user).
  - `ws.Protect Password:="SIRUS", UserInterfaceOnly:=True, DrawingObjects:=False` (VBA peut écrire + AddComment).
- Appelé dans `SIRUS_GenererFeuilleAnalyse` après `GriserChampsAutoHerite`.

### Refonte config Produit (25 champs)

`PeuplerProduit` refondu avec nouvel ordre métier (validation Thomas) :
Famille Generique, DTR, Designation manuelle, Node Type, Ref RAM Report, Fournisseur, Référence fournisseur, Motif NA (liste NA/NA MECA/NA KIT), Drawing Number, MTBF intrinsèque, DVL HARD/SOFT générique, Prix Spare manuel, Lead Time Spare, Source Spare, Date Source Spare, Prix Repair manuel, Lead Time Repair, Source Repair, Date Source Repair, Reference datasheet, Reference MUL installation/maintenance/software/interface.
- Prix Spare/Repair manuel + Date Source = `Obligatoire = "Non"` (saisie libre).
- Nouveaux champs marqués `??_*_a_definir` (à confirmer via `Test_DiagnostiquerEcartsConfig_Toutes`).
- Champs retirés (hors scope SIRUS) : MoQ Spare, Réparabilité D&IS, Prix REPAIR catalogue, Lead time REPAIR catalogue, LCS_Repair, Data_Catalogue_Repair, Désignation BdD (Auto).
- `G_SIRUS_Listes.bas` : ajout `SIRUS_LIST_MOTIFNA` = `NA / NA MECA / NA KIT`.

### Nettoyage ProjetSLI

Retrait des 6 champs hors scope SIRUS (Ordre 280-330) : Requis client Acceptance/Clausing/MoP, Requis fournisseur Clausing, PSLI préliminaire/final.

### Phase 4 entrée A — Extraction ATM/EMMST → SIRUS (étape D)

- **`E_Extraction_Main.Main_Extraction_Params(sFichier, sLivrable, sLangue, [bAfficherMsgFin])`** (nouvelle, paramétrée, découplée UserForm) : ouvre le fichier, valide, lance Extraction. L'ancien `Main_Extraction` devient wrapper rétrocompat. Retourne le `Workbook` livrable ouvert.
- **`G_SIRUS_Extraction.bas`** (nouveau module) :
  - `SIRUS_DeverserExtractionDansSIRUS(sTable, sFichier, sLangue, [sVersion])` : orchestrateur. ATM → 2 feuilles (`SIRUS_InfoATM_Extraction` + `SIRUS_TachesATM_Extraction`). EMMST → 1 feuille.
  - `DeverserFeuilleVersSIRUS` : mapping `ColonneBdD → index col SIRUS` via ligne 2 (entête masqué) avec `NormaliserNomBdD` (trim + lowercase) pour matching tolérant. MsgBox diagnostic : N mappées / N ignorées + liste des colonnes ignorées.
  - Helper `ForcerFormatTexte` : applique `NumberFormat = "@"` sur les champs `TypeChamp = "Texte"` du config (notamment Rang) pour éviter conversion "1-1" → date.

### Extraction ATM tolérante (`E_Extraction_ATM_Tolerant.bas` — nouveau module)

Algo de détection par **libellé** au lieu de positions fixes (D4/D5/I6/I7/I8/I9). Préserve **toute** la logique métier (DTR=Left(13), HARD/SOFT=Right(4), prefix REPLACE pour EN, Version="A1", catégorie chapitre, Connu/Inconnu BdD, mapping temps total → col N HARD / O SOFT).

- `ExtraireInfoATM_T` : `LireValeurApresLibelle` cherche libellé dans `A1:N12`, lit cellule droite (ou dessous si droite vide). Fallback positions historiques.
- `ExtraireTachesATM_T` : `TrouverLigneHeaderTaches` (cherche "Description" + "Outillage" en col B), puis `MapperColonnesTaches` (par libellé d'entête) avec fallback positions. `EstFeuilleVide_T` cherche pattern "x-y" en col B.
- `ValiderFormatLivrable` (E_Extraction_Main) : assoupli — mots-clés ("Maintenance", "ATM", "Tâches", "Tasks", "MTA") dans `A1:J10` d'**au moins 1 fiche DTR**, au lieu du titre exact en B2 de **chaque** fiche.
- Validé sur le fichier BaWu POOL (61 fiches DTR TCMS HARD).

### Refonte ATM V1 — HARD/SOFT piloté par BdD

- **`PrecalculerTypesDTR_BdD`** (nouveau) : parcourt `Taches_tab` (déjà chargé par `RecupererDonneesBDD_ATM`) et alimente `dictTypeDTR = {DTR → HARD/SOFT/BOTH/""}`. Plus de dépendance à `tHARD_SOFT` (DATA_ASL).
- `GetTypeDTR(dtr)` helper d'accès. `""` = pas de tâche en BdD → fiche HARD vide + "À faire" sommaire (continuité métier).
- `PrecalculerLignesSommaire` calcule `mapLigneHARD(i)` : DTR BOTH = 2 lignes consécutives (HARD puis SOFT) au lieu d'avoir toutes les SOFT en fin de tableau.
- `RemplirIndexV1` pose les lignes consécutives (même désignation, sans suffixe " (SOFT)").
- `PoserLienSommaireSOFT` (remplace `AjouterLigneSommaireSOFT`) : pose le lien hypertexte sur la ligne pré-allouée.
- `MarquerStatutIndex`, `CreerLienHypertexte`, supports/frames "NA" utilisent `mapLigneHARD(indexDTR)` au lieu de `indexDTR + 5`.

### Fix nommage doublons fiches ATM V1

Bug : pour un DTR BOTH, la fiche SOFT était nommée `DTR_TCMS (2)` (sans suffixe). Cause : `CompterDoublonsFeuilles` comparait sur baseName seul → confondait HARD et SOFT du même DTR. `NommerFeuilleV1` (cpt > 0) construisait le nom sans préserver le suffixe.
- `CompterDoublonsFeuilles` : nouveau param `typeHS`, compare avec **nom complet** `baseName + _HARD/_SOFT`.
- `NommerFeuilleV1` (cpt > 0) : `nomComplet & " (" & cpt & ")"` au lieu de `baseName & " (cpt+1)"`. Suffixe préservé : `DTR_TCMS_HARD (1)`.

### Remapping rangs SOFT (template 4 chapitres)

Le template `ATM_SOFT`/`MTA_SOFT` n'a que **4 marqueurs** (pas 5 comme HARD). Mapping rang BdD → chapitre SOFT (validé Thomas) :
- `1-x` BdD → Chapitre 1 SOFT
- `2-x` BdD → Chapitre 2 SOFT
- `3-x` BdD → Chapitre 3 SOFT (rang inchangé)
- `4-x` BdD → Chapitre 3 SOFT, renuméroté `3-(maxRang3 + x)`
- `5-x` BdD → Chapitre 4 SOFT, renuméroté `4-x`

Implémentation : pré-passe calcule `maxRang3` (plus grand x des rangs 3-x SOFT). `RemplirTachesV1` Select Case sur `rangChiffre`. `InsererTacheV1` accepte un param optionnel `rangAffiche` qui prime sur la valeur BdD. **Comportement HARD strictement inchangé**.

`destSOFT.Columns("C").WrapText = False` après remplissage SOFT : description sur 1 seule ligne.

### MTBF/DVL : SOFT sans MTBF + DVL contextuelle + scalabilité

- `RemplirDonneesRAM_V1(dest, indexDTR, typeHS)` refondue : pour SOFT, MTBF n'est plus posé en I4 (logique métier). DVL choisie selon `typeHS` (et plus `tHARD_SOFT`).
- Helpers `CalculerMTBF` (hiérarchie fonctionnel/intrinsèque/générique selon gate + RAM Report) et `CalculerDVL` (col P/R pour HARD, col U/W pour SOFT selon gate).
- **`PoserValeurApresLibelle`** : `Range.Find` des libellés "MTBF" / "Périodicit"/"Periodicit"/"Periodicity" dans `A1:K12`, écrit cellule droite. Fallback I4/I5. Rend la génération **scalable** aux versions de template (positions peuvent bouger).

### Phase 5 SIRUS — Envoi ATMAP

**`G_SIRUS_ATMAP.bas`** étendu (API publique) :
- `SIRUS_NumeroTicketSuivant()` : GET SharePoint sur ATMAP_TICKETS filtré `Typedemande='Import'`, parse `Numero_ticket` au format `IMP-YYYY-NNN`, retourne max+1 de l'année en cours. **Passage d'année** : retour à 001 si aucun ticket pour l'année en cours.
- `SIRUS_ConstruirePJ(feuilles, sNomBase)` : copie les feuilles SIRUS du projet vers un classeur temp dans `%TEMP%`, sauve en `.xlsx`. Multi-feuilles ATM → 1 fichier unique avec les 2 feuilles (logique : 1 livrable = 1 PJ).
- `SIRUS_AttacherFichierAuTicket(ticketID, digest, cheminFichier)` : upload PJ via `XMLHTTP POST` + `ADODB.Stream` binaire.
- Wrappers publics : `SIRUS_GetFormDigest`, `SIRUS_GetCurrentUserID`, `SIRUS_CreerTicketEnvoi` (defaults : statut="Nouveau", typedemande="Import", outilID=43).

### UI Workflow 3 zones (UserForm Frame_SIRUS_v2)

**`G_SIRUS_UserForm.bas`** étendu (handlers `SIRUS_UI_*`) :
- `SIRUS_PhaseInit_v2` : init au démarrage (combo Table/Op, masque sous-zones, peuple `Cmb_SIRUS_Priorite` avec High/Medium/Low + `Style = fmStyleDropDownList`, masque Frames Analyse/Envoi).
- `SIRUS_RafraichirComboFeuilles` : scanne `ActiveWorkbook` et alimente `Cmb_SIRUS_FeuilleAnalyse` avec toutes les feuilles `SIRUS_<Table>_<Op>` présentes (exclut Config, Diag, Historique, _Snap). Permet de retrouver les feuilles déjà générées dans une session précédente.
- `SIRUS_UI_BtnGenererClick` : dispatch Créer/Modifier/Extraction. Pour Modifier ProjetSLI/ProjetASL : clé depuis `Cmb_SIRUS_Projet`. Pour Extraction : appelle `SIRUS_DeverserExtractionDansSIRUS`.
- `SIRUS_UI_BtnAnalyserClick` : appelle `SIRUS_AnalyserFeuille`, remplit la ListView NOK, active `Frame_SIRUS_Envoi` si MeQ ≥ 70%, pré-remplit la Zone 3 via `SIRUS_UI_PreRemplirEnvoi`.
- `SIRUS_UI_BtnEnvoyerClick` complet (9 étapes) : validation → confirmation `vbYesNo` → PJ → digest → userID → extraction numéro → ticket → upload PJ → cleanup + statut.
- `SIRUS_UI_PrioriteChange` : MAJ `Txt_SIRUS_Attentes` selon priorité (1/3/5 jours).
- `SIRUS_UI_ProjetListChange` : MAJ Description si Frame Envoi visible. Pour ProjetSLI/ProjetASL, ignoré (la source est la feuille).
- `SIRUS_UI_PreRemplirEnvoi` : auto-rempli Sujet (`IMP-YYYY-00X - Import <Op> <Table>`), Description (`<Op> de <N> référence(s) en <Table> pour le projet <ProjetContextuel>`), Priorité Medium par défaut.
- `SIRUS_UI_BtnEditerFeuilleClick` (porte de sortie) : active la feuille du combo + `UserForm.Hide`. Retour via Ctrl+W.

### Affichage progressif des frames

- Au démarrage : Frame Analyse + Frame Envoi cachés.
- Après Génération : Frame Analyse visible.
- Après Analyse MeQ ≥ 70% : Frame Envoi visible + pré-remplissage auto.

### ListView NOK : rouge gras + clé selon table

- `SIRUS_RemplirListViewResultats` : appel `SIRUS_ConfigListViewResultats` au début (idempotent — recrée les ColumnHeaders, sinon `SubItems(1)` lève "Invalid property value").
- Clé affichée : col 3 (DTR) pour ProjetASL, col 2 (DTR ou Nom projet) pour les autres.
- `ForeColor = RGB(163,45,45)` + `Bold = True` sur item et subitems.
- Helper `SafeCStr` (gère Null/Empty/Error/autres types).

### Projet contextuel pour la description ATMAP

- `SIRUS_ProjetPourDescription` : pour ProjetSLI/ProjetASL, lit col 2 ligne 3 de la feuille SIRUS active (= projet en cours), pas `Project_List`. Fallback `Project_List` si feuille vide.
- `SIRUS_MajDescriptionEnvoi` : helper unifié pour reconstruire la description avec la bonne logique projet.

### Freeze panes par table

`SIRUS_GenererFeuilleAnalyse` fige automatiquement les colonnes selon la table :
- Produit / InfoATM / TachesATM / EMMST / ProjetSLI : col 1-2 (Statut + DTR ou Nom projet)
- ProjetASL : col 1-3 (Statut + Projet + DTR)

### Multi-lignes ProjetASL (`PreRemplirDepuisBdD`)

Bug : pour ProjetASL Modifier avec 1 clé projet (N DTR), seule la 1re ligne du Recordset était écrite (`If Not REQSEL.EOF` au lieu de boucle).
Fix : `Do While Not REQSEL.EOF` + `MoveNext` + `ligne = ligne + 1` à chaque tour.

### Bouton dans le Sommaire ATM V1 SOFT

- Pour les DTR BOTH : 2 lignes consécutives dans Tableau1 (HARD puis SOFT), même désignation, lien hypertexte distinct vers chaque fiche. Statut "À faire" sur ligne SOFT si pas de tâches SOFT en BdD.

### Nettoyage encoding (446 occurrences)

5 modules avec caractères corrompus `U+FFFD` (cp1252 mal importé) remplacés par leur équivalent ASCII selon contexte :
- `C_EMMST_Main.bas` (156), `E_Extraction_Main.bas` (2), `G_SIRUS_ATMAP.bas` (91), `G_SIRUS_UserForm.bas` (75), `UserForm.frm` (122).
- Mots métier traités : Creer, Generer, Selection, donnees, recuperer, repertoire, etape, procedure, executer, definir, designation, necessaire, ecrire, complete, different, presence, present, specifique, succes, securite, controle, meme, plutot, periode, periodicite, etc.

### Docs créées / mises à jour

- **`Docs/EXTRACTION_LOGIQUE_METIER.md`** (nouveau) : référence complète de la logique métier de l'extraction ATM/EMMST (algo actuel + ce qui doit être préservé dans la refonte tolérante).
- **`Docs/UI_SIRUS_CONTROLES.md`** : enrichi avec Zone 3 ATMAP détaillée (récap + saisie + action + bouton Editer + freeze panes + affichage progressif + ComboBox Priorité + coloration ListView + handler `Project_List_Change`).
- **`Docs/6_2026_BaWu POOL_ATM_1A.xlsx`** : fichier de référence (61 fiches DTR TCMS HARD) utilisé comme cas de test pour la refonte tolérante.
- **`Docs/Spec_Recettes_ECOSLI.xlsx`** : 33 lignes ajoutées (F3.6.8-12, F3.7.12-14, F6.5-7, F7 SIRUS complet F7.1 à F7.6.3) en version cible V2.

### Décisions verrouillées

- **HARD/SOFT depuis BdD** (pas DATA_ASL) pour ATM V1 — décision Thomas 2026-06-10.
- **Doublons PBS** : suffixe `(1)`, `(2)`... uniquement si même DTR doublé en PBS pour le même type (HARD ou SOFT séparément).
- **Template SOFT ATM V1** : 4 chapitres seulement, rangs 3+4 fusionnés dans Chap 3 avec renumération continue, rang 5 → Chap 4 renuméroté 4-x.
- **Numéro ticket Import** : auto-incrément SharePoint, format `IMP-YYYY-00X`, passage d'année avec retour à 001.
- **Multi-feuilles ATM (Extraction)** : 1 ticket = 1 PJ unique contenant les 2 feuilles SIRUS (logique 1 livrable = 1 demande).
- **Projet pour ProjetSLI/ProjetASL** : pris depuis la feuille SIRUS active (col 2 ligne 3), pas `Project_List`.

### Reste à faire

- Tester end-to-end le workflow SIRUS (Sprint 1 du plan).
- Phase 3.2 : snapshot Modifier (feuille `_Snap` veryHidden + diff dans PJ ATMAP).
- Phase 2.1 : enrichir `G_SIRUS_Regles` (EMMST, InfoATM, ProjetSLI complets).
- Enrichir `PeuplerInfoATM` / `PeuplerTachesATM` / `PeuplerEMMST` pour couvrir 100% des colonnes brutes d'extraction (cf. MsgBox diagnostic).
- Phase 4 entrée B (`SIRUS_DetecterEcartsPBS`) + Phase 6 (Historique, Paramètres, Recette).

---

## 2026-06-09 (Phase 1) — Session Claude Code (avec Thomas LE LARGE)

### Périmètre
**Phase 1 SIRUS** : générateur de feuille d'analyse. Module dédié `G_SIRUS_Generation.bas`.

### `G_SIRUS_Generation.bas` (nouveau module)
- **`SIRUS_GenererFeuilleAnalyse(table, operation, [cles])`** — point d'entrée public.
  - Crée la feuille `SIRUS_<Table>_<Operation>` dans **`ActiveWorkbook`** (= classeur projet, jamais le `.xlam`).
  - **Confirmation utilisateur** (`MsgBox vbYesNo`) si la feuille existe déjà.
  - Lit le config via `SIRUS_LireConfig(table)`.
  - **Ligne 1** = libellé `Champ` humain (gras, fond bleu clair, hauteur 30, wrap).
  - **Ligne 2** = `ColonneBdD` technique (gris, italique, masquée). Sert d'index pour le moteur Phase 2 et la PJ ATMAP Phase 5.
  - Colonne `Statut` en bout pour le retour du moteur d'analyse.
  - **Listes déroulantes** (`xlValidateList`) sur les colonnes `TypeChamp = "Liste"` ou `"Booleen"`, source `SIRUS_ListeEnChaine(ListeSource)`. Couvre 1000 lignes.
  - **Champs `Source = "Auto"` ou `"Herite"`** : fond gris clair (verrouillage réel reporté Phase 2).
  - **`Operation = "Modifier"`** : ouvre ADODB (`Open_CONN`/`open_REQSEL`/`Close_CONN`), SELECT par clé sur la table BdD, remplit ligne par ligne. Clé absente → colonne Statut = "Cle non trouvee" en rouge.
  - Sortie : Activate + Freeze panes sous l'entête.
- **Helpers privés** : `PoserEntetes`, `AppliquerListes`, `GriserChampsAutoHerite`, `PreRemplirDepuisBdD`, `NomTableSQL` (mapping `Produit` → `Produit$`, `InfoATM` → `Info_ATM$`, etc.), `FeuilleExiste`.
- **Constantes publiques** : `SIRUS_OP_CREER`, `SIRUS_OP_MODIFIER`.
- **Procédures de test** : `Test_Generer_Produit_Creer`, `Test_Generer_Produit_Modifier` (avec 2 DTR fictifs à remplacer par des DTR réels), `Test_Generer_ProjetSLI_Creer`, `Test_Generer_EMMST_Creer`.

### Décisions
- **Localisation** : feuille dans `ActiveWorkbook` (cohérent avec `DATA_ASL` et l'export PJ ATMAP). Erreur explicite si `ActiveWorkbook` = `.xlam`.
- **Entête sur 2 lignes** : ligne 1 humaine + ligne 2 technique masquée (validé Thomas).
- **Confirmation écrasement** : `MsgBox vbYesNo`, défaut "Non" pour éviter les pertes accidentelles.
- **Clé de filtrage Modifier** = `ColonneBdD` du 1er champ du config (Ordre 10) — DTR pour la plupart, `Nom projet_PROJET SLI` pour ProjetSLI.
- **Champs `ColonneBdD` vides** (cas du MTBF intrinsèque info Produit dans ProjetASL) : ignorés dans le SELECT, laissés vides côté feuille.

### Reportés Phase 2
- Snapshot des valeurs d'origine pour le diff Modifier.
- Verrouillage réel des cellules `Auto`/`Herite` (avec protection feuille).
- Évaluation des règles `Source = "Saisie"` obligatoires / conditionnelles via `G_SIRUS_Regles`.
- Calcul du taux de complétude (MeQ) et écriture du Statut.

### Reste à faire
- Importer `G_SIRUS_Generation.bas` dans le `.xlam`.
- Ouvrir un classeur projet (`Ctrl + N` ou ECOSLI.xlsm), l'activer.
- Tester `Test_Generer_Produit_Creer` → la feuille `SIRUS_Produit_Creer` apparaît dans le classeur projet avec 14 colonnes + Statut, listes déroulantes sur Node Type (LRU/SRU/…), colonne Designation BdD grisée.
- Tester `Test_Generer_Produit_Modifier` avec des DTR réels présents en BdD.
- **Phase 2** : `G_SIRUS_Regles.bas` (règles obligatoires/alertes en dur) + généraliser `SIRUS_AnalyserProduit` en `SIRUS_AnalyserFeuille(table)`.

---

## 2026-06-09 (suite) — Session Claude Code (avec Thomas LE LARGE)

### Périmètre
Alignement de `ColonneBdD` sur les vrais noms BdD (référentiel `Champs_BDD.xlsx` fourni par Thomas). Refactor complet des 6 `Peupler*`.

### Changements
- `.gitignore` : exclusion d'`Champs_BDD.xlsx` (2.2 MB binaire) et des CSV temporaires `_tmp_*.csv`.
- `G_SIRUS_Config.bas` — tous les `Peupler*` mis à jour avec les vrais noms BdD :
  - Produit (14 champs) : `Désignation manuelle_PRODUIT`, `Node type Produit (LRU/SRU…)_PRODUIT`, `Référence fournisseur_PRODUIT`, `Drawing number` (sans suffixe table), `MoQ spare catalogue_PRODUIT`, `MTBF intrinsèque Produit (h) issu de la datasheet _PRODUIT` (avec espace avant `_PRODUIT`), `MotifNonApplicabilité_PRODUIT`, `Prix SPARE/REPAIR manuel_PRODUIT`, `Dates de validité des prix SPARE/REPAIR connus sur Projet (ASL)_PRODUIT`, `Désignation_PRODUIT` (champ BdD lu en Auto).
  - InfoATM (8 champs) : `DTR_INFOATM`, `HARD/SOFT` (sans suffixe), `DescriptifATMFR/EN_INFOATM`, `PLANATM_INFOATM`, `Reparabilite_INFOATM`, `Poids_INFOATM`, `Temps générique de la tâche_INFOATM`.
  - TachesATM (7 champs) : `DTR_TACHEATM`, `Rang_TACHEATM`, `Langue_TACHEATM`, `HARD/SOFT`, `Description_TACHEATM`, `CritèresDimensionnels_TACHEATM` (pour Critère Nm — **à confirmer**), Banc test marqué `??_BancTest_a_definir` (pas trouvé dans Champs_BDD).
  - EMMST (38 champs au lieu de 24) : DTR + dimensions nu/conditionné (Longueur/Largeur/Hauteur/Poids) + conditionnement (Unité, Date péremption, Durée péremption, Réutilisable) + marquage/manutention (Étiquette, Marquage, Manutention, Outillage spécifique, Précaution, Type outil) + stockage (14 champs booléens et textes : Position, Extérieur, Sous abri, Air libre, Hors poussière, Hors eau, Intérieur, Chaud, Froid, Hors lumière, Gerbable, Moyen, Précaution, Maintenance) + commentaires (Emballage, Marquage, Manutention, Stockage, Transport).
  - ProjetSLI (15 champs) : `Nom projet_PROJET SLI` (cle), `Type de train_PROJET SLI`, `Durée de circulation_PROJET SLI`, `Nombre de train_PROJET SLI` (singulier), `Servicing day by year_PROJET SLI`, `PH/RH/Km by year_PROJET SLI`, `Durée de garantie_PROJET SLI`, `Devise_PROJET SLI`, `MAN hour cost_PROJET SLI`, `Niv de maintenance_PROJET SLI`. **3 champs sans correspondance Champs_BDD** marqués `??_*_a_definir` : Reference PBS/EAP, Reference RAM report, LCC scope total.
  - ProjetASL (12 champs au lieu de 11) : `PROJET_PROJETASL` (cle), `DTR_PROJETASL`, `Désignation_PROJETASL`, `Q PBS_PROJETASL`, `Node typed PROJET_PROJETASL` (pré-rempli depuis Produit), MTBF intrinsèque en lecture pure (ColonneBdD vide, `SourceHeritage=Produit.MTBF intrinsèque...`), DVL séparé en **HARD** et **SOFT** (au lieu d'un seul "DVL générique" du brief — à confirmer), `MTBF fonctionnel (h) issu du RAM report projet_PROJETASL`, `Lead time SPARE/REPAIR (jours)_PROJETASL`, `Testabilité (DTR du BTE...)_PROJETASL`.

Total après refactor : ~94 champs (vs 79 avant), majoritairement grâce à l'extension EMMST.

### Extension du config (validation Thomas 2026-06-09)
Extensions Produit, ProjetSLI, ProjetASL pour couvrir le périmètre métier réel. Total config passe de **~98 à ~158 champs**.

- **Produit (14 → 20)** : ajout de 6 colonnes Repair/catalogue (Réparabilité D&IS Service France VBN Repair Catalog, Prix REPAIR catalogue à date, Lead time REPAIR catalogue/manuel, LCS_Repair, Data_Catalogue_Repair).

- **ProjetSLI (15 → 62)** — refactor majeur. Périmètre fixé avec Thomas :
  - **Inclus** : identité (Nom projet, Projet/Offre, Region, CT Code, Type de train, Client final), configuration (Durée circulation, Nombre trains, Durée garantie, Niv maintenance, Nombre DTR EAP/PBS, Nombre fournisseurs CDRL, Type de pack, UO grand père, Statut), profil de mission (Servicing/PH/RH/Km by year), économique (Devise, MAN hour cost, LCC scope total), GDP (4 champs), Requis client/fournisseur (4), documents PSLI préliminaire/final, références (PBS/EAP, RAM report, Version Providor/BASL), **interlocuteurs (RSLI, RASL, LU + 7 LU + 6 VBN + 11 DL par scope + Autres)**.
  - **Exclus** : dates de jalons (~90 colonnes), livrables intermédiaires PROVIDOR/ATM/EMMST/MUL, prix Formation/BTE/MCO/OBSO, LCC par sous-domaine, colonnes formules Excel (Mise en qualité, Correcteur, Date de MEQ), F206-F209 vides.

- **ProjetASL (12 → 19)** — ajouts : Sous Système, Prix Repair/Préventif/After Market Spare/After Market Repair, STBI/DAT, LCS. Refactor du MTBF intrinsèque (passe d'hérité info pure à saisie pré-remplie depuis Produit, conforme à la structure BdD). Exclus : ID Projet (ASL) auto, ISP TIS HORS FRANCE (archivé), Mise en qualité, Correcteur, Date de MEQ (formules Excel).

- **DVL HARD/SOFT (ProjetASL)** : suffixe corrigé `_PROJET ASL` (avec espace, coquille de Champs_BDD.xlsx) → `_PROJETASL` (sans espace, conforme aux autres ProjetASL).

### Diagnostic colonnes BdD vers feuille Excel
- `Test_DiagnostiquerColonnesBdDFeuille` : génère une feuille `SIRUS_Diag_Colonnes` dans `ActiveWorkbook` avec les vrais noms BdD non tronqués (colonne C format texte, largeur 80, colonne D longueur). Évite la troncature de la fenêtre Immediate de VBE.
- `Test_DiagnostiquerEcartsConfig_Toutes` reste utilisable pour le diff config↔BdD (case-sensitive).

### Convention `Open_CONN` / `Close_CONN`
Mémoire conventions VBA #1 mise à jour : ces 2 fonctions **prennent la chaîne de connexion en argument**. SIRUS utilise BdD TRAVAIL avec helper privé `ChaineConnexionSIRUS = Replace(CONNEXION_EXCEL_BDD_TRAVAIL, "UserProfile", Environ$("UserProfile"))`.

### Ambiguïtés résolues (validation Thomas)
1. **TachesATM / Critère Nm** ✅ : confirmé `CritèresDimensionnels_TACHEATM`.
2. **TachesATM / Banc test** ✅ : pas de champ Banc test — les BTE sont couverts par `OutillageStandard_TACHEATM` et `OutillageSpécifique_TACHEATM`. Mon "Banc test" remplacé par ces 2 outillages (rang 2.1/5.1 + NodeType >= SRU).
3. **EMMST** ✅ : ajout de `MoyenTransport_EMMST`, `OutillageSpecifiqueTransport_EMMST`, `Version_EMMST` au peuplement (présents dans les SQL existants mais absents de Champs_BDD.xlsx).
4. **ProjetSLI** ✅ : vrais noms BdD posés — `Réf PBS/EAP_PROJET SLI`, `Réf RAM report projet_PROJET SLI`, `LCC valeur scope total_PROJET SLI`.
5. **ProjetASL / DVL** ✅ : les 2 champs (HARD et SOFT) confirmés — une ligne ProjetASL peut être Hard et Soft.
6. **Encodage UTF-8 en clair** ✅ : Thomas confirme que les SQL ADODB en chaîne acceptent les accents directs (sans `ChrW`). Convention #6 mise à jour en mémoire pour acter l'exception ColonneBdD / chaînes SQL.

Total champs après corrections : **~98** (14 + 8 + 8 + 41 + 15 + 12).

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
