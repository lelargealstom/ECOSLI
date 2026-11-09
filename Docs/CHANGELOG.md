# Change Log — ECOSLI.xlam

Suivi des modifications apportées au complément Excel `ECOSLI.xlam` au fil des sessions de développement.

Format : `[Date] — Auteur — Périmètre`, puis liste des changements par module.

---

## 2026-11-09 — Bench v2 ventilation + implémentation cache commentaires EMMST (Plan B1 adopté) + refactor dans C_EMMST_Main

Session focus : bench v1 (2026-11-06) → validation fonctionnelle de la lecture `Workbooks.Open` (commentaires Stockage lus jusqu'à **1007 chars** vs 255 avec ADO), mais coût prohibitif (37,93 s scénario A, 453,84 s scénario B). Bench v2 pour localiser les 27 s inexpliquées. Verdict bench v2 : les 27 s sont dans la **restauration `Calculation = xlCalculationAutomatic` quand la BdD vient d'OneDrive** (A1 = 26,71 s, A2 sans restauration = 0,02 s, A3 copie locale `%TEMP%` = 0,22 s). Combo "OneDrive + restauration auto" éliminé en copiant la BdD en local avant ouverture. **Plan B1 adopté** avec cache session.

Puis refactor : à la demande de Thomas, consolidation du cache commentaires dans le module `C_EMMST_Main.bas` existant au lieu d'un module séparé `D_EMMST_CacheCommentaires.bas` (règle "garder le développement EMMST dans un seul module").

Commits : `228d76f`, `742199a`, `9c0b5e5`, `e9b60e0`.

### Commit `9c0b5e5` — Implémentation cache commentaires EMMST + injection dans EMMST_tab

Nouveau module `PROD/Modules/D_EMMST_CacheCommentaires.bas` — cache session Dictionary(DTR → Array(10 commentaires)) peuplé lazy au 1er appel de `Main_EMMST`.

**API publique** :
- `ChargerEtInjecterCommentairesEMMST()` — point d'entrée unique appelé depuis `C_EMMST_Main.RecupererDonneesBDD_EMMST` juste après `EMMST_tab = RecordsetToArray(REQSEL)`. Au 1er appel : déclenche le chargement du cache (14 s). Aux suivants : réutilise le Dictionary mémoire (0 s). Écrase ensuite dans `EMMST_tab` les 10 colonnes commentaires tronquées par ADO avec les valeurs complètes du cache. DTR absents du cache = valeurs ADO tronquées conservées (dégradation gracieuse).

**API privée** :
- `ChargerCache()` — pipeline complet en 5 étapes avec `F_Progression` :
  1. `FileCopy BdD → %TEMP%\ECOSLI_bdd_cache_<timestamp>.xlsb` (0,07 s)
  2. `Workbooks.Open ReadOnly invisible, Notify:=False, UpdateLinks:=0` (12,86 s)
  3. 10 × `Range.Value` lecture bloc + DTR (0,02 s)
  4. Indexation Dictionary(DTR → Array(10)) (~ms)
  5. `wb.Close SaveChanges:=False` + `Kill` copie (1,14 + 0,1 s)
- `TrouverColonneBdD` — `Find xlValues/xlWhole` + fallback `xlPart` (convention 2)
- `ExtraireDataSourceDepuisConnStr` — parsing de `Data Source=` dans la chaîne de connexion

**Intégration dans `C_EMMST_Main.RecupererDonneesBDD_EMMST`** : ajout d'un `Call ChargerEtInjecterCommentairesEMMST` juste après `EMMST_tab = RecordsetToArray(REQSEL)` (ligne 227). Aucune modification de la requête SQL ADO ni du pipeline existant (`Cadrage_EMMST` continue à peupler `EMMST_type` normalement).

**Barre de progression** : 5 étapes via `F_Progression` (pattern ATM V1/V2/Providor). La barre est visible dans le UserForm principal, StatusBar Excel idem, TopMost API Windows actif.

**État Application** : non touché par le module cache. `TurnOffFunctionality` (appelé dans `Main_EMMST` avant `RecupererDonneesBDD_EMMST`) a déjà positionné `ScreenUpdating/EnableEvents/Calculation/DisplayAlerts`. `TurnOnFunctionality` restaure en fin de `Main_EMMST`. Zéro impact sur le combo OneDrive-Calculation identifié par bench v2.

**Respect conventions projet** :
- Pas de handler défensif (convention 3) — erreurs remontent naturellement.
- `With ws` + `.Cells(...)` qualifiés (convention 12).
- `Find` xlValues/xlWhole + fallback xlPart (convention 2).
- ASCII pur sans accents (convention 6).
- UDT pas utilisé (Dictionary + Array Variant → OK convention 11).
- Pas d'`Attribute VB_Name` (pattern copy/paste GitHub → VBE depuis 2026-10-23).

### À tester en VBE

1. Pull du commit `9c0b5e5`.
2. Importer `D_EMMST_CacheCommentaires.bas` dans VBE.
3. Récupérer la modif d'une ligne dans `C_EMMST_Main.bas` (ou ré-importer le module).
4. Générer EMMST depuis le UserForm :
   - 1er clic : barre "Chargement commentaires EMMST" visible pendant ~14 s (5 étapes), puis pipeline EMMST normal.
   - Clics suivants (même session) : aucune pause supplémentaire.
5. Ouvrir une fiche EMMST générée pour un DTR connu avec commentaire long (ex. DTR1000140360 pour Stockage → 1007 chars attendus). Comparer avec l'ancienne version (plafonnait à 255).
6. Non-régression : un DTR absent de la BdD (nouveau DTR ajouté dans DATA_ASL mais pas encore enregistré) doit garder ses commentaires ADO tronqués, pas crasher.

### Suppression des modules de test (optionnel)

Les 3 modules de diagnostic (`X_Test_TypeGuessRows.bas`, `X_Test_OuvertureBdDMasquee.bas`, `X_Test_OuvertureBdD_V2.bas`) restent dans le repo pour traçabilité historique. Peuvent être supprimés par Thomas une fois le cache validé en prod.

### Commit `115198a` — Suppression des 3 modules de test diagnostic (session EMMST)

Thomas valide la suppression des 3 modules de diagnostic/bench créés pendant la session EMMST (Piste 1 → bench v1 → bench v2 → Plan B1 adopté) :
- `PROD/Modules/X_Test_TypeGuessRows.bas` (commit `67043bc`, 2026-11-06)
- `PROD/Modules/X_Test_OuvertureBdDMasquee.bas` (commit `a40a02c`, 2026-11-06)
- `PROD/Modules/X_Test_OuvertureBdD_V2.bas` (commit `228d76f`, 2026-11-09)

Vérification pré-suppression : les 7 fonctions définies dans ces 3 modules (`Test_TypeGuessRows_Commentaires`, `MesurerLongueursCommentaires`, `Test_OuvertureBdDMasquee_Tous`, `BenchOuvertureBdD`, `LireCommentairesEtMesurer`, `Test_OuvertureBdD_V2`, `BenchVentile`, `LireCommentairesBlocSilencieux`) ne sont référencées que dans leurs propres modules. Aucune Sub/Function externe ne les appelle. Suppression sans impact sur `Main_EMMST`, `UserForm_Initialize` ou toute autre entrée du pipeline ECOSLI.

**Préservé** : `PROD/Modules/X_Test_Appel_BDD.bas` (pré-existant à cette session, pas touché).

Traçabilité : les 3 modules restent accessibles via l'historique git sur les commits précités si besoin de rejouer un diagnostic.

### Commit `9c993b2` — Nettoyage Debug.Print bruyants au Ctrl+W (SIRUS init)

Thomas a identifié ~10 Debug.Print qui s'affichent dans la fenêtre Immediate à chaque Ctrl+W (ouverture du UserForm ECOSLI). Enchaînement d'appels : `W_AppelUsf.Appel_Usf` → `UserForm.Show` → `UserForm_Initialize` (L660) → `SIRUS_InitGeneration` (L27) → `SIRUS_RemplirListViewGeneration`, qui faisait plusieurs `Debug.Print` d'init + diag de mapping.

Suppression des 10 lignes de bruit dans `G_SIRUS_PBS.bas` :
- L155 : `Debug.Print "[DATA_ASL] Ligne d'entete utilisee : " & DATA_ASL_HeaderRow`
- L448-449 : `Debug.Print "SIRUS_RemplirListViewGeneration : …"`
- L450 : `Debug.Print "  Colonnes prix mappees (0 = introuvable) :"`
- L451-456 : 6 lignes `Debug.Print "    PrixXxx = col " & …`

**Conservés** (signaux utiles pour debug futur, non bruyants en cas normal) :
- L461 : `Debug.Print "ERREUR SIRUS_RemplirListViewGeneration : "` (branche erreur).
- L516, L629, L667 : `Debug.Print` fallback/erreur sur `ChargerSetDTR_BdD`, `ChargerRowsTableBdD_Optim`.
- L225 : `Debug.Print` conditionnel sur colonnes optionnelles manquantes (silencieux dans le cas Thomas).

**Note préservée** : le fallback `[Projet_ASL$] SELECT cible echoue (…), fallback SELECT *` (L629) s'affiche toujours au Ctrl+W. Cause pas élucidée : `Test_DiagnostiquerEcartsConfig_Toutes` rapporte 0 écart config ↔ BdD sur les 6 tables, pourtant le SELECT cible échoue à l'exécution de `ChargerRowsTableBdD_Optim` sur `Projet_ASL`. Hypothèse à explorer plus tard : typage ACE sur une colonne du config, ou nom de colonne avec un caractère spécial qui s'échappe mal dans le SQL construit. Non bloquant (le fallback `SELECT *` fait le job).

Encoding `G_SIRUS_PBS.bas` : ASCII pur (0 high bytes avant modif, 0 après). Édition via Edit tool safe. Vérification post-modif : pas de BOM ajouté, 0 FFFD, 10 Debug.Print supprimés (35 → 25).

### Commit `f3e2e87` — Fix "Utilisation incorrecte de Null" ligne 638 injection commentaires

Thomas a remonté une erreur 94 "Utilisation incorrecte de Null" à l'exécution de `ChargerEtInjecterCommentairesEMMST`, ligne 638 : `dtr = CStr(EMMST_tab(iLigne, colDTR))`. Cause : ADO retourne `Null` (pas `Empty`) pour une cellule BdD vide, et `CStr(Null)` lève erreur 94.

Correction appliquée via PowerShell UTF-8 safe (règle encoding `C_EMMST_Main.bas`) : garde la ligne protégée par un `If IsNull(...) Then GoTo SuivantLigne` + filtre `LenB(dtr) = 0` + label `SuivantLigne:` devant `Next iLigne` pour continuer la boucle proprement.

```vba
For iLigne = 1 To UBound(EMMST_tab, 1)
    If IsNull(EMMST_tab(iLigne, colDTR)) Then GoTo SuivantLigne
    dtr = CStr(EMMST_tab(iLigne, colDTR))
    If LenB(dtr) = 0 Then GoTo SuivantLigne
    If m_dictCommentaires.Exists(dtr) Then
        arrComm = m_dictCommentaires(dtr)
        For iCol = LBound(colsBdD) To UBound(colsBdD)
            EMMST_tab(iLigne, colsTab(iCol)) = arrComm(iCol)
        Next iCol
    End If
SuivantLigne:
Next iLigne
```

Le label `SuivantLigne:` sert d'`early continue` (pattern autorisé par convention 3 — pas de handler défensif mais `GoTo <label>` fonctionnel OK).

Pas de pattern similaire dans `ChargerCacheCommentaires` ligne 743 (`dtrs(iLigne, 1)`) : la source est `Range.Value` d'Excel qui retourne `Empty` (pas `Null`) pour une cellule vide, et `CStr(Empty) = ""` ne crash pas. Le filtre `LenB(dtr) > 0` existant suffit.

### Commit `559efc2` — Fix accents cassés sur noms de colonnes PBS dans C_EMMST_Main

L'édit tool Claude avait remplacé 10 caractères accentués par le replacement character U+FFFD (affiché `` dans VBE) sur 5 lignes de `C_EMMST_Main.bas`. Impact critique : les `Range("Providor[…accents cassés…]")` échouaient à la compilation/exécution, et les chemins de fichier template contenaient le dossier "Qualit" (sans accent) → `FileCopy` KO.

Lignes réparées (via PowerShell `ReadAllText` / `WriteAllText` UTF-8 sans BOM pour contourner le bug de l'édit tool sur fichiers UTF-8) :
- L175 : `construite à partir` + `désignation` (nom colonne `Providor[DESIGNATION (mise en forme…)]`)
- L176 : `SOUS SYSTÈME` + `copiés` + `classifiés` + `génériques` (nom colonne `Providor[SOUS SYSTÈME…]`)
- L216 / L219 : `Qualité` dans les 2 chemins template EMMST / PHLST
- L365 : `Sous système` dans `TabIndex.ListColumns("Sous système")`

Vérification post-fix : 0 caractères U+FFFD restants, 10 paires UTF-8 latines valides, pas de BOM ajouté (cohérent avec l'encoding du reste du fichier).

**Leçon :** `C_EMMST_Main.bas` est en UTF-8 sans BOM (comme `D_Providor_Maintenance.bas` identifié le 2026-10-28). L'édit tool Claude n'est pas safe sur ces fichiers s'il touche à des lignes contenant des accents. Pour toute modif d'une ligne accentuée en UTF-8 : utiliser PowerShell avec `[System.IO.File]::ReadAllText($path, [System.Text.Encoding]::UTF8)` + `New-Object System.Text.UTF8Encoding $false` à l'écriture, pas l'édit tool.

### Commit `e9b60e0` — Refactor : cache commentaires dans C_EMMST_Main (module unique EMMST)

Thomas a demandé de garder tout le développement EMMST dans un seul module, en ajoutant uniquement des procédures. Consolidation du module séparé `D_EMMST_CacheCommentaires.bas` (créé au commit `9c0b5e5`) dans `C_EMMST_Main.bas` :

- **Déclarations module** ajoutées en tête de `C_EMMST_Main.bas` après `Option Explicit` : `m_dictCommentaires` (Private Dictionary) + 10 constantes `N_*` (Private Const, noms des colonnes BdD).
- **4 procédures** ajoutées en fin de `C_EMMST_Main.bas` :
  - `ChargerEtInjecterCommentairesEMMST` (Public) — point d'entrée.
  - `ChargerCacheCommentaires` (Private) — 5 étapes avec `F_Progression` (renommée depuis `ChargerCache` pour éviter l'ambiguïté dans un module partagé).
  - `TrouverColonneBdD_EMMST` (Private) — suffixe `_EMMST` pour éviter toute collision même en Private.
  - `ExtraireDataSourceDepuisConnStr` (Private).
- **Suppression** de `PROD/Modules/D_EMMST_CacheCommentaires.bas`.

Aucune modification de la logique métier. Le point d'injection dans `RecupererDonneesBDD_EMMST` reste `Call ChargerEtInjecterCommentairesEMMST` juste après `EMMST_tab = RecordsetToArray(REQSEL)`.

**Règle projet actée** (ajoutée en mémoire `feedback_vba_conventions.md`) : les évolutions EMMST restent dans `C_EMMST_Main.bas` + les modules `C_EMMST_*.bas` existants, pas de nouveau module dédié sauf accord explicite.

---

## 2026-11-09 (début de session) — Bench v2 ventilation fine du coût d'ouverture BdD EMMST + 2 hypothèses testées

Session focus : le bench v1 du 2026-11-06 a validé fonctionnellement la lecture via `Workbooks.Open` (commentaires Stockage lus jusqu'à **1007 chars** vs 255 avec ADO), mais révélé un coût prohibitif — 37,93 s pour le scénario A "instance courante masquée" (dont 27,66 s inexpliquées en "Setup+Fermeture") et 453,84 s pour le scénario B "instance séparée". Scénario B écarté définitivement. Bench v2 livré pour localiser précisément les 27 s du scénario A et tester 2 pistes d'optimisation.

Commits : `228d76f`.

### Commit `228d76f` — Module de bench v2 `X_Test_OuvertureBdD_V2`

Ajout de `PROD/Modules/X_Test_OuvertureBdD_V2.bas` — bench exécutable depuis VBE, 3 scénarios séquentiels avec ventilation fine en 6 phases.

Ventilation par phase (toutes chronométrées via `Timer`) :
- Setup Application (`ScreenUpdating`/`EnableEvents`/`Calculation`/`DisplayAlerts`)
- Copie locale `FileCopy` (A3 uniquement)
- `Workbooks.Open` ReadOnly
- Lecture bloc 10 colonnes + DTR via `Range.Value` (silencieuse pour mesurer uniquement le coût brut)
- `wb.Close SaveChanges:=False`
- Restauration Application (partielle si A2) + `Kill` copie (A3 uniquement)

3 scénarios :
- **A1 — baseline** : identique au scénario A du bench v1 (instance courante masquée), mais avec ventilation fine. Permet de voir si les 27 s sont dans `wb.Close` ou dans la restauration `Calculation`.
- **A2 — sans restauration Calculation** : on laisse `Application.Calculation = xlCalculationManual` à la fin au lieu de restaurer en auto. Teste l'hypothèse "les 27 s = recalcul global déclenché par `Calculation = xlCalculationAutomatic` sur le projet utilisateur qui contient `DATA_ASL` avec formules lourdes". Filet de sécurité : remise en auto après le bench pour ne pas laisser le classeur utilisateur en manuel.
- **A3 — copie locale `%TEMP%`** : `FileCopy chemin_BdD → %TEMP%\ECOSLI_bench_bdd_<timestamp>.xlsb` avant ouverture, suppression `Kill` après. Teste l'hypothèse "les 10 s d'ouverture + N s de fermeture = latence OneDrive sur I/O réseau sur le chemin `Alstom\DIS Service VBN...`".

Respecte les conventions projet : `With ws` + `.Cells` qualifiés (convention 12), pas de handler défensif (convention 3), `Find` avec `xlValues`/`xlWhole` + fallback `xlPart` (convention 2), commentaires ASCII pur (convention 6), pas d'`Attribute VB_Name`.

**Garantie non-destructive** : `ReadOnly:=True` + `SaveChanges:=False` + `Kill` ciblé uniquement sur le fichier copié en `%TEMP%` avec nom horodaté unique.

### À tester en VBE

1. Pull du commit `228d76f`.
2. Importer `X_Test_OuvertureBdD_V2.bas` dans le VBE.
3. Ctrl+G + F5 dans `Test_OuvertureBdD_V2`.
4. Comparer la ventilation A1, A2, A3. Grille de lecture :
   - Si `Restauration app` chute à ~0 s en A2 et coûte 20+ s en A1 → coupable = recalcul auto sur `ActiveWorkbook`. Solution : garder manual pendant la lecture, restaurer auto seulement à la fin du pipeline EMMST global.
   - Si `Ouverture + Fermeture` chute drastiquement en A3 → coupable = latence OneDrive. Solution : copie locale une fois par session ECOSLI, réutilisée par toutes les générations.
   - Si A3 reste cher mais A2 soulage → combiner les 2.
   - Si les 2 restent chers → bascule Plan B2 (script PowerShell export CSV préalable).

---

## 2026-11-06 — Diagnostic troncature 255 chars commentaires EMMST (ACE OLE DB) + règle projet CLAUDE.md + bench plan B

Session focus : Thomas constate que les commentaires EMMST (`CommentaireStockage_EMMST`, `CommentaireEmballage_EMMST`, etc.) sont tronqués à 255 caractères lors de la génération EMMST. Cause : le driver `Microsoft.ACE.OLEDB.12.0` échantillonne les 8 premières lignes pour deviner le type d'une colonne ; si aucune des 8 ne dépasse 255 chars → colonne typée `Text(255)` → troncature silencieuse. Thomas exclut d'ouvrir la BdD en `Workbooks.Open` (coût perf). On reste donc sur ADO et on cherche à contourner la limite via la chaîne de connexion.

**Verdict Piste 1 : KO.** Les 5 variantes testées (`TypeGuessRows=0`, `MaxScanRows=0`, avec/sans `IMEX=1`) retournent exactement les mêmes longueurs max que la chaîne actuelle, avec plafond systématique à 255 chars sur 8 des 10 colonnes commentaires (seul `CommentaireMarquage` FR/EN n'a pas de valeurs > 255 en BdD, donc n'est pas tronqué). ACE 12 installé sur le poste Alstom ignore ces paramètres dans la chaîne de connexion (comportement documenté Microsoft). Troncature confirmée : DTR0000157911 plafonne à 255 sur 5 colonnes à la fois (pattern qui ne peut être du hasard).

**Bascule vers Plan B1** : ouvrir la BdD TRAVAIL en Excel natif, invisible, ReadOnly, uniquement pour lire les 10 colonnes commentaires via `Range.Value` (pas de limite 255), puis fermer. Bench livré pour arbitrer objectivement entre instance Excel courante masquée vs instance Excel séparée via `CreateObject`.

Commits : `67043bc`, `d1d941a`, `a40a02c`.

Création en parallèle de `CLAUDE.md` à la racine : règle "mettre à jour `CHANGELOG.md` et les docs à chaque évolution de l'outil".

### Commit `67043bc` — Module de test `X_Test_TypeGuessRows`

Ajout de `PROD/Modules/X_Test_TypeGuessRows.bas` — diagnostic autonome exécutable depuis VBE.

Contenu du module :
- `Test_TypeGuessRows_Commentaires` (Public, point d'entrée F5) — lance 5 variantes de `Extended Properties` sur la chaîne de connexion de la BdD TRAVAIL :
  - **A** `Excel 12.0 Macro;HDR=Yes;IMEX=1` (chaîne actuelle, référence)
  - **B** `+ TypeGuessRows=0`
  - **C** `+ MaxScanRows=0`
  - **D** `+ TypeGuessRows=0;MaxScanRows=0`
  - **E** sans IMEX `+ TypeGuessRows=0;MaxScanRows=0`
- `MesurerLongueursCommentaires` (Private) — scanne `EMMST$`, calcule la longueur max observée sur les 10 colonnes commentaires (5 FR + 5 EN), imprime dans la fenêtre Immediate le DTR porteur du max.
- `ForcerExtendedProperties` (Private) — substitue le bloc `Extended Properties='...'` dans la chaîne de connexion sans toucher au `Data Source=`.
- `PadRight` (Private) — utilitaire d'alignement pour la sortie Immediate.

Respecte les conventions projet : utilise `Open_CONN` / `open_REQSEL` / `Close_CONN` existants, pas de handler d'erreur défensif, commentaires français sans accents (ASCII pur).

Verdict attendu :
- Si toutes les variantes plafonnent à 255 → ACE ignore les paramètres dans la chaîne (comportement documenté Microsoft). Plan B nécessaire (ouverture BdD ou modification registre).
- Si une variante remonte > 255 → Piste 1 validée, on adopte la chaîne gagnante dans la chaîne de connexion EMMST de `Z_Variables.bas` ou dans `C_EMMST_Main.RecupererDonneesBDD_EMMST`.

### À tester en VBE (statut 2026-11-06)

1. ✅ Pull du commit `67043bc` sur poste Alstom.
2. ✅ Importer `X_Test_TypeGuessRows.bas` dans le VBE (projet `ECOSLI.xlam`).
3. ✅ Ctrl+G pour ouvrir la fenêtre Immediate, placer le curseur dans `Test_TypeGuessRows_Commentaires`, F5.
4. ✅ Lire la sortie : toutes les variantes plafonnent à 255 sur 8 colonnes.
5. ✅ Décision : Piste 1 invalidée → bascule Plan B1 (bench ci-dessous).

### Commit `a40a02c` — Module de bench `X_Test_OuvertureBdDMasquee`

Ajout de `PROD/Modules/X_Test_OuvertureBdDMasquee.bas` — bench exécutable depuis VBE pour arbitrer entre les 2 variantes de Plan B1.

Contenu du module :
- `Test_OuvertureBdDMasquee_Tous` (Public, point d'entrée F5) — lance les 2 scénarios en séquence et imprime les temps + longueurs max par colonne.
- `BenchOuvertureBdD` (Private) — paramètre `creerInstanceSeparee As Boolean` :
  - **Scénario A** (False) : ouvre la BdD dans l'instance Excel courante avec `ScreenUpdating=False`, `EnableEvents=False`, `Calculation=xlCalculationManual`, `DisplayAlerts=False`. Restauration de l'état à la fin.
  - **Scénario B** (True) : crée une nouvelle instance via `CreateObject("Excel.Application")`, `Visible=False`, mêmes optimisations. L'instance est `Quit`-tée à la fin.
  - Ouverture : `Workbooks.Open(..., UpdateLinks:=0, ReadOnly:=True, Notify:=False)`.
  - Fermeture : `wb.Close SaveChanges:=False`.
  - Chronomètre les 3 phases : ouverture, lecture, setup+fermeture.
- `LireCommentairesEtMesurer` (Private) — trouve les 11 colonnes via `Find` sur ligne 1 (pattern convention 2 : `xlValues`/`xlWhole` + fallback `xlPart`), lit chaque colonne en bloc via `Range.Value` (pattern qui serait utilisé en prod pour peupler un `Dictionary(DTR → commentaires)`), mesure la longueur max par colonne.
- `TrouverColonne`, `ExtraireDataSource`, `PadRight` (Private) — utilitaires.

Respecte les conventions projet : `Range(Cells, Cells)` qualifié via `With ws` (convention 12), pas de handler défensif (convention 3), pas d'`Attribute VB_Name` (convention copy/paste GitHub → VBE), commentaires ASCII pur sans accents (convention 6).

**Garantie non-destructive** : `ReadOnly:=True` à l'ouverture + `SaveChanges:=False` à la fermeture.

### À tester en VBE (bench Plan B1)

1. Pull du commit `a40a02c`.
2. Importer `X_Test_OuvertureBdDMasquee.bas` dans le VBE.
3. Ctrl+G + F5 dans `Test_OuvertureBdDMasquee_Tous`.
4. Lire les temps affichés pour chaque scénario + vérifier que les longueurs max dépassent maintenant 255 sur les colonnes concernées (sinon la BdD elle-même n'a pas de commentaires longs — à investiguer).
5. Grille d'arbitrage :
   - `Total < 2 s` → Plan B1 adopté directement.
   - `2 s ≤ Total < 5 s` → négociable selon fréquence d'usage EMMST.
   - `Total ≥ 5 s` → envisager Plan B2 (export CSV externe).
   - Si écart A vs B `< 0.5 s` → retenir A (instance courante, plus simple).

---

## 2026-10-30 — Providor CM temps unique selon Gate + nettoyage code mort + MsgBox RAM Report CGR consolidé

Session focus : finir d'aligner la Corrective Maintenance sur la logique "source unique selon Gate" (symétrique au PM d'il y a 2 jours), nettoyer le code mort généré par la suppression de 4 colonnes dans le template Providor côté Thomas, puis refondre l'avertissement RAM Report à CGR+ en un seul MsgBox informatif non bloquant. Puis fix d'une régression `cs_col` oubliée.

4 commits : `47556e1`, `6593efe`, `f6e83ec`, `4d03517`.

### Commit `47556e1` — CM : `RemplirTempsReparation_Corr` source unique selon Gate

Symétrique de la refonte PM `b5fce46` du 2026-10-28. Plus de hiérarchie analytique → générique → famille après CGR. Une seule colonne lue selon la Gate :
- **Avant CGR** (PGR, GO ORDER, BA, IGR) : `tTps_CorrPrev_Gen_HARD` uniquement.
- **Après CGR** (CGR, FEI, IQA) : `tTps_CorrPrev_Ana_HARD` uniquement.

Cellule `RepairTime` vide + `ColorIndex 44` si DTR absent de `tDTR`, ou si source non numérique / nulle / `0`. **Plus de commentaires Excel** (`.AddComment`) posés sur la cellule (demande explicite Thomas).

`CalculerTempsDefaut` devient orpheline (plus aucun appelant). Laissée en place par convention règle 8, même situation que `CalculerTempsFamille` déjà orpheline depuis `b5fce46`.

Volume : −60 lignes nettes dans `D_Providor_Maintenance.bas`.

### Commit `6593efe` — Nettoyage code mort après suppression de 4 colonnes template Providor

Thomas a supprimé côté template 4 colonnes du tableau structuré `Providor` :
- `Tx déffaillance générique`, `Tx déffaillance fonctionnel`
- `UNITE de la maintenance corrective`, `UNITE de la maintenance preventive`

Impact fonctionnel : **aucun**. Les 4 colonnes étaient référencées uniquement par du code mort. Nettoyage :

- **`G_SIRUS_PBS.bas`** : retrait de 4 champs du Type `DATA_ASL_Col` (`TxDef_Gen`, `TxDef_Fonct`, `UniteMaintCorrective`, `UniteMaintPreventive`) + des 4 appels `TrouverColonne` correspondants dans `MapColonnesDATA_ASL`. Ces champs n'étaient jamais lus ailleurs dans PROD.
- **`Ybis_CadrageTemplate.bas`** : **fichier supprimé entièrement**. La Sub `Cadrage_ECOSLI` qu'il contenait n'était jamais appelée.
- **`Z_Variables.bas`** : retrait du `Type ECOSLI_Col` (37 champs) et de la variable `Public ECOSLI_Type As ECOSLI_Col`. Devenus définitivement orphelins avec la disparition de `Ybis_CadrageTemplate.bas`.

Volume : −97 lignes nettes sur 3 fichiers.

### Commit `f6e83ec` — Alerte RAM Report CGR consolidée + non bloquante

Refonte du dispositif de validation RAM Report. Plus de dialogue bloquant Oui/Non, plus de MsgBox "données complètes". À partir de la Gate CGR, un seul MsgBox informatif OK (non bloquant) qui résume les 3 points annoncés par Thomas :

**`D_Providor_Main.bas`** :
- `ValiderDonneesRAMProvidor` : réduite à l'orchestration (positionne `AvantCGR` + appelle `AfficherAlerteRAMReportCGR` conditionnellement). Toujours retourne True.
- **Nouvelle Sub `AfficherAlerteRAMReportCGR`** : construit le message en 3 sections, positionne `RAMReportDATA`, affiche MsgBox `vbInformation` OK avec titre "Providor - Donnees RAM Report a CGR" :
  - **== Durées de vie limitée ==** : HARDWARE/SOFTWARE → "OK (colonne analytique utilisée)" ou "colonne analytique vide, bascule sur la colonne générique (Offre et PGR)". Même critère que le fallback effectif dans `remplissage_prev` (via `AuMoinsUneValeurNumerique`).
  - **== Durées de tâche analytique ==** : "Corrective HARDWARE : X DTR sur Y sans valeur" + "Preventive SOFTWARE : X DTR sur Y sans valeur". Si > 0, ajoute "→ renseigner les taches ATM en BdD".
  - **== MTBF ==** : "X DTR sur Y sans MTBF fonctionnel" + rappel de la hiérarchie "1) MTBF fonctionnel 2) MTBF intrinsèque 3) vide + orange".
- **Nouveau helper Private `EstValeurRenseignee(v As Variant) As Boolean`** : vrai si v est numérique non nul, ni vide, ni Null, ni chaîne vide. Utilisé pour compter les DTR sans valeur.

**`D_Providor_Maintenance.bas`** :
- `AuMoinsUneValeurNumerique` : passée Private → **Public** pour réutilisation croisée depuis `AfficherAlerteRAMReportCGR`.
- `remplissage_prev` : retrait du MsgBox local de fallback DVL (variable `msgFallback` + 2 concaténations + bloc `If MsgBox`). L'avertissement unique en entrée via `AfficherAlerteRAMReportCGR` remplace ce MsgBox local → **un seul point d'information utilisateur** dans tout le pipeline Providor.

`RAMReportDATA` est toujours positionnée (True si tout OK, False sinon) — préservée car utilisée par `B_ATM_V1.bas:824,866`.

### Commit `4d03517` — Fix : `Public cs_col As cs_col` manquant dans Z_Variables

Régression détectée par Thomas en VBE à la compilation. Le `Type cs_col` était bien déclaré depuis le commit `2fda14d` du 2026-10-29, mais la variable publique `Public cs_col As cs_col` manquait dans la section "INSTANCES PUBLIQUES DES TYPES" (lignes 687-693 de `Z_Variables.bas`). Les 7 autres `*_col` sont bien là. Ajout de la 8e ligne entre `pm_col` et `sl_col` par cohérence d'ordre.

### À tester en VBE

1. Gate avant CGR → aucun MsgBox s'affiche (comme avant la refonte).
2. Gate ≥ CGR, données complètes → MsgBox s'affiche avec toutes sections "OK (colonne analytique utilisée)" + "0 DTR sur Y sans valeur" partout.
3. Gate ≥ CGR, colonne analytique DVL HARDWARE vide → section HARDWARE "bascule sur la colonne générique (Offre et PGR)".
4. Gate ≥ CGR, quelques DTR sans tps correctif → "X DTR sur Y sans valeur. → renseigner les taches ATM en BdD".
5. Projet DTR BOTH avec `DureeCirculation = 45`, DVL HARD = 30, DVL SOFT = 10 → ligne SOFT PM porte `"3"` dans Periodic Cancelled occurences (test de non-régression commit `edceabc`).
6. `remplissage_corr` : un DTR avec `tTps_CorrPrev_Ana_HARD` vide après CGR → `RepairTime` vide + orange (plus de temps famille).

---

## 2026-10-29 (fin de journée) — Retrait de tous les handlers d'erreur défensifs dans PROD

Session focus : Thomas s'est heurté à une erreur masquée par un `ErrorHandler` dans `InitialiserCadrageColonnes`. Il demande le retrait complet de tous les handlers défensifs dans l'ensemble du dossier PROD pour que les erreurs VBA remontent naturellement à la ligne exacte (surlignage jaune VBE) — conforme à la convention projet règle 3 (inversée depuis 2026-07-20 pour faciliter le debug avec les ingénieurs).

Commit unique : `311c8bd`. **−213 lignes nettes** sur **15 modules PROD**.

### Modules modifiés et labels retirés

| Module | Labels défensifs retirés |
|---|---|
| `A_1_ModEvents.bas` | ErreurLecture |
| `A_2_ModUpdate.bas` | ErrorHandler (x2) |
| `B_ATM_Main.bas` | ErrHandler |
| `B_ATM_V1.bas` | ErrHandler |
| `C_EMMST_Main.bas` | ErrHandler (x2) |
| `D_Providor_Finalisation.bas` | InvalidValue, ErrorHandler, SafeExit, Fail |
| `D_Providor_Main.bas` | ErrHandler (x3) |
| `D_Providor_Utilitaires.bas` | ErrorHandler, ValidationFailed (cas spécial) |
| `E_Extraction_ATM.bas` | ErrorHandler (x2) |
| `E_Extraction_ATM_Tolerant.bas` | ErrorHandler (x2) |
| `E_Extraction_EMMST.bas` | ErrorHandler |
| `E_Extraction_Main.bas` | ErrorHandler (x2), ValidationFailed |
| `F_FonctionsGlobales.bas` | SiErreur, ErrHandler (x2), EH |
| `G_SIRUS_ATMAP.bas` | ErrCreate, ErrExcel, ErrRead, ErrDigest, ErrUser |
| `G_SIRUS_Analyse.bas` | ErrBdD (x2) |

### Logique de suppression (automatisée via PowerShell cp1252)

Pour chaque occurrence d'un label défensif :
1. La ligne `On Error GoTo <Label>` en tête de Sub/Function est retirée.
2. Le bloc `<Label>: ... MsgBox / Err.Raise / etc.` jusqu'au `End Sub` ou `End Function` exclu est retiré.
3. Le `Exit Sub` ou `Exit Function` immédiatement avant le label (séparé par des lignes vides) est retiré.

### Conservé (règle 3b et 3c de la convention)

- Les `On Error Resume Next` **fonctionnels** (test d'existence de feuille / clé Dictionary, CopyFile idempotent, patterns UI ListView.Clear tolérants) avec `On Error GoTo 0` juste après.
- Les `On Error GoTo 0` seuls (reset du handler).
- Les labels `Sortie:` / `SortieBdD:` utilisés comme early-exit pour nettoyage (ex. `Close_CONN` après `If REQSEL Is Nothing Then GoTo Sortie`).

### Cas spécial : `D_Providor_Utilitaires.ValiderCadrageColonnes`

Dans cette Function, le label `ValidationFailed:` est référencé par **6 `GoTo ValidationFailed` fonctionnels** dans le corps (`If bs_col.NodeLevel = 0 Then GoTo ValidationFailed`, etc.) — pas uniquement par le `On Error GoTo ValidationFailed` défensif. Suppression nuancée :
- La ligne `On Error GoTo ValidationFailed` est retirée.
- Le label `ValidationFailed:` et la ligne `ValiderCadrageColonnes = False` sont préservés pour que les 6 `GoTo` continuent de fonctionner.
- Seul le `MsgBox` dans le bloc est retiré.

### Perte intentionnelle de contexte en cleanup

- Dans `G_SIRUS_ATMAP.AttachExcelToTicket`, le bloc `ErrExcel:` faisait du cleanup (`xlWb.Close False` + `xlApp.Quit`) après l'affichage de l'erreur. Supprimé → en cas d'erreur, l'instance Excel reste en mémoire. Acceptable car l'erreur sera visible en VBE et corrigible immédiatement.
- Dans `F_FonctionsGlobales.ExecNonQuery`, le bloc `EH:` ré-émettait l'erreur avec le SQL en contexte. Supprimé → l'erreur brute remonte sans le SQL en annotation.

---

## 2026-10-29 (apres-midi) — Providor : remplissage Control Sheet ligne 19 selon langue FR/EN

Session focus : ajouter dans la feuille `Control Sheet` du Providor un remplissage automatique de la ligne 19 (sous l'en-tête ligne 18) avec un texte structuré en 6 blocs selon la langue (FR / EN), les titres en gras, la date du jour, l'utilisateur Windows et l'issue `"A1"` en dur.

Commit unique : `2fda14d`. **+162 lignes** avec un nouveau module.

### Nouveau module `PROD/Modules/D_Providor_ControlSheet.bas`

- `Public Sub remplissage_ControlSheet` : pose les 4 valeurs en ligne 19 de la feuille `Control Sheet` :
  - **Content of Modification** : texte structuré 6 blocs (General Parameters, Breakdown Structure, Material List, Tool List, Preventive maintenance, Reliability & corr maintenance) selon la langue, titres mis en gras via `.Characters(start, len).Font.Bold = True`, `WrapText = True`.
  - **Date** : `Date` (date du jour VBA).
  - **Author** : `Application.UserName`.
  - **Issue** : `"A1"` en dur.
- 2 `Private Function` `ConstruireContenuControlSheet_FR` / `_EN` qui retournent la chaîne complète. Les `"..."` et `"XXX"` sont des placeholders à compléter manuellement par l'utilisateur. L'ellipsis `…` (U+2026) est posée via `ChrW(&H2026)` dans les 2 sections General Parameters (comme dans le texte original fourni par Thomas).

### `PROD/Modules/Z_Variables.bas`

- `Public CS As Worksheet` ajoutée dans la section des feuilles globales (après `Public PM As Worksheet`).
- Nouveau `Public Type cs_col` avec 4 champs : `Content`, `datecontrolsheet` (nommé ainsi pour éviter la collision avec le mot-clé VBA `Date`), `Author`, `Issue`.

### `PROD/Modules/D_Providor_Utilitaires.bas`

- `InitialiserCadrageColonnes` : nouvelle section `CONTROL SHEET (ligne 18)` avec 4 appels `Find` sur les headers exacts `"Content of Modification"`, `"Date"`, `"Author"`, `"Issue"`. Variable `ligneteteCS = 18` déclarée en haut avec les autres `lignetete*`.

### `PROD/Modules/D_Providor_Main.bas`

- `GenererProvidor` : `Set CS = wb.Worksheets("Control Sheet")` ajouté après `Set TL`.
- `GenererProvidor` : `Call remplissage_ControlSheet` ajouté après `Call remplissage_GeneralParameters`.

### Décisions annexes

- **Alignement de ponctuation EN** : les titres EN originaux fournis par Thomas manquaient parfois le `" :"` (`Breakdown structure`, `Material List`, `Tool List`). Ajouté partout pour cohérence avec la version FR et pour que la détection par `InStr` lors de l'application du gras matche correctement les 6 titres.
- **Module séparé** : nouvelle Sub placée dans un nouveau module `D_Providor_ControlSheet.bas` plutôt que dans `D_Providor_Finalisation.bas` existant, pour isolation et pour éviter la confusion avec `remplissage_skillAndcontrolSheets` (nom historique qui concerne Skill List + Tool List, pas Control Sheet).

---

## 2026-10-29 (soir) — Providor PM : colonne Periodic Cancelled occurences pour DTR BOTH

Session focus : ajouter dans la feuille Preventive Maintenance du Providor une colonne qui liste, pour chaque DTR en `BOTH`, les occurrences de la tâche SOFTWARE qui doivent être annulées parce qu'elles coïncident temporellement avec un remplacement HARDWARE (le remplacement complet du DTR rend inutile l'upload logiciel programmé au même instant).

Commit unique : `edceabc`. Testé OK par Thomas.

### Règle métier implémentée

Pour chaque DTR `BOTH` :
- `n * DVL_SOFT ≤ DureeCirculation` → occurrences SOFT à n = 1, 2, 3...
- `m * DVL_HARD ≤ DureeCirculation` → occurrences HARD à m = 1, 2, 3...
- Une occurrence SOFT numéro `n` est **annulée** si `n * DVL_SOFT` est un multiple entier ≥ 1 de `DVL_HARD` (et ≤ `DureeCirculation`).
- `pm_col.PeriodicCancelledOcc` de la ligne SOFT reçoit la liste des `n` séparés par virgules (ex. `"3"`, `"2,4,6"`).

Exemple validé par Thomas : DureeCirculation = 45 ans, DVL HARD = 30, DVL SOFT = 10 → `"3"` posé sur la ligne SOFT (la 3e occurrence SOFT à 30 ans coïncide avec la 1re HARD à 30 ans). Les occurrences 1 (10 ans), 2 (20 ans) et 4 (40 ans) sont conservées, la 5e (50 ans) dépasse la durée.

### `PROD/Modules/Z_Variables.bas`

- `Type pm_col` : ajout `PeriodicCancelledOcc As Integer` entre `FrequencyUnit` et `TaskTime`.
- Nouvelle variable globale `Public DureeCirculation As Variant` (en années), déclarée après `Public langue As String`.

### `PROD/Modules/D_Providor_Utilitaires.bas`

- `InitialiserCadrageColonnes`, section Preventive Maintenance : nouvelle ligne de cadrage auto-`Find` sur le header `"Periodic Cancelled occurences"` (le header doit exister dans la ligne 3 de la feuille PM du template Providor PROD, après `Frequency unit`).

### `PROD/Modules/D_Providor_Main.bas`

- `RecupererDonneesPBS_Providor` : ajout de la ligne `tHARD_SOFT = DS1.Range("Providor[HARD/SOFT]")` à la fin du bloc de récupération des tableaux depuis l'onglet Providor. Jusqu'ici la variable globale `tHARD_SOFT` (déclarée `Public` dans `Z_Variables.bas`) n'était alimentée que par le pipeline ATM (`B_ATM_Main.bas`).
- `RecupererDonneesPBS_Providor` : appel de la nouvelle Sub `RecupererDureeCirculationProjet` juste après `tHARD_SOFT`.
- **Nouvelle Private Sub `RecupererDureeCirculationProjet`** : requête SQL dédiée sur `[Projet_SLI$]`/`[Duree de circulation_PROJET SLI]` filtrée par `Code_Projet` (variable globale). Stocke le résultat dans `DureeCirculation`. La connexion BdD est déjà ouverte en amont par `Btn_Providor_Click` (ligne 222 de `UserForm.frm`) avant l'appel à `Main_Providor`.

### `PROD/Modules/D_Providor_Maintenance.bas`

- `InsererLignesPreventives` : après l'appel à `InsererLigneSOFT`, nouvel appel à `PoserPeriodicCancelledOcc(LastRowInCol(PM, pm_col.TaskDescription), index, sourceHARD(index, 1), sourceSOFT(index, 1))`. On passe les 2 durées de vie pour le calcul des coïncidences.
- **Nouvelle Private Sub `PoserPeriodicCancelledOcc(lignePM, index, dvlHARD, dvlSOFT)`** :
  - Sort immédiatement si `tHARD_SOFT(index, 1) <> "BOTH"` → cellule laissée vide.
  - Sort immédiatement si l'une des 3 valeurs (`DureeCirculation`, `dvlSOFT`, `dvlHARD`) est non numérique ou = 0.
  - Boucle `n = 1, 2, 3...` tant que `n * dvlSOFT ≤ DureeCirculation`. Pour chaque `n`, teste si `instantSOFT / dvlHARD` est un entier ≥ 1 (avec epsilon 1e-7 pour les flottants, bien que les DVL métier soient entières en pratique).
  - Pose la liste des `n` annulés dans `pm_col.PeriodicCancelledOcc` séparés par virgules si au moins un `n` a été capté.

### Décisions annexes

- Approche α retenue : calcul inline dans `InsererLignesPreventives` plutôt que post-boucle. Évite d'avoir à inverser `LIST_NODE` pour retrouver le DTR correspondant à chaque ligne SOFT.
- Pas de header d'erreur défensif dans `RecupererDureeCirculationProjet` : si `Code_Projet` ne matche aucun projet dans `Projet_SLI`, `DureeCirculation` tombe à 0 → `PoserPeriodicCancelledOcc` sort proprement sans rien poser.

### À retester si refonte pipeline

- Si quelqu'un fait un `Close_CONN` entre `Btn_Providor_Click` et `remplissage_prev`, `RecupererDureeCirculationProjet` plantera. À ce stade (2026-10-29), rien ne ferme la connexion dans le pipeline Providor avant la fin.

---

## 2026-10-29 — Providor PM : Skill FR/EN + colonne Package Reference regroupée par fréquence

Session focus : finaliser l'alignement de `InsererLigneSOFT` sur la convention métier (Skill = Electrician, pas Technician), ajouter la distinction FR/EN sur le libellé Skill en Preventive (déjà présente en Corrective), et introduire une nouvelle colonne **Package reference** dans la feuille Preventive Maintenance du Providor qui regroupe les lignes par même fréquence de maintenance via une lettre (A, B, C...).

Commit unique : `4a9c0dd`.

### `PROD/Modules/Z_Variables.bas`

- `Type pm_col` : ajout du champ `PackageReference As Integer` entre `PrevMaintType` et `OnOff`.

### `PROD/Modules/D_Providor_Utilitaires.bas`

- `InitialiserCadrageColonnes`, section Preventive Maintenance : nouvelle ligne de cadrage auto-`Find` sur le header `"Package reference"` (le header doit exister dans la ligne 3 de la feuille PM du template Providor PROD).

### `PROD/Modules/D_Providor_Maintenance.bas`

- **Skill FR/EN pour les 2 types de ligne** :
  - `InsererLigneHARD` : `SkillName = IIf(langue = "FR", "Electricien", "Electrician")` (avant : `"Electrician"` en dur).
  - `InsererLigneSOFT` : même expression (avant : `"Technician"` en dur, non conforme à la convention métier).
- **Appel de la nouvelle Sub** `RemplirPackageReference` à la fin de `remplissage_prev`, juste avant `CleanUp:`, après la boucle `For i = 1 To UBound(tDTR)`.
- **Nouvelle Private Sub `RemplirPackageReference`** :
  - Parcourt la feuille PM à partir de la ligne 4 (fin détectée via `LastRowInCol(PM, pm_col.TaskDescription)`).
  - Collecte les valeurs numériques uniques de `pm_col.FrequencyValue` dans un `Scripting.Dictionary` (via `CreateObject`).
  - Trie ces fréquences par ordre croissant (tri à bulles, taille généralement petite).
  - Attribue A, B, C, ... à chaque fréquence unique dans l'ordre croissant.
  - Au-delà de 26 fréquences distinctes : AA, AB, ..., ZZ, AAA, ...
  - Lignes sans fréquence numérique exploitable (vide, non numérique) : pose `"-"`.
  - Les lignes HARD et SOFT d'un même DTR qui partagent la même fréquence partagent la même lettre (le regroupement est purement basé sur la valeur de FrequencyValue, pas sur le DTR).
- **Nouvelle Private Function `IndexToLettreExcel(n As Long) As String`** : convertit un index base 1 en lettre style Excel (1 → `A`, 26 → `Z`, 27 → `AA`, 28 → `AB`, ...).

### Observation technique

- Encodings mixtes constatés dans le projet PROD : `Z_Variables.bas` et `D_Providor_Utilitaires.bas` sont en **cp1252** (byte 0xE9 pour `é`), `D_Providor_Maintenance.bas` est en **UTF-8** sans BOM (bytes 0xC3 0xA9 pour `é`). Script d'édition PowerShell adapté par fichier.

### À tester en VBE

1. `langue = "FR"` → toutes les lignes PM HARD et SOFT doivent avoir `SkillName = "Electricien"`.
2. `langue = "EN"` → `"Electrician"` partout.
3. Projet avec 3 fréquences distinctes (ex. 5, 10, 20) sur plusieurs DTR → Package reference : `A` pour les lignes @ 5, `B` pour les lignes @ 10, `C` pour les lignes @ 20. Un même DTR BOTH avec HARD @ 20 et SOFT @ 10 doit porter `B` sur sa ligne HARD et `A` sur sa ligne SOFT.
4. DTR avec fréquence vide / non numérique → `"-"` dans Package reference.

---

## 2026-10-28 — Simplification règles Providor Preventive et Corrective Maintenance

Session focus : nettoyer les règles de remplissage des feuilles **Preventive Maintenance (PM)** et **Corrective Maintenance (CM)** du livrable Providor côté **PROD**. Objectif : supprimer les hiérarchies de fallback internes au profit de règles claires pilotées par la Gate, et aligner la logique PM/CM avec la vraie attente métier (plus de fallback générique silencieux après CGR).

Tous les changements concernent `PROD/Modules/D_Providor_Maintenance.bas`.

### Commit `62acd59` — PM : fallback générique si colonne CGR vide + fix needSOFT

- `remplissage_prev` branche **après CGR** : si la colonne analytique durée de vie `tDureevie_Ana_HARD` (resp. `tDureevie_Ana_SOFT`) ne contient **aucune valeur numérique** sur toute la plage DTR, bascule automatique sur `tDureevie_Gen_HARD` (resp. `tDureevie_Gen_SOFT`). Les deux bascules sont indépendantes. Un MsgBox d'avertissement consolidé liste quelles sources ont basculé.
- Vient en complément du dialogue amont `ValiderDonneesRAMProvidor` (inchangé) — l'utilisateur peut avoir répondu "Oui, continuer avec les données intrinsèques" et bénéficier quand même du fallback automatique colonne par colonne.
- Les temps de réparation NE sont PAS concernés par ce fallback : `sourceTpsHARD = tTps_CorrPrev_Ana_HARD` et `sourceTpsSOFT = tTps_Prev_Ana_SOFT` conservés tels quels dans la branche après CGR.
- **Correction indépendante** : `InsererLignesPreventives.needSOFT` → retrait de la condition `And tNTI(index - 1, 1) = 3`. Test désormais identique à `needHARD` : une ligne SOFT est insérée dès que la durée de vie SOFT est numérique et ≠ 0, sans contrainte de NodeType.
- Nouveau helper privé `AuMoinsUneValeurNumerique(ByRef t As Variant) As Boolean` : teste si un tableau 2D (nx1) contient au moins une valeur numérique. Un `0` compte comme renseigné, les cellules vides et chaînes vides non.

### Commit `b5fce46` — PM : RemplirTempsReparationPrev — source unique selon Gate + HARD/SOFT

- Suppression de la hiérarchie à 3 étages (`sourcePrimaire` → `sourceSecondaire` → `CalculerTempsFamille`) dans `RemplirTempsReparationPrev`.
- Signature passée de 4 → 3 paramètres : `(ligne, index, source)`. La source est désormais déterminée **une seule fois en amont** dans `remplissage_prev` selon Gate + HARD/SOFT.
- Corps réduit à un seul `If/Else` :
  - Si `source(index, 1)` est numérique, non vide, `<> "#N/A"`, `<> 0` → pose le temps + commentaire `"Tps Generique"` (avant CGR) ou `"Tps Analytique"` (après CGR).
  - Sinon → cellule PM vide, commentaire `"Aucun temps de reparation renseigne"`, `ColorIndex = 44`.
- Appels dans `InsererLigneHARD` (ligne ~561) et `InsererLigneSOFT` (ligne ~591) alignés : 4e argument supprimé (`tTps_CorrPrev_Gen_HARD` / `tTps_Prev_Gen_SOFT`).
- **Code mort** : `CalculerTempsFamille` (lignes ~674-699) n'a plus aucun appelant. Non supprimée (règle 8 : pas de refactoring hors demande explicite). À trancher dans une session ultérieure.

### Commit `feffef6` — PM : retrait suppression dernière ligne en fin de `remplissage_prev`

- Suppression du bloc final `If dlPM > 3 Then PM.Cells(dlPM, pm_col.NodeIndex).EntireRow.Delete`. Résidu de l'ancien algorithme qui ajoutait systématiquement une ligne vide à nettoyer en fin de boucle. Dans la version refactorisée où chaque DTR insère 0, 1 ou 2 lignes selon ses propres conditions, plus aucune ligne vide à supprimer. `remplissage_prev` se termine directement par `CleanUp:`.

### Commit `aed7869` — CM : retrait fallback MTBF générique après CGR

- `RemplirMTBF_Corr` branche **après CGR** : suppression du 3e étage de la hiérarchie (`tMTBF_gen`). La sélection se limite à :
  1. `tMTBF_fct` (MTBF fonctionnel du RAM Report)
  2. `tMTBF_Intr` (MTBF intrinsèque du RAM Report)
  3. Sinon : cellule vide + `ColorIndex = 44` (orange) + `MTBFSource = "Pas de MTBF trouve en BDD"`
- Avant CGR (PGR / GO ORDER / BA / IGR) : comportement **inchangé**, toujours `tMTBF_gen` uniquement.
- Commentaire d'en-tête de la Sub + commentaire du bloc `APRES CGR` alignés sur la nouvelle logique (ligne header réécrite en ASCII conforme convention projet : `' APRES CGR : HIERARCHIE FONCTIONNEL -> INTRINSEQUE`).

### Observation technique

- Le fichier `D_Providor_Maintenance.bas` est en **UTF-8 sans BOM**, pas en cp1252 comme la mémoire des conventions l'indiquait. Les 2 premiers commits de la session touchaient uniquement des portions ASCII donc passaient en lisant via `GetEncoding(1252)` byte-par-byte. Le dernier commit (modifications de lignes contenant des accents dans `RemplirMTBF_Corr`) a dû passer par `[System.Text.Encoding]::UTF8` en lecture et `UTF8Encoding $false` en écriture. À confirmer si l'ensemble des `.bas` PROD partagent le même encoding UTF-8 — si oui, la convention projet est à mettre à jour.

### À tester en VBE

1. **PM fallback** : générer un Providor sur un projet Gate ≥ CGR avec `tDureevie_Ana_HARD` vide et `tDureevie_Ana_SOFT` renseignée → MsgBox affiche uniquement la bascule HARD, et les lignes PM HARD prennent les valeurs génériques pendant que les lignes SOFT utilisent l'analytique.
2. **PM needSOFT** : un DTR avec `tDureevie_SOFT` numérique ≠ 0 et NTI ≠ 3 doit maintenant produire une ligne SOFT (avant il n'en produisait pas).
3. **PM temps** : un DTR avec `tTps_CorrPrev_Ana_HARD` vide après CGR doit afficher une cellule TaskTime vide + orange + commentaire `"Aucun temps de reparation renseigne"` (plus de fallback vers temps famille 0.3h / 0.45h).
4. **CM MTBF** : un DTR après CGR avec `tMTBF_fct` et `tMTBF_Intr` vides doit afficher cellule MTBF vide + orange + `MTBFSource = "Pas de MTBF trouve en BDD"` (plus de fallback vers `tMTBF_gen`).

---

## 2026-10-23 — Port matrice MeQ Produit v0.2 vers PROD + alignement BdD sans accents

Session focus : porter la **matrice MeQ Produit v0.2** (développée en PREPROD en juillet 2026) côté **PROD** pour que le MeQ affiché dans la ListView Generation au démarrage du UserForm soit calculé selon les règles v0.2 (dénominateur variable 3→19 champs selon NodeType + MotifPreventif) au lieu des règles v0.1 ad-hoc (dénominateur bloqué à 6). Alignement du config SIRUS sur la BdD TRAVAIL migrée **sans accents ET sans apostrophe `'`** côté Alstom. Nettoyage headers VBE pour permettre le copier-coller GitHub → Editor VBA.

### Commit `9478f40` — Port MeQ v0.2 PREPROD → PROD + headers VBE copiables

- **`PROD/Modules/G_SIRUS_Config.bas`** :
  - Ajout `Public Const SIRUS_CONFIG_VERSION As String = "2026-10-23-v4"` (bump) + helper `SIRUS_ColonneBdDClePourTable` + fonction `SIRUS_VerifierConfigAJour` (rebuild auto si version obsolète).
  - Écriture `ws.Range("O1").Value = SIRUS_CONFIG_VERSION` dans `SIRUS_InitConfigSheet` pour permettre la détection.
  - Nouveau `PeuplerProduit` : matrice v0.2 (25 champs, dont **4 virtuels** `Source = "SnapShotPrix"` alimentés depuis `Produit_HistoPrix`). Noms BdD alignés sur schéma 07/2026 (`Designation_PRODUIT`, `MotifNonApplicabilite_PRODUIT`, `Motif_NonApplicabilitePreventif_PRODUIT`, `Prix SPARE catalogue a date_PRODUIT`, etc.).
- **`PROD/Modules/G_SIRUS_Regles.bas`** :
  - Routing `If sTable = "Produit" Then SIRUS_RegleObligatoire = EstObligatoire_Produit(sColonneBdD, ligne)` en tête de `SIRUS_RegleObligatoire`.
  - Ajout `EstObligatoire_Produit` : 3 base + 16 conditionnels avec exemptions DVL HARD/SOFT selon MotifPreventif et exclusions REPAIR si NodeType LRUNR/SRUNR.
  - Retrait des 2 Cases Produit obsolètes v0.1 (MTBF + MotifNA).
  - Remplacement `Test_Regle_Produit_MTBFConditionnel` par `Test_Regle_Produit_MatriceMeQ` qui vérifie les 9 cas (3/14/15/16/17/18/19 champs obligatoires).
- **`PROD/Modules/G_SIRUS_Analyse.bas`** :
  - Ajout `Public Function SIRUS_MeQLigne(ligneDict, sTable, configCol) As Object` : helper unifié de comptage MeQ (source unique de vérité). Formule `nbRempli / nbObl`, poids uniforme, skip Auto/Herite + placeholders `??_*_a_definir`.
- **`PROD/Modules/G_SIRUS_PBS.bas`** :
  - `CalculerMeQDTR_Optim` refactoré : délègue au helper `SIRUS_MeQLigne` au lieu de refaire le comptage inline. Signature identique.
- **`PROD/Forms/UserForm.frm`** :
  - Ajout `Call G_SIRUS_Config.SIRUS_VerifierConfigAJour` avant `SIRUS_InitGeneration` dans `UserForm_Initialize`. Rebuild silencieux de `SIRUS_Config` au premier Ctrl+W après bump.
- **Headers VBE copiables** :
  - Retrait de la ligne `Attribute VB_Name = "..."` sur **40 .bas** de `PROD/Modules/` (PREPROD l'avait déjà fait).
  - Retrait du bloc header `VERSION 5.00` + `Begin {GUID} ... End` + `Attribute VB_*` sur les **4 .frm** (`UserForm.frm` et `Barre_Progression.frm` × PROD + PREPROD). Les .frm démarrent maintenant par `Option Explicit` ou `Sub afficher()`.

### Commit `57f6403` — Alignement PeuplerXxx + Regles.bas sur BdD sans accents (v5)

Suite au nettoyage BdD TRAVAIL côté Alstom (toutes colonnes passées sans accents), alignement des `PeuplerXxx` hors Produit et des Cases TachesATM dans `G_SIRUS_Regles` sur les nouveaux noms. Objectif : éliminer les ~30 messages `[Table$] SELECT cible echoue (Aucune valeur donnée pour un ou plusieurs des paramètres requis.), fallback SELECT *` au démarrage du UserForm.

- **`G_SIRUS_Config.bas` (PROD + PREPROD)** : 30 renommages `ColonneBdD` (accents → sans accent).
  - InfoATM (1), TachesATM (2), EMMST (14), ProjetSLI (10), ProjetASL (5).
  - + SourceHeritage MTBF intrinsèque Produit sans accent (pointeur vers nouveau nom Produit).
- **`G_SIRUS_Regles.bas` (PROD)** : 5 Cases/GetVal (retrait `ChrW(232)`/`ChrW(233)`).
  - `CriteresDimensionnels_TACHEATM`, `OutillageSpecifique_TACHEATM`, `NbOperateurs_TACHEATM`, `MTBF intrinseque Produit...`, tests Torque Tighten + alerte MTBF.
- **Bump `SIRUS_CONFIG_VERSION`** : `2026-10-23-v4` → `2026-10-23-v5`.
- **Encoding `windows-1252` préservé byte-à-byte** via PowerShell `[System.Text.Encoding]::GetEncoding(1252)` + `ReadAllBytes/WriteAllBytes` : l'Edit tool corrompt les accents cp1252 (byte `0xE9` → `0xEF 0xBF 0xBD` UTF-8 replacement). **Règle structurante pour toute édition de .bas/.frm latin-1 avec accents** : utiliser PowerShell, pas Edit.

### Commit `7cf3862` — Nettoyage final 5 absences diag config vs BdD TRAVAIL (v6)

Dernières absences remontées par `Test_DiagnostiquerEcartsConfig_Toutes` après `57f6403`.

- **`G_SIRUS_Config.bas` (PROD + PREPROD)** :
  - Produit `Ref RAM Report_PRODUIT` → `Reference RAM report produit dans AEC_PRODUIT`.
  - ProjetSLI `Nombre de DTR dans l'EAP ou la PBS` → ` l EAP ou la PBS` (**apostrophe ASCII `0x27` retirée**, remplacée par espace — confirmation que la BdD retire aussi les apostrophes, pas seulement les accents).
  - ProjetSLI `InterlocuteurLU_PROJET SLI` : **retiré du config** (colonne inexistante en BdD).
  - ProjetASL `Prix spare (DEVISE du projet (SLI))_PROJETASL` → `ISP TIS HORS FRANCE (archive)_PROJETASL` (renommage 2026-07-28 confirmé).
  - ProjetASL `Testabilite (... relatif a l'equipement ...)` → apostrophe retirée.
- **Bump `SIRUS_CONFIG_VERSION`** : v5 → **v6**.

### Impact MeQ Produit (vérifié par `Test_Regle_Produit_MatriceMeQ`)

Pour un même DTR, le dénominateur passe de 6 (règles v0.1 ad-hoc) à 14-19 champs selon le profil :

| MotifNA / MotifPrev / NodeType | PROD avant | PROD après |
|---|---:|---:|
| NA / * / * | 6 | **3** |
| OK / HARD / LRU | 6 | **18** |
| OK / HARD / LRUNR | 6 | **15** |
| OK / BOTH / LRU | 6 | **19** |
| OK / BOTH / LRUNR | 6 | **16** |
| OK / NA / SRUNR | 6 | **14** |

### Diag final

`Test_DiagnostiquerEcartsConfig_Toutes` après le port + les 2 alignements (et après que Thomas a retiré le `à` restant de Testabilite côté BdD) : **0 absent sur toutes les tables** (Produit 24, InfoATM 8, TachesATM 8, EMMST 41, ProjetSLI 54, ProjetASL 19).

### Convention mise à jour

- `feedback_vba_conventions.md` règle 6 : BdD TRAVAIL migrée sans accents **ET sans apostrophe** depuis 2026-10-23. Les `ColonneBdD` du config SIRUS doivent être sans accent et sans apostrophe en clair (seule exception : ellipsis U+2026 via `ChrW(&H2026)`).
- Règle structurante ajoutée : pour toute édition .bas/.frm cp1252/latin-1 avec accents, utiliser PowerShell avec encoding forcé windows-1252 (l'Edit tool réencode en UTF-8 et perd les accents en cp1252 → caractères de remplacement `0xEF 0xBF 0xBD`).

### Chantier identifié pour sessions suivantes

- **Validation terrain MeQ Produit v0.2 en PROD** : re-tester les 9 cas de la matrice sur des DTR réels (pas seulement sur Dictionary synthétique de `Test_Regle_Produit_MatriceMeQ`). Comparer MeQ ListView Generation vs MeQ attendu.
- **Évaluation d'impact pour migration BdD vers SharePoint REST** (Étape 1 du handoff `HANDOFF_ECOSLI_CONNEXION_BDD_SHAREPOINT.md`) : inventaire fait côté dev, **16 points PROD + 32 points PREPROD** à adapter. Charge estimée **~12h / 1,5 j**. 10 questions à trancher avec Thomas avant Étape 2 (URL exacte SharePoint PROD vs PREPROD, délai de contrôle, fallback hors ligne, etc.).

---

## 2026-09-21 / 2026-09-23 — Split du dépôt PROD / PREPROD

Session focus : séparer proprement le complément **actif en prod** (ASL, template ECOSLI.xlsm avec formules intégrées) du complément **en préprod** (SIRUS intégré, template `ECOSLI_DATA_ASL.xlsx` sans formules). Objectif : itérer sur PREPROD sans casser la version utilisée en prod, et pouvoir déposer les deux codes distinctement depuis les deux PC de dev (perso + Alstom).

### Nouvelle structure du dépôt

```
ECOSLI/
├── Champs_BDD.xlsx        (partagé, gitignored)
├── README.md, GITHUB.md   (partagés)
├── Docs/                  (partagé)
├── Templates/             (livrables partagés)
│   ├── ATM_FR-Template VBN v2.xlsx
│   ├── ATMC2.0_{EN,FR}-Template.xlsx
│   ├── MTA_EN-Template VBN v2.xlsx
│   ├── Template EMMST-FR_V02.xlsx
│   ├── Template PHLST-ENG_V02.xlsx
│   ├── Providor v5.021.xlsb
│   └── logo_alstom.png
├── PROD/
│   ├── Classes/ Forms/ Modules/   (code deposé depuis PC Alstom le 2026-09-23)
│   └── Templates/ECOSLI.xlsm       (gitignored, non tracké)
└── PREPROD/
    ├── Classes/ Forms/ Modules/    (code SIRUS + Template ex-racine)
    └── Templates/
        ├── ECOSLI_DATA_ASL.xlsx
        ├── _build_template.ps1
        └── _patch_template_v2.ps1
```

### Décisions structurantes

- **Livrables métier partagés** (ATM/ATMC/MTA/EMMST/PHLST/Providor + logo) : une seule copie à la racine `Templates/`, pas de duplication.
- **Templates ECOSLI dispatchés** : `ECOSLI.xlsm` (avec formules) côté PROD, `ECOSLI_DATA_ASL.xlsx` (sans formules) côté PREPROD.
- **`ECOSLI_newtemplate.xlsx` supprimé** : ancienne itération de la nouvelle version, remplacée par `ECOSLI_DATA_ASL.xlsx`.
- **`ECOSLI.xlsm` racine supprimé** : copie de travail obsolète.
- **`.gitignore`** : la règle existante `ECOSLI.xlsm` (sans path) s'applique donc à `PROD/Templates/ECOSLI.xlsm` — non tracké, vit uniquement en local sur chaque PC.
- **`Docs/`** : reste partagé (référentiel documentaire commun aux deux versions).
- **`Docs/MIGRATION_PATHS_SHAREPOINT.md`** : ajouté (inventaire pour migration BdD SharePoint, en attente, non exécutée).

### Commits

- `35f8a9f` — Split repo PROD / PREPROD : déplacement code SIRUS + dispatch templates (53 fichiers, renames 100% détectés).
- `12c5268` — Ajout dossiers vides PROD/{Classes,Forms,Modules} via `.gitkeep` pour visibilité GitHub avant dépôt Alstom.
- Dépôt du code PROD depuis PC Alstom le 2026-09-23 (commit séparé depuis ce PC).

### Impact sur les sessions futures

- Tout nouveau dev SIRUS ou template sans formules → `PREPROD/`.
- Fix urgent sur la version en prod → `PROD/`.
- Les chemins cités dans les mémoires / docs pointent désormais vers `PREPROD/Modules/...` (pas racine).
- `Docs/MIGRATION_PATHS_SHAREPOINT.md` reste valide mais ses paths sont préfixés `PREPROD/` — à mettre à jour le jour où la migration SharePoint sera lancée.

---

## 2026-09-16 — Fix Providor Maintenance Corrective : Material Name LRUNR/SRUNR + tampons non nettoyés

Session focus : deux bugs consécutifs dans `D_Providor_Maintenance.remplissage_corr`, remontés par Thomas depuis la prod ASL.

### Bug 1 — Material Name `_REPAIR` collé sur des non-réparables

**Symptôme** : dans la feuille Corrective Maintenance générée, la colonne `Material Name` reçoit le suffixe `_REPAIR` pour des DTR marqués `SRUNR` (non-réparables). Métier veut : **aucun** matériel (ni Name ni Qty) sur les lignes `LRUNR` / `SRUNR`.

**Cause** : `RemplirMaterialName` teste `TamponNodeLevel = 3` pour appliquer `_REPAIR`. Or `NodeLevel = 3` regroupe `SRU` (réparable) ET `SRUNR` (non-réparable) — cf. `F_FonctionsGlobales.NodeTypeToLevel`. Le niveau numérique n'est pas discriminant.

**Fix** — ajout d'un 4ème tampon `TamponNodeType` porté depuis BS :

- **`Z_Variables.bas` `Type cm_col`** : ajout du champ `TamponNodeType As Integer` (après `TamponNodeName`).
- **`D_Providor_Utilitaires.bas` `InitialiserCadrageColonnes`** (bloc `With CM`) : `cm_col.TamponNodeType = cm_col.FailureName + 4`. Consomme la colonne CM à `FailureName + 4` (vérifiée vide dans le template Providor v5.021).
- **`D_Providor_Maintenance.bas` `CopierDonneesBS_vers_CM`** : recopie `BS.Cells(ligneSrc, bs_col.NodeType)` vers `CM.Cells(ligneDst, cm_col.TamponNodeType)`.
- **`D_Providor_Maintenance.bas` `RemplirMaterialName`** : early-exit `If sType = "LRUNR" Or sType = "SRUNR" Then Exit Sub` avant la logique `_REPAIR`. `sType = UCase$(Trim$(...))` pour robustesse. Aucun ChrW nécessaire (tags ASCII pur).

Décision **pas de refactor** du numérique vers le tag partout : `TamponNodeLevel` reste utilisé aux l. 43, 139, 248, 253, 273 (`> 2`, `<> 2`, recopie `MaintenanceLevel`) où la sémantique numérique est correcte. Ajout ciblé du 4ème tampon = patch minimal, pas de régression sur le reste.

### Bug 2 — Colonnes tampons non nettoyées quand on génère "d'un coup"

**Symptôme** : après le fix du bug 1, quand Thomas lance la génération complète du Providor (pipeline BS + CM + PM + ...), les colonnes tampons `TamponNodeLevel..TamponNodeType` restent visibles dans la feuille CM. Quand on lance `remplissage_corr` isolément, elles sont nettoyées normalement.

**Cause racine** : les 2 `ClearContents` finaux utilisent le pattern `CM.Range(Cells(a, b), Cells(c, d))` où les `Cells(...)` internes ne sont **pas qualifiés** avec la feuille. En VBA, `Cells(...)` non qualifié réfère à `ActiveSheet`. Quand `remplissage_corr` s'exécute isolément, ActiveSheet == CM par accident et ça marche. Dans le pipeline complet, ActiveSheet a bougé (feuille précédente du pipeline), donc `CM.Range(...)` reçoit deux adresses issues d'une autre feuille — erreur silencieuse masquée par le `On Error Resume Next` défensif en tête de sub.

**Fix** — qualifier les `Cells(...)` internes avec la feuille correspondante :
- **`remplissage_corr` l. 58, 79** (2 `ClearContents` tampons) : `CM.Range(CM.Cells(...), CM.Cells(...))`.
- **l. 68-70** (List Node BS — 3 lignes) : `BS.Range(BS.Cells(...), BS.Cells(...))`.
- **l. 73-76** (Node Index CM — 4 lignes) : `CM.Range(CM.Cells(...), CM.Cells(...))`.

Bugs latents identiques sur les 3 autres blocs, corrigés dans la foulée (même pattern, mêmes conditions de déclenchement).

**Note convention** : le `On Error Resume Next` en tête de `remplissage_corr` (l. 21) est un handler défensif qui a masqué ces bugs pendant plusieurs générations. À retirer une fois la stabilisation confirmée en prod ASL (cf. règle projet — pas de handler défensif, sauf `On Error Resume Next` fonctionnels ponctuels). Idem `remplissage_prev` en dessous.

### Recette

Ouvrir un projet avec au moins 1 `SRU` + 1 `SRUNR` de rang 3 → lancer la génération Providor complète → dans la feuille CM :
- ligne SRU → `MaterialName = <NodeName>_REPAIR`, `MaterialQty = 1`
- ligne SRUNR → `MaterialName` vide, `MaterialQty` vide
- colonnes tampons (5 cols après `FailureName`) → toutes vides en fin d'exécution

---

## 2026-07-28 — Livrable client MeQ Produit + fix format prix Refresh DATA_ASL

Session focus : produire un livrable HTML autonome présentant l'algorithme de calcul MeQ de la table Produit à un public d'ingénieurs métier (non-gouvernance BdD). Puis corriger un bug d'affichage des prix dans le Refresh DATA_ASL.

### Fix — Renommage BdD Prix Spare archive (ISP TIS HORS FRANCE)

**Contexte** : Thomas a renommé côté BdD la colonne "Prix Spare archive" de `Projet_ASL$`. Nouveau nom réel : **`ISP TIS HORS FRANCE (archive)_PROJETASL`** (avec suffixe `_PROJETASL` conforme à la convention de la table).

**Correctif appliqué à tout le codebase** :
- `G_PBS_Refresh.bas` sub `RafraichirPrixArchive` (l. 689, 705, 721) — commentaire header + SELECT + `REQSEL.Fields`.
- `G_SIRUS_Config.bas` l. 563 — `AjouterChamp("ProjetASL", "Prix Spare archive", ...)` : `ColonneBdD` alignée. Le libellé humain "Prix Spare archive" est conservé (côté MeQ / config, décision à confirmer plus tard si Thomas veut un renommage sémantique aligné sur "ISP TIS").
- `G_SIRUS_Generation.bas` l. 651 — `TrouverColParColonneBdD` de `SIRUS_EnrichirProjetASL_DepuisDATA_ASL`.
- `Docs/SIRUS_REGLES.md` l. 59 — tableau ProjetASL.

**Bonus** — retrait du `On Error Resume Next` défensif dans `AppliquerFormatPrix` (règle 3 conventions VBA — permet à Thomas de voir remonter l'origine exacte si une colonne prix disparaît du template).

### Fix — Nom colonne BdD Node Type Produit (ellipsis Unicode)

**Symptôme** : le bouton `Btn_Rafraichir_DataPBS` levait l'erreur ADODB `'  ' n'est pas un nom valide. Assurez-vous qu'il ne comprend pas de caractères ou de ponctuation incorrects et qu'il n'est pas trop long.`

**Cause** : même racine que le fix EUR ci-dessous. Le nom réel de la colonne BdD est `Node type Produit (LRU/SRU…)_PRODUIT` avec un ellipsis Unicode `…` (U+2026), pas trois points ASCII `...`. Le fichier `.bas` étant ASCII pur, le caractère avait été perdu et remplacé par `...`. ACE OLE DB ne trouvait donc pas la colonne dans le SELECT.

**Fix** :
- `G_PBS_Refresh.bas` sub `RafraichirDepuisProduit` (l. 212-265) : local `sColNodeProd = "Node type Produit (LRU/SRU" & ChrW(&H2026) & ")_PRODUIT"`, utilisé dans le SELECT et le `GetFieldSafe`.
- `G_SIRUS_Regles.bas` `SIRUS_RegleObligatoire` case Outillage Standard/Spécifique (l. 80-88) : idem pour le `GetVal`. Bug silencieux corrigé — la règle renvoyait systématiquement `False` parce que la clé du dict ne matchait pas, donc les champs `OutillageStandard_TACHEATM`/`OutillageSpecifique_TACHEATM` n'étaient jamais marqués obligatoires sur les tâches rang 2.1/5.1.

**Note** : `G_SIRUS_Config.bas` l. 551 contient aussi `LRU/SRU...` dans le champ `SourceHeritage` de ProjetASL.NodeType. Ce champ est stocké mais jamais lu ailleurs (métadonnée décorative pour l'instant) — pas de fix nécessaire. À harmoniser lors du chantier futur d'héritage cross-table.

### Fix — Format prix DATA_ASL Refresh (symbole EUR)

**Symptôme** : après un Refresh, les 10 colonnes de prix de DATA_ASL (Z, AA, AB, AE, AH, AJ, AL, AM, AN, AO) affichaient `1 234,56 ?` au lieu de `1 234,56 €`.

**Cause** : le fichier `Modules/G_PBS_Refresh.bas` est en ASCII pur. Le symbole `€` (0x80 en Windows-1252 / U+20AC en Unicode) présent dans la chaîne de format `fmt` de `AppliquerFormatPrix` avait été perdu à un moment (probable export/import avec encoding incompatible) et remplacé littéralement par `?` (0x3F). `AppliquerFormatPrix` étant appelée à la fin de `RafraichirCalculs`, elle **écrase** systématiquement le format des colonnes prix à chaque Refresh — donc corriger le template Excel n'aurait rien changé.

**Fix** : dans `G_PBS_Refresh.bas` sub `AppliquerFormatPrix` (l. 942-970), introduction d'un local `sEuro = ChrW(&H20AC)` et recomposition de `fmt` par concaténation. Le fichier `.bas` reste 100 % ASCII, insensible à toute conversion d'encoding future. Même convention que `ChrW(&H2026)` utilisé pour l'ellipsis du NodeType Produit dans `G_SIRUS_Regles.EstObligatoire_Produit`.

### Livrable client HTML

Session focus : produire un livrable HTML autonome présentant l'algorithme de calcul MeQ de la table Produit à un public d'ingénieurs métier (non-gouvernance BdD).

### Livrable créé

- **`Docs/SIRUS_MeQ_Produit.html`** : document standalone (aucune dépendance JS/CDN, ouvrable double-clic hors ligne, imprimable A4 via `@media print`).

### Contenu du livrable

- **Formule MeQ** : `MeQ = champs remplis / champs attendus × 100`, poids uniforme (1 champ = 1 point).
- **Pilotes du contexte** : `Motif NA`, `Motif Préventif`, `Node Type` — les 3 variables qui font varier le dénominateur.
- **Arbre de décision SVG inline** — cœur du document :
  - Racine "Fiche Produit" en haut.
  - 3 niveaux de questions : `Motif NA ?` (2 branches) → `Motif Préventif ?` (4 branches) → `Réparable ?` (2 branches ×4).
  - **9 feuilles terminales** en base de l'arbre avec le nombre de champs attendus en gros + contexte lisible sous le nombre.
  - Feuille min (3, orange) et feuille max (19, vert) mises en évidence ; 7 feuilles neutres.
- **Tableau récap des champs** : 19 champs comptés répartis en 5 blocs (Base, Identification, Cycle préventif, Prix Spare, Prix Repair) + ligne grise des 6 champs libres exclus du MeQ.

### Cible / ton

- Public : ingénieurs métier qui ne connaissent pas la gouvernance BdD.
- Ton pédagogique, vocabulaire explicité (DTR, LRU, SRU, LRUNR, MTBF, DVL, MoQ) dans la 1re itération puis réduit à l'essentiel dans la version finale simplifiée.
- Zéro jargon SIRUS/ATMAP dans le corps du document (mention uniquement en footer comme source de vérité).

### Itérations

- **v1** : version détaillée (5 blocs colorés, encart vocabulaire, encart SnapShotPrix, exemple explicite, résumé — ~500 lignes).
- **v2 (retenue)** : simplification demandée par Thomas — focus sur l'algorithme, arbre SVG au centre avec les 9 cas terminaux en base, reste réduit au minimum (~250 lignes).

### Alignement avec la source de vérité

Livrable construit à partir de `Docs/SIRUS_REGLES.md` (section Table Produit) et vérifié contre `G_SIRUS_Regles.EstObligatoire_Produit` (7 étages de décision). Aucune modification de code VBA cette session.

---

## 2026-07-08 — Sprint 2 bis : Consolidation MeQ + ADODB pur SnapShotPrix

Session focus : rendre la formule MeQ identique entre la ListView Generation et l'analyse SIRUS feuille pour un même DTR + basculer la lecture SnapShotPrix sur ADODB pur (BdD fermée) après ajustements côté BdD.

### Consolidation MeQ (helper unique `SIRUS_MeQLigne`)

**Problème identifié** : divergence entre 2 chemins de calcul MeQ.
- `G_SIRUS_PBS.CalculerMeQDTR_Optim` (ListView Generation) → compte 1-champ = 1-unité (`nbRempli / nbObl`).
- `G_SIRUS_Analyse.EvaluerChampLigne` (analyse feuille) → utilise `PoidsMeQ` cumulé (`poidsRempli / poidsObl`).

Avec `PoidsMeQ` non uniforme dans le config Produit (DTR=3, Designation=2, MTBF=2, autres=1), le taux MeQ pour un même DTR différait entre les 2 chemins.

**Nouvelle formule unique** : `MeQ = nbRempli / nbObl` (comptage 1-champ = 1-unité), avec `nbObl` variable selon les règles conditionnelles. `PoidsMeQ` **n'est plus utilisé** pour le calcul (décision Q5 2026-07-07 : poids uniforme).

**Implémentation** :
- **`G_SIRUS_Regles.SIRUS_MeQLigne(ligneDict, sTable, configCol)`** : nouveau helper Public, source unique de vérité. Retourne `Dictionary { nbObl As Long, nbRempli As Long, taux As Double, manquants As Collection (de SIRUS_Champ_Config) }`. Skip Auto/Herite/placeholders `??_*`, applique Obligatoire=Oui direct, Conditionnel via `SIRUS_RegleObligatoire` (routing matrice Produit).
- **`G_SIRUS_PBS.CalculerMeQDTR_Optim`** : refonte, délègue à `SIRUS_MeQLigne`. Cumule `nbObl`/`nbRempli` sur les rows du DTR (moyenne pour tables multi-rows).
- **`G_SIRUS_Analyse.SIRUS_AnalyserFeuille`** : loop refactoré en 3 phases claires :
  - Phase 1 : `SIRUS_MeQLigne` (calcul MeQ + liste manquants)
  - Phase 2 : `AjouterAlerteManquant` (par def manquant)
  - Phase 3 : `AjouterAlerteMetier` (SIRUS_RegleAlerte par def)
  Suppression de `EvaluerChampLigne` (obsolète). Nouvelles subs Private `AjouterAlerteManquant` + `AjouterAlerteMetier`.

**Garantie** : MeQ ListView Generation = MeQ analyse SIRUS feuille pour un même DTR. Validé sur DTR0000120004 (LRUNR + Motif Preventif HARD) : 11/15 = 73,33 % dans les 2 chemins.

**Rappel** : l'analyse ne bloque jamais l'envoi ATMAP (décision v0.2 verrouillée). Le seuil MeQ historique 70 % est purement informatif.

### Bascule SnapShotPrix sur ADODB pur (BdD fermée)

**Historique itératif** :
- v1 : ADODB direct → NULL sur les DTR alphanumériques (bug ACE mixed-type).
- v2 : ADODB + `TypeGuessRows=0` dans les Extended Properties → inefficace chez Thomas (32 DTR numériques seulement).
- v3 : lecture directe de la feuille (bulk array) → fonctionne mais requiert l'ouverture de la BdD dans Excel (1 fois / session).
- **v4 (retenue)** : ADODB pur redevient viable après 2 modifs BdD côté Thomas 2026-07-08 :
  1. DTR alphanumériques déplacés en tête de `Produit_SnapShotPrix` (ACE type Text en scannant ses 8 premières lignes).
  2. Format cellule colonne DTR passe de "Général" à "Texte".
  Combiné à `IMEX=1` déjà présent dans la connection string, tous les DTR reviennent lisibles. Test `Test_SIRUS_TenterADODBPurPourDTR` valide : 29 748 DTR alphanum lus sur 29 788 lignes totales.

**Effets** :
- `G_SIRUS_PBS.ChargerSnapShotPrixDernierParDTR` : ADODB pur + cache session (`m_dictSnapShotPrixCache`). BdD ne s'ouvre **plus jamais** dans Excel.
- `DumpSnapShotBrutPourDTR` : idem, ADODB pur.
- Suppression des helpers direct-read (`OuvrirFeuilleSnapShotPrix`, `FermerBdDSiOuverte`, `ExtraireCheminBdDSirus`, `TrouverColParLibelleSirus`).
- Suppression de `Test_SIRUS_TenterADODBPurPourDTR` (mission accomplie).
- Cache session `m_dictSnapShotPrixCache` toujours en place : évite le re-scan des ~30 k lignes à chaque appel. Reset via `SIRUS_ResetCacheSnapShotPrix`.

### Diagnostic ajouté

- **`Test_SIRUS_DiagPreFillProduit(dtr)`** dans `G_SIRUS_PBS.bas` : dump complet pour un DTR — état du cache, config Produit (champs `Source='SnapShotPrix'`), contenu dict SnapShotPrix, cross-check config ↔ dict. Utile pour diagnostiquer des mismatches ColonneBdD ou un cache pollué.

### Nettoyage

- Suppression du `Attribute VB_Name = "..."` en tête des 41 fichiers `Modules/*.bas` (facilite le copier-coller manuel dans VBE — VBE gère le nom du module automatiquement).

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
