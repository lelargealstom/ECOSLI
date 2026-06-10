# Change Log — ECOSLI.xlam

Suivi des modifications apportées au complément Excel `ECOSLI.xlam` au fil des sessions de développement.

Format : `[Date] — Auteur — Périmètre`, puis liste des changements par module.

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
