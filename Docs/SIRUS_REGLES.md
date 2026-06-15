# SIRUS — Règles métier de validation

Référence des règles appliquées par le moteur d'analyse `G_SIRUS_Analyse.SIRUS_AnalyserFeuille` via `G_SIRUS_Regles` + `G_SIRUS_Config`.

> ⚠️ **Source unique de vérité** : ce document décrit ce qui est implémenté dans `G_SIRUS_Config.bas` (`Peupler*`) + `G_SIRUS_Regles.bas` (`SIRUS_RegleObligatoire`, `SIRUS_RegleAlerte`). Toute évolution doit être appliquée dans le code ET ici.

## Logique générale

Pour chaque ligne de la feuille d'analyse, et pour chaque champ du config :

1. **Obligatoire = "Oui"** dans le config → champ doit être rempli, sinon alerte "Obligatoire manquant".
2. **Obligatoire = "Conditionnel"** → on appelle `SIRUS_RegleObligatoire(table, colonneBdD, ligne)` qui renvoie `True/False` selon le contexte de la ligne (ex. NodeType, MotifNA…). Si `True` et le champ est vide, alerte.
3. **`SIRUS_RegleAlerte(table, colonneBdD, ligne)`** est appelée systématiquement pour tous les champs. Renvoie `""` (rien) ou un libellé d'alerte métier (cohérence transverse : MTBF fonctionnel < intrinsèque, projet inconnu…).

Le **taux MeQ** = somme(`PoidsMeQ` des obligatoires remplis) / somme(`PoidsMeQ` des obligatoires). Seuil d'autorisation envoi = **70%**.

L'analyse **ne bloque jamais** — décision verrouillée roadmap. Sauf cas explicite (cf. ProjetASL projet inconnu ci-dessous, blocage à l'envoi).

---

## Table ProjetASL

Couvre les opérations Création et Modification d'une entrée de la table BdD `Projet_ASL$`.

### Ordre des colonnes de la feuille (verrouillé)

| Col | Libellé humain (Champ) | ColonneBdD | Ordre config |
|:---:|---|---|:---:|
| 1 | Statut | `_statut` | — |
| 2 | Nom projet | `PROJET_PROJETASL` | 10 |
| 3 | Sous Système | `Sous Système_PROJETASL` | 20 |
| 4 | **DTR (clé)** | `DTR_PROJETASL` | 30 |
| 5 | Designation PBS | `Désignation_PROJETASL` | 40 |
| 6 | Qte | `Q PBS_PROJETASL` | 50 |
| 7 | Node Type | `Node typed PROJET_PROJETASL` | 60 |
| 8 | MTBF intrinseque Projet | `MTBF intrinsèque Projet (h) ..._PROJETASL` | 70 |
| 9 | MTBF fonctionnel | `MTBF fonctionnel (h) ..._PROJETASL` | 80 |
| 10 | DVL fonctionnel HARD | `DUREE DE VIE LIMITEE ... HARDWARE_PROJET ASL` | 90 |
| 11 | DVL fonctionnel SOFT | `DUREE DE VIE LIMITEE ... SOFTWARE_PROJET ASL` | 100 |
| 12 | Prix Spare archive | `Prix spare (DEVISE du projet (SLI))_PROJETASL` | 110 |
| 13 | Lead Time Spare archive | `Lead time SPARE (jours)_PROJETASL` | 120 |
| 14 | Prix Repair archive | `Prix repair (DEVISE du projet (SLI))_PROJETASL` | 130 |
| 15 | Lead Time Repair archive | `Lead time REPAIR (jours)_PROJETASL` | 140 |
| 16 | Testabilite | `Testabilité ..._PROJETASL` | 150 |
| 17 | Prix Preventif | `Prix préventif (DEVISE du projet (SLI))_PROJETASL` | 160 |
| 18 | Prix After Market Spare | `Prix after market Spare_PROJETASL` | 170 |
| 19 | Prix After Market Repair | `Prix after market Repair_PROJETASL` | 180 |
| 20 | STBI/DAT | `STBI/DAT_PROJETASL` | 190 |
| 21 | LCS | `LCS` | 200 |
| 22 | Commentaires (auto-rempli analyse) | `_commentaires` | — |

### Champs obligatoires (`Obligatoire = "Oui"`)

| Champ | PoidsMeQ |
|---|:---:|
| Nom projet | 3 |
| Sous Système | 2 |
| DTR | 3 |
| Designation PBS | 2 |
| Qte | 1 |
| MTBF fonctionnel | 2 |
| Prix Spare archive | 2 |
| Lead Time Spare archive | 1 |

Testabilité (liste Oui/Non) reste en saisie possible mais **non obligatoire** (décision 2026-06-15) — n'entre pas dans le calcul MeQ Generation ni dans le MeQ SIRUS d'envoi ATMAP.

### Champs conditionnels (`Obligatoire = "Conditionnel"`)

Évalués via `SIRUS_RegleObligatoire` :

| Champ | Règle |
|---|---|
| **Prix Repair archive** `Prix repair (DEVISE du projet (SLI))_PROJETASL` | Obligatoire si `Node typed PROJET_PROJETASL` ∈ (`LRU`, `SRU`). Sinon, saisie libre. |
| **Lead Time Repair archive** `Lead time REPAIR (jours)_PROJETASL` | Idem. |

**Justification métier** : un KIT / VISSERIE / ASSEMBLY n'a pas de cycle de réparation, donc pas de prix ni de lead time Repair attendus.

### Alertes de cohérence (`SIRUS_RegleAlerte`)

| Champ déclencheur | Condition | Message |
|---|---|---|
| `MTBF fonctionnel (h) issu du RAM report projet_PROJETASL` | MTBF fonctionnel < MTBF intrinsèque Produit | `MTBF fonctionnel < MTBF intrinseque : verifier RAM report` |
| `PROJET_PROJETASL` | Le nom de projet n'existe pas dans la table BdD `Projet_SLI$` | `Projet inconnu dans ProjetSLI : import impossible` |

### Enrichissement depuis DATA_ASL (Modifier)

`SIRUS_EnrichirProjetASL_DepuisDATA_ASL` (G_SIRUS_Generation) est appelée après le pré-remplissage BdD lors d'un ProjetASL Modifier. Croise la PBS du classeur projet (DATA_ASL) avec la feuille SIRUS pour overwrite les champs reflétant la PBS courante.

| Champ ProjetASL (feuille SIRUS) | Colonne DATA_ASL (mapping `DATA_ASL_type`) | Source |
|---|---|---|
| Nom Projet | — | `Project_List.Value` (combo principale UserForm) |
| Sous Système | `SousSysteme` | DATA_ASL |
| DTR | `DTR_PBS` | DATA_ASL |
| Designation PBS | `Designation_PBS` | DATA_ASL |
| Qte | `Qte` | DATA_ASL |
| Node Type | `NodeType` | DATA_ASL |
| MTBF intrinsèque Projet | `MTBF_Intr` | DATA_ASL (RAM report spécifique projet CGR) |
| MTBF fonctionnel | `MTBF_Fonct` | DATA_ASL (RAM report spécifique projet CGR) |
| DVL fonctionnel HARD | `DureeVie_Ana_HARD` | DATA_ASL (analytique = spécifique projet) |
| DVL fonctionnel SOFT | `DureeVie_Ana_SOFT` | DATA_ASL |
| Lead Time Spare archive | `LeadTimeSpareCalend` | DATA_ASL (jours calendaires) |
| Lead Time Repair archive | `LeadTimeRepairCalend` | DATA_ASL (jours calendaires) |
| **Prix Spare archive** | cascade ↓ | (cf. ci-dessous) |
| **Prix Repair archive** | cascade ↓ | (cf. ci-dessous) |
| Testabilité, Prix Préventif, Prix After Market Spare/Repair, STBI/DAT, LCS | — | Saisie manuelle (pas de source DATA_ASL) |

**Cascade prix Spare/Repair** (`CascadePrixDATA_ASL`, reprend la logique de `D_Providor_BreakDownStructure.RemplirPrixSpare/Repair`) :

1. **Fournisseur** (`CoutSpare_AlstomFR_FRS` / `PrixRepair_Alstom_FRS`) — *Cout spare ALSTOM FRANCE FOURNISSEUR* / *Prix repair ALSTOM FOURNISSEUR*
2. **Catalogue** (`ISP_TIS_FCS` / `PrixRepa_ISP`) — *ISP TIS FRANCE (FCS)* / *Prix repa (ISP Catalogue Répa)*
3. **Manuel** (`PrixManuelSpare` / `PrixManuelRepair`) — *Prix Manuel Spare* / *Prix Manuel Repair*
4. **Archive** (`PrixSpare_ProjetArchive` / `PrixRepair_ProjetArchive`) — *Prix Spare sur un projet archive* / *Prix Repair sur un projet archive*

Une valeur est considérée "connue" si `IsNumeric` ET `<> 0`. La 1re source connue gagne. Si aucune n'est trouvée, la valeur déjà posée par le pré-remplissage BdD est conservée (pas d'écrasement).

### Garde-fou à l'envoi ATMAP

⚠️ **Bloquant** : `SIRUS_UI_BtnEnvoyerClick` (G_SIRUS_UserForm) refuse l'envoi du ticket si **au moins une ligne** a un projet absent de `Projet_SLI$`. MsgBox vbCritical avec la liste des projets fautifs, le ticket n'est pas créé.

Justification : la cohérence référentielle SLI → ASL est imposée par la gouvernance Data Owner. Pas de ProjetASL sans ProjetSLI préalable.

Implémentation : `SIRUS_ProjetsASL_Inconnus(feuilles())` scanne la col 2 (Nom projet) de chaque feuille `SIRUS_ProjetASL_*` à envoyer, croise avec le cache `G_SIRUS_Regles.SIRUS_ProjetsSLI()` (chargé une fois depuis la BdD TRAVAIL au 1er appel de la session). Reset du cache via `SIRUS_ResetCacheProjetsSLI`.

### Layout / mise en forme

- Freeze panes : colonnes 1-4 (Statut + Nom Projet + Sous Système + DTR).
- Alignement vertical : centré sur toute la zone de données.
- Suffixe entête : ` *` rouge pour `Obligatoire = Oui`, ` (*)` rouge pour `Conditionnel`.
- Cellules `Source = Auto` / `Herite` (Node Type, MTBF intrinseque Projet, DVL fonctionnel HARD/SOFT) : fond gris + verrouillées par protection feuille (mot de passe `SIRUS`).
- Colonne Commentaires (dernière) : auto-remplie par l'analyse — 1 alerte = texte simple, N alertes = liste à puces (`- alerte\n- alerte\n...` avec WrapText).

---

## MeQ Generation — ListView_Generation du UserForm

La `ListView_Generation` (onglet Generation du UserForm) affiche pour chaque DTR de DATA_ASL :

| Colonne | Contenu | Couleur |
|---|---|---|
| DTR | Identifiant PBS | — |
| Nb projets | Nombre de projets distincts où ce DTR existe en `Projet_ASL$` | Vert si ≥ 1, rouge gras si 0 |
| ATM - Info | % MeQ par DTR | Vert si 100%, rouge gras sinon, "Absent" si DTR absent |
| ATM - Taches | idem (moyenne sur toutes les tâches du DTR) | idem |
| EMMST | idem | idem |
| Produit | idem | idem |

Le tooltip de ligne (`ListItem.ToolTipText`, hover sur la cellule DTR) affiche la liste des projets connus pour ce DTR : `Projets connus (N) : ProjA, ProjB, ProjC`. MSComctlLib v6 ne supporte pas le tooltip par sub-item, donc c'est niveau ligne.

### Calcul du MeQ (cols ATM-Info / ATM-Taches / EMMST / Produit)

- **Source des champs** : `SIRUS_LireConfig(<table>)` (chargée 1× par table en début de remplissage, plus N× — optim 2026-06-15).
- **Numérateur** : champs `Obligatoire = "Oui"` **remplis** + champs `Obligatoire = "Conditionnel"` dont `SIRUS_RegleObligatoire` renvoie `True` et qui sont remplis.
- **Dénominateur** : tous les champs `Oui` + Conditionnels actifs (selon la règle).
- **Multi-rows par DTR** (TachesATM) : moyenne pondérée (sum num / sum den) sur toutes les rows.
- **Tables couvertes** : Produit, InfoATM, TachesATM, EMMST (ProjetSLI hors scope).

### Optimisations de performance (2026-06-15)

- `SIRUS_LireConfig` chargée **5 fois au total** (1× par table) au lieu de 5×N (N = nombre de DTR) — économie majeure pour N>50.
- `SELECT` ciblé : ne charge que les colonnes du config `Source = "Saisie"` (en excluant les placeholders `??_*_a_definir`). EMMST passe de ~50 colonnes à ~15. Fallback automatique `SELECT *` si une colonne du config n'est pas alignée avec la BdD (trace `Debug.Print`).
- Helper `NzSafeStr` pour gérer les `Null` ADO sans crash `CStr(Null)`.

### Refresh

Bouton **Rafraichir** (`SIRUS_RafraichirListViewGeneration`) à actionner après changement de PBS ou modification BdD. Pas de refresh automatique sur `Project_List_Change` (le compte de projets ne dépend pas du projet sélectionné depuis qu'on a viré le filtre).

Implémentation : `G_SIRUS_PBS.ChargerRowsTableBdD_Optim` + `CalculerMeQDTR_Optim` + `ProjetsConnusPourDTR` + `AjouterSubItemMeQ` / `AjouterSubItemNbProjets`.

## Tables non encore documentées

Les règles des tables Produit, InfoATM, TachesATM, EMMST, ProjetSLI sont dans `G_SIRUS_Regles.bas` (Phase 2 MVP — 10 règles) et `G_SIRUS_Config.bas` (`Peupler*`). À documenter ici au fur et à mesure de l'enrichissement (cf. roadmap Phase 2.1).

---

## Comment ajouter une nouvelle règle

1. **Identifier** la table et le `ColonneBdD` exact (cf. `SIRUS_Config` → onglet `tblSIRUSConfig` via `Test_DiagnostiquerEcartsConfig_Toutes`).
2. **Choisir** : règle d'obligation (champ vide = problème) ou règle d'alerte (incohérence métier).
3. **Coder** dans `G_SIRUS_Regles.bas` :
   - Ajouter un `Case "<Table>|<ColonneBdD>"` dans `SIRUS_RegleObligatoire` ou `SIRUS_RegleAlerte`.
   - Utiliser `GetVal(ligne, "<autreColonneBdD>")` pour inspecter les autres champs de la ligne.
4. **Tester** unitairement via une `Public Sub Test_Regle_<...>()` (cf. exemples en fin de `G_SIRUS_Regles.bas`).
5. **Documenter ici**.

Pattern verrouillé (validation Thomas) : **`Select Case (Table | ColonneBdD)`** — pas de mini-langage évaluable. Les conditions descriptives du config (`ConditionDescription`) sont purement humaines.
