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
| Testabilite (liste Oui/Non) | 1 |

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
