# DATA_ASL — Spec du template figé (Sprint 1 bis)

Spec validée Thomas 2026-06-25 + livré 2026-06-26. Template figé : `Templates/ECOSLI_DATA_ASL.xlsx` (et script PowerShell de génération : `Templates/_build_template.ps1` + patch `Templates/_patch_template_v2.ps1`).

> **Statut** : ✅ Sprint 1 bis fonctionnellement complet. Refresh end-to-end validé avec DTRs numériques (8600xxx) ET alphanumériques (DTRxxxxxx, ATXyyyyy). Voir `CHANGELOG.md` entrée 2026-06-25 → 2026-06-26 pour le détail des modifs et du point dur **bypass ADO** (bug Microsoft ACE OLE DB sur .xlsb mixed-type).

> **Objectif global** : tuer toutes les formules à liens externes SharePoint dans DATA_ASL. Toutes les colonnes "issues BdD" sont remplies par le code VBA via ADODB (cliqué via bouton Rafraichir). DATA_ASL ne contient plus que :
> - Colonnes **saisie ASL** (vert) : remplies par l'ingénieur
> - Colonnes **calcul local** (bleu) : formules Excel internes sans lien externe
> - Colonnes **constantes** : valeurs en dur ("PH", "Y")
> - Colonnes **Refresh** (rouge) : remplies par VBA en valeurs pures via ADODB

---

## Structure du template — 41 colonnes (réparties en 5 blocs)

| # | Col | Libellé | Type | Source / Calcul |
|:-:|:-:|---|:-:|---|
| **Bloc 1 — Identification PBS (saisie ASL)** ||||||
| 1 | A | DTR Reference (issue de PBS) | 🟢 saisie | — |
| 2 | B | DESIGNATION (issue de PBS) | 🟢 saisie | — |
| 3 | C | SOUS SYSTÈME | 🟢 saisie | — |
| 4 | D | Qte | 🟢 saisie | — |
| **Bloc 2 — Groupe Générique** ||||||
| 5 | E | DESIGNATION (issue BdD) | 🔴 Refresh | `Produit$.[Désignation_PRODUIT]` WHERE `[DTR] = <DTR>` |
| 6 | F | DESIGNATION (Providor) | 🟦 calcul | `=IF([@[DESIGNATION (issue BdD)]]="";[@[DESIGNATION (issue de PBS)]];[@[DESIGNATION (issue BdD)]])` |
| 7 | G | Node typed (LRU/SRU…) | 🔴 Refresh | `Produit$.[Node type Produit (LRU/SRU…)_PRODUIT]` WHERE `[DTR] = <DTR>` |
| 8 | H | HARD/SOFT | 🔴 Refresh (calc SQL) | Agrégation `Taches_ATM$` (cf. SQL §A) → "BOTH" / "HARD" / "SOFT" / "" |
| 9 | I | NTI | 🟦 calcul | `=IF([@[Node typed (LRU/SRU…)]]="LRU";"3";IF([@[Node typed (LRU/SRU…)]]="LRUNR";"3";IF([@[Node typed (LRU/SRU…)]]="SRU";"3+";IF([@[Node typed (LRU/SRU…)]]="SRUNR";"3+";""))))` |
| **Bloc 3 — Groupe Données RAM** ||||||
| 10 | J | MTBF intrinsèque RAM Report spécifique | 🟢 saisie | — |
| 11 | K | MTBF fonctionnel RAM Report spécifique | 🟢 saisie | — |
| 12 | L | MTBF générique (offre+PGR) | 🔴 Refresh | `Produit$.[MTBF intrinsèque Produit (h) issu de la datasheet _PRODUIT]` WHERE `[DTR] = <DTR>` |
| 13 | M | Tx défaillance générique | 🟦 calcul | `=IFERROR(1/[@[MTBF générique (offre et PGR)]];"")` |
| 14 | N | Tx défaillance fonctionnel | 🟦 calcul | `=IFERROR(1/[@[MTBF fonctionnel RAM Report spécifique]];"")` |
| **Bloc 4 — Groupe Maintenance** ||||||
| 15 | O | Réf plan ensemble | 🔴 Refresh | `Produit$.[Drawing number]` WHERE `[DTR] = <DTR>` |
| 16 | P | Durée tâche correctif+préventif HARD générique | 🔴 Refresh | `Info_ATM$.[Temps générique de la tâche_INFOATM]` WHERE `[DTR_INFOATM] = <DTR>` AND `[HARD/SOFT] = 'HARD'` |
| 17 | Q | Durée tâche correctif+préventif HARD analytique | 🔴 Refresh (agrégat SQL) | SUM `Info_ATM$.[Temps analytique cumulé de la tache_INFOATM]` (cf. SQL §B HARD) |
| 18 | R | UNITÉ maintenance corrective | 🟦 constante | `"PH"` (en dur, écrit par Refresh) |
| 19 | S | DVL HARD générique | 🔴 Refresh | `Produit$.[Durée de vie générique Produit HARDWARE (année)_PRODUIT]` |
| 20 | T | DVL HARD spécifique | 🟢 saisie | — |
| 21 | U | DVL SOFT générique | 🔴 Refresh | `Produit$.[Durée de vie générique Produit (année) SOFTWARE _PRODUIT]` |
| 22 | V | DVL SOFT spécifique | 🟢 saisie | — |
| 23 | W | Durée tâche préventif SOFT générique | 🟦 calcul | `=IF([@[HARD/SOFT]]="BOTH";0,5;"")` |
| 24 | X | Durée tâche préventif SOFT analytique | 🔴 Refresh (agrégat SQL) | SUM `Info_ATM$.[Temps analytique cumulé de la tache_INFOATM]` (cf. SQL §B SOFT) |
| 25 | Y | UNITÉ maintenance préventive | 🟦 constante | `"Y"` (en dur, écrit par Refresh) |
| **Bloc 5 — Groupe Prix** ||||||
| 26 | Z | Prix FCS *(était `Cout total`)* | 🔴 Refresh | `SnapShotPrix$` Spare + FCS (cf. SQL §C) |
| 27 | AA | Prix spare catalogue *(était `ISP TIS FRANCE (FCS)`)* | 🔴 Refresh | `SnapShotPrix$` Spare + D&IS Catalogue (cf. SQL §C) |
| 28 | AB | Prix Spare calculé *(était `ISP TIS FRANCE (calculé)`)* | 🟦 calcul | `=IFERROR([@[Prix FCS]]*$AB$2;"")` *(coeff `1,0557` en ligne 2 col AB)* |
| 29 | AC | **Ref Source Spare** *(NOUVEAU)* | 🔴 Refresh | `SnapShotPrix$.[Ref Catalogue]` de la ligne Spare D&IS retenue (cf. SQL §C) |
| 30 | AD | Lead Time SPARE (jours calendaires) | 🔴 Refresh | `SnapShotPrix$.[LeadTime]` Spare D&IS (cf. SQL §C) |
| 31 | AE | Prix repair catalogue *(était `Prix repa (ISP Catalogue Répa)`)* | 🔴 Refresh | `SnapShotPrix$` Repair + D&IS Catalogue (cf. SQL §C) |
| 32 | AF | **Ref Source Repair** *(NOUVEAU)* | 🔴 Refresh | `SnapShotPrix$.[Ref Catalogue]` de la ligne Repair D&IS retenue (cf. SQL §C) |
| 33 | AG | Lead Time REPAIR (jours calendaires) | 🔴 Refresh | `SnapShotPrix$.[LeadTime]` Repair D&IS (cf. SQL §C) |
| 34 | AH | Cout spare FOURNISSEUR Not ISP | 🟢 saisie | — |
| 35 | AI | Lead Time SPARE FOURNISSEUR | 🟢 saisie | — |
| 36 | AJ | Cout Repair FOURNISSEUR Not ISP | 🟢 saisie | — |
| 37 | AK | Lead Time REPAIR FOURNISSEUR | 🟢 saisie | — |
| 38 | AL | Cout spare ALSTOM FRANCE FOURNISSEUR | 🟦 calcul | `=IF(ISBLANK([@[Cout spare FOURNISSEUR Not ISP]]);"";[@[Cout spare FOURNISSEUR Not ISP]]*$AL$2)` *(coeff en AL$2)* |
| 39 | AM | Prix repair ALSTOM FOURNISSEUR | 🟦 calcul | `=IF(ISBLANK([@[Cout Repair FOURNISSEUR Not ISP]]);"";[@[Cout Repair FOURNISSEUR Not ISP]]*$AM$2)` *(coeff en AM$2)* |
| 40 | AN | Prix Spare projet archive | 🔴 Refresh | `Projet_ASL$.[ISP TIS HORS FRANCE (archivé)]` AVG (cf. SQL §D Spare) |
| 41 | AO | Prix Repair projet archive | 🔴 Refresh | `Projet_ASL$.[Prix repair (DEVISE du projet (SLI))_PROJETASL]` AVG (cf. SQL §D Repair) |

### Suppressions par rapport à la version actuelle

- ❌ `Prix Manuel Spare` (était AL) — obsolète, plus de tracking manuel
- ❌ `Prix Manuel Repair` (était AM) — idem
- ❌ `Last Prix SP` (était AP) — remplacé par `Date Source` directe dans `SnapShotPrix`
- ❌ `Ref Source` (était AQ) — remplacée par `Ref Source Spare` (AC) + `Ref Source Repair` (AF)
- ❌ Colonne `AR` vide

### Ajouts

- ✅ `Ref Source Spare` (AC) — référence du catalogue ayant fourni le prix spare retenu
- ✅ `Ref Source Repair` (AF) — référence du catalogue ayant fourni le prix repair retenu

### Renommages

| Avant | Après |
|---|---|
| `Cout total` | **`Prix FCS`** |
| `ISP TIS FRANCE (FCS)` | **`Prix spare catalogue`** |
| `ISP TIS FRANCE (calculé)` | **`Prix Spare calculé`** |
| `Prix repa (ISP Catalogue Répa)` | **`Prix repair catalogue`** |

---

## Ligne 2 — Coefficients (à conserver au-dessus des headers)

| Cellule | Valeur | Utilisé par |
|---|:-:|---|
| `AB$2` | `1,0557` | Prix Spare calculé (col AB) |
| `AL$2` | `<coeff Alstom>` | Cout spare ALSTOM FRANCE FOURNISSEUR (col AL) |
| `AM$2` | `<coeff Alstom>` | Prix repair ALSTOM FOURNISSEUR (col AM) |

Les autres cellules de la ligne 2 sont libres.

**Ligne 3 = ligne d'entête** (auto-détectée par `MapColonnesDATA_ASL` via le libellé ancre `DTR Reference (issue de PBS)`).

---

## Code couleur (à appliquer sur les entêtes ligne 3)

| Couleur fond | RGB (Excel `Interior.Color`) | Sémantique | Cols concernées |
|---|---|---|---|
| 🟢 **Vert clair** | `5296274` (≈ `RGB(82, 220, 80)`) | Saisie ASL — l'utilisateur remplit | A, B, C, D, J, K, T, V, AH, AI, AJ, AK |
| 🔴 **Rouge** | `255` (≈ `RGB(255, 0, 0)`) | Refresh — rempli par le bouton via ADODB | E, G, H, L, O, P, Q, S, U, X, Z, AA, AC, AD, AE, AF, AG, AN, AO |
| 🟦 **Bleu clair** | `15773696` (≈ `RGB(0, 176, 240)`) | Calcul — recalculé en valeurs pures par `G_PBS_Refresh.RafraichirCalculs` (zéro formule dans le fichier — décision 2026-06-28) | F, I, M, N, W, AB, AL, AM |
| ⬜ **Gris clair** | (au choix Thomas) | Constante en dur écrite par Refresh | R ("PH"), Y ("Y") |

Ces couleurs sont **fonctionnelles** : elles permettent à l'ingénieur ASL de savoir d'un coup d'œil quelles cellules sont à sa charge (vert) vs auto-remplies (rouge / bleu / gris).

> **Évolution 2026-06-28** : les colonnes 🟦 (bleu) ne contiennent plus de formules Excel — `G_PBS_Refresh.RafraichirCalculs` les écrit en valeurs pures à chaque clic sur **Rafraichir Data PBS**. Le PowerShell `_patch_template_v2.ps1` continue de poser les formules à la génération du template, mais elles sont écrasées au 1er Refresh.

---

## Requêtes SQL ADODB — détail

Toutes les requêtes ciblent la BdD TRAVAIL via `CONNEXION_EXCEL_BDD_TRAVAIL`. Pré-agrégation côté SQL pour minimiser les allers-retours.

### §A — HARD/SOFT (col H)

Stratégie : 1 requête agrégée par DTR.

```sql
SELECT
    [DTR_TACHEATM],
    SUM(IIF([HARD/SOFT]='HARD', 1, 0)) AS nbHard,
    SUM(IIF([HARD/SOFT]='SOFT', 1, 0)) AS nbSoft
FROM [Taches_ATM$]
WHERE [DTR_TACHEATM] IN (<liste DTR DATA_ASL>)
GROUP BY [DTR_TACHEATM]
```

VBA post-traite : si nbHard>0 et nbSoft>0 → "BOTH" ; si nbHard>0 → "HARD" ; si nbSoft>0 → "SOFT" ; sinon "".

### §B — Durée tâche analytique HARD/SOFT (cols Q et X)

```sql
SELECT
    [DTR_INFOATM],
    [HARD/SOFT],
    SUM(VAL([Temps analytique cumulé de la tache_INFOATM])) AS tempsTotal
FROM [Info_ATM$]
WHERE [DTR_INFOATM] IN (<liste DTR>)
  AND ([HARD/SOFT] = 'HARD' OR [HARD/SOFT] = 'SOFT')
GROUP BY [DTR_INFOATM], [HARD/SOFT]
```

VBA : dict `{(DTR, HardSoft) → tempsTotal}` → écrit en col Q si HARD, X si SOFT.

### §C — Prix Catalogue (cols Z, AA, AC, AD, AE, AF, AG)

**Implémentation actuelle** (décision Thomas 2026-06-28) : le Refresh ne fait **aucune cascade sur les prix** — il dépose chaque ligne BdD dans sa colonne dédiée :

| Col | Contenu | Filtre BdD `SnapShotPrix$` |
|:-:|---|---|
| Z (`Prix FCS`) | MAX(`Date Source`) parmi les lignes FCS Spare | `Type=Spare` ET `Ref Catalogue LIKE '%FCS%'` |
| AA (`Prix spare catalogue`) | MAX(`Date Source`) parmi les lignes D&IS Spare | `Type=Spare` ET `Ref Catalogue LIKE '%D&IS Catalogue%'` |
| AC (`Ref Source Spare`) | Ref D&IS retenue **sinon** Ref FCS retenue (cascade D&IS → FCS pour le **label seulement**, pas le prix) | dépend de AA/Z |
| AD (`Lead Time SPARE`) | LeadTime de la ligne AA retenue | idem AA |
| AE (`Prix repair catalogue`) | MAX(`Date Source`) parmi les lignes D&IS Repair | `Type=Repair` ET `Ref Catalogue LIKE '%D&IS Catalogue%'` |
| AF (`Ref Source Repair`) | Ref D&IS Repair retenue (strict, pas de FCS Repair en BdD) | idem AE |
| AG (`Lead Time REPAIR`) | LeadTime de la ligne AE retenue | idem AE |

Aucune ligne avec ref autre que D&IS ou FCS n'est conservée. **Pas de sous-requête JOIN** : on lit toute la feuille `Produit_SnapShotPrix` en bulk array (bypass ADO pour contourner le bug ACE mixed-type sur la col DTR) et on filtre côté VBA via un Dictionary `{DTR|TYPE_CATEGORIE → MAX(Date)}`.

**La sous-priorité D&IS / FCS pour le prix se fait en aval** dans `G_SIRUS_PBS.CalculerPrixSpare` (ListView Generation) : si AA est vide mais AB l'est (= Prix FCS × coeff 1,0557, recalculé par `RafraichirCalculs`), on retient AB comme prix Spare avec source = `Ref Source Spare` (qui contient alors la ref FCS grâce au fallback de AC).

### §D — Prix archive multi-projets (cols AN, AO)

**Implémentation actuelle** (décision Thomas 2026-06-28) : pas de colonne Date dans `Projet_ASL$`, donc on retient le **projet alphabétiquement premier** (StrComp ascendant) parmi ceux ayant un prix > 0, **projet courant exclu**. Pas d'AVG.

```sql
SELECT [DTR_PROJETASL] & '' AS DTR_str,
       [PROJET_PROJETASL],
       [ISP TIS HORS FRANCE (archivé)],
       [Prix repair (DEVISE du projet (SLI))_PROJETASL]
FROM   [Projet_ASL$]
```

VBA post-traite : 2 dicts séparés (Spare / Repair) `{DTR_UC → Array(projetMin, prix)}`. Pour chaque ligne BdD :
- Si `prix > 0` et `projet <> projetCourant` :
  - Si projet plus petit que celui déjà retenu (StrComp asc) → on remplace
  - Sinon → on conserve

Le projet retenu est récupérable par la ListView Generation via `dictProjetASL_Rows` (cache de `G_SIRUS_PBS.ChargerRowsTableBdD_Optim`) pour reconstruire le label `"Projet : <nom>"`.

### §E — Colonnes calculées (cols I, N, O, AB, AL, AM)

**Décision Thomas 2026-06-28** : zéro formule Excel dans le fichier. `G_PBS_Refresh.RafraichirCalculs` écrit en valeurs pures (avec `Application.WorksheetFunction.Round` à 2 décimales pour les prix) :

| Col | Formule appliquée | Source | Cellule vide si |
|:-:|---|---|---|
| I (`NTI`) | `Select Case NodeType` → "3" / "3+" / "" | col G | NodeType ∉ {LRU, LRUNR, SRU, SRUNR} |
| N (`Tx déf fonct`) | `1 / MTBF fonctionnel` | col L | MTBF ≤ 0 ou non numérique |
| O (`Tx déf gen`) | `1 / MTBF générique` | col M | MTBF ≤ 0 ou non numérique |
| AB (`Prix Spare Calculé`) | `Prix FCS × cellule AB2` | col Z × 1,0557 | source ≤ 0 ou non numérique |
| AL (`Cout Spare Alstom`) | `Cout Spare FRS × cellule AL2` | col AH × 1,17 | source ≤ 0 ou non numérique |
| AM (`Prix Repair Alstom`) | `Cout Repair FRS × cellule AM2` | col AJ × 1,75 | source ≤ 0 ou non numérique |

Helper `EstNumeriqueNonNul(v)` filtre les zéros pour éviter d'écrire `0,00 €` qui serait réinterprété par la ListView Generation comme un prix Fournisseur connu (cascade Frs → D&IS → ...).

Format **Comptabilité €** (`_-* #,##0.00\ "€"_-;-* #,##0.00\ "€"_-;_-* "-"??\ "€"_-;_-@_-`) appliqué via `AppliquerFormatPrix(lo)` aux 10 colonnes prix (Z/AA/AB/AE/AH/AJ/AL/AM/AN/AO) à chaque Refresh.

---

## Migration des fichiers projet existants (legacy)

Les classeurs projets déjà déployés chez les ingés ASL contiennent encore les anciennes formules à liens externes. Routine de migration one-shot à intégrer dans `A_2_ModUpdate` au démarrage si la version du template DATA_ASL détectée est < seuil :

1. Détection : si la cellule `A1` ne porte pas un marqueur `"DATA_ASL v2"` (à définir), considérer comme legacy.
2. Action :
   - Pour chaque colonne de DATA_ASL : si la cellule de ligne 4 contient une formule avec `"https://"`, convertir toute la colonne en valeurs pures (`Range.Value = Range.Value`).
   - Renommer les colonnes (Cout total → Prix FCS, etc.) si présentes avec l'ancien libellé.
   - Supprimer les colonnes obsolètes (Prix Manuel Spare/Repair, Last Prix SP, Ref Source, AR).
   - Ajouter les nouvelles colonnes (Ref Source Spare, Ref Source Repair) si absentes.
   - Marqueur `A1 = "DATA_ASL v2"` posé.
   - Trace dans une feuille `_PurgeLog_DATA_ASL` (audit).

À implémenter après le code Refresh, pour éviter de stresser sur la migration en même temps que le code initial.

---

## Plan d'action

### Étape 1 — Thomas applique la spec dans le template
1. Ouvrir `Templates/ECOSLI.xlsm` → feuille `DATA_ASL`.
2. Réordonner les colonnes selon le tableau ci-dessus (41 colonnes finales).
3. Renommer les 4 colonnes (Prix FCS, Prix spare catalogue, Prix Spare calculé, Prix repair catalogue).
4. Insérer les 2 nouvelles colonnes (`Ref Source Spare` en AC, `Ref Source Repair` en AF).
5. Supprimer les 5 colonnes obsolètes (Prix Manuel Spare, Prix Manuel Repair, Last Prix SP, Ref Source, AR).
6. Vider toutes les cellules rouges (Refresh) — supprimer les formules `=XLOOKUP(https://...)`.
7. Vérifier les calculs locaux des cellules bleues (F, I, M, N, W, AB, AL, AM) — formules à conserver sans lien externe.
8. Vérifier les coefficients en ligne 2 (AB$2, AL$2, AM$2).
9. Appliquer les couleurs de fond selon le code couleur §"Code couleur".
10. Commit du nouveau template.

### Étape 2 — Claude code le bouton Refresh
1. Mettre à jour `Type DATA_ASL_Col` dans `G_SIRUS_PBS.bas` (suppr/ajout/rename champs).
2. Mettre à jour `MapColonnesDATA_ASL` avec les nouveaux libellés exacts.
3. Créer le module `G_PBS_Refresh.bas` avec `SIRUS_RafraichirDATA_ASL(wb As Workbook)` qui :
   - Liste les DTRs de DATA_ASL (col A).
   - Exécute les 4 requêtes SQL §A à §D (avec cache).
   - Écrit en valeurs pures dans les colonnes 🔴 Refresh.
   - Écrit les constantes "PH" (R) et "Y" (Y).
   - Trace `_RefreshLog`.
4. Ajouter `Btn_Rafraichir_DATA_ASL` dans le UserForm (page Generation).
5. Câbler le handler `Btn_Rafraichir_DATA_ASL_Click()` → `G_PBS_Refresh.SIRUS_RafraichirDATA_ASL(ActiveWorkbook)`.

### Étape 3 — Recette
- Test sur 1 classeur projet réel avec quelques DTR connus en BdD.
- Vérifier que toutes les cellules rouges sont remplies correctement.
- Vérifier les cas multi-projets pour Prix archive (AVG).
- Test downstream : générer un livrable Providor/ATM/EMMST → doit fonctionner identiquement (DATA_ASL toujours lisible).

### Étape 4 — Migration legacy (plus tard)
Implémentation de la routine `PurgerEtMigrerDATA_ASL` dans `A_2_ModUpdate`.
