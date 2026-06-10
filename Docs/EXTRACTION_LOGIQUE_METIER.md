# Extraction ATM / EMMST — Logique métier

Document de référence des règles métier appliquées par les modules `E_Extraction_*` lors de l'extraction des fiches livrable VBN. Base pour la refonte tolérante (Phase 4 entrée A SIRUS).

> ⚠️ **Algorithme actuel = positions fixes.** La refonte vise à remplacer les coordonnées en dur (`D4`, `I6`, `B13`...) par une **détection par libellé** plus tolérante aux variations de template. Les règles métier ci-dessous **doivent être préservées** dans la refonte.

---

## 1. Détection des fiches DTR dans le fichier

Une feuille du classeur est une fiche DTR ATM/EMMST si son **nom contient "DTR"** (`ws.Name Like "*DTR*"`).

→ Comptage initial par `CompterFichesDTR(wb)`. Rejet si 0.

## 2. Validation du format

**Ancien (strict)** : la cellule `B2` de **chaque** fiche DTR contient exactement le titre attendu ("Analyse de Tâches de Maintenance" / "Maintenance Tasks Analysis" pour ATM).

**Nouveau (tolérant, depuis 2026-06-10)** : au moins une fiche DTR contient un mot-clé caractéristique dans la zone `A1:J10` (cf. `ValiderFormatLivrable` v2). Mots-clés :
- ATM FR : "Maintenance", "ATM", "Tâches", "Taches"
- ATM EN : "Maintenance", "MTA", "Tasks"
- EMMST FR : "Emballage", "EMMST", "Marquage", "Stockage"
- EMMST EN : "Packaging", "PHLST", "Labelling", "Storage"

---

## 3. Identifiant DTR (clé)

Depuis le **nom de feuille** :
```vba
dtr = Left(ws.Name, 13)
```
Les 13 premiers caractères du nom de feuille sont le code DTR (convention VBN).

## 4. Détection HARD / SOFT

Depuis le **nom de feuille** :
```vba
If Right(nomFeuille, 4) = "HARD" Then "HARD"
ElseIf Right(nomFeuille, 4) = "SOFT" Then "SOFT"
Else ""
```
Le suffixe du nom de feuille indique le type. Si ni l'un ni l'autre → vide (cas indéterminé).

---

## 5. Cartouche Info_ATM — positions fixes actuelles

Sur chaque fiche DTR ATM (FR) :

| Cellule | Champ | Champ BdD cible |
|---|---|---|
| `D4` | Désignation | `DesignationATM_INFOATM` |
| `D5` | Descriptif FR | `DescriptifATMFR_INFOATM` |
| `D5` (préfixé "REPLACE ") | Descriptif EN | `DescriptifATMEN_INFOATM` |
| `D10` | Commentaires FR | `CommentaireATM_INFOATM` |
| `D10` | Commentaires EN | `CommentairesATMEN_INFOATM` |
| `I6` | Poids | `Poids_INFOATM` |
| `I7` | Niveau maintenance | (utilisé pour Taches, pas Info) |
| `I8` | Plan | `PLANATM_INFOATM` |
| `I9` | Réparabilité | `Reparabilite_INFOATM` |
| (calculé) | DTR | `DTR_INFOATM` ← `Left(ws.Name, 13)` |
| (calculé) | Version | `Version_INFOATM` ← `"A1"` en dur |
| (recherche) | Temps total | `Temps générique de la tâche HARDWARE/SOFTWARE_INFOATM` |

→ Localisation FR/EN laissées **vides** par l'algo actuel.

## 6. Temps total — recherche par libellé

Déjà tolérant (utilise `Range.Find`) :

```vba
texteRecherche = IIf(langue = "FR", "Temps total :", "Total time :")
Set cellTemps = ws.Cells.Find(What:=texteRecherche, LookIn:=xlValues, lookat:=xlWhole)
RecupererTempsTotal = SafeNum(ws.Cells(cellTemps.Row, 10).Value, 0)
```

→ Trouve la ligne contenant "Temps total :", puis lit la valeur en **colonne J (10)** sur la même ligne.

Le temps est ensuite mappé selon HARD/SOFT :
- HARD → col N (Temps générique HARDWARE)
- SOFT → col O (Temps générique SOFTWARE)
- vide → col N par défaut

---

## 7. Détection feuille vide (pas de tâches)

```vba
EstFeuilleVide = (Cells(14,3).Value = "" AND _
                  Cells(17,3).Value = "" AND _
                  Cells(20,3).Value = "" AND _
                  Cells(24,3).Value = "")
```
4 lignes témoins (14, 17, 20, 24 col C) — si toutes vides, fiche considérée vide.

---

## 8. Extraction des tâches — algo

Plage scannée : ligne **13** à `Range("B10000").End(xlUp).Row` (dernière ligne remplie en col B).

Pour chaque ligne de cette plage :

| Contenu col B (B) | Sens | Action |
|---|---|---|
| Vide (`Len = 0`) | Ligne séparateur | Ignorer |
| 1 caractère (`Len = 1`) | Marqueur de chapitre (A/B/C/D) | Mémoriser la **catégorie** = `Cells(i, 3).Value` (col C) |
| Plus d'1 caractère | Tâche avec rang (ex. `2-1`, `3.2`...) | Extraire la tâche |

### Colonnes d'une ligne de tâche

| Col | Champ | Champ BdD cible |
|---|---|---|
| B (2) | Rang (préfixé `'` pour forcer texte) | `Rang_TACHEATM` |
| C (3) | Description | `Description_TACHEATM` |
| E (5) | Critères dimensionnels | `CritèresDimensionnels_TACHEATM` |
| F (6) | Outillage spécifique | `OutillageSpécifique_TACHEATM` |
| G (7) | Outillage standard | `OutillageStandard_TACHEATM` |
| H (8) | Pièces de rechange | `PiècedeRechangeConsommable_TACHEATM` |
| I (9) | Nb opérateurs | `NbOpérateurs_TACHEATM` |
| J (10) | Temps estimé | `TempsEstiméHeures_TACHEATM` |

→ Col D (4) n'est **pas** lue par l'algo actuel (peut contenir des sous-titres).

### Champs calculés / dérivés (par tâche)

| Champ BdD | Source |
|---|---|
| `DTR_TACHEATM` | `Left(ws.Name, 13)` |
| `Langue_TACHEATM` | param d'entrée (`FR` ou `EN`) |
| `HARD/SOFT` | `Right(ws.Name, 4)` |
| `Categorie_TACHEATM` | dernière catégorie mémorisée (chapitre courant) |
| `Niveau qualif_TACHEATM` | `ws.Range("I7").Value` (du cartouche) |
| `Temps analytique cumulé...HARD` | `tempsTotal` si HARD |
| `Temps analytique cumulé...SOFT` | `tempsTotal` si SOFT |
| `Pictogramme ATM2.0_TACHEATM` | vide |
| `Texte commentaire picto2_TACHEATM` | vide |

---

## 9. Vérification DTR en BdD

Pour chaque DTR extrait :
- Recordset BdD `Info_ATM$` : `RechercheValeurdansRS(REQSEL_INFO, dtr, "DTR_INFOATM")`
- Recordset BdD `Taches_ATM$` : `RechercheValeurdansRS(REQSEL_TACHES, dtr, "DTR_TACHEATM")`

Le résultat alimente la colonne `R` (`Connu en BDD ?`) avec `"Connu"` ou `"Inconnu"`.

---

## 10. Sorties — feuilles `Info_ATM` et `Taches_ATM`

Créées dans le **fichier livrable ouvert** (pas dans le projet), avec :
- Ligne 1 : en-têtes = vrais noms BdD (Champs_BDD.xlsx)
- Lignes 2+ : une ligne par DTR (Info_ATM) ou par tâche (Taches_ATM)
- Colonne R (18) : indicateur "Connu / Inconnu"

Format BdD prêt pour import → c'est ce que SIRUS consomme ensuite.

---

## 11. Cas particuliers / spécificités identifiées

- **Préfixe `REPLACE `** dans Descriptif EN : si Descriptif FR n'est pas vide, EN = "REPLACE " + FR. C'est une convention de marquage métier pour signaler une traduction à faire.
- **Préfixe `'`** sur le Rang : force Excel à traiter `2-1` comme texte et pas comme une formule de soustraction.
- **Localisation FR/EN** : laissées vides par l'algo actuel (cellules pas exploitées).
- **Wrap text désactivé** en fin d'extraction sur les 2 feuilles produites (mise en forme).

---

## 12. Refonte tolérante — règles à préserver impérativement

Lors de la refonte, **ces règles métier ne doivent pas être touchées** :

1. ✅ DTR = 13 premiers chars du nom de feuille
2. ✅ HARD/SOFT = suffixe du nom de feuille
3. ✅ Mapping tempsTotal → colonne N (HARD) / O (SOFT)
4. ✅ Préfixe `REPLACE ` pour Descriptif EN si FR non vide
5. ✅ Préfixe `'` sur le Rang
6. ✅ Catégorie de chapitre = col C de la ligne où col B fait 1 caractère
7. ✅ Marquage "Connu/Inconnu" en col R via recordset BdD
8. ✅ Version forcée à `"A1"`

Ce qui **peut/doit évoluer** :

| Élément actuel (rigide) | Évolution proposée |
|---|---|
| Cellules cartouche `D4`/`D5`/`D10`/`I6`/`I7`/`I8`/`I9` | Détection par libellé + lecture cellule adjacente droite/dessous |
| Ligne début tâches `= 13` | Détection du 1er chapitre (col B = "A" sur une ligne) |
| Colonnes tâches fixées B→J | Détection des en-têtes de colonne juste avant le marqueur "A" |
| Test feuille vide sur C14/C17/C20/C24 | Test si aucun marqueur de chapitre trouvé dans col B |
| Validation B2 = titre exact | Mots-clés dans A1:J10 (déjà fait) |
