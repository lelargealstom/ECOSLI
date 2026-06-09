# ATM V2 — Pistes d'optimisation restantes

État au **2026-06-09**, après les commits :
- `e76ef02` : cache positions chapitres + `End(xlUp)` + `Calculation=manual`
- `8fe9b9f` : ATM V1 BOTH remplit aussi feuille SOFT
- `3a7665b` : Insert + CopyOrigin (V1)
- `d109c47` : SQL direct par DTR (V1)
- `7ec2561` : Index Dictionary après les Cadrage_*
- `ad8996b` : Option C — index Dictionary + fallback linéaire

**Contraintes verrouillées** (validation Thomas) :
- ✅ Garder **1 fichier Excel par DTR** (pas de classeur unifié)
- ✅ Garder la **comparaison initiale avec la bibliothèque illustrée** (`BIBLIO_PATH`) — si fichier illustré présent, copie directe sans génération

---

## Pistes triées par impact

### 🔥 Gros gain — Piste 2 : Écriture en bulk au lieu de N `EntireRow.Insert`

**Problème** : pour chaque tâche, le code actuel fait :
1. `EntireRow.Insert` (1 ligne) ⇒ Excel décale tout le contenu en dessous
2. Met à jour les compteurs `LigneProchainChap` / `ligneDebut` de tous les chapitres suivants (boucle `For k = numChapitre To UBound(chapitres)`)
3. Écrit 4 cellules (B, C, D, L) en 4 appels COM séparés
4. Pour Chapitre C avec sous-chapitre : 2 `Insert` au lieu d'1

Avec N tâches dans un DTR : **N insertions + N×4 mises à jour de compteurs + 4×N appels COM d'écriture**.

**Refactor proposé** :
- **1re passe** sur `Taches_tab[PMB..PMH]` : pour chaque chapitre, compter les tâches et préparer un tableau 2D `arr(1..n, 1..11)` avec les valeurs à écrire.
- **2e passe** : 1 seul `Range(...).EntireRow.Insert` de N lignes par chapitre, puis 1 seul `Range(...).Resize(n, 11).Value = arr`.

**Gain estimé** : 3-5× sur les fiches chargées (rang 3/4 nombreux).
**Risque** : refactor des fonctions `InsererTacheChapitre` + `InsererTacheChapitreC`. Bien tester le Chapitre C avec sous-chapitres (B1, B2…).

**Fichiers concernés** : `B_ATM_V2.bas` (lignes 350-540 environ).

---

### 🟡 Gain moyen — Piste 4 : `Workbooks.Open` invisible dans une 2ᵉ instance Excel

**Problème** : chaque DTR ouvre un classeur dans l'instance principale Excel, ce qui force un rendu UI même avec `ScreenUpdating = False` (notification taskbar, ribbon, etc.), et empoisonne l'environnement de l'utilisateur (qui voit clignoter les fichiers).

**Refactor proposé** :
```vba
Dim xlBg As Object
Set xlBg = CreateObject("Excel.Application")
xlBg.Visible = False
xlBg.DisplayAlerts = False
xlBg.Calculation = xlCalculationManual
' ... boucle DTR : xlBg.Workbooks.Open(...) ...
xlBg.Quit
Set xlBg = Nothing
```

**Risque** : gérer proprement `xlBg.Quit` (sinon Excel.exe fantôme dans le Task Manager). Penser à un `On Error` qui ferme l'instance en cas de plantage.

**Gain estimé** : modéré, surtout ressenti côté UX (pas de clignotement).

---

### 🟡 Gain moyen — Piste 5 : Hardcoder les colonnes BDD via constantes (style V1)

**Problème** : V2 dépend toujours de `Cadrage_Taches_ATM` + `Taches_type.dtr` (Find sur les en-têtes du RecordSet) alors que l'ordre du SELECT dans `B_ATM_Commun.RecupererDonneesBDD_ATM` est figé.

**Refactor proposé** : copier le pattern V1 (`B_ATM_V1.bas` l. 7-36) avec des `Private Const RC_TACHES_*` et supprimer la phase Cadrage.

**Gain** : petit côté perf (1 fonction de moins par run), gros côté lisibilité + cohérence avec V1.

---

### 🟡 Gain moyen — Piste 6 : Filtrer par langue côté SQL

**Problème** : `RecupererDonneesBDD_ATM(2)` filtre `WHERE [HARD/SOFT]='HARD'` mais récupère **toutes les langues**. Le filtre par langue est fait dans la boucle VBA (`B_ATM_V2.bas` l. 285 : `If Taches_tab(j, langue) <> langue Then GoTo NextTache`).

**Refactor proposé** : ajouter `AND [Langue_TACHEATM] = '<langue>'` dans le SELECT V2.

**Gain** : divise `Taches_tab` par 2 + supprime le filtre VBA.

**Attention** : ne pas impacter V1 (qui a besoin des 2 langues pour `HARD_SOFT()` ou autre). Faire ça uniquement dans la branche `versionATM = 2` de `B_ATM_Commun`.

---

### 🟢 Petits gains / propreté

**Piste 8 — `ligvide As Integer` → `As Long`**
Dans `InsererTacheChapitreC` ligne 426 : `Dim ligvide As Integer`. V1 a déjà été corrigé. Risque latent d'overflow si fiche dépasse 32767 lignes (peu probable, mais 0 coût à figer).

**Piste 9 — Écriture d'1 tâche en 1 appel COM**
Remplacer :
```vba
.Range("B" & ligvide) = ...
.Range("C" & ligvide) = ...
.Range("D" & ligvide) = ...
.Range("L" & ligvide) = ...
```
par :
```vba
.Range("B" & ligvide & ":L" & ligvide).Value = Array(rang, cat, desc, "", "", "", "", "", "", "", temps)
```
Réduit le nombre d'appels COM par tâche (×4 → ×1).

---

### 🚀 Refactor plus profond — Piste 10 : Génération en mémoire + duplication de feuille

**Idée** : ouvrir UN seul workbook "modèle" en background (instance Excel séparée), dupliquer la feuille `DESCRIPTION_TACHE` par DTR (`Sheets("...").Copy`), remplir, exporter chaque feuille en `SaveAs` indépendant. Évite N `Workbooks.Open`/`Save`/`Close` complets.

**Gain potentiel** : massif sur projets 100+ DTR.
**Coût** : gros refactor, beaucoup de logique à revoir (couverture, mise en page, formats).
**Contraintes à valider** : compatibilité avec la sortie 1 fichier/DTR (oui, via `Sheets.Copy ToNewBook`).

À envisager seulement si les Pistes 2+4+5+6 ne suffisent pas.

---

## Ordre recommandé pour la prochaine session ATM V2

1. **Piste 2** (écriture en bulk) — le plus gros gain restant, contenu, validable rapidement.
2. **Piste 5** (constantes RC_*) — clarté + cohérence V1, petit refactor.
3. **Piste 6** (filtre SQL par langue) — 1 ligne SQL, 1 ligne VBA en moins.
4. **Piste 9** (écriture COM unique) — quick win après le refactor Piste 2.
5. **Piste 8** (Long) — 1 ligne, à faire en passant.
6. **Piste 4** (instance Excel séparée) — quand on aura besoin du gain UX.
7. **Piste 10** — refactor profond, en dernier recours.
