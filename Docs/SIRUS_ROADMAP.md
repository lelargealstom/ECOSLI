# SIRUS — Roadmap de développement

> Plan d'avancement du module **SIRUS** dans `ECOSLI.xlam`.
> Document vivant : à jour des décisions prises en sessions de dev jusqu'au **2026-06-09**.
> Référence : brief v0.2 (voir aussi `Docs/CHANGELOG.md` pour l'historique détaillé).

---

## État global

| Phase | Périmètre | État | Module(s) |
|---|---|:---:|---|
| 0 — Socle config | `SIRUS_Config` + `SIRUS_Champ_Config` (~158 champs sur 6 tables) | ✅ | `G_SIRUS_Config.bas`, `SIRUS_Champ_Config.cls` |
| 0 bis — Listes en code | `SIRUS_GetListe(cle)` + 7 listes | ✅ | `G_SIRUS_Listes.bas` |
| 1 — Générateur de feuille | `SIRUS_GenererFeuilleAnalyse(table, op, [cles])` + diag BdD | ✅ | `G_SIRUS_Generation.bas` |
| 2 — Moteur d'analyse | `G_SIRUS_Regles` + généralisation `SIRUS_AnalyserFeuille` | 🔴 | à créer |
| 3 — Surfaces UI | radar UserForm + verrouillage cellules + diff Modifier | 🟡 | partiel |
| 4 — Entrées A et B | consommation ATM fournisseur + détection écarts PBS | 🔴 | à créer |
| 5 — Envoi ATMAP | PJ normalisée + branchement ticket | 🟡 | ATMAP déjà fonctionnel |
| 6 — Industrialisation | logs, paramètres, perfs, recette | 🔴 | à créer |

Légende : ✅ fait — 🟡 partiel — 🔴 à faire.

---

## Phase 2 — Moteur d'analyse `G_SIRUS_Regles`

**Objectif** : remplacer la pseudo-syntaxe `ConditionDescription` (purement descriptive aujourd'hui) par des règles VBA en dur, façon « D1 » validée avec Thomas.

### Module à créer : `G_SIRUS_Regles.bas`

```vba
' API publique
Public Function SIRUS_RegleObligatoire(ByVal sTable As String, _
                                       ByVal sColonneBdD As String, _
                                       ByRef ligne As Object) As Boolean
    ' ligne = Dictionary {ColonneBdD -> valeur} de la ligne en cours
    ' Renvoie True si le champ est obligatoire pour cette ligne
End Function

Public Function SIRUS_RegleAlerte(ByVal sTable As String, _
                                  ByVal sColonneBdD As String, _
                                  ByRef ligne As Object) As String
    ' Renvoie "" si OK, sinon le libellé de l'alerte
End Function
```

### Pattern d'implémentation

`Select Case (Table, ColonneBdD)` pour chaque règle conditionnelle :

```vba
Select Case sTable & "|" & sColonneBdD
    Case "Produit|Prix SPARE manuel_PRODUIT"
        ' Obligatoire si NodeType LRU/SRU
        SIRUS_RegleObligatoire = (ligne("Node type Produit (LRU/SRU…)_PRODUIT") = "LRU" _
                              Or ligne("Node type Produit (LRU/SRU…)_PRODUIT") = "SRU")
    Case "TachesATM|Description_TACHEATM"
        ' Obligatoire si Rang commence par 3 ou 4
        Dim r As String : r = CStr(ligne("Rang_TACHEATM"))
        SIRUS_RegleObligatoire = (Left$(r, 2) = "3." Or Left$(r, 2) = "4.")
    ' …
End Select
```

### `SIRUS_AnalyserFeuille(table)`

Généralisation de `G_SIRUS_Analyse.SIRUS_AnalyserProduit`. À écrire dans `G_SIRUS_Analyse.bas` :

```vba
Public Function SIRUS_AnalyserFeuille(ByVal sTable As String) As Double
    ' Renvoie le taux de complétude (0..1) de la feuille SIRUS_<table>_<op>
    ' active dans ActiveWorkbook.
    ' Pour chaque ligne, pour chaque champ du config :
    '   1. Si Source = "Saisie" et SIRUS_RegleObligatoire(...) = True et cellule vide
    '      → ajouter Alerte "Champ obligatoire manquant"
    '   2. Si SIRUS_RegleAlerte(...) <> "" → ajouter Alerte avec libellé
    '   3. Calcul taux MeQ = somme(PoidsMeQ remplis) / somme(PoidsMeQ obligatoires)
    ' Écrit le statut inline (colonne Statut) en couleur + remplit la ListView Erreurs.
End Function
```

### Critères de réussite Phase 2

- [ ] Sur une feuille `SIRUS_Produit_Creer` avec 2 DTR (un LRU et un SRUNR), `SIRUS_AnalyserFeuille` détecte correctement les obligations conditionnelles (Prix Spare manuel obligatoire pour LRU, 1 prix Spare sur 3 pour SRUNR).
- [ ] `SIRUS_AnalyserProduit` historique est remplacé sans casser ses appels.
- [ ] Taux de complétude affiché en pied de feuille.

---

## Phase 3 — Surfaces UI

### 3.1 Verrouillage des cellules `Auto` / `Herite`

Dans `SIRUS_GenererFeuilleAnalyse` (Phase 1), les colonnes `Source = "Auto"` ou `"Herite"` sont juste **grisées visuellement**. Pour les rendre **non éditables**, ajouter :

```vba
plage.Locked = True
ws.Protect Password:="SIRUS", UserInterfaceOnly:=True
```

Et inversement, déverrouiller explicitement les cellules `Saisie` avant le `Protect`. Permet à VBA d'écrire dans toutes les cellules (via `UserInterfaceOnly`) mais bloque l'utilisateur sur les `Auto/Herite`.

### 3.2 Snapshot Modifier pour le diff

Quand `operation = "Modifier"`, créer une feuille jumelle `SIRUS_<Table>_<Op>_Snap` (`veryHidden`) avec une copie des valeurs originales lues depuis la BdD. Au moment de l'envoi ATMAP, comparer ligne par ligne pour produire le diff `ValeurActuelle` / `NouvelleValeur` exigé par le contrat PJ.

```vba
' Dans PreRemplirDepuisBdD, après écriture des valeurs :
Set wsSnap = wb.Worksheets.Add(After:=ws)
wsSnap.Name = ws.Name & "_Snap"
ws.UsedRange.Copy
wsSnap.Range("A1").PasteSpecial xlPasteValues
wsSnap.Visible = xlSheetVeryHidden
```

### 3.3 Radar dans le UserForm

Onglet **Générer / Vérifier** existant : ajouter une ListView_Ecarts qui synthétise les écarts détectés à la dernière génération de feuille SIRUS. Clic sur une ligne → ouvre la feuille et sélectionne la cellule fautive (`Application.Goto`).

**Source des données** : `G_SIRUS_PBS.SIRUS_RemplirListViewGeneration` produit déjà une vue PBS↔BdD. À enrichir avec les écarts du moteur Phase 2.

---

## Phase 4 — Entrées A et B

### Parcours A — Consommer une ATM fournisseur

Les feuilles produites par `E_Extraction_ATM.CreerFeuillesExtractionATM` sont déjà **au format BdD** (`Info ATM` et `Tâches ATM` séparées). À brancher comme entrée SIRUS :

```vba
Public Sub SIRUS_AnalyserFeuilleExistante(ByVal sTable As String, _
                                          ByVal sNomFeuille As String)
    ' 1. Vérifier que la feuille existe et a le bon format
    ' 2. Croiser avec la BdD via ChargerSetDTR_BdD :
    '    - DTR déjà en base → Modification implicite
    '    - DTR absent → Création implicite
    ' 3. Appliquer le moteur SIRUS_AnalyserFeuille(table) dessus
End Sub
```

### Parcours B — Détection proactive des écarts PBS

À la génération des livrables (Main_ATM, Main_EMMST, Main_Providor), scanner `DATA_ASL` et identifier :
- DTR absents en BdD (à créer)
- Champs projet manquants (MTBF fonctionnel, Lead times, …)
- Divergences PBS↔BdD (Désignation, Node Type)

```vba
Public Sub SIRUS_DetecterEcartsPBS()
    ' Lit DATA_ASL, croise avec les 5 dicos BdD (déjà préchargés par
    ' G_SIRUS_PBS.SIRUS_RemplirListViewGeneration), affiche un récap des
    ' DTR à corriger dans la BdD avant génération.
    ' Bouton "Ouvrir SIRUS pour corriger" qui pré-remplit la feuille
    ' SIRUS_ProjetASL_Modifier avec les DTR concernés.
End Sub
```

---

## Phase 5 — Envoi ATMAP

`G_SIRUS_ATMAP.bas` est **fonctionnel** (validation Thomas) : `CreateTicketREST`, `AttachExcelToTicket`, `GetHeadersForTable`, `GetFormDigest`, `GetCurrentUserID`, parseurs JSON, `Test_Creer_Ticket_ATMAP`.

### Reste à faire

1. **Figer le contrat de la PJ** (brief v0.2 section 9) :
   - Bloc en-tête : `Table`, `Opération` (Créer/Modifier), `Demandeur`, `Date`, `Projet` (si ASL), `TauxComplétude`.
   - Bloc données : colonnes = `ColonneBdD` dans l'ordre `GetHeadersForTable(tableName)`. Pour `Modifier` : `ColonneBdD`, `ValeurActuelle`, `NouvelleValeur` (diff explicite via snapshot Phase 3.2).
   - Truncation **à 64 chars** pour matcher la limite ACE OLE DB côté Data Owner (cf. `TronquerColonneBdD`).

2. **Procédure `SIRUS_EnvoyerDemande(table)`** :
   ```vba
   Public Sub SIRUS_EnvoyerDemande(ByVal sTable As String)
       ' 1. Vérifier qu'une feuille SIRUS_<Table>_<Op> active existe
       ' 2. Vérifier qu'elle a passé l'analyse (taux MeQ > seuil métier)
       ' 3. Construire la PJ normalisée (Phase 5.1)
       ' 4. Sauvegarder PJ dans un .xlsx temporaire
       ' 5. CreateTicketREST(digest, userID, ...)
       ' 6. AttachExcelToTicket(ticketID, ..., PJ_path)
       ' 7. Supprimer le .xlsx temp
       ' 8. MsgBox de confirmation au demandeur avec ID du ticket
   End Sub
   ```

3. **Branchement UserForm** : bouton "Envoyer la demande" sur l'onglet SIRUS qui appelle `SIRUS_EnvoyerDemande(table active)`.

### Workflow de validation côté Data Owner (hors périmètre SIRUS)

1. Ticket reçu dans ATMAP (SharePoint List `ATMAP_TICKETS`).
2. Data Owner ouvre la PJ Excel attachée.
3. Vérifie les valeurs proposées (création) ou le diff (modification).
4. Copie-colle dans la BdD Excel partagée (geste manuel, point de contrôle de gouvernance).
5. Met le ticket en statut "Traité" dans ATMAP.

Le commentaire / rejet éventuel se fait dans ATMAP, pas dans la feuille SIRUS (le demandeur n'a pas la main sur la BdD).

---

## Phase 6 — Industrialisation

### 6.1 Journal des envois

Feuille `SIRUS_Historique` (`veryHidden` dans le `.xlam`) qui trace : Date, Demandeur, Table, Opération, DTRs concernés, TicketID, TauxMeQ. Permet à un demandeur de retrouver ses envois passés.

### 6.2 Paramètres globaux

Module `G_SIRUS_Parametres.bas` (ou réutiliser `SIRUS_Config` avec une table `__Params__`) :
- `SEUIL_TAUX_COMPLETUDE` (défaut 0.7)
- `BDD_TARGET` ("TRAVAIL" ou "PROD" — switcher pour les tests)
- `ATMAP_BASE_URL` (déjà en const dans `G_SIRUS_ATMAP`)
- `MOT_DE_PASSE_PROTECTION_FEUILLES` (pour le `Protect`)

### 6.3 Recette utilisateur

Plan de tests par phase :
- **Produit Créer** : 3 DTR (LRU + SRUNR + KIT) avec et sans prix.
- **Produit Modifier** : 2 DTR existants en BdD, modif d'1 champ chacun.
- **ProjetSLI Créer** : 1 projet complet avec interlocuteurs.
- **ProjetASL Créer** : 5 DTR sur 1 projet existant.
- **InfoATM + TachesATM** : 1 DTR avec 3 tâches rang 2, 3, 4.
- **EMMST** : 1 DTR avec dimensions + stockage.

À documenter dans un futur `Docs/SIRUS_RECETTE.md` ou directement dans `CHANGELOG.md`.

---

## Améliorations transverses (toutes phases)

### Cache global BdD

`G_SIRUS_PBS.ChargerSetDTR_BdD` est appelé 5 fois à chaque `SIRUS_RemplirListViewGeneration`. À chaque ouverture du UserForm, on relit toute la colonne DTR des 5 tables BdD. Possible amélioration : cache avec TTL (1h ?) et bouton **Rafraîchir** explicite dans l'UI.

### Conventions de nommage

- Modules SIRUS : préfixe `G_SIRUS_*.bas`.
- Procédures publiques SIRUS : préfixe `SIRUS_*`.
- Feuilles générées : préfixe `SIRUS_<Table>_<Op>`.
- Feuilles internes : `SIRUS_Config`, `SIRUS_Historique`, `SIRUS_Diag_*` (veryHidden).

### Internationalisation

Le projet a une logique FR/EN (cf. `langue` global). Pour SIRUS, libellés `Champ` du config sont actuellement en français sans accents. À voir si extension EN nécessaire (côté Data Owner).

### Tests automatisés

Procédures `Test_*` (déjà la convention du projet, cf. `Test_MapColonnesDATA_ASL`, `Test_EvaluerQualitePrix`, `Test_LireConfig`, `Test_GetListe`, `Test_Generer_Produit_Creer`, `Test_DiagnostiquerEcartsConfig_Toutes`, etc.). À étendre Phase 2+ : `Test_RegleObligatoire`, `Test_AnalyserFeuille`, `Test_EnvoyerDemande_Mock`.

---

## Décisions verrouillées (rappel)

1. **Prix HORS scope** de SIRUS. La fonction `EvaluerQualitePrix` de `G_SIRUS_PBS` n'est pas utilisée par le moteur d'analyse SIRUS.
2. **L'analyse ne bloque jamais** : statuts OK / Alerte, jamais "Bloquant". Taux de complétude pondéré + cas "vide justifié".
3. **Une feuille par (table × opération)**.
4. **Commentaire / rejet Data Owner = workflow ATMAP** (hors SIRUS).
5. **Listes en code** (`G_SIRUS_Listes`), pas en feuille.
6. **Règles en dur** (`G_SIRUS_Regles`), pas de mini-langage.
7. **`ColonneBdD` = vrai nom BdD** (avec accents UTF-8). Limite ACE 64 chars gérée par `TronquerColonneBdD`.
8. **BdD TRAVAIL** pour SIRUS, pas PROD.
9. **Feuilles SIRUS dans `ActiveWorkbook`** (classeur projet), pas dans le `.xlam`.
10. **Entête sur 2 lignes** : ligne 1 humaine + ligne 2 `ColonneBdD` masquée.

---

## Prochaine session — proposition d'ordre

1. **Phase 3.1** (verrouillage cellules Auto/Hérite) — ~30 min, gain UX immédiat.
2. **Phase 2** (moteur d'analyse `G_SIRUS_Regles` + `SIRUS_AnalyserFeuille`) — gros morceau, ~1-2 sessions.
3. **Phase 3.2** (snapshot Modifier pour le diff) — préparé en Phase 1, à finaliser.
4. **Phase 5** (contrat PJ + `SIRUS_EnvoyerDemande`) — ATMAP est prêt, c'est juste l'assemblage.
5. **Phase 4** (entrées A et B) — peut attendre que les phases 2-3-5 soient stabilisées.
