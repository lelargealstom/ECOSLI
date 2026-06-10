# UI SIRUS — Contrôles à ajouter dans le UserForm

Document de référence pour Thomas. Contrôles à créer **manuellement dans le VBE Designer** (Excel ne permet pas de modifier `UserForm.frm` proprement par script). Les noms doivent être respectés à l'identique — c'est sur eux que `G_SIRUS_UserForm.bas` est câblé.

Placer le tout dans un **`Frame_SIRUS_v2`** (ou MultiPage existant, onglet "SIRUS"). Au choix.

---

## Zone 1 — Générer la feuille d'analyse

### Bloc Opération

| Type | Name | Caption | Group |
|---|---|---|---|
| OptionButton | `Opt_SIRUS_Creer` | "Créer" | `grpOp_SIRUS` |
| OptionButton | `Opt_SIRUS_Modifier` | "Modifier" | `grpOp_SIRUS` |
| OptionButton | `Opt_SIRUS_Extraction` | "Extraction (depuis livrable)" | `grpOp_SIRUS` |

→ Mettre `Opt_SIRUS_Creer.Value = True` par défaut. Les 3 partagent le même `GroupName` pour exclusivité.

### Bloc Table

| Type | Name | Caption |
|---|---|---|
| Label | `Lbl_SIRUS_Table` | "Table :" |
| ComboBox | `Cmb_SIRUS_Table` | (vide) |

La liste est peuplée dynamiquement par `SIRUS_Op_Click` (côté code).

### Bloc Modifier (visible si Op = Modifier)

| Type | Name | Caption / Propriétés |
|---|---|---|
| Frame | `Frame_SIRUS_Modifier` | "Clés à modifier" |
| Label | `Lbl_SIRUS_Cles` | "Clés DTR (une par ligne) :" |
| TextBox | `Txt_SIRUS_Cles` | `MultiLine = True`, `ScrollBars = fmScrollBarsVertical`, hauteur ~60 |
| Label | `Lbl_SIRUS_Projet` | "Projet :" |
| ComboBox | `Cmb_SIRUS_Projet` | (vide, alimentée à la 1re utilisation via `RecupProjetsList()`) |

**Logique d'affichage dynamique** (gérée par `SIRUS_MajZoneModifier`) :

| Op | Table | Visible |
|---|---|---|
| Modifier | ProjetSLI, ProjetASL | `Lbl_SIRUS_Projet` + `Cmb_SIRUS_Projet` |
| Modifier | autres (Produit, InfoATM, ...) | `Lbl_SIRUS_Cles` + `Txt_SIRUS_Cles` |

Justification : pour ProjetSLI/ProjetASL, la clé BdD est un **nom de projet** (pas un DTR), et l'utilisateur doit le choisir dans la liste existante en base. Pour les autres tables, on garde la saisie libre multi-DTR.

### Bloc Extraction (visible si Op = Extraction)

| Type | Name | Caption / Propriétés |
|---|---|---|
| Frame | `Frame_SIRUS_Extraction` | "Source du livrable" |
| OptionButton | `Opt_SIRUS_V1` | "V1", `GroupName = "grpVersionSIRUS"` |
| OptionButton | `Opt_SIRUS_V2` | "V2", `GroupName = "grpVersionSIRUS"`, `Value = True` |
| OptionButton | `Opt_SIRUS_FR` | "FR", `GroupName = "grpLangueSIRUS"`, `Value = True` |
| OptionButton | `Opt_SIRUS_EN` | "EN", `GroupName = "grpLangueSIRUS"` |
| Label | `Lbl_SIRUS_Fichier` | "Fichier :" |
| TextBox | `Txt_SIRUS_Fichier` | `Locked = True` (sera rempli par le bouton) |
| CommandButton | `Btn_SIRUS_Parcourir` | "Parcourir..." |
| Label | `Lbl_SIRUS_NbFiches` | "" (rempli après parcours) |

→ Les options Version V1/V2 ne sont pertinentes que si la table sélectionnée est "ATM". Le code masque automatiquement si EMMST sélectionné.

### Bouton de génération

| Type | Name | Caption |
|---|---|---|
| CommandButton | `Btn_SIRUS_Generer` | "Générer la feuille d'analyse" |

---

## Zone 2 — Analyser

| Type | Name | Caption / Propriétés |
|---|---|---|
| Frame | `Frame_SIRUS_Analyse` | "2. Analyser" |
| Label | `Lbl_SIRUS_FeuilleActive` | "Feuille active :" |
| ComboBox | `Cmb_SIRUS_FeuilleAnalyse` | (rempli par le code après génération) |
| CommandButton | `Btn_SIRUS_EditerFeuille` | "Éditer cette feuille" (active la feuille + masque le UserForm) |
| CommandButton | `Btn_SIRUS_Analyser` | "Analyser la feuille" |
| Label | `Lbl_SIRUS_StatutAnalyse` | "" (résumé OK/NOK/MeQ) |
| ListView | `ListView_SIRUS_Resultats` | ActiveX MSComctlLib, `View = lvwReport`, `FullRowSelect = True`, `Gridlines = True`, `LabelEdit = lvwManual` |

> ⚠️ **Important** : la propriété `ShowModal` du UserForm doit être à **`False`** (VBE → UserForm → F4 Properties → ShowModal = False) pour permettre l'interaction libre avec Excel pendant que le UserForm est ouvert. Le `.Show vbModeless` au lancement (W_AppelUsf) ne suffit pas si la propriété designer est restée à True.
> Le bouton `Btn_SIRUS_EditerFeuille` est une porte de sortie pragmatique : il masque le UserForm + active la feuille. Retour via `Ctrl+W`.

Colonnes attendues sur `ListView_SIRUS_Resultats` (configurées par `SIRUS_ConfigListViewResultats`) :
1. Ligne (40)
2. Clé (120)
3. Alertes (largeur restante)

---

## Zone 3 — Envoyer ATMAP (Phase 5)

### Bloc Récap (lecture seule, auto-rempli)

| Type | Name | Source |
|---|---|---|
| Frame | `Frame_SIRUS_Envoi` | "3. Envoyer ATMAP" |
| Label | `Lbl_SIRUS_Demandeur` | `Application.UserName` |
| Label | `Lbl_SIRUS_Date` | `Format(Now, "dd/mm/yyyy")` |
| Label | `Lbl_SIRUS_OpRecap` | Recopie l'Op active (Créer/Modifier/Extraction) |
| Label | `Lbl_SIRUS_TableRecap` | Recopie la Table active |
| Label | `Lbl_SIRUS_FeuillesRecap` | Liste feuilles + nb lignes + MeQ |

### Bloc Saisie (pré-rempli, modifiable)

| Type | Name | Pré-rempli | Obligatoire |
|---|---|---|---|
| TextBox | `Txt_SIRUS_Sujet` | `IMP-YYYY-00X - Import <Op> <Table>` | ✅ |
| TextBox `MultiLine=True` | `Txt_SIRUS_Description` | `<Op> de <N> référence(s) en <Table> pour le projet <Projet>` | ✅ |
| ComboBox | `Cmb_SIRUS_Priorite` | "Medium" par défaut | ✅ |
| TextBox `MultiLine=True` | `Txt_SIRUS_Attentes` | "Traitement sous X jours" selon priorité | non |

**Liste Priorité** : `High`, `Medium`, `Low`.

**Pré-remplissage Attentes** selon Priorité (re-calculé à chaque changement de `Cmb_SIRUS_Priorite`) :
- `High`   → `Traitement sous 1 jour`
- `Medium` → `Traitement sous 3 jours`
- `Low`    → `Traitement sous 5 jours`

### Bloc Action

| Type | Name | Caption |
|---|---|---|
| CommandButton | `Btn_SIRUS_Envoyer` | "Envoyer la demande ATMAP" |
| Label | `Lbl_SIRUS_StatutEnvoi` | "" (rempli après envoi : ID ticket + couleur verte/rouge) |

### Logique métier

**Numéro de ticket `IMP-YYYY-00X`** : généré automatiquement par requête SharePoint sur la liste `ID_Import` :
- Filtre `Typedemande = "Import"`
- Récupère le `ID_Import` le plus haut **de l'année en cours**
- Incrémente de 1
- Format : `IMP-2026-001`, `IMP-2026-002`, ...
- **Passage d'année** : si aucun ID trouvé pour l'année en cours, recommence à `001`.

**Projet** : récupéré depuis `Project_List` (combobox principale du UserForm en haut). Si vide → MsgBox bloquante demandant à l'utilisateur de choisir un projet avant envoi.

**Multi-feuilles (cas Extraction ATM)** :
- L'extraction ATM produit 2 feuilles : `SIRUS_InfoATM_Extraction` + `SIRUS_TachesATM_Extraction`.
- Côté envoi : **1 seul ticket** avec **1 PJ Excel unique** contenant les 2 feuilles d'analyse dans un même classeur (logique : 1 livrable = 1 demande d'import).
- Extraction EMMST : 1 feuille → 1 PJ.

**Mapping vers l'API ATMAP** (`CreateTicketREST`) :
- `numeroTicket` ← `Txt_SIRUS_Sujet` (préfixe `IMP-YYYY-00X`)
- `sujet` ← `Txt_SIRUS_Sujet`
- `statut` ← `"Nouveau"` (en dur)
- `description` ← `Txt_SIRUS_Description`
- `attentes` ← `Txt_SIRUS_Attentes`
- `typedemande` ← `"Import"` (en dur, c'est le filtre SharePoint pour l'auto-numérotation)
- `priorite` ← `Cmb_SIRUS_Priorite.Value`
- `userID` ← `GetCurrentUserID(digest)` (helper existant `G_SIRUS_ATMAP`)

### Garde-fous

`Btn_SIRUS_Envoyer.Enabled = False` tant que :
- Aucune feuille analysée OU taux MeQ < seuil (déjà géré)
- `Project_List.Value` vide (côté validation à l'envoi : MsgBox)
- `Txt_SIRUS_Sujet` ou `Txt_SIRUS_Description` vides
- `Cmb_SIRUS_Priorite` non choisie

**Confirmation `MsgBox vbYesNo`** avant l'appel `CreateTicketREST` (envoi irréversible).

### Câblage événements (à ajouter dans le code-behind UserForm)

```vba
Private Sub Cmb_SIRUS_Priorite_Change()
    Call SIRUS_UI_PrioriteChange
End Sub

Private Sub Btn_SIRUS_Envoyer_Click()
    Call SIRUS_UI_BtnEnvoyerClick
End Sub
```

---

## Câblage des événements (à ajouter dans le code du UserForm)

Dans le code-behind du UserForm, ajouter les **stubs d'événements** qui appellent les procédures du module `G_SIRUS_UserForm` :

```vba
Private Sub Opt_SIRUS_Creer_Click()
    Call SIRUS_Op_Click("Creer")
End Sub

Private Sub Opt_SIRUS_Modifier_Click()
    Call SIRUS_Op_Click("Modifier")
End Sub

Private Sub Opt_SIRUS_Extraction_Click()
    Call SIRUS_Op_Click("Extraction")
End Sub

Private Sub Cmb_SIRUS_Table_Change()
    Call SIRUS_Table_Change
End Sub

Private Sub Btn_SIRUS_Parcourir_Click()
    Call SIRUS_UI_BtnParcourirClick
End Sub

Private Sub Btn_SIRUS_Generer_Click()
    Call SIRUS_UI_BtnGenererClick
End Sub

Private Sub Btn_SIRUS_EditerFeuille_Click()
    Call SIRUS_UI_BtnEditerFeuilleClick
End Sub

Private Sub Btn_SIRUS_Analyser_Click()
    Call SIRUS_UI_BtnAnalyserClick
End Sub

Private Sub ListView_SIRUS_Resultats_DblClick()
    Call SIRUS_UI_LvResultatsDblClick
End Sub

Private Sub Btn_SIRUS_Envoyer_Click()
    Call SIRUS_UI_BtnEnvoyerClick
End Sub

' Met a jour automatiquement la Description ATMAP quand le user change
' de projet dans la liste principale (apres une analyse OK).
Private Sub Project_List_Change()
    Call SIRUS_UI_ProjetListChange
End Sub
```

## Affichage progressif des frames

Au démarrage du UserForm (`SIRUS_PhaseInit_v2`) :
- `Frame_SIRUS_Modifier.Visible = False`
- `Frame_SIRUS_Extraction.Visible = False`
- **`Frame_SIRUS_Analyse.Visible = False`** ← caché tant qu'aucune feuille n'a été générée
- **`Frame_SIRUS_Envoi.Visible = False`** ← caché tant que l'analyse n'a pas atteint le seuil MeQ

Progression :
- **Après Génération** → `Frame_SIRUS_Analyse` devient visible (Zone 2)
- **Après Analyse avec MeQ ≥ 70%** → `Frame_SIRUS_Envoi` devient visible (Zone 3) + pré-remplissage auto

## ComboBox Priorité

Au démarrage du UserForm (`SIRUS_PhaseInit_v2`), la ComboBox est peuplée auto :
- `Style = fmStyleDropDownList` (= 2) → force la sélection dans la liste, pas de saisie libre
- Items : `High`, `Medium`, `Low`

Le changement de valeur déclenche `SIRUS_UI_PrioriteChange` qui met à jour `Txt_SIRUS_Attentes` (1/3/5 jours).

## Freeze panes sur les feuilles SIRUS_*

`SIRUS_GenererFeuilleAnalyse` fige automatiquement les colonnes selon la table :

| Table | Colonnes figées | Détail |
|---|---|---|
| Produit, InfoATM, TachesATM, EMMST | 1-2 | Statut + DTR |
| ProjetSLI | 1-2 | Statut + Nom projet |
| ProjetASL | 1-3 | Statut + Projet + DTR |

## Coloration ListView_SIRUS_Resultats

Les lignes NOK sont affichées en **rouge gras** (RGB(163, 45, 45)). Permet à l'ingé ASL de voir immédiatement les lignes à corriger.

**Note de nommage** : les handlers UI portent désormais le préfixe `SIRUS_UI_*` pour les distinguer clairement de l'API métier `SIRUS_GenererFeuilleAnalyse`, `SIRUS_AnalyserFeuille` (qui peuvent être appelées depuis du code, hors UserForm).

Et dans `UserForm_Initialize`, ajouter à la fin :

```vba
Call SIRUS_PhaseInit_v2
```
