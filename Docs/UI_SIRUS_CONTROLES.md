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
| CommandButton | `Btn_SIRUS_Analyser` | "Analyser la feuille" |
| Label | `Lbl_SIRUS_StatutAnalyse` | "" (résumé OK/NOK/MeQ) |
| ListView | `ListView_SIRUS_Resultats` | ActiveX MSComctlLib, `View = lvwReport`, `FullRowSelect = True`, `Gridlines = True`, `LabelEdit = lvwManual` |

Colonnes attendues sur `ListView_SIRUS_Resultats` (configurées par `SIRUS_ConfigListViewResultats`) :
1. Ligne (40)
2. Clé (120)
3. Alertes (largeur restante)

---

## Zone 3 — Envoyer ATMAP

| Type | Name | Caption / Propriétés |
|---|---|---|
| Frame | `Frame_SIRUS_Envoi` | "3. Envoyer ATMAP" |
| Label | `Lbl_SIRUS_StatutEnvoi` | "" (résumé feuilles à envoyer) |
| CommandButton | `Btn_SIRUS_Envoyer` | "Envoyer la demande ATMAP" |

(Détails Phase 5 — pour l'instant le bouton sera désactivé tant que Phase 5 n'est pas branchée.)

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

Private Sub Btn_SIRUS_Analyser_Click()
    Call SIRUS_UI_BtnAnalyserClick
End Sub

Private Sub ListView_SIRUS_Resultats_DblClick()
    Call SIRUS_UI_LvResultatsDblClick
End Sub

Private Sub Btn_SIRUS_Envoyer_Click()
    Call SIRUS_UI_BtnEnvoyerClick
End Sub
```

**Note de nommage** : les handlers UI portent désormais le préfixe `SIRUS_UI_*` pour les distinguer clairement de l'API métier `SIRUS_GenererFeuilleAnalyse`, `SIRUS_AnalyserFeuille` (qui peuvent être appelées depuis du code, hors UserForm).

Et dans `UserForm_Initialize`, ajouter à la fin :

```vba
Call SIRUS_PhaseInit_v2
```
