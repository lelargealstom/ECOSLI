Attribute VB_Name = "G_SIRUS_UserForm"

' -- Constante : colonnes de la ListBox Produit -------------
' Col 0 = DTR | Col 1 = Designation | Col 2 = Node Type
' Col 3 = Fournisseur | Col 4 = Ref. Fourn. | Col 5 = Drawing Nb

' -- Collection de stockage ---------------------------------
Private lignesProduit As Collection
Public Sub SIRUS_Init()
    Set lignesProduit = New Collection
    UserForm.ListBox_Produit.Clear
End Sub

' ============================================================
' SIRUS_InitGeneration
' Point d'entree unique appele par UserForm_Initialize :
'   - configure les 4 ListView SIRUS
'   - active le lookup BdD
'   - remplit la ListView_Generation depuis DATA_ASL + BdD
' Regroupe ce qui etait des lignes eparpillees dans Initialize
' (risque d'oublier de decommenter une ligne intermediaire).
' ============================================================
Public Sub SIRUS_InitGeneration()
    Call SIRUS_ConfigListViewProduit
    Call SIRUS_ConfigListViewTaches
    Call SIRUS_ConfigListViewATM
    Call SIRUS_ConfigListViewGeneration
    G_SIRUS_PBS.SIRUS_LookupDTR_Active = True
    Call G_SIRUS_PBS.SIRUS_RemplirListViewGeneration
End Sub
' ============================================================
' CONFIGURATION LISTVIEW_PRODUIT a dans UserForm_Initialize
' ============================================================
Public Sub SIRUS_ConfigListViewProduit()

    Dim lv As MSComctlLib.ListView
    Set lv = UserForm.ListView_Produit

    ' -- Vue rapport (tableau) ------------------------------
    lv.View = lvwReport
    lv.FullRowSelect = True
    lv.Gridlines = True
    lv.LabelEdit = lvwManual

    ' -- Vider les colonnes existantes ---------------------
    lv.ColumnHeaders.Clear
    lv.ListItems.Clear

    ' -- Creer les colonnes avec largeurs ------------------
    lv.ColumnHeaders.Add , , "Type", 60
    lv.ColumnHeaders.Add , , "DTR", 80
    lv.ColumnHeaders.Add , , "Designation", 150
    lv.ColumnHeaders.Add , , "Node Type", 80
    lv.ColumnHeaders.Add , , "Fournisseur", 100
    lv.ColumnHeaders.Add , , "Ref. Fournisseur", 100
    lv.ColumnHeaders.Add , , "Drawing Number", 100
    lv.ColumnHeaders.Add , , "MTBF Intr.", 75
    lv.ColumnHeaders.Add , , "DVL HARD", 70
    lv.ColumnHeaders.Add , , "DVL SOFT", 70

End Sub
' ============================================================
' CONFIGURATION LISTVIEW_PRODUIT a dans UserForm_Initialize
' ============================================================
Public Sub SIRUS_ConfigListViewGeneration()

    Dim lv As MSComctlLib.ListView
    Set lv = UserForm.ListView_Generation

    ' -- Vue rapport (tableau) ------------------------------
    lv.View = lvwReport
    lv.FullRowSelect = True
    lv.Gridlines = True
    lv.LabelEdit = lvwManual

    ' -- Vider les colonnes existantes ---------------------
    lv.ColumnHeaders.Clear
    lv.ListItems.Clear

    ' -- Creer les colonnes avec largeurs ------------------
    ' 10 colonnes : DTR + 5 MeQ tables + 2x (Prix, Source)
    ' Cascade prix : Spare = Frs -> Cat D&IS -> FCS calcule -> Archive
    '                Repair = Frs -> Cat D&IS -> Archive
    lv.ColumnHeaders.Add , , "DTR", 80
    lv.ColumnHeaders.Add , , "Nb projets", 70
    lv.ColumnHeaders.Add , , "ATM - Info", 70
    lv.ColumnHeaders.Add , , "ATM - Taches", 80
    lv.ColumnHeaders.Add , , "EMMST", 60
    lv.ColumnHeaders.Add , , "Produit", 60
    lv.ColumnHeaders.Add , , "Prix Spare", 80
    lv.ColumnHeaders.Add , , "Source Spare", 150
    lv.ColumnHeaders.Add , , "Prix Repair", 80
    lv.ColumnHeaders.Add , , "Source Repair", 150

End Sub
' ============================================================
' CONFIGURATION LISTVIEW_TACHESATM a dans UserForm_Initialize
' ============================================================
Public Sub SIRUS_ConfigListViewTaches()

    Dim lv As MSComctlLib.ListView
    Set lv = UserForm.ListView_Taches

    ' -- Vue rapport (tableau) ------------------------------
    lv.View = lvwReport
    lv.FullRowSelect = True
    lv.Gridlines = True
    lv.LabelEdit = lvwManual

    ' -- Vider les colonnes existantes ---------------------
    lv.ColumnHeaders.Clear
    lv.ListItems.Clear

    ' -- Creer les colonnes avec largeurs ------------------
    lv.ColumnHeaders.Add , , "DTR", 80
    lv.ColumnHeaders.Add , , "Langue", 40
    lv.ColumnHeaders.Add , , "Rang", 40
    lv.ColumnHeaders.Add , , "Type", 40
    lv.ColumnHeaders.Add , , "Categorie", 60
    lv.ColumnHeaders.Add , , "Niv. Qualif", 60
    lv.ColumnHeaders.Add , , "Description", 100
    lv.ColumnHeaders.Add , , "Criteres Dim", 60
    lv.ColumnHeaders.Add , , "Out. Spe", 60
    lv.ColumnHeaders.Add , , "Out. Standard", 60
    lv.ColumnHeaders.Add , , "Pieces Rech", 60
    lv.ColumnHeaders.Add , , "Nb Operateurs", 40
    lv.ColumnHeaders.Add , , "Tps estima (h)", 80
    lv.ColumnHeaders.Add , , "Picto", 60
    lv.ColumnHeaders.Add , , "Texte Picto", 80

End Sub

' ============================================================
' CONFIGURER LA LISTVIEW D'ERREURS
' ============================================================

Public Sub SIRUS_ConfigListViewErreurs()

    Dim lv As MSComctlLib.ListView
    Set lv = UserForm.ListView_Erreurs

    lv.View = lvwReport
    lv.FullRowSelect = True
    lv.Gridlines = True
    lv.LabelEdit = lvwManual

    lv.ColumnHeaders.Clear
    lv.ListItems.Clear

    lv.ColumnHeaders.Add , , "Ligne", 45
    lv.ColumnHeaders.Add , , "DTR", 90
    lv.ColumnHeaders.Add , , "Controle", 110
    lv.ColumnHeaders.Add , , "Action corrective", 300

End Sub
' ============================================================
' CONFIGURATION LISTVIEW_ATM a dans UserForm_Initialize
' ============================================================
Public Sub SIRUS_ConfigListViewATM()

    Dim lv As MSComctlLib.ListView
    Set lv = UserForm.ListView_ATM

    ' -- Vue rapport (tableau) ------------------------------
    lv.View = lvwReport
    lv.FullRowSelect = True
    lv.Gridlines = True
    lv.LabelEdit = lvwManual

    ' -- Vider les colonnes existantes ---------------------
    lv.ColumnHeaders.Clear
    lv.ListItems.Clear

    ' -- Creer les colonnes avec largeurs ------------------
    lv.ColumnHeaders.Add , , "Table", 40
    lv.ColumnHeaders.Add , , "DTR", 60
    lv.ColumnHeaders.Add , , "Descr_FR Tache", 80
    lv.ColumnHeaders.Add , , "Descr_EN Tache", 80
    lv.ColumnHeaders.Add , , "Reparabilite", 40
    lv.ColumnHeaders.Add , , "Type", 40
    lv.ColumnHeaders.Add , , "Poids", 20
    lv.ColumnHeaders.Add , , "Tps_gen_HARD(h)", 100
    lv.ColumnHeaders.Add , , "Tps_gen_SOFT(h)", 60
    lv.ColumnHeaders.Add , , "Tps_ana_cumule(h)", 60
    lv.ColumnHeaders.Add , , "Out. Standard", 60
    lv.ColumnHeaders.Add , , "Pieces Rech", 60
    lv.ColumnHeaders.Add , , "Nb Operateurs", 40
    lv.ColumnHeaders.Add , , "Tps estima (h)", 40
    lv.ColumnHeaders.Add , , "Picto", 60
    lv.ColumnHeaders.Add , , "Texte Picto", 80

End Sub
' ============================================================
' AJOUTER UNE LIGNE DANS LISTVIEW_PRODUIT
' ============================================================
Public Sub SIRUS_AddToListViewProduit(vals() As String, TypeMod As String)

    Dim lv   As MSComctlLib.ListView
    Dim item As MSComctlLib.ListItem

    Set lv = UserForm.ListView_Produit

    ' Creer la ligne avec la valeur de la premiere colonne (DTR)
    Set item = lv.ListItems.Add(, , TypeMod)

    ' Ajouter les sous-elements (colonnes 2 a 10)
    item.SubItems(1) = vals(1)  ' Designation
    item.SubItems(2) = vals(2)  ' Designation
    item.SubItems(3) = vals(3)  ' Node Type
    item.SubItems(4) = vals(4)  ' Fournisseur
    item.SubItems(5) = vals(5)  ' Ref. Fournisseur
    item.SubItems(6) = vals(6)  ' Drawing Number
    item.SubItems(7) = vals(7)  ' MTBF Intr.
    item.SubItems(8) = vals(8)  ' DVL HARD
    item.SubItems(9) = vals(9)  ' DVL SOFT

End Sub

' ============================================================
' CHANGEMENT TYPE DEMANDE (NavTab Creation/Modif/Suppression)
' ============================================================
Public Sub SIRUS_SetType(typeVal As String)

    With UserForm
        Select Case typeVal

            Case "Creation"
                ' DTR = TextBox libre, blanc
                .DTRProduit_txt.Enabled = True
                .DTRProduit_txt.BackColor = RGB(255, 255, 255)
                ' Tous les champs actifs
                .Designation_txt.Enabled = True
                .NodeType_list.Enabled = True
                .FRS_list.Enabled = True
                .RefFRS_txt.Enabled = True
                .DrawingNb_txt.Enabled = True
                .MTBF_Intr_txt.Enabled = True
                .DVL_HARD_txt.Enabled = True
                .DVL_SOFT_txt.Enabled = True

            Case "Modification"
                ' DTR = lookup BdD (fond bleu clair)
                .DTRProduit_txt.Enabled = True
                .DTRProduit_txt.BackColor = RGB(214, 228, 241)
                ' Tous les champs actifs a pre-remplis depuis BdD apres saisie DTR
                .Designation_txt.Enabled = True
                .NodeType_list.Enabled = True
                .FRS_list.Enabled = True
                .RefFRS_txt.Enabled = True
                .DrawingNb_txt.Enabled = True
                .MTBF_Intr_txt.Enabled = True
                .DVL_HARD_txt.Enabled = True
                .DVL_SOFT_txt.Enabled = True

            Case "Suppression"
                ' DTR seul requis, tout le reste grise
                .DTRProduit_txt.Enabled = True
                .DTRProduit_txt.BackColor = RGB(255, 235, 235)
                .Designation_txt.Enabled = False
                .NodeType_list.Enabled = False
                .FRS_list.Enabled = False
                .RefFRS_txt.Enabled = False
                .DrawingNb_txt.Enabled = False
                .MTBF_Intr_txt.Enabled = False
                .DVL_HARD_txt.Enabled = False
                .DVL_SOFT_txt.Enabled = False

        End Select
    End With

End Sub
' ============================================================
' BOUTON "+ Ajouter la ligne"
' ============================================================
Public Sub SIRUS_AjouterLigneProduit(TypeMod As String)

    If lignesProduit Is Nothing Then Set lignesProduit = New Collection

    Dim dtr   As String: dtr = Trim(UserForm.DTRProduit_txt.Value)
    Dim des   As String: des = Trim(UserForm.Designation_txt.Value)
    Dim node  As String: node = Trim(UserForm.NodeType_list.Value)
    Dim fourn As String: fourn = Trim(UserForm.FRS_list.Value)
    Dim ref   As String: ref = Trim(UserForm.RefFRS_txt.Value)
    Dim drw   As String: drw = Trim(UserForm.DrawingNb_txt.Value)
    Dim MTBF  As String: MTBF = Trim(UserForm.MTBF_Intr_txt.Value)
    Dim dvlh  As String: dvlh = Trim(UserForm.DVL_HARD_txt.Value)
    Dim dvls  As String: dvls = Trim(UserForm.DVL_SOFT_txt.Value)

    ' -- Validation selon type de demande -------------------
    Dim typeActif As String
    typeActif = SIRUS_GetTypeActif()

    Dim missing As String
    missing = ""

    If typeActif = "Suppression" Then
        If dtr = "" Then missing = "DTR"
    Else
        If dtr = "" Then missing = missing & "DTR / "
        If des = "" Then missing = missing & "Designation / "
        If node = "" Then missing = missing & "Node Type / "
        If fourn = "" Then missing = missing & "Fournisseur / "
        If ref = "" Then missing = missing & "Ref. Fournisseur / "
        If drw = "" Then missing = missing & "Drawing Number / "
        If MTBF = "" Then missing = missing & "MTBF Intrinseque / "
        If dvlh = "" Then missing = missing & "DVL HARD / "
    End If

    If missing <> "" Then
        MsgBox "Champs obligatoires manquants : " & _
               Left(missing, Len(missing) - 3), vbExclamation, "SIRUS"
        Exit Sub
    End If

    ' -- Verifier doublon DTR dans la collection ------------
    Dim item As Variant
    For Each item In lignesProduit
        If IsArray(item) Then
            If UCase(Trim(item(0))) = UCase(dtr) Then
                MsgBox "DTR '" & dtr & "' deja present dans la liste.", _
                       vbExclamation, "SIRUS"
                Exit Sub
            End If
        End If
    Next item

    ' -- Ajouter a la collection ----------------------------
    Dim vals(9) As String
    vals(0) = dtr: vals(1) = des: vals(2) = node
    vals(3) = fourn: vals(4) = ref: vals(5) = drw
    vals(6) = MTBF: vals(7) = dvlh: vals(8) = dvls

    lignesProduit.Add vals

    Call SIRUS_AddToListViewProduit(vals, TypeMod)

    ' -- Vider les champs apres ajout ----------------------
    Call SIRUS_EffacerChampsProduit

    ' -- Mettre le focus sur DTR pour saisie suivante -------
    UserForm.DTRProduit_txt.SetFocus

End Sub
' ============================================================
' BOUTON "Effacer"
' ============================================================
Public Sub SIRUS_EffacerChampsProduit()

    With UserForm
        .DTRProduit_txt.Value = ""
        .Designation_txt.Value = ""
        .NodeType_list.Value = ""
        .FRS_list.Value = ""
        .RefFRS_txt.Value = ""
        .DrawingNb_txt.Value = ""
        .MTBF_Intr_txt.Value = ""
        .DVL_HARD_txt.Value = ""
        .DVL_SOFT_txt.Value = ""
    End With

End Sub
' ============================================================
' BOUTON "Effacer la selection"
' ============================================================
Public Sub SIRUS_EffacerSelectionProduitProduit()

    Dim lb  As MSForms.ListBox
    Dim idx As Long

    Set lb = UserForm.ListBox_Produit
    idx = lb.ListIndex

    If idx = -1 Then
        MsgBox "Selectionnez une ligne a supprimer.", vbExclamation, "SIRUS"
        Exit Sub
    End If

    ' Reconstruire la collection sans la ligne supprimee
    Dim newCol As New Collection
    Dim i      As Long
    Dim item   As Variant

    i = 0
    For Each item In lignesProduit
        If i <> idx Then newCol.Add item
        i = i + 1
    Next item

    Set lignesProduit = newCol

    ' Supprimer de la ListBox
    lb.RemoveItem idx

End Sub

' ============================================================
' AJOUTER UNE TeCHE DANS LISTVIEW_TACHES
' ============================================================
Public Sub SIRUS_AjouterTacheATM()

    ' -- Recuperation des valeurs ---------------------------
    Dim dtr     As String: dtr = Trim(UserForm.DTRATM_txt.Value)
    Dim langue  As String: langue = Trim(UserForm.LangueATM_list.Value)
    Dim rang    As String: rang = Trim(UserForm.RangATM_list.Value)
    Dim hs      As String: hs = Trim(UserForm.TypeATM_list.Value)
    Dim cat     As String: cat = Trim(UserForm.CategorieATM_txt.Value)
    Dim nq      As String: nq = Trim(UserForm.NivQualifATM_txt.Value)
    Dim desc    As String: desc = Trim(UserForm.DescriptionTachesATM_txt.Value)
    Dim critdim As String: critdim = Trim(UserForm.CritDimATM_txt.Value)
    Dim outspec As String: outspec = Trim(UserForm.OutSpecATM_txt.Value)
    Dim outstd  As String: outstd = Trim(UserForm.OutStandardATM_txt.Value)
    Dim piecesR As String: piecesR = ""   ' a brancher si tu as le champ
    Dim nbops   As String: nbops = ""     ' a brancher si tu as le champ
    Dim tps     As String: tps = Trim(UserForm.TpsTachesATM_txt.Value)
    Dim picto   As String: picto = Trim(UserForm.PictoATM_list.Value)
    Dim txtpic  As String: txtpic = Trim(UserForm.TxtPictoATM_txt.Value)

    ' -- Validation champs obligatoires --------------------
    Dim missing As String
    missing = ""

    If dtr = "" Then missing = missing & "DTR / "
    If langue = "" Then missing = missing & "Langue / "
    If rang = "" Then missing = missing & "Rang / "
    If hs = "" Then missing = missing & "Type (HARD/SOFT) / "
    If cat = "" Then missing = missing & "Categorie / "
    If nq = "" Then missing = missing & "Niveau qualif / "
    If desc = "" Then missing = missing & "Description / "
    If tps = "" Then missing = missing & "Temps estima / "
    If picto = "" Then missing = missing & "Picto / "
    If txtpic = "" Then missing = missing & "Texte Picto / "

    ' Outillage standard toujours obligatoire
    If outstd = "" Then missing = missing & "Out. Standard / "

    ' Outillage specifique obligatoire si rang 2-1 ou 5-1
    If (rang = "2-1" Or rang = "5-1") And outspec = "" Then
        missing = missing & "Out. Spec. (obligatoire rang " & rang & ") / "
    End If

    ' Criteres dimensionnels a obligatoire si description contient
    ' "Torque", "Tighten" (EN) ou "Serrer" + "couple" (FR)
    Dim descLower As String
    descLower = LCase(desc)
    Dim needCrit As Boolean
    needCrit = (InStr(descLower, "torque") > 0 Or _
                InStr(descLower, "tighten") > 0 Or _
                (InStr(descLower, "serrer") > 0 And InStr(descLower, "couple") > 0))

    If needCrit And critdim = "" Then
        missing = missing & "Criteres dim. (serrage detecte) / "
    End If

    If missing <> "" Then
        MsgBox "Champs obligatoires manquants :" & vbCrLf & _
               Left(missing, Len(missing) - 3), vbExclamation, "SIRUS a ATM"
        Exit Sub
    End If

    ' -- Verification doublon DTR + Rang + Langue ----------
    Dim lv   As MSComctlLib.ListView
    Dim li   As MSComctlLib.ListItem

    Set lv = UserForm.ListView_Taches

    For Each li In lv.ListItems
        If UCase(li.Text) = UCase(dtr) And _
           UCase(li.SubItems(2)) = UCase(rang) And _
           UCase(li.SubItems(1)) = UCase(langue) Then
            MsgBox "La combinaison DTR '" & dtr & "' / Rang '" & rang & _
                   "' / Langue '" & langue & "' existe deja.", _
                   vbExclamation, "SIRUS a ATM"
            Exit Sub
        End If
    Next li

    ' -- Ajouter la ligne dans le ListView -----------------
    Set li = lv.ListItems.Add(, , dtr)

    li.SubItems(1) = langue
    li.SubItems(2) = rang
    li.SubItems(3) = hs
    li.SubItems(4) = cat
    li.SubItems(5) = nq
    li.SubItems(6) = desc
    li.SubItems(7) = critdim
    li.SubItems(8) = outspec
    li.SubItems(9) = outstd
    li.SubItems(10) = piecesR
    li.SubItems(11) = nbops
    li.SubItems(12) = tps
    li.SubItems(13) = picto
    li.SubItems(14) = txtpic

    

    ' -- Vider les champs de saisie -------------------------
    Call SIRUS_EffacerChampsTaches

End Sub

' ============================================================
' EFFACER LES CHAMPS DE SAISIE TACHES
' ============================================================

Public Sub SIRUS_EffacerChampsTaches()

    With UserForm
        .RangATM_list.Value = ""
        .CategorieATM_txt.Value = ""
        .TpsTachesATM_txt.Value = ""
        .DescriptionTachesATM_txt.Value = ""
        .OutStandardATM_txt.Value = ""
        .OutSpecATM_txt.Value = ""
        .PictoATM_list.Value = ""
        .TxtPictoATM_txt.Value = ""
    End With

End Sub

' ============================================================
' SUPPRIMER LA LIGNE SELECTIONNeE DANS LISTVIEW_TACHES
' ============================================================

Public Sub SIRUS_SupprimerTacheATM()

    Dim lv As MSComctlLib.ListView
    Set lv = UserForm.ListView_Taches

    If lv.SelectedItem Is Nothing Then
        MsgBox "Selectionnez une teche a supprimer.", vbExclamation, "SIRUS"
        Exit Sub
    End If

    lv.ListItems.Remove lv.SelectedItem.index

End Sub
' ============================================================
' RECUPeRER LE TYPE DE DEMANDE ACTIF (NavTab)
' ============================================================
Private Function SIRUS_GetTypeActif() As String

    Select Case UserForm.Type_Import_Tab.SelectedItem.index
        Case 0: SIRUS_GetTypeActif = "Creation"
        Case 1: SIRUS_GetTypeActif = "Modification"
        Case 2: SIRUS_GetTypeActif = "Suppression"
        Case Else: SIRUS_GetTypeActif = "Creation"
    End Select

End Function

' =========================================================================
' =========================================================================
' UI SIRUS v2 (Phase 2/3.2/5) - cablage des nouveaux controles
' =========================================================================
' Cf. Docs/UI_SIRUS_CONTROLES.md pour la liste des controles a creer dans
' le UserForm. Les noms utilises ici doivent matcher ceux du Designer.
'
' Workflow utilisateur :
'   1. Choisir Op (Creer/Modifier/Extraction) + Table (+ fichier si Extraction)
'   2. Bouton Generer -> feuille SIRUS_<Table>_<Op> creee
'   3. Bouton Analyser -> Statut OK/NOK ecrit + ListView remplie des NOK
'   4. Bouton Envoyer -> ticket ATMAP (Phase 5, TODO)
' =========================================================================

' Note : la liste des feuilles generees est ecrite directement dans
' Cmb_SIRUS_FeuilleAnalyse a la generation (pas de variable d'etat
' module pour eviter les ennuis avec les variables Object/Collection
' privees qui peuvent etre dans un etat instable apres recompilation
' du projet VBA).

' -------------------------------------------------------------------------
' SIRUS_PhaseInit_v2 : a appeler dans UserForm_Initialize
' -------------------------------------------------------------------------
Public Sub SIRUS_PhaseInit_v2()
    On Error Resume Next

    ' Op par defaut : Creer
    UserForm.Opt_SIRUS_Creer.Value = True

    ' Combo Langue / Version par defaut
    UserForm.Opt_SIRUS_V2.Value = True
    UserForm.Opt_SIRUS_FR.Value = True

    ' Init combo Table (depend de l'Op)
    Call SIRUS_Op_Click("Creer")

    ' Sous-zones cachees au depart
    UserForm.Frame_SIRUS_Modifier.Visible = False
    UserForm.Frame_SIRUS_Extraction.Visible = False

    ' Combo Projet / TextBox DTR : etat par defaut (caches, basculement
    ' gere par SIRUS_MajZoneModifier lors du choix Modifier + Table).
    On Error Resume Next
    UserForm.Lbl_SIRUS_Projet.Visible = False
    UserForm.Cmb_SIRUS_Projet.Visible = False
    On Error GoTo 0

    ' Affichage progressif : Frame Analyse + Frame Envoi caches tant que
    ' l etape precedente n est pas faite.
    On Error Resume Next
    UserForm.Frame_SIRUS_Analyse.Visible = False
    UserForm.Frame_SIRUS_Envoi.Visible = False
    On Error GoTo 0

    ' Peupler la ComboBox Priorite (3 valeurs fixes) + style liste fermee
    ' (Style = 2 = fmStyleDropDownList : empeche la saisie libre, force
    ' la selection dans la liste).
    On Error Resume Next
    With UserForm.Cmb_SIRUS_Priorite
        .Clear
        .Style = 2   ' fmStyleDropDownList
        .AddItem "High"
        .AddItem "Medium"
        .AddItem "Low"
    End With
    On Error GoTo 0

    ' Scanner le classeur actif pour alimenter le combo des feuilles SIRUS
    ' (utile si feuilles deja generees dans la session). Si au moins une
    ' feuille trouvee, afficher la Frame Analyse pour permettre d analyser
    ' tout de suite sans regenerer.
    Call SIRUS_RafraichirComboFeuilles
    On Error Resume Next
    UserForm.Frame_SIRUS_Analyse.Visible = (UserForm.Cmb_SIRUS_FeuilleAnalyse.ListCount > 0)
    On Error GoTo 0

    ' ListView Resultats
    Call SIRUS_ConfigListViewResultats

    ' Bouton Envoyer desactive tant qu'il n'y a pas d'analyse OK
    UserForm.Btn_SIRUS_Envoyer.Enabled = False

    ' Vider le combo des feuilles d'analyse
    UserForm.Cmb_SIRUS_FeuilleAnalyse.Clear

    On Error GoTo 0
End Sub

' -------------------------------------------------------------------------
' SIRUS_Op_Click : appele par Opt_SIRUS_Creer/Modifier/Extraction_Click
' Met a jour la liste Table et la visibilite des sous-zones.
' -------------------------------------------------------------------------
Public Sub SIRUS_Op_Click(ByVal sOp As String)

    With UserForm.Cmb_SIRUS_Table
        .Clear
        Select Case sOp
            Case "Creer", "Modifier"
                .AddItem "Produit"
                .AddItem "InfoATM"
                .AddItem "TachesATM"
                .AddItem "EMMST"
                .AddItem "ProjetSLI"
                .AddItem "ProjetASL"
            Case "Extraction"
                .AddItem "ATM"
                .AddItem "EMMST"
        End Select
        If .ListCount > 0 Then .ListIndex = 0
    End With

    UserForm.Frame_SIRUS_Modifier.Visible = (sOp = "Modifier")
    UserForm.Frame_SIRUS_Extraction.Visible = (sOp = "Extraction")

    ' Adapter Modifier : Combo Projet vs TextBox DTR multiligne
    Call SIRUS_MajZoneModifier
End Sub

' -------------------------------------------------------------------------
' SIRUS_Table_Change : declenche la mise a jour des sous-zones :
'   - Version V1/V2 visible si Op=Extraction ET Table=ATM
'   - Combo Projet visible si Op=Modifier ET Table=ProjetSLI/ProjetASL
' -------------------------------------------------------------------------
Public Sub SIRUS_Table_Change()
    Dim bExtraction As Boolean
    Dim sTable As String

    bExtraction = UserForm.Opt_SIRUS_Extraction.Value
    sTable = CStr(UserForm.Cmb_SIRUS_Table.Value)

    On Error Resume Next
    UserForm.Opt_SIRUS_V1.Visible = (bExtraction And sTable = "ATM")
    UserForm.Opt_SIRUS_V2.Visible = (bExtraction And sTable = "ATM")
    On Error GoTo 0

    Call SIRUS_MajZoneModifier
End Sub

' -------------------------------------------------------------------------
' SIRUS_MajZoneModifier : selon Op = Modifier ET Table choisie, bascule
' entre la TextBox DTR multiligne et le Combo Projet (alimente depuis BdD).
'   - ProjetSLI / ProjetASL -> Combo Projet visible, TextBox cachee
'   - autres tables          -> TextBox visible, Combo Projet cache
' -------------------------------------------------------------------------
Private Sub SIRUS_MajZoneModifier()
    Dim bModifier As Boolean
    Dim sTable As String
    Dim bEstProjet As Boolean

    bModifier = UserForm.Opt_SIRUS_Modifier.Value
    sTable = CStr(UserForm.Cmb_SIRUS_Table.Value)
    bEstProjet = (sTable = "ProjetSLI" Or sTable = "ProjetASL")

    On Error Resume Next
    ' Combo Projet visible si Modifier + table Projet*
    UserForm.Lbl_SIRUS_Projet.Visible = (bModifier And bEstProjet)
    UserForm.Cmb_SIRUS_Projet.Visible = (bModifier And bEstProjet)
    ' TextBox DTR : visible si Modifier ET table NON Projet*
    UserForm.Lbl_SIRUS_Cles.Visible = (bModifier And Not bEstProjet)
    UserForm.Txt_SIRUS_Cles.Visible = (bModifier And Not bEstProjet)
    On Error GoTo 0

    ' Alimenter le combo Projet a la 1re utilisation
    If bModifier And bEstProjet Then
        Call SIRUS_AlimenterComboProjets
    End If
End Sub

' -------------------------------------------------------------------------
' SIRUS_AlimenterComboProjets : peuple Cmb_SIRUS_Projet depuis la BdD.
' Idempotent : ne recharge pas si deja peuple. Utilise RecupProjetsList
' (existant dans F_FonctionsGlobales).
' -------------------------------------------------------------------------
Private Sub SIRUS_AlimenterComboProjets()
    Dim arr As Variant
    Dim i As Long

    On Error Resume Next
    If UserForm.Cmb_SIRUS_Projet.ListCount > 0 Then Exit Sub
    On Error GoTo 0

    arr = RecupProjetsList()
    If IsEmpty(arr) Then Exit Sub

    With UserForm.Cmb_SIRUS_Projet
        .Clear
        For i = LBound(arr) To UBound(arr)
            If LenB(CStr(arr(i))) > 0 Then .AddItem CStr(arr(i))
        Next i
        If .ListCount > 0 Then .ListIndex = 0
    End With
End Sub

' -------------------------------------------------------------------------
' SIRUS_Parcourir_Fichier : FileDialog -> Txt_SIRUS_Fichier
' -------------------------------------------------------------------------
' Note de nommage 2026-06-10 (clarification) :
'   - SIRUS_GenererFeuilleAnalyse (G_SIRUS_Generation) = API du generateur
'     de feuille (cree SIRUS_<Table>_<Op>). Reste le point d'entree metier.
'   - SIRUS_UI_* (G_SIRUS_UserForm)                    = handlers de clic
'     UI qui lisent les controles UserForm et appellent l'API metier.
' Le prefixe SIRUS_UI_ evite la confusion entre les deux niveaux.
Public Sub SIRUS_UI_BtnParcourirClick()
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "SIRUS - Selectionner le livrable ATM ou EMMST"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Fichiers Excel", "*.xlsx; *.xlsm; *.xlsb"
        If .Show = -1 Then
            UserForm.Txt_SIRUS_Fichier.Value = .SelectedItems(1)
            UserForm.Lbl_SIRUS_NbFiches.Caption = "Fichier selectionne. Le decompte se fera a la generation."
        End If
    End With
End Sub

' -------------------------------------------------------------------------
' SIRUS_Generer_FeuilleAnalyse : dispatch Creer/Modifier/Extraction
' -------------------------------------------------------------------------
Public Sub SIRUS_UI_BtnGenererClick()
    Dim sOp As String, sTable As String

    sOp = SIRUS_GetOpActive()
    sTable = CStr(UserForm.Cmb_SIRUS_Table.Value)

    If LenB(sTable) = 0 Then
        MsgBox "Selectionner une table.", vbExclamation, "SIRUS"
        Exit Sub
    End If

    ' Reset du combo Feuille d'analyse (Zone 2)
    UserForm.Cmb_SIRUS_FeuilleAnalyse.Clear

    Select Case sOp
        Case "Creer"
            Call SIRUS_GenererFeuilleAnalyse(sTable, SIRUS_OP_CREER)
            Call SIRUS_AjouterFeuilleAuCombo("SIRUS_" & sTable & "_Creer")

        Case "Modifier"
            Dim cles As Variant
            Dim sProjet As String
            ' Source des cles :
            '   - ProjetASL Modifier : Project_List (combo principale en haut
            '     du UserForm). C'est le projet sur lequel l ingenieur ASL
            '     fait sa PBS dans ECOSLI -> on enrichit ensuite la feuille
            '     SIRUS avec les donnees de DATA_ASL.
            '   - ProjetSLI Modifier : Cmb_SIRUS_Projet (combo dans Frame_Modifier).
            '   - Autres tables : TextBox DTR multiligne.
            If sTable = "ProjetASL" Then
                sProjet = Trim$(CStr(UserForm.Project_List.Value))
                If LenB(sProjet) = 0 Then
                    MsgBox "Selectionner d abord un projet dans la liste principale " & _
                           "(Project_List) en haut du UserForm.", _
                           vbExclamation, "SIRUS"
                    Exit Sub
                End If
                cles = Array(sProjet)
            ElseIf sTable = "ProjetSLI" Then
                sProjet = CStr(UserForm.Cmb_SIRUS_Projet.Value)
                If LenB(sProjet) = 0 Then
                    MsgBox "Selectionner un projet dans la liste.", vbExclamation, "SIRUS"
                    Exit Sub
                End If
                cles = Array(sProjet)
            Else
                cles = SIRUS_ParserCles(UserForm.Txt_SIRUS_Cles.Value)
                If UBound(cles) < LBound(cles) Then
                    MsgBox "Saisir au moins une cle DTR.", vbExclamation, "SIRUS"
                    Exit Sub
                End If
            End If
            Call SIRUS_GenererFeuilleAnalyse(sTable, SIRUS_OP_MODIFIER, cles)
            ' Pour ProjetASL Modifier : croiser ensuite avec DATA_ASL (PBS
            ' du classeur projet actif). Overwrite des 6 champs PBS sur les
            ' lignes deja pre-remplies par BdD + ajout des DTR de la PBS
            ' qui n existent pas encore en BdD ProjetASL.
            If sTable = "ProjetASL" Then
                Call SIRUS_EnrichirProjetASL_DepuisDATA_ASL(SIRUS_OP_MODIFIER, sProjet)
            End If
            Call SIRUS_AjouterFeuilleAuCombo("SIRUS_" & sTable & "_Modifier")

        Case "Extraction"
            Dim sFichier As String, sLangue As String, sVersion As String
            sFichier = CStr(UserForm.Txt_SIRUS_Fichier.Value)
            If LenB(sFichier) = 0 Then
                MsgBox "Selectionner un fichier livrable.", vbExclamation, "SIRUS"
                Exit Sub
            End If

            ' Langue + Version depuis les radio buttons SIRUS
            If UserForm.Opt_SIRUS_FR.Value Then
                sLangue = "FR"
            Else
                sLangue = "EN"
            End If
            If UserForm.Opt_SIRUS_V1.Value Then
                sVersion = "V1"
            Else
                sVersion = "V2"
            End If

            ' 1. Creer les feuilles SIRUS_*_Extraction vierges (layout standard)
            Select Case sTable
                Case "ATM"
                    Call SIRUS_GenererFeuilleAnalyse("InfoATM", SIRUS_OP_EXTRACTION)
                    Call SIRUS_GenererFeuilleAnalyse("TachesATM", SIRUS_OP_EXTRACTION)
                    Call SIRUS_AjouterFeuilleAuCombo("SIRUS_InfoATM_Extraction")
                    Call SIRUS_AjouterFeuilleAuCombo("SIRUS_TachesATM_Extraction")
                Case "EMMST"
                    Call SIRUS_GenererFeuilleAnalyse("EMMST", SIRUS_OP_EXTRACTION)
                    Call SIRUS_AjouterFeuilleAuCombo("SIRUS_EMMST_Extraction")
            End Select

            ' 2. Deverser les donnees du fichier livrable dans les feuilles SIRUS
            '    (G_SIRUS_Extraction). Lance Main_Extraction_Params + mapping
            '    colonnes brutes -> SIRUS via la ligne 2 (ColonneBdD).
            Call SIRUS_DeverserExtractionDansSIRUS(sTable, sFichier, sLangue, sVersion)
    End Select

    ' Affichage progressif (assoupli 2026-06-11) : Frame Analyse ET Frame
    ' Envoi visibles ensemble apres Generation. Le seuil MeQ ne gate plus
    ' la visibilite de l envoi (decision Thomas : on permet d envoyer
    ' tout le temps pour l instant ; le blocage MeQ reviendra plus tard).
    On Error Resume Next
    UserForm.Frame_SIRUS_Analyse.Visible = True
    UserForm.Frame_SIRUS_Envoi.Visible = True
    UserForm.Btn_SIRUS_Envoyer.Enabled = True
    Call SIRUS_UI_PreRemplirEnvoi
    On Error GoTo 0
End Sub

' -------------------------------------------------------------------------
' SIRUS_RafraichirComboFeuilles : scanne ActiveWorkbook et alimente
' Cmb_SIRUS_FeuilleAnalyse avec TOUTES les feuilles SIRUS_<Table>_<Op>
' presentes (utile si l utilisateur a deja genere des feuilles dans une
' session precedente, ou si le UserForm vient d etre rouvert).
' Exclut les feuilles internes : SIRUS_Config, SIRUS_Diag_*, SIRUS_Historique.
' -------------------------------------------------------------------------
Public Sub SIRUS_RafraichirComboFeuilles()
    Dim ws As Worksheet
    Dim selActuelle As String
    Dim i As Long
    Dim bExiste As Boolean

    On Error Resume Next
    If ActiveWorkbook Is Nothing Then Exit Sub
    selActuelle = CStr(UserForm.Cmb_SIRUS_FeuilleAnalyse.Value)

    With UserForm.Cmb_SIRUS_FeuilleAnalyse
        .Clear
        For Each ws In ActiveWorkbook.Worksheets
            If Left$(ws.Name, 6) = "SIRUS_" Then
                ' Exclure feuilles internes (Config, Diag, Historique)
                If InStr(1, ws.Name, "SIRUS_Config", vbTextCompare) = 0 _
                   And InStr(1, ws.Name, "SIRUS_Diag", vbTextCompare) = 0 _
                   And InStr(1, ws.Name, "SIRUS_Historique", vbTextCompare) = 0 _
                   And InStr(1, ws.Name, "_Snap", vbTextCompare) = 0 Then
                    .AddItem ws.Name
                End If
            End If
        Next ws

        ' Restaurer la selection precedente si encore presente, sinon 1re
        bExiste = False
        For i = 0 To .ListCount - 1
            If .List(i) = selActuelle Then
                .ListIndex = i
                bExiste = True
                Exit For
            End If
        Next i
        If Not bExiste And .ListCount > 0 Then .ListIndex = 0
    End With
    On Error GoTo 0
End Sub

' -------------------------------------------------------------------------
' SIRUS_AjouterFeuilleAuCombo : ajoute un nom de feuille dans le combo
' Cmb_SIRUS_FeuilleAnalyse et selectionne la 1ere si c'est la 1re ajoutee.
' Pas de doublon : verifie l'existence avant ajout.
' -------------------------------------------------------------------------
Private Sub SIRUS_AjouterFeuilleAuCombo(ByVal sNomFeuille As String)
    Dim i As Long
    Dim bExiste As Boolean

    With UserForm.Cmb_SIRUS_FeuilleAnalyse
        ' Eviter les doublons (cas relance generation)
        bExiste = False
        For i = 0 To .ListCount - 1
            If .List(i) = sNomFeuille Then
                bExiste = True
                Exit For
            End If
        Next i
        If Not bExiste Then .AddItem sNomFeuille
        If .ListIndex < 0 And .ListCount > 0 Then .ListIndex = 0
    End With
End Sub

' -------------------------------------------------------------------------
' SIRUS_Analyser_Feuille : appelle SIRUS_AnalyserFeuille + remplit ListView
' -------------------------------------------------------------------------
Public Sub SIRUS_UI_BtnAnalyserClick()
    Dim sFeuille As String
    Dim sTable As String
    Dim sOp As String
    Dim taux As Double

    sFeuille = CStr(UserForm.Cmb_SIRUS_FeuilleAnalyse.Value)
    If LenB(sFeuille) = 0 Then
        MsgBox "Aucune feuille a analyser. Generer d'abord une feuille.", _
               vbExclamation, "SIRUS"
        Exit Sub
    End If

    ' Parser "SIRUS_<Table>_<Op>" -> sTable, sOp
    Call SIRUS_ParserNomFeuille(sFeuille, sTable, sOp)

    taux = SIRUS_AnalyserFeuille(sTable, sOp)

    ' Resume statut + activer envoi (gating MeQ retire pour l instant -
    ' decision Thomas 2026-06-11 : on permet d envoyer tout le temps).
    UserForm.Lbl_SIRUS_StatutAnalyse.Caption = _
        "Taux de completude (MeQ) : " & Format(taux * 100, "0.0") & "%"

    UserForm.Btn_SIRUS_Envoyer.Enabled = True

    ' Remplir la ListView des NOK
    Call SIRUS_RemplirListViewResultats(sFeuille)

    ' Frame Envoi visible (deja affichee depuis Generation), MAJ recap envoi.
    On Error Resume Next
    UserForm.Frame_SIRUS_Envoi.Visible = True
    Call SIRUS_UI_PreRemplirEnvoi
    On Error GoTo 0
End Sub

' -------------------------------------------------------------------------
' SIRUS_UI_BtnEditerFeuilleClick : porte de sortie vers la feuille
' -------------------------------------------------------------------------
' Active la feuille SIRUS selectionnee dans Cmb_SIRUS_FeuilleAnalyse et
' masque le UserForm pour permettre la saisie libre dans Excel.
' Le user peut revenir au UserForm via le raccourci Ctrl+W (W_AppelUsf).
' Solution pragmatique si le UserForm reste bloquant malgre vbModeless
' (cas ou ShowModal=True dans le designer).
Public Sub SIRUS_UI_BtnEditerFeuilleClick()
    Dim sFeuille As String
    Dim ws As Worksheet

    sFeuille = CStr(UserForm.Cmb_SIRUS_FeuilleAnalyse.Value)
    If LenB(sFeuille) = 0 Then
        MsgBox "Selectionner d'abord une feuille dans la liste Zone 2.", _
               vbExclamation, "SIRUS"
        Exit Sub
    End If

    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets(sFeuille)
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "Feuille '" & sFeuille & "' introuvable dans le classeur actif.", _
               vbExclamation, "SIRUS"
        Exit Sub
    End If

    ' Activer la feuille + placer le curseur sur la 1re cellule de saisie (col 2)
    ws.Activate
    ws.Cells(3, 2).Select

    ' Masquer le UserForm (le user ramene via Ctrl+W)
    UserForm.Hide
End Sub

' -------------------------------------------------------------------------
' SIRUS_GotoLigneFautive : double-clic ListView -> Activer la cellule cle
' -------------------------------------------------------------------------
Public Sub SIRUS_UI_LvResultatsDblClick()
    ' Double-clic sur une alerte = goto sur la cellule fautive (ligne+colChamp
    ' parses depuis la key "alerte_<idx>_<ligne>_<colChamp>" stockee par
    ' SIRUS_RemplirListViewResultats).
    Dim lv As MSComctlLib.ListView
    Dim li As MSComctlLib.ListItem
    Dim ligne As Long, colChamp As Long
    Dim sFeuille As String
    Dim ws As Worksheet
    Dim parts() As String

    Set lv = UserForm.ListView_SIRUS_Resultats
    Set li = lv.SelectedItem
    If li Is Nothing Then Exit Sub

    parts = Split(li.key, "_")
    If UBound(parts) < 3 Then Exit Sub
    ligne = CLng(parts(2))
    colChamp = CLng(parts(3))
    If colChamp < 1 Then colChamp = 3   ' fallback : cellule cle

    sFeuille = CStr(UserForm.Cmb_SIRUS_FeuilleAnalyse.Value)
    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets(sFeuille)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Application.Goto ws.Cells(ligne, colChamp), Scroll:=True
End Sub

' -------------------------------------------------------------------------
' SIRUS_Envoyer_Demande : Phase 5, TODO
' -------------------------------------------------------------------------
Public Sub SIRUS_UI_BtnEnvoyerClick()
    Dim sProjet As String, sSujet As String, sDesc As String
    Dim sAttentes As String, sPriorite As String
    Dim sNumeroImport As String, sNumeroVBN As String
    Dim digest As String
    Dim userID As Long, ticketID As Long
    Dim cheminPJ As String
    Dim resultTicket As String, resultPJ As String
    Dim reponse As VbMsgBoxResult
    Dim feuilles() As String

    cheminPJ = ""
    On Error GoTo Sortie

    ' -- 1. Validation des inputs ----------------------------
    sProjet = CStr(UserForm.Project_List.Value)
    sSujet = Trim$(CStr(UserForm.Txt_SIRUS_Sujet.Value))
    sDesc = Trim$(CStr(UserForm.Txt_SIRUS_Description.Value))
    sAttentes = Trim$(CStr(UserForm.Txt_SIRUS_Attentes.Value))
    sPriorite = CStr(UserForm.Cmb_SIRUS_Priorite.Value)

    If LenB(sProjet) = 0 Then
        MsgBox "Selectionner un projet dans la liste principale Project_List " & _
               "avant l envoi.", vbExclamation, "SIRUS"
        Exit Sub
    End If
    If LenB(sSujet) = 0 Then
        MsgBox "Le sujet est obligatoire.", vbExclamation, "SIRUS"
        Exit Sub
    End If
    If LenB(sDesc) = 0 Then
        MsgBox "La description est obligatoire.", vbExclamation, "SIRUS"
        Exit Sub
    End If
    If LenB(sPriorite) = 0 Then
        MsgBox "Selectionner une priorite (High/Medium/Low).", vbExclamation, "SIRUS"
        Exit Sub
    End If

    feuilles = SIRUS_RecupererFeuillesAEnvoyer()
    If UBound(feuilles) < LBound(feuilles) Then
        MsgBox "Aucune feuille SIRUS a envoyer. Generer + analyser une feuille d abord.", _
               vbExclamation, "SIRUS"
        Exit Sub
    End If

    ' -- 1 bis. Garde-fou ProjetASL : aucun projet inconnu en ProjetSLI ----
    '    Critere metier (Thomas 2026-06-11) : impossible de creer/modifier
    '    un ProjetASL pour un code projet qui n existe pas dans ProjetSLI.
    Dim sProjetsInconnus As String
    sProjetsInconnus = SIRUS_ProjetsASL_Inconnus(feuilles)
    If LenB(sProjetsInconnus) > 0 Then
        MsgBox "Envoi bloque : projet(s) inconnu(s) dans ProjetSLI :" & vbCrLf & vbCrLf & _
               sProjetsInconnus & vbCrLf & vbCrLf & _
               "Creer d abord le projet en ProjetSLI avant d ajouter des entrees ProjetASL.", _
               vbCritical, "SIRUS - Envoi impossible"
        Exit Sub
    End If

    ' -- 2. Confirmation utilisateur -------------------------
    reponse = MsgBox("Confirmer l envoi du ticket ATMAP ?" & vbCrLf & vbCrLf & _
                     "Sujet     : " & sSujet & vbCrLf & _
                     "Priorite  : " & sPriorite & vbCrLf & _
                     "Projet    : " & sProjet & vbCrLf & _
                     "Feuilles  : " & (UBound(feuilles) - LBound(feuilles) + 1) & " a envoyer", _
                     vbYesNo + vbQuestion + vbDefaultButton2, _
                     "SIRUS - Confirmation envoi ATMAP")
    If reponse <> vbYes Then Exit Sub

    ' -- 3. Construire la PJ Excel (toutes feuilles SIRUS du combo) ----
    UserForm.Lbl_SIRUS_StatutEnvoi.Caption = "Construction de la PJ Excel..."
    DoEvents
    cheminPJ = SIRUS_ConstruirePJ(feuilles, "SIRUS_PJ")
    If LenB(cheminPJ) = 0 Then
        UserForm.Lbl_SIRUS_StatutEnvoi.Caption = "Echec construction PJ"
        UserForm.Lbl_SIRUS_StatutEnvoi.ForeColor = RGB(163, 45, 45)
        Exit Sub
    End If

    ' -- 4. Recuperer digest SharePoint ----------------------
    UserForm.Lbl_SIRUS_StatutEnvoi.Caption = "Authentification SharePoint..."
    DoEvents
    digest = SIRUS_GetFormDigest()
    If LenB(digest) = 0 Then
        UserForm.Lbl_SIRUS_StatutEnvoi.Caption = "Echec digest SharePoint"
        UserForm.Lbl_SIRUS_StatutEnvoi.ForeColor = RGB(163, 45, 45)
        On Error Resume Next
        Kill cheminPJ
        On Error GoTo 0
        Exit Sub
    End If

    userID = SIRUS_GetCurrentUserID(digest)
    If userID = -1 Then
        UserForm.Lbl_SIRUS_StatutEnvoi.Caption = "Echec recuperation utilisateur"
        UserForm.Lbl_SIRUS_StatutEnvoi.ForeColor = RGB(163, 45, 45)
        On Error Resume Next
        Kill cheminPJ
        On Error GoTo 0
        Exit Sub
    End If

    ' -- 5. Extraire le numero d'import du sujet (format "IMP-YYYY-00X - ...")
    '       et generer le Numero_ticket VBN_NN cote SP (auto-increment).
    Dim posSep As Long
    posSep = InStr(1, sSujet, " - ")
    If posSep > 0 Then
        sNumeroImport = Trim$(Left$(sSujet, posSep - 1))
    Else
        sNumeroImport = sSujet
    End If

    sNumeroVBN = SIRUS_NumeroVBNSuivant()
    If LenB(sNumeroVBN) = 0 Then
        UserForm.Lbl_SIRUS_StatutEnvoi.Caption = "Echec generation Numero_ticket VBN"
        UserForm.Lbl_SIRUS_StatutEnvoi.ForeColor = RGB(163, 45, 45)
        On Error Resume Next
        Kill cheminPJ
        On Error GoTo 0
        Exit Sub
    End If

    ' -- 6. Creer le ticket ATMAP ----------------------------
    '       Numero_ticket   = VBN_NN (auto-incremente cote SP)
    '       Numero_Import   = IMP-YYYY-NNN (extrait du sujet)
    '       Date_estimee    = aujourd hui + N jours selon priorite
    Dim sDateEstimee As String
    sDateEstimee = SIRUS_DateEstimeeISO(sPriorite)

    UserForm.Lbl_SIRUS_StatutEnvoi.Caption = "Creation du ticket..."
    DoEvents
    resultTicket = SIRUS_CreerTicketEnvoi(digest, userID, sNumeroVBN, _
                                            sSujet, sDesc, sAttentes, sPriorite, ticketID, _
                                            sNumeroImport, sDateEstimee)

    If ticketID = -1 Then
        UserForm.Lbl_SIRUS_StatutEnvoi.Caption = "Echec creation ticket"
        UserForm.Lbl_SIRUS_StatutEnvoi.ForeColor = RGB(163, 45, 45)
        MsgBox "Echec creation ticket :" & vbCrLf & resultTicket, vbCritical, "SIRUS"
        On Error Resume Next
        Kill cheminPJ
        On Error GoTo 0
        Exit Sub
    End If

    ' -- 7. Attacher la PJ -----------------------------------
    UserForm.Lbl_SIRUS_StatutEnvoi.Caption = "Upload de la PJ..."
    DoEvents
    resultPJ = SIRUS_AttacherFichierAuTicket(ticketID, digest, cheminPJ)

    ' -- 8. Cleanup PJ temp ----------------------------------
    On Error Resume Next
    Kill cheminPJ
    On Error GoTo 0

    ' -- 9. Statut final visible -----------------------------
    UserForm.Lbl_SIRUS_StatutEnvoi.Caption = "OK - " & sNumeroVBN & _
                                              " / " & sNumeroImport & _
                                              " (SP Id " & ticketID & ")"
    UserForm.Lbl_SIRUS_StatutEnvoi.ForeColor = RGB(39, 80, 10)

    MsgBox "Ticket cree avec succes !" & vbCrLf & vbCrLf & _
           "Numero ticket  : " & sNumeroVBN & vbCrLf & _
           "Numero import  : " & sNumeroImport & vbCrLf & _
           "SharePoint ID  : " & ticketID & vbCrLf & vbCrLf & _
           resultPJ, vbInformation, "SIRUS - Envoi termine"
    Exit Sub

Sortie:
    UserForm.Lbl_SIRUS_StatutEnvoi.Caption = "Erreur : " & Err.Description
    UserForm.Lbl_SIRUS_StatutEnvoi.ForeColor = RGB(163, 45, 45)
    On Error Resume Next
    If LenB(cheminPJ) > 0 Then Kill cheminPJ
    On Error GoTo 0
End Sub

' =========================================================================
' SIRUS_UI_PrioriteChange : appele par Cmb_SIRUS_Priorite_Change
' =========================================================================
' Met a jour le pre-remplissage des Attentes selon la priorite choisie.
' L utilisateur peut ensuite editer librement.
Public Sub SIRUS_UI_PrioriteChange()
    Dim sPriorite As String
    Dim nbJours As Long
    Dim sLabel As String
    On Error Resume Next
    sPriorite = CStr(UserForm.Cmb_SIRUS_Priorite.Value)
    nbJours = SIRUS_JoursTraitementSelonPriorite(sPriorite)
    If nbJours <= 1 Then
        sLabel = "Traitement sous " & nbJours & " jour"
    Else
        sLabel = "Traitement sous " & nbJours & " jours"
    End If
    UserForm.Txt_SIRUS_Attentes.Value = sLabel
    On Error GoTo 0
End Sub

' =========================================================================
' SIRUS_UI_ProjetListChange
' =========================================================================
' Handler du Change event de Project_List (combobox principale du UserForm).
' Met a jour le texte de Description si la Zone Envoi est visible (= apres
' analyse OK). Permet a l ingenieur ASL de voir immediatement le projet
' dans la description sans re-cliquer sur Analyser.
Public Sub SIRUS_UI_ProjetListChange()
    Dim sOp As String, sTable As String

    On Error Resume Next
    ' Si Zone Envoi pas encore visible, rien a faire
    If Not UserForm.Frame_SIRUS_Envoi.Visible Then Exit Sub

    ' Pour ProjetSLI / ProjetASL, le projet de la description NE vient PAS
    ' de Project_List (liste du haut) mais du projet de la feuille SIRUS
    ' en cours (cf. SIRUS_ProjetPourDescription). Donc on ne reagit au
    ' changement de Project_List que pour les autres tables.
    sTable = CStr(UserForm.Cmb_SIRUS_Table.Value)
    If sTable = "ProjetSLI" Or sTable = "ProjetASL" Then Exit Sub

    ' Re-construire la description (logique unifiee dans un helper)
    Call SIRUS_MajDescriptionEnvoi
    On Error GoTo 0
End Sub

' =========================================================================
' SIRUS_ProjetPourDescription
' =========================================================================
' Renvoie le nom de projet a utiliser dans la description ATMAP :
'   - Produit/InfoATM/TachesATM/EMMST : Project_List (liste du haut)
'   - ProjetSLI / ProjetASL : col 2 ligne 3 de la feuille active
'     (= 1ere valeur "Nom projet" saisie dans la feuille). Fallback :
'     Project_List si feuille vide.
Private Function SIRUS_ProjetPourDescription() As String
    Dim sTable As String, sFeuille As String, sProjet As String
    Dim ws As Worksheet

    sTable = CStr(UserForm.Cmb_SIRUS_Table.Value)
    sFeuille = CStr(UserForm.Cmb_SIRUS_FeuilleAnalyse.Value)

    If sTable = "ProjetSLI" Or sTable = "ProjetASL" Then
        ' Lire la valeur projet dans la feuille SIRUS
        On Error Resume Next
        Set ws = ActiveWorkbook.Worksheets(sFeuille)
        On Error GoTo 0
        If Not ws Is Nothing Then
            sProjet = Trim$(SafeCStr(ws.Cells(3, 2).Value))
        End If
        If LenB(sProjet) = 0 Then
            ' Fallback Project_List (cas Creer avec feuille vide)
            sProjet = CStr(UserForm.Project_List.Value)
        End If
    Else
        sProjet = CStr(UserForm.Project_List.Value)
    End If

    SIRUS_ProjetPourDescription = sProjet
End Function

' Helper unifie : reconstruit la description avec la bonne logique projet
Private Sub SIRUS_MajDescriptionEnvoi()
    Dim sOp As String, sTable As String, sProjet As String
    Dim sFeuilleActive As String
    Dim nbLignes As Long
    Dim sProjetAff As String

    sOp = SIRUS_GetOpActive()
    sTable = CStr(UserForm.Cmb_SIRUS_Table.Value)
    sProjet = SIRUS_ProjetPourDescription()
    sFeuilleActive = CStr(UserForm.Cmb_SIRUS_FeuilleAnalyse.Value)
    nbLignes = SIRUS_CompterLignesFeuille(sFeuilleActive)

    If LenB(sProjet) > 0 Then
        sProjetAff = sProjet
    Else
        sProjetAff = "<projet non selectionne>"
    End If

    On Error Resume Next
    UserForm.Txt_SIRUS_Description.Value = sOp & " de " & nbLignes & _
        " reference(s) en " & sTable & " pour le projet " & sProjetAff
    On Error GoTo 0
End Sub

' =========================================================================
' SIRUS_UI_PreRemplirEnvoi
' =========================================================================
' Pre-remplit la Zone 3 apres une analyse reussie. A appeler depuis
' SIRUS_UI_BtnAnalyserClick a la fin (cf. SIRUS_RemplirListViewResultats).
' Auto-rempli :
'   - Demandeur (Application.UserName), Date (today)
'   - Op + Table recopies des controles Zone 1
'   - Recap feuilles
'   - Sujet : IMP-YYYY-00X (via SIRUS_NumeroTicketSuivant) + Op + Table
'   - Description : "<Op> de <N> reference(s) en <Table> pour le projet <Projet>"
'   - Priorite : Medium par defaut
'   - Attentes : selon priorite via SIRUS_UI_PrioriteChange
Public Sub SIRUS_UI_PreRemplirEnvoi()
    Dim sOp As String, sTable As String, sProjet As String
    Dim sFeuilleActive As String
    Dim nbLignes As Long
    Dim numTicket As String

    On Error Resume Next

    sOp = SIRUS_GetOpActive()
    sTable = CStr(UserForm.Cmb_SIRUS_Table.Value)
    sFeuilleActive = CStr(UserForm.Cmb_SIRUS_FeuilleAnalyse.Value)
    sProjet = CStr(UserForm.Project_List.Value)

    ' Recap auto-rempli
    UserForm.Lbl_SIRUS_Demandeur.Caption = Application.UserName
    UserForm.Lbl_SIRUS_Date.Caption = Format$(Now, "dd/mm/yyyy")
    UserForm.Lbl_SIRUS_OpRecap.Caption = sOp
    UserForm.Lbl_SIRUS_TableRecap.Caption = sTable
    UserForm.Lbl_SIRUS_FeuillesRecap.Caption = SIRUS_ResumerFeuillesPourEnvoi()

    ' Sujet auto (numero ticket suivant via SharePoint)
    numTicket = SIRUS_NumeroTicketSuivant()
    UserForm.Txt_SIRUS_Sujet.Value = numTicket & " - Import " & sOp & " " & sTable

    ' Description auto via helper (logique projet selon table)
    Call SIRUS_MajDescriptionEnvoi

    ' Priorite defaut Medium + pre-remplir Attentes
    If LenB(CStr(UserForm.Cmb_SIRUS_Priorite.Value)) = 0 Then
        UserForm.Cmb_SIRUS_Priorite.Value = "Medium"
    End If
    Call SIRUS_UI_PrioriteChange

    On Error GoTo 0
End Sub

' =========================================================================
' Helpers UI Phase 5
' =========================================================================

' Recupere les noms des feuilles SIRUS du combo Cmb_SIRUS_FeuilleAnalyse
' (= toutes les feuilles generees pour la session courante).
Private Function SIRUS_RecupererFeuillesAEnvoyer() As String()
    Dim i As Long, n As Long
    Dim resultat() As String

    On Error Resume Next
    n = UserForm.Cmb_SIRUS_FeuilleAnalyse.ListCount
    On Error GoTo 0

    If n = 0 Then
        ReDim resultat(0 To -1)
        SIRUS_RecupererFeuillesAEnvoyer = resultat
        Exit Function
    End If

    ReDim resultat(0 To n - 1)
    For i = 0 To n - 1
        resultat(i) = CStr(UserForm.Cmb_SIRUS_FeuilleAnalyse.List(i))
    Next i
    SIRUS_RecupererFeuillesAEnvoyer = resultat
End Function

' Resume textuel des feuilles a envoyer (pour Lbl_SIRUS_FeuillesRecap).
Private Function SIRUS_ResumerFeuillesPourEnvoi() As String
    Dim feuilles() As String
    Dim i As Long, s As String
    feuilles = SIRUS_RecupererFeuillesAEnvoyer()
    For i = LBound(feuilles) To UBound(feuilles)
        If LenB(s) > 0 Then s = s & ", "
        s = s & feuilles(i)
    Next i
    SIRUS_ResumerFeuillesPourEnvoi = s
End Function

' Scan les feuilles SIRUS_ProjetASL_* a envoyer, collecte les noms de projets
' uniques en col 2 et renvoie ceux absents de ProjetSLI (BdD TRAVAIL).
' Renvoie une chaine vide si tout est connu, ou un texte multi-lignes
' avec un projet inconnu par ligne sinon. Helper du garde-fou envoi.
Private Function SIRUS_ProjetsASL_Inconnus(ByRef feuilles() As String) As String
    Dim i As Long
    Dim sNom As String, sTable As String, sOp As String
    Dim ws As Worksheet
    Dim ligneMax As Long, ligne As Long
    Dim sProjet As String
    Dim dictProjSLI As Object, dictInconnus As Object
    Dim sOut As String, k As Variant

    For i = LBound(feuilles) To UBound(feuilles)
        sNom = feuilles(i)
        Call SIRUS_ParserNomFeuille(sNom, sTable, sOp)
        If sTable = "ProjetASL" Then
            On Error Resume Next
            Set ws = ActiveWorkbook.Worksheets(sNom)
            On Error GoTo 0
            If ws Is Nothing Then GoTo NextFeuille

            ' Charger le cache ProjetSLI une fois (lazy load).
            If dictProjSLI Is Nothing Then Set dictProjSLI = SIRUS_ProjetsSLI()
            If dictProjSLI Is Nothing Then
                ' Echec connexion BdD : on ne bloque pas (l alerte radar
                ' apparaitra deja cote feuille). Cas degrade tolere.
                Exit Function
            End If
            If dictInconnus Is Nothing Then
                Set dictInconnus = CreateObject("Scripting.Dictionary")
                dictInconnus.CompareMode = 1
            End If

            ligneMax = ws.Cells(ws.rows.count, 2).End(xlUp).Row
            For ligne = 3 To ligneMax
                sProjet = Trim$(SafeCStr(ws.Cells(ligne, 2).Value))
                If LenB(sProjet) > 0 Then
                    If Not dictProjSLI.Exists(UCase$(sProjet)) Then
                        If Not dictInconnus.Exists(UCase$(sProjet)) Then
                            dictInconnus.Add UCase$(sProjet), sProjet
                        End If
                    End If
                End If
            Next ligne
        End If
NextFeuille:
        Set ws = Nothing
    Next i

    If dictInconnus Is Nothing Then
        SIRUS_ProjetsASL_Inconnus = ""
        Exit Function
    End If

    For Each k In dictInconnus.Keys
        If LenB(sOut) > 0 Then sOut = sOut & vbLf
        sOut = sOut & "- " & dictInconnus(k)
    Next k
    SIRUS_ProjetsASL_Inconnus = sOut
End Function

' Compte les lignes de donnees (apres entete) d une feuille SIRUS.
' Scanne la colonne cle metier (table-aware via SIRUS_ColonneClePourTable).
Private Function SIRUS_CompterLignesFeuille(ByVal sNomFeuille As String) As Long
    Dim ws As Worksheet
    Dim n As Long
    Dim sTable As String, sOp As String
    Dim colCle As Long

    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets(sNomFeuille)
    On Error GoTo 0
    If ws Is Nothing Then
        SIRUS_CompterLignesFeuille = 0
        Exit Function
    End If

    Call SIRUS_ParserNomFeuille(sNomFeuille, sTable, sOp)
    colCle = SIRUS_ColonneClePourTable(sTable)

    n = ws.Cells(ws.rows.count, colCle).End(xlUp).Row - 2
    If n < 0 Then n = 0
    SIRUS_CompterLignesFeuille = n
End Function

' =========================================================================
' Helpers UI SIRUS v2
' =========================================================================

Private Function SIRUS_GetOpActive() As String
    If UserForm.Opt_SIRUS_Creer.Value Then
        SIRUS_GetOpActive = "Creer"
    ElseIf UserForm.Opt_SIRUS_Modifier.Value Then
        SIRUS_GetOpActive = "Modifier"
    ElseIf UserForm.Opt_SIRUS_Extraction.Value Then
        SIRUS_GetOpActive = "Extraction"
    Else
        SIRUS_GetOpActive = ""
    End If
End Function

Private Function SIRUS_ParserCles(ByVal sTexte As String) As Variant
    ' Split sur newline + filtre lignes vides + trim. Renvoie tableau 0-based.
    Dim arr As Variant
    Dim i As Long, k As Long
    Dim s As String
    Dim resultat() As String

    sTexte = Replace(sTexte, vbCrLf, vbLf)
    sTexte = Replace(sTexte, vbCr, vbLf)
    arr = Split(sTexte, vbLf)

    ReDim resultat(0 To UBound(arr))
    k = 0
    For i = LBound(arr) To UBound(arr)
        s = Trim$(CStr(arr(i)))
        If LenB(s) > 0 Then
            resultat(k) = s
            k = k + 1
        End If
    Next i

    If k = 0 Then
        SIRUS_ParserCles = Array()
    Else
        ReDim Preserve resultat(0 To k - 1)
        SIRUS_ParserCles = resultat
    End If
End Function

Private Sub SIRUS_ParserNomFeuille(ByVal sNom As String, _
                                    ByRef sTable As String, _
                                    ByRef sOp As String)
    ' "SIRUS_Produit_Creer" -> sTable="Produit", sOp="Creer"
    Dim parts() As String
    parts = Split(sNom, "_")
    If UBound(parts) >= 2 Then
        sTable = parts(1)
        sOp = parts(2)
    End If
End Sub

Public Sub SIRUS_ConfigListViewResultats()
    Dim lv As MSComctlLib.ListView
    On Error Resume Next
    Set lv = UserForm.ListView_SIRUS_Resultats
    If lv Is Nothing Then Exit Sub
    On Error GoTo 0

    ' Nouveau layout (refonte 2026-06-11) : 1 ligne par alerte.
    '   Cle    : DTR (toutes tables) ou Projet (ProjetSLI) — col 3 feuille
    '   Champ  : libelle humain du champ fautif (def.Champ)
    '   Alertes: detail de l alerte sur ce champ precis
    With lv
        .View = lvwReport
        .FullRowSelect = True
        .Gridlines = True
        .LabelEdit = lvwManual
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "Cle", 120, lvwColumnLeft
        .ColumnHeaders.Add , , "Champ", 180, lvwColumnLeft
        .ColumnHeaders.Add , , "Alertes", 300, lvwColumnLeft
        .ListItems.Clear
    End With
End Sub

Private Sub SIRUS_RemplirListViewResultats(ByVal sFeuille As String)
    ' Refonte 2026-06-11 : une ligne par alerte (au lieu d'une ligne par DTR
    ' avec alertes concatenees). Lit la collection structuree sirusAlertes
    ' alimentee par SIRUS_AnalyserFeuille (G_SIRUS_Analyse).
    ' Chaque entree sirusAlertes = Array(cle, champ, alerte, ligne, colChamp).
    Dim lv As MSComctlLib.ListView
    Dim li As MSComctlLib.ListItem
    Dim it As Variant
    Dim sCle As String, sChamp As String, sAlerte As String
    Dim nLigne As Long, nColChamp As Long
    Dim key As String

    ' Re-configurer la ListView (idempotent : recree les ColumnHeaders).
    ' Critique : sans ColumnHeaders, SubItems(1) leve "Invalid property value".
    Call SIRUS_ConfigListViewResultats

    Set lv = UserForm.ListView_SIRUS_Resultats
    lv.ListItems.Clear

    If sirusAlertes Is Nothing Then Exit Sub
    If sirusAlertes.count = 0 Then Exit Sub

    Dim idx As Long: idx = 0
    For Each it In sirusAlertes
        sCle = SafeCStr(it(0))
        sChamp = SafeCStr(it(1))
        sAlerte = SafeCStr(it(2))
        nLigne = CLng(it(3))
        nColChamp = CLng(it(4))
        idx = idx + 1

        ' Key technique : "ligne|colChamp" pour le goto en double-clic
        key = "alerte_" & idx & "_" & nLigne & "_" & nColChamp

        On Error Resume Next
        Set li = lv.ListItems.Add(, key, sCle)
        li.SubItems(1) = sChamp
        li.SubItems(2) = sAlerte
        ' Coloration rouge gras des alertes
        li.ForeColor = RGB(163, 45, 45)
        li.Bold = True
        li.ListSubItems(1).ForeColor = RGB(163, 45, 45)
        li.ListSubItems(1).Bold = True
        li.ListSubItems(2).ForeColor = RGB(163, 45, 45)
        On Error GoTo 0
    Next it
End Sub

' SafeCStr : conversion en chaine resistante a Null / Error / Empty.
' Evite les "Invalid use of Null" et "Invalid property value" quand on lit
' des cellules contenant des valeurs speciales (#N/A, formules vides, etc.).
Private Function SafeCStr(ByVal v As Variant) As String
    On Error Resume Next
    If IsNull(v) Then
        SafeCStr = ""
    ElseIf IsEmpty(v) Then
        SafeCStr = ""
    ElseIf IsError(v) Then
        SafeCStr = ""
    Else
        SafeCStr = CStr(v)
    End If
    If Err.Number <> 0 Then
        SafeCStr = ""
        Err.Clear
    End If
    On Error GoTo 0
End Function
