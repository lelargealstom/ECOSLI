Attribute VB_Name = "G_SIRUS_UserForm"

' -- Constante : colonnes de la ListBox Produit -------------
' Col 0 = DTR | Col 1 = D�signation | Col 2 = Node Type
' Col 3 = Fournisseur | Col 4 = R�f. Fourn. | Col 5 = Drawing Nb

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
' CONFIGURATION LISTVIEW_PRODUIT � dans UserForm_Initialize
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

    ' -- Cr�er les colonnes avec largeurs ------------------
    lv.ColumnHeaders.Add , , "Type", 60
    lv.ColumnHeaders.Add , , "DTR", 80
    lv.ColumnHeaders.Add , , "D�signation", 150
    lv.ColumnHeaders.Add , , "Node Type", 80
    lv.ColumnHeaders.Add , , "Fournisseur", 100
    lv.ColumnHeaders.Add , , "R�f. Fournisseur", 100
    lv.ColumnHeaders.Add , , "Drawing Number", 100
    lv.ColumnHeaders.Add , , "MTBF Intr.", 75
    lv.ColumnHeaders.Add , , "DVL HARD", 70
    lv.ColumnHeaders.Add , , "DVL SOFT", 70

End Sub
' ============================================================
' CONFIGURATION LISTVIEW_PRODUIT � dans UserForm_Initialize
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

    ' -- Cr�er les colonnes avec largeurs ------------------
    lv.ColumnHeaders.Add , , "DTR", 80
    lv.ColumnHeaders.Add , , "Projet_ASL", 70
    lv.ColumnHeaders.Add , , "ATM - Info", 70
    lv.ColumnHeaders.Add , , "ATM - Taches", 80
    lv.ColumnHeaders.Add , , "EMMST", 60
    lv.ColumnHeaders.Add , , "Produit", 60
    lv.ColumnHeaders.Add , , "Qualite Prix Spare", 130
    lv.ColumnHeaders.Add , , "Source Ecart Spare", 100
    lv.ColumnHeaders.Add , , "Source Manque Spare", 130
    lv.ColumnHeaders.Add , , "Qualite Prix Repair", 130
    lv.ColumnHeaders.Add , , "Source Ecart Repair", 100
    lv.ColumnHeaders.Add , , "Source Manque Repair", 130

End Sub
' ============================================================
' CONFIGURATION LISTVIEW_TACHESATM � dans UserForm_Initialize
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

    ' -- Cr�er les colonnes avec largeurs ------------------
    lv.ColumnHeaders.Add , , "DTR", 80
    lv.ColumnHeaders.Add , , "Langue", 40
    lv.ColumnHeaders.Add , , "Rang", 40
    lv.ColumnHeaders.Add , , "Type", 40
    lv.ColumnHeaders.Add , , "Cat�gorie", 60
    lv.ColumnHeaders.Add , , "Niv. Qualif", 60
    lv.ColumnHeaders.Add , , "Description", 100
    lv.ColumnHeaders.Add , , "Crit�res Dim", 60
    lv.ColumnHeaders.Add , , "Out. Sp�", 60
    lv.ColumnHeaders.Add , , "Out. Standard", 60
    lv.ColumnHeaders.Add , , "Pieces Rech", 60
    lv.ColumnHeaders.Add , , "Nb Op�rateurs", 40
    lv.ColumnHeaders.Add , , "Tps estim� (h)", 80
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
    lv.ColumnHeaders.Add , , "Contr�le", 110
    lv.ColumnHeaders.Add , , "Action corrective", 300

End Sub
' ============================================================
' CONFIGURATION LISTVIEW_ATM � dans UserForm_Initialize
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

    ' -- Cr�er les colonnes avec largeurs ------------------
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
    lv.ColumnHeaders.Add , , "Nb Op�rateurs", 40
    lv.ColumnHeaders.Add , , "Tps estim� (h)", 40
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

    ' Cr�er la ligne avec la valeur de la premi�re colonne (DTR)
    Set item = lv.ListItems.Add(, , TypeMod)

    ' Ajouter les sous-�l�ments (colonnes 2 � 10)
    item.SubItems(1) = vals(1)  ' D�signation
    item.SubItems(2) = vals(2)  ' D�signation
    item.SubItems(3) = vals(3)  ' Node Type
    item.SubItems(4) = vals(4)  ' Fournisseur
    item.SubItems(5) = vals(5)  ' R�f. Fournisseur
    item.SubItems(6) = vals(6)  ' Drawing Number
    item.SubItems(7) = vals(7)  ' MTBF Intr.
    item.SubItems(8) = vals(8)  ' DVL HARD
    item.SubItems(9) = vals(9)  ' DVL SOFT

End Sub

' ============================================================
' CHANGEMENT TYPE DEMANDE (NavTab Cr�ation/Modif/Suppression)
' ============================================================
Public Sub SIRUS_SetType(typeVal As String)

    With UserForm
        Select Case typeVal

            Case "Cr�ation"
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
                ' Tous les champs actifs � pr�-remplis depuis BdD apr�s saisie DTR
                .Designation_txt.Enabled = True
                .NodeType_list.Enabled = True
                .FRS_list.Enabled = True
                .RefFRS_txt.Enabled = True
                .DrawingNb_txt.Enabled = True
                .MTBF_Intr_txt.Enabled = True
                .DVL_HARD_txt.Enabled = True
                .DVL_SOFT_txt.Enabled = True

            Case "Suppression"
                ' DTR seul requis, tout le reste gris�
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
        If des = "" Then missing = missing & "D�signation / "
        If node = "" Then missing = missing & "Node Type / "
        If fourn = "" Then missing = missing & "Fournisseur / "
        If ref = "" Then missing = missing & "R�f. Fournisseur / "
        If drw = "" Then missing = missing & "Drawing Number / "
        If MTBF = "" Then missing = missing & "MTBF Intrins�que / "
        If dvlh = "" Then missing = missing & "DVL HARD / "
    End If

    If missing <> "" Then
        MsgBox "Champs obligatoires manquants : " & _
               Left(missing, Len(missing) - 3), vbExclamation, "SIRUS"
        Exit Sub
    End If

    ' -- V�rifier doublon DTR dans la collection ------------
    Dim item As Variant
    For Each item In lignesProduit
        If IsArray(item) Then
            If UCase(Trim(item(0))) = UCase(dtr) Then
                MsgBox "DTR '" & dtr & "' d�j� pr�sent dans la liste.", _
                       vbExclamation, "SIRUS"
                Exit Sub
            End If
        End If
    Next item

    ' -- Ajouter � la collection ----------------------------
    Dim vals(9) As String
    vals(0) = dtr: vals(1) = des: vals(2) = node
    vals(3) = fourn: vals(4) = ref: vals(5) = drw
    vals(6) = MTBF: vals(7) = dvlh: vals(8) = dvls

    lignesProduit.Add vals

    Call SIRUS_AddToListViewProduit(vals, TypeMod)

    ' -- Vider les champs apr�s ajout ----------------------
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
' BOUTON "Effacer la s�lection"
' ============================================================
Public Sub SIRUS_EffacerSelectionProduitProduit()

    Dim lb  As MSForms.ListBox
    Dim idx As Long

    Set lb = UserForm.ListBox_Produit
    idx = lb.ListIndex

    If idx = -1 Then
        MsgBox "S�lectionnez une ligne � supprimer.", vbExclamation, "SIRUS"
        Exit Sub
    End If

    ' Reconstruire la collection sans la ligne supprim�e
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
' AJOUTER UNE T�CHE DANS LISTVIEW_TACHES
' ============================================================
Public Sub SIRUS_AjouterTacheATM()

    ' -- R�cup�ration des valeurs ---------------------------
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
    Dim piecesR As String: piecesR = ""   ' � brancher si tu as le champ
    Dim nbops   As String: nbops = ""     ' � brancher si tu as le champ
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
    If cat = "" Then missing = missing & "Cat�gorie / "
    If nq = "" Then missing = missing & "Niveau qualif / "
    If desc = "" Then missing = missing & "Description / "
    If tps = "" Then missing = missing & "Temps estim� / "
    If picto = "" Then missing = missing & "Picto / "
    If txtpic = "" Then missing = missing & "Texte Picto / "

    ' Outillage standard toujours obligatoire
    If outstd = "" Then missing = missing & "Out. Standard / "

    ' Outillage sp�cifique obligatoire si rang 2-1 ou 5-1
    If (rang = "2-1" Or rang = "5-1") And outspec = "" Then
        missing = missing & "Out. Sp�c. (obligatoire rang " & rang & ") / "
    End If

    ' Crit�res dimensionnels � obligatoire si description contient
    ' "Torque", "Tighten" (EN) ou "Serrer" + "couple" (FR)
    Dim descLower As String
    descLower = LCase(desc)
    Dim needCrit As Boolean
    needCrit = (InStr(descLower, "torque") > 0 Or _
                InStr(descLower, "tighten") > 0 Or _
                (InStr(descLower, "serrer") > 0 And InStr(descLower, "couple") > 0))

    If needCrit And critdim = "" Then
        missing = missing & "Crit�res dim. (serrage d�tect�) / "
    End If

    If missing <> "" Then
        MsgBox "Champs obligatoires manquants :" & vbCrLf & _
               Left(missing, Len(missing) - 3), vbExclamation, "SIRUS � ATM"
        Exit Sub
    End If

    ' -- V�rification doublon DTR + Rang + Langue ----------
    Dim lv   As MSComctlLib.ListView
    Dim li   As MSComctlLib.ListItem

    Set lv = UserForm.ListView_Taches

    For Each li In lv.ListItems
        If UCase(li.Text) = UCase(dtr) And _
           UCase(li.SubItems(2)) = UCase(rang) And _
           UCase(li.SubItems(1)) = UCase(langue) Then
            MsgBox "La combinaison DTR '" & dtr & "' / Rang '" & rang & _
                   "' / Langue '" & langue & "' existe d�j�.", _
                   vbExclamation, "SIRUS � ATM"
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
' SUPPRIMER LA LIGNE S�LECTIONN�E DANS LISTVIEW_TACHES
' ============================================================

Public Sub SIRUS_SupprimerTacheATM()

    Dim lv As MSComctlLib.ListView
    Set lv = UserForm.ListView_Taches

    If lv.SelectedItem Is Nothing Then
        MsgBox "S�lectionnez une t�che � supprimer.", vbExclamation, "SIRUS"
        Exit Sub
    End If

    lv.ListItems.Remove lv.SelectedItem.index

End Sub
' ============================================================
' R�CUP�RER LE TYPE DE DEMANDE ACTIF (NavTab)
' ============================================================
Private Function SIRUS_GetTypeActif() As String

    Select Case UserForm.Type_Import_Tab.SelectedItem.index
        Case 0: SIRUS_GetTypeActif = "Cr�ation"
        Case 1: SIRUS_GetTypeActif = "Modification"
        Case 2: SIRUS_GetTypeActif = "Suppression"
        Case Else: SIRUS_GetTypeActif = "Cr�ation"
    End Select

End Function

