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
Public Sub SIRUS_Parcourir_Fichier()
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
Public Sub SIRUS_Generer_FeuilleAnalyse()
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
            ' Source des cles : Combo Projet pour ProjetSLI/ProjetASL,
            ' TextBox DTR multiligne pour les autres tables.
            If sTable = "ProjetSLI" Or sTable = "ProjetASL" Then
                Dim sProjet As String
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
            Call SIRUS_AjouterFeuilleAuCombo("SIRUS_" & sTable & "_Modifier")

        Case "Extraction"
            Dim sFichier As String
            sFichier = CStr(UserForm.Txt_SIRUS_Fichier.Value)
            If LenB(sFichier) = 0 Then
                MsgBox "Selectionner un fichier livrable.", vbExclamation, "SIRUS"
                Exit Sub
            End If
            ' Etape D a venir : appel a SIRUS_DeverserExtractionDansSIRUS.
            ' Pour l'instant on cree la feuille vierge SIRUS_<Table>_Extraction
            ' pour valider le squelette UI/dispatch.
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
            MsgBox "Extraction : feuilles SIRUS creees vierges. Le deversement " & _
                   "des donnees du fichier sera branche a la prochaine etape (D).", _
                   vbInformation, "SIRUS - Extraction"
    End Select
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
Public Sub SIRUS_Analyser_Feuille()
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

    ' Resume statut + activer envoi si taux OK
    UserForm.Lbl_SIRUS_StatutAnalyse.Caption = _
        "Taux de completude (MeQ) : " & Format(taux * 100, "0.0") & "%"

    UserForm.Btn_SIRUS_Envoyer.Enabled = (taux >= 0.7)

    ' Remplir la ListView des NOK
    Call SIRUS_RemplirListViewResultats(sFeuille)
End Sub

' -------------------------------------------------------------------------
' SIRUS_GotoLigneFautive : double-clic ListView -> Activer la cellule cle
' -------------------------------------------------------------------------
Public Sub SIRUS_GotoLigneFautive()
    Dim lv As MSComctlLib.ListView
    Dim li As MSComctlLib.ListItem
    Dim ligne As Long
    Dim sFeuille As String
    Dim ws As Worksheet

    Set lv = UserForm.ListView_SIRUS_Resultats
    Set li = lv.SelectedItem
    If li Is Nothing Then Exit Sub

    ligne = CLng(li.Text)
    sFeuille = CStr(UserForm.Cmb_SIRUS_FeuilleAnalyse.Value)
    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets(sFeuille)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Application.Goto ws.Cells(ligne, 2), Scroll:=True
End Sub

' -------------------------------------------------------------------------
' SIRUS_Envoyer_Demande : Phase 5, TODO
' -------------------------------------------------------------------------
Public Sub SIRUS_Envoyer_Demande()
    MsgBox "Envoi ATMAP : a brancher Phase 5 (etape G).", vbInformation, "SIRUS"
End Sub

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

    With lv
        .View = lvwReport
        .FullRowSelect = True
        .Gridlines = True
        .LabelEdit = lvwManual
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "Ligne", 50, lvwColumnLeft
        .ColumnHeaders.Add , , "Cle", 120, lvwColumnLeft
        .ColumnHeaders.Add , , "Alertes", 350, lvwColumnLeft
        .ListItems.Clear
    End With
End Sub

Private Sub SIRUS_RemplirListViewResultats(ByVal sFeuille As String)
    ' Lit la colonne Statut (col 1) de la feuille active et liste les NOK
    ' dans la ListView. Cle = col 2 (1er champ saisie). Detail = commentaire
    ' de cellule.
    Dim lv As MSComctlLib.ListView
    Dim ws As Worksheet
    Dim li As MSComctlLib.ListItem
    Dim ligne As Long, ligneMax As Long
    Dim statut As String, cle As String, detail As String
    Dim cellStatut As Range

    Set lv = UserForm.ListView_SIRUS_Resultats
    lv.ListItems.Clear

    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets(sFeuille)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    ligneMax = ws.Cells(ws.rows.count, 2).End(xlUp).Row
    For ligne = 3 To ligneMax
        statut = CStr(ws.Cells(ligne, 1).Value)
        If InStr(1, statut, "NOK", vbTextCompare) > 0 _
           Or InStr(1, statut, "non trouv", vbTextCompare) > 0 Then
            cle = CStr(ws.Cells(ligne, 2).Value)
            detail = ""
            Set cellStatut = ws.Cells(ligne, 1)
            If Not cellStatut.Comment Is Nothing Then
                detail = cellStatut.Comment.Text
                detail = Replace(detail, vbLf, " | ")
            End If
            Set li = lv.ListItems.Add(, , CStr(ligne))
            li.SubItems(1) = cle
            li.SubItems(2) = detail
        End If
    Next ligne
End Sub
