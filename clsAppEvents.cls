VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAppEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ==========================================
' Classe: clsAppEvents
' Description: Gestionnaire d'evenements Application
' Surveille les changements dans les feuilles de calcul
' ==========================================
Option Explicit

Private WithEvents App As Application
Attribute App.VB_VarHelpID = -1

' ==========================================
' Initialisation
' ==========================================
Public Sub Initialize()
    Set App = Application
End Sub

' ==========================================
' Termination
' ==========================================
Public Sub Terminate()
    Set App = Nothing
End Sub

' ==========================================
' Evenement: Changement dans une feuille
' Gere RS (colonne K=11) et SIG (colonne P=16)
' Gere aussi la navigation FAO depuis J3
' ==========================================
'Private Sub App_SheetChange(ByVal Sh As Object, ByVal Target As Range)
'    Dim ws As Worksheet
'    Dim changedCell As Range
'    Dim rowIndex As Long
'    Dim isFR As Boolean
'    Dim toDefine As String
'    Dim statutDispo As String
'    Dim colStatutDispo As Long
'    Dim colStatusUpdate As Long
'    Dim colStatusDesc As Long
'    Dim colObsoTreat As Long
'    Dim colFAOStatus As Long
'
'    On Error GoTo ErrorHandler
'
'    ' === GESTION NAVIGATION FAO ===
'    ' Si c'est la feuille FAO ou OAS et cellule I2 (dropdown navigation par reference)
'    If (Sh.name = "FAO" Or Sh.name = "OAS") And Target.Address = "$I$2" Then
'        Call HandleFAONavigationByIndex(Sh, Target.Value)
'        Exit Sub
'    End If
'
'    ' Verifier si c'est une feuille Monitoring Scope
'    If Sh.name <> "Monitoring Scope" Then Exit Sub
'
'    Set ws = Sh
'
'    ' Desactiver les evenements pour eviter les boucles infinies
'    Application.EnableEvents = False
'
'    ' Charger la configuration pour determiner le scope et la langue
'    Call ModConfiguration.LoadConfiguration(ws.Parent)
'
'    isFR = (ModMain.g_CurrentLang = ModMain.Lang_FR)
'    toDefine = IIf(isFR, "A definir", "To be defined")
'
'    ' Determiner les colonnes selon le scope (apres ajout Reparabilite)
'    If ModMain.g_CurrentScope = ModMain.Scope_RS Then
'        ' RS : Statut dispo = K(11), Status update = L(12), Status desc = W(23), Obso treat = X(24)
'        colStatutDispo = 11
'        colStatusUpdate = 12
'        colStatusDesc = 23
'        colObsoTreat = 24
'        colFAOStatus = 0 ' Pas de FAO pour RS
'    Else
'        ' SIG : Statut dispo = P(16), Status update = Q(17), Status desc = AA(27), Obso treat = AB(28), FAO Status = AH(34)
'        colStatutDispo = 16
'        colStatusUpdate = 17
'        colStatusDesc = 27
'        colObsoTreat = 28
'        colFAOStatus = 34
'    End If
'
'    ' Verifier si le changement est dans la colonne Statut de disponibilite ou Statut FAO
'    For Each changedCell In Target.Cells
'        rowIndex = changedCell.Row
'
'        ' Ignorer les en-tetes
'        If rowIndex < 3 Then GoTo NextCell
'
'        ' Gestion du Statut de disponibilite
'        If changedCell.Column = colStatutDispo Then
'            statutDispo = UCase(Trim(changedCell.Value))
'
'            ' Mettre a jour la date de mise a jour du statut
'            ws.Cells(rowIndex, colStatusUpdate).Value = Format(Date, IIf(isFR, "DD/MM/YYYY", "MM/DD/YYYY"))
'
'            ' Si statut n'est pas Current/Actif et colonnes vides, pre-remplir
'            If statutDispo <> "" And statutDispo <> "CURRENT" And statutDispo <> "ACTIF" Then
'                ' Description du statut
'                If Trim(ws.Cells(rowIndex, colStatusDesc).Value) = "" Then
'                    ws.Cells(rowIndex, colStatusDesc).Value = toDefine
'                End If
'
'                ' Traitement obsolescence
'                If Trim(ws.Cells(rowIndex, colObsoTreat).Value) = "" Then
'                    ws.Cells(rowIndex, colObsoTreat).Value = toDefine
'                End If
'            End If
'
'            ' Pour SIG, gerer les colonnes FAO
'            If ModMain.g_CurrentScope = ModMain.Scope_SIG Then
'                Call ModFAO.FillFAOColumns(ws, rowIndex)
'
'                ' Mettre a jour l'onglet FAO
'                Call ModFAO.CreateOrUpdateFAOSheet(ws.Parent)
'            End If
'        End If
'
'        ' Gestion du Statut FAO (SIG uniquement)
'        If colFAOStatus > 0 And changedCell.Column = colFAOStatus Then
'            Call ModFAO.UpdateClosingDate(ws, rowIndex)
'
'            ' Mettre a jour l'onglet FAO
'            Call ModFAO.CreateOrUpdateFAOSheet(ws.Parent)
'        End If
'
'NextCell:
'    Next changedCell
'
'CleanUp:
'    ' Toujours reactiver les evenements
'    Application.EnableEvents = True
'    Exit Sub
'
'ErrorHandler:
'    Debug.Print "App_SheetChange Error: " & Err.description
'    Resume CleanUp
'End Sub

' ==========================================
' Gerer la navigation FAO depuis le dropdown J7 (par reference FAO)
' ==========================================
'Private Sub HandleFAONavigationByIndex(ByVal wsFAO As Worksheet, ByVal selectedFAO As Variant)
'    Dim wsMonitoring As Worksheet
'    Dim lastRowMS As Long, i As Long
'    Dim faoRef As String
'    Dim faoIndex As Long
'    Dim count As Long
'
'    On Error Resume Next
'
'    If Trim(CStr(selectedFAO)) = "" Then Exit Sub
'
'    Set wsMonitoring = wsFAO.Parent.Worksheets("Monitoring Scope")
'    If wsMonitoring Is Nothing Then Exit Sub
'
'    ' Trouver l'index de la FAO selectionnee par sa reference
'    count = 0
'    lastRowMS = wsMonitoring.Cells(rows.count, 3).End(xlUp).Row ' Colonne C
'
'    For i = 3 To lastRowMS
'        faoRef = Trim(wsMonitoring.Cells(i, 29).Value) ' Colonne AC = 29 (FAO Reference)
'
'        If faoRef <> "" And faoRef <> "Non applicable" And faoRef <> "Not applicable" Then
'            count = count + 1
'            If faoRef = CStr(selectedFAO) Then
'                faoIndex = count
'                Exit For
'            End If
'        End If
'    Next i
'
'    ' Si trouve, mettre a jour L5 et rafraichir l'affichage
'    If faoIndex > 0 Then
'        wsFAO.Range("L5").Value = faoIndex
'        Call ModFAO.RefreshFAODisplay(wsFAO)
'    End If
'
'    On Error GoTo 0
'End Sub


