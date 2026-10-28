'--------------------------------------------------------------------------------------------------------
'
'                                       Generateur Livrables SLI
'                                       Module D_Providor_Maintenance (Version Refactorisée)
'                                       Developpement : LE LARGE Thomas (LGM/Alstom)
'                                       Module de remplissage maintenance corrective et préventive
'                                       Refactorisation : Février 2026
'
'-----------------------------------------------------------------------------------------------------------
Option Compare Text
Option Explicit

' =========================================================================
' PROCÉDURE: REMPLISSAGE CORRECTIVE MAINTENANCE
' =========================================================================
Public Sub remplissage_corr()
' Remplit la feuille de maintenance corrective (CM)

    Dim i As Integer, j As Integer
    
    On Error Resume Next
    
    Application.ThisWorkbook.Activate
    
    ' Vérifications
    If ML.Range("C4") = "" Then
        MsgBox "Merci de vous assurer que l'onglet Material List est bien rempli."
        Exit Sub
    End If
    
    ' Calculer limites
    dlBS = LastRowInCol(BS, bs_col.NodeName)
    dlML = LastRowInCol(ML, ml_col.Name)
    
    ' Copier données BS vers CM avec des colonnes tampons
    Call CopierDonneesBS_vers_CM
    
    ' Calculer limite CM
    dlCM = LastRowInCol(CM, cm_col.NodeName)
    
    ' Remplir données pour chaque ligne
    For i = 4 To dlCM
        If IsNumeric(CM.Cells(i, cm_col.TamponNodeLevel)) And CM.Cells(i, cm_col.TamponNodeLevel) > 2 Then
            Call RemplirLigneCorrectiveMaintenance(i)
        Else
            CM.Range("A" & i & ":EG" & i).Interior.ColorIndex = 15
        End If
    Next i
    
    ' Format colonnes
    CM.Columns(cm_col.FailureRate).NumberFormat = "0.00"
    CM.Columns(cm_col.FailureMultiplier).NumberFormat = "0.00E+00"
    
    ' Calculer limite CM
    dlCM = LastRowInCol(CM, cm_col.NodeName)
    
    ' Supprimer colonnes tampons
    CM.Range(Cells(4, cm_col.TamponNodeLevel), Cells(dlCM, cm_col.TamponNodeType)).ClearContents
    
    ' Supprimer supports
    For i = dlCM To 4 Step -1
        If InStr(1, CM.Cells(i, cm_col.NodeName), "support") <> 0 Then
            CM.rows(i).Delete
        End If
    Next i
    
    ' Modifier List Node BS
    BS.Range(BS.Cells(4, bs_col.ListNode), BS.Cells(dlBS, bs_col.ListNode)).ClearContents
    BS.Cells(4, bs_col.ListNode).FormulaLocal = "=B4&"" ""&@Implement_ListNode(A4)&C4"
    BS.Range(BS.Cells(4, bs_col.ListNode), BS.Cells(dlBS, bs_col.ListNode)).FillDown
    
    ' Modifier Node Index CM
    CM.Range(Cells(4, cm_col.NodeIndex), Cells(dlCM, cm_col.NodeIndex)).ClearContents
    CM.Cells(4, cm_col.NodeIndex).FormulaLocal = "=@REL_Parent(A4)"
    CM.Range(CM.Cells(4, cm_col.NodeIndex), CM.Cells(dlCM, cm_col.NodeIndex)).FillDown
    CM.Range(CM.Cells(4, cm_col.NodeIndex), CM.Cells(dlCM, cm_col.NodeIndex)).Interior.ColorIndex = 3
    
    'On nettoie les colonnes tampons de la CM
    CM.Cells(4, cm_col.TamponNodeLevel).Select
    CM.Range(CM.Cells(4, cm_col.TamponNodeLevel), CM.Cells(dlCM, cm_col.TamponNodeType)).ClearContents
    CM.Cells(4, cm_col.NodeName).Select
    
End Sub
Private Sub CopierDonneesBS_vers_CM()
' Copie données de Breakdown Structure vers Corrective Maintenance

    Dim ligneSrc As Long
    Dim ligneDst As Long
    
    ' Copier données ligne par ligne
    For i = 5 To dlBS
        ligneSrc = i
        ligneDst = i - 1  ' BS ligne 5 ? CM ligne 4
        
        ' Node's name (BS col D ? CM col A)
        CM.Cells(ligneDst, cm_col.NodeName).Value = BS.Cells(ligneSrc, bs_col.ListNode).Value
        
        ' Colonnes Tampon
        ' Ajuster selon ta structure CM (vérifier quel champ correspond)
        CM.Cells(ligneDst, cm_col.TamponNodeLevel).Value = BS.Cells(ligneSrc, bs_col.NodeLevel).Value
        ' Part number / DTR (BS col O ? CM col F)
        CM.Cells(ligneDst, cm_col.TamponDTR).Value = BS.Cells(ligneSrc, bs_col.PartNumber).Value
        'List Node
        CM.Cells(ligneDst, cm_col.TamponNodeName).Value = BS.Cells(ligneSrc, bs_col.NodeName).Value
        'Node Type
        CM.Cells(ligneDst, cm_col.TamponNodeType).Value = BS.Cells(ligneSrc, bs_col.NodeType).Value
        
    Next i
    
    ' Restaurer header "Node's name" en CM col A ligne 3
    CM.Cells(3, cm_col.NodeName).Value = "Node's name"
    
End Sub
' =========================================================================
' PROCÉDURE: REMPLIR LIGNE CORRECTIVE MAINTENANCE
' =========================================================================
Private Sub RemplirLigneCorrectiveMaintenance(ByVal i As Integer)
' Remplit une ligne de maintenance corrective
    
    ' MTBF et Temps Réparation
    Call RemplirMTBF_Corr(i)
    Call RemplirTempsReparation_Corr(i)
    
    ' Unité MTBF : POWER HUP HOUR par défaut
    CM.Cells(i, cm_col.MTBFUnit) = "POWER-UP HOUR"
    
    ' Failure Mode et Task Description
    Call RemplirFailureMode(i)
    
    ' Failure Rate (basé sur MTBF)
    If IsNumeric(CM.Cells(i, cm_col.MTBF)) And CM.Cells(i, cm_col.MTBF) <> 0 Then
        CM.Cells(i, cm_col.FailureRate) = (1 / CM.Cells(i, cm_col.MTBF)) * 1000000
        CM.Cells(i, cm_col.FailureRate).Value = CM.Cells(i, cm_col.FailureRate)
        CM.Cells(i, cm_col.FailureRate).NumberFormat = "0.0000"
    End If
    
    ' Others data
    CM.Cells(i, cm_col.SequenceNo) = 1
    CM.Cells(i, cm_col.FailureMultiplier) = 0.000001
    CM.Cells(i, cm_col.FailureRateUnit) = "PH"
    CM.Cells(i, cm_col.CMNo) = 1
    CM.Cells(i, cm_col.MaintenanceLevel) = CM.Cells(i, cm_col.TamponNodeLevel)
    
    ' Skill Name
    Call RemplirSkillName(i)
    
    ' Material Name
    Call RemplirMaterialName(i)
    
    ' Tool List (Dataplug adapter)
    Call RemplirToolList(i)
    
End Sub

' =========================================================================
' PROCÉDURE: REMPLIR MTBF CORRECTIVE (CORRIGÉE)
' =========================================================================
Private Sub RemplirMTBF_Corr(ByVal i As Integer)
' Remplit MTBF selon gate en utilisant directement les tableaux variants
'
' Logique AvantCGR (PGR, GO ORDER, BA, IGR):
'   - MTBF générique uniquement (tMTBF_gen)
'
' Logique AprèsCGR (CGR, FEI, IQA):
'   - 1. MTBF fonctionnel (tMTBF_fct)
'   - 2. MTBF intrinsèque (tMTBF_Intr)
'   - 3. MTBF générique (tMTBF_gen)
'   - 4. Vide + coloré orange si rien

    Dim j As Integer
    Dim dtrCherche As String
    Dim indexDTR As Integer
    
    dtrCherche = CM.Cells(i, cm_col.TamponDTR).Value
    
    ' Trouver l'index du DTR dans les tableaux
    indexDTR = 0
    For j = LBound(tDTR) To UBound(tDTR)
        If tDTR(j, 1) = dtrCherche Then
            indexDTR = j
            Exit For
        End If
    Next j
    
    ' Si DTR non trouvé
    If indexDTR = 0 Then
        CM.Cells(i, cm_col.MTBF).Value = ""
        CM.Cells(i, cm_col.MTBFSource) = "DTR non trouvé dans PBS"
        CM.Cells(i, cm_col.MTBF).Interior.ColorIndex = 44
        Exit Sub
    End If
    
    ' =====================================================================
    ' AVANT CGR : MTBF GÉNÉRIQUE UNIQUEMENT
    ' =====================================================================
    If AvantCGR Then
        If IsNumeric(tMTBF_gen(indexDTR, 1)) And tMTBF_gen(indexDTR, 1) <> 0 And Not IsNull(tMTBF_gen(indexDTR, 1)) Then
            CM.Cells(i, cm_col.MTBF).Value = tMTBF_gen(indexDTR, 1)
            CM.Columns(cm_col.MTBF).NumberFormat = "General"
            CM.Cells(i, cm_col.MTBFSource) = "MTBF générique"
        Else
            CM.Cells(i, cm_col.MTBF).Value = ""
            CM.Cells(i, cm_col.MTBFSource) = "MTBF générique inconnu en BDD"
            CM.Cells(i, cm_col.MTBF).Interior.ColorIndex = 44
        End If
        Exit Sub
    End If
    
    ' =====================================================================
    ' APRÈS CGR : HIÉRARCHIE FONCTIONNEL ? INTRINSÈQUE ? GÉNÉRIQUE
    ' =====================================================================
    
    ' Tentative 1 : MTBF Fonctionnel
    If IsNumeric(tMTBF_fct(indexDTR, 1)) And tMTBF_fct(indexDTR, 1) <> 0 And Not IsNull(tMTBF_fct(indexDTR, 1)) Then
        CM.Cells(i, cm_col.MTBF).Value = tMTBF_fct(indexDTR, 1)
        CM.Columns(cm_col.MTBF).NumberFormat = "General"
        CM.Cells(i, cm_col.MTBFSource) = "MTBF fonctionnel du RAM Report"
        Exit Sub
    End If
    
    ' Tentative 2 : MTBF Intrinsèque
    If IsNumeric(tMTBF_Intr(indexDTR, 1)) And tMTBF_Intr(indexDTR, 1) <> 0 And Not IsNull(tMTBF_Intr(indexDTR, 1)) Then
        CM.Cells(i, cm_col.MTBF).Value = tMTBF_Intr(indexDTR, 1)
        CM.Columns(cm_col.MTBF).NumberFormat = "General"
        CM.Cells(i, cm_col.MTBFSource) = "MTBF intrinsèque du RAM Report"
        Exit Sub
    End If
    
    ' Tentative 3 : MTBF Générique
    If IsNumeric(tMTBF_gen(indexDTR, 1)) And tMTBF_gen(indexDTR, 1) <> 0 And Not IsNull(tMTBF_gen(indexDTR, 1)) Then
        CM.Cells(i, cm_col.MTBF).Value = tMTBF_gen(indexDTR, 1)
        CM.Columns(cm_col.MTBF).NumberFormat = "General"
        CM.Cells(i, cm_col.MTBFSource) = "MTBF générique"
        Exit Sub
    End If
    
    ' Aucun MTBF trouvé
    CM.Cells(i, cm_col.MTBF).Value = ""
    CM.Cells(i, cm_col.MTBFSource) = "Pas de MTBF trouvé en BDD"
    CM.Cells(i, cm_col.MTBF).Interior.ColorIndex = 44
    
End Sub

' =========================================================================
' PROCÉDURE: REMPLIR FAILURE MODE
' =========================================================================
Private Sub RemplirFailureMode(ByVal i As Integer)
' Remplit Failure Mode et Task Description selon langue
    
    If langue = "FR" Then
        If CM.Cells(i, cm_col.TamponNodeLevel) > 2 Then
            CM.Cells(i, cm_col.FailureName) = CM.Cells(i, cm_col.TamponNodeName) & " HS"
        End If
        CM.Cells(i, cm_col.TaskDescription) = "REMPLACER " & CM.Cells(i, cm_col.TamponNodeName)
    Else
        If CM.Cells(i, cm_col.TamponNodeLevel) > 2 Then
            CM.Cells(i, cm_col.FailureName) = CM.Cells(i, cm_col.TamponNodeName) & " OUT OF ORDER"
        End If
        CM.Cells(i, cm_col.TaskDescription) = "REPLACE " & CM.Cells(i, cm_col.TamponNodeName)
    End If
    
End Sub

' =========================================================================
' PROCÉDURE: REMPLIR SKILL NAME
' =========================================================================
Private Sub RemplirSkillName(ByVal i As Integer)
' Remplit Skill Name selon maintenance level
    
    If CM.Cells(i, cm_col.MaintenanceLevel) = "3" Then
        CM.Cells(i, cm_col.OnOff) = "On"
    ElseIf CM.Cells(i, cm_col.MaintenanceLevel) = "4" Then
        CM.Cells(i, cm_col.OnOff) = "Off"
    End If
    
    If CM.Cells(i, cm_col.TamponNodeLevel) <> 2 Then
        If langue = "FR" Then
            CM.Cells(i, cm_col.SkillName) = "Electricien"
        Else
            CM.Cells(i, cm_col.SkillName) = "Electrician"
        End If
        CM.Cells(i, cm_col.SkillQty) = 1
    End If
    
End Sub

' =========================================================================
' PROCÉDURE: REMPLIR TEMPS RÉPARATION CORRECTIVE
' =========================================================================
Private Sub RemplirTempsReparation_Corr(ByVal i As Integer)
' Remplit temps réparation selon gate en utilisant directement les tableaux variants
'
' Logique AvantCGR (PGR, GO ORDER, BA, IGR):
'   - Temps générique HARD uniquement (tTps_CorrPrev_Gen_HARD)
'
' Logique AprèsCGR (CGR, FEI, IQA):
'   - 1. Temps analytique HARD (tTps_CorrPrev_Ana_HARD)
'   - 2. Temps générique HARD (tTps_CorrPrev_Gen_HARD)
'   - 3. Temps famille (Dataplug 0.45, autre 0.3)

    Dim j As Integer
    Dim dtrCherche As String
    Dim indexDTR As Integer
    Dim tempsDefaut As Double
    
    dtrCherche = CM.Cells(i, cm_col.TamponDTR)
    
    ' Trouver l'index du DTR
    indexDTR = 0
    For j = LBound(tDTR) To UBound(tDTR)
        If tDTR(j, 1) = dtrCherche Then
            indexDTR = j
            Exit For
        End If
    Next j
    
    ' Si DTR non trouvé
    If indexDTR = 0 Then
        tempsDefaut = CalculerTempsDefaut(CM.Cells(i, cm_col.NodeName))
        CM.Cells(i, cm_col.RepairTime) = tempsDefaut
        CM.Columns(cm_col.RepairTime).NumberFormat = "General"
        CM.Cells(i, cm_col.RepairTime).AddComment
        CM.Cells(i, cm_col.RepairTime).Comment.Visible = False
        CM.Cells(i, cm_col.RepairTime).Comment.Text Text:="DTR non trouvé - Temps famille"
        CM.Cells(i, cm_col.RepairTime).Interior.ColorIndex = 44
        Exit Sub
    End If
    
    ' =====================================================================
    ' AVANT CGR : TEMPS GÉNÉRIQUE UNIQUEMENT
    ' =====================================================================
    If AvantCGR Then
        If IsNumeric(tTps_CorrPrev_Gen_HARD(indexDTR, 1)) And tTps_CorrPrev_Gen_HARD(indexDTR, 1) <> 0 And _
           Not IsNull(tTps_CorrPrev_Gen_HARD(indexDTR, 1)) Then
            CM.Cells(i, cm_col.RepairTime) = tTps_CorrPrev_Gen_HARD(indexDTR, 1)
            CM.Columns(cm_col.RepairTime).NumberFormat = "General"
            CM.Cells(i, cm_col.RepairTime).AddComment
            CM.Cells(i, cm_col.RepairTime).Comment.Visible = False
            CM.Cells(i, cm_col.RepairTime).Comment.Text Text:="Temps générique"
        Else
            tempsDefaut = CalculerTempsDefaut(CM.Cells(i, cm_col.NodeName))
            CM.Cells(i, cm_col.RepairTime) = tempsDefaut
            CM.Columns(cm_col.RepairTime).NumberFormat = "General"
            CM.Cells(i, cm_col.RepairTime).AddComment
            CM.Cells(i, cm_col.RepairTime).Comment.Visible = False
            CM.Cells(i, cm_col.RepairTime).Comment.Text Text:="Pas de temps en BDD - Temps famille"
            CM.Cells(i, cm_col.RepairTime).Interior.ColorIndex = 44
        End If
        Exit Sub
    End If
    
    ' =====================================================================
    ' APRÈS CGR : HIÉRARCHIE ANALYTIQUE ? GÉNÉRIQUE ? FAMILLE
    ' =====================================================================
    
    ' Tentative 1 : Temps analytique
    If IsNumeric(tTps_CorrPrev_Ana_HARD(indexDTR, 1)) And tTps_CorrPrev_Ana_HARD(indexDTR, 1) <> 0 And _
       Not IsNull(tTps_CorrPrev_Ana_HARD(indexDTR, 1)) Then
        CM.Cells(i, cm_col.RepairTime) = tTps_CorrPrev_Ana_HARD(indexDTR, 1)
        CM.Columns(cm_col.RepairTime).NumberFormat = "General"
        CM.Cells(i, cm_col.RepairTime).AddComment
        CM.Cells(i, cm_col.RepairTime).Comment.Visible = False
        CM.Cells(i, cm_col.RepairTime).Comment.Text Text:="Temps analytique RAM Report"
        Exit Sub
    End If
    
    ' Tentative 2 : Temps générique
    If IsNumeric(tTps_CorrPrev_Gen_HARD(indexDTR, 1)) And tTps_CorrPrev_Gen_HARD(indexDTR, 1) <> 0 And _
       Not IsNull(tTps_CorrPrev_Gen_HARD(indexDTR, 1)) Then
        CM.Cells(i, cm_col.RepairTime) = tTps_CorrPrev_Gen_HARD(indexDTR, 1)
        CM.Columns(cm_col.RepairTime).NumberFormat = "General"
        CM.Cells(i, cm_col.RepairTime).AddComment
        CM.Cells(i, cm_col.RepairTime).Comment.Visible = False
        CM.Cells(i, cm_col.RepairTime).Comment.Text Text:="Temps générique"
        Exit Sub
    End If
    
    ' Tentative 3 : Temps famille
    tempsDefaut = CalculerTempsDefaut(CM.Cells(i, cm_col.NodeName))
    CM.Cells(i, cm_col.RepairTime) = tempsDefaut
    CM.Columns(cm_col.RepairTime).NumberFormat = "General"
    CM.Cells(i, cm_col.RepairTime).AddComment
    CM.Cells(i, cm_col.RepairTime).Comment.Visible = False
    CM.Cells(i, cm_col.RepairTime).Comment.Text Text:="Pas de temps en BDD - Temps famille"
    CM.Cells(i, cm_col.RepairTime).Interior.ColorIndex = 44
    
End Sub

' =========================================================================
' PROCÉDURE: REMPLIR MATERIAL NAME
' =========================================================================
Private Sub RemplirMaterialName(ByVal i As Integer)
' Remplit Material Name selon type (Stdex ? _REPAIR)
    
    Dim sType As String
    
    If CM.Cells(i, cm_col.FailureName) = "-" Then Exit Sub
    
    sType = UCase$(Trim$(CM.Cells(i, cm_col.TamponNodeType).Value))
    If sType = "LRUNR" Or sType = "SRUNR" Then
        CM.Cells(i, cm_col.MaterialName) = CM.Cells(i, cm_col.TamponNodeName)
    ElseIf sType = "LRU" Or sType = "SRU" Then
        CM.Cells(i, cm_col.MaterialName) = CM.Cells(i, cm_col.TamponNodeName) & "_REPAIR"
    End If
        
    CM.Cells(i, cm_col.MaterialQty) = 1
    
End Sub

' =========================================================================
' PROCÉDURE: REMPLIR TOOL LIST
' =========================================================================
Private Sub RemplirToolList(ByVal i As Integer)
' Ajoute Dataplug adapter si nécessaire
    
    If InStr(1, CM.Cells(i, cm_col.FailureName), "DATAPLUG") <> 0 And InStr(1, CM.Cells(i, cm_col.FailureName), "GES") = 0 Then
        CM.Cells(i, cm_col.ToolName) = "DATAPLUG ADAPTER"
        CM.Cells(i, cm_col.ToolQty) = 1
    End If
    
End Sub

' =========================================================================
' PROCÉDURE: REMPLISSAGE PRÉVENTIVE MAINTENANCE (REFACTORISÉE)
' =========================================================================
Public Sub remplissage_prev()
' Remplit la feuille de maintenance préventive
' Version refactorisée : unifie les gates en utilisant des paramètres

    Dim i As Integer
    Dim sourceHARD As Variant, sourceSOFT As Variant
    Dim sourceTpsHARD As Variant, sourceTpsSOFT As Variant
    
    On Error Resume Next
    
    TurnOffFunctionality
    
    ' Préparer LIST_NODE
    Call PreparerListNode
    
    Application.ThisWorkbook.Activate
    
    ' Vérifications
    If ML.Range("C4") = "" Then
        MsgBox "Merci de vous assurer que l'onglet Material List est bien rempli."
        GoTo CleanUp
    End If
    
    ' Déterminer sources selon gate
    If AvantCGR Then
        ' Avant CGR : données génériques
        sourceHARD = tDureevie_Gen_HARD
        sourceSOFT = tDureevie_Gen_SOFT
        sourceTpsHARD = tTps_CorrPrev_Gen_HARD
        sourceTpsSOFT = tTps_Prev_Gen_SOFT
    Else
        ' Après CGR : données analytiques
        Dim msgFallback As String
        msgFallback = ""

        If AuMoinsUneValeurNumerique(tDureevie_Ana_HARD) Then
            sourceHARD = tDureevie_Ana_HARD
        Else
            sourceHARD = tDureevie_Gen_HARD
            msgFallback = msgFallback & "- Duree de vie HARDWARE : aucune valeur analytique (CGR) renseignee, utilisation de la colonne generique (offre et PGR)." & vbCrLf
        End If

        If AuMoinsUneValeurNumerique(tDureevie_Ana_SOFT) Then
            sourceSOFT = tDureevie_Ana_SOFT
        Else
            sourceSOFT = tDureevie_Gen_SOFT
            msgFallback = msgFallback & "- Duree de vie SOFTWARE : aucune valeur analytique (CGR) renseignee, utilisation de la colonne generique (offre et PGR)." & vbCrLf
        End If

        sourceTpsHARD = tTps_CorrPrev_Ana_HARD
        sourceTpsSOFT = tTps_Prev_Ana_SOFT

        If Len(msgFallback) > 0 Then
            MsgBox "Preventive Maintenance - Fallback sur les donnees Offre/PGR :" & vbCrLf & vbCrLf & msgFallback, _
                   vbInformation, "Providor"
        End If
    End If
    
    ' Boucle unique sur tous les DTR
    For i = 1 To UBound(tDTR)
        Call InsererLignesPreventives(i, sourceHARD, sourceSOFT, sourceTpsHARD, sourceTpsSOFT)
    Next i
    
    ' Supprimer dernière ligne uniquement si le tableau est plein
    dlPM = LastRowInCol(PM, pm_col.TaskDescription)
    
    If dlPM > 3 Then
        PM.Cells(dlPM, pm_col.NodeIndex).EntireRow.Delete
    End If
    
CleanUp:
    TurnOnFunctionality
    
End Sub

' =========================================================================
' PROCÉDURE: PRÉPARER LIST NODE
' =========================================================================
Private Sub PreparerListNode()
' Prépare le tableau LIST_NODE depuis BS
    
    Dim j As Integer
    
    ReDim LIST_NODE(1 To UBound(tDTR), 1 To 1)
    dlBS = BS.Cells(3, bs_col.NodeName).End(xlDown).row + 1
    
    TurnOnFunctionality
    
    j = 1
    For i = 4 To dlBS
        If BS.Cells(i, bs_col.NodeLevel).Value >= 3 Then
            For j = j To UBound(LIST_NODE)
                LIST_NODE(j, 1) = BS.Cells(i, bs_col.ListNode)
                j = j + 1
                Exit For
            Next j
        End If
    Next i
    
    TurnOffFunctionality
    
End Sub

' =========================================================================
' PROCÉDURE: INSÉRER LIGNES PRÉVENTIVES
' =========================================================================
Private Sub InsererLignesPreventives(ByVal index As Integer, _
                                     ByRef sourceHARD As Variant, ByRef sourceSOFT As Variant, _
                                     ByRef sourceTpsHARD As Variant, ByRef sourceTpsSOFT As Variant)
' Détermine et insère les lignes préventives nécessaires (HARD, SOFT ou les deux)
    
    Dim needHARD As Boolean, needSOFT As Boolean
    
    ' Déterminer quelles lignes créer
    needHARD = sourceHARD(index, 1) <> 0 And IsNumeric(sourceHARD(index, 1))
    needSOFT = sourceSOFT(index, 1) <> 0 And IsNumeric(sourceSOFT(index, 1))
    
    ' Insérer lignes nécessaires
    If needHARD Then
        Call InsererLigneHARD(index, sourceHARD, sourceTpsHARD)
    End If
    
    If needSOFT Then
        Call InsererLigneSOFT(index, sourceSOFT, sourceTpsSOFT)
    End If
    
End Sub

' =========================================================================
' PROCÉDURE: INSÉRER LIGNE HARD
' =========================================================================
Private Sub InsererLigneHARD(ByVal index As Integer, ByRef Source As Variant, ByRef sourceTps As Variant)
' Insère une ligne HARD dans PM
    
    Dim dlPM As Long
    
    dlPM = LastRowInCol(PM, pm_col.TaskDescription) + 1
    
    With PM
        ' Infos génériques
        .Cells(dlPM, pm_col.NodeIndex) = LIST_NODE(index, 1)
        .Cells(dlPM, pm_col.SequenceNo) = 1
        .Cells(dlPM, pm_col.PrevMaintType) = "Progr"
        .Cells(dlPM, pm_col.FrequencyUnit) = "Y"
        .Cells(dlPM, pm_col.SkillQty) = 1
        
        ' Infos HARD
        .Cells(dlPM, pm_col.TaskDescription) = "REPLACE " & tDES_MEF(index, 1)
        .Cells(dlPM, pm_col.OnOff) = IIf(tNTI(index, 1) = 3, "On", "Off")
        .Cells(dlPM, pm_col.FrequencyValue) = Source(index, 1)
        .Cells(dlPM, pm_col.SkillName) = "Electrician"
        .Cells(dlPM, pm_col.MaterialName) = tDES_MEF(index, 1)
        .Cells(dlPM, pm_col.MaterialQty) = 1
        
        ' Temps réparation
        Call RemplirTempsReparationPrev(dlPM, index, sourceTps)
    End With
    
End Sub

' =========================================================================
' PROCÉDURE: INSÉRER LIGNE SOFT
' =========================================================================
Private Sub InsererLigneSOFT(ByVal index As Integer, ByRef Source As Variant, ByRef sourceTps As Variant)
' Insère une ligne SOFT dans PM
    
    Dim dlPM As Long
    
    dlPM = LastRowInCol(PM, pm_col.TaskDescription) + 1
    
    With PM
        ' Infos génériques
        .Cells(dlPM, pm_col.NodeIndex) = LIST_NODE(index, 1)
        .Cells(dlPM, pm_col.SequenceNo) = 1
        .Cells(dlPM, pm_col.PrevMaintType) = "Progr"
        .Cells(dlPM, pm_col.FrequencyUnit) = "Y"
        .Cells(dlPM, pm_col.SkillQty) = 1
        
        ' Infos SOFT
        .Cells(dlPM, pm_col.TaskDescription) = "SOFTWARE UPLOAD " & tDES_MEF(index, 1)
        .Cells(dlPM, pm_col.OnOff) = IIf(tNTI(index, 1) = 3, "On", "Off")
        .Cells(dlPM, pm_col.FrequencyValue) = Source(index, 1)
        .Cells(dlPM, pm_col.SkillName) = "Technician"
        
        ' Temps réparation
        Call RemplirTempsReparationPrev(dlPM, index, sourceTps)
    End With
    
End Sub

' =========================================================================
' PROCÉDURE: REMPLIR TEMPS RÉPARATION PRÉVENTIVE
' =========================================================================
Private Sub RemplirTempsReparationPrev(ByVal ligne As Long, ByVal index As Integer, _
                                       ByRef source As Variant)
' Remplit le temps de reparation depuis une source unique (Gate + HARD/SOFT determinent
' la colonne en amont dans remplissage_prev). Pas de fallback vers une autre source.
' Cellule vide si source non numerique, ou vide, ou #N/A, ou 0.

    If IsNumeric(source(index, 1)) And source(index, 1) <> "" And _
       source(index, 1) <> "#N/A" And source(index, 1) <> 0 Then
        PM.Cells(ligne, pm_col.TaskTime) = source(index, 1)
        Call AddCommentPM(ligne, IIf(AvantCGR, "Tps Generique", "Tps Analytique"))
    Else
        Call AddCommentPM(ligne, "Aucun temps de reparation renseigne")
        PM.Cells(ligne, pm_col.TaskTime).Interior.ColorIndex = 44
    End If

End Sub
' =========================================================================
' FONCTION: CALCULER TEMPS DÉFAUT
' =========================================================================
Private Function CalculerTempsDefaut(ByVal designation As String) As Double
' Calcule temps réparation par défaut selon famille
' Dataplug (hors GES) ? 0.45h
' Autre ? 0.3h
    
    If InStr(1, designation, "DATAPLUG") <> 0 And InStr(1, designation, "GES") = 0 Then
        CalculerTempsDefaut = 0.45
    Else
        CalculerTempsDefaut = 0.3
    End If
    
End Function
' =========================================================================
' FONCTION: CALCULER TEMPS FAMILLE
' =========================================================================
Private Function CalculerTempsFamille(ByVal designation As String, ByVal systeme As String) As Variant
' Calcule temps réparation selon famille de produit
' Hiérarchie : Dataplug (0.45) ? PACIS (0.3) ? Autre (0.2) ? Vide ("")
    
    ' Dataplug (hors GES)
    If InStr(1, designation, "DATAPLUG") <> 0 And InStr(1, designation, "GES") = 0 Then
        CalculerTempsFamille = 0.45
        Exit Function
    End If
    
    ' PACIS family
    If InStr(1, systeme, "PA") <> 0 Or InStr(1, systeme, "PIS") <> 0 Or InStr(1, systeme, "CCTV") <> 0 Then
        CalculerTempsFamille = 0.3
        Exit Function
    End If
    
    ' Pas de sous-système
    If systeme = "" Then
        CalculerTempsFamille = ""
        Exit Function
    End If
    
    ' Autre
    CalculerTempsFamille = 0.2
    
End Function

' =========================================================================
' PROCÉDURE: ADD COMMENT PM
' =========================================================================
Private Sub AddCommentPM(ByVal ligne As Long, ByVal texte As String)
' Ajoute un commentaire sur la cellule M (temps)

    On Error Resume Next
    PM.Cells(ligne, pm_col.TimeSource) = texte
    
End Sub

' =========================================================================
' FONCTION: AU MOINS UNE VALEUR NUMERIQUE
' =========================================================================
Private Function AuMoinsUneValeurNumerique(ByRef t As Variant) As Boolean
' Verifie si le tableau 2D (nx1) contient au moins une valeur numerique.
' Un 0 compte comme valeur renseignee. Les cellules vides ou chaines vides ne comptent pas.
    Dim k As Long
    For k = LBound(t, 1) To UBound(t, 1)
        If Not IsEmpty(t(k, 1)) Then
            If Len(CStr(t(k, 1))) > 0 Then
                If IsNumeric(t(k, 1)) Then
                    AuMoinsUneValeurNumerique = True
                    Exit Function
                End If
            End If
        End If
    Next k
End Function
