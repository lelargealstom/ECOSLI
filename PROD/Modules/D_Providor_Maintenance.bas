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

        ' Propager ID_PBS du BS vers CM (lignes composants uniquement)
        ' Utilise par RemplirMTBF_Corr pour gerer les DTR dupliques.
        If dictBSRowToIDPBS.Exists(CLng(ligneSrc)) Then
            dictCMRowToIDPBS(CLng(ligneDst)) = dictBSRowToIDPBS(CLng(ligneSrc))
        End If
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
'   - 3. Vide + coloré orange si rien

    Dim j As Integer
    Dim dtrCherche As String
    Dim indexDTR As Integer
    
    ' Lookup direct par dict Ligne CM -> ID_PBS (resout DTR dupliques :
    ' meme DTR dans plusieurs sous-systemes avec MTBF differents).
    If dictCMRowToIDPBS.Exists(CLng(i)) Then
        indexDTR = CLng(dictCMRowToIDPBS(CLng(i)))
    Else
        indexDTR = 0
    End If
    
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
    ' APRES CGR : HIERARCHIE FONCTIONNEL -> INTRINSEQUE
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
' Pose le temps de reparation corrective selon la Gate (source unique, pas de fallback).
'   Avant CGR (PGR, GO ORDER, BA, IGR) : tTps_CorrPrev_Gen_HARD uniquement.
'   Apres CGR (CGR, FEI, IQA)          : tTps_CorrPrev_Ana_HARD uniquement.
' Cellule vide + ColorIndex 44 si source non numerique / nulle, ou si DTR absent de la PBS.

    Dim j As Integer
    Dim dtrCherche As String
    Dim indexDTR As Integer
    Dim source As Variant

    ' Lookup direct par dict Ligne CM -> ID_PBS (resout DTR dupliques :
    ' meme DTR sur plusieurs lignes PBS avec temps de repa differents).
    If dictCMRowToIDPBS.Exists(CLng(i)) Then
        indexDTR = CLng(dictCMRowToIDPBS(CLng(i)))
    Else
        indexDTR = 0
    End If

    If indexDTR > 0 Then
        If AvantCGR Then
            source = tTps_CorrPrev_Gen_HARD(indexDTR, 1)
        Else
            source = tTps_CorrPrev_Ana_HARD(indexDTR, 1)
        End If
    End If

    If IsNumeric(source) And Not IsEmpty(source) And source <> "" And source <> 0 Then
        CM.Cells(i, cm_col.RepairTime) = source
        CM.Columns(cm_col.RepairTime).NumberFormat = "General"
    Else
        CM.Cells(i, cm_col.RepairTime) = ""
        CM.Cells(i, cm_col.RepairTime).Interior.ColorIndex = 44
    End If

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

        If AuMoinsUneValeurNumerique(tDureevie_Ana_HARD) Then
            sourceHARD = tDureevie_Ana_HARD
        Else
            sourceHARD = tDureevie_Gen_HARD
        End If

        If AuMoinsUneValeurNumerique(tDureevie_Ana_SOFT) Then
            sourceSOFT = tDureevie_Ana_SOFT
        Else
            sourceSOFT = tDureevie_Gen_SOFT
        End If

        sourceTpsHARD = tTps_CorrPrev_Ana_HARD
        sourceTpsSOFT = tTps_Prev_Ana_SOFT

    End If
    
    ' Boucle unique sur tous les DTR
    For i = 1 To UBound(tDTR)
        Call InsererLignesPreventives(i, sourceHARD, sourceSOFT, sourceTpsHARD, sourceTpsSOFT)
    Next i
    
    ' Regroupement des lignes par FrequencyValue : pose la Package Reference (A, B, C...)
    Call RemplirPackageReference
    
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
        ' Si le DTR est BOTH, poser dans la ligne SOFT la liste des occurrences annulees
        Call PoserPeriodicCancelledOcc(LastRowInCol(PM, pm_col.TaskDescription), index, sourceHARD(index, 1), sourceSOFT(index, 1))
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
        .Cells(dlPM, pm_col.SkillName) = IIf(langue = "FR", "Electricien", "Electrician")
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
        .Cells(dlPM, pm_col.SkillName) = IIf(langue = "FR", "Electricien", "Electrician")
        
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
Public Function AuMoinsUneValeurNumerique(ByRef t As Variant) As Boolean
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

' =========================================================================
' PROCEDURE: REMPLIR PACKAGE REFERENCE
' =========================================================================
Private Sub RemplirPackageReference()
' Remplit la colonne PackageReference de la feuille Preventive Maintenance.
' Regroupe les lignes par meme FrequencyValue (ordre croissant des frequences uniques).
' Lettres attribuees : 1er niveau A, B, C, ..., Z, puis AA, AB, ..., ZZ, AAA, ...
' Lignes sans frequence numerique exploitable : pose "-".

    Dim dlPM As Long
    Dim i As Long
    Dim valFreq As Variant
    Dim dictLettre As Object
    Dim arrFreq() As Double
    Dim nbFreq As Long
    Dim k As Long
    Dim tmp As Double
    Dim swapped As Boolean

    dlPM = LastRowInCol(PM, pm_col.TaskDescription)
    If dlPM < 4 Then Exit Sub

    ' Collecter les frequences uniques numeriques
    Set dictLettre = CreateObject("Scripting.Dictionary")
    ReDim arrFreq(1 To dlPM - 3)
    nbFreq = 0
    For i = 4 To dlPM
        valFreq = PM.Cells(i, pm_col.FrequencyValue).Value
        If IsNumeric(valFreq) And Not IsEmpty(valFreq) Then
            If Len(CStr(valFreq)) > 0 Then
                If Not dictLettre.Exists(CDbl(valFreq)) Then
                    dictLettre.Add CDbl(valFreq), ""
                    nbFreq = nbFreq + 1
                    arrFreq(nbFreq) = CDbl(valFreq)
                End If
            End If
        End If
    Next i

    If nbFreq = 0 Then
        ' Rien a regrouper : poser "-" sur toutes les lignes
        For i = 4 To dlPM
            PM.Cells(i, pm_col.PackageReference).Value = "-"
        Next i
        Exit Sub
    End If

    ' Trier les frequences par ordre croissant (bulle, nbFreq petit en pratique)
    Do
        swapped = False
        For k = 1 To nbFreq - 1
            If arrFreq(k) > arrFreq(k + 1) Then
                tmp = arrFreq(k)
                arrFreq(k) = arrFreq(k + 1)
                arrFreq(k + 1) = tmp
                swapped = True
            End If
        Next k
    Loop While swapped

    ' Assigner A, B, C,... selon l'ordre croissant
    For k = 1 To nbFreq
        dictLettre(arrFreq(k)) = IndexToLettreExcel(k)
    Next k

    ' Poser la lettre dans chaque ligne (ou "-" si non numerique)
    For i = 4 To dlPM
        valFreq = PM.Cells(i, pm_col.FrequencyValue).Value
        If IsNumeric(valFreq) And Not IsEmpty(valFreq) Then
            If Len(CStr(valFreq)) > 0 Then
                PM.Cells(i, pm_col.PackageReference).Value = dictLettre(CDbl(valFreq))
            Else
                PM.Cells(i, pm_col.PackageReference).Value = "-"
            End If
        Else
            PM.Cells(i, pm_col.PackageReference).Value = "-"
        End If
    Next i

End Sub

' =========================================================================
' FONCTION: INDEX TO LETTRE EXCEL
' =========================================================================
Private Function IndexToLettreExcel(ByVal n As Long) As String
' Convertit un entier n (base 1) en lettre style Excel : 1 -> A, 26 -> Z, 27 -> AA, 28 -> AB, ...
    Dim result As String
    Dim rest As Long
    Do While n > 0
        n = n - 1
        rest = n Mod 26
        result = Chr(65 + rest) & result
        n = n \ 26
    Loop
    IndexToLettreExcel = result
End Function

' =========================================================================
' PROCEDURE: POSER PERIODIC CANCELLED OCC
' =========================================================================
Private Sub PoserPeriodicCancelledOcc(ByVal lignePM As Long, ByVal index As Integer, _
                                      ByVal dvlHARD As Variant, ByVal dvlSOFT As Variant)
' Pose dans la ligne SOFT de PM la liste des occurrences SOFT annulees.
' Une occurrence SOFT (n * DVL_SOFT) est annulee si elle coincide avec un instant HARD
' (m * DVL_HARD) sur la DureeCirculation du projet (ans).
' Rien pose si tHARD_SOFT(index) <> "BOTH", ou si DureeCirculation / DVL_SOFT / DVL_HARD invalides.

    If tHARD_SOFT(index, 1) <> "BOTH" Then Exit Sub
    If Not IsNumeric(DureeCirculation) Then Exit Sub
    If CDbl(DureeCirculation) = 0 Then Exit Sub
    If Not IsNumeric(dvlSOFT) Then Exit Sub
    If CDbl(dvlSOFT) = 0 Then Exit Sub
    If Not IsNumeric(dvlHARD) Then Exit Sub
    If CDbl(dvlHARD) = 0 Then Exit Sub

    Dim duree As Double
    Dim dSOFT As Double
    Dim dHARD As Double
    Dim n As Long
    Dim instantSOFT As Double
    Dim ratio As Double
    Dim listeCancelled As String

    duree = CDbl(DureeCirculation)
    dSOFT = CDbl(dvlSOFT)
    dHARD = CDbl(dvlHARD)
    listeCancelled = ""

    n = 1
    Do
        instantSOFT = n * dSOFT
        If instantSOFT > duree Then Exit Do
        If instantSOFT >= dHARD Then
            ratio = instantSOFT / dHARD
            If Abs(ratio - Int(ratio)) < 0.0000001 Then
                If Len(listeCancelled) = 0 Then
                    listeCancelled = CStr(n)
                Else
                    listeCancelled = listeCancelled & "," & CStr(n)
                End If
            End If
        End If
        n = n + 1
    Loop

    If Len(listeCancelled) > 0 Then
        PM.Cells(lignePM, pm_col.PeriodicCancelledOcc).Value = listeCancelled
    End If

End Sub
