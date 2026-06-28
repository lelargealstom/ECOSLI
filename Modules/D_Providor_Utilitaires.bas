Attribute VB_Name = "D_Providor_Utilitaires"
'--------------------------------------------------------------------------------------------------------
'
'                                       Generateur Livrables SLI
'                                       Module D_Providor_Utilitaires (Version Refactorisee)
'                                       Developpement : LE LARGE Thomas (LGM/Alstom)
'                                       Module contenant les fonctions utilitaires Providor
'                                       Refactorisation : Fevrier 2026
'
'-----------------------------------------------------------------------------------------------------------
Option Compare Text
Option Explicit
' =========================================================================
' FONCTION: TROUVER INDEX DTR
' =========================================================================
Public Function TrouverIndexDTR(ByVal dtr As String) As Integer
' Trouve l'index d'un DTR dans le tableau tDTR
' Retourne 0 si non trouve

    Dim i As Integer
    
    For i = 1 To UBound(tDTR, 1)
        If tDTR(i, 1) = dtr Then
            TrouverIndexDTR = i
            Exit Function
        End If
    Next
    
    TrouverIndexDTR = 0
    
End Function
' =========================================================================
' PROCEDURE: INITIALISER CADRAGE COLONNES (AUTO-FIND)
' =========================================================================
Public Sub InitialiserCadrageColonnes()
' Initialise toutes les colonnes via Find automatique
' ? appeler APRES copie du template et ouverture du workbook
'
' IMPORTANT : Suppose que les feuilles BS, ML, CM, PM, SL, TL, GP
'             sont deja declarees publiques et pointent vers le bon workbook
    Dim ligneteteBS, ligneteteSL, ligneteteML, ligneteteTL, ligneteteCM, lignetetePM As Integer
    
    'Definition des lignes d'entete dans chaque feuille
    ligneteteBS = 3
    ligneteteSL = 3
    ligneteteML = 3
    ligneteteTL = 3
    ligneteteCM = 3
    lignetetePM = 3
    
    On Error GoTo ErrorHandler
    
    ' ======================================================================
    ' BREAKDOWN STRUCTURE (ligne 4)
    ' ======================================================================
    With BS
        bs_col.NodeLevel = .rows(ligneteteBS).Find(What:="Node level", LookIn:=xlValues, lookat:=xlWhole).Column
        bs_col.NodeIndex = .rows(ligneteteBS).Find(What:="Node's index", LookIn:=xlValues, lookat:=xlWhole).Column
        bs_col.NodeName = .rows(ligneteteBS).Find(What:="Node's name", LookIn:=xlValues, lookat:=xlWhole).Column
        bs_col.ListNode = .rows(ligneteteBS).Find(What:="List node (Hidden column)", LookIn:=xlValues, lookat:=xlWhole).Column
        bs_col.Quantity = .rows(ligneteteBS).Find(What:="Quantity", LookIn:=xlValues, lookat:=xlWhole).Column
        bs_col.Update = .rows(ligneteteBS).Find(What:="Update", LookIn:=xlValues, lookat:=xlWhole).Column
        bs_col.PartNumber = .rows(ligneteteBS).Find(What:="DTR code", LookIn:=xlValues, lookat:=xlWhole).Column
        bs_col.NodeType = .rows(ligneteteBS).Find(What:="Node type", LookIn:=xlValues, lookat:=xlWhole).Column
        bs_col.MaintenanceLevel = .rows(ligneteteBS).Find(What:="Maintenance level", LookIn:=xlValues, lookat:=xlWhole).Column
    End With
    
    ' ======================================================================
    ' MATERIAL LIST (ligne 3)
    ' ======================================================================
    With ML
        ml_col.Update = .rows(ligneteteML).Find(What:="Update", LookIn:=xlValues, lookat:=xlWhole).Column
        ml_col.Name = .rows(ligneteteML).Find(What:="NAME", LookIn:=xlValues, lookat:=xlWhole).Column
        ml_col.PartNumber = .rows(ligneteteML).Find(What:="Alstom Part Number", LookIn:=xlValues, lookat:=xlWhole).Column
        ml_col.Cost = .rows(ligneteteML).Find(What:="Unit-cost for replacement", LookIn:=xlValues, lookat:=xlWhole).Column
        ml_col.CostSource = .rows(ligneteteML).Find(What:="Cost validity date", LookIn:=xlValues, lookat:=xlWhole).Column
        ml_col.CurrencyUnit = .rows(ligneteteML).Find(What:="Currency unit", LookIn:=xlValues, lookat:=xlWhole).Column
        ml_col.MaterialType = .rows(ligneteteML).Find(What:="Material type", LookIn:=xlValues, lookat:=xlWhole).Column
    End With
    
    ' ======================================================================
    ' CORRECTIVE MAINTENANCE (ligne 3)
    ' ======================================================================
    With CM
        cm_col.NodeName = .rows(ligneteteCM).Find(What:="Node's index & name", LookIn:=xlValues, lookat:=xlWhole).Column
        cm_col.NodeIndex = .rows(ligneteteCM).Find(What:="Node index", LookIn:=xlValues, lookat:=xlWhole).Column
        cm_col.SequenceNo = .rows(ligneteteCM).Find(What:="Sequence No", LookIn:=xlValues, lookat:=xlWhole).Column
        cm_col.FailureName = .rows(ligneteteCM).Find(What:="Failure mode name / Item name", LookIn:=xlValues, lookat:=xlWhole).Column
        cm_col.TamponNodeLevel = cm_col.FailureName + 1
        cm_col.TamponDTR = cm_col.FailureName + 2
        cm_col.TamponNodeName = cm_col.FailureName + 3
        cm_col.FailureRate = .rows(ligneteteCM).Find(What:="Failure rate", LookIn:=xlValues, lookat:=xlWhole).Column
        cm_col.FailureMultiplier = .rows(ligneteteCM).Find(What:="Multiplier", LookIn:=xlValues, lookat:=xlWhole).Column
        cm_col.FailureRateUnit = .rows(ligneteteCM).Find(What:="Failure rate unit", LookIn:=xlValues, lookat:=xlWhole).Column
        cm_col.CMNo = .rows(ligneteteCM).Find(What:="CM No", LookIn:=xlValues, lookat:=xlWhole).Column
        cm_col.TaskDescription = .rows(ligneteteCM).Find(What:="Task description", LookIn:=xlValues, lookat:=xlWhole).Column
        cm_col.OnOff = .rows(ligneteteCM).Find(What:="On / Off", LookIn:=xlValues, lookat:=xlWhole).Column
        cm_col.RepairTime = .rows(ligneteteCM).Find(What:="Repair Time (in hours)", LookIn:=xlValues, lookat:=xlWhole).Column
        cm_col.TimeSource = .rows(ligneteteCM).Find(What:="Procedure No", LookIn:=xlValues, lookat:=xlWhole).Column
        cm_col.MaintenanceLevel = .rows(ligneteteCM).Find(What:="Maintenance level", LookIn:=xlValues, lookat:=xlWhole).Column
        cm_col.SkillName = .rows(ligneteteCM).Find(What:="Skill name", LookIn:=xlValues, lookat:=xlWhole).Column
        cm_col.SkillQty = .rows(ligneteteCM).Find(What:="Qty", LookIn:=xlValues, lookat:=xlWhole).Column
        cm_col.MaterialName = .rows(ligneteteCM).Find(What:="Material name", LookIn:=xlValues, lookat:=xlWhole).Column
        cm_col.MaterialQty = cm_col.MaterialName + 1
        cm_col.ToolName = .rows(ligneteteCM).Find(What:="Tool name", LookIn:=xlValues, lookat:=xlWhole).Column
        cm_col.ToolQty = cm_col.ToolName + 1
        cm_col.JustificationTaskFR = .rows(ligneteteCM).Find(What:="Justification_Tache_C_FR", LookIn:=xlValues, lookat:=xlWhole).Column
        cm_col.JustificationTaskEN = .rows(ligneteteCM).Find(What:="Justification_Tache_C_EN", LookIn:=xlValues, lookat:=xlWhole).Column
        cm_col.MTBF = .rows(ligneteteCM).Find(What:="MTBF", LookIn:=xlValues, lookat:=xlWhole).Column
        cm_col.MTBFUnit = .rows(ligneteteCM).Find(What:="MTBF unit", LookIn:=xlValues, lookat:=xlWhole).Column
        cm_col.MTBFSource = .rows(ligneteteCM).Find(What:="Commentaire Alstom (taux de defaillance)", LookIn:=xlValues, lookat:=xlWhole).Column
    End With
    
    ' ======================================================================
    ' PREVENTIVE MAINTENANCE (ligne 3)
    ' ======================================================================
    With PM
        pm_col.NodeIndex = .rows(lignetetePM).Find(What:="Node's index & name", LookIn:=xlValues, lookat:=xlWhole).Column
        pm_col.SequenceNo = .rows(lignetetePM).Find(What:="Sequence No", LookIn:=xlValues, lookat:=xlWhole).Column
        pm_col.TaskDescription = .rows(lignetetePM).Find(What:="Task description", LookIn:=xlValues, lookat:=xlWhole).Column
        pm_col.PrevMaintType = .rows(lignetetePM).Find(What:="Preventive maintenance type", LookIn:=xlValues, lookat:=xlWhole).Column
        pm_col.OnOff = .rows(lignetetePM).Find(What:="On / Off (on/off)", LookIn:=xlValues, lookat:=xlWhole).Column
        pm_col.FrequencyValue = .rows(lignetetePM).Find(What:="Frequency value", LookIn:=xlValues, lookat:=xlWhole).Column
        pm_col.FrequencyUnit = .rows(lignetetePM).Find(What:="Frequency unit", LookIn:=xlValues, lookat:=xlWhole).Column
        pm_col.TaskTime = .rows(lignetetePM).Find(What:="Task time (in hours)", LookIn:=xlValues, lookat:=xlWhole).Column
        pm_col.TimeSource = .rows(lignetetePM).Find(What:="Procedure N?", LookIn:=xlValues, lookat:=xlWhole).Column
        pm_col.SkillName = .rows(lignetetePM).Find(What:="Skill name", LookIn:=xlValues, lookat:=xlWhole).Column
        pm_col.SkillQty = .rows(lignetetePM).Find(What:="Qty", LookIn:=xlValues, lookat:=xlWhole).Column
        pm_col.MaterialName = .rows(lignetetePM).Find(What:="Material name", LookIn:=xlValues, lookat:=xlWhole).Column
        pm_col.MaterialQty = pm_col.MaterialName + 1
        pm_col.JustificationTaskFR = .rows(lignetetePM).Find(What:="Justification de la tache", LookIn:=xlValues, lookat:=xlWhole).Column
        pm_col.JustificationTaskEN = .rows(lignetetePM).Find(What:="Justification of the task", LookIn:=xlValues, lookat:=xlWhole).Column
    End With
    
    ' ======================================================================
    ' SKILL LIST (ligne 3)
    ' ======================================================================
    With SL
        sl_col.Update = .rows(ligneteteSL).Find(What:="Update", LookIn:=xlValues, lookat:=xlWhole).Column
        sl_col.SkillCategory = .rows(ligneteteSL).Find(What:="SKILL CATEGORY", LookIn:=xlValues, lookat:=xlWhole).Column
        sl_col.ProfessionCategory = .rows(ligneteteSL).Find(What:="Profession category", LookIn:=xlValues, lookat:=xlWhole).Column
        sl_col.ManHourCost = .rows(ligneteteSL).Find(What:="Man hour cost", LookIn:=xlValues, lookat:=xlWhole).Column
        sl_col.CurrencyUnit = .rows(ligneteteSL).Find(What:="Currency unit", LookIn:=xlValues, lookat:=xlWhole).Column
    End With
    
    ' ======================================================================
    ' TOOL LIST (ligne 3)
    ' ======================================================================
    With TL
        tl_col.Update = .rows(ligneteteTL).Find(What:="Update", LookIn:=xlValues, lookat:=xlWhole).Column
        tl_col.Name = .rows(ligneteteTL).Find(What:="NAME", LookIn:=xlValues, lookat:=xlWhole).Column
        tl_col.PartNumber = .rows(ligneteteTL).Find(What:="Alstom Part Number", LookIn:=xlValues, lookat:=xlWhole).Column
        tl_col.AcquisitionCost = .rows(ligneteteTL).Find(What:="Unit-cost for replacement", LookIn:=xlValues, lookat:=xlWhole).Column
        tl_col.Currency = .rows(ligneteteTL).Find(What:="Currency unit", LookIn:=xlValues, lookat:=xlWhole).Column
        tl_col.ToolType = .rows(ligneteteTL).Find(What:="Tool type", LookIn:=xlValues, lookat:=xlWhole).Column
    End With
    
    ' ======================================================================
    ' GENERAL PARAMETERS (lignes 1-2)
    ' ======================================================================
    With GP
'        gp_col.Unit = .Cells.Find(What:="Unit used ?", LookIn:=xlValues, LookAt:=xlWhole).Column
        gp_lig.Pjct_acronym = .Columns(1).Find(What:="Project acronym", LookIn:=xlValues, lookat:=xlWhole).Row
        gp_lig.Pjct_name = .Columns(1).Find(What:="Project name", LookIn:=xlValues, lookat:=xlWhole).Row
        gp_lig.Pjct_productname = .Columns(1).Find(What:="Product name", LookIn:=xlValues, lookat:=xlWhole).Row
        gp_lig.Pjct_number = .Columns(1).Find(What:="Project number", LookIn:=xlValues, lookat:=xlWhole).Row
        gp_lig.LCC_length = .Columns(1).Find(What:="LCC length (years)", LookIn:=xlValues, lookat:=xlWhole).Row
        gp_lig.LCC_nbassets = .Columns(1).Find(What:="Number of assets", LookIn:=xlValues, lookat:=xlWhole).Row
        gp_lig.LCC_warranty = .Columns(1).Find(What:="Warranty period (in month)", LookIn:=xlValues, lookat:=xlWhole).Row
        gp_col.OpUnits_Description = .rows(2).Find(What:="Description", LookIn:=xlValues, lookat:=xlWhole).Column
        gp_col.OpUnits_Acronym = .rows(2).Find(What:="Acronym", LookIn:=xlValues, lookat:=xlWhole).Column
        gp_col.OpUnits_YearlyValue = .rows(2).Find(What:="Yearly value", LookIn:=xlValues, lookat:=xlWhole).Column
        gp_col.OpUnits_Unit = .rows(2).Find(What:="Operating dependant unit", LookIn:=xlValues, lookat:=xlWhole).Column
        gp_col.CurrUnit_Currtablename = .rows(35).Find(What:="Currency table name", LookIn:=xlValues, lookat:=xlWhole).Column
        gp_col.CurrUnit_Currname = .rows(35).Find(What:="Currency name", LookIn:=xlValues, lookat:=xlWhole).Column
        gp_col.CurrUnit_Currabbrev = .rows(35).Find(What:="Currency abbreviation", LookIn:=xlValues, lookat:=xlWhole).Column
        gp_col.CurrUnit_changerate = .rows(35).Find(What:="Tool change rate (/Euro)", LookIn:=xlValues, lookat:=xlWhole).Column
        gp_col.CurrUnit_datereviewtoolrate = .rows(35).Find(What:="Date of review for tool rate", LookIn:=xlValues, lookat:=xlWhole).Column
        gp_col.CurrUnit_pjctchangerate = .rows(35).Find(What:="Project change rate (/ Euro)", LookIn:=xlValues, lookat:=xlWhole).Column
        gp_col.CurrUnit_datereviewpjctrate = .rows(35).Find(What:="Date of review for project rate", LookIn:=xlValues, lookat:=xlWhole).Column
    End With
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Erreur initialisation cadrage colonnes : " & vbCrLf & "Colonne non trouvee ou nom d'en-tete incorrect." & vbCrLf & vbCrLf & "Erreur : " & Err.description & vbCrLf & "Numero : " & Err.Number, vbCritical, "Erreur Cadrage Providor"
    
End Sub

' =========================================================================
' FONCTION: VALIDER CADRAGE COLONNES
' =========================================================================
Public Function ValiderCadrageColonnes() As Boolean
' Verifie que toutes les colonnes ont ete initialisees (> 0)
' Retourne True si OK, False sinon
    
    On Error GoTo ValidationFailed
    
    ' Verifier qu'au moins une variable de chaque type est > 0
    If bs_col.NodeLevel = 0 Then GoTo ValidationFailed
    If ml_col.Name = 0 Then GoTo ValidationFailed
    If cm_col.NodeName = 0 Then GoTo ValidationFailed
    If pm_col.NodeIndex = 0 Then GoTo ValidationFailed
    If sl_col.SkillCategory = 0 Then GoTo ValidationFailed
    If tl_col.Name = 0 Then GoTo ValidationFailed
    'If gp_col.ParameterName = 0 Then GoTo ValidationFailed
    
    ValiderCadrageColonnes = True
    Exit Function
    
ValidationFailed:
    ValiderCadrageColonnes = False
    MsgBox "Cadrage colonnes non initialise correctement !" & vbCrLf & vbCrLf & "Verifier que InitialiserCadrageColonnes() a ete appele APRES ouverture du template." & vbCrLf & "Verifier que tous les en-tetes sont presents dans le template.", vbCritical, "Erreur Validation Cadrage"
    
End Function

' =========================================================================
' PROCEDURE: AFFICHER CADRAGE (DEBUG)
' =========================================================================
Public Sub AfficherCadrageColonnes()
' Affiche toutes les colonnes cadrees (pour debug)
    
    Dim msg As String
    
    msg = "=== CADRAGE COLONNES PROVIDOR (Types Structures) ===" & vbCrLf & vbCrLf
    
    msg = msg & "BREAKDOWN STRUCTURE (bs_col) :" & vbCrLf
    msg = msg & "  NodeLevel = " & bs_col.NodeLevel & vbCrLf
    msg = msg & "  NodeIndex = " & bs_col.NodeIndex & vbCrLf
    msg = msg & "  NodeName = " & bs_col.NodeName & vbCrLf
    msg = msg & "  PartNumber = " & bs_col.PartNumber & vbCrLf & vbCrLf
    
    msg = msg & "MATERIAL LIST (ml_col) :" & vbCrLf
    msg = msg & "  Name = " & ml_col.Name & vbCrLf
    msg = msg & "  Cost = " & ml_col.Cost & vbCrLf
    msg = msg & "  MaterialType = " & ml_col.MaterialType & vbCrLf & vbCrLf
    
    msg = msg & "CORRECTIVE MAINTENANCE (cm_col) :" & vbCrLf
    msg = msg & "  NodeName = " & cm_col.NodeName & vbCrLf
    msg = msg & "  MTBF = " & cm_col.MTBF & vbCrLf
    msg = msg & "  RepairTime = " & cm_col.RepairTime & vbCrLf & vbCrLf
    
    msg = msg & "PREVENTIVE MAINTENANCE (pm_col) :" & vbCrLf
    msg = msg & "  NodeIndex = " & pm_col.NodeIndex & vbCrLf
    msg = msg & "  TaskDescription = " & pm_col.TaskDescription & vbCrLf
    msg = msg & "  FrequencyValue = " & pm_col.FrequencyValue & vbCrLf & vbCrLf
    
    msg = msg & "SKILL LIST (sl_col) :" & vbCrLf
    msg = msg & "  SkillCategory = " & sl_col.SkillCategory & vbCrLf
    msg = msg & "  HourlyRate = " & sl_col.HourlyRate & vbCrLf & vbCrLf
    
    msg = msg & "TOOL LIST (tl_col) :" & vbCrLf
    msg = msg & "  Name = " & tl_col.Name & vbCrLf
    msg = msg & "  AcquisitionCost = " & tl_col.AcquisitionCost & vbCrLf & vbCrLf
    
    msg = msg & "GENERAL PARAMETERS (gp_col) :" & vbCrLf
    msg = msg & "  ParameterName = " & gp_col.ParameterName & vbCrLf
    msg = msg & "  ParameterValue = " & gp_col.ParameterValue & vbCrLf
    
    Debug.Print msg
    MsgBox msg, vbInformation, "Cadrage Colonnes Providor"
    
End Sub


