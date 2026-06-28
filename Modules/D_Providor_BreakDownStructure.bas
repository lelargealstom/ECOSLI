Attribute VB_Name = "D_Providor_BreakDownStructure"
'--------------------------------------------------------------------------------------------------------
'
'                                       Generateur Livrables SLI
'                                       Module D_Providor_BreakdownStructure (Version Refactorisée)
'                                       Developpement : LE LARGE Thomas (LGM/Alstom)
'                                       Module de remplissage BS et Material List
'                                       Refactorisation : Février 2026
'
'-----------------------------------------------------------------------------------------------------------
Option Compare Text
Option Explicit

' =========================================================================
' PROCÉDURE: REMPLISSAGE BREAKDOWN STRUCTURE
' =========================================================================
Public Sub remplissage_BreakdownStructure()
' Remplit la Breakdown Structure avec :
' - Node Index incrémental
' - Insertion automatique des Family (sous-systèmes)
' - Node Type ? Node Level
' - Gestion des changements de sous-système

    Dim i As Integer
    Dim curSS As String, lastSS As String
    Dim addFamily As Boolean
    Dim NodeLevel As Integer
    
    On Error Resume Next
    
    BS.Activate
    
    ' Vérifications préalables
    If IsEmpty(tSS) Then Exit Sub
    If UBound(tSS, 1) = 0 Then Exit Sub
    
    ' Calculer dernière ligne
    dlBS = LastRowInCol(BS, bs_col.NodeName) ' Col C
    If dlBS < 2 Then dlBS = 2
    
    ' Insérer ligne Projet si nécessaire
    If BS.Cells(dlBS, 3).Value <> Code_Projet Then
        nextIdx = 1
        BS.Cells(dlBS + 1, bs_col.NodeLevel).Value = 1                    ' Node Level
        BS.Cells(dlBS + 1, bs_col.NodeIndex).Value = Format(nextIdx, "000000000")  ' Node Index
        BS.Cells(dlBS + 1, bs_col.NodeName).Value = Code_Projet          ' Node Name
        BS.Cells(dlBS + 1, bs_col.Quantity).Value = 1                    ' Quantity
        BS.Cells(dlBS + 1, bs_col.Update).Value = "A0"                 ' Update
        BS.Cells(dlBS + 1, bs_col.NodeType).Value = "Breakdown structure" ' Node Type
        dlBS = dlBS + 1
    Else
        nextIdx = CLng(BS.Cells(dlBS, bs_col.NodeIndex).Value) + 1
    End If
    
    ' Insérer première Family si nécessaire
    curSS = NzStr(tSS(1, 1))
    If NzStr(BS.Cells(dlBS, bs_col.NodeType).Value) <> "Family" Or NzStr(BS.Cells(dlBS, bs_col.NodeName).Value) <> curSS Then
        BS.Cells(dlBS + 1, bs_col.NodeLevel).Value = 2
        BS.Cells(dlBS + 1, bs_col.NodeIndex).Value = Format(nextIdx + 1, "000000000")
        BS.Cells(dlBS + 1, bs_col.NodeName).Value = curSS
        BS.Cells(dlBS + 1, bs_col.Quantity).Value = 1
        BS.Cells(dlBS + 1, bs_col.Update).Value = "A0"
        BS.Cells(dlBS + 1, bs_col.NodeType).Value = "Family"
        dlBS = dlBS + 1
        nextIdx = nextIdx + 1
    End If
    lastSS = curSS
    
    ' Boucle sur tous les composants
    For i = 1 To UBound(tSS, 1)
        curSS = NzStr(tSS(i, 1))
        addFamily = (curSS <> lastSS)
        
        ' Insérer Family si changement sous-système
        If addFamily Then
            nextIdx = nextIdx + 1
            BS.Cells(dlBS + 1, bs_col.NodeLevel).Value = 2
            BS.Cells(dlBS + 1, bs_col.NodeIndex).Value = Format(nextIdx, "000000000")
            BS.Cells(dlBS + 1, bs_col.NodeName).Value = curSS
            BS.Cells(dlBS + 1, bs_col.Quantity).Value = 1
            BS.Cells(dlBS + 1, bs_col.Update).Value = "A0"
            BS.Cells(dlBS + 1, bs_col.NodeType).Value = "Family"
            dlBS = dlBS + 1
        End If
        
        ' Insérer ligne composant
        NodeLevel = NodeTypeToLevel(NzStr(tNodeType(i, 1)))
        nextIdx = nextIdx + 1
        
        BS.Cells(dlBS + 1, bs_col.NodeLevel).Value = NodeLevel                       ' A: Node Level
        BS.Cells(dlBS + 1, bs_col.NodeIndex).Value = Format(nextIdx, "000000000")    ' B: Node Index
        BS.Cells(dlBS + 1, bs_col.NodeName).Value = NzStr(tDES_MEF(i, 1))           ' C: Node Name
        BS.Cells(dlBS + 1, bs_col.Quantity).Value = SafeNum(tQTE(i, 1), 1)          ' E: Quantity
        BS.Cells(dlBS + 1, bs_col.Update).Value = "A0"                            ' I: Update
        BS.Cells(dlBS + 1, bs_col.PartNumber).Value = NzStr(tDTR(i, 1))              ' O: DTR
        BS.Cells(dlBS + 1, bs_col.NodeType).Value = NzStr(tNodeType(i, 1))         ' R: Node Type
        BS.Cells(dlBS + 1, bs_col.MaintenanceLevel).Value = NodeLevel                      ' S: Maint Level
        
        dlBS = dlBS + 1
        lastSS = curSS
    Next i
    
End Sub

' =========================================================================
' PROCÉDURE: REMPLISSAGE MATERIAL LIST (ORCHESTRATION)
' =========================================================================
Public Sub remplissage_MaterialList()
' Orchestre le remplissage de la Material List :
' 1. Calcul coefficients prix fournisseur
' 2. Copie données BS ? ML
' 3. Ajout lignes _REPAIR pour Stdex
' 4. Suppression supports
' 5. Remplissage prix (hiérarchie)
' 6. Gestion doublons

    On Error Resume Next
    
    Application.ThisWorkbook.Activate
    
    ' Calculer dernière ligne BS
    dlBS = LastRowInCol(BS, bs_col.NodeName)
    
    ' 1. Calculer coefficients prix fournisseur
    Call CalculerCoefficients
    
    ' 2. Copier données BS vers ML
    Call CopierDonneesBS_vers_ML
    
    ' 3. Remplir Material Type selon Node Type
    Call RemplirMaterialType
    
    ' Recalculer limite
    dlML = LastRowInCol(ML, ml_col.Name)
    
    ' 4. Ajouter lignes _REPAIR pour Stdex
    Call AjouterLignesRepair
    
    ' 5. Supprimer supports
    Call SupprimerSupports
    
    ' Recalculer limite
    dlML = LastRowInCol(ML, ml_col.Name)
    
    ' Récupérer tableaux ML pour gestion doublons
    DTR_tab_ML = ML.Range(ML.Cells(4, ml_col.PartNumber), ML.Cells(dlML, ml_col.PartNumber))
    Des_tab_ML = ML.Range(ML.Cells(4, ml_col.Name), ML.Cells(dlML, ml_col.Name))
    
    ' 6. Remplir prix avec hiérarchie
    Call RemplirPrix
    
    ' 7. Gérer doublons
    Call GererDoublons
    
    ' 8. Colorier lignes sans DTR
    Call ColorierLignesSansDTR
    
    ' Calculer et retour
    Application.Calculate
    ThisWorkbook.Activate
    ML.Range("A1").Select
    
End Sub

' =========================================================================
' PROCÉDURE: CALCULER COEFFICIENTS
' =========================================================================
Private Sub CalculerCoefficients()
' Calcule et applique les coefficients sur les prix fournisseur

    Dim i As Integer
    
    ' Coefficient Spare
    Coeff_Price_FRS_Spare = DS1.Cells(DS1.Cells.Find(What:="Cout spare FOURNISSEUR (Not ISP)", lookat:=xlWhole).Row - 1, _
                                       DS1.Cells.Find(What:="Cout spare FOURNISSEUR (Not ISP)", lookat:=xlWhole).Column)
    
    For i = 1 To UBound(tPrix_FRS_Spare)
        If IsNumeric(tPrix_FRS_Spare(i, 1)) And tPrix_FRS_Spare(i, 1) <> 0 Then
            tPrix_FRS_Spare(i, 1) = tPrix_FRS_Spare(i, 1) * Coeff_Price_FRS_Spare
        End If
    Next i
    
    ' Coefficient Repair
    Coeff_Price_FRS_Repair = DS1.Cells(DS1.Cells.Find(What:="Cout Repair FOURNISSEUR (Not ISP)", lookat:=xlWhole).Row - 1, _
                                        DS1.Cells.Find(What:="Cout Repair FOURNISSEUR (Not ISP)", lookat:=xlWhole).Column)
    
    For i = 1 To UBound(tPrix_FRS_Repair)
        If IsNumeric(tPrix_FRS_Repair(i, 1)) And tPrix_FRS_Repair(i, 1) <> 0 Then
            tPrix_FRS_Repair(i, 1) = tPrix_FRS_Repair(i, 1) * Coeff_Price_FRS_Repair
        End If
    Next i
    
End Sub

' =========================================================================
' PROCÉDURE: COPIER DONNÉES BS VERS ML
' =========================================================================
Private Sub CopierDonneesBS_vers_ML()
' Copie Node Name et DTR de BS vers ML
    
    Dim ligneSrc As Long
    Dim ligneDst As Long
    
    ' Désactiver calculs pour performance
    Application.Calculation = xlCalculationManual
    
    ' Copier ligne par ligne
    For i = 5 To dlBS
        ligneSrc = i
        ligneDst = i - 1  ' Décalage de 1 (BS ligne 5 ? ML ligne 4)
        
        ' Node Name
        ML.Cells(ligneDst, ml_col.Name).Value = BS.Cells(ligneSrc, bs_col.NodeName).Value
        
        ' Part Number / DTR
        ML.Cells(ligneDst, ml_col.PartNumber).Value = BS.Cells(ligneSrc, bs_col.PartNumber).Value
    Next i
    
End Sub

' =========================================================================
' PROCÉDURE: REMPLIR MATERIAL TYPE
' =========================================================================
Private Sub RemplirMaterialType()
' Remplit Material Type selon Node Type de BS
    
    Dim i As Integer
    
    For i = 5 To dlBS
        Select Case BS.Cells(i, bs_col.NodeType)
            Case "Family"
                ML.Cells(i - 1, ml_col.MaterialType) = ""
            Case "LRU", "SRU"
                ML.Cells(i - 1, ml_col.MaterialType) = "Stdex"
            Case "LRUNR", "SRUNR"
                ML.Cells(i - 1, ml_col.MaterialType) = "Disca"
            Case Else
                ML.Cells(i - 1, ml_col.MaterialType) = ""
        End Select
    Next i
    
End Sub

' =========================================================================
' PROCÉDURE: AJOUTER LIGNES REPAIR
' =========================================================================
Private Sub AjouterLignesRepair()
' Ajoute une ligne _REPAIR pour chaque Stdex
    
    Dim i As Integer
    
    i = 4
    Do While i <= dlML
        If ML.Cells(i, ml_col.MaterialType) = "Stdex" Then
            ' Insérer ligne
            ML.rows(i + 1).Insert
            ML.Cells(i + 1, ml_col.Name) = ML.Cells(i, ml_col.Name) & "_REPAIR"
            ML.Cells(i + 1, ml_col.PartNumber) = ML.Cells(i, ml_col.PartNumber)
            ML.Cells(i + 1, ml_col.MaterialType) = "Repai"
            i = i + 1
            dlML = LastRowInCol(ML, ml_col.Name)
        End If
        i = i + 1
    Loop
    
End Sub

' =========================================================================
' PROCÉDURE: SUPPRIMER SUPPORTS
' =========================================================================
Private Sub SupprimerSupports()
' Supprime les lignes contenant "support" dans le nom
    
    Dim i As Integer
    
    For i = dlML To 4 Step -1
        If InStr(1, ML.Cells(i, ml_col.Name).Value, "support") <> 0 Then
            ML.rows(i).Delete
        End If
    Next i
    
End Sub

' =========================================================================
' PROCÉDURE: REMPLIR PRIX (ORCHESTRATION)
' =========================================================================
Private Sub RemplirPrix()
' Orchestre le remplissage des prix selon Material Type
    
    Dim i As Integer
    
    For i = 4 To dlML
        ' Remplir champs communs
        ML.Cells(i, ml_col.Update) = "A0"
        ML.Cells(i, ml_col.CurrencyUnit) = "EUR"
        
        ' Remplir prix selon type
        Select Case ML.Cells(i, ml_col.MaterialType)
            Case "Repai"
                Call RemplirPrixRepair(i)
            Case "Stdex", "Disca"
                Call RemplirPrixSpare(i)
        End Select
        
        ' Convertir en valeur
        ML.Cells(i, ml_col.Cost) = ML.Cells(i, ml_col.Cost).Value
        
    Next i
    
End Sub

' =========================================================================
' PROCÉDURE: REMPLIR PRIX REPAIR
' =========================================================================
Private Sub RemplirPrixRepair(ByVal ligne As Integer)
' Remplit prix Repair avec hiérarchie :
' FRS ? Catalogue ? Manuel ? Archive ? Formule ISP*0.67

    Dim j As Integer
    Dim Designation As String
    
    Designation = Replace(ML.Cells(ligne, ml_col.Name).Value, "_REPAIR", "")
    
    For j = 1 To UBound(tDTR, 1)
        If ML.Cells(ligne, ml_col.PartNumber).Value = tDTR(j, 1) And Designation = tDES_MEF(j, 1) Then
            
            ' Hiérarchie prix
            If TrySetPrice(ligne, tPrix_FRS_Repair(j, 1), "Cout Repair fournisseur") Then Exit Sub
            If TrySetPrice(ligne, tPrix_Cat_Repair(j, 1), "Prix catalogue Repa") Then Exit Sub
            If TrySetPrice(ligne, tPrix_RepairProjet_Archive(j, 1), "Prix d'un projet archivé en BDD") Then Exit Sub
            
            ' Si aucun prix trouvé : ISP * 0.67
            If Not (IsNumeric(ML.Cells(ligne, ml_col.Cost))) Or ML.Cells(ligne, ml_col.Cost) = 0 Then
                ML.Cells(ligne, ml_col.Cost).Value = ML.Cells(ligne - 1, ml_col.Cost).Value * 0.67
                Call AddCommentML(ligne, "Aucun prix connu, ISP * 0,67")
                ML.Cells(ligne, ml_col.Cost).Interior.ColorIndex = 15
            End If
            
            Exit For
        End If
    Next j
    
End Sub

' =========================================================================
' PROCÉDURE: REMPLIR PRIX SPARE
' =========================================================================
Private Sub RemplirPrixSpare(ByVal ligne As Integer)
' Remplit prix Spare/Disca avec hiérarchie :
' FRS ? Catalogue ? Manuel ? Archive

    Dim j As Integer
    
    For j = 1 To UBound(tDTR, 1)
        If ML.Cells(ligne, ml_col.PartNumber).Value = tDTR(j, 1) And ML.Cells(ligne, ml_col.Name).Value = tDES_MEF(j, 1) Then
            
            ' Hiérarchie prix
            If TrySetPrice(ligne, tPrix_FRS_Spare(j, 1), "Cout Spare fournisseur") Then Exit Sub
            If TrySetPrice(ligne, tPrix_Cat_Spare(j, 1), "Prix catalogue Spare") Then Exit Sub
            If TrySetPrice(ligne, tPrix_SpareProjet_Archive(j, 1), "Prix d'un projet archivé en BDD") Then Exit Sub
            
            Exit For
        End If
    Next j
    
End Sub

' =========================================================================
' FONCTION: TRY SET PRICE
' =========================================================================
Private Function TrySetPrice(ByVal ligne As Integer, ByVal prix As Variant, ByVal commentaire As String) As Boolean
' Tente de définir un prix si valide
' Retourne True si prix défini, False sinon

    If IsNumeric(prix) And prix <> 0 Then
        ML.Cells(ligne, ml_col.Cost).Value = prix
        Call AddCommentML(ligne, commentaire)
        TrySetPrice = True
    Else
        TrySetPrice = False
    End If
    
End Function

' =========================================================================
' PROCÉDURE: ADD COMMENT ML
' =========================================================================
Private Sub AddCommentML(ByVal ligne As Integer, ByVal texte As String)
' Ajoute un commentaire sur la cellule L (prix)

    On Error Resume Next
    ML.Cells(ligne, ml_col.Cost).AddComment
    ML.Cells(ligne, ml_col.Cost).Comment.Visible = False
    ML.Cells(ligne, ml_col.Cost).Comment.Text Text:=texte
    
End Sub

' =========================================================================
' PROCÉDURE: GÉRER DOUBLONS
' =========================================================================
Private Sub GererDoublons()
' Détecte et marque les doublons :
' - Couples SPARE-REPAIR
' - DTR en doublon

    Dim i As Integer
    
    For i = 4 To dlML
        ' Test doublon couple SPARE-REPAIR
        If TestDoublonCouple(i) Then
            ML.Cells(i, ml_col.Name).AddComment
            ML.Cells(i, ml_col.Name).Comment.Visible = False
            ML.Cells(i, ml_col.Name).Comment.Text Text:="Doublon Couple SPARE - REPAIR"
        Else
            ' Test doublon DTR uniquement (si pas de couple)
            If Not ExisteCouple(i) And TestDoublonDTR(i) Then
                ML.Cells(i, ml_col.PartNumber).AddComment
                ML.Cells(i, ml_col.PartNumber).Comment.Visible = False
                ML.Cells(i, ml_col.PartNumber).Comment.Text Text:="Doublon DTR"
            End If
        End If
    Next i
    
End Sub

' =========================================================================
' FONCTION: TEST DOUBLON COUPLE
' =========================================================================
Private Function TestDoublonCouple(ByVal ligne As Integer) As Boolean
' Teste si doublon couple SPARE-REPAIR existe
    
    Dim Designation As String
    Dim designationRepair As String
    
    Designation = ML.Cells(ligne, ml_col.Name).Value
    
    If InStr(Designation, "_REPAIR") = 0 Then
        ' Ligne SPARE : chercher _REPAIR
        designationRepair = Designation & "_REPAIR"
        TestDoublonCouple = (IsInArrayNb(designationRepair, Des_tab_ML) >= 2 And InStr(ML.Cells(ligne, ml_col.PartNumber).Value, "DTR") >= 1)
    Else
        ' Ligne REPAIR : chercher SPARE
        TestDoublonCouple = (IsInArrayNb(Designation, Des_tab_ML) >= 2 And InStr(ML.Cells(ligne, ml_col.PartNumber).Value, "DTR") >= 1)
    End If
    
End Function

' =========================================================================
' FONCTION: EXISTE COUPLE
' =========================================================================
Private Function ExisteCouple(ByVal ligne As Integer) As Boolean
' Vérifie si un couple SPARE-REPAIR existe pour cette ligne
    
    Dim Designation As String
    Dim designationRepair As String
    
    Designation = ML.Range("C" & ligne).Value
    
    If InStr(Designation, "_REPAIR") = 0 Then
        designationRepair = Designation & "_REPAIR"
        ExisteCouple = (IsInArrayNb(designationRepair, Des_tab_ML) >= 1 And InStr(ML.Cells(ligne, ml_col.PartNumber).Value, "DTR") >= 1)
    Else
        ExisteCouple = (IsInArrayNb(Designation, Des_tab_ML) >= 1 And InStr(ML.Cells(ligne, ml_col.PartNumber).Value, "DTR") >= 1)
    End If
    
End Function

' =========================================================================
' FONCTION: TEST DOUBLON DTR
' =========================================================================
Private Function TestDoublonDTR(ByVal ligne As Integer) As Boolean
' Teste si doublon DTR existe
    
    TestDoublonDTR = (InStr(ML.Cells(ligne, ml_col.PartNumber).Value, "DTR") >= 1 And IsInArrayNb(ML.Cells(ligne, ml_col.PartNumber).Value, DTR_tab_ML) >= 2)
    
End Function

' =========================================================================
' PROCÉDURE: COLORIER LIGNES SANS DTR
' =========================================================================
Private Sub ColorierLignesSansDTR()
' Colorie en gris les lignes sans DTR
    
    Dim i As Integer
    
    For i = 4 To dlML
        If IsEmpty(ML.Cells(i, ml_col.PartNumber)) Then
            ML.Range("A" & i & ":DX" & i).Interior.ColorIndex = 15
        End If
    Next i
    
End Sub

