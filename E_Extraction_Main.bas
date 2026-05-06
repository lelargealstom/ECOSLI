Attribute VB_Name = "E_Extraction_Main"
'--------------------------------------------------------------------------------------------------------
'
'                                       Generateur Livrables SLI
'                                       Module L_Extraction_Main (Version Refactorisée)
'                                       Developpement : LE LARGE Thomas (LGM/Alstom)
'                                       Module principal d'extraction ATM et EMMST
'
'-----------------------------------------------------------------------------------------------------------
Option Compare Text
Option Explicit

' =========================================================================
' PROCÉDURE PRINCIPALE: MAIN EXTRACTION
' =========================================================================
Public Sub Main_Extraction()
' Procédure principale qui se lance après validation par le UserForm
' Orchestre l'extraction selon le livrable sélectionné (ATM ou EMMST)
    
    Dim fichier As String
    Dim nbFiches As Integer
    Dim validationOK As Boolean
    Dim Livrable As String
    
    On Error GoTo ErrorHandler
    
    TurnOffFunctionality
    
    ' ======================================================================
    ' ÉTAPE 1 : Récupération paramètres UserForm
    ' ======================================================================
    fichier = UserForm.lb_nomfichier_extr.Value
    Livrable = UserForm.Extr_Livrables_List.Value
    
    If UserForm.Langue_EN_Extr.Value = True Then
        langue = "EN"
    Else
        langue = "FR"
    End If
    
    ' ======================================================================
    ' ÉTAPE 2 : Ouverture fichier livrable
    ' ======================================================================
    Set wb = Workbooks.Open(fichier, UpdateLinks:=False, ReadOnly:=False)
    
    ' ======================================================================
    ' ÉTAPE 3 : Validation fichier
    ' ======================================================================
    nbFiches = CompterFichesDTR(wb)
    
    If nbFiches = 0 Then
        MsgBox "Aucune fiche " & Livrable & " n'est reconnue dans le fichier.", vbExclamation, "Extraction"
        GoTo CleanExit
    End If
    
    validationOK = ValiderFormatLivrable(wb, Livrable, langue)
    
    If Not validationOK Then
        MsgBox "Les fiches " & Livrable & " ne sont pas au format VBN, extraction impossible.", _
               vbCritical, "Erreur Format"
        GoTo CleanExit
    End If
    
    ' ======================================================================
    ' ÉTAPE 4 : Extraction selon livrable
    ' ======================================================================
    UserForm.Lb_Nb_Ref.Caption = "Il existe " & nbFiches & " références 'DTR' possédant une fiche " & _
                                  Livrable & " dans le document."
    
    Select Case Livrable
        Case "ATM"
            Call Extraction_Fiches_ATM(wb, langue, nbFiches)
        
        Case "EMMST"
            Call Extraction_Fiches_EMMST(wb, langue, nbFiches)
        
        Case Else
            MsgBox "Livrable non reconnu : " & Livrable, vbCritical
            GoTo CleanExit
    End Select
    
    ' ======================================================================
    ' ÉTAPE 5 : Confirmation
    ' ======================================================================
    MsgBox "Les données ont été extraites avec succès !" & vbCrLf & _
           "Format BDD prêt pour import.", vbInformation, "Extraction Terminée"
    
CleanExit:
    TurnOnFunctionality
    Exit Sub
    
ErrorHandler:
    MsgBox "Erreur lors de l'extraction : " & Err.description, vbCritical
    TurnOnFunctionality
End Sub

' =========================================================================
' FONCTION: COMPTER FICHES DTR
' =========================================================================
Private Function CompterFichesDTR(ByRef wb As Workbook) As Integer
' Compte le nombre de feuilles contenant "DTR" dans leur nom
    
    Dim ws As Worksheet
    Dim count As Integer
    
    count = 0
    
    For Each ws In wb.Worksheets
        If ws.Name Like "*DTR*" Then
            count = count + 1
        End If
    Next ws
    
    CompterFichesDTR = count
    
End Function

' =========================================================================
' FONCTION: VALIDER FORMAT LIVRABLE
' =========================================================================
Private Function ValiderFormatLivrable(ByRef wb As Workbook, _
                                       ByVal typeLivrable As String, _
                                       ByVal langue As String) As Boolean
' Vérifie que les fiches DTR sont bien au format VBN attendu
' Retourne True si format OK, False sinon
    
    Dim ws As Worksheet
    Dim titreAttendu As String
    Dim celluleTitre As Range
    
    On Error GoTo ValidationFailed
    
    ' Déterminer titre attendu selon livrable et langue
    Select Case typeLivrable
        Case "ATM"
            If langue = "FR" Then
                titreAttendu = "Analyse de Tâches de Maintenance"
            Else
                titreAttendu = "Maintenance Tasks Analysis"
            End If
        
        Case "EMMST"
            If langue = "FR" Then
                titreAttendu = "Emballage Marquage Manutention Stockage Transport"
            Else
                titreAttendu = "Packaging Labelling Handling Storage Transport"
            End If
        
        Case Else
            ValiderFormatLivrable = False
            Exit Function
    End Select
    
    ' Vérifier chaque feuille DTR
    For Each ws In wb.Worksheets
        If ws.Name Like "*DTR*" Then
            Set celluleTitre = ws.Cells(2, 2)
            
            If celluleTitre.Value <> titreAttendu Then
                ValiderFormatLivrable = False
                Exit Function
            End If
        End If
    Next ws
    
    ' Si aucune erreur ? format valide
    ValiderFormatLivrable = True
    Exit Function
    
ValidationFailed:
    ValiderFormatLivrable = False
    
End Function


