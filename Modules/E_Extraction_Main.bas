Attribute VB_Name = "E_Extraction_Main"
'--------------------------------------------------------------------------------------------------------
'
'                                       Generateur Livrables SLI
'                                       Module L_Extraction_Main (Version Refactoris�e)
'                                       Developpement : LE LARGE Thomas (LGM/Alstom)
'                                       Module principal d'extraction ATM et EMMST
'
'-----------------------------------------------------------------------------------------------------------
Option Compare Text
Option Explicit

' =========================================================================
' PROC�DURE PRINCIPALE: MAIN EXTRACTION (wrapper retrocompat)
' =========================================================================
' Conserve l'interface historique pilote par le UserForm (controles legacy
' lb_nomfichier_extr / Extr_Livrables_List / Langue_*_Extr). Delegue le
' travail a Main_Extraction_Params qui peut etre appele directement avec
' des parametres (notamment depuis SIRUS).
Public Sub Main_Extraction()
    Dim fichier As String, livrable As String, lang As String

    fichier = UserForm.lb_nomfichier_extr.Value
    livrable = UserForm.Extr_Livrables_List.Value
    If UserForm.Langue_EN_Extr.Value = True Then
        lang = "EN"
    Else
        lang = "FR"
    End If

    Call Main_Extraction_Params(fichier, livrable, lang, True)
End Sub

' =========================================================================
' Main_Extraction_Params : version parametree (decouplee du UserForm)
' =========================================================================
' Lance l'extraction ATM ou EMMST depuis un fichier livrable, sans
' dependance aux controles UserForm legacy. Utilisee par SIRUS (etape D).
'
' Parametres :
'   sFichier            : chemin complet du fichier livrable
'   sLivrable           : "ATM" ou "EMMST"
'   sLangue             : "FR" ou "EN"
'   bAfficherMsgFin     : True pour afficher le MsgBox de confirmation
'                         (False pour usage SIRUS qui poursuit avec deversement)
'
' Retour : reference vers le Workbook livrable ouvert avec les feuilles
'          "Info_ATM" / "Taches_ATM" / "EMMST" creees. Nothing si erreur.
'          Le caller est responsable de fermer le classeur.
Public Function Main_Extraction_Params(ByVal sFichier As String, _
                                        ByVal sLivrable As String, _
                                        ByVal sLangue As String, _
                                        Optional ByVal bAfficherMsgFin As Boolean = True) As Workbook
    Dim nbFiches As Integer
    Dim validationOK As Boolean

    On Error GoTo ErrorHandler

    TurnOffFunctionality

    ' Propager la langue au global (utilisee dans les SQL ADODB)
    langue = sLangue

    ' ETAPE 1 : Ouverture fichier livrable
    Set wb = Workbooks.Open(sFichier, UpdateLinks:=False, ReadOnly:=False)

    ' ETAPE 2 : Validation fichier
    nbFiches = CompterFichesDTR(wb)
    If nbFiches = 0 Then
        MsgBox "Aucune fiche " & sLivrable & " n'est reconnue dans le fichier.", _
               vbExclamation, "Extraction"
        wb.Close SaveChanges:=False
        Set wb = Nothing
        Set Main_Extraction_Params = Nothing
        TurnOnFunctionality
        Exit Function
    End If

    validationOK = ValiderFormatLivrable(wb, sLivrable, sLangue)
    If Not validationOK Then
        MsgBox "Les fiches " & sLivrable & " ne sont pas au format VBN, extraction impossible.", _
               vbCritical, "Erreur Format"
        wb.Close SaveChanges:=False
        Set wb = Nothing
        Set Main_Extraction_Params = Nothing
        TurnOnFunctionality
        Exit Function
    End If

    ' ETAPE 3 : Extraction selon livrable
    On Error Resume Next
    UserForm.Lb_Nb_Ref.Caption = "Il existe " & nbFiches & " references 'DTR' possedant une fiche " & _
                                  sLivrable & " dans le document."
    On Error GoTo ErrorHandler

    Select Case sLivrable
        Case "ATM"
            Call Extraction_Fiches_ATM(wb, sLangue, nbFiches)
        Case "EMMST"
            Call Extraction_Fiches_EMMST(wb, sLangue, nbFiches)
        Case Else
            MsgBox "Livrable non reconnu : " & sLivrable, vbCritical
            Set Main_Extraction_Params = Nothing
            TurnOnFunctionality
            Exit Function
    End Select

    ' ETAPE 4 : Confirmation (optionnelle)
    If bAfficherMsgFin Then
        MsgBox "Les donnees ont ete extraites avec succes !" & vbCrLf & _
               "Format BDD pret pour import.", vbInformation, "Extraction Terminee"
    End If

    TurnOnFunctionality
    Set Main_Extraction_Params = wb
    Exit Function

ErrorHandler:
    MsgBox "Erreur lors de l'extraction : " & Err.Description, vbCritical
    Set Main_Extraction_Params = Nothing
    TurnOnFunctionality
End Function

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
' V�rifie que les fiches DTR sont bien au format VBN attendu
' Retourne True si format OK, False sinon
    
    Dim ws As Worksheet
    Dim titreAttendu As String
    Dim celluleTitre As Range
    
    On Error GoTo ValidationFailed
    
    ' D�terminer titre attendu selon livrable et langue
    Select Case typeLivrable
        Case "ATM"
            If langue = "FR" Then
                titreAttendu = "Analyse de T�ches de Maintenance"
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
    
    ' V�rifier chaque feuille DTR
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


