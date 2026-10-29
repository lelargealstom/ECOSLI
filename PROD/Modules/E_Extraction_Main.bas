'--------------------------------------------------------------------------------------------------------
'
'                                       Generateur Livrables SLI
'                                       Module L_Extraction_Main (Version Refactorisee)
'                                       Developpement : LE LARGE Thomas (LGM/Alstom)
'                                       Module principal d'extraction ATM et EMMST
'
'-----------------------------------------------------------------------------------------------------------
Option Compare Text
Option Explicit

' =========================================================================
' PROCEDURE PRINCIPALE: MAIN EXTRACTION (wrapper retrocompat)
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
    Dim ws As Worksheet
    Dim feuillesDispo As String


    TurnOffFunctionality

    ' Propager la langue au global (utilisee dans les SQL ADODB)
    langue = sLangue

    ' ETAPE 1 : Ouverture fichier livrable
    Debug.Print "Main_Extraction_Params : ouverture du fichier '" & sFichier & "'"
    Set wb = Workbooks.Open(sFichier, UpdateLinks:=False, ReadOnly:=False)
    Debug.Print "  -> wb.Name = '" & wb.Name & "', " & wb.Worksheets.count & " feuille(s)."

    ' ETAPE 2 : Validation fichier
    nbFiches = CompterFichesDTR(wb)
    Debug.Print "  CompterFichesDTR -> " & nbFiches & " fiche(s) reconnue(s) (filtre nom Like '*DTR*')"
    If nbFiches = 0 Then
        ' Diagnostic : lister les feuilles disponibles pour identifier
        ' pourquoi aucune ne matche le pattern "*DTR*".
        feuillesDispo = ""
        For Each ws In wb.Worksheets
            If LenB(feuillesDispo) > 0 Then feuillesDispo = feuillesDispo & vbCrLf
            feuillesDispo = feuillesDispo & "  - " & ws.Name
        Next ws
        MsgBox "Aucune fiche " & sLivrable & " n'est reconnue dans le fichier." & vbCrLf & vbCrLf & _
               "Fichier : " & wb.Name & vbCrLf & _
               wb.Worksheets.count & " feuille(s) au total :" & vbCrLf & feuillesDispo & vbCrLf & vbCrLf & _
               "Le compteur cherche les feuilles dont le nom contient 'DTR' (sensible au pattern Like '*DTR*'). " & _
               "Verifier que le fichier est bien un livrable " & sLivrable & " au format VBN.", _
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

    Select Case sLivrable
        Case "ATM"
            ' Bascule vers la version tolerante (Phase 4 entree A).
            ' L'ancien Extraction_Fiches_ATM reste disponible comme
            ' fallback pour usage direct si besoin.
            Call Extraction_Fiches_ATM_Tolerant(wb, sLangue, nbFiches)
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
' Validation tolerante du format livrable VBN (revision 2026-06-10).
' Au lieu d'exiger que la cellule B2 de CHAQUE feuille DTR contienne le
' titre EXACT, on cherche un mot-cle caracteristique du livrable dans
' la zone d'entete A1:J10 d'AU MOINS UNE feuille DTR. Cela tolere :
'   - les variations de titre (Analyse de Taches / Analyse de Taches de Maintenance / MTA)
'   - le titre place dans une autre cellule de l'entete
'   - une eventuelle fiche d'exemple ou de garde non standard
'
' Si une seule fiche DTR contient le mot-cle, on accepte le format et on
' poursuit l'extraction. Les fiches non conformes seront ignorees plus
' loin par les fonctions Extraire* (qui sont elles-memes tolerantes).

    Dim ws As Worksheet
    Dim motsCles As Variant
    Dim mc As Variant
    Dim cell As Range
    Dim nbDTRTestees As Long
    Dim nbDTRMatch As Long


    ' Mots-cles a chercher (au moins 1 doit matcher dans la zone A1:J10)
    Select Case typeLivrable
        Case "ATM"
            If langue = "FR" Then
                motsCles = Array("Maintenance", "ATM", "Taches", "T" & ChrW(226) & "ches")
            Else
                motsCles = Array("Maintenance", "MTA", "Tasks")
            End If
        Case "EMMST"
            If langue = "FR" Then
                motsCles = Array("Emballage", "EMMST", "Marquage", "Stockage")
            Else
                motsCles = Array("Packaging", "PHLST", "Labelling", "Storage")
            End If
        Case Else
            ValiderFormatLivrable = False
            Exit Function
    End Select

    nbDTRTestees = 0
    nbDTRMatch = 0
    For Each ws In wb.Worksheets
        If ws.Name Like "*DTR*" Then
            nbDTRTestees = nbDTRTestees + 1
            ' Cherche au moins un mot-cle dans la zone d'entete (10x10)
            For Each cell In ws.Range("A1:J10")
                For Each mc In motsCles
                    If InStr(1, CStr(cell.Value), CStr(mc), vbTextCompare) > 0 Then
                        nbDTRMatch = nbDTRMatch + 1
                        GoTo NextWs
                    End If
                Next mc
            Next cell
NextWs:
        End If
    Next ws

    Debug.Print "ValiderFormatLivrable (" & typeLivrable & ", " & langue & ") : " & _
                nbDTRMatch & " / " & nbDTRTestees & " feuille(s) DTR ont matche au moins un mot-cle."

    ' On accepte si au moins 1 feuille DTR contient un mot-cle.
    ValiderFormatLivrable = (nbDTRMatch > 0)
End Function




