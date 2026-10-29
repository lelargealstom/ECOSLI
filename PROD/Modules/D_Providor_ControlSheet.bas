'--------------------------------------------------------------------------------------------------------
'
'                                       Generateur Livrables SLI
'                                       Module D_Providor_ControlSheet
'                                       Developpement : LE LARGE Thomas (LGM/Alstom)
'                                       Module de remplissage de la feuille Control Sheet du Providor
'                                       Creation : 2026-10-29
'
'-----------------------------------------------------------------------------------------------------------
Option Compare Text
Option Explicit

' =========================================================================
' PROCEDURE: REMPLISSAGE CONTROL SHEET
' =========================================================================
Public Sub remplissage_ControlSheet()
' Remplit la ligne 19 de la feuille Control Sheet (sous l en-tete ligne 18) :
'   - Content of Modification : texte structure selon la langue (titres en gras)
'   - Date                     : date du jour
'   - Author                   : nom utilisateur Windows (Application.UserName)
'   - Issue                    : "A1" en dur

    Dim contenu As String
    Dim titres(0 To 5) As String
    Dim i As Integer
    Dim startPos As Long

    If langue = "FR" Then
        contenu = ConstruireContenuControlSheet_FR()
        titres(0) = "General Parameters :"
        titres(1) = "Breakdown Structure :"
        titres(2) = "Material List :"
        titres(3) = "Tool List :"
        titres(4) = "Preventive maintenance :"
        titres(5) = "Reliability & corr maintenance :"
    Else
        contenu = ConstruireContenuControlSheet_EN()
        titres(0) = "General Parameters :"
        titres(1) = "Breakdown structure :"
        titres(2) = "Material List :"
        titres(3) = "Tool List :"
        titres(4) = "Preventive maintenance :"
        titres(5) = "Reliability & corr maintenance :"
    End If

    With CS
        .Cells(19, cs_col.Content).Value = contenu
        .Cells(19, cs_col.Content).WrapText = True

        ' Appliquer le gras sur chaque titre dans la cellule Content
        For i = 0 To UBound(titres)
            startPos = InStr(1, contenu, titres(i))
            If startPos > 0 Then
                .Cells(19, cs_col.Content).Characters(startPos, Len(titres(i))).Font.Bold = True
            End If
        Next i

        .Cells(19, cs_col.datecontrolsheet).Value = Date
        .Cells(19, cs_col.Author).Value = Application.UserName
        .Cells(19, cs_col.Issue).Value = "A1"
    End With

End Sub

' =========================================================================
' FONCTION: CONTENU CONTROL SHEET - FRANCAIS
' =========================================================================
Private Function ConstruireContenuControlSheet_FR() As String
' Retourne le texte complet a poser dans la cellule Content of Modification (FR).
' Les "..." et "…" (U+2026) sont des placeholders a completer manuellement par l utilisateur.

    Dim s As String

    s = "General Parameters :" & vbLf
    s = s & "Profil de mission issu de " & ChrW(&H2026) & vbLf
    s = s & vbLf
    s = s & "Breakdown Structure :" & vbLf
    s = s & "Construit à partir de la PBS référence ..." & vbLf
    s = s & vbLf
    s = s & "Material List :" & vbLf
    s = s & "Prix issus des dernières versions de catalogue spare et repair VBN" & vbLf
    s = s & "Prix fournisseurs convertis en Prix ISP VBN " & vbLf
    s = s & "Attention prix valables sur l'année en cours. " & vbLf
    s = s & "Les prix indiqués par les fournisseurs sont valides jusqu'à fin de garantie et augmenteront après (cela dépend de la fin de série du fournisseur)." & vbLf
    s = s & vbLf
    s = s & "Tool List :" & vbLf
    s = s & "Prix outillages issu de la dernière versions des catalogues VBN." & vbLf
    s = s & vbLf
    s = s & "Preventive maintenance :" & vbLf
    s = s & "Recommendations de tâches préventives issues de ... " & vbLf
    s = s & "Recommendations de tâches préventives issues des livrables fournisseurs" & vbLf
    s = s & vbLf
    s = s & "Reliability & corr maintenance :" & vbLf
    s = s & "Les MTBF indiqués sont les MTBF produits (issues de RAM report produits ou bien de Datasheet pour les COTS)" & vbLf
    s = s & "MTBF issues du RAM report référence ..."

    ConstruireContenuControlSheet_FR = s

End Function

' =========================================================================
' FONCTION: CONTENU CONTROL SHEET - ANGLAIS
' =========================================================================
Private Function ConstruireContenuControlSheet_EN() As String
' Retourne le texte complet a poser dans la cellule Content of Modification (EN).
' Les "..." et "…" (U+2026) et "XXX" sont des placeholders a completer manuellement.
' Les deux-points sont ajoutes partout pour cohérence avec la version FR.

    Dim s As String

    s = "General Parameters :" & vbLf
    s = s & "Mission profile derived from " & ChrW(&H2026) & vbLf
    s = s & vbLf
    s = s & "Breakdown structure :" & vbLf
    s = s & "Built with the file reference XXX" & vbLf
    s = s & vbLf
    s = s & "Material List :" & vbLf
    s = s & "The spare and repair price come from the VBN catalogue revision XXX" & vbLf
    s = s & "The spare price indicated by the supplier is valid until the end of warranty and will increase" & vbLf
    s = s & "afterwards (it depends on the end of serial production by the supplier). " & vbLf
    s = s & "Therefore the repair price through VBN is likely to be economically more interesting than the " & vbLf
    s = s & "spare price in the future (after serial)." & vbLf
    s = s & vbLf
    s = s & "Tool List :" & vbLf
    s = s & "The tool price comes from the VBN catalogue revision XXX" & vbLf
    s = s & vbLf
    s = s & "Preventive maintenance :" & vbLf
    s = s & "Preventive maintenance task recommendations derived from ... " & vbLf
    s = s & "Preventive maintenance task recommendations derived from supplier deliverables. " & vbLf
    s = s & vbLf
    s = s & "Reliability & corr maintenance :" & vbLf
    s = s & "The indicated MTBF values are product MTBF values, taken from product RAM reports or datasheets for COTS items. " & vbLf
    s = s & "MTBF values taken from the reference RAM report ..."

    ConstruireContenuControlSheet_EN = s

End Function
