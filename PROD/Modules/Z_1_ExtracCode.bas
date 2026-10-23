' ==========================================
' Module: ModPublish
' Description: Génère le .xlam depuis le .xlsm courant
' A placer dans le .xlsm uniquement (pas dans le complément)
' ==========================================
Option Explicit

Public Sub PublishToXlam()
    Dim xlsmPath   As String
    Dim xlamPath   As String
    Dim xlsmName   As String
    Dim xlamName   As String
    Dim wbSource   As Workbook
    Dim wbTarget   As Workbook
    Dim result     As VbMsgBoxResult

    'On Error GoTo ErrorHandler

    ' --- Chemins source et destination ---
    xlsmPath = GetLocalPath(ThisWorkbook.FullName)
    xlsmName = ThisWorkbook.Name
    xlamName = Replace(LCase(xlsmName), ".xlsm", ".xlam")
    xlamPath = GetLocalPath(ThisWorkbook.Path & "/" & xlamName)

    ' --- Confirmation ---
    result = MsgBox("Publication du complément :" & vbCrLf & vbCrLf & _
                    "Source  : " & xlsmName & vbCrLf & _
                    "Sortie  : " & xlamPath & vbCrLf & vbCrLf & _
                    "Le fichier .xlam existant sera remplacé." & vbCrLf & _
                    "Voulez-vous continuer ?", _
                    vbQuestion + vbYesNo, "Publier le complément")

    If result <> vbYes Then Exit Sub

    ' --- Désactiver les alertes pendant la copie ---
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

    ' --- Sauvegarder d'abord le .xlsm ---
    ThisWorkbook.Save

    ' --- Copier en .xlam ---
    ThisWorkbook.SaveCopyAs xlamPath


    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox "Complément publié avec succès !" & vbCrLf & vbCrLf & _
           xlamPath, vbInformation, "Publication terminée"

    Exit Sub

'ErrorHandler:
'    Application.DisplayAlerts = True
'    Application.ScreenUpdating = True
'    If Not wbTarget Is Nothing Then
'        wbTarget.Close False
'    End If
'    MsgBox "Erreur (" & Err.Number & ") : " & Err.description, vbCritical, "PublishToXlam"
    
End Sub

' ==========================================
' Convertit une URL SharePoint/OneDrive
' en chemin local via Environ("USERPROFILE")
' ==========================================
Private Function GetLocalPath(ByVal spPath As String) As String
    Dim localBase  As String
    Dim spBase     As String
    Dim relativePath As String

    ' Cas 1 : chemin déjà local (commence par une lettre de lecteur)
    If Mid(spPath, 2, 2) = ":\" Or Left(spPath, 2) = "\\" Then
        GetLocalPath = spPath
        Exit Function
    End If

    ' Cas 2 : URL SharePoint Alstom
    ' https://alstomgroup.sharepoint.com/sites/Svcs_DIS_Fr_VBN/Shared Documents/...
    ' ? C:\Users\XXXXXXX\Alstom\Svcs_DIS_Fr_VBN - Documents\...

    spBase = "https://alstomgroup.sharepoint.com/sites/Svcs_DIS_Fr_VBN/Shared%20Documents"
    If InStr(spPath, spBase) = 0 Then
        ' Essayer sans encodage %20
        spBase = "https://alstomgroup.sharepoint.com/sites/Svcs_DIS_Fr_VBN/Shared Documents"
    End If

    If InStr(spPath, spBase) > 0 Then
        ' Extraire la partie relative après /Shared Documents/
        relativePath = Mid(spPath, Len(spBase) + 2)  ' +2 pour le slash final

        ' Remplacer les %20 résiduels par des espaces
        relativePath = Replace(relativePath, "%20", " ")

        ' Remplacer les slashs URL par des antislashs Windows
        relativePath = Replace(relativePath, "/", "\")

        ' Construire le chemin local
        localBase = Environ("USERPROFILE") & "\Alstom\Svcs_DIS_Fr_VBN - Documents\"
        GetLocalPath = localBase & relativePath
        Exit Function
    End If

    ' Cas 3 : URL OneDrive générique
    ' https://d.docs.live.net/XXXX/... ? ne s'applique pas ici, mais sécurité
    GetLocalPath = spPath  ' retourne tel quel si non reconnu

End Function

