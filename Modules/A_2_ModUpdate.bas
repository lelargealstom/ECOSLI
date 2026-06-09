Attribute VB_Name = "A_2_ModUpdate"
Option Explicit

Private Const REG_KEY As String = "HKEY_CURRENT_USER\Software\AlstomSLI\AddinPath"

' Fonction pour le ruban (sans parametre)
Public Sub Update_AddIn()
    Call Update_AddIn_Entry(False)
End Sub

' Fonction principale
Private Sub Update_AddIn_Entry(Optional ByVal skipConfirm As Boolean = False)
    Dim sourcePath As String
    Dim destPath As String
    Dim tempPath As String
    Dim batPath As String
    Dim result As VbMsgBoxResult
    Dim wsh As Object

    On Error GoTo ErrorHandler

    Set wsh = CreateObject("WScript.Shell")

    ' Chemin source SharePoint (fixe)
    sourcePath = Environ("USERPROFILE") & "\Alstom\Svcs_DIS_Fr_VBN - Documents\01_Metiers\01_ISR\007 - Outils métier ISR\01. BDD SLI\11_Livraison de l'OUTIL excel BDD\ecosli.xlam"

    ' Verifier que le fichier source existe
    If Dir(sourcePath) = "" Then
        MsgBox "Fichier source introuvable sur SharePoint." & vbCrLf & vbCrLf & _
               "Verifiez que le dossier est synchronise.", vbExclamation, "Erreur"
        Set wsh = Nothing
        Exit Sub
    End If

    ' Recuperer le chemin de destination depuis le registre
    On Error Resume Next
    destPath = wsh.RegRead(REG_KEY)
    On Error GoTo ErrorHandler

    ' Si pas de chemin enregistre, demander a l'utilisateur (une seule fois)
    If destPath = "" Or Dir(destPath) = "" Then
        MsgBox "Premiere utilisation : vous devez indiquer ou se trouve votre add-in actuel." & vbCrLf & vbCrLf & _
               "(Cette question ne sera posee qu'une seule fois)", vbInformation

        destPath = SelectAddinPath()
        If destPath = "" Then
            Set wsh = Nothing
            Exit Sub
        End If

        ' Sauvegarder dans le registre
        wsh.RegWrite REG_KEY, destPath, "REG_SZ"
    End If

    ' Confirmation (sauf si deja confirmee)
    If Not skipConfirm Then
        Dim sourceDate As Date, destDate As Date
        sourceDate = FileDateTime(sourcePath)
        destDate = FileDateTime(destPath)

        result = MsgBox("Mise a jour de l'add-in :" & vbCrLf & vbCrLf & _
                       "Version actuelle : " & Format(destDate, "DD/MM/YYYY HH:MM") & vbCrLf & _
                       "Nouvelle version : " & Format(sourceDate, "DD/MM/YYYY HH:MM") & vbCrLf & vbCrLf & _
                       "Excel va se fermer et se rouvrir automatiquement.", _
                       vbQuestion + vbYesNo, "Mise a jour")

        If result <> vbYes Then
            Set wsh = Nothing
            Exit Sub
        End If
    End If

    ' Copier vers TEMP
    tempPath = Environ("TEMP") & "\new_addin.xlam"
    wsh.Run "cmd /c copy """ & sourcePath & """ """ & tempPath & """ /Y", 0, True

    ' Verifier copie
    If Dir(tempPath) = "" Then
        MsgBox "Erreur: La copie a echoue.", vbCritical
        Set wsh = Nothing
        Exit Sub
    End If

    ' Creer le script batch
    batPath = Environ("TEMP") & "\update_addin.bat"
    CreateUpdateBatch batPath, tempPath, destPath

    ' Lancer le batch (attend qu'Excel se ferme)
    wsh.Run """" & batPath & """", 1, False

    Set wsh = Nothing

    ' Fermer Excel
    Application.Quit

    Exit Sub

ErrorHandler:
    MsgBox "Erreur: " & Err.description, vbCritical
    If Not wsh Is Nothing Then Set wsh = Nothing
End Sub

' ==========================================
' Detecter automatiquement le chemin de l'add-in
' Cherche dans les add-ins charges celui qui contient "Obsolescence"
' Fallback: dossier AddIns par defaut
' ==========================================
Private Function DetectAddinPath() As String
    Dim addin As Workbook
    Dim addinName As String

    On Error Resume Next
    DetectAddinPath = ""

    ' Methode 1 : Chercher dans les classeurs ouverts (les add-ins sont des workbooks)
    For Each addin In Application.Workbooks
        If LCase(Right(addin.Name, 5)) = ".xlam" Then
            If InStr(1, addin.Name, "ecosli", vbTextCompare) > 0 Then
                DetectAddinPath = addin.FullName
                Exit Function
            End If
        End If
    Next addin

    ' Methode 2 : Chercher dans les add-ins installes
    Dim ai As addin
    For Each ai In Application.AddIns
        If ai.Installed Then
            If InStr(1, ai.Name, "ecosli", vbTextCompare) > 0 Then
                If Dir(ai.FullName) <> "" Then
                    DetectAddinPath = ai.FullName
                    Exit Function
                End If
            End If
        End If
    Next ai

    ' Methode 3 : Chercher dans le dossier AddIns par defaut
    Dim defaultPath As String
    defaultPath = Environ("APPDATA") & "\Microsoft\AddIns\ecosli.xlam"
    If Dir(defaultPath) <> "" Then
        DetectAddinPath = defaultPath
        Exit Function
    End If

    On Error GoTo 0
End Function

Private Function SelectAddinPath() As String
    Dim initialPath As String
    Dim result As Variant

    SelectAddinPath = ""

    ' Ouvrir dans le dossier AddIns par defaut
    initialPath = Environ("APPDATA") & "\Microsoft\AddIns\"
    On Error Resume Next
    ChDrive Left(initialPath, 1)
    ChDir initialPath
    On Error GoTo 0

    result = Application.GetOpenFilename( _
        FileFilter:="Add-in Excel (*.xlam), *.xlam", _
        Title:="Selectionner votre add-in ACTUEL (celui a remplacer)")

    If VarType(result) = vbBoolean Then
        SelectAddinPath = ""
    Else
        SelectAddinPath = CStr(result)
    End If
End Function

Private Sub CreateUpdateBatch(ByVal batPath As String, ByVal tempPath As String, ByVal destPath As String)
    Dim ps1Path As String
    Dim f As Integer
    Dim stream As Object

    ' Utiliser PowerShell avec encodage UTF-8
    ps1Path = Environ("TEMP") & "\update_addin.ps1"

    ' Ecrire le script PowerShell en UTF-8
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2 ' adTypeText
    stream.Charset = "UTF-8"
    stream.Open

    stream.WriteText "$host.UI.RawUI.WindowTitle = 'Mise a jour Add-in'" & vbCrLf
    stream.WriteText "Write-Host 'Mise a jour de l''add-in en cours...'" & vbCrLf
    stream.WriteText "Write-Host 'Attente de la fermeture d''Excel...'" & vbCrLf
    stream.WriteText "" & vbCrLf
    stream.WriteText "while (Get-Process -Name 'EXCEL' -ErrorAction SilentlyContinue) {" & vbCrLf
    stream.WriteText "    Start-Sleep -Seconds 2" & vbCrLf
    stream.WriteText "}" & vbCrLf
    stream.WriteText "" & vbCrLf
    stream.WriteText "Write-Host 'Excel ferme. Copie du fichier...'" & vbCrLf
    stream.WriteText "Start-Sleep -Seconds 1" & vbCrLf
    stream.WriteText "" & vbCrLf
    stream.WriteText "$source = '" & tempPath & "'" & vbCrLf
    stream.WriteText "$dest = '" & destPath & "'" & vbCrLf
    stream.WriteText "" & vbCrLf
    stream.WriteText "try {" & vbCrLf
    stream.WriteText "    Copy-Item -Path $source -Destination $dest -Force" & vbCrLf
    stream.WriteText "    Write-Host 'Mise a jour reussie !' -ForegroundColor Green" & vbCrLf
    stream.WriteText "    Write-Host 'Redemarrage d''Excel...'" & vbCrLf
    stream.WriteText "    Start-Process '" & Application.Path & "\EXCEL.EXE'" & vbCrLf
    stream.WriteText "    Remove-Item $source -ErrorAction SilentlyContinue" & vbCrLf
    stream.WriteText "} catch {" & vbCrLf
    stream.WriteText "    Write-Host 'ERREUR lors de la copie !' -ForegroundColor Red" & vbCrLf
    stream.WriteText "    Write-Host $_.Exception.Message" & vbCrLf
    stream.WriteText "    Read-Host 'Appuyez sur Entree pour fermer'" & vbCrLf
    stream.WriteText "}" & vbCrLf

    stream.SaveToFile ps1Path, 2 ' adSaveCreateOverWrite
    stream.Close
    Set stream = Nothing

    ' Creer le batch qui lance PowerShell (pour contourner les restrictions)
    f = FreeFile
    Open batPath For Output As #f
    Print #f, "@echo off"
    Print #f, "powershell -ExecutionPolicy Bypass -File """ & ps1Path & """"
    Print #f, "del """ & ps1Path & """ 2>NUL"
    Print #f, "del ""%~f0"""
    Close #f
End Sub

Public Sub ResetAddinPath()
    Dim wsh As Object
    On Error Resume Next
    Set wsh = CreateObject("WScript.Shell")
    wsh.RegDelete REG_KEY
    Set wsh = Nothing
    MsgBox "Chemin reinitialise. La prochaine mise a jour vous demandera de reselectionner l'add-in.", vbInformation
End Sub

Public Sub CheckForUpdatesOnStartup()
    Dim sourcePath As String
    Dim destPath As String
    Dim sourceDate As Date
    Dim destDate As Date
    Dim wsh As Object
    Dim result As VbMsgBoxResult

    On Error Resume Next

    Set wsh = CreateObject("WScript.Shell")

    ' Recuperer le chemin de l'add-in depuis le registre
    destPath = wsh.RegRead(REG_KEY)

    ' Si pas configure, detecter automatiquement le chemin de l'add-in
    If destPath = "" Or Dir(destPath) = "" Then
        destPath = DetectAddinPath()
        If destPath <> "" Then
            ' Sauvegarder dans le registre pour la prochaine fois
            wsh.RegWrite REG_KEY, destPath, "REG_SZ"
        Else
            Set wsh = Nothing
            Exit Sub
        End If
    End If

    ' Chemin du fichier source sur SharePoint
    sourcePath = Environ("USERPROFILE") & "\Alstom\Svcs_DIS_Fr_VBN - Documents\01_Metiers\01_ISR\007 - Outils métier ISR\01. BDD SLI\11_Livraison de l'OUTIL excel BDD\ecosli.xlam"

    ' Verifier que le fichier source existe
    If Dir(sourcePath) = "" Then
        Set wsh = Nothing
        Exit Sub
    End If

    ' Comparer les dates
    sourceDate = FileDateTime(sourcePath)
    destDate = FileDateTime(destPath)

    Set wsh = Nothing

    ' Si la source est plus recente (marge de 1 minute)
    If sourceDate > DateAdd("n", 1, destDate) Then
        result = MsgBox("Une nouvelle version de l'add-in est disponible !" & vbCrLf & vbCrLf & _
                       "Version installee : " & Format(destDate, "DD/MM/YYYY HH:MM") & vbCrLf & _
                       "Nouvelle version : " & Format(sourceDate, "DD/MM/YYYY HH:MM") & vbCrLf & vbCrLf & _
                       "Voulez-vous mettre a jour maintenant ?", _
                       vbQuestion + vbYesNo, "Mise a jour disponible")

        If result = vbYes Then
            Call Update_AddIn_Entry(True)  ' True = skip confirmation (deja faite)
        End If
    End If

    On Error GoTo 0
End Sub




