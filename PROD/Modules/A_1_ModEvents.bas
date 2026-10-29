Option Explicit
' Instance globale du gestionnaire d'evenements
Private g_AppEvents As clsAppEvents

' ==========================================
' Auto_Open - Execute automatiquement au chargement de l'add-in
' ==========================================
Public Sub Auto_Open()

    Call InitializeEventHandler

    ' Verifier si une mise a jour est disponible
    Call A_2_ModUpdate.CheckForUpdatesOnStartup
End Sub

' ==========================================
' Auto_Close - Execute automatiquement a la fermeture de l'add-in
' ==========================================
Public Sub Auto_Close()
    Call TerminateEventHandler
End Sub

' ==========================================
' Initialiser le gestionnaire d'evenements
' ==========================================
Public Sub InitializeEventHandler()
    If g_AppEvents Is Nothing Then
        Set g_AppEvents = New clsAppEvents
        g_AppEvents.Initialize
    End If
End Sub

' ==========================================
' Arreter le gestionnaire d'evenements
' A appeler a la fermeture de l'add-in
' ==========================================
Public Sub TerminateEventHandler()
    If Not g_AppEvents Is Nothing Then
        g_AppEvents.Terminate
        Set g_AppEvents = Nothing
    End If
End Sub

' ==========================================
' Verifier si le gestionnaire est actif
' ==========================================
Public Function IsEventHandlerActive() As Boolean
    IsEventHandlerActive = Not (g_AppEvents Is Nothing)
End Function
Public Sub Install()
     
    Dim FullPath, addinspath As String
    Dim version As String
    Dim lastRow As Integer
    Dim Workbook As Workbook
    
    TurnOffFunctionality
    
    Set Workbook = ThisWorkbook
    lastRow = Workbook.Sheets("Macro").Cells(Workbook.Sheets("Macro").rows.count, 2).End(xlUp).row
    version = Workbook.Sheets("Macro").Cells(lastRow - 1, 2).Value
    
    'vérifier si d'autres fichiers sont ouverts dans l'application
    For Each wb In Application.Workbooks
        If wb.Name <> "ECOSLI.xlsm" Then
            MsgBox "Attention, d’autres fichiers sont ouverts dans l’application Excel, veuillez les fermer et redémarrer l’installation", vbCritical
            'enregistrer le fichier
            'w.Save
            Exit Sub
        Else
            Select Case MsgBox("ECOSLI " & version & "  va etre installé.", vbOKCancel + vbInformation)
                Case vbOK
                    'on continue
                Case vbCancel
                    MsgBox "Installation annulée", vbInformation
                    Exit Sub
            End Select
        End If
    Next wb
    
    Set wb = Workbooks("ECOSLI.xlsm")
    
    addinspath = Application.UserLibraryPath
    FullPath = addinspath & "ecosli.xlam"
    
    ' 1)On cherche l'entrée dans la collection d'Add Ins
    For Each ai In Application.AddIns
        If StrComp(ai.Name, "ecosli.xlam", vbTextCompare) = 0 Or _
            (Len(ai.FullName) > 0 And StrComp(Dir$(ai.FullName), "ecosli.xlam", vbTextCompare) = 0) Then
            found = True
            Exit For
        End If
    Next ai
    
    ' 2) Si pas trouvée, on l'ajoute au chemin
    If Not found Then
        If Dir$(FullPath) = "" Then
            ' 1) Sauver silencieusement ce classeur en tant que .xlam dans la librairie des AddIns
            Application.DisplayAlerts = False
            ThisWorkbook.SaveAs Filename:=FullPath, FileFormat:=xlOpenXMLAddIn, CreateBackup:=False
            Application.DisplayAlerts = True
        
            ' 2) L’ajouter à la collection des compléments Excel (sans copie) puis l’activer
            On Error Resume Next
            Set ai = Application.AddIns.Add(Filename:=FullPath, CopyFile:=False)
            On Error GoTo 0
        
            ' Sécurité : si AddIns.Add n’a pas renvoyé l’objet, on le retrouve par le fichier
            If ai Is Nothing Then
                For Each ai In Application.AddIns
                    If StrComp(Dir$(ai.FullName), "ecosli.xlam", vbTextCompare) = 0 Then Exit For
                Next ai
            End If
    
            If Not ai Is Nothing Then
                ai.Installed = True
            Else
                ' (Optionnel) Lever une erreur claire si quelque chose ne va pas
                Err.Raise vbObjectError + 513, , "Impossible d’enregistrer/activer le complément ecosli.xlam."
            End If
        End If
        
        Else
        ' 3) Si le complement est trouvé, on fait l'update
            'a) Si installé, le désinstaller pour libérer le fichier
            If ai.Installed Then ai.Installed = False
            
            'b) Ecraser le suivant en resauvegardant le classeur en .xlam
            ThisWorkbook.SaveAs Filename:=FullPath, FileFormat:=xlOpenXMLAddIn, CreateBackup:=False

            'c) Reactivation du complément
            ai.Installed = True
        
    End If
    
    'ouverture d'un nouveau fichier pour travailler
    Set wb1 = Application.Workbooks.Add
    
    'On copie colle la feuille template ainsi que la feuille de tracabilité
    wb.Sheets("DATA_ASL").Copy Before:=wb1.Sheets(1)
    wb.Sheets("TRACA_SHEET").Copy Before:=wb1.Sheets(1)
    wb.Sheets("LISTES_DATA_CHOIX").Copy Before:=wb1.Sheets(1)
    wb1.Sheets("TRACA_SHEET").Range("Traca_tab").ListObject.DataBodyRange.Delete
    wb1.Sheets("LISTES_DATA_CHOIX").Visible = 2
    wb1.Sheets("Feuil1").Delete
    'On nettoie la PBS actuel
    wb1.Activate
    wb1.Sheets("DATA_ASL").Activate
    wb1.Sheets("DATA_ASL").Range("Providor[DESIGNATION (issue de PBS)]").Select
    Selection.ClearContents
    wb1.Sheets("DATA_ASL").Range("Providor[DTR Reference (issue de PBS)]").Select
    Selection.ClearContents
    wb1.Sheets("DATA_ASL").Range("Providor[SOUS SYSTÈME (copiés de la PBS et classifiés selon génériques)]").Select
    Selection.ClearContents
    wb1.Sheets("DATA_ASL").Range("Providor[Qte]").Select
    Selection.ClearContents
    wb1.Sheets("DATA_ASL").Range("Providor[MTBF intrinsèque (RAM report specifique projet (CGR))]").Select
    Selection.ClearContents
    wb1.Sheets("DATA_ASL").Range("Providor[MTBF fonctionnel (RAM report specifique projet (CGR))]").Select
    Selection.ClearContents
    wb1.Sheets("DATA_ASL").Range("Providor[DUREE DE VIE LIMITEE (RAM report specifique projet (CGR)) HARDWARE]").Select
    Selection.ClearContents
    wb1.Sheets("DATA_ASL").Range("Providor[DUREE DE VIE LIMITEE (RAM report specifique projet (CGR)) SOFTWARE]").Select
    Selection.ClearContents
    wb1.Sheets("DATA_ASL").Range("Providor[Cout spare FOURNISSEUR (Not ISP)]").Select
    Selection.ClearContents
    wb1.Sheets("DATA_ASL").Range("Providor[Lead Time SPARE FOURNISSEUR]").Select
    Selection.ClearContents
    wb1.Sheets("DATA_ASL").Range("Providor[Cout Repair FOURNISSEUR (Not ISP)]").Select
    Selection.ClearContents
    wb1.Sheets("DATA_ASL").Range("Providor[Lead Time REPAIR FOURNISSEUR]").Select
    Selection.ClearContents
    
    MsgBox "ECOSLI est enregistré en tant qu'Excel Addin et activé." & Chr(13) & "   => Vous pouvez utiliser ctrl+W pour ouvrir la fenêtre ECOSLI.", vbInformation
    
    Application.DisplayAlerts = False
    'fermeture du fichier d'install (attention, la macro en cours sera stoppée)
    Workbooks("ECOSLI.xlsm").Close
    MsgBox "La feuille qui vient de s'ouvrir est le template a utiliser pour la generation. Vous pouvez le sauver dans votre répertoire projet"
    
    wb1.Activate
    TurnOnFunctionality
    
End Sub
'Procédure de vérification de la version du complément ECO SLI
Public Sub VerifierVersionComplement()
    
    Dim cheminTxt As String
    Dim cheminMacro As String
    Dim fichier As Object
    Dim versionActuelle, versionDisponible As String
    
    versionActuelle = "1.1" ' à adapter selon ta version actuelle
    
    cheminTxt = Environ$("UserProfile") & "\Alstom\DIS Service VBN Pole ISR - Outils Métier ISR\01. BDD SLI\11_Livraison de l'OUTIL excel BDD\VersionComplement.txt"
    cheminMacro = Environ$("UserProfile") & "\Alstom\DIS Service VBN Pole ISR - Outils Métier ISR"
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fichier = fso.OpenTextFile(cheminTxt, 1)
    
    Do While Not fichier.AtEndOfStream
        ligne = fichier.ReadLine
        If InStr(ligne, "Version") > 0 Then
            versionDisponible = Trim(Split(ligne, "=")(1))
            Exit Do
        End If
    Loop
    
    fichier.Close
    
    If versionDisponible <> versionActuelle And versionDisponible <> "" Then
        MsgBox "Une nouvelle version du complément est disponible : " & versionDisponible & vbCrLf & _
               "Veuillez télécharger le dernier complément sur le chemin suivant." & vbCrLf & _
               cheminTxt, vbInformation, "Mise à jour disponible"
    End If
End Sub




