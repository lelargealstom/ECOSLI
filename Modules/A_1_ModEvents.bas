Option Explicit
' Instance globale du gestionnaire d'evenements
Private g_AppEvents As clsAppEvents

' ==========================================
' Auto_Open - Execute automatiquement au chargement de l'add-in
' ==========================================
Public Sub Auto_Open()

    Call InitializeEventHandler

    ' Binding raccourci Ctrl+W -> Appel_Usf (W_AppelUsf module).
    ' Plus fiable que l attribut VB_Invoke_Func qui pose probleme en
    ' reimport. OnKey s applique a l Application -> persiste tant que
    ' le complement est charge.
    Application.OnKey "^w", "Appel_Usf"

    ' Verifier si une mise a jour est disponible
    Call A_2_ModUpdate.CheckForUpdatesOnStartup
End Sub

' ==========================================
' Auto_Close - Execute automatiquement a la fermeture de l'add-in
' ==========================================
Public Sub Auto_Close()
    ' Libere le raccourci Ctrl+W (sinon il reste lie meme apres unload)
    Application.OnKey "^w"

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
    lastRow = Workbook.Sheets("Macro").Cells(Workbook.Sheets("Macro").rows.count, 2).End(xlUp).Row
    version = Workbook.Sheets("Macro").Cells(lastRow - 1, 2).Value
    
    'verifier si d'autres fichiers sont ouverts dans l'application
    For Each wb In Application.Workbooks
        If wb.Name <> "ECOSLI.xlsm" Then
            MsgBox "Attention, d'autres fichiers sont ouverts dans l'application Excel, veuillez les fermer et redemarrer l'installation", vbCritical
            'enregistrer le fichier
            'w.Save
            Exit Sub
        Else
            Select Case MsgBox("ECOSLI " & version & "  va etre installe.", vbOKCancel + vbInformation)
                Case vbOK
                    'on continue
                Case vbCancel
                    MsgBox "Installation annulee", vbInformation
                    Exit Sub
            End Select
        End If
    Next wb
    
    Set wb = Workbooks("ECOSLI.xlsm")
    
    addinspath = Application.UserLibraryPath
    FullPath = addinspath & "ecosli.xlam"
    
    ' 1)On cherche l'entree dans la collection d'Add Ins
    For Each ai In Application.AddIns
        If StrComp(ai.Name, "ecosli.xlam", vbTextCompare) = 0 Or (Len(ai.FullName) > 0 And StrComp(Dir$(ai.FullName), "ecosli.xlam", vbTextCompare) = 0) Then
            found = True
            Exit For
        End If
    Next ai
    
    ' 2) Si pas trouvee, on l'ajoute au chemin
    If Not found Then
        If Dir$(FullPath) = "" Then
            ' 1) Sauver silencieusement ce classeur en tant que .xlam dans la librairie des AddIns
            Application.DisplayAlerts = False
            ThisWorkbook.SaveAs Filename:=FullPath, FileFormat:=xlOpenXMLAddIn, CreateBackup:=False
            Application.DisplayAlerts = True
        
            ' 2) L'ajouter a la collection des complements Excel (sans copie) puis l'activer
            On Error Resume Next
            Set ai = Application.AddIns.Add(Filename:=FullPath, CopyFile:=False)
            On Error GoTo 0
        
            ' Securite : si AddIns.Add n'a pas renvoye l'objet, on le retrouve par le fichier
            If ai Is Nothing Then
                For Each ai In Application.AddIns
                    If StrComp(Dir$(ai.FullName), "ecosli.xlam", vbTextCompare) = 0 Then Exit For
                Next ai
            End If
    
            If Not ai Is Nothing Then
                ai.Installed = True
            Else
                ' (Optionnel) Lever une erreur claire si quelque chose ne va pas
                Err.Raise vbObjectError + 513, , "Impossible d'enregistrer/activer le complement ecosli.xlam."
            End If
        End If
        
        Else
        ' 3) Si le complement est trouve, on fait l'update
            'a) Si installe, le desinstaller pour liberer le fichier
            If ai.Installed Then ai.Installed = False
            
            'b) Ecraser le suivant en resauvegardant le classeur en .xlam
            ThisWorkbook.SaveAs Filename:=FullPath, FileFormat:=xlOpenXMLAddIn, CreateBackup:=False

            'c) Reactivation du complement
            ai.Installed = True
        
    End If
    
    'ouverture d'un nouveau fichier pour travailler
    Set wb1 = Application.Workbooks.Add
    
    'On copie colle la feuille template ainsi que la feuille de tracabilite
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
    wb1.Sheets("DATA_ASL").Range("Providor[SOUS SYSTEME]").Select
    Selection.ClearContents
    wb1.Sheets("DATA_ASL").Range("Providor[Qte]").Select
    Selection.ClearContents
    wb1.Sheets("DATA_ASL").Range("Providor[MTBF intrinseque RAM Report specifique]").Select
    Selection.ClearContents
    wb1.Sheets("DATA_ASL").Range("Providor[MTBF fonctionnel RAM Report specifique]").Select
    Selection.ClearContents
    wb1.Sheets("DATA_ASL").Range("Providor[DVL HARD specifique]").Select
    Selection.ClearContents
    wb1.Sheets("DATA_ASL").Range("Providor[DVL SOFT specifique]").Select
    Selection.ClearContents
    wb1.Sheets("DATA_ASL").Range("Providor[Cout spare FOURNISSEUR Not ISP]").Select
    Selection.ClearContents
    wb1.Sheets("DATA_ASL").Range("Providor[Lead Time SPARE FOURNISSEUR]").Select
    Selection.ClearContents
    wb1.Sheets("DATA_ASL").Range("Providor[Cout Repair FOURNISSEUR Not ISP]").Select
    Selection.ClearContents
    wb1.Sheets("DATA_ASL").Range("Providor[Lead Time REPAIR FOURNISSEUR]").Select
    Selection.ClearContents
    
    MsgBox "ECOSLI est enregistre en tant qu'Excel Addin et active." & Chr(13) & "   => Vous pouvez utiliser ctrl+W pour ouvrir la fenetre ECOSLI.", vbInformation
    
    Application.DisplayAlerts = False
    'fermeture du fichier d'install (attention, la macro en cours sera stoppee)
    Workbooks("ECOSLI.xlsm").Close
    MsgBox "La feuille qui vient de s'ouvrir est le template a utiliser pour la generation. Vous pouvez le sauver dans votre repertoire projet"
    
    wb1.Activate
    TurnOnFunctionality
    
End Sub
'Procedure de verification de la version du complement ECO SLI
Public Sub VerifierVersionComplement()
    
    Dim cheminTxt As String
    Dim cheminMacro As String
    Dim fichier As Object
    Dim versionActuelle, versionDisponible As String
    
    versionActuelle = "1.1" ' a adapter selon ta version actuelle
    
    cheminTxt = Environ$("UserProfile") & "\Alstom\Svcs_DIS_Fr_VBN - Documents\01_Metiers\01_ISR\007 - Outils metier ISR\01. BDD SLI\11_Livraison de l'OUTIL excel BDD\VersionComplement.txt"
    cheminMacro = Environ$("UserProfile") & "\Alstom\Svcs_DIS_Fr_VBN - Documents\01_Metiers\01_ISR\007 - Outils metier ISR"
    On Error GoTo ErreurLecture
    
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
        MsgBox "Une nouvelle version du complement est disponible : " & versionDisponible & vbCrLf & "Veuillez telecharger le dernier complement sur le chemin suivant." & vbCrLf & cheminTxt, vbInformation, "Mise a jour disponible"
    End If

    Exit Sub

ErreurLecture:
    MsgBox "Impossible de verifier la version du complement." & vbCrLf & "Verifiez que le fichier 'VersionComplement.txt' est bien synchronise via SharePoint.", vbExclamation
End Sub




