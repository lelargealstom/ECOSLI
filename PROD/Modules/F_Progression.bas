
'--------------------------------------------------------------------------------------------------------
'
'                                       Generateur Livrables SLI
'                                       Module F_Progression
'                                       Developpement : LE LARGE Thomas (LGM/Alstom)
'                                       UX progression : barre du UserForm + StatusBar Excel +
'                                       curseur sablier + DTR en cours + temps ecoule / restant
'                                       + UserForm "toujours devant" via API Windows.
'                                       A appeler depuis ATM V1/V2, EMMST, Providor, etc.
'                                       Juin 2026
'
'--------------------------------------------------------------------------------------------------------
Option Explicit

' =========================================================================
' API Windows pour forcer le UserForm en TopMost ("always on top") pendant
' la generation. Disponible nativement sur toutes les machines Windows
' (user32.dll). Pas de droits admin requis.
' =========================================================================
#If VBA7 Then
    Private Declare PtrSafe Function FindWindowA Lib "user32" ( _
        ByVal lpClassName As String, _
        ByVal lpWindowName As String) As LongPtr
    Private Declare PtrSafe Function SetWindowPos Lib "user32" ( _
        ByVal hWnd As LongPtr, _
        ByVal hWndInsertAfter As LongPtr, _
        ByVal x As Long, ByVal y As Long, _
        ByVal cx As Long, ByVal cy As Long, _
        ByVal wFlags As Long) As Long
#Else
    Private Declare Function FindWindowA Lib "user32" ( _
        ByVal lpClassName As String, _
        ByVal lpWindowName As String) As Long
    Private Declare Function SetWindowPos Lib "user32" ( _
        ByVal hWnd As Long, _
        ByVal hWndInsertAfter As Long, _
        ByVal x As Long, ByVal y As Long, _
        ByVal cx As Long, ByVal cy As Long, _
        ByVal wFlags As Long) As Long
#End If

Private Const HWND_TOPMOST    As Long = -1
Private Const HWND_NOTOPMOST  As Long = -2
Private Const SWP_NOMOVE      As Long = &H2
Private Const SWP_NOSIZE      As Long = &H1
Private Const SWP_NOACTIVATE  As Long = &H10

' Etat courant (Private, scope module)
Private prog_total As Long
Private prog_titre As String
Private prog_tDebut As Double      ' Timer() au demarrage
Private prog_actif As Boolean
Private prog_dernierTaux As Long   ' eviter de re-peindre pour rien
#If VBA7 Then
Private prog_hWndUserForm As LongPtr   ' handle Windows du UserForm, recupere 1x dans Progress_Demarrer
#Else
Private prog_hWndUserForm As Long
#End If

' =========================================================================
' Progress_Demarrer
' =========================================================================
' A appeler en debut de generation longue.
'   - lTotal : nombre total d'items a traiter (ex. UBound(tDTR))
'   - sTitre : libelle court affiche dans la StatusBar (ex. "ATM V2", "ATM V1", "EMMST")
'
' Effets :
'   - Curseur Excel en sablier
'   - StatusBar = "<titre> - demarrage..."
'   - Barre + label du UserForm rendus visibles, remis a 0%
'   - Demarre le chrono interne (Timer)
Public Sub Progress_Demarrer(ByVal lTotal As Long, ByVal sTitre As String)
    prog_total = lTotal
    prog_titre = sTitre
    prog_tDebut = Timer()
    prog_actif = True
    prog_dernierTaux = -1

    Application.Cursor = xlWait
    Application.StatusBar = sTitre & " - demarrage..."

    ' Rendre la barre du UserForm visible et la remettre a zero.
    ' On Error Resume Next pour tolerer le cas ou le UserForm n'est
    ' pas instancie (procedure lancee depuis VBE en standalone).
    On Error Resume Next
    UserForm.barre.Width = 0
    UserForm.barre.Visible = True
    UserForm.barretexte.Caption = "0%"
    UserForm.barretexte.Visible = True
    UserForm.barreimage.Visible = True
    UserForm.barrearriere.Visible = True
    On Error GoTo 0

    ' Recuperer le handle Windows du UserForm UNE FOIS (avant que le
    ' Caption ne soit modifie a chaque tick par Progress_Etape).
    ' Classe Windows du UserForm VBA : "ThunderDFrame" (Office 2007+),
    ' fallback "ThunderXFrame" pour versions plus anciennes.
    On Error Resume Next
    prog_hWndUserForm = FindWindowA("ThunderDFrame", UserForm.Caption)
    If prog_hWndUserForm = 0 Then
        prog_hWndUserForm = FindWindowA("ThunderXFrame", UserForm.Caption)
    End If
    On Error GoTo 0

    ' Activer mode TopMost ("always on top") pendant toute la generation.
    Call PinUserFormTopMost(True)

    DoEvents
End Sub

' =========================================================================
' Progress_Etape
' =========================================================================
' A appeler avant le traitement de chaque item.
'   - lCourant : numero de l'item en cours (1-based)
'   - sEtape   : optionnel, libelle d'etape ("Copie template", "Sauvegarde"...)
'   - sIdItem  : optionnel, identifiant lisible (ex. DTR, nom de fichier)
'
' Effets :
'   - Met a jour la barre + le label % du UserForm
'   - Met a jour la StatusBar Excel avec : titre, ratio, DTR, etape,
'     temps ecoule et temps restant estime
'   - DoEvents pour laisser Excel re-peindre
Public Sub Progress_Etape(ByVal lCourant As Long, _
                          Optional ByVal sEtape As String = "", _
                          Optional ByVal sIdItem As String = "")
    If Not prog_actif Then Exit Sub
    If prog_total <= 0 Then Exit Sub

    Dim taux As Long
    taux = CLng(lCourant / prog_total * 100)
    If taux < 0 Then taux = 0
    If taux > 100 Then taux = 100

    ' --- Temps ecoule et restant estime ---------------------
    Dim tEcoule As Double, tRestant As Double
    tEcoule = Timer() - prog_tDebut
    If lCourant > 0 And lCourant <= prog_total Then
        tRestant = tEcoule / lCourant * (prog_total - lCourant)
    Else
        tRestant = 0
    End If

    ' --- StatusBar Excel (visible si barre d'etat Excel activee) ----
    Dim sBar As String
    sBar = prog_titre & " - " & lCourant & "/" & prog_total & " (" & taux & "%)"
    If LenB(sIdItem) > 0 Then sBar = sBar & " - " & sIdItem
    If LenB(sEtape) > 0 Then sBar = sBar & " - " & sEtape
    sBar = sBar & " - ecoule " & FormaterDuree(tEcoule)
    If tRestant > 0 Then sBar = sBar & ", restant ~" & FormaterDuree(tRestant)
    Application.StatusBar = sBar

    ' --- Caption du UserForm (toujours visible, plus fiable que StatusBar) ----
    On Error Resume Next
    UserForm.Caption = "ECOSLI - " & prog_titre & " - " & lCourant & "/" & prog_total
    If LenB(sIdItem) > 0 Then UserForm.Caption = UserForm.Caption & " : " & sIdItem
    On Error GoTo 0

    ' --- Barre visuelle du UserForm -------------------------
    ' Ne re-peindre que si le taux a change (eviter le flicker).
    If taux <> prog_dernierTaux Then
        prog_dernierTaux = taux
        On Error Resume Next
        UserForm.barre.Width = lCourant / prog_total * UserForm.barreimage.Width
        UserForm.barretexte.Caption = taux & "%"
        UserForm.Repaint                       ' force le UserForm a se redessiner
        On Error GoTo 0
    End If
    ' Note : pas de UserForm.ZOrder (n'existe pas comme methode UserForm).
    ' Pour forcer le UserForm devant la fenetre Excel, il faudrait une API
    ' Windows (SetWindowPos) - report en Phase B si necessaire.

    ' --- Liberer le presse-papier (anti-erreur Clipboard) -----------
    ' Les Range.Copy et Sheets.Copy dans la boucle remplissent le
    ' presse-papier Excel qui finit par saturer ("Erreur du
    ' Presse-Papiers"). On le vide a chaque etape.
    On Error Resume Next
    Application.CutCopyMode = False
    On Error GoTo 0

    DoEvents
End Sub

' =========================================================================
' Progress_Terminer
' =========================================================================
' A appeler en fin de generation (avant MsgBox de confirmation).
' Restaure curseur, StatusBar, masque la barre du UserForm, trace
' la duree totale dans Immediate.
Public Sub Progress_Terminer()
    If Not prog_actif Then Exit Sub

    Dim tTotal As Double
    tTotal = Timer() - prog_tDebut
    Debug.Print prog_titre & " : termine en " & FormaterDuree(tTotal) & _
                " (" & prog_total & " items)"

    prog_actif = False

    ' Desactiver le mode TopMost AVANT toute autre operation (le UserForm
    ' redevient une fenetre normale et peut etre cachee par d'autres).
    Call PinUserFormTopMost(False)
    prog_hWndUserForm = 0

    Application.Cursor = xlDefault
    Application.StatusBar = False   ' restaure l'affichage natif Excel

    On Error Resume Next
    UserForm.barre.Width = 0
    UserForm.barretexte.Caption = ""
    UserForm.barre.Visible = False
    UserForm.barretexte.Visible = False
    UserForm.barreimage.Visible = False
    UserForm.barrearriere.Visible = False
    UserForm.Caption = "ECOSLI"     ' restaure le titre du UserForm
    Application.CutCopyMode = False
    On Error GoTo 0
End Sub

' =========================================================================
' PinUserFormTopMost
' =========================================================================
' Active (True) ou desactive (False) le mode "always on top" du UserForm.
' Utilise prog_hWndUserForm recupere une seule fois dans Progress_Demarrer.
' Si handle = 0 (UserForm pas trouve), no-op silencieux.
Private Sub PinUserFormTopMost(ByVal bActif As Boolean)
#If VBA7 Then
    Dim hAfter As LongPtr
#Else
    Dim hAfter As Long
#End If

    If prog_hWndUserForm = 0 Then Exit Sub

    If bActif Then
        hAfter = HWND_TOPMOST
    Else
        hAfter = HWND_NOTOPMOST
    End If

    ' SWP_NOMOVE + SWP_NOSIZE + SWP_NOACTIVATE : on ne change que l'ordre Z,
    ' pas la position, la taille ou le focus (pas de clignotement visuel).
    SetWindowPos prog_hWndUserForm, hAfter, 0, 0, 0, 0, _
                 SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE
End Sub

' =========================================================================
' FormaterDuree
' =========================================================================
' Formate une duree en secondes en "Xm Ys" ou "Ys".
Private Function FormaterDuree(ByVal secs As Double) As String
    Dim m As Long, s As Long
    If secs < 0 Then secs = 0
    m = Int(secs / 60)
    s = Int(secs - m * 60)
    If m > 0 Then
        FormaterDuree = m & "m " & s & "s"
    Else
        FormaterDuree = s & "s"
    End If
End Function

