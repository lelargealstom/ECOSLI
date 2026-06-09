Attribute VB_Name = "F_Progression"
'--------------------------------------------------------------------------------------------------------
'
'                                       Generateur Livrables SLI
'                                       Module F_Progression
'                                       Developpement : LE LARGE Thomas (LGM/Alstom)
'                                       UX progression : barre du UserForm + StatusBar Excel +
'                                       curseur sablier + DTR en cours + temps ecoule / restant.
'                                       A appeler depuis ATM V1/V2, EMMST, Providor, etc.
'                                       Juin 2026
'
'--------------------------------------------------------------------------------------------------------
Option Explicit

' Etat courant (Private, scope module)
Private prog_total As Long
Private prog_titre As String
Private prog_tDebut As Double      ' Timer() au demarrage
Private prog_actif As Boolean
Private prog_dernierTaux As Long   ' eviter de re-peindre pour rien

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
