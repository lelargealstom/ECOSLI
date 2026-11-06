' =========================================================================
' MODULE: X_Test_OuvertureBdDMasquee
' =========================================================================
' Auteur : Thomas LE LARGE - 2026-11-06
' Perimetre : benchmark plan B pour contourner la troncature 255 chars d'ACE
'             sur les commentaires EMMST, apres echec de la Piste 1
'             (cf. X_Test_TypeGuessRows + CHANGELOG 2026-11-06).
'
' But : mesurer le cout reel (en secondes) d'ouverture de la BdD TRAVAIL
'       en Excel natif, invisible, ReadOnly, lecture des 10 colonnes
'       commentaires en bloc, puis fermeture. Deux scenarios :
'       A) ouverture dans l'instance Excel courante (ScreenUpdating OFF).
'       B) ouverture dans une instance Excel separee (CreateObject).
'
' Usage : placer le curseur dans Test_OuvertureBdDMasquee_Tous puis F5.
'         Resultats dans la fenetre Immediate (Ctrl+G).
'
' NB : ne modifie jamais la BdD (ReadOnly + SaveChanges:=False a la fermeture).
' =========================================================================

Option Explicit

Public Sub Test_OuvertureBdDMasquee_Tous()
' Lance les 2 scenarios (instance courante puis instance separee) et
' imprime pour chacun : temps d'ouverture, temps de lecture, temps de
' fermeture, longueurs max par colonne. Verdict d'arbitrage objectif
' entre Plan B1 instance courante vs Plan B1 instance separee.

    Debug.Print String(72, "=")
    Debug.Print "Bench ouverture BdD TRAVAIL masquee - " & Format(Now, "yyyy-mm-dd hh:nn:ss")
    Debug.Print String(72, "=")

    Debug.Print
    Debug.Print "### Scenario A : instance Excel courante, masquee ###"
    Call BenchOuvertureBdD(creerInstanceSeparee:=False)

    Debug.Print
    Debug.Print "### Scenario B : instance Excel separee (CreateObject) ###"
    Call BenchOuvertureBdD(creerInstanceSeparee:=True)

    Debug.Print
    Debug.Print String(72, "=")
    Debug.Print "Verdict :"
    Debug.Print "  * si Total < 2 s : ouverture BdD acceptable, on adopte Plan B1."
    Debug.Print "  * si 2 s < Total < 5 s : negociable selon frequence d'usage."
    Debug.Print "  * si Total > 5 s : envisager Plan B2 (export CSV externe)."
    Debug.Print "  * comparer A vs B : si ecart < 0.5 s, A suffit (plus simple)."
    Debug.Print String(72, "=")

    MsgBox "Bench termine. Ouvre la fenetre Immediate (Ctrl+G) pour les temps.", _
           vbInformation, "Bench ouverture BdD masquee"

End Sub

Private Sub BenchOuvertureBdD(ByVal creerInstanceSeparee As Boolean)
' Chronometre l'ouverture, la lecture bloc des 10 colonnes commentaires
' + DTR, puis la fermeture. Imprime les longueurs max observees pour
' validation fonctionnelle (doivent depasser 255 si la BdD contient
' reellement des commentaires longs - test de la fin de la troncature).

    Dim t0 As Double
    Dim t1 As Double
    Dim t2 As Double
    Dim t3 As Double
    Dim appCible As Excel.Application
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim chemin As String

    Dim oldScreen As Boolean
    Dim oldEvents As Boolean
    Dim oldCalc As XlCalculation
    Dim oldAlerts As Boolean

    ' Sauvegarde etat Application courante (toujours, meme si on utilise
    ' une instance separee : ScreenUpdating OFF reduit le bruit UI).
    oldScreen = Application.ScreenUpdating
    oldEvents = Application.EnableEvents
    oldCalc = Application.Calculation
    oldAlerts = Application.DisplayAlerts

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False

    t0 = Timer

    If creerInstanceSeparee Then
        Set appCible = CreateObject("Excel.Application")
        appCible.Visible = False
        appCible.DisplayAlerts = False
        appCible.ScreenUpdating = False
        appCible.EnableEvents = False
        appCible.Calculation = xlCalculationManual
    Else
        Set appCible = Application
    End If

    chemin = ExtraireDataSource(Replace(CONNEXION_EXCEL_BDD_TRAVAIL, "UserProfile", Environ$("UserProfile")))
    Debug.Print "  Chemin BdD : " & chemin

    t1 = Timer

    Set wb = appCible.Workbooks.Open(Filename:=chemin, UpdateLinks:=0, _
                                     ReadOnly:=True, Notify:=False)
    Set ws = wb.Worksheets("EMMST")

    t2 = Timer

    Call LireCommentairesEtMesurer(ws)

    t3 = Timer

    wb.Close SaveChanges:=False
    Set wb = Nothing
    Set ws = Nothing

    If creerInstanceSeparee Then
        appCible.Quit
        Set appCible = Nothing
    End If

    ' Restauration etat
    Application.ScreenUpdating = oldScreen
    Application.EnableEvents = oldEvents
    Application.Calculation = oldCalc
    Application.DisplayAlerts = oldAlerts

    Debug.Print "  Ouverture : " & Format(t2 - t1, "0.000") & " s"
    Debug.Print "  Lecture   : " & Format(t3 - t2, "0.000") & " s"
    Debug.Print "  Setup+Fermeture : " & Format((t1 - t0) + (Timer - t3), "0.000") & " s"
    Debug.Print "  Total     : " & Format(Timer - t0, "0.000") & " s"

End Sub

Private Sub LireCommentairesEtMesurer(ByRef ws As Worksheet)
' Trouve les 11 colonnes (DTR + 10 commentaires) via Find en ligne 1,
' lit chaque colonne en bloc via Range.Value, calcule longueur max.
' Pattern "lecture bloc .Value" = ce qui serait utilise en prod pour
' peupler le Dictionary DTR -> commentaires.

    Dim colNoms As Variant
    Dim colIndices() As Long
    Dim iCol As Long
    Dim derniereLigne As Long
    Dim dtrs As Variant
    Dim valeurs As Variant
    Dim maxLong As Long
    Dim dtrLong As String
    Dim ligne As Long
    Dim v As Variant
    Dim lg As Long

    colNoms = Array("DTR_EMMST", _
                    "CommentaireEmballage_EMMST", _
                    "CommentaireMarquage_EMMST", _
                    "CommentaireManutention_EMMST", _
                    "CommentaireStockage_EMMST", _
                    "CommentaireTransport_EMMST", _
                    "CommentaireEmballageEN_EMMST", _
                    "CommentaireMarquageEN_EMMST", _
                    "CommentaireManutentionEN_EMMST", _
                    "CommentaireStockageEN_EMMST", _
                    "CommentaireTransportEN_EMMST")

    ReDim colIndices(LBound(colNoms) To UBound(colNoms))
    For iCol = LBound(colNoms) To UBound(colNoms)
        colIndices(iCol) = TrouverColonne(ws, CStr(colNoms(iCol)))
    Next iCol

    ' Derniere ligne de donnees : scan descendant sur la colonne DTR
    With ws
        derniereLigne = .Cells(.Rows.Count, colIndices(LBound(colNoms))).End(xlUp).Row

        ' Lecture bloc de la colonne DTR (2D)
        dtrs = .Range(.Cells(2, colIndices(0)), .Cells(derniereLigne, colIndices(0))).Value

        ' Pour chaque colonne commentaire : lecture bloc + mesure longueur max
        For iCol = LBound(colNoms) + 1 To UBound(colNoms)
            valeurs = .Range(.Cells(2, colIndices(iCol)), .Cells(derniereLigne, colIndices(iCol))).Value

            maxLong = 0
            dtrLong = ""
            For ligne = 1 To UBound(valeurs, 1)
                v = valeurs(ligne, 1)
                If Not IsEmpty(v) And Not IsNull(v) Then
                    lg = Len(CStr(v))
                    If lg > maxLong Then
                        maxLong = lg
                        dtrLong = CStr(dtrs(ligne, 1))
                    End If
                End If
            Next ligne

            Debug.Print "    " & PadRight(CStr(colNoms(iCol)), 36) & _
                        " max=" & Format(maxLong, "0000") & _
                        "  DTR=" & dtrLong
        Next iCol
    End With

End Sub

Private Function TrouverColonne(ByRef ws As Worksheet, ByVal nomColonne As String) As Long
' Cherche l'en-tete nomColonne en ligne 1 et retourne l'index de colonne.
' Pattern Find convention 2 du projet : xlValues + xlWhole, fallback xlPart.
' Leve une erreur si introuvable (pre-condition convention 3).

    Dim trouve As Range

    Set trouve = ws.Rows(1).Find(What:=nomColonne, LookIn:=xlValues, _
                                  LookAt:=xlWhole, MatchCase:=False)
    If trouve Is Nothing Then
        Set trouve = ws.Rows(1).Find(What:=nomColonne, LookIn:=xlValues, _
                                      LookAt:=xlPart, MatchCase:=False)
    End If

    If trouve Is Nothing Then
        Err.Raise vbObjectError + 520, "TrouverColonne", _
                  "Colonne introuvable dans la feuille EMMST : " & nomColonne
    End If

    TrouverColonne = trouve.Column

End Function

Private Function ExtraireDataSource(ByVal connStr As String) As String
' Extrait la valeur de Data Source= dans une chaine de connexion ADO.
' Forme attendue : Provider=...;Data Source=<chemin>;Extended Properties='...';

    Dim posStart As Long
    Dim posFin As Long
    Dim debut As Long

    posStart = InStr(1, connStr, "Data Source=", vbTextCompare)
    If posStart = 0 Then
        Err.Raise vbObjectError + 521, "ExtraireDataSource", _
                  "Chaine sans 'Data Source='."
    End If

    debut = posStart + Len("Data Source=")
    posFin = InStr(debut, connStr, ";")
    If posFin = 0 Then posFin = Len(connStr) + 1

    ExtraireDataSource = Mid$(connStr, debut, posFin - debut)

End Function

Private Function PadRight(ByVal s As String, ByVal n As Long) As String
' Complete s avec des espaces a droite pour atteindre n caracteres.

    If Len(s) >= n Then
        PadRight = s
    Else
        PadRight = s & Space$(n - Len(s))
    End If

End Function
