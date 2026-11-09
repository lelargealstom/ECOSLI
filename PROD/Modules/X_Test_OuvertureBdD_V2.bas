' =========================================================================
' MODULE: X_Test_OuvertureBdD_V2
' =========================================================================
' Auteur : Thomas LE LARGE - 2026-11-09
' Perimetre : bench v2 pour localiser les 27 s de "Setup+Fermeture" observees
'             au bench v1 (scenario A : instance courante masquee).
'
' But : ventiler finement les phases du cycle ouverture/lecture/fermeture
'       et tester 2 hypotheses pour reduire le cout total :
'       - A2 : laisser Calculation en manuel (ne pas restaurer en auto)
'              -> teste l'hypothese "recalcul global sur ActiveWorkbook ECOSLI".
'       - A3 : copier la BdD en local %TEMP% avant ouverture, supprimer apres
'              -> teste l'hypothese "latence OneDrive sur I/O reseau".
'
' Usage : placer le curseur dans Test_OuvertureBdD_V2 puis F5.
'         Resultats dans la fenetre Immediate (Ctrl+G).
'
' NB : ne modifie jamais la BdD (ReadOnly + SaveChanges:=False a la fermeture).
'      En A2, Application.Calculation est remise en xlCalculationAutomatic a
'      la fin du Test (pour ne pas laisser le classeur utilisateur en manuel).
' =========================================================================

Option Explicit

Public Sub Test_OuvertureBdD_V2()
' Lance 3 scenarios et imprime la ventilation par phase :
'   A1 : baseline bench v1 avec ventilation fine.
'   A2 : baseline sans restauration Calculation (manuel laisse jusqu'a la fin).
'   A3 : copie locale %TEMP% + ouverture locale + Kill.

    Debug.Print String(72, "=")
    Debug.Print "Bench V2 ouverture BdD TRAVAIL masquee - " & Format(Now, "yyyy-mm-dd hh:nn:ss")
    Debug.Print String(72, "=")

    Debug.Print
    Debug.Print "### Scenario A1 : baseline bench v1 + ventilation fine ###"
    Call BenchVentile(copieLocale:=False, restaurerCalculation:=True)

    Debug.Print
    Debug.Print "### Scenario A2 : sans restauration Calculation (manuel maintenu) ###"
    Call BenchVentile(copieLocale:=False, restaurerCalculation:=False)

    ' Filet de securite : si A2 a laisse Calculation en manuel, on remet en auto
    ' pour ne pas perturber la session utilisateur.
    Application.Calculation = xlCalculationAutomatic

    Debug.Print
    Debug.Print "### Scenario A3 : copie locale %TEMP% + ouverture locale + Kill ###"
    Call BenchVentile(copieLocale:=True, restaurerCalculation:=True)

    Debug.Print
    Debug.Print String(72, "=")
    Debug.Print "Lecture :"
    Debug.Print "  * A1 ventilation : identifie le poste couteux (close vs restauration)."
    Debug.Print "  * A2 vs A1 : si phase 'Restauration app' chute a ~0 s en A2 et"
    Debug.Print "    coutait 20+ s en A1, le coupable est le recalcul auto sur"
    Debug.Print "    ActiveWorkbook ECOSLI -> on peut adopter 'manuel pendant lecture'."
    Debug.Print "  * A3 vs A1 : si Ouverture + Fermeture chute drastiquement en A3,"
    Debug.Print "    le coupable est la latence OneDrive -> on peut adopter"
    Debug.Print "    la copie locale a chaque appel (ou une fois par session)."
    Debug.Print "  * Calculation remise en xlCalculationAutomatic a la fin."
    Debug.Print String(72, "=")

    MsgBox "Bench V2 termine. Ctrl+G pour la ventilation detaillee." & vbCrLf & _
           "Application.Calculation remise en xlCalculationAutomatic.", _
           vbInformation, "Bench V2"

End Sub

Private Sub BenchVentile(ByVal copieLocale As Boolean, ByVal restaurerCalculation As Boolean)
' Chronometre les 6 phases du cycle :
'   t0 -> t1 : setup Application (ScreenUpdating/EnableEvents/Calculation/Alerts)
'   t1 -> t2 : copie locale FileCopy si copieLocale, sinon nop
'   t2 -> t3 : Workbooks.Open ReadOnly
'   t3 -> t4 : lecture bloc 10 colonnes + DTR via Range.Value
'   t4 -> t5 : wb.Close SaveChanges:=False
'   t5 -> t6 : restauration Application (partielle si restaurerCalculation=False)
'              + Kill copie locale si copieLocale.

    Dim t0 As Double, t1 As Double, t2 As Double
    Dim t3 As Double, t4 As Double, t5 As Double, t6 As Double

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim cheminBdD As String
    Dim cheminTravail As String

    Dim oldScreen As Boolean
    Dim oldEvents As Boolean
    Dim oldCalc As XlCalculation
    Dim oldAlerts As Boolean

    oldScreen = Application.ScreenUpdating
    oldEvents = Application.EnableEvents
    oldCalc = Application.Calculation
    oldAlerts = Application.DisplayAlerts

    t0 = Timer

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False

    t1 = Timer

    cheminBdD = ExtraireDataSource(Replace(CONNEXION_EXCEL_BDD_TRAVAIL, "UserProfile", Environ$("UserProfile")))

    If copieLocale Then
        cheminTravail = Environ$("TEMP") & "\ECOSLI_bench_bdd_" & Format(Now, "hhnnss") & ".xlsb"
        FileCopy cheminBdD, cheminTravail
    Else
        cheminTravail = cheminBdD
    End If

    t2 = Timer

    Set wb = Workbooks.Open(Filename:=cheminTravail, UpdateLinks:=0, _
                             ReadOnly:=True, Notify:=False)
    Set ws = wb.Worksheets("EMMST")

    t3 = Timer

    Call LireCommentairesBlocSilencieux(ws)

    t4 = Timer

    wb.Close SaveChanges:=False
    Set wb = Nothing
    Set ws = Nothing

    t5 = Timer

    Application.ScreenUpdating = oldScreen
    Application.EnableEvents = oldEvents
    If restaurerCalculation Then Application.Calculation = oldCalc
    Application.DisplayAlerts = oldAlerts

    If copieLocale Then
        Kill cheminTravail
    End If

    t6 = Timer

    Debug.Print "  Setup Application         : " & Format(t1 - t0, "0.000") & " s"
    If copieLocale Then
        Debug.Print "  Copie locale FileCopy     : " & Format(t2 - t1, "0.000") & " s"
    End If
    Debug.Print "  Ouverture Workbooks.Open  : " & Format(t3 - t2, "0.000") & " s"
    Debug.Print "  Lecture bloc 10 colonnes  : " & Format(t4 - t3, "0.000") & " s"
    Debug.Print "  Fermeture wb.Close        : " & Format(t5 - t4, "0.000") & " s"
    Debug.Print "  Restauration app" & _
                IIf(restaurerCalculation, " (incl. Calculation)", " (sans Calculation)") & _
                IIf(copieLocale, " + Kill copie", "") & _
                " : " & Format(t6 - t5, "0.000") & " s"
    Debug.Print "  ------------------------------------"
    Debug.Print "  TOTAL                     : " & Format(t6 - t0, "0.000") & " s"

End Sub

Private Sub LireCommentairesBlocSilencieux(ByRef ws As Worksheet)
' Meme logique que bench v1 (Find + lecture bloc Range.Value) mais sans
' Debug.Print : on veut mesurer uniquement le cout brut de la lecture, pas
' la serialisation vers la fenetre Immediate.
' La variable dummy force VBA a consommer le tableau 2D retourne par
' Range.Value (sinon le compilateur pourrait optimiser l'appel).

    Dim colNoms As Variant
    Dim colIndices() As Long
    Dim iCol As Long
    Dim derniereLigne As Long
    Dim dtrs As Variant
    Dim valeurs As Variant
    Dim dummy As Long

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

    With ws
        derniereLigne = .Cells(.Rows.Count, colIndices(0)).End(xlUp).Row
        dtrs = .Range(.Cells(2, colIndices(0)), .Cells(derniereLigne, colIndices(0))).Value
        dummy = UBound(dtrs, 1)

        For iCol = LBound(colNoms) + 1 To UBound(colNoms)
            valeurs = .Range(.Cells(2, colIndices(iCol)), .Cells(derniereLigne, colIndices(iCol))).Value
            dummy = UBound(valeurs, 1)
        Next iCol
    End With

End Sub

Private Function TrouverColonne(ByRef ws As Worksheet, ByVal nomColonne As String) As Long
' Cherche l'en-tete nomColonne en ligne 1 et retourne l'index de colonne.
' Pattern Find convention 2 : xlValues + xlWhole, fallback xlPart.

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
