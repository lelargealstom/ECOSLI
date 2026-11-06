' =========================================================================
' MODULE: X_Test_TypeGuessRows
' =========================================================================
' Auteur : Thomas LE LARGE - 2026-11-06
' Perimetre : diagnostic troncature 255 chars des commentaires EMMST via ACE OLE DB.
'
' But : verifier si le driver ACE 12 installe sur ce poste respecte les
'       parametres TypeGuessRows=0 / MaxScanRows=0 dans Extended Properties
'       de la chaine de connexion. Si oui, une des variantes doit faire
'       apparaitre des longueurs > 255 sur au moins une colonne commentaire.
'
' Usage : placer le curseur dans Test_TypeGuessRows_Commentaires puis F5.
'         Resultats dans la fenetre Immediate (Ctrl+G dans VBE).
' =========================================================================

Option Explicit

Public Sub Test_TypeGuessRows_Commentaires()
' Lance 5 variantes de chaine de connexion et imprime la longueur max
' observee sur chacune des 10 colonnes de commentaires de la table EMMST$.
' Verdict : si une variante B/C/D/E remonte > 255 la ou A plafonne a 255,
' le parametre est respecte -> Piste 1 validee.

    Dim variantes As Variant
    Dim libelles As Variant
    Dim iVar As Long

    libelles = Array( _
        "A - actuel (HDR=Yes;IMEX=1)", _
        "B - TypeGuessRows=0 seul", _
        "C - MaxScanRows=0 seul", _
        "D - TypeGuessRows=0;MaxScanRows=0", _
        "E - sans IMEX + TypeGuessRows=0;MaxScanRows=0")

    variantes = Array( _
        "Excel 12.0 Macro;HDR=Yes;IMEX=1", _
        "Excel 12.0 Macro;HDR=Yes;IMEX=1;TypeGuessRows=0", _
        "Excel 12.0 Macro;HDR=Yes;IMEX=1;MaxScanRows=0", _
        "Excel 12.0 Macro;HDR=Yes;IMEX=1;TypeGuessRows=0;MaxScanRows=0", _
        "Excel 12.0 Macro;HDR=Yes;TypeGuessRows=0;MaxScanRows=0")

    Debug.Print String(72, "=")
    Debug.Print "Test TypeGuessRows / MaxScanRows sur BdD TRAVAIL - " & Format(Now, "yyyy-mm-dd hh:nn:ss")
    Debug.Print String(72, "=")

    For iVar = LBound(variantes) To UBound(variantes)
        Debug.Print
        Debug.Print "--- Variante " & libelles(iVar) & " ---"
        Call MesurerLongueursCommentaires(CStr(variantes(iVar)))
    Next iVar

    Debug.Print
    Debug.Print String(72, "=")
    Debug.Print "Verdict :"
    Debug.Print "  * si toutes les variantes plafonnent a 255 -> ACE ignore"
    Debug.Print "    les parametres dans la chaine de connexion (comportement"
    Debug.Print "    documente par Microsoft). Plan B necessaire."
    Debug.Print "  * si une variante remonte > 255 sur au moins une colonne,"
    Debug.Print "    on l'adopte dans la chaine de connexion EMMST."
    Debug.Print String(72, "=")

    MsgBox "Test termine. Ouvre la fenetre Immediate (Ctrl+G) pour comparer." & vbCrLf & _
           "Une longueur > 255 sur une variante B/C/D/E valide la Piste 1.", _
           vbInformation, "Test TypeGuessRows"

End Sub

Private Sub MesurerLongueursCommentaires(ByVal extendedProperties As String)
' Ouvre la BdD TRAVAIL avec les Extended Properties fournies, parcourt
' EMMST$ et imprime la longueur max observee sur les 10 colonnes
' commentaires (FR + EN).

    Dim chaine As String
    Dim colNoms As Variant
    Dim maxLong() As Long
    Dim dtrLong() As String
    Dim iCol As Long
    Dim v As Variant
    Dim lg As Long
    Dim nbLignes As Long

    colNoms = Array( _
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

    ReDim maxLong(LBound(colNoms) To UBound(colNoms))
    ReDim dtrLong(LBound(colNoms) To UBound(colNoms))

    chaine = Replace(CONNEXION_EXCEL_BDD_TRAVAIL, "UserProfile", Environ$("UserProfile"))
    chaine = ForcerExtendedProperties(chaine, extendedProperties)
    Debug.Print "ConnStr : " & chaine

    Call Open_CONN(chaine)
    Call open_REQSEL( _
        "SELECT [DTR_EMMST], [CommentaireEmballage_EMMST], [CommentaireMarquage_EMMST], " & _
        "[CommentaireManutention_EMMST], [CommentaireStockage_EMMST], [CommentaireTransport_EMMST], " & _
        "[CommentaireEmballageEN_EMMST], [CommentaireMarquageEN_EMMST], [CommentaireManutentionEN_EMMST], " & _
        "[CommentaireStockageEN_EMMST], [CommentaireTransportEN_EMMST] FROM [EMMST$]")

    If Not REQSEL.EOF Then
        REQSEL.MoveFirst
        Do Until REQSEL.EOF
            nbLignes = nbLignes + 1
            For iCol = LBound(colNoms) To UBound(colNoms)
                v = REQSEL.Fields(CStr(colNoms(iCol))).Value
                If Not IsNull(v) Then
                    lg = Len(CStr(v))
                    If lg > maxLong(iCol) Then
                        maxLong(iCol) = lg
                        dtrLong(iCol) = CStr(REQSEL.Fields("DTR_EMMST").Value)
                    End If
                End If
            Next iCol
            REQSEL.MoveNext
        Loop
    End If

    Call Close_CONN(chaine)

    Debug.Print "  Lignes scannees : " & nbLignes
    For iCol = LBound(colNoms) To UBound(colNoms)
        Debug.Print "  " & PadRight(CStr(colNoms(iCol)), 36) & _
                    " max=" & Format(maxLong(iCol), "000") & _
                    "  DTR=" & dtrLong(iCol)
    Next iCol

End Sub

Private Function ForcerExtendedProperties(ByVal connStr As String, ByVal newExt As String) As String
' Remplace la portion Extended Properties='...' de connStr par newExt.
' Forme attendue : ...Extended Properties='Excel 12.0 Macro;HDR=Yes;IMEX=1';

    Dim posStart As Long
    Dim posQuoteOuvrante As Long
    Dim posQuoteFermante As Long

    posStart = InStr(1, connStr, "Extended Properties=", vbTextCompare)
    If posStart = 0 Then
        Err.Raise vbObjectError + 514, "ForcerExtendedProperties", _
                  "Chaine sans 'Extended Properties='."
    End If

    posQuoteOuvrante = InStr(posStart, connStr, "'")
    posQuoteFermante = InStr(posQuoteOuvrante + 1, connStr, "'")
    If posQuoteOuvrante = 0 Or posQuoteFermante = 0 Then
        Err.Raise vbObjectError + 515, "ForcerExtendedProperties", _
                  "Quotes simples manquantes autour d'Extended Properties."
    End If

    ForcerExtendedProperties = Left$(connStr, posQuoteOuvrante) & _
                               newExt & _
                               Mid$(connStr, posQuoteFermante)

End Function

Private Function PadRight(ByVal s As String, ByVal n As Long) As String
' Complete s avec des espaces a droite pour atteindre n caracteres.
' Utilitaire d'alignement dans la fenetre Immediate.

    If Len(s) >= n Then
        PadRight = s
    Else
        PadRight = s & Space$(n - Len(s))
    End If

End Function
