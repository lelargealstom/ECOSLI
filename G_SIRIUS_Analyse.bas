Attribute VB_Name = "G_SIRIUS_Analyse"
' ============================================================
' CONSTANTES ANALYSE
' ============================================================
Private Const SEUIL_DVL_MAX     As Double = 50
Private Const SEUIL_KO_BLOQUANT As Double = 0.2  ' 20%

' ============================================================
' STRUCTURE RÉSULTAT ANALYSE
' ============================================================
' Stocké dans une Collection globale pour accès depuis UI
Public resultatsAnalyse As Collection

' ============================================================
' POINT D'ENTRÉE — LANCER L'ANALYSE PRODUIT
' ============================================================

Public Sub SIRIUS_AnalyserProduit()

    Dim lv     As MSComctlLib.ListView
    Dim lvErr  As MSComctlLib.ListView

    Set lv = UserForm.ListView_Produit

    ' -- Vérifier qu'il y a des données --------------------
    If lv.ListItems.count = 0 Then
        MsgBox "Aucune ligne à analyser.", vbExclamation, "SIRIUS — Analyse"
        Exit Sub
    End If

    ' -- Initialiser la collection de résultats ------------
    Set resultatsAnalyse = New Collection

    ' -- Masquer le frame de saisie ------------------------
    UserForm.Frame_Saisie.Visible = False

    ' -- Configurer la ListView d'erreurs ------------------
    Call SIRIUS_ConfigListViewErreurs

    Set lvErr = UserForm.ListView_Erreurs

    ' -- Compteurs -----------------------------------------
    Dim nbTotal         As Long
    Dim nbObligatOK     As Long: nbObligatOK = 0
    Dim nbObligatKO     As Long: nbObligatKO = 0
    Dim nbDoublons      As Long: nbDoublons = 0
    Dim nbRefKO         As Long: nbRefKO = 0
    Dim nbRefWarn       As Long: nbRefWarn = 0
    Dim nbTypeKO        As Long: nbTypeKO = 0
    Dim nbExistBdD      As Long: nbExistBdD = 0

    nbTotal = lv.ListItems.count

    ' -- Récupérer les DTR existants en BdD ----------------
    Dim dtrBdD As Collection
    Set dtrBdD = SIRIUS_GetDTRFromBdD()

    ' -- Récupérer les DTR déjà saisis pour doublon interne-
    Dim dtrSaisis As New Collection

    ' -- Boucle sur chaque ligne ---------------------------
    Dim li       As MSComctlLib.ListItem
    Dim ligneNum As Long
    ligneNum = 1

    For Each li In lv.ListItems

        ' Colonnes selon SIRIUS_ConfigListViewProduit :
        ' Text    = DTR
        ' SI(1)   = Désignation
        ' SI(2)   = Node Type
        ' SI(3)   = Fournisseur
        ' SI(4)   = Réf. Fournisseur
        ' SI(5)   = Drawing Number
        ' SI(6)   = MTBF Intr.
        ' SI(7)   = DVL HARD
        ' SI(8)   = DVL SOFT

        Dim dtr     As String: dtr = Trim(li.Text)
        Dim des     As String: des = Trim(li.SubItems(1))
        Dim node    As String: node = Trim(li.SubItems(2))
        Dim fourn   As String: fourn = Trim(li.SubItems(3))
        Dim ref     As String: ref = Trim(li.SubItems(4))
        Dim drw     As String: drw = Trim(li.SubItems(5))
        Dim MTBF    As String: MTBF = Trim(li.SubItems(6))
        Dim dvlh    As String: dvlh = Trim(li.SubItems(7))
        Dim dvls    As String: dvls = Trim(li.SubItems(8))

        Dim ligneOK As Boolean: ligneOK = True

        ' -- CONTRÔLE 1 — Champs obligatoires --------------
        Dim champsMissing As String: champsMissing = ""

        If dtr = "" Then champsMissing = champsMissing & "DTR / "
        If des = "" Then champsMissing = champsMissing & "Désignation / "
        If node = "" Then champsMissing = champsMissing & "Node Type / "
        If fourn = "" Then champsMissing = champsMissing & "Fournisseur / "
        If ref = "" Then champsMissing = champsMissing & "Réf. Fournisseur / "
        If drw = "" Then champsMissing = champsMissing & "Drawing Number / "
        If MTBF = "" Then champsMissing = champsMissing & "MTBF Intrinsèque / "
        If dvlh = "" Then champsMissing = champsMissing & "DVL HARD / "

        If champsMissing <> "" Then
            nbObligatKO = nbObligatKO + 1
            ligneOK = False
            Call SIRIUS_AddErreur(lvErr, ligneNum, dtr, _
                "Champ obligatoire", _
                "Renseigner : " & Left(champsMissing, Len(champsMissing) - 3), _
                "KO")
        Else
            nbObligatOK = nbObligatOK + 1
        End If

        ' -- CONTRÔLE 2 — Doublon DTR interne --------------
        If dtr <> "" Then
            Dim isDupInternal As Boolean: isDupInternal = False
            Dim dtrCheck As Variant
            For Each dtrCheck In dtrSaisis
                If UCase(dtrCheck) = UCase(dtr) Then
                    isDupInternal = True
                    Exit For
                End If
            Next dtrCheck

            If isDupInternal Then
                nbDoublons = nbDoublons + 1
                ligneOK = False
                Call SIRIUS_AddErreur(lvErr, ligneNum, dtr, _
                    "Doublon DTR", _
                    "DTR déjà présent dans la liste de saisie", _
                    "KO")
            Else
                dtrSaisis.Add dtr
            End If
        End If

        ' -- CONTRÔLE 3 — DTR existant en BdD --------------
        If dtr <> "" Then
            Dim existBdD As Boolean: existBdD = False
            Dim dtrBdDItem As Variant
            For Each dtrBdDItem In dtrBdD
                If UCase(dtrBdDItem) = UCase(dtr) Then
                    existBdD = True
                    Exit For
                End If
            Next dtrBdDItem

            If existBdD Then
                nbExistBdD = nbExistBdD + 1
                Call SIRIUS_AddErreur(lvErr, ligneNum, dtr, _
                    "DTR existant BdD", _
                    "DTR déjà présent en BdD — vérifier si Création ou Modification", _
                    "WARN")
            End If
        End If

        ' -- CONTRÔLE 4 — Types de données -----------------

        ' MTBF — décimal > 0
        If MTBF <> "" Then
            If Not SIRIUS_IsDecimal(MTBF) Then
                nbTypeKO = nbTypeKO + 1
                ligneOK = False
                Call SIRIUS_AddErreur(lvErr, ligneNum, dtr, _
                    "Type MTBF", _
                    "MTBF Intrinsèque doit être un nombre décimal > 0 (valeur : " & MTBF & ")", _
                    "KO")
            End If
        End If

        ' DVL HARD — décimal > 0 et <= 50
        If dvlh <> "" Then
            If Not SIRIUS_IsDecimal(dvlh) Then
                nbTypeKO = nbTypeKO + 1
                ligneOK = False
                Call SIRIUS_AddErreur(lvErr, ligneNum, dtr, _
                    "Type DVL HARD", _
                    "DVL HARD doit être un nombre décimal (valeur : " & dvlh & ")", _
                    "KO")
            ElseIf CDbl(dvlh) > SEUIL_DVL_MAX Then
                nbTypeKO = nbTypeKO + 1
                ligneOK = False
                Call SIRIUS_AddErreur(lvErr, ligneNum, dtr, _
                    "DVL HARD > 50", _
                    "DVL HARD = " & dvlh & " — durée de vie limitée ne peut pas dépasser 50 ans", _
                    "KO")
            End If
        End If

        ' DVL SOFT — décimal > 0 et <= 50 (si renseigné)
        If dvls <> "" Then
            If Not SIRIUS_IsDecimal(dvls) Then
                nbTypeKO = nbTypeKO + 1
                ligneOK = False
                Call SIRIUS_AddErreur(lvErr, ligneNum, dtr, _
                    "Type DVL SOFT", _
                    "DVL SOFT doit être un nombre décimal (valeur : " & dvls & ")", _
                    "KO")
            ElseIf CDbl(dvls) > SEUIL_DVL_MAX Then
                nbTypeKO = nbTypeKO + 1
                ligneOK = False
                Call SIRIUS_AddErreur(lvErr, ligneNum, dtr, _
                    "DVL SOFT > 50", _
                    "DVL SOFT = " & dvls & " — durée de vie limitée ne peut pas dépasser 50 ans", _
                    "KO")
            End If
        End If

        ligneNum = ligneNum + 1
    Next li

    ' -- Calcul du taux KO ---------------------------------
    Dim nbKOTotal  As Long
    Dim tauxKO     As Double
    nbKOTotal = nbObligatKO + nbDoublons + nbTypeKO
    If nbTotal > 0 Then tauxKO = nbKOTotal / nbTotal Else tauxKO = 0

    ' -- Afficher les résultats ----------------------------
    Call SIRIUS_AfficherResultatsAnalyse( _
        nbTotal, nbObligatOK, nbObligatKO, _
        nbDoublons, nbExistBdD, nbRefWarn, _
        nbTypeKO, tauxKO)

    ' -- Afficher le frame analyse -------------------------
    UserForm.Frame_Analyse.Visible = True

    ' Activer ou bloquer le bouton demande selon taux KO
    UserForm.Btn_CreerDemande.Enabled = (tauxKO <= SEUIL_KO_BLOQUANT)

    If tauxKO > SEUIL_KO_BLOQUANT Then
        UserForm.lbl_StatutAnalyse.Caption = _
            "BLOQUÉ — Taux d'erreurs " & Format(tauxKO * 100, "0.0") & _
            "% > 20% — corrigez avant d'envoyer"
        UserForm.lbl_StatutAnalyse.ForeColor = RGB(163, 45, 45)
    Else
        UserForm.lbl_StatutAnalyse.Caption = _
            "Analyse terminée — " & nbKOTotal & " erreur(s) · " & _
            Format(tauxKO * 100, "0.0") & "% KO"
        UserForm.lbl_StatutAnalyse.ForeColor = IIf(nbKOTotal = 0, RGB(39, 80, 10), RGB(186, 117, 23))
    End If

End Sub

' ============================================================
' RETOUR VERS LA SAISIE
' ============================================================

Public Sub SIRIUS_RetourSaisie()

    ' Masquer l'analyse, réafficher la saisie
    UserForm.Frame_Analyse.Visible = False
    UserForm.Frame_Saisie.Visible = True

    ' ListView_Produit reste intacte — pas de clear

End Sub

' ============================================================
' RÉCUPÉRER LES DTR EXISTANTS EN BDD
' ============================================================

Private Function SIRIUS_GetDTRFromBdD() As Collection

    Dim col As New Collection

    On Error GoTo ErrBdD

    Set REQSEL = CreateObject("ADODB.Recordset")

    STRSEL = "SELECT [DTR Reference (issue de PBS)_PRODUIT] FROM [Produit$]"
    Call open_REQSEL(STRSEL)

    Do While Not REQSEL.EOF
        Dim val As String
        val = ""
        If Not IsNull(REQSEL.Fields(0).Value) Then
            val = Trim(REQSEL.Fields(0).Value)
        End If
        If val <> "" Then
            On Error Resume Next
            col.Add val, val
            On Error GoTo ErrBdD
        End If
        REQSEL.MoveNext
    Loop

    Set SIRIUS_GetDTRFromBdD = col
    Exit Function

ErrBdD:
    ' En cas d'erreur BdD — on retourne une collection vide
    Set SIRIUS_GetDTRFromBdD = New Collection

End Function

' ============================================================
' AJOUTER UNE ERREUR DANS LA LISTVIEW
' ============================================================

Private Sub SIRIUS_AddErreur(lvErr As MSComctlLib.ListView, _
                              ligne As Long, _
                              dtr As String, _
                              controle As String, _
                              action As String, _
                              niveau As String)

    Dim li As MSComctlLib.ListItem
    Set li = lvErr.ListItems.Add(, , CStr(ligne))

    li.SubItems(1) = dtr
    li.SubItems(2) = controle
    li.SubItems(3) = action

    ' Stocker aussi dans la collection globale
    resultatsAnalyse.Add Array(ligne, dtr, controle, action, niveau)

End Sub


' ============================================================
' AFFICHER LES RÉSULTATS DANS LES LABELS DU FRAME_ANALYSE
' ============================================================

Private Sub SIRIUS_AfficherResultatsAnalyse( _
    nbTotal As Long, _
    nbObligatOK As Long, nbObligatKO As Long, _
    nbDoublons As Long, _
    nbExistBdD As Long, _
    nbRefWarn As Long, _
    nbTypeKO As Long, _
    tauxKO As Double)

    With UserForm
        ' Champs obligatoires
        .lbl_ObligatVal.Caption = nbObligatOK & "/" & nbTotal
        .lbl_ObligatStatut.Caption = IIf(nbObligatKO = 0, "OK", nbObligatKO & " KO")
        .lbl_ObligatStatut.ForeColor = IIf(nbObligatKO = 0, RGB(39, 80, 10), RGB(163, 45, 45))

        ' Unicité DTR
        .lbl_DoublonsVal.Caption = IIf(nbDoublons = 0, "OK", nbDoublons & " KO")
        .lbl_DoublonsStatut.Caption = IIf(nbDoublons = 0, "OK", "Bloquant")
        .lbl_DoublonsStatut.ForeColor = IIf(nbDoublons = 0, RGB(39, 80, 10), RGB(163, 45, 45))

        ' Cohérence référentielle BdD
        .lbl_RefVal.Caption = IIf(nbExistBdD = 0, "OK", nbExistBdD & " exist.")
        .lbl_RefStatut.Caption = IIf(nbExistBdD = 0, "OK", "Alerte")
        .lbl_RefStatut.ForeColor = IIf(nbExistBdD = 0, RGB(39, 80, 10), RGB(186, 117, 23))

        ' Types de données
        .lbl_TypesVal.Caption = IIf(nbTypeKO = 0, "OK", nbTypeKO & " KO")
        .lbl_TypesStatut.Caption = IIf(nbTypeKO = 0, "OK", "Bloquant")
        .lbl_TypesStatut.ForeColor = IIf(nbTypeKO = 0, RGB(39, 80, 10), RGB(163, 45, 45))

        ' DTR connus en BdD
        .lbl_ExistVal.Caption = IIf(nbExistBdD = 0, "0", nbExistBdD & " exist.")
        .lbl_ExistStatut.Caption = IIf(nbExistBdD = 0, "OK", "Info")
        .lbl_ExistStatut.ForeColor = IIf(nbExistBdD = 0, RGB(39, 80, 10), RGB(24, 95, 165))
    End With

End Sub

' ============================================================
' UTILITAIRE — VÉRIFIER QU'UNE VALEUR EST UN DÉCIMAL VALIDE
' ============================================================

Private Function SIRIUS_IsDecimal(val As String) As Boolean
    On Error Resume Next
    Dim d As Double
    d = CDbl(val)
    SIRIUS_IsDecimal = (Err.Number = 0 And d > 0)
    On Error GoTo 0
End Function

