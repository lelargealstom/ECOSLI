' ============================================================
' G_SIRUS_Analyse - Module d'analyse SIRUS
' ============================================================
' Deux points d'entree publics coexistent :
'   - SIRUS_AnalyserProduit (legacy) : analyse depuis ListView_Produit du UserForm
'   - SIRUS_AnalyserFeuille  (Phase 2, l. 391) : analyse depuis feuille SIRUS_<Table>_<Op>
'     Pilote par G_SIRUS_Config + G_SIRUS_Regles. Statut en col 1 + commentaire de cellule.
' ============================================================
' CONSTANTES ANALYSE
' ============================================================
Private Const SEUIL_DVL_MAX     As Double = 50
Private Const SEUIL_KO_BLOQUANT As Double = 0.2  ' 20%

' ============================================================
' STRUCTURE RESULTAT ANALYSE
' ============================================================
' Stocké dans une Collection globale pour accès depuis UI
Public resultatsAnalyse As Collection

' sirusAlertes : alertes structurees produites par SIRUS_AnalyserFeuille.
' Chaque entree = Array(cle, champ, texteAlerte, ligne, colChamp). Consomme
' par SIRUS_RemplirListViewResultats (UserForm) pour afficher 1 ligne par
' alerte (au lieu d 1 ligne par DTR avec alertes concatenees).
'   cle       : DTR (ou Nom projet pour ProjetSLI) lu en col 3
'   champ     : libelle humain du config (def.Champ)
'   alerte    : "Obligatoire manquant" ou retour de SIRUS_RegleAlerte
'   ligne     : numero de ligne dans la feuille (pour goto)
'   colChamp  : numero de la colonne du champ (pour goto la cellule fautive)
Public sirusAlertes As Collection

' ============================================================
' POINT D'ENTREE A LANCER L'ANALYSE PRODUIT
' ============================================================

Public Sub SIRUS_AnalyserProduit()

    Dim lv     As MSComctlLib.ListView
    Dim lvErr  As MSComctlLib.ListView

    Set lv = UserForm.ListView_Produit

    ' -- Vérifier qu'il y a des données --------------------
    If lv.ListItems.count = 0 Then
        MsgBox "Aucune ligne à analyser.", vbExclamation, "SIRUS - Analyse"
        Exit Sub
    End If

    ' -- Initialiser la collection de résultats ------------
    Set resultatsAnalyse = New Collection

    ' -- Masquer le frame de saisie ------------------------
    UserForm.Frame_Saisie.Visible = False

    ' -- Configurer la ListView d'erreurs ------------------
    Call SIRUS_ConfigListViewErreurs

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
    Set dtrBdD = SIRUS_GetDTRFromBdD()

    ' -- Récupérer les DTR déjà saisis pour doublon interne-
    Dim dtrSaisis As New Collection

    ' -- Boucle sur chaque ligne ---------------------------
    Dim li       As MSComctlLib.ListItem
    Dim ligneNum As Long
    ligneNum = 1

    For Each li In lv.ListItems

        ' Colonnes selon SIRUS_ConfigListViewProduit :
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

        ' -- CONTROLE 1 : Champs obligatoires --------------
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
            Call SIRUS_AddErreur(lvErr, ligneNum, dtr, _
                "Champ obligatoire", _
                "Renseigner : " & Left(champsMissing, Len(champsMissing) - 3), _
                "KO")
        Else
            nbObligatOK = nbObligatOK + 1
        End If

        ' -- CONTROLE 2 : Doublon DTR interne --------------
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
                Call SIRUS_AddErreur(lvErr, ligneNum, dtr, _
                    "Doublon DTR", _
                    "DTR déjà présent dans la liste de saisie", _
                    "KO")
            Else
                dtrSaisis.Add dtr
            End If
        End If

        ' -- CONTROLE 3 : DTR existant en BdD --------------
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
                Call SIRUS_AddErreur(lvErr, ligneNum, dtr, _
                    "DTR existant BdD", _
                    "DTR déjà présent en BdD à vérifier si Création ou Modification", _
                    "WARN")
            End If
        End If

        ' -- CONTROLE 4 Types de donnes -----------------

        ' MTBF à décimal > 0
        If MTBF <> "" Then
            If Not SIRUS_IsDecimal(MTBF) Then
                nbTypeKO = nbTypeKO + 1
                ligneOK = False
                Call SIRUS_AddErreur(lvErr, ligneNum, dtr, _
                    "Type MTBF", _
                    "MTBF Intrinsèque doit être un nombre décimal > 0 (valeur : " & MTBF & ")", _
                    "KO")
            End If
        End If

        ' DVL HARD à décimal > 0 et <= 50
        If dvlh <> "" Then
            If Not SIRUS_IsDecimal(dvlh) Then
                nbTypeKO = nbTypeKO + 1
                ligneOK = False
                Call SIRUS_AddErreur(lvErr, ligneNum, dtr, _
                    "Type DVL HARD", _
                    "DVL HARD doit être un nombre décimal (valeur : " & dvlh & ")", _
                    "KO")
            ElseIf CDbl(dvlh) > SEUIL_DVL_MAX Then
                nbTypeKO = nbTypeKO + 1
                ligneOK = False
                Call SIRUS_AddErreur(lvErr, ligneNum, dtr, _
                    "DVL HARD > 50", _
                    "DVL HARD = " & dvlh & " La durée de vie limitée ne peut pas dépasser 50 ans", _
                    "KO")
            End If
        End If

        ' DVL SOFT à décimal > 0 et <= 50 (si renseigné)
        If dvls <> "" Then
            If Not SIRUS_IsDecimal(dvls) Then
                nbTypeKO = nbTypeKO + 1
                ligneOK = False
                Call SIRUS_AddErreur(lvErr, ligneNum, dtr, _
                    "Type DVL SOFT", _
                    "DVL SOFT doit être un nombre décimal (valeur : " & dvls & ")", _
                    "KO")
            ElseIf CDbl(dvls) > SEUIL_DVL_MAX Then
                nbTypeKO = nbTypeKO + 1
                ligneOK = False
                Call SIRUS_AddErreur(lvErr, ligneNum, dtr, _
                    "DVL SOFT > 50", _
                    "DVL SOFT = " & dvls & " La durée de vie limitée ne peut pas dépasser 50 ans", _
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
    Call SIRUS_AfficherResultatsAnalyse( _
        nbTotal, nbObligatOK, nbObligatKO, _
        nbDoublons, nbExistBdD, nbRefWarn, _
        nbTypeKO, tauxKO)

    ' -- Afficher le frame analyse -------------------------
    UserForm.Frame_Analyse.Visible = True

    ' Activer ou bloquer le bouton demande selon taux KO
    UserForm.Btn_CreerDemande.Enabled = (tauxKO <= SEUIL_KO_BLOQUANT)

    If tauxKO > SEUIL_KO_BLOQUANT Then
        UserForm.lbl_StatutAnalyse.Caption = _
            "BLOQUE | Taux d'erreurs " & Format(tauxKO * 100, "0.0") & _
            "% > 20% | corrigez avant d'envoyer"
        UserForm.lbl_StatutAnalyse.ForeColor = RGB(163, 45, 45)
    Else
        UserForm.lbl_StatutAnalyse.Caption = _
            "Analyse terminée : " & nbKOTotal & " erreur(s) ! " & _
            Format(tauxKO * 100, "0.0") & "% KO"
        UserForm.lbl_StatutAnalyse.ForeColor = IIf(nbKOTotal = 0, RGB(39, 80, 10), RGB(186, 117, 23))
    End If

End Sub

' ============================================================
' RETOUR VERS LA SAISIE
' ============================================================

Public Sub SIRUS_RetourSaisie()

    ' Masquer l'analyse, réafficher la saisie
    UserForm.Frame_Analyse.Visible = False
    UserForm.Frame_Saisie.Visible = True

    ' ListView_Produit reste intacte à pas de clear

End Sub

' ============================================================
' RÉCUPÉRER LES DTR EXISTANTS EN BDD
' ============================================================

Private Function SIRUS_GetDTRFromBdD() As Collection

    Dim col As New Collection


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
        End If
        REQSEL.MoveNext
    Loop

    Set SIRUS_GetDTRFromBdD = col
End Function

' ============================================================
' AJOUTER UNE ERREUR DANS LA LISTVIEW
' ============================================================

Private Sub SIRUS_AddErreur(lvErr As MSComctlLib.ListView, _
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
' AFFICHER LES RESULTATS DANS LES LABELS DU FRAME_ANALYSE
' ============================================================

Private Sub SIRUS_AfficherResultatsAnalyse( _
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
' UTILITAIRE A VERIFIER QU'UNE VALEUR EST UN DECIMAL VALIDE
' ============================================================

Private Function SIRUS_IsDecimal(val As String) As Boolean
    On Error Resume Next
    Dim d As Double
    d = CDbl(val)
    SIRUS_IsDecimal = (Err.Number = 0 And d > 0)
    On Error GoTo 0
End Function

' =========================================================================
' PHASE 2 - SIRUS_AnalyserFeuille(sTable, [op])
' =========================================================================
' Analyse la feuille SIRUS_<sTable>_<op> active dans ActiveWorkbook :
'   - lit l'entete ligne 2 (ColonneBdD) + lit le config via SIRUS_LireConfig
'   - pour chaque ligne de donnees (a partir de la ligne 3) :
'       1. construit Dictionary {ColonneBdD -> valeur}
'       2. pour chaque champ Source = "Saisie" :
'           - manquant + (Obligatoire = "Oui" OU SIRUS_RegleObligatoire = True)
'             => alerte "Champ obligatoire manquant: <Champ>"
'           - SIRUS_RegleAlerte <> "" => ajoute l'alerte
'           - PoidsMeQ : cumul (rempli + obligatoire) pour calcul taux MeQ
'       3. ecrit le Statut dans la colonne en bout (vert OK / orange ALERTE)
'   - retourne le taux MeQ global (0..1) = sum(PoidsMeQ rempli oblig)/sum(PoidsMeQ oblig)
'
' Ne bloque jamais (decision verrouillee SIRUS roadmap #2). Pas de bloquant,
' juste OK / ALERTE et un taux de completude.
Public Function SIRUS_AnalyserFeuille(ByVal sTable As String, _
                                       Optional ByVal sOperation As String = "Creer") As Double

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim nomFeuille As String
    Dim col As Collection
    Dim def As SIRUS_Champ_Config
    Dim mapColBdD_Index As Object  ' Dictionary {ColonneBdD -> num colonne feuille}
    Dim nbCol As Long, colStatut As Long
    Dim ligne As Long, ligneMax As Long
    Dim ligneDict As Object
    Dim sStatutLigne As String
    Dim nAlertesLigne As Long
    Dim poidsLigneNum As Double, poidsLigneDen As Double
    Dim totNum As Double, totDen As Double
    Dim ligneVide As Boolean

    On Error GoTo Sortie

    ' -- 1. Localiser le classeur projet ---------------------
    Set wb = ActiveWorkbook
    If wb Is Nothing Then
        MsgBox "Aucun classeur actif.", vbExclamation, "SIRUS - Analyse"
        Exit Function
    End If

    ' -- 2. Localiser la feuille SIRUS_<Table>_<Op> ----------
    nomFeuille = "SIRUS_" & sTable & "_" & sOperation
    If Len(nomFeuille) > 31 Then nomFeuille = Left$(nomFeuille, 31)
    On Error Resume Next
    Set ws = wb.Worksheets(nomFeuille)
    On Error GoTo Sortie
    If ws Is Nothing Then
        MsgBox "Feuille '" & nomFeuille & "' introuvable dans '" & wb.Name & "'." & vbCrLf & _
               "Generer d'abord la feuille via SIRUS_GenererFeuilleAnalyse.", _
               vbExclamation, "SIRUS - Analyse"
        Exit Function
    End If

    ' -- 3. Charger le config -------------------------------
    Set col = SIRUS_LireConfig(sTable)
    If col.count = 0 Then
        MsgBox "Aucun champ defini dans SIRUS_Config pour '" & sTable & "'.", _
               vbCritical, "SIRUS - Analyse"
        Exit Function
    End If
    nbCol = col.count
    ' Layout refonte 2026-06-10 :
    '   col 1            : Statut (ecrit par cette fonction)
    '   col 2..(nbCol+1) : champs du config
    '   col (nbCol+2)    : Commentaires (saisie libre user)
    colStatut = 1

    ' -- 4. Construire mapping ColonneBdD -> num colonne feuille -----
    '    Lit la ligne 2 (entete BdD) en sautant col 1 (Statut). Les
    '    champs sont en col 2..(nbCol+1).
    Set mapColBdD_Index = CreateObject("Scripting.Dictionary")
    Dim i As Long
    Dim cellBdD As String
    For i = 2 To nbCol + 1
        cellBdD = CStr(ws.Cells(2, i).Value)
        If LenB(cellBdD) > 0 Then mapColBdD_Index(cellBdD) = i
    Next i

    ' -- 5. Determiner la derniere ligne de donnees ---------
    '    On scanne la colonne cle metier (table-aware via
    '    SIRUS_ColonneClePourTable) : DTR ou Nom projet selon la table.
    Dim colCle As Long
    colCle = SIRUS_ColonneClePourTable(sTable)
    ligneMax = ws.Cells(ws.rows.count, colCle).End(xlUp).row
    If ligneMax < 3 Then
        MsgBox "La feuille '" & nomFeuille & "' ne contient aucune donnee a analyser.", _
               vbInformation, "SIRUS - Analyse"
        Exit Function
    End If

    ' -- 5 bis. Localiser la colonne Commentaires (col N+2, marquee
    '   "_commentaires" en ligne 2). Permet d ecrire le detail des alertes
    '   directement dans la feuille.
    Dim colCommentaires As Long
    colCommentaires = TrouverColonneCommentaires(ws)

    ' -- 5 ter. Reset de la collection structuree des alertes -----
    Set sirusAlertes = New Collection
    Dim sCle As String

    ' -- 6. Boucle d'analyse par ligne ----------------------
    totNum = 0
    totDen = 0
    For ligne = 3 To ligneMax
        ' Construire le Dictionary {ColonneBdD -> valeur} de la ligne
        Set ligneDict = LireLigneEnDictionary(ws, ligne, mapColBdD_Index)

        ' Cle metier de la ligne (table-aware) :
        '   ProjetASL : col 4 (DTR)
        '   Produit   : col 3 (DTR, apres Famille Generique)
        '   autres    : col 2 (DTR ou Nom projet pour ProjetSLI)
        sCle = SafeCStrAnalyse(ws.Cells(ligne, colCle).Value)

        ' Skip ligne vide (toutes valeurs vides)
        ligneVide = EstLigneVide(ligneDict)
        If ligneVide Then
            ws.Cells(ligne, colStatut).Value = ""
            If colCommentaires > 0 Then ws.Cells(ligne, colCommentaires).Value = ""
            GoTo NextLigne
        End If

        ' Evaluer les regles pour chaque champ du config
        sStatutLigne = ""
        nAlertesLigne = 0
        poidsLigneNum = 0
        poidsLigneDen = 0

        For Each def In col
            Call EvaluerChampLigne(def, sTable, ligneDict, _
                                   sCle, ligne, mapColBdD_Index, _
                                   sStatutLigne, nAlertesLigne, _
                                   poidsLigneNum, poidsLigneDen)
        Next def

        ' Ecrire le Statut court en col 1 + detail en commentaire de cellule
        ' (survol). Format simple : OK (vert) / NOK (rouge avec n alertes).
        Call EcrireStatutCellule(ws.Cells(ligne, colStatut), nAlertesLigne, sStatutLigne)

        ' Ecrire le detail des alertes dans la colonne Commentaires
        ' (visible inline, en plus du survol sur Statut). Vide si OK.
        ' Format : 1 alerte = texte simple. N alertes (N > 1) = liste a
        ' puces avec "- " devant chaque ligne, separees par vbLf.
        If colCommentaires > 0 Then
            If nAlertesLigne = 0 Then
                ws.Cells(ligne, colCommentaires).Value = ""
            Else
                ws.Cells(ligne, colCommentaires).Value = _
                    FormaterAlertesCommentaire(sStatutLigne, nAlertesLigne)
                ws.Cells(ligne, colCommentaires).WrapText = True
            End If
        End If

        totNum = totNum + poidsLigneNum
        totDen = totDen + poidsLigneDen

NextLigne:
    Next ligne

    ' -- 7. Taux MeQ global ---------------------------------
    If totDen > 0 Then
        SIRUS_AnalyserFeuille = totNum / totDen
    Else
        SIRUS_AnalyserFeuille = 1#
    End If

    MsgBox "Analyse '" & nomFeuille & "' terminee." & vbCrLf & vbCrLf & _
           "Taux de completude (MeQ) : " & Format(SIRUS_AnalyserFeuille * 100, "0.0") & "%" & vbCrLf & _
           "Lignes analysees : " & (ligneMax - 2), _
           vbInformation, "SIRUS - Analyse"
    Exit Function

Sortie:
    MsgBox "Erreur SIRUS_AnalyserFeuille : " & Err.description, vbCritical
End Function

' =========================================================================
' EvaluerChampLigne (helper interne)
' =========================================================================
' Applique les 3 controles SIRUS sur un champ d'une ligne :
'   - obligatoire (config Oui OU regle conditionnelle)
'   - regle d'alerte (coherence metier)
'   - cumul poids MeQ (pour le taux global)
' Modifie en sortie le statut ligne + compteurs (passes ByRef) ET pousse
' chaque alerte dans la collection structuree sirusAlertes (UI consume).
Private Sub EvaluerChampLigne(ByVal def As SIRUS_Champ_Config, _
                              ByVal sTable As String, _
                              ByRef ligneDict As Object, _
                              ByVal sCle As String, _
                              ByVal nLigne As Long, _
                              ByRef mapColBdD_Index As Object, _
                              ByRef sStatutLigne As String, _
                              ByRef nAlertes As Long, _
                              ByRef poidsNum As Double, _
                              ByRef poidsDen As Double)

    Dim sVal As String
    Dim bRempli As Boolean
    Dim bObligConfig As Boolean
    Dim bObligRegle As Boolean
    Dim sAlerte As String
    Dim colChamp As Long

    ' Champs sans ColonneBdD (ex. infos pures heritees, MTBF intrins
    ' note dans ProjetASL) : pas evalues.
    If LenB(def.ColonneBdD) = 0 Then Exit Sub

    ' Champs Auto/Herite : pas de controle obligatoire cote utilisateur
    ' (rempli par le moteur ou herite d'une autre table).
    If def.Source = "Auto" Or def.Source = "Herite" Then Exit Sub

    sVal = ""
    If ligneDict.Exists(def.ColonneBdD) Then
        sVal = CStr(ligneDict(def.ColonneBdD))
    End If
    bRempli = (LenB(Trim$(sVal)) > 0)

    bObligConfig = (def.Obligatoire = "Oui")
    bObligRegle = False
    If def.Obligatoire = "Conditionnel" Then
        bObligRegle = SIRUS_RegleObligatoire(sTable, def.ColonneBdD, ligneDict)
    End If

    ' Numero de colonne du champ dans la feuille (pour goto cellule depuis UI)
    colChamp = 0
    If mapColBdD_Index.Exists(def.ColonneBdD) Then
        colChamp = CLng(mapColBdD_Index(def.ColonneBdD))
    End If

    ' Cas 1 : Champ obligatoire manquant
    If (bObligConfig Or bObligRegle) And Not bRempli Then
        nAlertes = nAlertes + 1
        sStatutLigne = AjouterAlerte(sStatutLigne, "Obligatoire manquant '" & def.Champ & "'")
        sirusAlertes.Add Array(sCle, def.Champ, "Obligatoire manquant", nLigne, colChamp)
    End If

    ' Cas 2 : Regle d'alerte metier
    sAlerte = SIRUS_RegleAlerte(sTable, def.ColonneBdD, ligneDict)
    If LenB(sAlerte) > 0 Then
        nAlertes = nAlertes + 1
        sStatutLigne = AjouterAlerte(sStatutLigne, def.Champ & ": " & sAlerte)
        sirusAlertes.Add Array(sCle, def.Champ, sAlerte, nLigne, colChamp)
    End If

    ' Cas 3 : Cumul poids MeQ
    '    Denominateur = somme PoidsMeQ des champs obligatoires (oui + cond.).
    '    Numerateur = somme PoidsMeQ des champs obligatoires effectivement remplis.
    If (bObligConfig Or bObligRegle) And def.PoidsMeQ > 0 Then
        poidsDen = poidsDen + def.PoidsMeQ
        If bRempli Then poidsNum = poidsNum + def.PoidsMeQ
    End If
End Sub

' Formate les alertes pour la colonne Commentaires :
'   nAlertes = 1 -> texte simple (pas de puce, plus lisible pour 1 ligne)
'   nAlertes > 1 -> "- alerte1" & vbLf & "- alerte2" & vbLf & ...
' sStatutLigne est la chaine concatenee par AjouterAlerte (separateur " | ").
Private Function FormaterAlertesCommentaire(ByVal sStatutLigne As String, _
                                             ByVal nAlertes As Long) As String
    Dim arr() As String
    Dim i As Long
    Dim sOut As String

    If nAlertes <= 1 Then
        FormaterAlertesCommentaire = sStatutLigne
        Exit Function
    End If

    arr = Split(sStatutLigne, " | ")
    For i = LBound(arr) To UBound(arr)
        If LenB(arr(i)) > 0 Then
            If LenB(sOut) > 0 Then sOut = sOut & vbLf
            sOut = sOut & "- " & arr(i)
        End If
    Next i
    FormaterAlertesCommentaire = sOut
End Function

' Recherche la colonne marquee "_commentaires" en ligne 2 (entete BdD).
' Renvoie 0 si pas trouve (ne casse pas l analyse).
Private Function TrouverColonneCommentaires(ByVal ws As Worksheet) As Long
    Dim c As Long
    Dim lastCol As Long
    On Error Resume Next
    lastCol = ws.Cells(2, ws.Columns.count).End(xlToLeft).Column
    If lastCol < 1 Then lastCol = 1
    On Error GoTo 0
    For c = 1 To lastCol
        If CStr(ws.Cells(2, c).Value) = "_commentaires" Then
            TrouverColonneCommentaires = c
            Exit Function
        End If
    Next c
    TrouverColonneCommentaires = 0
End Function

' Variante locale de SafeCStr pour ce module (le SafeCStr public est dans
' G_SIRUS_UserForm, on garde l independance).
Private Function SafeCStrAnalyse(ByVal v As Variant) As String
    On Error Resume Next
    If IsNull(v) Or IsEmpty(v) Or IsError(v) Then
        SafeCStrAnalyse = ""
    Else
        SafeCStrAnalyse = CStr(v)
    End If
    If Err.Number <> 0 Then
        SafeCStrAnalyse = ""
        Err.Clear
    End If
    On Error GoTo 0
End Function

Private Function AjouterAlerte(ByVal sStatut As String, ByVal sNouvelle As String) As String
    If LenB(sStatut) = 0 Then
        AjouterAlerte = sNouvelle
    Else
        AjouterAlerte = sStatut & " | " & sNouvelle
    End If
End Function

' Ecrit le statut court (OK / NOK) en couleur dans la cellule, et place
' le detail des alertes en commentaire de cellule (affiche au survol).
Private Sub EcrireStatutCellule(ByVal cellStatut As Range, _
                                 ByVal nAlertes As Long, _
                                 ByVal sDetail As String)
    ' Reset du commentaire existant
    On Error Resume Next
    cellStatut.ClearComments
    On Error GoTo 0

    If nAlertes = 0 Then
        cellStatut.Value = "OK"
        cellStatut.Font.Color = RGB(39, 80, 10)        ' vert
        cellStatut.Interior.Color = RGB(220, 240, 220) ' fond vert clair
    Else
        cellStatut.Value = "NOK (" & nAlertes & ")"
        cellStatut.Font.Color = RGB(163, 45, 45)       ' rouge
        cellStatut.Interior.Color = RGB(255, 220, 220) ' fond rouge clair
        ' Detail des alertes en commentaire de cellule (survol)
        If LenB(sDetail) > 0 Then
            cellStatut.AddComment Replace(sDetail, " | ", vbLf)
        End If
    End If
End Sub

Private Function LireLigneEnDictionary(ByVal ws As Worksheet, _
                                        ByVal ligne As Long, _
                                        ByRef mapColBdD_Index As Object) As Object
    ' Construit le Dictionary {ColonneBdD -> valeur} pour une ligne donnee.
    Dim d As Object
    Dim k As Variant
    Set d = CreateObject("Scripting.Dictionary")
    For Each k In mapColBdD_Index.Keys
        d(CStr(k)) = ws.Cells(ligne, CLng(mapColBdD_Index(k))).Value
    Next k
    Set LireLigneEnDictionary = d
End Function

Private Function EstLigneVide(ByRef ligneDict As Object) As Boolean
    Dim k As Variant
    For Each k In ligneDict.Keys
        If LenB(CStr(ligneDict(k))) > 0 Then
            EstLigneVide = False
            Exit Function
        End If
    Next k
    EstLigneVide = True
End Function

' =========================================================================
' SIRUS_MeQLigne : calcul unifie du MeQ pour une ligne
' =========================================================================
' Source unique de verite pour le calcul MeQ. Utilise par :
'   - G_SIRUS_PBS.CalculerMeQDTR_Optim (ListView Generation)
'   - (futur) G_SIRUS_Analyse.SIRUS_AnalyserFeuille (actuellement helper
'     interne EvaluerChampLigne : a basculer en Phase 2.1 pour alignement).
'
' Formule : MeQ = nbRempli / nbObl
'   nbObl = nombre de champs obligatoires POUR CETTE LIGNE (variable selon
'           les regles conditionnelles ex. matrice Produit v0.2 : 3/14..19)
'   nbRempli = nombre de ces obligatoires effectivement remplis
'
' Comptage 1-champ = 1-unite (PoidsMeQ ignore : la ponderation n'est pas
' active pour l'instant, decision Thomas Q5 2026-07-07).
'
' Skip :
'   - champs sans ColonneBdD (infos pures, heritages du config)
'   - placeholders ??_*_a_definir
'   - Source in {Auto, Herite} (rempli automatiquement ou herite -> pas
'     concerne par la saisie utilisateur)
'   - Source = "SnapShotPrix" est INCLUS dans le comptage (source virtuelle
'     mais compte comme un champ MeQ).
'
' Retour : Dictionary {
'   nbObl       As Long
'   nbRempli    As Long
'   taux        As Double  (nbRempli / nbObl, 1.0 si nbObl = 0)
'   manquants   As Collection (de SIRUS_Champ_Config, les obligatoires vides)
' }
Public Function SIRUS_MeQLigne(ByRef ligneDict As Object, ByVal sTable As String, _
                                ByVal configCol As Collection) As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    result("nbObl") = 0&
    result("nbRempli") = 0&
    result("taux") = 1#

    Dim manquants As Collection
    Set manquants = New Collection

    If configCol Is Nothing Or ligneDict Is Nothing Then
        Set result("manquants") = manquants
        Set SIRUS_MeQLigne = result
        Exit Function
    End If

    Dim def As SIRUS_Champ_Config
    Dim nbObl As Long, nbRempli As Long
    Dim bOblig As Boolean
    Dim sVal As String

    For Each def In configCol
        If LenB(def.ColonneBdD) = 0 Then GoTo NextDef
        If Left$(def.ColonneBdD, 3) = "??_" Then GoTo NextDef
        If def.Source = "Auto" Or def.Source = "Herite" Then GoTo NextDef

        bOblig = False
        Select Case def.Obligatoire
            Case "Oui":          bOblig = True
            Case "Conditionnel": bOblig = SIRUS_RegleObligatoire(sTable, def.ColonneBdD, ligneDict)
        End Select

        If Not bOblig Then GoTo NextDef

        nbObl = nbObl + 1

        sVal = ""
        If ligneDict.Exists(def.ColonneBdD) Then
            sVal = CStr(ligneDict(def.ColonneBdD))
        End If
        If LenB(Trim$(sVal)) > 0 Then
            nbRempli = nbRempli + 1
        Else
            manquants.Add def
        End If
NextDef:
    Next def

    result("nbObl") = nbObl
    result("nbRempli") = nbRempli
    If nbObl > 0 Then
        result("taux") = nbRempli / nbObl
    Else
        result("taux") = 1#
    End If
    Set result("manquants") = manquants
    Set SIRUS_MeQLigne = result
End Function

' =========================================================================
' TESTS Phase 2
' =========================================================================

Public Sub Test_AnalyserFeuille_Produit_Creer()
    Dim taux As Double
    taux = SIRUS_AnalyserFeuille("Produit", "Creer")
    Debug.Print "Taux MeQ Produit Creer = " & Format(taux * 100, "0.0") & "%"
End Sub

Public Sub Test_AnalyserFeuille_TachesATM_Creer()
    Dim taux As Double
    taux = SIRUS_AnalyserFeuille("TachesATM", "Creer")
    Debug.Print "Taux MeQ TachesATM Creer = " & Format(taux * 100, "0.0") & "%"
End Sub


