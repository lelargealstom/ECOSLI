Attribute VB_Name = "G_PBS_Refresh"
'--------------------------------------------------------------------------------------------------------
'
'                                       Generateur Livrables SLI
'                                       Module G_PBS_Refresh
'                                       Developpement : LE LARGE Thomas (LGM/Alstom)
'                                       Sprint 1 bis : bouton Rafraichir DATA_ASL
'                                       Remplace les formules a liens externes par des
'                                       valeurs pures ecrites via ADODB.
'                                       Juin 2026
'
'--------------------------------------------------------------------------------------------------------
'
' Architecture :
'   - SIRUS_RafraichirDATA_ASL (point d entree public)
'   - 4 sub-routines de refresh par bloc (Produit, HARD/SOFT, InfoATM, Prix, Archive)
'   - Helpers : chargement DTRs, cellules projet/date, traces _RefreshLog,
'     deprotection / reprotection feuille
'
' Strategie SQL : pas de clause WHERE [DTR] IN (...) lourde -> on charge
' tout le RecordSet de la table BdD une fois, on filtre cote VBA en
' parcourant le RecordSet et en regardant si chaque DTR appartient au dict
' de DATA_ASL. Plus robuste que des grosses IN, et le RecordSet n a typi-
' quement que quelques milliers de lignes.
'
' La feuille DATA_ASL est protegee (mdp LGM69000, UserInterfaceOnly=True).
' On deprotege au debut, on reprotege a la fin (UserInterfaceOnly ne
' persiste pas a la sauvegarde, donc geste defensif).
'
'--------------------------------------------------------------------------------------------------------
Option Explicit
Option Compare Text

' =========================================================================
' CONSTANTES
' =========================================================================
Private Const SIRUS_PROTECTION_DATA_ASL As String = "LGM69000"
Private Const SHEET_DATA_ASL            As String = "DATA_ASL"
Private Const LISTOBJ_DATA_ASL          As String = "Providor"
Private Const REFRESH_LOG_SHEET         As String = "_RefreshLog"

' Mapping libelles BdD -> indices colonnes DATA_ASL (1-based, dans le LO).
' Source unique de verite : si on renomme une colonne du LO, on modifie ici.
Private Const COL_DTR                   As Long = 1
Private Const COL_DESIGNATION_PBS       As Long = 2
Private Const COL_SOUS_SYSTEME          As Long = 3
Private Const COL_QTE                   As Long = 4
Private Const COL_DESIGNATION_BDD       As Long = 5
Private Const COL_DESIGNATION_PROVIDOR  As Long = 6
Private Const COL_NODE_TYPED            As Long = 7
Private Const COL_HARD_SOFT             As Long = 8
Private Const COL_NTI                   As Long = 9
Private Const COL_REF_PLAN              As Long = 10
Private Const COL_MTBF_INTR             As Long = 11
Private Const COL_MTBF_FONCT            As Long = 12
Private Const COL_MTBF_GEN              As Long = 13
Private Const COL_TX_DEF_FONCT          As Long = 14
Private Const COL_TX_DEF_GEN            As Long = 15
Private Const COL_DUREE_TACHE_HARD_GEN  As Long = 16
Private Const COL_DUREE_TACHE_HARD_ANA  As Long = 17
Private Const COL_UNITE_MAINT_CORR      As Long = 18
Private Const COL_DVL_HARD_GEN          As Long = 19
Private Const COL_DVL_HARD_SPEC         As Long = 20
Private Const COL_DVL_SOFT_GEN          As Long = 21
Private Const COL_DVL_SOFT_SPEC         As Long = 22
Private Const COL_DUREE_TACHE_SOFT_GEN  As Long = 23
Private Const COL_DUREE_TACHE_SOFT_ANA  As Long = 24
Private Const COL_UNITE_MAINT_PREV      As Long = 25
Private Const COL_PRIX_FCS              As Long = 26
Private Const COL_PRIX_SPARE_CAT        As Long = 27
Private Const COL_PRIX_SPARE_CALC       As Long = 28
Private Const COL_REF_SOURCE_SPARE      As Long = 29
Private Const COL_LEAD_TIME_SPARE       As Long = 30
Private Const COL_PRIX_REPAIR_CAT       As Long = 31
Private Const COL_REF_SOURCE_REPAIR     As Long = 32
Private Const COL_LEAD_TIME_REPAIR      As Long = 33
Private Const COL_COUT_SPARE_FRS        As Long = 34
Private Const COL_LEAD_TIME_SPARE_FRS   As Long = 35
Private Const COL_COUT_REPAIR_FRS       As Long = 36
Private Const COL_LEAD_TIME_REPAIR_FRS  As Long = 37
Private Const COL_COUT_SPARE_ALSTOM     As Long = 38
Private Const COL_PRIX_REPAIR_ALSTOM    As Long = 39
Private Const COL_PRIX_SPARE_ARCHIVE    As Long = 40
Private Const COL_PRIX_REPAIR_ARCHIVE   As Long = 41

' =========================================================================
' POINT D ENTREE PUBLIC (cable au bouton Btn_Rafraichir_DATA_ASL)
' =========================================================================
Public Sub SIRUS_RafraichirDATA_ASL(Optional ByVal wbParam As Workbook = Nothing)
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim dictDtrs As Object
    Dim sProjet As String
    Dim tStart As Double
    Dim sErreur As String

    ' -- 1. Localiser le classeur cible -----------------------
    If wbParam Is Nothing Then Set wb = ActiveWorkbook Else Set wb = wbParam
    If wb Is Nothing Then
        MsgBox "Aucun classeur actif.", vbExclamation, "Rafraichir DATA_ASL"
        Exit Sub
    End If

    ' -- 2. Localiser la feuille DATA_ASL + ListObject --------
    On Error Resume Next
    Set ws = wb.Worksheets(SHEET_DATA_ASL)
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "Feuille '" & SHEET_DATA_ASL & "' introuvable dans " & wb.Name & ".", _
               vbCritical, "Rafraichir DATA_ASL"
        Exit Sub
    End If

    On Error Resume Next
    Set lo = ws.ListObjects(LISTOBJ_DATA_ASL)
    On Error GoTo 0
    If lo Is Nothing Then
        MsgBox "ListObject '" & LISTOBJ_DATA_ASL & "' introuvable dans " & SHEET_DATA_ASL & ".", _
               vbCritical, "Rafraichir DATA_ASL"
        Exit Sub
    End If

    ' -- 3. Projet courant depuis UserForm Project_List -------
    sProjet = LireProjetCourant()
    If LenB(sProjet) = 0 Then
        MsgBox "Selectionner d abord un projet dans la liste principale du UserForm ECOSLI" & vbCrLf & _
               "(Ctrl+W pour ouvrir le UserForm).", _
               vbExclamation, "Rafraichir DATA_ASL"
        Exit Sub
    End If

    ' -- 4. Charger les DTRs de DATA_ASL ----------------------
    Set dictDtrs = ChargerDtrsPBS(lo)
    If dictDtrs.count = 0 Then
        MsgBox "Aucun DTR trouve dans DATA_ASL.", vbInformation, "Rafraichir DATA_ASL"
        Exit Sub
    End If

    ' -- 5. Lancer le refresh ---------------------------------
    tStart = Timer
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Call DeProtegerFeuille(ws, wb)

    On Error GoTo CleanUp
    sErreur = ""

    Call RafraichirDepuisProduit(lo, dictDtrs)
    Call RafraichirHARDSOFT(lo, dictDtrs)
    Call RafraichirDepuisInfoATM(lo, dictDtrs)
    Call RafraichirConstantes(lo, dictDtrs)
    Call RafraichirPrixCatalogue(lo, dictDtrs)
    Call RafraichirPrixArchive(lo, dictDtrs, sProjet)
    Call EcrireCellulesProjetDate(wb, sProjet)
    Call TracerRefreshLog(wb, sProjet, dictDtrs.count, Timer - tStart)

CleanUp:
    If Err.Number <> 0 Then sErreur = Err.description
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Call ProtegerFeuille(ws, wb)

    If LenB(sErreur) > 0 Then
        MsgBox "Erreur Rafraichir DATA_ASL :" & vbCrLf & vbCrLf & sErreur, _
               vbCritical, "Rafraichir DATA_ASL"
    Else
        MsgBox "Rafraichissement termine." & vbCrLf & vbCrLf & _
               dictDtrs.count & " DTR(s) traite(s)." & vbCrLf & _
               "Projet : " & sProjet & vbCrLf & _
               "Duree : " & Format$(Timer - tStart, "0.0") & " s.", _
               vbInformation, "Rafraichir DATA_ASL"
    End If
End Sub

' =========================================================================
' LECTURE PROJET COURANT
' =========================================================================
Private Function LireProjetCourant() As String
    ' Recupere le projet depuis UserForm.Project_List si le UserForm est
    ' charge, sinon retourne "" (l appelant bloque).
    On Error Resume Next
    LireProjetCourant = Trim$(CStr(UserForm.Project_List.Value))
    If Err.Number <> 0 Then LireProjetCourant = ""
    On Error GoTo 0
End Function

' =========================================================================
' CHARGEMENT DES DTRs DE DATA_ASL
' =========================================================================
Private Function ChargerDtrsPBS(ByVal lo As ListObject) As Object
    ' Renvoie un Dictionary { DTR_UCase -> row_index dans le DataBodyRange (1-based) }.
    ' Permet ensuite d ecrire en O(1) la valeur correspondante.
    Dim dict As Object
    Dim col As Range
    Dim ligne As Long
    Dim sDTR As String

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    If lo.DataBodyRange Is Nothing Then
        Set ChargerDtrsPBS = dict
        Exit Function
    End If

    Set col = lo.ListColumns(COL_DTR).DataBodyRange
    For ligne = 1 To col.Rows.count
        sDTR = Trim$(NzS(col.Cells(ligne, 1).Value))
        If LenB(sDTR) > 0 Then
            If Not dict.Exists(UCase$(sDTR)) Then dict.Add UCase$(sDTR), ligne
        End If
    Next ligne

    Set ChargerDtrsPBS = dict
End Function

' =========================================================================
' BLOC 1 : RAFRAICHIR DEPUIS Produit$
' =========================================================================
' Colonnes ecrites : Designation BdD (E), Node typed (G), Ref plan (J),
'                    MTBF generique (M), DVL HARD generique (S), DVL SOFT generique (U)
' SELECT * sur Produit$ et lecture champ par champ tolerante (cf. pattern
' G_SIRUS_PBS.ChargerRowsTableBdD_Optim).
Private Sub RafraichirDepuisProduit(ByVal lo As ListObject, ByVal dictDtrs As Object)
    Dim cnx As String
    Dim sDTR As String, sUC As String
    Dim ligne As Long
    Dim colDesi As Range, colNode As Range, colPlan As Range
    Dim colMTBFGen As Range, colDVLHardGen As Range, colDVLSoftGen As Range

    cnx = Replace(CONNEXION_EXCEL_BDD_TRAVAIL, "UserProfile", Environ$("UserProfile"))
    On Error GoTo Sortie
    Call Open_CONN(cnx)
    Call open_REQSEL("SELECT * FROM [Produit$]")

    If REQSEL Is Nothing Then GoTo Sortie
    If REQSEL.EOF And REQSEL.BOF Then GoTo Sortie

    Set colDesi = lo.ListColumns(COL_DESIGNATION_BDD).DataBodyRange
    Set colNode = lo.ListColumns(COL_NODE_TYPED).DataBodyRange
    Set colPlan = lo.ListColumns(COL_REF_PLAN).DataBodyRange
    Set colMTBFGen = lo.ListColumns(COL_MTBF_GEN).DataBodyRange
    Set colDVLHardGen = lo.ListColumns(COL_DVL_HARD_GEN).DataBodyRange
    Set colDVLSoftGen = lo.ListColumns(COL_DVL_SOFT_GEN).DataBodyRange

    Do While Not REQSEL.EOF
        sDTR = Trim$(NzS(GetFieldSafe(REQSEL, "DTR")))
        If LenB(sDTR) > 0 Then
            sUC = UCase$(sDTR)
            If dictDtrs.Exists(sUC) Then
                ligne = CLng(dictDtrs(sUC))
                colDesi.Cells(ligne, 1).Value = NzS(GetFieldSafe(REQSEL, "Désignation_PRODUIT"))
                colNode.Cells(ligne, 1).Value = NzS(GetFieldSafe(REQSEL, "Node type Produit (LRU/SRU…)_PRODUIT"))
                colPlan.Cells(ligne, 1).Value = NzS(GetFieldSafe(REQSEL, "Drawing number"))
                colMTBFGen.Cells(ligne, 1).Value = GetFieldSafe(REQSEL, "MTBF intrinsèque Produit (h) issu de la datasheet _PRODUIT")
                colDVLHardGen.Cells(ligne, 1).Value = GetFieldSafe(REQSEL, "Durée de vie générique Produit HARDWARE (année)_PRODUIT")
                colDVLSoftGen.Cells(ligne, 1).Value = GetFieldSafe(REQSEL, "Durée de vie générique Produit (année) SOFTWARE _PRODUIT")
            End If
        End If
        REQSEL.MoveNext
    Loop

Sortie:
    On Error Resume Next
    Call Close_CONN(cnx)
    On Error GoTo 0
End Sub

' =========================================================================
' BLOC 2 : RAFRAICHIR HARD/SOFT (col H)
' =========================================================================
' Agregat depuis Taches_ATM$ : pour chaque DTR, compter HARD et SOFT
' separement. Resultat : "BOTH" / "HARD" / "SOFT" / "".
Private Sub RafraichirHARDSOFT(ByVal lo As ListObject, ByVal dictDtrs As Object)
    Dim cnx As String
    Dim sDTR As String, sUC As String, sHS As String
    Dim ligne As Long
    Dim colHS As Range
    Dim dictAgreg As Object   ' { DTR_UC -> "H" / "S" / "HS" }

    cnx = Replace(CONNEXION_EXCEL_BDD_TRAVAIL, "UserProfile", Environ$("UserProfile"))
    On Error GoTo Sortie
    Call Open_CONN(cnx)
    Call open_REQSEL("SELECT [DTR_TACHEATM], [HARD/SOFT] FROM [Taches_ATM$]")

    If REQSEL Is Nothing Then GoTo Sortie
    If REQSEL.EOF And REQSEL.BOF Then GoTo Sortie

    Set dictAgreg = CreateObject("Scripting.Dictionary")
    dictAgreg.CompareMode = vbTextCompare

    Do While Not REQSEL.EOF
        sDTR = Trim$(NzS(REQSEL.Fields("DTR_TACHEATM").Value))
        sHS = UCase$(Trim$(NzS(REQSEL.Fields("HARD/SOFT").Value)))
        If LenB(sDTR) > 0 Then
            sUC = UCase$(sDTR)
            If dictDtrs.Exists(sUC) Then
                Dim existing As String
                existing = ""
                If dictAgreg.Exists(sUC) Then existing = CStr(dictAgreg(sUC))
                If sHS = "HARD" Then
                    If existing = "" Then dictAgreg(sUC) = "H" _
                    ElseIf existing = "S" Then dictAgreg(sUC) = "HS"
                ElseIf sHS = "SOFT" Then
                    If existing = "" Then dictAgreg(sUC) = "S" _
                    ElseIf existing = "H" Then dictAgreg(sUC) = "HS"
                End If
            End If
        End If
        REQSEL.MoveNext
    Loop

    ' Ecrire col H pour chaque DTR
    Set colHS = lo.ListColumns(COL_HARD_SOFT).DataBodyRange
    Dim key As Variant
    For Each key In dictDtrs.Keys
        ligne = CLng(dictDtrs(key))
        If dictAgreg.Exists(key) Then
            Select Case CStr(dictAgreg(key))
                Case "HS": colHS.Cells(ligne, 1).Value = "BOTH"
                Case "H":  colHS.Cells(ligne, 1).Value = "HARD"
                Case "S":  colHS.Cells(ligne, 1).Value = "SOFT"
                Case Else: colHS.Cells(ligne, 1).Value = ""
            End Select
        Else
            colHS.Cells(ligne, 1).Value = ""
        End If
    Next key

Sortie:
    On Error Resume Next
    Call Close_CONN(cnx)
    On Error GoTo 0
End Sub

' =========================================================================
' BLOC 3 : RAFRAICHIR DEPUIS Info_ATM$
' =========================================================================
' Colonne P (Duree tache corr+prev HARD generique) : Temps generique de la
'   tache_INFOATM par (DTR, HARD/SOFT='HARD') -- premier match.
' Colonnes Q (HARD analytique) et X (SOFT analytique) : SUM(Temps
'   analytique cumule de la tache_INFOATM) par (DTR, HARD/SOFT).
Private Sub RafraichirDepuisInfoATM(ByVal lo As ListObject, ByVal dictDtrs As Object)
    Dim cnx As String
    Dim sDTR As String, sUC As String, sHS As String
    Dim ligne As Long
    Dim colP As Range, colQ As Range, colX As Range
    Dim dictGenHard As Object  ' { DTR_UC -> tempsGenHard }
    Dim dictAnaHard As Object  ' { DTR_UC -> SUM tempsAnaHard }
    Dim dictAnaSoft As Object  ' { DTR_UC -> SUM tempsAnaSoft }
    Dim vTempsGen As Variant, vTempsAna As Variant

    cnx = Replace(CONNEXION_EXCEL_BDD_TRAVAIL, "UserProfile", Environ$("UserProfile"))
    On Error GoTo Sortie
    Call Open_CONN(cnx)
    Call open_REQSEL("SELECT [DTR_INFOATM], [HARD/SOFT], " & _
                     "[Temps générique de la tâche_INFOATM], " & _
                     "[Temps analytique cumulé de la tache_INFOATM] " & _
                     "FROM [Info_ATM$]")

    If REQSEL Is Nothing Then GoTo Sortie
    If REQSEL.EOF And REQSEL.BOF Then GoTo Sortie

    Set dictGenHard = CreateObject("Scripting.Dictionary")
    Set dictAnaHard = CreateObject("Scripting.Dictionary")
    Set dictAnaSoft = CreateObject("Scripting.Dictionary")
    dictGenHard.CompareMode = vbTextCompare
    dictAnaHard.CompareMode = vbTextCompare
    dictAnaSoft.CompareMode = vbTextCompare

    Do While Not REQSEL.EOF
        sDTR = Trim$(NzS(REQSEL.Fields("DTR_INFOATM").Value))
        sHS = UCase$(Trim$(NzS(REQSEL.Fields("HARD/SOFT").Value)))
        If LenB(sDTR) > 0 Then
            sUC = UCase$(sDTR)
            If dictDtrs.Exists(sUC) Then
                vTempsGen = REQSEL.Fields("Temps générique de la tâche_INFOATM").Value
                vTempsAna = REQSEL.Fields("Temps analytique cumulé de la tache_INFOATM").Value

                ' Premier match HARD pour Temps generique (col P)
                If sHS = "HARD" Then
                    If Not dictGenHard.Exists(sUC) Then
                        If IsNumeric(vTempsGen) Then dictGenHard(sUC) = CDbl(vTempsGen)
                    End If
                End If

                ' Cumul analytique
                If IsNumeric(vTempsAna) Then
                    If sHS = "HARD" Then
                        dictAnaHard(sUC) = NzD(dictAnaHard, sUC) + CDbl(vTempsAna)
                    ElseIf sHS = "SOFT" Then
                        dictAnaSoft(sUC) = NzD(dictAnaSoft, sUC) + CDbl(vTempsAna)
                    End If
                End If
            End If
        End If
        REQSEL.MoveNext
    Loop

    ' Ecrire les 3 colonnes
    Set colP = lo.ListColumns(COL_DUREE_TACHE_HARD_GEN).DataBodyRange
    Set colQ = lo.ListColumns(COL_DUREE_TACHE_HARD_ANA).DataBodyRange
    Set colX = lo.ListColumns(COL_DUREE_TACHE_SOFT_ANA).DataBodyRange

    Dim key As Variant
    For Each key In dictDtrs.Keys
        ligne = CLng(dictDtrs(key))
        If dictGenHard.Exists(key) Then
            colP.Cells(ligne, 1).Value = dictGenHard(key)
        Else
            colP.Cells(ligne, 1).Value = Empty
        End If
        If dictAnaHard.Exists(key) Then
            colQ.Cells(ligne, 1).Value = dictAnaHard(key)
        Else
            colQ.Cells(ligne, 1).Value = Empty
        End If
        If dictAnaSoft.Exists(key) Then
            colX.Cells(ligne, 1).Value = dictAnaSoft(key)
        Else
            colX.Cells(ligne, 1).Value = Empty
        End If
    Next key

Sortie:
    On Error Resume Next
    Call Close_CONN(cnx)
    On Error GoTo 0
End Sub

' =========================================================================
' BLOC 4 : ECRIRE LES CONSTANTES PH / Y
' =========================================================================
' Col R (UNITE maintenance corrective) = "PH" en dur.
' Col Y (UNITE maintenance preventive) = "Y" en dur.
' Toujours, pour toutes les lignes ayant un DTR.
Private Sub RafraichirConstantes(ByVal lo As ListObject, ByVal dictDtrs As Object)
    Dim colR As Range, colY As Range
    Dim key As Variant
    Dim ligne As Long

    Set colR = lo.ListColumns(COL_UNITE_MAINT_CORR).DataBodyRange
    Set colY = lo.ListColumns(COL_UNITE_MAINT_PREV).DataBodyRange

    For Each key In dictDtrs.Keys
        ligne = CLng(dictDtrs(key))
        colR.Cells(ligne, 1).Value = "PH"
        colY.Cells(ligne, 1).Value = "Y"
    Next key
End Sub

' =========================================================================
' BLOC 5 : RAFRAICHIR LES PRIX CATALOGUE depuis SnapShotPrix$
' =========================================================================
' SELECT * FROM [SnapShotPrix$]
' Pour chaque DTR de DATA_ASL, on garde la ligne de date la plus recente
' pour chaque combo (Type, Ref Catalogue contient FCS ou D&IS Catalogue) :
'   Spare + FCS                 -> col Z (Prix FCS)
'   Spare + D&IS Catalogue      -> col AA (Prix Spare Catalogue),
'                                  col AC (Ref Source Spare),
'                                  col AD (Lead Time SPARE)
'   Repair + D&IS Catalogue     -> col AE (Prix Repair Catalogue),
'                                  col AF (Ref Source Repair),
'                                  col AG (Lead Time REPAIR)
Private Sub RafraichirPrixCatalogue(ByVal lo As ListObject, ByVal dictDtrs As Object)
    Dim cnx As String
    Dim sDTR As String, sUC As String, sType As String, sRefCat As String
    Dim sTypeRefUC As String
    Dim vDate As Variant, vPrix As Variant, vLT As Variant
    Dim dictBest As Object  ' { DTR_UC|"SPARE_FCS"|... -> Array(date, prix, ref, lt) }
    Dim key As Variant
    Dim ligne As Long

    cnx = Replace(CONNEXION_EXCEL_BDD_TRAVAIL, "UserProfile", Environ$("UserProfile"))
    On Error GoTo Sortie
    Call Open_CONN(cnx)
    Call open_REQSEL("SELECT [DTR], [Type], [Ref Catalogue], [Date Source], [Prix], [LeadTime] " & _
                     "FROM [SnapShotPrix$]")

    If REQSEL Is Nothing Then GoTo Sortie
    If REQSEL.EOF And REQSEL.BOF Then GoTo Sortie

    Set dictBest = CreateObject("Scripting.Dictionary")
    dictBest.CompareMode = vbTextCompare

    Do While Not REQSEL.EOF
        sDTR = Trim$(NzS(REQSEL.Fields("DTR").Value))
        If LenB(sDTR) > 0 Then
            sUC = UCase$(sDTR)
            If dictDtrs.Exists(sUC) Then
                sType = UCase$(Trim$(NzS(REQSEL.Fields("Type").Value)))
                sRefCat = UCase$(NzS(REQSEL.Fields("Ref Catalogue").Value))
                vDate = REQSEL.Fields("Date Source").Value
                vPrix = REQSEL.Fields("Prix").Value
                vLT = REQSEL.Fields("LeadTime").Value

                sTypeRefUC = ""
                If sType = "SPARE" Then
                    If InStr(1, sRefCat, "D&IS CATALOGUE", vbTextCompare) > 0 Then
                        sTypeRefUC = "SPARE_DIS"
                    ElseIf InStr(1, sRefCat, "FCS", vbTextCompare) > 0 Then
                        sTypeRefUC = "SPARE_FCS"
                    End If
                ElseIf sType = "REPAIR" Then
                    If InStr(1, sRefCat, "D&IS CATALOGUE", vbTextCompare) > 0 Then
                        sTypeRefUC = "REPAIR_DIS"
                    End If
                End If

                If LenB(sTypeRefUC) > 0 And IsDate(vDate) Then
                    Dim k As String
                    k = sUC & "|" & sTypeRefUC
                    Dim conserver As Boolean
                    conserver = True
                    If dictBest.Exists(k) Then
                        Dim prev As Variant
                        prev = dictBest(k)
                        If CDate(prev(0)) >= CDate(vDate) Then conserver = False
                    End If
                    If conserver Then
                        dictBest(k) = Array(vDate, vPrix, REQSEL.Fields("Ref Catalogue").Value, vLT)
                    End If
                End If
            End If
        End If
        REQSEL.MoveNext
    Loop

    ' Ecriture des 6 colonnes
    Dim colZ As Range, colAA As Range, colAC As Range, colAD As Range
    Dim colAE As Range, colAF As Range, colAG As Range
    Set colZ = lo.ListColumns(COL_PRIX_FCS).DataBodyRange
    Set colAA = lo.ListColumns(COL_PRIX_SPARE_CAT).DataBodyRange
    Set colAC = lo.ListColumns(COL_REF_SOURCE_SPARE).DataBodyRange
    Set colAD = lo.ListColumns(COL_LEAD_TIME_SPARE).DataBodyRange
    Set colAE = lo.ListColumns(COL_PRIX_REPAIR_CAT).DataBodyRange
    Set colAF = lo.ListColumns(COL_REF_SOURCE_REPAIR).DataBodyRange
    Set colAG = lo.ListColumns(COL_LEAD_TIME_REPAIR).DataBodyRange

    For Each key In dictDtrs.Keys
        ligne = CLng(dictDtrs(key))
        ' Prix FCS
        If dictBest.Exists(key & "|SPARE_FCS") Then
            colZ.Cells(ligne, 1).Value = dictBest(key & "|SPARE_FCS")(1)
        Else
            colZ.Cells(ligne, 1).Value = Empty
        End If
        ' Prix Spare Catalogue + Ref + LT
        If dictBest.Exists(key & "|SPARE_DIS") Then
            colAA.Cells(ligne, 1).Value = dictBest(key & "|SPARE_DIS")(1)
            colAC.Cells(ligne, 1).Value = dictBest(key & "|SPARE_DIS")(2)
            colAD.Cells(ligne, 1).Value = dictBest(key & "|SPARE_DIS")(3)
        Else
            colAA.Cells(ligne, 1).Value = Empty
            colAC.Cells(ligne, 1).Value = Empty
            colAD.Cells(ligne, 1).Value = Empty
        End If
        ' Prix Repair Catalogue + Ref + LT
        If dictBest.Exists(key & "|REPAIR_DIS") Then
            colAE.Cells(ligne, 1).Value = dictBest(key & "|REPAIR_DIS")(1)
            colAF.Cells(ligne, 1).Value = dictBest(key & "|REPAIR_DIS")(2)
            colAG.Cells(ligne, 1).Value = dictBest(key & "|REPAIR_DIS")(3)
        Else
            colAE.Cells(ligne, 1).Value = Empty
            colAF.Cells(ligne, 1).Value = Empty
            colAG.Cells(ligne, 1).Value = Empty
        End If
    Next key

Sortie:
    On Error Resume Next
    Call Close_CONN(cnx)
    On Error GoTo 0
End Sub

' =========================================================================
' BLOC 6 : RAFRAICHIR LES PRIX ARCHIVE depuis Projet_ASL$
' =========================================================================
' SELECT * FROM [Projet_ASL$] WHERE [PROJET_PROJETASL] <> '<projet courant>'
' Pour chaque DTR, AVG des valeurs > 0 trouvees dans :
'   [ISP TIS HORS FRANCE (archivé)]                               -> col AN
'   [Prix repair (DEVISE du projet (SLI))_PROJETASL]              -> col AO
Private Sub RafraichirPrixArchive(ByVal lo As ListObject, ByVal dictDtrs As Object, _
                                    ByVal sProjetCourant As String)
    Dim cnx As String
    Dim sDTR As String, sUC As String, sProjet As String
    Dim ligne As Long
    Dim dictSpareSum As Object   ' { DTR_UC -> sum }
    Dim dictSpareCnt As Object   ' { DTR_UC -> count }
    Dim dictRepairSum As Object
    Dim dictRepairCnt As Object
    Dim vSpare As Variant, vRepair As Variant
    Dim key As Variant
    Dim colAN As Range, colAO As Range

    cnx = Replace(CONNEXION_EXCEL_BDD_TRAVAIL, "UserProfile", Environ$("UserProfile"))
    On Error GoTo Sortie
    Call Open_CONN(cnx)
    Call open_REQSEL("SELECT [DTR_PROJETASL], [PROJET_PROJETASL], " & _
                     "[ISP TIS HORS FRANCE (archivé)], " & _
                     "[Prix repair (DEVISE du projet (SLI))_PROJETASL] " & _
                     "FROM [Projet_ASL$]")

    If REQSEL Is Nothing Then GoTo Sortie
    If REQSEL.EOF And REQSEL.BOF Then GoTo Sortie

    Set dictSpareSum = CreateObject("Scripting.Dictionary")
    Set dictSpareCnt = CreateObject("Scripting.Dictionary")
    Set dictRepairSum = CreateObject("Scripting.Dictionary")
    Set dictRepairCnt = CreateObject("Scripting.Dictionary")
    dictSpareSum.CompareMode = vbTextCompare
    dictSpareCnt.CompareMode = vbTextCompare
    dictRepairSum.CompareMode = vbTextCompare
    dictRepairCnt.CompareMode = vbTextCompare

    Do While Not REQSEL.EOF
        sDTR = Trim$(NzS(REQSEL.Fields("DTR_PROJETASL").Value))
        sProjet = Trim$(NzS(REQSEL.Fields("PROJET_PROJETASL").Value))
        If LenB(sDTR) > 0 And UCase$(sProjet) <> UCase$(sProjetCourant) Then
            sUC = UCase$(sDTR)
            If dictDtrs.Exists(sUC) Then
                vSpare = REQSEL.Fields("ISP TIS HORS FRANCE (archivé)").Value
                vRepair = REQSEL.Fields("Prix repair (DEVISE du projet (SLI))_PROJETASL").Value

                If IsNumeric(vSpare) Then
                    If CDbl(vSpare) > 0 Then
                        dictSpareSum(sUC) = NzD(dictSpareSum, sUC) + CDbl(vSpare)
                        dictSpareCnt(sUC) = NzD(dictSpareCnt, sUC) + 1
                    End If
                End If
                If IsNumeric(vRepair) Then
                    If CDbl(vRepair) > 0 Then
                        dictRepairSum(sUC) = NzD(dictRepairSum, sUC) + CDbl(vRepair)
                        dictRepairCnt(sUC) = NzD(dictRepairCnt, sUC) + 1
                    End If
                End If
            End If
        End If
        REQSEL.MoveNext
    Loop

    Set colAN = lo.ListColumns(COL_PRIX_SPARE_ARCHIVE).DataBodyRange
    Set colAO = lo.ListColumns(COL_PRIX_REPAIR_ARCHIVE).DataBodyRange

    For Each key In dictDtrs.Keys
        ligne = CLng(dictDtrs(key))
        If dictSpareCnt.Exists(key) Then
            colAN.Cells(ligne, 1).Value = dictSpareSum(key) / dictSpareCnt(key)
        Else
            colAN.Cells(ligne, 1).Value = Empty
        End If
        If dictRepairCnt.Exists(key) Then
            colAO.Cells(ligne, 1).Value = dictRepairSum(key) / dictRepairCnt(key)
        Else
            colAO.Cells(ligne, 1).Value = Empty
        End If
    Next key

Sortie:
    On Error Resume Next
    Call Close_CONN(cnx)
    On Error GoTo 0
End Sub

' =========================================================================
' ECRITURE CELLULES PROJET + DATE
' =========================================================================
' Cellules nommees DATA_ASL_PROJET et DATA_ASL_DERNIERE_MAJ (cf. spec
' template). Si les cellules nommees n'existent pas, on ecrit en dur en
' I1 et I2.
Private Sub EcrireCellulesProjetDate(ByVal wb As Workbook, ByVal sProjet As String)
    Dim ws As Worksheet
    Set ws = wb.Worksheets(SHEET_DATA_ASL)

    On Error Resume Next
    wb.Names("DATA_ASL_PROJET").RefersToRange.Value = sProjet
    If Err.Number <> 0 Then
        ws.Cells(1, 9).Value = sProjet   ' I1
        Err.Clear
    End If
    wb.Names("DATA_ASL_DERNIERE_MAJ").RefersToRange.Value = _
        Format$(Now(), "yyyy-mm-dd hh:mm")
    If Err.Number <> 0 Then
        ws.Cells(2, 9).Value = Format$(Now(), "yyyy-mm-dd hh:mm")
        Err.Clear
    End If
    ws.Cells(1, 9).Font.Italic = False
    ws.Cells(1, 9).Font.Color = vbBlack
    ws.Cells(2, 9).Font.Italic = False
    ws.Cells(2, 9).Font.Color = vbBlack
    On Error GoTo 0
End Sub

' =========================================================================
' TRACE _RefreshLog
' =========================================================================
' Append d une ligne dans la feuille _RefreshLog : Date / Projet / nbDTR /
' Duree (secondes) / Utilisateur. Cree la feuille si absente.
Private Sub TracerRefreshLog(ByVal wb As Workbook, ByVal sProjet As String, _
                              ByVal nbDtrs As Long, ByVal duree As Double)
    Dim ws As Worksheet
    Dim ligne As Long

    On Error Resume Next
    Set ws = wb.Worksheets(REFRESH_LOG_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.count))
        ws.Name = REFRESH_LOG_SHEET
        ws.Visible = xlSheetVeryHidden
        ws.Range("A1:E1").Value = Array("Date", "Projet", "Nb DTR", "Duree (s)", "Utilisateur")
        ws.Range("A1:E1").Font.Bold = True
    End If

    ligne = ws.Cells(ws.rows.count, 1).End(xlUp).Row + 1
    ws.Cells(ligne, 1).Value = Now()
    ws.Cells(ligne, 1).NumberFormat = "yyyy-mm-dd hh:mm:ss"
    ws.Cells(ligne, 2).Value = sProjet
    ws.Cells(ligne, 3).Value = nbDtrs
    ws.Cells(ligne, 4).Value = Format$(duree, "0.0")
    ws.Cells(ligne, 5).Value = Environ$("USERNAME")
End Sub

' =========================================================================
' DEPROTECTION / REPROTECTION
' =========================================================================
Private Sub DeProtegerFeuille(ByVal ws As Worksheet, ByVal wb As Workbook)
    On Error Resume Next
    wb.Unprotect SIRUS_PROTECTION_DATA_ASL
    ws.Unprotect SIRUS_PROTECTION_DATA_ASL
    On Error GoTo 0
End Sub

Private Sub ProtegerFeuille(ByVal ws As Worksheet, ByVal wb As Workbook)
    On Error Resume Next
    ws.Protect Password:=SIRUS_PROTECTION_DATA_ASL, _
               DrawingObjects:=True, Contents:=True, Scenarios:=True, _
               UserInterfaceOnly:=True, _
               AllowFormattingCells:=False, AllowFormattingColumns:=False, _
               AllowFormattingRows:=False, AllowInsertingColumns:=False, _
               AllowInsertingRows:=True, AllowInsertingHyperlinks:=True, _
               AllowDeletingColumns:=False, AllowDeletingRows:=True, _
               AllowSorting:=True, AllowFiltering:=True, AllowUsingPivotTables:=True
    wb.Protect Password:=SIRUS_PROTECTION_DATA_ASL, Structure:=True
    On Error GoTo 0
End Sub

' =========================================================================
' HELPERS
' =========================================================================

' Lecture safe d un champ ADO : retourne Empty si le champ n existe pas.
Private Function GetFieldSafe(ByVal rs As Object, ByVal sNomChamp As String) As Variant
    On Error Resume Next
    GetFieldSafe = rs.Fields(sNomChamp).Value
    If Err.Number <> 0 Then GetFieldSafe = Empty
    On Error GoTo 0
End Function

' Conversion safe Variant -> String (Null/Empty/Error -> "")
Private Function NzS(ByVal v As Variant) As String
    On Error Resume Next
    If IsNull(v) Or IsEmpty(v) Or IsError(v) Then
        NzS = ""
    Else
        NzS = CStr(v)
    End If
    If Err.Number <> 0 Then NzS = "" : Err.Clear
    On Error GoTo 0
End Function

' Lookup Double dans un dict avec defaut 0.
Private Function NzD(ByVal dict As Object, ByVal key As Variant) As Double
    If dict.Exists(key) Then NzD = CDbl(dict(key)) Else NzD = 0
End Function

' =========================================================================
' RESET CACHE (a appeler apres refresh BdD, non utilise pour l instant)
' =========================================================================
Public Sub SIRUS_ResetCacheRefresh()
    ' Pas de cache module-level dans cette version (rechargement complet
    ' a chaque clic). Stub pour symetrie API future.
End Sub
