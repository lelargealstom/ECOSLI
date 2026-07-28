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
' La feuille DATA_ASL et le workbook ne sont plus proteges du tout
' (demande utilisateur 2026-06-29 : Thomas a deja deprotege l ensemble
' a la main). Plus de DeProtegerFeuille / ProtegerFeuille.
'
'--------------------------------------------------------------------------------------------------------
Option Explicit
Option Compare Text

' =========================================================================
' CONSTANTES
' =========================================================================
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
        MsgBox "Feuille '" & SHEET_DATA_ASL & "' introuvable dans " & wb.Name & ".", vbCritical, "Rafraichir DATA_ASL"
        Exit Sub
    End If

    On Error Resume Next
    Set lo = ws.ListObjects(LISTOBJ_DATA_ASL)
    On Error GoTo 0
    If lo Is Nothing Then
        MsgBox "ListObject '" & LISTOBJ_DATA_ASL & "' introuvable dans " & SHEET_DATA_ASL & ".", vbCritical, "Rafraichir DATA_ASL"
        Exit Sub
    End If

    ' -- 3. Projet courant depuis UserForm Project_List -------
    sProjet = LireProjetCourant()
    If LenB(sProjet) = 0 Then
        MsgBox "Selectionner d abord un projet dans la liste principale du UserForm ECOSLI" & vbCrLf & "(Ctrl+W pour ouvrir le UserForm).", vbExclamation, "Rafraichir DATA_ASL"
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

    Call RafraichirDepuisProduit(lo, dictDtrs)
    Call RafraichirHARDSOFT(lo, dictDtrs)
    Call RafraichirDepuisInfoATM(lo, dictDtrs)
    Call RafraichirConstantes(lo, dictDtrs)
    Call RafraichirPrixCatalogue(lo, dictDtrs)
    Call RafraichirPrixArchive(lo, dictDtrs, sProjet)
    Call RafraichirCalculs(ws, lo, dictDtrs)
    Call EcrireCellulesProjetDate(wb, sProjet)
    Call TracerRefreshLog(wb, sProjet, dictDtrs.count, Timer - tStart)

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    ' Pas de reprotection (decision 2026-06-29 : feuille DATA_ASL libre)

    ' Rafraichir la ListView Generation pour refleter les nouvelles
    ' valeurs DATA_ASL sans que l ingenieur ait a fermer/rouvrir le
    ' UserForm. Wrap en On Error : si la ListView n est pas dispo
    ' (UserForm pas charge), on ignore silencieusement.
    On Error Resume Next
    Call G_SIRUS_PBS.SIRUS_RemplirListViewGeneration
    On Error GoTo 0

    MsgBox "Rafraichissement termine." & vbCrLf & vbCrLf & dictDtrs.count & " DTR(s) traite(s)." & vbCrLf & "Projet : " & sProjet & vbCrLf & "Duree : " & Format$(Timer - tStart, "0.0") & " s.", vbInformation, "Rafraichir DATA_ASL"
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
    Call Open_CONN(cnx)
    ' SELECT explicit + cast DTR en String via "& ''" (force ACE OLE DB
    ' a retourner les DTRs alphanumeriques meme si les 8 premieres lignes
    ' sont numeriques).
    Call open_REQSEL("SELECT [DTR] & '' AS DTR_str, [Designation_PRODUIT], [Node type Produit (LRU/SRU...)_PRODUIT], [Drawing number], [MTBF intrinseque Produit (h) issu de la datasheet _PRODUIT], [Duree de vie generique Produit HARDWARE (annee)_PRODUIT], [Duree de vie generique Produit (annee) SOFTWARE _PRODUIT] FROM [Produit$]")

    If REQSEL Is Nothing Then GoTo Sortie
    If REQSEL.EOF And REQSEL.BOF Then GoTo Sortie

    Set colDesi = lo.ListColumns(COL_DESIGNATION_BDD).DataBodyRange
    Set colNode = lo.ListColumns(COL_NODE_TYPED).DataBodyRange
    Set colPlan = lo.ListColumns(COL_REF_PLAN).DataBodyRange
    Set colMTBFGen = lo.ListColumns(COL_MTBF_GEN).DataBodyRange
    Set colDVLHardGen = lo.ListColumns(COL_DVL_HARD_GEN).DataBodyRange
    Set colDVLSoftGen = lo.ListColumns(COL_DVL_SOFT_GEN).DataBodyRange

    Do While Not REQSEL.EOF
        sDTR = Trim$(NzS(GetFieldSafe(REQSEL, "DTR_str")))
        If LenB(sDTR) > 0 Then
            sUC = UCase$(sDTR)
            If dictDtrs.Exists(sUC) Then
                ligne = CLng(dictDtrs(sUC))
                colDesi.Cells(ligne, 1).Value = NzS(GetFieldSafe(REQSEL, "Designation_PRODUIT"))
                colNode.Cells(ligne, 1).Value = NzS(GetFieldSafe(REQSEL, "Node type Produit (LRU/SRU...)_PRODUIT"))
                colPlan.Cells(ligne, 1).Value = NzS(GetFieldSafe(REQSEL, "Drawing number"))
                colMTBFGen.Cells(ligne, 1).Value = GetFieldSafe(REQSEL, "MTBF intrinseque Produit (h) issu de la datasheet _PRODUIT")
                colDVLHardGen.Cells(ligne, 1).Value = GetFieldSafe(REQSEL, "Duree de vie generique Produit HARDWARE (annee)_PRODUIT")
                colDVLSoftGen.Cells(ligne, 1).Value = GetFieldSafe(REQSEL, "Duree de vie generique Produit (annee) SOFTWARE _PRODUIT")
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
    Call Open_CONN(cnx)
    ' Cast DTR en String (cf. note dans RafraichirDepuisProduit)
    Call open_REQSEL("SELECT [DTR_TACHEATM] & '' AS DTR_str, [HARD/SOFT] FROM [Taches_ATM$]")

    If REQSEL Is Nothing Then GoTo Sortie
    If REQSEL.EOF And REQSEL.BOF Then GoTo Sortie

    Set dictAgreg = CreateObject("Scripting.Dictionary")
    dictAgreg.CompareMode = vbTextCompare

    Do While Not REQSEL.EOF
        sDTR = Trim$(NzS(REQSEL.Fields("DTR_str").Value))
        sHS = UCase$(Trim$(NzS(REQSEL.Fields("HARD/SOFT").Value)))
        If LenB(sDTR) > 0 Then
            sUC = UCase$(sDTR)
            If dictDtrs.Exists(sUC) Then
                Dim existing As String
                existing = ""
                If dictAgreg.Exists(sUC) Then existing = CStr(dictAgreg(sUC))
                If sHS = "HARD" Then
                    Select Case existing
                        Case "":  dictAgreg(sUC) = "H"
                        Case "S": dictAgreg(sUC) = "HS"
                    End Select
                ElseIf sHS = "SOFT" Then
                    Select Case existing
                        Case "":  dictAgreg(sUC) = "S"
                        Case "H": dictAgreg(sUC) = "HS"
                    End Select
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
    Call Open_CONN(cnx)
    ' Cast DTR en String (cf. note dans RafraichirDepuisProduit)
    Call open_REQSEL("SELECT [DTR_INFOATM] & '' AS DTR_str, [HARD/SOFT], [Temps generique de la tache_INFOATM], [Temps analytique cumule de la tache_INFOATM] FROM [Info_ATM$]")

    If REQSEL Is Nothing Then GoTo Sortie
    If REQSEL.EOF And REQSEL.BOF Then GoTo Sortie

    Set dictGenHard = CreateObject("Scripting.Dictionary")
    Set dictAnaHard = CreateObject("Scripting.Dictionary")
    Set dictAnaSoft = CreateObject("Scripting.Dictionary")
    dictGenHard.CompareMode = vbTextCompare
    dictAnaHard.CompareMode = vbTextCompare
    dictAnaSoft.CompareMode = vbTextCompare

    Do While Not REQSEL.EOF
        sDTR = Trim$(NzS(REQSEL.Fields("DTR_str").Value))
        sHS = UCase$(Trim$(NzS(REQSEL.Fields("HARD/SOFT").Value)))
        If LenB(sDTR) > 0 Then
            sUC = UCase$(sDTR)
            If dictDtrs.Exists(sUC) Then
                vTempsGen = REQSEL.Fields("Temps generique de la tache_INFOATM").Value
                vTempsAna = REQSEL.Fields("Temps analytique cumule de la tache_INFOATM").Value

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
    ' Decision Thomas 2026-06-28 : le Refresh ne fait PAS de cascade sur
    ' les PRIX, mais SI une mini-cascade sur le label catalogue AC :
    '   - col Z  Prix FCS              = MAX(Date Source) parmi les lignes FCS
    '   - col AA Prix Spare Catalogue  = MAX(Date) parmi les lignes D&IS (strict)
    '   - col AC Ref Source Spare      = ref D&IS si dispo, sinon ref FCS
    '   - col AE Prix Repair Catalogue = MAX(Date) parmi les lignes D&IS Repair
    '   - col AF Ref Source Repair     = ref D&IS Repair (pas de FCS Repair)
    ' Justification AC en cascade : la ListView Generation cascade Spare
    ' D&IS -> FCS pour le prix (via col AB = Prix FCS x 1,0557) ; elle a
    ' donc besoin du libelle catalogue retenu en aval. Si on laissait AC
    ' vide quand seul FCS est connu, la Source afficherait "FCS" en dur
    ' sans traceability sur la ref Catalogue exacte.
    ' Pas de fallback "AUTRE" : si ni D&IS ni FCS ne sont connus, AA/AC
    ' restent vides et la ListView tombe sur Archive.
    '
    ' IMPORTANT : on N UTILISE PAS ADODB pour cette table. Le bug Microsoft
    ' ACE OLE DB sur .xlsb avec colonne DTR mixed-type (numerique +
    ' alphanumerique) fait que les valeurs alphanumeriques sont retournees
    ' comme NULL (le scan TypeGuessRows precede l evaluation des expressions
    ' donc meme '[DTR] & '''' ne suffit pas). On bypass en ouvrant le
    ' fichier .xlsb dans Excel en read-only et en lisant la feuille en
    ' bulk array.
    Dim sCheminBdD As String
    Dim wbBdD As Workbook, wb As Workbook
    Dim wsBdD As Worksheet
    Dim arr As Variant
    Dim ligne As Long, derLigne As Long
    ' Note : pas de "cDate" (collision avec fonction VBA CDate), ni "cType"
    ' (collision avec mot-cle Type). On prefixe avec colIdx pour clarte.
    Dim colIdxDTR As Long, colIdxType As Long, colIdxRef As Long
    Dim colIdxDate As Long, colIdxPrix As Long, colIdxLT As Long
    Dim maxCol As Long
    Dim sDTR As String, sUC As String, sType As String, sRefCat As String
    Dim sCategorie As String, sKeyCat As String
    Dim vDate As Variant, vPrix As Variant, vLT As Variant
    Dim dictBest As Object
    Dim key As Variant
    Dim ligneCell As Long
    Dim wasOpen As Boolean

    sCheminBdD = ExtraireCheminBdD()
    If LenB(sCheminBdD) = 0 Then Exit Sub

    ' Ne pas refermer la BdD si elle etait deja ouverte chez l utilisateur
    wasOpen = False
    For Each wb In Application.Workbooks
        If LCase$(wb.FullName) = LCase$(sCheminBdD) Then
            Set wbBdD = wb
            wasOpen = True
            Exit For
        End If
    Next wb

    If wbBdD Is Nothing Then
        Set wbBdD = Application.Workbooks.Open(sCheminBdD, UpdateLinks:=0, ReadOnly:=True, Notify:=False, IgnoreReadOnlyRecommended:=True)
    End If

    Set wsBdD = wbBdD.Worksheets("Produit_SnapShotPrix")

    ' Mapper les colonnes par libelle de la ligne 1
    colIdxDTR = TrouverColParLibelle(wsBdD, "DTR")
    colIdxType = TrouverColParLibelle(wsBdD, "Type")
    colIdxRef = TrouverColParLibelle(wsBdD, "Ref Catalogue")
    colIdxDate = TrouverColParLibelle(wsBdD, "Date Source")
    colIdxPrix = TrouverColParLibelle(wsBdD, "Prix")
    colIdxLT = TrouverColParLibelle(wsBdD, "LeadTime")
    If colIdxDTR = 0 Or colIdxType = 0 Or colIdxRef = 0 Then GoTo CleanupBdD

    ' Lire tout en bulk dans un array 2D (bien plus rapide que cell-by-cell)
    derLigne = wsBdD.Cells(wsBdD.rows.count, colIdxDTR).End(xlUp).Row
    If derLigne < 2 Then GoTo CleanupBdD
    maxCol = Application.WorksheetFunction.Max(colIdxDTR, colIdxType, colIdxRef, colIdxDate, colIdxPrix, colIdxLT)
    arr = wsBdD.Range(wsBdD.Cells(2, 1), wsBdD.Cells(derLigne, maxCol)).Value

    Set dictBest = CreateObject("Scripting.Dictionary")
    dictBest.CompareMode = vbTextCompare

    For ligne = 1 To UBound(arr, 1)
        sDTR = Trim$(NzS(arr(ligne, colIdxDTR)))
        ' BdD peut stocker les DTR numeriques en Double (.0 a stripper)
        If Right$(sDTR, 2) = ".0" Then sDTR = Left$(sDTR, Len(sDTR) - 2)
        If LenB(sDTR) > 0 Then
            sUC = UCase$(sDTR)
            If dictDtrs.Exists(sUC) Then
                sType = UCase$(Trim$(NzS(arr(ligne, colIdxType))))
                sRefCat = UCase$(NzS(arr(ligne, colIdxRef)))
                vDate = arr(ligne, colIdxDate)
                vPrix = arr(ligne, colIdxPrix)
                vLT = arr(ligne, colIdxLT)

                ' Categoriser la Ref Catalogue : on retient uniquement DIS
                ' et FCS. Tout autre catalogue est ignore (decision 2026-06-28
                ' : pas de fallback "AUTRE", la ListView gere la cascade).
                sCategorie = ""
                If InStr(1, sRefCat, "D&IS CATALOGUE", vbTextCompare) > 0 Then
                    sCategorie = "DIS"
                ElseIf InStr(1, sRefCat, "FCS", vbTextCompare) > 0 Then
                    sCategorie = "FCS"
                End If

                ' On ne stocke que les types Spare/Repair avec Date valide
                ' et categorie DIS ou FCS reconnue.
                If LenB(sCategorie) > 0 And (sType = "SPARE" Or sType = "REPAIR") And IsDate(vDate) Then
                    sKeyCat = sUC & "|" & sType & "_" & sCategorie
                    Dim conserver As Boolean
                    conserver = True
                    If dictBest.Exists(sKeyCat) Then
                        Dim prev As Variant
                        prev = dictBest(sKeyCat)
                        If CDate(prev(0)) >= CDate(vDate) Then conserver = False
                    End If
                    If conserver Then
                        dictBest(sKeyCat) = Array(vDate, vPrix, arr(ligne, colIdxRef), vLT)
                    End If
                End If
            End If
        End If
    Next ligne

CleanupBdD:
    ' Fermer la BdD UNIQUEMENT si on l a ouverte dans cette Sub
    If Not wasOpen And Not wbBdD Is Nothing Then
        wbBdD.Close SaveChanges:=False
        Set wbBdD = Nothing
    End If

    If dictBest Is Nothing Then GoTo Sortie

    ' Ecriture des 7 colonnes Prix
    Dim colZ As Range, colAA As Range, colAC As Range, colAD As Range
    Dim colAE As Range, colAF As Range, colAG As Range
    Set colZ = lo.ListColumns(COL_PRIX_FCS).DataBodyRange
    Set colAA = lo.ListColumns(COL_PRIX_SPARE_CAT).DataBodyRange
    Set colAC = lo.ListColumns(COL_REF_SOURCE_SPARE).DataBodyRange
    Set colAD = lo.ListColumns(COL_LEAD_TIME_SPARE).DataBodyRange
    Set colAE = lo.ListColumns(COL_PRIX_REPAIR_CAT).DataBodyRange
    Set colAF = lo.ListColumns(COL_REF_SOURCE_REPAIR).DataBodyRange
    Set colAG = lo.ListColumns(COL_LEAD_TIME_REPAIR).DataBodyRange

    Dim choix As Variant
    For Each key In dictDtrs.Keys
        ligne = CLng(dictDtrs(key))

        ' --- Prix FCS (col Z) : strict FCS, MAX(Date) ----------
        If dictBest.Exists(key & "|SPARE_FCS") Then
            colZ.Cells(ligne, 1).Value = dictBest(key & "|SPARE_FCS")(1)
        Else
            colZ.Cells(ligne, 1).Value = Empty
        End If

        ' --- Prix Spare Catalogue + LT (AA/AD) : strict D&IS, MAX(Date).
        ' --- Ref Source Spare (AC) : cascade D&IS -> FCS pour le label
        '     uniquement (le prix AA reste D&IS strict ; le prix FCS vit
        '     en col Z + AB calcule).
        If dictBest.Exists(key & "|SPARE_DIS") Then
            choix = dictBest(key & "|SPARE_DIS")
            colAA.Cells(ligne, 1).Value = choix(1)
            colAC.Cells(ligne, 1).Value = choix(2)
            colAD.Cells(ligne, 1).Value = choix(3)
        Else
            colAA.Cells(ligne, 1).Value = Empty
            colAD.Cells(ligne, 1).Value = Empty
            ' Fallback label AC sur la ref FCS si dispo.
            If dictBest.Exists(key & "|SPARE_FCS") Then
                colAC.Cells(ligne, 1).Value = dictBest(key & "|SPARE_FCS")(2)
            Else
                colAC.Cells(ligne, 1).Value = Empty
            End If
        End If

        ' --- Prix Repair Catalogue + Ref + LT (AE/AF/AG) :
        '     strict D&IS, MAX(Date). Pas de FCS Repair en BdD.
        If dictBest.Exists(key & "|REPAIR_DIS") Then
            choix = dictBest(key & "|REPAIR_DIS")
            colAE.Cells(ligne, 1).Value = choix(1)
            colAF.Cells(ligne, 1).Value = choix(2)
            colAG.Cells(ligne, 1).Value = choix(3)
        Else
            colAE.Cells(ligne, 1).Value = Empty
            colAF.Cells(ligne, 1).Value = Empty
            colAG.Cells(ligne, 1).Value = Empty
        End If
    Next key
    Exit Sub

Sortie:
    ' Fallback cleanup en cas d erreur avant CleanupBdD
    On Error Resume Next
    If Not wasOpen And Not wbBdD Is Nothing Then wbBdD.Close SaveChanges:=False
    On Error GoTo 0
End Sub

' Helper : extrait le chemin du fichier BdD depuis CONNEXION_EXCEL_BDD_TRAVAIL.
' Format attendu : "Provider=...;Data Source=<chemin>;Extended Properties=..."
Private Function ExtraireCheminBdD() As String
    Dim cnx As String
    Dim posDS As Long, posEnd As Long
    cnx = Replace(CONNEXION_EXCEL_BDD_TRAVAIL, "UserProfile", Environ$("UserProfile"))
    posDS = InStr(1, cnx, "Data Source=", vbTextCompare)
    If posDS = 0 Then Exit Function
    posDS = posDS + Len("Data Source=")
    posEnd = InStr(posDS, cnx, ";")
    If posEnd = 0 Then posEnd = Len(cnx) + 1
    ExtraireCheminBdD = Mid$(cnx, posDS, posEnd - posDS)
End Function

' Helper : trouve la colonne dans la ligne 1 de ws par libelle (case-insensitive).
' Renvoie 0 si non trouve.
Private Function TrouverColParLibelle(ByVal ws As Worksheet, ByVal sLibelle As String) As Long
    Dim cell As Range
    On Error Resume Next
    Set cell = ws.Rows(1).Find(What:=sLibelle, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    If Not cell Is Nothing Then TrouverColParLibelle = cell.Column
    On Error GoTo 0
End Function

' =========================================================================
' BLOC 6 : RAFRAICHIR LES PRIX ARCHIVE depuis Projet_ASL$
' =========================================================================
' SELECT * FROM [Projet_ASL$] WHERE [PROJET_PROJETASL] <> '<projet courant>'
' Decision Thomas 2026-06-28 : pas de colonne Date dans Projet_ASL, donc
' on retient le projet alphabetiquement premier (StrComp asc) pour chaque
' DTR. Pas d AVG. Le prix retenu = celui du projet retenu, source
' tracable par la ListView Generation via dictProjetASL_Rows.
'   [ISP TIS HORS FRANCE (archive)]                  -> col AN
'   [Prix repair (DEVISE du projet (SLI))_PROJETASL] -> col AO
Private Sub RafraichirPrixArchive(ByVal lo As ListObject, ByVal dictDtrs As Object, ByVal sProjetCourant As String)
    Dim cnx As String
    Dim sDTR As String, sUC As String, sProjet As String
    Dim ligne As Long
    Dim dictSpare As Object  ' { DTR_UC -> Array(projetMin, prix) }
    Dim dictRepair As Object ' { DTR_UC -> Array(projetMin, prix) }
    Dim vSpare As Variant, vRepair As Variant
    Dim key As Variant
    Dim colAN As Range, colAO As Range
    Dim prevSpare As Variant, prevRepair As Variant

    cnx = Replace(CONNEXION_EXCEL_BDD_TRAVAIL, "UserProfile", Environ$("UserProfile"))
    Call Open_CONN(cnx)
    ' Cast DTR en String (cf. note dans RafraichirDepuisProduit)
    Call open_REQSEL("SELECT [DTR_PROJETASL] & '' AS DTR_str, [PROJET_PROJETASL], [ISP TIS HORS FRANCE (archive)], [Prix repair (DEVISE du projet (SLI))_PROJETASL] FROM [Projet_ASL$]")

    If REQSEL Is Nothing Then GoTo Sortie
    If REQSEL.EOF And REQSEL.BOF Then GoTo Sortie

    Set dictSpare = CreateObject("Scripting.Dictionary")
    Set dictRepair = CreateObject("Scripting.Dictionary")
    dictSpare.CompareMode = vbTextCompare
    dictRepair.CompareMode = vbTextCompare

    Do While Not REQSEL.EOF
        sDTR = Trim$(NzS(REQSEL.Fields("DTR_str").Value))
        sProjet = Trim$(NzS(REQSEL.Fields("PROJET_PROJETASL").Value))
        If LenB(sDTR) > 0 And LenB(sProjet) > 0 And UCase$(sProjet) <> UCase$(sProjetCourant) Then
            sUC = UCase$(sDTR)
            If dictDtrs.Exists(sUC) Then
                vSpare = REQSEL.Fields("ISP TIS HORS FRANCE (archive)").Value
                vRepair = REQSEL.Fields("Prix repair (DEVISE du projet (SLI))_PROJETASL").Value

                ' Spare : on retient le projet alpha 1er (StrComp ascendant)
                ' parmi les projets ayant un prix Spare > 0.
                If IsNumeric(vSpare) Then
                    If CDbl(vSpare) > 0 Then
                        If dictSpare.Exists(sUC) Then
                            prevSpare = dictSpare(sUC)
                            If StrComp(sProjet, CStr(prevSpare(0)), vbTextCompare) < 0 Then
                                dictSpare(sUC) = Array(sProjet, CDbl(vSpare))
                            End If
                        Else
                            dictSpare(sUC) = Array(sProjet, CDbl(vSpare))
                        End If
                    End If
                End If
                ' Repair : meme logique, dict separe (un projet peut avoir
                ' Spare connu mais pas Repair, ou inversement).
                If IsNumeric(vRepair) Then
                    If CDbl(vRepair) > 0 Then
                        If dictRepair.Exists(sUC) Then
                            prevRepair = dictRepair(sUC)
                            If StrComp(sProjet, CStr(prevRepair(0)), vbTextCompare) < 0 Then
                                dictRepair(sUC) = Array(sProjet, CDbl(vRepair))
                            End If
                        Else
                            dictRepair(sUC) = Array(sProjet, CDbl(vRepair))
                        End If
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
        If dictSpare.Exists(key) Then
            colAN.Cells(ligne, 1).Value = dictSpare(key)(1)
        Else
            colAN.Cells(ligne, 1).Value = Empty
        End If
        If dictRepair.Exists(key) Then
            colAO.Cells(ligne, 1).Value = dictRepair(key)(1)
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
' BLOC 7 : RAFRAICHIR LES COLONNES CALCULEES
' =========================================================================
' Decision Thomas 2026-06-28 : zero formule dans DATA_ASL. Les 6 colonnes
' "calcul local" sont reconstruites en valeurs pures a chaque Refresh,
' apres que les colonnes sources (Refresh + saisie ASL) soient toutes
' renseignees.
'
'   F  COL_DESIGNATION_PROVIDOR = Designation PBS (col B) si connue,
'                                  sinon Designation BdD (col E). Sert
'                                  comme cle de jointure DTR/Designation
'                                  pour generer Providor / ATM / EMMST.
'   I  COL_NTI                = Select Case Node typed (col G)
'                                LRU/LRUNR -> "3" ; SRU/SRUNR -> "3+" ; sinon ""
'   N  COL_TX_DEF_FONCT       = 1 / MTBF fonctionnel (col L) si > 0
'   O  COL_TX_DEF_GEN         = 1 / MTBF generique   (col M) si > 0
'   AB COL_PRIX_SPARE_CALC    = Prix FCS (col Z)               * coeff AB2
'   AL COL_COUT_SPARE_ALSTOM  = Cout spare FRS Not ISP (col AH) * coeff AL2
'   AM COL_PRIX_REPAIR_ALSTOM = Cout Repair FRS Not ISP (col AJ)* coeff AM2
'
' Regle source vide (Thomas 2026-06-28) : si la source est Empty, Null,
' non-numerique OU egale a 0, la cellule cible reste vide (pas de "0,00 ?"
' qui serait reinterprete par la ListView comme un prix Fournisseur connu).
'
' Prix arrondis a 2 decimales avant ecriture (sinon les multiplications
' par coeff genere des longs decimaux qui s affichent comme "373,179.39"
' avec separateur de milliers).
'
' Format Comptabilite ? applique aux colonnes de prix (Z/AA/AB/AE/AH/
' AJ/AL/AM/AN/AO) en sortie pour homogeneiser l affichage "350,00 ?".
'
' Les coefficients sont lus depuis la ligne 2 du worksheet (hors ListObject,
' au-dessus de la ligne d en-tete row 3).
Private Sub RafraichirCalculs(ByVal ws As Worksheet, ByVal lo As ListObject, ByVal dictDtrs As Object)
    Dim colDesPBS As Range, colDesBDD As Range, colDesProv As Range
    Dim colNode As Range, colNTI As Range
    Dim colMTBFFonct As Range, colMTBFGen As Range
    Dim colTxFonct As Range, colTxGen As Range
    Dim colPrixFCS As Range, colPrixSpareCalc As Range
    Dim colCoutSpareFRS As Range, colCoutSpareAlstom As Range
    Dim colCoutRepairFRS As Range, colPrixRepairAlstom As Range
    Dim coefAB As Variant, coefAL As Variant, coefAM As Variant
    Dim key As Variant
    Dim ligne As Long
    Dim sNode As String, sDesPBS As String, sDesBDD As String
    Dim v As Variant

    Set colDesPBS = lo.ListColumns(COL_DESIGNATION_PBS).DataBodyRange
    Set colDesBDD = lo.ListColumns(COL_DESIGNATION_BDD).DataBodyRange
    Set colDesProv = lo.ListColumns(COL_DESIGNATION_PROVIDOR).DataBodyRange
    Set colNode = lo.ListColumns(COL_NODE_TYPED).DataBodyRange
    Set colNTI = lo.ListColumns(COL_NTI).DataBodyRange
    Set colMTBFFonct = lo.ListColumns(COL_MTBF_FONCT).DataBodyRange
    Set colMTBFGen = lo.ListColumns(COL_MTBF_GEN).DataBodyRange
    Set colTxFonct = lo.ListColumns(COL_TX_DEF_FONCT).DataBodyRange
    Set colTxGen = lo.ListColumns(COL_TX_DEF_GEN).DataBodyRange
    Set colPrixFCS = lo.ListColumns(COL_PRIX_FCS).DataBodyRange
    Set colPrixSpareCalc = lo.ListColumns(COL_PRIX_SPARE_CALC).DataBodyRange
    Set colCoutSpareFRS = lo.ListColumns(COL_COUT_SPARE_FRS).DataBodyRange
    Set colCoutSpareAlstom = lo.ListColumns(COL_COUT_SPARE_ALSTOM).DataBodyRange
    Set colCoutRepairFRS = lo.ListColumns(COL_COUT_REPAIR_FRS).DataBodyRange
    Set colPrixRepairAlstom = lo.ListColumns(COL_PRIX_REPAIR_ALSTOM).DataBodyRange

    ' Coefficients ligne 2 (hors ListObject : row 2 < header row 3)
    coefAB = ws.Cells(2, COL_PRIX_SPARE_CALC).Value
    coefAL = ws.Cells(2, COL_COUT_SPARE_ALSTOM).Value
    coefAM = ws.Cells(2, COL_PRIX_REPAIR_ALSTOM).Value

    For Each key In dictDtrs.Keys
        ligne = CLng(dictDtrs(key))

        ' --- F : DESIGNATION (Providor) -------------------------
        '         PBS prio (saisie ASL), sinon fallback BdD (Refresh).
        '         Cle de jointure DTR/Designation pour Providor/ATM/EMMST.
        sDesPBS = Trim$(NzS(colDesPBS.Cells(ligne, 1).Value))
        sDesBDD = Trim$(NzS(colDesBDD.Cells(ligne, 1).Value))
        If LenB(sDesPBS) > 0 Then
            colDesProv.Cells(ligne, 1).Value = sDesPBS
        ElseIf LenB(sDesBDD) > 0 Then
            colDesProv.Cells(ligne, 1).Value = sDesBDD
        Else
            colDesProv.Cells(ligne, 1).ClearContents
        End If

        ' --- I : NTI ---------------------------------------------
        sNode = UCase$(Trim$(NzS(colNode.Cells(ligne, 1).Value)))
        Select Case sNode
            Case "LRU", "LRUNR"
                colNTI.Cells(ligne, 1).Value = "3"
            Case "SRU", "SRUNR"
                colNTI.Cells(ligne, 1).Value = "3+"
            Case Else
                colNTI.Cells(ligne, 1).ClearContents
        End Select

        ' --- N : Tx def fonct = 1 / MTBF fonctionnel ------------
        v = colMTBFFonct.Cells(ligne, 1).Value
        If EstNumeriqueNonNul(v) Then
            colTxFonct.Cells(ligne, 1).Value = 1 / CDbl(v)
        Else
            colTxFonct.Cells(ligne, 1).ClearContents
        End If

        ' --- O : Tx def gen = 1 / MTBF generique ----------------
        v = colMTBFGen.Cells(ligne, 1).Value
        If EstNumeriqueNonNul(v) Then
            colTxGen.Cells(ligne, 1).Value = 1 / CDbl(v)
        Else
            colTxGen.Cells(ligne, 1).ClearContents
        End If

        ' --- AB : Prix Spare Calcule = Prix FCS * coeff AB2 -----
        '         arrondi 2 decimales
        v = colPrixFCS.Cells(ligne, 1).Value
        If EstNumeriqueNonNul(v) And EstNumeriqueNonNul(coefAB) Then
            colPrixSpareCalc.Cells(ligne, 1).Value = ArrondiPrix(CDbl(v) * CDbl(coefAB))
        Else
            colPrixSpareCalc.Cells(ligne, 1).ClearContents
        End If

        ' --- AL : Cout Spare Alstom = Cout Spare FRS * coeff AL2 -
        v = colCoutSpareFRS.Cells(ligne, 1).Value
        If EstNumeriqueNonNul(v) And EstNumeriqueNonNul(coefAL) Then
            colCoutSpareAlstom.Cells(ligne, 1).Value = ArrondiPrix(CDbl(v) * CDbl(coefAL))
        Else
            colCoutSpareAlstom.Cells(ligne, 1).ClearContents
        End If

        ' --- AM : Prix Repair Alstom = Cout Repair FRS * coeff AM2
        v = colCoutRepairFRS.Cells(ligne, 1).Value
        If EstNumeriqueNonNul(v) And EstNumeriqueNonNul(coefAM) Then
            colPrixRepairAlstom.Cells(ligne, 1).Value = ArrondiPrix(CDbl(v) * CDbl(coefAM))
        Else
            colPrixRepairAlstom.Cells(ligne, 1).ClearContents
        End If
    Next key

    ' Format Comptabilite ? sur toutes les colonnes de prix (cf. comment
    ' tete de sub). Idempotent : reapplique a chaque Refresh.
    Call AppliquerFormatPrix(lo)
End Sub

' Helper : True si v est numerique ET different de 0. Filtre Empty, Null,
' "", strings non-numeriques et 0. Utilise par RafraichirCalculs pour
' eviter d ecrire "0,00 ?" quand la source est vide (qui serait alors
' interprete par la ListView Generation comme un prix Fournisseur connu).
Private Function EstNumeriqueNonNul(ByVal v As Variant) As Boolean
    If IsEmpty(v) Then Exit Function
    If IsNull(v) Then Exit Function
    If IsError(v) Then Exit Function
    If Not IsNumeric(v) Then Exit Function
    If CDbl(v) = 0 Then Exit Function
    EstNumeriqueNonNul = True
End Function

' Helper : arrondi commercial a 2 decimales (Application.Round, pas le
' Round VBA qui fait du banker's rounding).
Private Function ArrondiPrix(ByVal v As Double) As Double
    ArrondiPrix = Application.WorksheetFunction.Round(v, 2)
End Function

' Helper : applique le format Comptabilite EUR (style "350,00 EUR") aux
' colonnes de prix de DATA_ASL. Format VBA US-style (decimal=point,
' thousand=comma) ; Excel le traduit en local (FR : decimal=virgule,
' thousand=espace, donc "1 234,56 EUR").
'
' Pour les cellules vides, ce format affiche rien (Empty -> blank).
' Pour les zeros (qu on ne devrait pas avoir vu la regle EstNumerique
' NonNul mais defensive), il affiche "-   EUR" qui est l indicateur
' standard de "valeur zero" en comptabilite.
Private Sub AppliquerFormatPrix(ByVal lo As ListObject)
    Dim fmt As String
    Dim sEuro As String
    Dim cols As Variant
    Dim i As Long
    Dim colIdx As Variant

    ' Symbole EUR construit via ChrW(&H20AC) pour survivre a tout export
    ' / re-import du .bas quel que soit l encoding (ANSI, UTF-8, Latin1).
    ' Le fichier .bas est ASCII pur : ecrire le symbole directement dans
    ' une chaine litterale a deja ete perdu (remplace par '?' -> format
    ' Comptabilite affichait "1 234,56 ?" au lieu de "1 234,56 EUR").
    sEuro = ChrW(&H20AC)

    ' Format Comptabilite standard. Le \  represente un espace insecable
    ' avant le symbole. Les ?? alignent les "-" sur les decimales.
    fmt = "_-* #,##0.00\ """ & sEuro & """_-;-* #,##0.00\ """ & sEuro & """_-;_-* ""-""??\ """ & sEuro & """_-;_-@_-"

    cols = Array(COL_PRIX_FCS, COL_PRIX_SPARE_CAT, COL_PRIX_SPARE_CALC, COL_PRIX_REPAIR_CAT, COL_COUT_SPARE_FRS, COL_COUT_REPAIR_FRS, COL_COUT_SPARE_ALSTOM, COL_PRIX_REPAIR_ALSTOM, COL_PRIX_SPARE_ARCHIVE, COL_PRIX_REPAIR_ARCHIVE)

    On Error Resume Next
    For i = LBound(cols) To UBound(cols)
        colIdx = cols(i)
        lo.ListColumns(CLng(colIdx)).DataBodyRange.NumberFormat = fmt
    Next i
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
    wb.Names("DATA_ASL_DERNIERE_MAJ").RefersToRange.Value = Format$(Now(), "yyyy-mm-dd hh:mm")
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
Private Sub TracerRefreshLog(ByVal wb As Workbook, ByVal sProjet As String, ByVal nbDtrs As Long, ByVal duree As Double)
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

' DeProtegerFeuille / ProtegerFeuille supprimees (decision Thomas
' 2026-06-29 : plus aucune protection sur DATA_ASL ni sur le workbook).

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

' =========================================================================
' DIAGNOSTIC : lister les feuilles / tables de la BdD TRAVAIL
' =========================================================================
' A executer manuellement dans le VBE (F5) pour verifier le nom EXACT de
' la feuille SnapShotPrix dans la BdD. Output dans la fenetre Immediate
' (Ctrl+G dans VBE).
Public Sub Test_DumpFeuillesBdD()
    Dim cnx As String
    Dim cn As Object
    Dim rs As Object

    cnx = Replace(CONNEXION_EXCEL_BDD_TRAVAIL, "UserProfile", Environ$("UserProfile"))

    Set cn = CreateObject("ADODB.Connection")
    cn.Open cnx
    Set rs = cn.OpenSchema(20)   ' adSchemaTables = 20

    Debug.Print String(60, "=")
    Debug.Print "TABLES / FEUILLES de la BdD TRAVAIL"
    Debug.Print String(60, "=")
    Do While Not rs.EOF
        Debug.Print "  [" & rs.Fields("TABLE_TYPE").Value & "] " & rs.Fields("TABLE_NAME").Value
        rs.MoveNext
    Loop

    rs.Close
    cn.Close
End Sub

' =========================================================================
' DIAGNOSTIC : patterns de DTR dans Produit_SnapShotPrix
' Classifie chaque DTR par forme : "Numerique pur" (8600xxx), "DTRxxxxx",
' "ATXxxxxx", "Autre alphanumerique". Compte le total et donne 5 exemples
' par categorie.
' =========================================================================
Public Sub Test_CompterPatternsDTR()
    Dim cnx As String
    Dim sDTR As String, sCat As String
    Dim dictCount As Object, dictExamples As Object
    Dim key As Variant

    cnx = Replace(CONNEXION_EXCEL_BDD_TRAVAIL, "UserProfile", Environ$("UserProfile"))
    Call Open_CONN(cnx)
    ' Cast DTR en String pour recuperer aussi les DTR alphanumeriques
    Call open_REQSEL("SELECT DISTINCT [DTR] & '' AS DTR_str FROM [Produit_SnapShotPrix$]")

    If REQSEL Is Nothing Then GoTo Sortie
    If REQSEL.EOF And REQSEL.BOF Then GoTo Sortie

    Set dictCount = CreateObject("Scripting.Dictionary")
    Set dictExamples = CreateObject("Scripting.Dictionary")
    dictCount.CompareMode = vbTextCompare
    dictExamples.CompareMode = vbTextCompare

    Do While Not REQSEL.EOF
        sDTR = Trim$(NzS(REQSEL.Fields("DTR_str").Value))
        If Right$(sDTR, 2) = ".0" Then sDTR = Left$(sDTR, Len(sDTR) - 2)
        If LenB(sDTR) > 0 Then
            sCat = ""
            If IsNumeric(sDTR) Then
                sCat = "Numerique pur"
            ElseIf UCase$(Left$(sDTR, 3)) = "DTR" Then
                sCat = "DTRxxxxx"
            ElseIf UCase$(Left$(sDTR, 3)) = "ATX" Then
                sCat = "ATXxxxxx"
            Else
                sCat = "Autre alphanumerique"
            End If
            If dictCount.Exists(sCat) Then
                dictCount(sCat) = CLng(dictCount(sCat)) + 1
            Else
                dictCount.Add sCat, 1
                dictExamples.Add sCat, ""
            End If
            If Len(CStr(dictExamples(sCat))) < 200 Then
                dictExamples(sCat) = CStr(dictExamples(sCat)) & sDTR & " ; "
            End If
        End If
        REQSEL.MoveNext
    Loop

    Debug.Print String(60, "=")
    Debug.Print "Patterns DTR distincts dans Produit_SnapShotPrix"
    Debug.Print String(60, "=")
    For Each key In dictCount.Keys
        Debug.Print Format$(CLng(dictCount(key)), "00000") & " x  [" & key & "]"
        Debug.Print "       Exemples : " & dictExamples(key)
    Next key

Sortie:
    On Error Resume Next
    Call Close_CONN(cnx)
    On Error GoTo 0
End Sub

' =========================================================================
' DIAGNOSTIC : dry-run de RafraichirPrixCatalogue sur UN DTR donne.
' Affiche TOUTES les lignes SnapShotPrix matchant le DTR, avec :
'   - leur Type / Ref Catalogue / Prix / Date Source
'   - le verdict du filtre VBA (SPARE_DIS / SPARE_FCS / REPAIR_DIS / skip)
'   - la ligne finalement retenue par categorie (MAX Date Source)
' Modifier la constante TEST_DTR pour tester avec ton DTR connu.
' =========================================================================
Public Sub Test_RafraichirPrixDryRun()
    Const TEST_DTR As String = "8600361"   ' <-- modifier ici si besoin
    Dim cnx As String
    Dim sDTR As String, sType As String, sRefCat As String, sTypeRefUC As String
    Dim vDate As Variant, vPrix As Variant, vLT As Variant
    Dim dictBest As Object
    Dim key As Variant
    Dim nbLignesDTR As Long, nbMatchFiltre As Long

    cnx = Replace(CONNEXION_EXCEL_BDD_TRAVAIL, "UserProfile", Environ$("UserProfile"))
    Call Open_CONN(cnx)
    ' Cast DTR en String (cf. note RafraichirDepuisProduit)
    Call open_REQSEL("SELECT [DTR] & '' AS DTR_str, [Type], [Prix], [LeadTime], [Ref Catalogue], [Date Source] FROM [Produit_SnapShotPrix$]")

    If REQSEL Is Nothing Then GoTo Sortie
    If REQSEL.EOF And REQSEL.BOF Then GoTo Sortie

    Set dictBest = CreateObject("Scripting.Dictionary")
    dictBest.CompareMode = vbTextCompare

    Debug.Print String(60, "=")
    Debug.Print "Dry-run RafraichirPrixCatalogue pour DTR = " & TEST_DTR
    Debug.Print String(60, "=")

    Do While Not REQSEL.EOF
        ' Normaliser le DTR (BdD peut etre Double, .0 a stripper)
        sDTR = Trim$(NzS(REQSEL.Fields("DTR_str").Value))
        If Right$(sDTR, 2) = ".0" Then sDTR = Left$(sDTR, Len(sDTR) - 2)

        If sDTR = TEST_DTR Then
            nbLignesDTR = nbLignesDTR + 1
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

            Debug.Print "  [" & sType & "] '" & REQSEL.Fields("Ref Catalogue").Value & "'  Prix=" & NzS(vPrix) & "  Date=" & NzS(vDate) & "  IsDate=" & IsDate(vDate) & "  -> " & IIf(sTypeRefUC = "", "(skip)", sTypeRefUC)

            If LenB(sTypeRefUC) > 0 And IsDate(vDate) Then
                nbMatchFiltre = nbMatchFiltre + 1
                Dim k As String
                k = sTypeRefUC
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
        REQSEL.MoveNext
    Loop

    Debug.Print String(60, "-")
    Debug.Print "Total lignes BdD pour " & TEST_DTR & " : " & nbLignesDTR
    Debug.Print "Lignes matchant un filtre (D&IS/FCS) : " & nbMatchFiltre
    Debug.Print "Lignes retenues (max Date par categorie) : " & dictBest.count
    Debug.Print String(60, "-")
    Debug.Print "Lignes retenues :"
    For Each key In dictBest.Keys
        Debug.Print "  " & key & "  ->  Date=" & dictBest(key)(0) & "  Prix=" & dictBest(key)(1) & "  Ref='" & dictBest(key)(2) & "'  LT=" & dictBest(key)(3)
    Next key

Sortie:
    On Error Resume Next
    Call Close_CONN(cnx)
    On Error GoTo 0
End Sub

' =========================================================================
' DIAGNOSTIC : liste les valeurs distinctes de [Ref Catalogue] dans
' Produit_SnapShotPrix avec nb d occurrences par (Type, Ref Catalogue).
' Permet de connaitre les VRAIS libelles de catalogue presents en BdD.
' =========================================================================
Public Sub Test_DistinctRefCatalogue()
    Dim cnx As String
    Dim dict As Object
    Dim sType As String, sRef As String, sKey As String
    Dim key As Variant
    Dim total As Long

    cnx = Replace(CONNEXION_EXCEL_BDD_TRAVAIL, "UserProfile", Environ$("UserProfile"))
    Call Open_CONN(cnx)
    Call open_REQSEL("SELECT [Type], [Ref Catalogue] FROM [Produit_SnapShotPrix$]")

    If REQSEL Is Nothing Then GoTo Sortie
    If REQSEL.EOF And REQSEL.BOF Then GoTo Sortie

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    Do While Not REQSEL.EOF
        sType = NzS(REQSEL.Fields("Type").Value)
        sRef = NzS(REQSEL.Fields("Ref Catalogue").Value)
        sKey = "[" & sType & "]  " & sRef
        If dict.Exists(sKey) Then
            dict(sKey) = CLng(dict(sKey)) + 1
        Else
            dict.Add sKey, 1
        End If
        total = total + 1
        REQSEL.MoveNext
    Loop

    Debug.Print String(60, "=")
    Debug.Print "Couples (Type, Ref Catalogue) distincts dans Produit_SnapShotPrix"
    Debug.Print "Total rows scannees : " & total & "  -  Couples distincts : " & dict.count
    Debug.Print String(60, "=")
    For Each key In dict.Keys
        Debug.Print Format$(CLng(dict(key)), "0000") & " x  " & key
    Next key

Sortie:
    On Error Resume Next
    Call Close_CONN(cnx)
    On Error GoTo 0
End Sub

' =========================================================================
' DIAGNOSTIC : dump des colonnes + 5 premieres lignes de Produit_SnapShotPrix
' =========================================================================
' A executer manuellement dans le VBE (F5). Output dans Immediate.
' Permet de voir le vrai schema (noms de colonnes, types) et le format
' des valeurs reelles pour Type, Ref Catalogue, etc.
Public Sub Test_DumpSnapShotPrix()
    Dim cnx As String
    Dim f As Object
    Dim i As Long

    cnx = Replace(CONNEXION_EXCEL_BDD_TRAVAIL, "UserProfile", Environ$("UserProfile"))
    Call Open_CONN(cnx)
    Call open_REQSEL("SELECT * FROM [Produit_SnapShotPrix$]")

    If REQSEL Is Nothing Then
        Debug.Print "REQSEL = Nothing (la requete a echoue)"
        GoTo Sortie
    End If

    Debug.Print String(60, "=")
    Debug.Print "Schema de [Produit_SnapShotPrix$] :"
    Debug.Print String(60, "=")
    For Each f In REQSEL.Fields
        Debug.Print "  '" & f.Name & "'  (Type ADO=" & f.Type & ")"
    Next f

    If REQSEL.EOF And REQSEL.BOF Then
        Debug.Print "(RecordSet vide)"
        GoTo Sortie
    End If

    Debug.Print String(60, "-")
    Debug.Print "5 PREMIERES LIGNES :"
    Debug.Print String(60, "-")
    i = 0
    Do While Not REQSEL.EOF And i < 5
        Debug.Print "--- Ligne " & (i + 1) & " ---"
        For Each f In REQSEL.Fields
            Debug.Print "  " & f.Name & " = '" & NzS(f.Value) & "'"
        Next f
        REQSEL.MoveNext
        i = i + 1
    Loop

Sortie:
    On Error Resume Next
    Call Close_CONN(cnx)
    On Error GoTo 0
End Sub
