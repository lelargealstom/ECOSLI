' =========================================================================
' MODULE: D_EMMST_CacheCommentaires
' =========================================================================
' Auteur : Thomas LE LARGE - 2026-11-09
' Perimetre : contournement de la troncature 255 chars d'ACE OLE DB sur
'             les 10 colonnes commentaires EMMST (Emballage, Marquage,
'             Manutention, Stockage, Transport, FR + EN).
'
' Strategie : cache session Dictionary(DTR -> Array(10 commentaires)),
' peuple lazy au 1er appel de Main_EMMST via copie locale %TEMP% +
' Workbooks.Open ReadOnly invisible + lecture bloc Range.Value + fermeture
' + Kill. Les appels suivants de la session reutilisent le cache memoire.
'
' Pattern issu du bench v2 scenario A3 (14 s au 1er appel, 0 s aux suivants).
' Cf. Docs/CHANGELOG.md session 2026-11-09 pour l'historique complet :
' Piste 1 ADO KO -> bench v1 scenarios A/B -> bench v2 A1/A2/A3 -> ici.
'
' Appele par : C_EMMST_Main.RecupererDonneesBDD_EMMST juste apres
' EMMST_tab = RecordsetToArray(REQSEL).
'
' Progression : 5 etapes via F_Progression (pattern ATM V1/V2/Providor).
' =========================================================================

Option Explicit

' Cache session persistant pendant toute la duree de chargement de ECOSLI.xlam.
' Cle   : DTR (String).
' Valeur : Array(0 To 9) Variant - ordre FR puis EN :
'   0 Emballage FR, 1 Marquage FR, 2 Manutention FR, 3 Stockage FR,
'   4 Transport FR, 5 Emballage EN, 6 Marquage EN, 7 Manutention EN,
'   8 Stockage EN, 9 Transport EN.
Private m_dictCommentaires As Object

' Noms des 10 colonnes commentaires - ordre FR puis EN, utilise a la fois
' pour la lecture BdD (via Find) et pour la reinjection dans EMMST_tab (via
' Arraypos). L'ordre DOIT etre coherent (meme index 0-9 source/destination).
Private Const N_EMBALLAGE_FR As String = "CommentaireEmballage_EMMST"
Private Const N_MARQUAGE_FR As String = "CommentaireMarquage_EMMST"
Private Const N_MANUTENTION_FR As String = "CommentaireManutention_EMMST"
Private Const N_STOCKAGE_FR As String = "CommentaireStockage_EMMST"
Private Const N_TRANSPORT_FR As String = "CommentaireTransport_EMMST"
Private Const N_EMBALLAGE_EN As String = "CommentaireEmballageEN_EMMST"
Private Const N_MARQUAGE_EN As String = "CommentaireMarquageEN_EMMST"
Private Const N_MANUTENTION_EN As String = "CommentaireManutentionEN_EMMST"
Private Const N_STOCKAGE_EN As String = "CommentaireStockageEN_EMMST"
Private Const N_TRANSPORT_EN As String = "CommentaireTransportEN_EMMST"


' =========================================================================
' PROCEDURE: ChargerEtInjecterCommentairesEMMST
' =========================================================================
' Point d'entree unique appele depuis C_EMMST_Main.RecupererDonneesBDD_EMMST
' juste apres EMMST_tab = RecordsetToArray(REQSEL).
'
' Au 1er appel de la session : charge le cache (14 s typiques).
' Aux appels suivants : reutilise le Dictionary memoire (0 s).
'
' Ecrase dans EMMST_tab les 10 colonnes commentaires tronquees par ADO avec
' les valeurs completes. Les DTR absents du cache (nouveaux ou BdD desync)
' sont laisses intacts (donc valeur tronquee ADO conservee).
Public Sub ChargerEtInjecterCommentairesEMMST()

    Dim colDTR As Long
    Dim colsBdD As Variant
    Dim colsTab() As Long
    Dim iCol As Long
    Dim iLigne As Long
    Dim dtr As String
    Dim arrComm As Variant

    If m_dictCommentaires Is Nothing Then
        Call ChargerCache
    End If

    colsBdD = Array(N_EMBALLAGE_FR, N_MARQUAGE_FR, N_MANUTENTION_FR, _
                    N_STOCKAGE_FR, N_TRANSPORT_FR, _
                    N_EMBALLAGE_EN, N_MARQUAGE_EN, N_MANUTENTION_EN, _
                    N_STOCKAGE_EN, N_TRANSPORT_EN)

    colDTR = Arraypos(EMMST_tab, "DTR_EMMST")
    If colDTR = -1 Then
        Err.Raise vbObjectError + 540, "ChargerEtInjecterCommentairesEMMST", _
                  "Colonne DTR_EMMST introuvable dans EMMST_tab."
    End If

    ReDim colsTab(LBound(colsBdD) To UBound(colsBdD))
    For iCol = LBound(colsBdD) To UBound(colsBdD)
        colsTab(iCol) = Arraypos(EMMST_tab, CStr(colsBdD(iCol)))
        If colsTab(iCol) = -1 Then
            Err.Raise vbObjectError + 541, "ChargerEtInjecterCommentairesEMMST", _
                      "Colonne introuvable dans EMMST_tab : " & colsBdD(iCol)
        End If
    Next iCol

    ' Parcours des lignes de donnees (ligne 0 = en-tetes dans RecordsetToArray)
    For iLigne = 1 To UBound(EMMST_tab, 1)
        dtr = CStr(EMMST_tab(iLigne, colDTR))
        If m_dictCommentaires.Exists(dtr) Then
            arrComm = m_dictCommentaires(dtr)
            For iCol = LBound(colsBdD) To UBound(colsBdD)
                EMMST_tab(iLigne, colsTab(iCol)) = arrComm(iCol)
            Next iCol
        End If
    Next iLigne

End Sub


' =========================================================================
' PROCEDURE: ChargerCache
' =========================================================================
' 1er appel uniquement : copie BdD TRAVAIL vers %TEMP%, ouvre en Excel
' invisible ReadOnly, lit les 10 colonnes commentaires + DTR via Range.Value,
' ferme sans sauvegarder, supprime la copie. 5 etapes via F_Progression.
'
' Decoupage de cout (bench v2 scenario A3 sur poste Alstom, 998 DTR) :
'   - FileCopy local       : 0.07 s  (elimine latence OneDrive sur close)
'   - Workbooks.Open       : 12.86 s (incompressible, parsing .xlsb)
'   - Range.Value x 10     : 0.02 s  (lecture bloc, 10 colonnes)
'   - wb.Close             : 1.14 s
'   - Restauration + Kill  : 0.22 s
'   - TOTAL                : 14.40 s
'
' Application.ScreenUpdating / EnableEvents / Calculation / DisplayAlerts
' sont deja positionnes par TurnOffFunctionality appele dans Main_EMMST
' avant RecupererDonneesBDD_EMMST. On ne touche pas a l'etat Application
' ici (restauration sous controle de TurnOnFunctionality en fin de Main).
Private Sub ChargerCache()

    Dim cheminBdD As String
    Dim cheminTravail As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim colDTR As Long
    Dim colsBdD As Variant
    Dim colsIdx() As Long
    Dim iCol As Long
    Dim derniereLigne As Long
    Dim dtrs As Variant
    Dim tabCom0 As Variant, tabCom1 As Variant, tabCom2 As Variant
    Dim tabCom3 As Variant, tabCom4 As Variant, tabCom5 As Variant
    Dim tabCom6 As Variant, tabCom7 As Variant, tabCom8 As Variant
    Dim tabCom9 As Variant
    Dim tabsParCol As Variant
    Dim iLigne As Long
    Dim dtr As String
    Dim arrComm(0 To 9) As Variant

    Set m_dictCommentaires = CreateObject("Scripting.Dictionary")
    m_dictCommentaires.CompareMode = vbTextCompare

    colsBdD = Array(N_EMBALLAGE_FR, N_MARQUAGE_FR, N_MANUTENTION_FR, _
                    N_STOCKAGE_FR, N_TRANSPORT_FR, _
                    N_EMBALLAGE_EN, N_MARQUAGE_EN, N_MANUTENTION_EN, _
                    N_STOCKAGE_EN, N_TRANSPORT_EN)

    Call Progress_Demarrer(5, "Chargement commentaires EMMST")

    ' --- Etape 1 : copie locale %TEMP% -----------------------------------
    Call Progress_Etape(1, "Copie BdD en local", "")
    cheminBdD = ExtraireDataSourceDepuisConnStr( _
        Replace(CONNEXION_EXCEL_BDD_TRAVAIL, "UserProfile", Environ$("UserProfile")))
    cheminTravail = Environ$("TEMP") & "\ECOSLI_bdd_cache_" & _
                    Format(Now, "yyyymmdd_hhnnss") & ".xlsb"
    FileCopy cheminBdD, cheminTravail

    ' --- Etape 2 : ouverture BdD invisible ReadOnly ----------------------
    Call Progress_Etape(2, "Ouverture BdD invisible", "")
    Set wb = Workbooks.Open(Filename:=cheminTravail, UpdateLinks:=0, _
                             ReadOnly:=True, Notify:=False)
    Set ws = wb.Worksheets("EMMST")

    colDTR = TrouverColonneBdD(ws, "DTR_EMMST")
    ReDim colsIdx(LBound(colsBdD) To UBound(colsBdD))
    For iCol = LBound(colsBdD) To UBound(colsBdD)
        colsIdx(iCol) = TrouverColonneBdD(ws, CStr(colsBdD(iCol)))
    Next iCol

    ' --- Etape 3 : lecture bloc Range.Value (10 colonnes + DTR) ----------
    Call Progress_Etape(3, "Lecture commentaires", "")
    With ws
        derniereLigne = .Cells(.Rows.count, colDTR).End(xlUp).Row
        dtrs = .Range(.Cells(2, colDTR), .Cells(derniereLigne, colDTR)).Value

        tabCom0 = .Range(.Cells(2, colsIdx(0)), .Cells(derniereLigne, colsIdx(0))).Value
        tabCom1 = .Range(.Cells(2, colsIdx(1)), .Cells(derniereLigne, colsIdx(1))).Value
        tabCom2 = .Range(.Cells(2, colsIdx(2)), .Cells(derniereLigne, colsIdx(2))).Value
        tabCom3 = .Range(.Cells(2, colsIdx(3)), .Cells(derniereLigne, colsIdx(3))).Value
        tabCom4 = .Range(.Cells(2, colsIdx(4)), .Cells(derniereLigne, colsIdx(4))).Value
        tabCom5 = .Range(.Cells(2, colsIdx(5)), .Cells(derniereLigne, colsIdx(5))).Value
        tabCom6 = .Range(.Cells(2, colsIdx(6)), .Cells(derniereLigne, colsIdx(6))).Value
        tabCom7 = .Range(.Cells(2, colsIdx(7)), .Cells(derniereLigne, colsIdx(7))).Value
        tabCom8 = .Range(.Cells(2, colsIdx(8)), .Cells(derniereLigne, colsIdx(8))).Value
        tabCom9 = .Range(.Cells(2, colsIdx(9)), .Cells(derniereLigne, colsIdx(9))).Value
    End With
    tabsParCol = Array(tabCom0, tabCom1, tabCom2, tabCom3, tabCom4, _
                       tabCom5, tabCom6, tabCom7, tabCom8, tabCom9)

    ' --- Etape 4 : indexation Dictionary ---------------------------------
    Call Progress_Etape(4, "Indexation DTR", "")
    For iLigne = 1 To UBound(dtrs, 1)
        dtr = CStr(dtrs(iLigne, 1))
        If LenB(dtr) > 0 Then
            For iCol = 0 To 9
                arrComm(iCol) = tabsParCol(iCol)(iLigne, 1)
            Next iCol
            ' Assignation d'un Array a un Dictionary = copie par valeur en
            ' VBA. La modification de arrComm a l'iteration suivante ne
            ' modifie pas la valeur deja stockee dans le Dictionary.
            m_dictCommentaires(dtr) = arrComm
        End If
    Next iLigne

    ' --- Etape 5 : fermeture propre + Kill copie -------------------------
    Call Progress_Etape(5, "Fermeture BdD", "")
    wb.Close SaveChanges:=False
    Set wb = Nothing
    Set ws = Nothing
    Kill cheminTravail

    Call Progress_Terminer

End Sub


' =========================================================================
' FONCTION: TrouverColonneBdD
' =========================================================================
' Cherche nomColonne en ligne 1 de ws et retourne l'index de colonne.
' Pattern Find convention 2 : xlValues + xlWhole, fallback xlPart.
Private Function TrouverColonneBdD(ByRef ws As Worksheet, _
                                    ByVal nomColonne As String) As Long
    Dim trouve As Range

    Set trouve = ws.Rows(1).Find(What:=nomColonne, LookIn:=xlValues, _
                                  LookAt:=xlWhole, MatchCase:=False)
    If trouve Is Nothing Then
        Set trouve = ws.Rows(1).Find(What:=nomColonne, LookIn:=xlValues, _
                                      LookAt:=xlPart, MatchCase:=False)
    End If

    If trouve Is Nothing Then
        Err.Raise vbObjectError + 542, "TrouverColonneBdD", _
                  "Colonne introuvable dans la feuille EMMST : " & nomColonne
    End If

    TrouverColonneBdD = trouve.Column
End Function


' =========================================================================
' FONCTION: ExtraireDataSourceDepuisConnStr
' =========================================================================
' Extrait la valeur de Data Source= dans une chaine de connexion ADO.
' Forme attendue : Provider=...;Data Source=<chemin>;Extended Properties='...';
Private Function ExtraireDataSourceDepuisConnStr(ByVal connStr As String) As String
    Dim posStart As Long
    Dim posFin As Long
    Dim debut As Long

    posStart = InStr(1, connStr, "Data Source=", vbTextCompare)
    If posStart = 0 Then
        Err.Raise vbObjectError + 543, "ExtraireDataSourceDepuisConnStr", _
                  "Chaine sans 'Data Source='."
    End If

    debut = posStart + Len("Data Source=")
    posFin = InStr(debut, connStr, ";")
    If posFin = 0 Then posFin = Len(connStr) + 1

    ExtraireDataSourceDepuisConnStr = Mid$(connStr, debut, posFin - debut)
End Function
