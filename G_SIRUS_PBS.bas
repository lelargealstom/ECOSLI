Attribute VB_Name = "G_SIRUS_PBS"
'--------------------------------------------------------------------------------------------------------
'
'                                       Generateur Livrables SLI
'                                       Module G_SIRUS_PBS
'                                       Developpement : LE LARGE Thomas (LGM/Alstom)
'                                       Cadrage de la feuille DATA_ASL et alimentation de la
'                                       ListView_Generation du UserForm ECOSLI.
'                                       Mai 2026
'
'  Convention : libelles DATA_ASL construits avec ChrW pour les accents afin que le fichier
'  reste compatible cp1252 / UTF-8 a l'import VBA. Les commentaires sont en francais sans accents.
'
'-----------------------------------------------------------------------------------------------------------
Option Compare Text
Option Explicit

' =========================================================================
' TYPE STRUCTURE - COLONNES DATA_ASL
' Une entree par colonne utile de la PBS du projet actif.
' La valeur 0 indique que la colonne n'a pas ete trouvee a l'etape de cadrage.
' Extensible : ajouter un champ ici + sa ligne dans MapColonnesDATA_ASL.
' =========================================================================
Public Type DATA_ASL_Col
    DTR_PBS                       As Long
    Designation_PBS               As Long
    Designation_BdD               As Long
    Designation_Providor          As Long
    SousSysteme                   As Long
    Qte                           As Long
    NodeType                      As Long
    HardSoft                      As Long
    NTI                           As Long
    RefPlan                       As Long
    MTBF_Gen                      As Long
    TxDef_Gen                     As Long
    DureeTacheCorrPrev_Gen_HARD   As Long
    MTBF_Intr                     As Long
    MTBF_Fonct                    As Long
    TxDef_Fonct                   As Long
    DureeTacheCorrPrev_Ana_HARD   As Long
    UniteMaintCorrective          As Long
    DureeVie_Gen_HARD             As Long
    DureeVie_Ana_HARD             As Long
    DureeVie_Gen_SOFT             As Long
    DureeTachePrev_Gen_SOFT       As Long
    DureeVie_Ana_SOFT             As Long
    DureeTachePrev_Ana_SOFT       As Long
    UniteMaintPreventive          As Long
    CoutTotal                     As Long
    ISP_TIS_FCS                   As Long
    ISP_TIS_Calc                  As Long
    LeadTimeSpareCalend           As Long
    PrixRepa_ISP                  As Long
    LeadTimeRepairCalend          As Long
    CoutSpare_FRS_NotISP          As Long
    LeadTimeSpare_FRS             As Long
    CoutRepair_FRS_NotISP         As Long
    LeadTimeRepair_FRS            As Long
    CoutSpare_AlstomFR_FRS        As Long
    PrixRepair_Alstom_FRS         As Long
    PrixManuelSpare               As Long
    PrixManuelRepair              As Long
    PrixSpare_ProjetArchive       As Long
    PrixRepair_ProjetArchive      As Long
End Type

Public DATA_ASL_type As DATA_ASL_Col
Public DATA_ASL_HeaderRow As Long          ' Ligne d'entete de DATA_ASL (auto-detectee si 0)

' Flag de configuration : True pour activer le lookup BdD des 4 colonnes
' ATM-Info / ATM-Taches / EMMST / Produit. False = colonnes laissees vides.
' Permet de couper le lookup BdD sans toucher au code de remplissage.
Public SIRUS_LookupDTR_Active As Boolean

' Caches DTR par table BdD - charges une fois par appel.
Private dictProjetASL As Object
Private dictInfoATM   As Object
Private dictTachesATM As Object
Private dictEMMST     As Object
Private dictProduit   As Object


' =========================================================================
' FONCTION: VerifierFeuilleDATA_ASL
' Verifie que la feuille DATA_ASL existe dans le classeur fourni et
' contient au moins une ligne sous l'entete.
'   Entree  wb : Workbook a inspecter
'   Sortie  ws : Worksheet pointe sur DATA_ASL si trouvee (ByRef)
'   Retour     : True si la feuille est exploitable, False sinon
' =========================================================================
Public Function VerifierFeuilleDATA_ASL(ByVal wb As Workbook, ByRef ws As Worksheet) As Boolean

    Dim feuille As Worksheet

    If wb Is Nothing Then
        VerifierFeuilleDATA_ASL = False
        Exit Function
    End If

    On Error Resume Next
    Set feuille = wb.Worksheets("DATA_ASL")
    On Error GoTo 0

    If feuille Is Nothing Then
        MsgBox "La feuille 'DATA_ASL' est introuvable dans le classeur '" & wb.Name & "'." & vbCrLf & _
               "Ouvre un classeur projet contenant la PBS avant de lancer ECOSLI.", _
               vbExclamation, "ECOSLI - PBS"
        VerifierFeuilleDATA_ASL = False
        Exit Function
    End If

    Set ws = feuille

    If DATA_ASL_HeaderRow = 0 Then DATA_ASL_HeaderRow = 1

    If feuille.UsedRange.rows.count < (DATA_ASL_HeaderRow + 1) Then
        VerifierFeuilleDATA_ASL = False
        Exit Function
    End If

    VerifierFeuilleDATA_ASL = True

End Function


' =========================================================================
' FONCTION: MapColonnesDATA_ASL
' Cadre les colonnes de la feuille DATA_ASL via Find sur la ligne d'entete.
'
' Colonnes critiques (DTR + 6 prix utilises pour la qualite) : si l'une
' est introuvable, la fonction retourne False et affiche un MsgBox
' recapitulatif. Les autres colonnes sont optionnelles : leur absence
' est juste tracee dans Debug.Print pour faciliter l'extension future.
'
'   Entree  ws : Worksheet DATA_ASL
'   Retour     : True si toutes les colonnes critiques sont mappees
' =========================================================================
Public Function MapColonnesDATA_ASL(ByVal ws As Worksheet) As Boolean

    Dim manqueCritique As String
    Dim manqueOptionnel As String
    Dim ligneAuto As Long

    ' Auto-detection de la ligne d'entete via le libelle ancre DTR.
    ' On cherche sur toute la feuille car DATA_ASL peut avoir un en-tete
    ' qui n'est pas en ligne 1 (titres, fusionnees, lignes de filtres...).
    ligneAuto = DetecterLigneEntete(ws, "DTR Reference (issue de PBS)")
    If ligneAuto > 0 Then
        DATA_ASL_HeaderRow = ligneAuto
    ElseIf DATA_ASL_HeaderRow = 0 Then
        DATA_ASL_HeaderRow = 1
    End If
    Debug.Print "[DATA_ASL] Ligne d'entete utilisee : " & DATA_ASL_HeaderRow

    With DATA_ASL_type
        ' --- Identification --------------------------------------------------
        .DTR_PBS = TrouverColonne(ws, "DTR Reference (issue de PBS)", True, manqueCritique, manqueOptionnel)
        .Designation_PBS = TrouverColonne(ws, "DESIGNATION (issue de PBS)", False, manqueCritique, manqueOptionnel)
        .Designation_BdD = TrouverColonne(ws, "DESIGNATION (issue de BdD)", False, manqueCritique, manqueOptionnel)
        .Designation_Providor = TrouverColonne(ws, _
            "DESIGNATION (mise en forme pour copier dans Providor : construite " & ChrW(224) & _
            " partir de celle de la PBS et de la d" & ChrW(233) & _
            "signation issue de BdD (ORCHESTRA)))", _
            False, manqueCritique, manqueOptionnel)
        .SousSysteme = TrouverColonne(ws, _
            "SOUS SYST" & ChrW(200) & "ME (copi" & ChrW(233) & "s de la PBS et classifi" & _
            ChrW(233) & "s selon g" & ChrW(233) & "n" & ChrW(233) & "riques)", _
            False, manqueCritique, manqueOptionnel)
        .Qte = TrouverColonne(ws, "Qte", False, manqueCritique, manqueOptionnel)
        .NodeType = TrouverColonne(ws, "Node typed (LRU/SRU" & ChrW(8230) & ")", False, manqueCritique, manqueOptionnel)
        .HardSoft = TrouverColonne(ws, "HARD/SOFT", False, manqueCritique, manqueOptionnel)
        .NTI = TrouverColonne(ws, "NTI", False, manqueCritique, manqueOptionnel)
        .RefPlan = TrouverColonne(ws, "r" & ChrW(233) & "f du plan d'ensemble (issue de la BdD)", False, manqueCritique, manqueOptionnel)

        ' --- RAM -------------------------------------------------------------
        .MTBF_Gen = TrouverColonne(ws, "MTBF g" & ChrW(233) & "n" & ChrW(233) & "rique (offre et PGR)", False, manqueCritique, manqueOptionnel)
        .TxDef_Gen = TrouverColonne(ws, "Tx d" & ChrW(233) & "ffaillance g" & ChrW(233) & "n" & ChrW(233) & "rique", False, manqueCritique, manqueOptionnel)
        .DureeTacheCorrPrev_Gen_HARD = TrouverColonne(ws, _
            "Dur" & ChrW(233) & "e de la t" & ChrW(226) & "che corrective et pr" & ChrW(233) & _
            "ventive g" & ChrW(233) & "n" & ChrW(233) & "rique (offre et PGR) HARDWARE", _
            False, manqueCritique, manqueOptionnel)
        .MTBF_Intr = TrouverColonne(ws, "MTBF intrins" & ChrW(232) & "que (RAM report specifique projet (CGR))", False, manqueCritique, manqueOptionnel)
        .MTBF_Fonct = TrouverColonne(ws, "MTBF fonctionnel (RAM report specifique projet (CGR))", False, manqueCritique, manqueOptionnel)
        .TxDef_Fonct = TrouverColonne(ws, "Tx d" & ChrW(233) & "ffaillance fonctionnel", False, manqueCritique, manqueOptionnel)
        .DureeTacheCorrPrev_Ana_HARD = TrouverColonne(ws, _
            "Dur" & ChrW(233) & "e de la t" & ChrW(226) & "che corrective et pr" & ChrW(233) & _
            "ventive analytique (CGR) HARDWARE", _
            False, manqueCritique, manqueOptionnel)
        .UniteMaintCorrective = TrouverColonne(ws, "UNITE de la maintenance corrective", False, manqueCritique, manqueOptionnel)

        ' --- Duree de vie / taches preventives -----------------------------
        .DureeVie_Gen_HARD = TrouverColonne(ws, "DUREE DE VIE LIMITEE g" & ChrW(233) & "n" & ChrW(233) & "rique (offre et PGR) HARDWARE", False, manqueCritique, manqueOptionnel)
        .DureeVie_Ana_HARD = TrouverColonne(ws, "DUREE DE VIE LIMITEE (RAM report specifique projet (CGR)) HARDWARE", False, manqueCritique, manqueOptionnel)
        .DureeVie_Gen_SOFT = TrouverColonne(ws, "DUREE DE VIE LIMITEE g" & ChrW(233) & "n" & ChrW(233) & "rique (offre et PGR) SOFTWARE", False, manqueCritique, manqueOptionnel)
        .DureeTachePrev_Gen_SOFT = TrouverColonne(ws, _
            "Dur" & ChrW(233) & "e de la t" & ChrW(226) & "che pr" & ChrW(233) & _
            "ventive g" & ChrW(233) & "n" & ChrW(233) & "rique SOFTWARE (offre et PGR)", _
            False, manqueCritique, manqueOptionnel)
        .DureeVie_Ana_SOFT = TrouverColonne(ws, "DUREE DE VIE LIMITEE (RAM report specifique projet (CGR)) SOFTWARE", False, manqueCritique, manqueOptionnel)
        .DureeTachePrev_Ana_SOFT = TrouverColonne(ws, _
            "Dur" & ChrW(233) & "e de la t" & ChrW(226) & "che pr" & ChrW(233) & _
            "ventive analytique SOFTWARE (CGR)", _
            False, manqueCritique, manqueOptionnel)
        .UniteMaintPreventive = TrouverColonne(ws, "UNITE de la maintenance pr" & ChrW(233) & "ventive", False, manqueCritique, manqueOptionnel)

        ' --- Couts et prix --------------------------------------------------
        ' Critiques : les 6 prix utilises pour la qualite Spare/Repair
        .CoutTotal = TrouverColonne(ws, "Cout total", False, manqueCritique, manqueOptionnel)
        .ISP_TIS_FCS = TrouverColonne(ws, "ISP TIS FRANCE (FCS)", True, manqueCritique, manqueOptionnel)
        .ISP_TIS_Calc = TrouverColonne(ws, "ISP TIS FRANCE (calcul" & ChrW(233) & ")", False, manqueCritique, manqueOptionnel)
        .LeadTimeSpareCalend = TrouverColonne(ws, "Lead Time SPARE (en jours calendaires)", False, manqueCritique, manqueOptionnel)
        .PrixRepa_ISP = TrouverColonne(ws, "Prix repa (ISP Catalogue R" & ChrW(233) & "pa)", True, manqueCritique, manqueOptionnel)
        .LeadTimeRepairCalend = TrouverColonne(ws, "Lead Time REPAIR (en jours calendaires)", False, manqueCritique, manqueOptionnel)
        .CoutSpare_FRS_NotISP = TrouverColonne(ws, "Cout spare FOURNISSEUR (Not ISP)", False, manqueCritique, manqueOptionnel)
        .LeadTimeSpare_FRS = TrouverColonne(ws, "Lead Time SPARE FOURNISSEUR", False, manqueCritique, manqueOptionnel)
        .CoutRepair_FRS_NotISP = TrouverColonne(ws, "Cout Repair FOURNISSEUR (Not ISP)", False, manqueCritique, manqueOptionnel)
        .LeadTimeRepair_FRS = TrouverColonne(ws, "Lead Time REPAIR FOURNISSEUR", False, manqueCritique, manqueOptionnel)
        .CoutSpare_AlstomFR_FRS = TrouverColonne(ws, "Cout spare ALSTOM FRANCE FOURNISSEUR", True, manqueCritique, manqueOptionnel)
        .PrixRepair_Alstom_FRS = TrouverColonne(ws, "Prix repair ALSTOM FOURNISSEUR", True, manqueCritique, manqueOptionnel)
        .PrixManuelSpare = TrouverColonne(ws, "Prix Manuel Spare", True, manqueCritique, manqueOptionnel)
        .PrixManuelRepair = TrouverColonne(ws, "Prix Manuel Repair", True, manqueCritique, manqueOptionnel)
        .PrixSpare_ProjetArchive = TrouverColonne(ws, "Prix Spare sur un projet archive", False, manqueCritique, manqueOptionnel)
        .PrixRepair_ProjetArchive = TrouverColonne(ws, "Prix Repair sur un projet archive", False, manqueCritique, manqueOptionnel)
    End With

    If Len(manqueOptionnel) > 0 Then
        Debug.Print "[DATA_ASL] Colonnes optionnelles introuvables : " & manqueOptionnel
    End If

    If Len(manqueCritique) > 0 Then
        MsgBox "Cadrage DATA_ASL impossible." & vbCrLf & vbCrLf & _
               "Colonnes critiques introuvables sur la ligne " & DATA_ASL_HeaderRow & " :" & vbCrLf & _
               manqueCritique, vbCritical, "ECOSLI - PBS"
        MapColonnesDATA_ASL = False
        Exit Function
    End If

    MapColonnesDATA_ASL = True

End Function


' =========================================================================
' FONCTION: TrouverColonne
' Cherche un libelle exact sur la ligne d'entete et retourne son numero
' de colonne. Si introuvable, accumule le libelle dans le bon log
' (critique ou optionnel) et retourne 0.
' =========================================================================
Private Function TrouverColonne(ByVal ws As Worksheet, ByVal libelle As String, ByVal critique As Boolean, _
                                ByRef manqueCritique As String, ByRef manqueOptionnel As String) As Long

    Dim cel As Range

    On Error Resume Next
    ' Tentative match exact (xlWhole)
    Set cel = ws.rows(DATA_ASL_HeaderRow).Find(What:=libelle, LookIn:=xlValues, lookat:=xlWhole, MatchCase:=False)
    ' Fallback match partiel (xlPart) - tolere espaces / sauts de ligne / suffixes
    If cel Is Nothing Then
        Set cel = ws.rows(DATA_ASL_HeaderRow).Find(What:=libelle, LookIn:=xlValues, lookat:=xlPart, MatchCase:=False)
    End If
    On Error GoTo 0

    If cel Is Nothing Then
        TrouverColonne = 0
        If critique Then
            manqueCritique = manqueCritique & "  - " & libelle & vbCrLf
        Else
            manqueOptionnel = manqueOptionnel & libelle & " ; "
        End If
    Else
        TrouverColonne = cel.Column
    End If

End Function


' =========================================================================
' FONCTION: DetecterLigneEntete
' Cherche un libelle ancre dans toute la feuille et retourne la ligne ou
' il est trouve (xlWhole d'abord, xlPart en fallback). 0 si introuvable.
' =========================================================================
Private Function DetecterLigneEntete(ByVal ws As Worksheet, ByVal libelleAncre As String) As Long

    Dim cel As Range

    On Error Resume Next
    Set cel = ws.Cells.Find(What:=libelleAncre, LookIn:=xlValues, lookat:=xlWhole, MatchCase:=False, _
                            SearchOrder:=xlByRows, SearchDirection:=xlNext)
    If cel Is Nothing Then
        Set cel = ws.Cells.Find(What:=libelleAncre, LookIn:=xlValues, lookat:=xlPart, MatchCase:=False, _
                                SearchOrder:=xlByRows, SearchDirection:=xlNext)
    End If
    On Error GoTo 0

    If cel Is Nothing Then
        DetecterLigneEntete = 0
    Else
        DetecterLigneEntete = cel.Row
    End If

End Function


' =========================================================================
' SUB: SIRUS_RemplirListViewGeneration
' Alimente la ListView_Generation du UserForm avec, pour chaque DTR de
' la PBS du classeur actif :
'   Col 1  : DTR
'   Col 2  : Projet_ASL - "OUI" si DTR present dans Projet_ASL$, "NON" sinon
'   Col 3  : ATM-Info   - "OK"  si DTR present dans Info_ATM$
'   Col 4  : ATM-Tache  - "OK"  si DTR present dans Taches_ATM$
'   Col 5  : EMMST      - "OK"  si DTR present dans EMMST$
'   Col 6  : Produit    - "OK"  si DTR present dans Produit$
'   Col 7  : Qualite Prix Spare
'   Col 8  : Source Ecart Spare
'   Col 9  : Source Manque Spare
'   Col 10 : Qualite Prix Repair
'   Col 11 : Source Ecart Repair
'   Col 12 : Source Manque Repair
'
' Source : feuille DATA_ASL du ActiveWorkbook au moment de l'appel.
' Suppose la connexion BdD ouverte (cf. UserForm_Initialize).
'
' Requetes SQL emises (une par lookup, executees une seule fois par appel
' puis cachees en Dictionary pour test d'appartenance en O(1)) :
'   1. SELECT [DTR]            FROM [Projet_ASL$]
'   2. SELECT [DTR_INFOATM]    FROM [Info_ATM$]
'   3. SELECT [DTR_TACHEATM]   FROM [Taches_ATM$]
'   4. SELECT [DTR_EMMST]      FROM [EMMST$]
'   5. SELECT [DTR]            FROM [Produit$]
' =========================================================================
Public Sub SIRUS_RemplirListViewGeneration()

    Dim wsPBS As Worksheet
    Dim wb As Workbook
    Dim lv As MSComctlLib.ListView
    Dim ligne As Long
    Dim derniereLigne As Long
    Dim dtr As String
    Dim vMan As Variant, vFrs As Variant, vCat As Variant
    Dim resSpare As Variant, resRepair As Variant
    Dim li As MSComctlLib.ListItem
    Dim nbAjoutees As Long

    On Error GoTo Sortie

    Debug.Print "=== SIRUS_RemplirListViewGeneration : DEBUT ==="
    Debug.Print "  ActiveWorkbook = " & IIf(ActiveWorkbook Is Nothing, "<Nothing>", ActiveWorkbook.Name)
    Debug.Print "  SIRUS_LookupDTR_Active = " & SIRUS_LookupDTR_Active

    Set wb = ActiveWorkbook
    If wb Is Nothing Then
        Debug.Print "  ABORT : ActiveWorkbook est Nothing."
        GoTo Sortie
    End If

    Set lv = UserForm.ListView_Generation
    If lv Is Nothing Then
        Debug.Print "  ABORT : UserForm.ListView_Generation est Nothing."
        GoTo Sortie
    End If

    ' Reset systematique pour tolerer DATA_ASL absent / vide sans ListView residuelle
    lv.ListItems.Clear

    If Not VerifierFeuilleDATA_ASL(wb, wsPBS) Then
        Debug.Print "  ABORT : VerifierFeuilleDATA_ASL = False (DATA_ASL absente ou vide dans '" & wb.Name & "')."
        GoTo Sortie
    End If
    Debug.Print "  DATA_ASL OK : ws=" & wsPBS.Name & ", UsedRange.rows=" & wsPBS.UsedRange.rows.count

    If Not MapColonnesDATA_ASL(wsPBS) Then
        Debug.Print "  ABORT : MapColonnesDATA_ASL = False (colonne critique introuvable)."
        GoTo Sortie
    End If
    Debug.Print "  Mapping DATA_ASL OK : DTR_PBS=col" & DATA_ASL_type.DTR_PBS & _
                ", HeaderRow=" & DATA_ASL_HeaderRow

    ' Caches DTR BdD - charges seulement si le lookup est active.
    If SIRUS_LookupDTR_Active Then
        Debug.Print "  Chargement des 5 caches DTR BdD..."
        Set dictProjetASL = ChargerSetDTR_BdD("Projet_ASL$", "DTR")
        Set dictInfoATM = ChargerSetDTR_BdD("Info_ATM$", "DTR_INFOATM")
        Set dictTachesATM = ChargerSetDTR_BdD("Taches_ATM$", "DTR_TACHEATM")
        Set dictEMMST = ChargerSetDTR_BdD("EMMST$", "DTR_EMMST")
        Set dictProduit = ChargerSetDTR_BdD("Produit$", "DTR")
    Else
        Set dictProjetASL = Nothing
        Set dictInfoATM = Nothing
        Set dictTachesATM = Nothing
        Set dictEMMST = Nothing
        Set dictProduit = Nothing
    End If

    ' Derniere ligne sur la colonne DTR
    If DATA_ASL_type.DTR_PBS = 0 Then
        Debug.Print "  ABORT : DATA_ASL_type.DTR_PBS = 0 (colonne DTR non mappee)."
        GoTo Sortie
    End If
    derniereLigne = wsPBS.Cells(wsPBS.rows.count, DATA_ASL_type.DTR_PBS).End(xlUp).Row
    Debug.Print "  Derniere ligne DTR_PBS = " & derniereLigne & " (HeaderRow = " & DATA_ASL_HeaderRow & ")"
    If derniereLigne <= DATA_ASL_HeaderRow Then
        Debug.Print "  ABORT : aucune ligne de donnees sous l'entete."
        GoTo Sortie
    End If

    Application.ScreenUpdating = False
    nbAjoutees = 0

    For ligne = DATA_ASL_HeaderRow + 1 To derniereLigne

        dtr = Trim$(CStr(wsPBS.Cells(ligne, DATA_ASL_type.DTR_PBS).Value))
        If Len(dtr) = 0 Then GoTo SuivanteLigne

        Set li = lv.ListItems.Add(, , dtr)
        nbAjoutees = nbAjoutees + 1
        ' Col 2 : Projet_ASL utilise OUI/NON (presence binaire critique avant generation).
        ' Cols 3-6 : presence informative en OK / vide.
        If SIRUS_LookupDTR_Active Then
            li.SubItems(1) = IIf(SetContient(dictProjetASL, dtr), "OUI", "NON")
        Else
            li.SubItems(1) = ""
        End If
        li.SubItems(2) = IIf(SetContient(dictInfoATM, dtr), "OK", "")
        li.SubItems(3) = IIf(SetContient(dictTachesATM, dtr), "OK", "")
        li.SubItems(4) = IIf(SetContient(dictEMMST, dtr), "OK", "")
        li.SubItems(5) = IIf(SetContient(dictProduit, dtr), "OK", "")

        ' Qualite SPARE : Prix Manuel Spare / Cout spare ALSTOM FRANCE FOURNISSEUR / ISP TIS FRANCE (FCS)
        vMan = LireCellule(wsPBS, ligne, DATA_ASL_type.PrixManuelSpare)
        vFrs = LireCellule(wsPBS, ligne, DATA_ASL_type.CoutSpare_AlstomFR_FRS)
        vCat = LireCellule(wsPBS, ligne, DATA_ASL_type.ISP_TIS_FCS)
        resSpare = EvaluerQualitePrix(vMan, vFrs, vCat)
        li.SubItems(6) = resSpare(0)
        li.SubItems(7) = resSpare(1)
        li.SubItems(8) = resSpare(2)

        ' Qualite REPAIR : Prix Manuel Repair / Prix repair ALSTOM FOURNISSEUR / Prix repa (ISP Catalogue Repa)
        vMan = LireCellule(wsPBS, ligne, DATA_ASL_type.PrixManuelRepair)
        vFrs = LireCellule(wsPBS, ligne, DATA_ASL_type.PrixRepair_Alstom_FRS)
        vCat = LireCellule(wsPBS, ligne, DATA_ASL_type.PrixRepa_ISP)
        resRepair = EvaluerQualitePrix(vMan, vFrs, vCat)
        li.SubItems(9) = resRepair(0)
        li.SubItems(10) = resRepair(1)
        li.SubItems(11) = resRepair(2)

SuivanteLigne:
    Next ligne

    Debug.Print "  TERMINE : " & nbAjoutees & " ligne(s) ajoutee(s) dans la ListView_Generation."
    Debug.Print "=== SIRUS_RemplirListViewGeneration : FIN ==="
    Application.ScreenUpdating = True
    Exit Sub

Sortie:
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then
        Debug.Print "  ERREUR " & Err.Number & " : " & Err.Description & _
                    " (source : " & Err.Source & ")"
    End If
    Debug.Print "=== SIRUS_RemplirListViewGeneration : FIN (sortie anticipee) ==="
End Sub


' =========================================================================
' SUB: SIRUS_RafraichirListViewGeneration
' Bouton de rafraichissement page Generation : vide la ListView et
' rappelle la routine d'alimentation. A brancher sur le clic d'un
' bouton "Rafraichir PBS".
' =========================================================================
Public Sub SIRUS_RafraichirListViewGeneration()
    Call SIRUS_RemplirListViewGeneration
End Sub


' =========================================================================
' FONCTION: ChargerSetDTR_BdD
' Charge la liste des DTR d'une table BdD dans un Dictionary
' { DTR -> True } pour permettre un test d'appartenance en O(1).
' Suppose Open_CONN deja appelee (UserForm_Initialize).
'
' Implementation : GetRows materialise tout le RecordSet en tableau 2D
' en une seule operation. Plus robuste qu'une boucle MoveNext / EOF qui
' peut s'interrompre prematurement sur les sources Excel ODBC (curseur
' forward-only fragile).
' =========================================================================
Private Function ChargerSetDTR_BdD(ByVal nomTable As String, ByVal colonneDTR As String) As Object

    Dim dict As Object
    Dim STRSEL As String
    Dim arr As Variant
    Dim i As Long
    Dim valeur As String

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    On Error GoTo Sortie

    STRSEL = "SELECT [" & colonneDTR & "] FROM [" & nomTable & "]"
    Call open_REQSEL(STRSEL)

    If REQSEL Is Nothing Then
        Debug.Print "  [" & nomTable & "] ABORT : REQSEL = Nothing (Open_CONN appele ?)"
        GoTo Sortie
    End If
    If REQSEL.State <> 1 Then
        Debug.Print "  [" & nomTable & "] ABORT : REQSEL.State = " & REQSEL.State & " (attendu 1=adStateOpen)"
        GoTo Sortie
    End If
    If REQSEL.EOF And REQSEL.BOF Then
        Debug.Print "  [" & nomTable & "] table vide (EOF et BOF)"
        GoTo Sortie
    End If

    ' Extraction complete en une passe. arr est dimensionne (champs x lignes).
    arr = REQSEL.GetRows()

    For i = LBound(arr, 2) To UBound(arr, 2)
        valeur = NzStr(arr(0, i))
        If Len(valeur) > 0 Then
            If Not dict.Exists(valeur) Then dict.Add valeur, True
        End If
    Next i

    Debug.Print "  [" & nomTable & "] DTR charges : " & dict.count

Sortie:
    If Err.Number <> 0 Then
        Debug.Print "  [" & nomTable & "] ERREUR " & Err.Number & " : " & Err.Description
    End If
    On Error Resume Next
    If Not REQSEL Is Nothing Then
        If REQSEL.State = 1 Then REQSEL.Close
    End If
    On Error GoTo 0
    Set ChargerSetDTR_BdD = dict

End Function


' =========================================================================
' FONCTION: SetContient
' Test d'appartenance d'un DTR a un Dictionary. Robuste a Nothing / vide.
' =========================================================================
Private Function SetContient(ByVal dict As Object, ByVal dtr As String) As Boolean
    If dict Is Nothing Then Exit Function
    If Len(dtr) = 0 Then Exit Function
    SetContient = dict.Exists(dtr)
End Function


' =========================================================================
' FONCTION: LireCellule
' Lecture protegee : retourne Empty si la colonne vaut 0 (non mappee).
' =========================================================================
Private Function LireCellule(ByVal ws As Worksheet, ByVal ligne As Long, ByVal colonne As Long) As Variant
    If colonne = 0 Then
        LireCellule = Empty
        Exit Function
    End If
    LireCellule = ws.Cells(ligne, colonne).Value
End Function


' =========================================================================
' FONCTION: EvaluerQualitePrix
' Calcule les 3 indicateurs de qualite pour un trio (Manuel / Fournisseur /
' Catalogue) selon les regles SIRUS :
'
'   0/3 connus                -> "NOK - Manque" / "-" / "Manuel + Fournisseur + Catalogue"
'   1/3 connus                -> "NOK - Manque" / "-" / les 2 sources manquantes
'   2/3 connus, ecart >  seuil -> "NOK - Manque & Ecart" / source la plus
'                                ecartee de la mediane / source manquante
'   2/3 connus, ecart <= seuil -> "NOK - Manque" / "-" / source manquante
'   3/3 connus, ecart >  seuil -> "NOK - Ecart" / source la plus ecartee / "-"
'   3/3 connus, ecart <= seuil -> "OK" / "-" / "-"
'
' Ordre des sources : 1 = Manuel, 2 = Fournisseur, 3 = Catalogue.
' Retour : Array(qualite, sourceEcart, sourceManque) -- toutes des String.
' =========================================================================
Public Function EvaluerQualitePrix(ByVal vManuel As Variant, ByVal vFournisseur As Variant, ByVal vCatalogue As Variant, _
                                   Optional ByVal seuil As Double = 0.25) As Variant

    Dim names(1 To 3) As String
    Dim vals(1 To 3) As Double
    Dim known(1 To 3) As Boolean
    Dim nb As Long
    Dim qualite As String, srcEcart As String, srcManque As String
    Dim manquantes As String
    Dim i As Long
    Dim med As Double
    Dim diff As Double, bestDiff As Double
    Dim bestIdx As Long

    names(1) = "Manuel"
    names(2) = "Fournisseur"
    names(3) = "Catalogue"

    nb = 0
    If EstNumeriqueConnu(vManuel) Then known(1) = True: vals(1) = CDbl(vManuel): nb = nb + 1
    If EstNumeriqueConnu(vFournisseur) Then known(2) = True: vals(2) = CDbl(vFournisseur): nb = nb + 1
    If EstNumeriqueConnu(vCatalogue) Then known(3) = True: vals(3) = CDbl(vCatalogue): nb = nb + 1

    ' Liste des sources manquantes (separees par " + ")
    manquantes = ""
    For i = 1 To 3
        If Not known(i) Then
            If Len(manquantes) > 0 Then manquantes = manquantes & " + "
            manquantes = manquantes & names(i)
        End If
    Next i

    Select Case nb
        Case 0, 1
            qualite = "NOK - Manque"
            srcEcart = "-"
            srcManque = manquantes

        Case 2, 3
            med = MedianeConnus(vals, known)

            ' Recherche de la source la plus ecartee de la mediane
            bestDiff = -1#
            bestIdx = 0
            For i = 1 To 3
                If known(i) Then
                    If med = 0 Then
                        diff = IIf(vals(i) = 0, 0#, 1E+99)
                    Else
                        diff = Abs(vals(i) - med) / Abs(med)
                    End If
                    If diff > bestDiff Then
                        bestDiff = diff
                        bestIdx = i
                    End If
                End If
            Next i

            If bestDiff > seuil Then
                If nb = 2 Then
                    qualite = "NOK - Manque & Ecart"
                    srcEcart = names(bestIdx)
                    srcManque = manquantes
                Else
                    qualite = "NOK - Ecart"
                    srcEcart = names(bestIdx)
                    srcManque = "-"
                End If
            Else
                If nb = 2 Then
                    qualite = "NOK - Manque"
                    srcEcart = "-"
                    srcManque = manquantes
                Else
                    qualite = "OK"
                    srcEcart = "-"
                    srcManque = "-"
                End If
            End If
    End Select

    EvaluerQualitePrix = Array(qualite, srcEcart, srcManque)

End Function


' =========================================================================
' FONCTION: EstNumeriqueConnu
' True si la valeur est exploitable comme prix (non vide, non erreur,
' numerique). 0 est considere connu.
' =========================================================================
Private Function EstNumeriqueConnu(ByVal v As Variant) As Boolean
    If IsError(v) Then Exit Function
    If IsMissing(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    If Trim$(CStr(v)) = "" Then Exit Function
    If Not IsNumeric(v) Then Exit Function
    EstNumeriqueConnu = True
End Function


' =========================================================================
' FONCTION: MedianeConnus
' Mediane des valeurs marquees connues dans known(). Retourne 0 si vide.
' =========================================================================
Private Function MedianeConnus(ByRef vals() As Double, ByRef known() As Boolean) As Double

    Dim arr(1 To 3) As Double
    Dim n As Long, i As Long, j As Long, t As Double

    n = 0
    For i = 1 To 3
        If known(i) Then
            n = n + 1
            arr(n) = vals(i)
        End If
    Next i

    If n = 0 Then Exit Function

    For i = 1 To n - 1
        For j = i + 1 To n
            If arr(j) < arr(i) Then t = arr(i): arr(i) = arr(j): arr(j) = t
        Next j
    Next i

    If n Mod 2 = 1 Then
        MedianeConnus = arr((n + 1) \ 2)
    Else
        MedianeConnus = (arr(n \ 2) + arr(n \ 2 + 1)) / 2
    End If

End Function


' =========================================================================
' SUB: Test_MapColonnesDATA_ASL
' Procedure executable depuis l'editeur VBA pour tester le cadrage sur
' la feuille DATA_ASL du classeur actif (sans passer par le UserForm).
' Affiche les numeros de colonne dans la fenetre Immediate.
' =========================================================================
Public Sub Test_MapColonnesDATA_ASL()

    Dim ws As Worksheet

    If Not VerifierFeuilleDATA_ASL(ActiveWorkbook, ws) Then
        Debug.Print "[TEST] Feuille DATA_ASL non exploitable."
        Exit Sub
    End If

    If Not MapColonnesDATA_ASL(ws) Then
        Debug.Print "[TEST] Mapping echoue (colonnes critiques manquantes)."
        Exit Sub
    End If

    Debug.Print String(60, "-")
    Debug.Print "[TEST] Cadrage DATA_ASL OK (ligne entete = " & DATA_ASL_HeaderRow & ")"
    Debug.Print "  DTR_PBS                = " & DATA_ASL_type.DTR_PBS
    Debug.Print "  Designation_PBS        = " & DATA_ASL_type.Designation_PBS
    Debug.Print "  SousSysteme            = " & DATA_ASL_type.SousSysteme
    Debug.Print "  PrixManuelSpare        = " & DATA_ASL_type.PrixManuelSpare
    Debug.Print "  CoutSpare_AlstomFR_FRS = " & DATA_ASL_type.CoutSpare_AlstomFR_FRS
    Debug.Print "  ISP_TIS_FCS            = " & DATA_ASL_type.ISP_TIS_FCS
    Debug.Print "  PrixManuelRepair       = " & DATA_ASL_type.PrixManuelRepair
    Debug.Print "  PrixRepair_Alstom_FRS  = " & DATA_ASL_type.PrixRepair_Alstom_FRS
    Debug.Print "  PrixRepa_ISP           = " & DATA_ASL_type.PrixRepa_ISP

End Sub


' =========================================================================
' SUB: Test_EvaluerQualitePrix
' Tests unitaires basiques sur EvaluerQualitePrix. Affiche les resultats
' dans la fenetre Immediate.
' =========================================================================
Public Sub Test_EvaluerQualitePrix()

    Dim r As Variant

    r = EvaluerQualitePrix(Empty, Empty, Empty)
    Debug.Print "[0/3]                " & r(0) & " | Ecart=" & r(1) & " | Manque=" & r(2)

    r = EvaluerQualitePrix(100, Empty, Empty)
    Debug.Print "[1/3 Manuel]         " & r(0) & " | Ecart=" & r(1) & " | Manque=" & r(2)

    r = EvaluerQualitePrix(100, 110, Empty)
    Debug.Print "[2/3 ecart faible]   " & r(0) & " | Ecart=" & r(1) & " | Manque=" & r(2)

    r = EvaluerQualitePrix(100, 200, Empty)
    Debug.Print "[2/3 ecart fort]     " & r(0) & " | Ecart=" & r(1) & " | Manque=" & r(2)

    r = EvaluerQualitePrix(100, 110, 105)
    Debug.Print "[3/3 ecart faible]   " & r(0) & " | Ecart=" & r(1) & " | Manque=" & r(2)

    r = EvaluerQualitePrix(100, 110, 300)
    Debug.Print "[3/3 ecart fort]     " & r(0) & " | Ecart=" & r(1) & " | Manque=" & r(2)

End Sub


