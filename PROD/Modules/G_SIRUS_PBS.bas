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

' Caches rows par table BdD - charges une fois par appel.
' Structure : Dictionary { DTR (UCase trim) -> Collection of Dictionary
'                          { ColonneBdD -> valeur } }.
' Une Collection par DTR car certaines tables ont plusieurs lignes par DTR
' (cas TachesATM : N taches par DTR ; ProjetASL : N projets par DTR si non
' filtre). Permet la moyenne MeQ sur toutes les lignes.
Private dictProjetASL_Rows As Object
Private dictInfoATM_Rows   As Object
Private dictTachesATM_Rows As Object
Private dictEMMST_Rows     As Object
Private dictProduit_Rows   As Object


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
        DetecterLigneEntete = cel.row
    End If

End Function


' =========================================================================
' SUB: SIRUS_RemplirListViewGeneration
' Alimente la ListView_Generation du UserForm avec, pour chaque DTR de
' la PBS du classeur actif :
'   Col 1  : DTR
'   Col 2  : Projet_ASL - % MeQ (vert 100% / rouge < 100% / "Absent" si pas de row)
'   Col 3  : ATM-Info   - % MeQ
'   Col 4  : ATM-Tache  - % MeQ (moyenne sur toutes les taches du DTR)
'   Col 5  : EMMST      - % MeQ
'   Col 6  : Produit    - % MeQ
'   Col 7  : Qualite Prix Spare
'   Col 8  : Source Ecart Spare
'   Col 9  : Source Manque Spare
'   Col 10 : Qualite Prix Repair
'   Col 11 : Source Ecart Repair
'   Col 12 : Source Manque Repair
'
' MeQ = champs Obligatoire="Oui" remplis / champs Obligatoire="Oui" attendus,
' + Conditionnels evalues via SIRUS_RegleObligatoire (G_SIRUS_Regles).
' Source : SIRUS_LireConfig(<table>) pour la liste des champs.
' Source : feuille DATA_ASL du ActiveWorkbook au moment de l'appel.
' Suppose la connexion BdD ouverte (cf. UserForm_Initialize).
'
' Requetes SQL emises (une par lookup, executees une seule fois par appel
' puis cachees en Dictionary de rows pour calcul MeQ) :
'   1. SELECT * FROM [Projet_ASL$]
'   2. SELECT * FROM [Info_ATM$]
'   3. SELECT * FROM [Taches_ATM$]
'   4. SELECT * FROM [EMMST$]
'   5. SELECT * FROM [Produit$]
' =========================================================================
Public Sub SIRUS_RemplirListViewGeneration()

    Dim wsPBS As Worksheet
    Dim wb As Workbook
    Dim lv As MSComctlLib.ListView
    Dim ligne As Long, derniereLigne As Long
    Dim dtr As String
    Dim vMan As Variant, vFrs As Variant, vCat As Variant
    Dim resSpare As Variant, resRepair As Variant
    Dim li As MSComctlLib.ListItem
    Dim nbAjoutees As Long

    On Error GoTo Sortie

    Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Sub
    Set lv = UserForm.ListView_Generation
    If lv Is Nothing Then Exit Sub
    lv.ListItems.Clear

    If Not VerifierFeuilleDATA_ASL(wb, wsPBS) Then Exit Sub
    If Not MapColonnesDATA_ASL(wsPBS) Then Exit Sub
    If DATA_ASL_type.DTR_PBS = 0 Then Exit Sub

    ' Caches config SIRUS par table - lus UNE SEULE FOIS au lieu de
    ' 5 x N (perf : evite N x 5 lectures de la feuille SIRUS_Config).
    Dim cfgProjetASL As Collection, cfgInfoATM As Collection
    Dim cfgTachesATM As Collection, cfgEMMST As Collection, cfgProduit As Collection
    If SIRUS_LookupDTR_Active Then
        Set cfgProjetASL = SIRUS_LireConfig("ProjetASL")
        Set cfgInfoATM = SIRUS_LireConfig("InfoATM")
        Set cfgTachesATM = SIRUS_LireConfig("TachesATM")
        Set cfgEMMST = SIRUS_LireConfig("EMMST")
        Set cfgProduit = SIRUS_LireConfig("Produit")
    End If

    ' Caches rows BdD - une seule requete SELECT (cible) par table.
    ' ProjetASL : pas de filtre projet (on veut le count global de projets
    ' connus par DTR, cf. nouvelle colonne "N projets" ci-dessous).
    If SIRUS_LookupDTR_Active Then
        Set dictProjetASL_Rows = ChargerRowsTableBdD_Optim("Projet_ASL$", "DTR_PROJETASL", cfgProjetASL, "PROJET_PROJETASL")
        Set dictInfoATM_Rows = ChargerRowsTableBdD_Optim("Info_ATM$", "DTR_INFOATM", cfgInfoATM)
        Set dictTachesATM_Rows = ChargerRowsTableBdD_Optim("Taches_ATM$", "DTR_TACHEATM", cfgTachesATM)
        Set dictEMMST_Rows = ChargerRowsTableBdD_Optim("EMMST$", "DTR_EMMST", cfgEMMST)
        Set dictProduit_Rows = ChargerRowsTableBdD_Optim("Produit$", "DTR", cfgProduit)
    End If

    derniereLigne = wsPBS.Cells(wsPBS.rows.count, DATA_ASL_type.DTR_PBS).End(xlUp).row
    If derniereLigne <= DATA_ASL_HeaderRow Then Exit Sub

    Application.ScreenUpdating = False

    For ligne = DATA_ASL_HeaderRow + 1 To derniereLigne
        dtr = Trim$(CStr(wsPBS.Cells(ligne, DATA_ASL_type.DTR_PBS).Value))
        If Len(dtr) = 0 Then GoTo SuivanteLigne

        Set li = lv.ListItems.Add(, , dtr)
        nbAjoutees = nbAjoutees + 1

        ' Col 2 : ProjetASL = nombre de projets distincts ou ce DTR existe
        '         en BdD (au lieu d'un % MeQ). La liste des projets est
        '         affichee en tooltip de ligne (ListItem.ToolTipText).
        ' Cols 3 a 6 : % MeQ par DTR pour les 4 autres tables BdD.
        '              Vert 100%, rouge sinon, "Absent" si pas de row.
        If SIRUS_LookupDTR_Active Then
            Dim nProjets As Long
            Dim sListeProjets As String
            Call ProjetsConnusPourDTR(dictProjetASL_Rows, dtr, nProjets, sListeProjets)
            Call AjouterSubItemNbProjets(li, nProjets)

            ' Tooltip de ligne avec la liste des projets (ListItem.ToolTipText
            ' est niveau ligne, pas par cellule -- MSComctlLib v6 ne supporte
            ' pas le tooltip par sub-item. Hover sur la 1re cellule (DTR)
            ' affiche le texte).
            On Error Resume Next
            If nProjets > 0 Then
                li.TooltipText = "Projets connus (" & nProjets & ") :" & vbLf & sListeProjets
            Else
                li.TooltipText = "Aucun projet connu pour ce DTR"
            End If
            On Error GoTo 0

            Call AjouterSubItemMeQ(li, CalculerMeQDTR_Optim(dictInfoATM_Rows, dtr, "InfoATM", cfgInfoATM))
            Call AjouterSubItemMeQ(li, CalculerMeQDTR_Optim(dictTachesATM_Rows, dtr, "TachesATM", cfgTachesATM))
            Call AjouterSubItemMeQ(li, CalculerMeQDTR_Optim(dictEMMST_Rows, dtr, "EMMST", cfgEMMST))
            Call AjouterSubItemMeQ(li, CalculerMeQDTR_Optim(dictProduit_Rows, dtr, "Produit", cfgProduit))
        Else
            Call AjouterSubItemColore(li, "", False)
            Call AjouterSubItemColore(li, "", False)
            Call AjouterSubItemColore(li, "", False)
            Call AjouterSubItemColore(li, "", False)
            Call AjouterSubItemColore(li, "", False)
        End If

        ' Qualite SPARE (Manuel / Alstom FRS / ISP catalogue)
        vMan = LireCellule(wsPBS, ligne, DATA_ASL_type.PrixManuelSpare)
        vFrs = LireCellule(wsPBS, ligne, DATA_ASL_type.CoutSpare_AlstomFR_FRS)
        vCat = LireCellule(wsPBS, ligne, DATA_ASL_type.ISP_TIS_FCS)
        resSpare = EvaluerQualitePrix(vMan, vFrs, vCat)
        Call AjouterSubItemColore(li, CStr(resSpare(0)), False, CouleurQualite(CStr(resSpare(0)), CStr(resSpare(2))))
        Call AjouterSubItemColore(li, CStr(resSpare(1)), False)
        Call AjouterSubItemColore(li, CStr(resSpare(2)), False)

        ' Qualite REPAIR (Manuel / Alstom FRS / ISP catalogue)
        vMan = LireCellule(wsPBS, ligne, DATA_ASL_type.PrixManuelRepair)
        vFrs = LireCellule(wsPBS, ligne, DATA_ASL_type.PrixRepair_Alstom_FRS)
        vCat = LireCellule(wsPBS, ligne, DATA_ASL_type.PrixRepa_ISP)
        resRepair = EvaluerQualitePrix(vMan, vFrs, vCat)
        Call AjouterSubItemColore(li, CStr(resRepair(0)), False, CouleurQualite(CStr(resRepair(0)), CStr(resRepair(2))))
        Call AjouterSubItemColore(li, CStr(resRepair(1)), False)
        Call AjouterSubItemColore(li, CStr(resRepair(2)), False)

SuivanteLigne:
    Next ligne

    Debug.Print "SIRUS_RemplirListViewGeneration : " & nbAjoutees & _
                " ligne(s) (Lookup=" & SIRUS_LookupDTR_Active & ")"
    Debug.Print "  Colonnes prix mappees (0 = introuvable) :"
    Debug.Print "    PrixManuelSpare        = col " & DATA_ASL_type.PrixManuelSpare
    Debug.Print "    CoutSpare_AlstomFR_FRS = col " & DATA_ASL_type.CoutSpare_AlstomFR_FRS
    Debug.Print "    ISP_TIS_FCS            = col " & DATA_ASL_type.ISP_TIS_FCS
    Debug.Print "    PrixManuelRepair       = col " & DATA_ASL_type.PrixManuelRepair
    Debug.Print "    PrixRepair_Alstom_FRS  = col " & DATA_ASL_type.PrixRepair_Alstom_FRS
    Debug.Print "    PrixRepa_ISP           = col " & DATA_ASL_type.PrixRepa_ISP

Sortie:
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then
        Debug.Print "ERREUR SIRUS_RemplirListViewGeneration : " & Err.description
    End If
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
' Renvoie un Dictionary { DTR -> True } des DTR presents dans la colonne
' demandee de la table BdD. Une seule requete ADODB SELECT, RecordSet
' materialise en tableau via GetRows pour eviter MoveNext sur ODBC Excel.
' Suppose Open_CONN deja appelee (par UserForm_Initialize).

    Dim dict As Object
    Dim arr As Variant
    Dim i As Long
    Dim valeur As String

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    On Error GoTo Sortie

    Call open_REQSEL("SELECT [" & colonneDTR & "] FROM [" & nomTable & "]")
    If REQSEL Is Nothing Then GoTo Sortie
    If REQSEL.EOF And REQSEL.BOF Then GoTo Sortie

    arr = REQSEL.GetRows()
    For i = LBound(arr, 2) To UBound(arr, 2)
        valeur = NzStr(arr(0, i))
        If Len(valeur) > 0 And Not dict.Exists(valeur) Then dict.Add valeur, True
    Next i

Sortie:
    If Err.Number <> 0 Then
        Debug.Print "[" & nomTable & "] ChargerSetDTR_BdD : " & Err.description
    End If
    On Error Resume Next
    If Not REQSEL Is Nothing Then If REQSEL.State = 1 Then REQSEL.Close
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
' FONCTION: NzSafeStr
' Conversion en String resistante a Null / Empty / Error. Contrairement
' a NzStr (F_FonctionsGlobales) qui ne gere pas Null et crash sur
' CStr(Null) ("Utilisation incorrecte de Null"). Indispensable pour les
' valeurs ADO BdD (cellules Excel vides renvoient Null).
' =========================================================================
Private Function NzSafeStr(ByVal v As Variant) As String
    On Error Resume Next
    If IsNull(v) Then
        NzSafeStr = ""
    ElseIf IsEmpty(v) Then
        NzSafeStr = ""
    ElseIf IsError(v) Then
        NzSafeStr = ""
    Else
        NzSafeStr = CStr(v)
    End If
    If Err.Number <> 0 Then
        NzSafeStr = ""
        Err.Clear
    End If
    On Error GoTo 0
End Function


' =========================================================================
' FONCTION: ChargerRowsTableBdD_Optim (version optimisee)
' Charge les lignes d'une table BdD dans un Dictionary indexe par DTR.
' Chaque entree = Collection de Dictionary { ColonneBdD -> valeur }.
'
' Optimisation perf : ne SELECT que les colonnes utiles (la cle DTR, les
' fields Saisie du config + colonnes "extras" pour les regles condition-
' nelles et la liste des projets ProjetASL). Au lieu d'un SELECT * qui
' transfere ~50 colonnes pour EMMST, on reduit aux ~10-15 utiles.
'
' configCol  : la Collection SIRUS_LireConfig(table) pre-chargee. Utilisee
'              pour deduire la liste des colonnes a SELECT.
' sExtraCol  : nom d'une colonne supplementaire a inclure (ex.
'              "PROJET_PROJETASL" pour ProjetASL meme si non obligatoire).
' =========================================================================
Private Function ChargerRowsTableBdD_Optim(ByVal nomTable As String, _
                                            ByVal colonneDTR As String, _
                                            ByVal configCol As Collection, _
                                            Optional ByVal sExtraCol As String = "") As Object
    Dim dict As Object
    Dim sql As String
    Dim dtrVal As String
    Dim rowDict As Object
    Dim col As Collection
    Dim f As Object

    Dim def As SIRUS_Champ_Config
    Dim dictCols As Object
    Dim sCols As String
    Dim sColTrim As String

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    ' -- Construire la liste des colonnes a selectionner ----
    '    Tous les fields Saisie du config + cle DTR + extra eventuel.
    '    Dictionnaire pour eviter doublons.
    Set dictCols = CreateObject("Scripting.Dictionary")
    dictCols.CompareMode = vbTextCompare
    dictCols(colonneDTR) = True
    If LenB(sExtraCol) > 0 Then dictCols(sExtraCol) = True
    If Not configCol Is Nothing Then
        For Each def In configCol
            If LenB(def.ColonneBdD) > 0 And def.Source = "Saisie" Then
                ' Skip placeholders ??_*_a_definir (n'existent pas en BdD).
                If Left$(def.ColonneBdD, 3) <> "??_" Then
                    sColTrim = def.ColonneBdD
                    If Not dictCols.Exists(sColTrim) Then dictCols(sColTrim) = True
                End If
            End If
        Next def
    End If

    Dim k As Variant
    For Each k In dictCols.Keys
        If LenB(sCols) > 0 Then sCols = sCols & ", "
        sCols = sCols & "[" & CStr(k) & "]"
    Next k

    If LenB(sCols) = 0 Then sCols = "*"
    sql = "SELECT " & sCols & " FROM [" & nomTable & "]"

    ' Tentative 1 : SELECT cible. Fallback SELECT * si une colonne est
    ' rejetee par ACE (ex. nom BdD non aligne avec config).
    On Error Resume Next
    Call open_REQSEL(sql)
    If Err.Number <> 0 Then
        Debug.Print "[" & nomTable & "] SELECT cible echoue (" & Err.description & "), fallback SELECT *"
        Err.Clear
        On Error GoTo Sortie
        Call open_REQSEL("SELECT * FROM [" & nomTable & "]")
    Else
        On Error GoTo Sortie
    End If
    If REQSEL Is Nothing Then GoTo Sortie
    If REQSEL.EOF And REQSEL.BOF Then GoTo Sortie

    Do While Not REQSEL.EOF
        dtrVal = Trim$(NzSafeStr(REQSEL.Fields(colonneDTR).Value))
        If LenB(dtrVal) > 0 Then
            Set rowDict = CreateObject("Scripting.Dictionary")
            rowDict.CompareMode = vbTextCompare
            For Each f In REQSEL.Fields
                ' Ne stocke jamais Null (provoque "Utilisation incorrecte
                ' de Null" plus tard sur CStr/Trim). Vide = "".
                If IsNull(f.Value) Then
                    rowDict(f.Name) = ""
                Else
                    rowDict(f.Name) = f.Value
                End If
            Next f

            If dict.Exists(dtrVal) Then
                Set col = dict(dtrVal)
            Else
                Set col = New Collection
                dict.Add dtrVal, col
            End If
            col.Add rowDict
        End If
        REQSEL.MoveNext
    Loop

Sortie:
    If Err.Number <> 0 Then
        Debug.Print "[" & nomTable & "] ChargerRowsTableBdD_Optim : " & Err.description
    End If
    On Error Resume Next
    If Not REQSEL Is Nothing Then If REQSEL.State = 1 Then REQSEL.Close
    On Error GoTo 0
    Set ChargerRowsTableBdD_Optim = dict
End Function


' =========================================================================
' FONCTION: ProjetsConnusPourDTR
' Pour la colonne ProjetASL : compte les projets distincts dans Projet_ASL$
' ou ce DTR existe, et construit une chaine "ProjA, ProjB, ProjC" pour
' affichage en tooltip de ligne.
' Renvoie nProjets (ByRef) et sListe (ByRef).
' =========================================================================
Private Sub ProjetsConnusPourDTR(ByVal dictTable As Object, _
                                  ByVal dtr As String, _
                                  ByRef nProjets As Long, _
                                  ByRef sListe As String)
    Dim row As Object
    Dim rows As Collection
    Dim sProjet As String
    Dim dictDistincts As Object

    nProjets = 0
    sListe = ""

    If dictTable Is Nothing Then Exit Sub
    If Not dictTable.Exists(dtr) Then Exit Sub
    Set rows = dictTable(dtr)

    Set dictDistincts = CreateObject("Scripting.Dictionary")
    dictDistincts.CompareMode = vbTextCompare

    For Each row In rows
        If row.Exists("PROJET_PROJETASL") Then
            sProjet = Trim$(NzSafeStr(row("PROJET_PROJETASL")))
            If LenB(sProjet) > 0 And Not dictDistincts.Exists(sProjet) Then
                dictDistincts(sProjet) = True
                If LenB(sListe) > 0 Then sListe = sListe & vbLf
                sListe = sListe & "- " & sProjet
            End If
        End If
    Next row

    nProjets = dictDistincts.count
End Sub


' =========================================================================
' SUB: AjouterSubItemNbProjets
' Ajoute un ListSubItem affichant le nombre de projets connus :
'   nProjets = 0  -> "0" en rouge gras
'   nProjets >= 1 -> "N" en vert
' =========================================================================
Private Sub AjouterSubItemNbProjets(li As MSComctlLib.ListItem, ByVal nProjets As Long)
    Dim lsi As MSComctlLib.ListSubItem
    Set lsi = li.ListSubItems.Add(, , CStr(nProjets))
    If nProjets >= 1 Then
        lsi.ForeColor = RGB(0, 128, 0)
    Else
        lsi.ForeColor = RGB(192, 0, 0)
        lsi.Bold = True
    End If
End Sub


' =========================================================================
' FONCTION: CalculerMeQDTR_Optim (version optimisee)
' Recoit la configCol pre-chargee (au lieu d'appeler SIRUS_LireConfig
' a chaque DTR : economie 5 x N -> 5 appels au total).
'
' Renvoie :
'   -1   si DTR absent de la table (a afficher "Absent")
'   0..1 sinon (a multiplier par 100 pour pourcentage)
'
' Pour les tables a plusieurs rows par DTR (TachesATM), moyenne (somme
' numerateurs / somme denominateurs) sur toutes les rows.
' =========================================================================
Private Function CalculerMeQDTR_Optim(ByVal dictTable As Object, _
                                       ByVal dtr As String, _
                                       ByVal sTable As String, _
                                       ByVal configCol As Collection) As Double
    Dim col As Collection
    Dim def As SIRUS_Champ_Config
    Dim row As Object
    Dim rows As Collection
    Dim totNum As Double, totDen As Double
    Dim sVal As String
    Dim bRequis As Boolean

    If dictTable Is Nothing Then
        CalculerMeQDTR_Optim = -1
        Exit Function
    End If
    If Not dictTable.Exists(dtr) Then
        CalculerMeQDTR_Optim = -1
        Exit Function
    End If

    Set col = configCol
    If col Is Nothing Then Set col = SIRUS_LireConfig(sTable)
    If col.count = 0 Then
        CalculerMeQDTR_Optim = -1
        Exit Function
    End If

    Set rows = dictTable(dtr)

    For Each row In rows
        For Each def In col
            ' Skip champs sans ColonneBdD (info pure, heritages)
            If LenB(def.ColonneBdD) = 0 Then GoTo NextDef
            ' Skip champs Auto / Herite (pas saisis cote BdD)
            If def.Source = "Auto" Or def.Source = "Herite" Then GoTo NextDef

            bRequis = (def.Obligatoire = "Oui")
            If Not bRequis Then
                If def.Obligatoire = "Conditionnel" Then
                    bRequis = SIRUS_RegleObligatoire(sTable, def.ColonneBdD, row)
                End If
            End If

            If bRequis Then
                totDen = totDen + 1
                sVal = ""
                If row.Exists(def.ColonneBdD) Then sVal = NzSafeStr(row(def.ColonneBdD))
                If LenB(Trim$(sVal)) > 0 Then totNum = totNum + 1
            End If
NextDef:
        Next def
    Next row

    If totDen > 0 Then
        CalculerMeQDTR_Optim = totNum / totDen
    Else
        ' Aucune obligation -> considere 100% (cas tables sans champ
        ' obligatoire pour ce DTR dans le config).
        CalculerMeQDTR_Optim = 1
    End If
End Function


' =========================================================================
' SUB: AjouterSubItemMeQ
' Ajoute un ListSubItem au ListItem avec affichage % MeQ + coloration :
'   meq < 0  -> "Absent" rouge gras
'   meq = 1  -> "100%"   vert
'   sinon    -> "XX%"    rouge
' =========================================================================
Private Sub AjouterSubItemMeQ(li As MSComctlLib.ListItem, ByVal meq As Double)
    Dim lsi As MSComctlLib.ListSubItem
    Dim sTexte As String
    Dim lCouleur As Long

    If meq < 0 Then
        sTexte = "Absent"
        lCouleur = RGB(192, 0, 0)
    ElseIf meq >= 0.9999 Then
        sTexte = "100%"
        lCouleur = RGB(0, 128, 0)
    Else
        sTexte = Format$(meq * 100, "0") & "%"
        lCouleur = RGB(192, 0, 0)
    End If

    Set lsi = li.ListSubItems.Add(, , sTexte)
    lsi.ForeColor = lCouleur
    If lCouleur = RGB(192, 0, 0) Then lsi.Bold = True
End Sub


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
' SUB: AjouterSubItemColore
' Ajoute un ListSubItem au ListItem en parametre, avec coloration optionnelle.
'   - bColoreOuiNon : si True, applique vert (OUI) / rouge (NON) automatiquement.
'   - lColorOverride : si <> -1, force cette couleur (prioritaire sur bColoreOuiNon).
' Necessite d'utiliser ListSubItems.Add partout dans une meme ligne (pas
' melanger avec li.SubItems(n) =).
' =========================================================================
Private Sub AjouterSubItemColore(li As MSComctlLib.ListItem, _
                                  ByVal sTexte As String, _
                                  ByVal bColoreOuiNon As Boolean, _
                                  Optional ByVal lColorOverride As Long = -1)
    Dim lsi As MSComctlLib.ListSubItem
    Set lsi = li.ListSubItems.Add(, , sTexte)
    If lColorOverride <> -1 Then
        lsi.ForeColor = lColorOverride
        ' Gras sur les statuts critiques (NOK *).
        If Left$(sTexte, 3) = "NOK" Then lsi.Bold = True
    ElseIf bColoreOuiNon Then
        Select Case sTexte
            Case "OUI": lsi.ForeColor = RGB(0, 128, 0)        ' vert fonce
            Case "NON"
                lsi.ForeColor = RGB(192, 0, 0)                ' rouge
                lsi.Bold = True                                ' NON en gras
        End Select
    End If
End Sub


' =========================================================================
' FONCTION: CouleurQualite
' Couleur associee a un statut qualite prix (OK / NOK *) pour la ListView.
' Retourne -1 si pas de couleur particuliere (texte noir par defaut).
'
' Cas special : "NOK - Manque" avec sourceManque = "Fournisseur" seul est
' considere normal (le prix fournisseur n'est pas toujours connu en
' avance, c'est attendu). Pas de coloration ni de gras dans ce cas.
' =========================================================================
Private Function CouleurQualite(ByVal sQualite As String, _
                                 Optional ByVal sSourceManque As String = "") As Long
    ' Exception : manque uniquement du prix Fournisseur = cas normal
    If sQualite = "NOK - Manque" And sSourceManque = "Fournisseur" Then
        CouleurQualite = -1
        Exit Function
    End If

    Select Case sQualite
        Case "OK":                    CouleurQualite = RGB(0, 128, 0)        ' vert
        Case "NOK - Ecart":           CouleurQualite = RGB(255, 140, 0)      ' orange
        Case "NOK - Manque":          CouleurQualite = RGB(255, 140, 0)      ' orange
        Case "NOK - Manque & Ecart":  CouleurQualite = RGB(192, 0, 0)        ' rouge
        Case Else:                    CouleurQualite = -1
    End Select
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

