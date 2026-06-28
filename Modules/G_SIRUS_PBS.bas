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
    PrixFCS                       As Long
    PrixSpareCatalogue            As Long
    PrixSpareCalcule              As Long
    RefSourceSpare                As Long
    LeadTimeSpareCalend           As Long
    PrixRepairCatalogue           As Long
    RefSourceRepair               As Long
    LeadTimeRepairCalend          As Long
    CoutSpare_FRS_NotISP          As Long
    LeadTimeSpare_FRS             As Long
    CoutRepair_FRS_NotISP         As Long
    LeadTimeRepair_FRS            As Long
    CoutSpare_AlstomFR_FRS        As Long
    PrixRepair_Alstom_FRS         As Long
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
        .Designation_BdD = TrouverColonne(ws, "DESIGNATION (issue BdD)", False, manqueCritique, manqueOptionnel)
        .Designation_Providor = TrouverColonne(ws, "DESIGNATION (Providor)", False, manqueCritique, manqueOptionnel)
        .SousSysteme = TrouverColonne(ws, "SOUS SYST" & ChrW(200) & "ME", False, manqueCritique, manqueOptionnel)
        .Qte = TrouverColonne(ws, "Qte", False, manqueCritique, manqueOptionnel)
        .NodeType = TrouverColonne(ws, "Node typed (LRU/SRU" & ChrW(8230) & ")", False, manqueCritique, manqueOptionnel)
        .HardSoft = TrouverColonne(ws, "HARD/SOFT", False, manqueCritique, manqueOptionnel)
        .NTI = TrouverColonne(ws, "NTI", False, manqueCritique, manqueOptionnel)
        .RefPlan = TrouverColonne(ws, "R" & ChrW(233) & "f plan ensemble", False, manqueCritique, manqueOptionnel)

        ' --- RAM -------------------------------------------------------------
        .MTBF_Gen = TrouverColonne(ws, "MTBF g" & ChrW(233) & "n" & ChrW(233) & "rique (offre+PGR)", False, manqueCritique, manqueOptionnel)
        .TxDef_Gen = TrouverColonne(ws, "Tx d" & ChrW(233) & "faillance g" & ChrW(233) & "n" & ChrW(233) & "rique", False, manqueCritique, manqueOptionnel)
        .DureeTacheCorrPrev_Gen_HARD = TrouverColonne(ws, _
            "Dur" & ChrW(233) & "e t" & ChrW(226) & "che correctif+pr" & ChrW(233) & _
            "ventif HARD g" & ChrW(233) & "n" & ChrW(233) & "rique", _
            False, manqueCritique, manqueOptionnel)
        .MTBF_Intr = TrouverColonne(ws, "MTBF intrins" & ChrW(232) & "que RAM Report sp" & ChrW(233) & "cifique", False, manqueCritique, manqueOptionnel)
        .MTBF_Fonct = TrouverColonne(ws, "MTBF fonctionnel RAM Report sp" & ChrW(233) & "cifique", False, manqueCritique, manqueOptionnel)
        .TxDef_Fonct = TrouverColonne(ws, "Tx d" & ChrW(233) & "faillance fonctionnel", False, manqueCritique, manqueOptionnel)
        .DureeTacheCorrPrev_Ana_HARD = TrouverColonne(ws, _
            "Dur" & ChrW(233) & "e t" & ChrW(226) & "che correctif+pr" & ChrW(233) & _
            "ventif HARD analytique", _
            False, manqueCritique, manqueOptionnel)
        .UniteMaintCorrective = TrouverColonne(ws, "UNIT" & ChrW(201) & " maintenance corrective", False, manqueCritique, manqueOptionnel)

        ' --- Duree de vie / taches preventives -----------------------------
        ' DureeVie_Ana_* mappe la colonne "DVL <type> specifique" du nouveau
        ' template (= ce qui etait appele "analytique" / "RAM report specifique
        ' projet (CGR)" dans l'ancien). Saisie ASL.
        .DureeVie_Gen_HARD = TrouverColonne(ws, "DVL HARD g" & ChrW(233) & "n" & ChrW(233) & "rique", False, manqueCritique, manqueOptionnel)
        .DureeVie_Ana_HARD = TrouverColonne(ws, "DVL HARD sp" & ChrW(233) & "cifique", False, manqueCritique, manqueOptionnel)
        .DureeVie_Gen_SOFT = TrouverColonne(ws, "DVL SOFT g" & ChrW(233) & "n" & ChrW(233) & "rique", False, manqueCritique, manqueOptionnel)
        .DureeTachePrev_Gen_SOFT = TrouverColonne(ws, _
            "Dur" & ChrW(233) & "e t" & ChrW(226) & "che pr" & ChrW(233) & _
            "ventif SOFT g" & ChrW(233) & "n" & ChrW(233) & "rique", _
            False, manqueCritique, manqueOptionnel)
        .DureeVie_Ana_SOFT = TrouverColonne(ws, "DVL SOFT sp" & ChrW(233) & "cifique", False, manqueCritique, manqueOptionnel)
        .DureeTachePrev_Ana_SOFT = TrouverColonne(ws, _
            "Dur" & ChrW(233) & "e t" & ChrW(226) & "che pr" & ChrW(233) & _
            "ventif SOFT analytique", _
            False, manqueCritique, manqueOptionnel)
        .UniteMaintPreventive = TrouverColonne(ws, "UNIT" & ChrW(201) & " maintenance pr" & ChrW(233) & "ventive", False, manqueCritique, manqueOptionnel)

        ' --- Couts et prix --------------------------------------------------
        ' Critiques : les 4 prix necessaires aux cascades Spare/Repair de la
        ' ListView_Generation. Cascade Spare : Fournisseur -> Cat D&IS ->
        ' FCS calcule -> Archive. Cascade Repair : Fournisseur -> Cat D&IS ->
        ' Archive.
        .PrixFCS = TrouverColonne(ws, "Prix FCS", False, manqueCritique, manqueOptionnel)
        .PrixSpareCatalogue = TrouverColonne(ws, "Prix spare catalogue", True, manqueCritique, manqueOptionnel)
        .PrixSpareCalcule = TrouverColonne(ws, "Prix Spare calcul" & ChrW(233), True, manqueCritique, manqueOptionnel)
        .RefSourceSpare = TrouverColonne(ws, "Ref Source Spare", False, manqueCritique, manqueOptionnel)
        .LeadTimeSpareCalend = TrouverColonne(ws, "Lead Time SPARE (jours calendaires)", False, manqueCritique, manqueOptionnel)
        .PrixRepairCatalogue = TrouverColonne(ws, "Prix repair catalogue", True, manqueCritique, manqueOptionnel)
        .RefSourceRepair = TrouverColonne(ws, "Ref Source Repair", False, manqueCritique, manqueOptionnel)
        .LeadTimeRepairCalend = TrouverColonne(ws, "Lead Time REPAIR (jours calendaires)", False, manqueCritique, manqueOptionnel)
        .CoutSpare_FRS_NotISP = TrouverColonne(ws, "Cout spare FOURNISSEUR Not ISP", False, manqueCritique, manqueOptionnel)
        .LeadTimeSpare_FRS = TrouverColonne(ws, "Lead Time SPARE FOURNISSEUR", False, manqueCritique, manqueOptionnel)
        .CoutRepair_FRS_NotISP = TrouverColonne(ws, "Cout Repair FOURNISSEUR Not ISP", False, manqueCritique, manqueOptionnel)
        .LeadTimeRepair_FRS = TrouverColonne(ws, "Lead Time REPAIR FOURNISSEUR", False, manqueCritique, manqueOptionnel)
        .CoutSpare_AlstomFR_FRS = TrouverColonne(ws, "Cout spare ALSTOM FRANCE FOURNISSEUR", True, manqueCritique, manqueOptionnel)
        .PrixRepair_Alstom_FRS = TrouverColonne(ws, "Prix repair ALSTOM FOURNISSEUR", True, manqueCritique, manqueOptionnel)
        .PrixSpare_ProjetArchive = TrouverColonne(ws, "Prix Spare projet archive", False, manqueCritique, manqueOptionnel)
        .PrixRepair_ProjetArchive = TrouverColonne(ws, "Prix Repair projet archive", False, manqueCritique, manqueOptionnel)
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
'   Col 2  : Nb projets ProjetASL (count global ; tooltip = liste projets)
'   Col 3  : ATM-Info   - % MeQ
'   Col 4  : ATM-Tache  - % MeQ (moyenne sur toutes les taches du DTR)
'   Col 5  : EMMST      - % MeQ
'   Col 6  : Produit    - % MeQ
'   Col 7  : Prix Spare    (Format #,##0.00 ou vide)
'   Col 8  : Source Spare  ("Fournisseur" / Ref D&IS / "FCS" / "Projet : NOM")
'   Col 9  : Prix Repair   (Format #,##0.00 ou vide)
'   Col 10 : Source Repair ("Fournisseur" / Ref D&IS / "Projet : NOM")
'
' Cascade prix (priorite descendante) :
'   Spare  : Fournisseur -> Cat D&IS -> Cat FCS calcule (col AB) -> Archive
'   Repair : Fournisseur -> Cat D&IS                              -> Archive
'
' Source affichee = origine de la valeur prix retenue :
'   - Fournisseur : libelle en dur
'   - Cat D&IS    : contenu de Ref Source Spare (col AC) ou Repair (col AF)
'   - Cat FCS     : "FCS" en dur (Spare uniquement)
'   - Archive     : "Projet : <nom>" du projet alphabetiquement premier dans
'                   ProjetASL pour ce DTR, projet courant exclu
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
    Dim sProjetCourant As String
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

    ' Garantir que la connexion BdD est ouverte. RafraichirPrixArchive
    ' (G_PBS_Refresh) fait un Close_CONN a la fin -> Base = Nothing.
    ' Sans ce guard, ChargerRowsTableBdD_Optim plante silencieusement et
    ' la ListView affiche "Absent" / 0 partout apres un Refresh.
    If SIRUS_LookupDTR_Active Then
        On Error Resume Next
        If Base Is Nothing Then
            Call Open_CONN(Replace(CONNEXION_EXCEL_BDD_TRAVAIL, "UserProfile", Environ$("UserProfile")))
        ElseIf Base.State <> 1 Then
            Call Open_CONN(Replace(CONNEXION_EXCEL_BDD_TRAVAIL, "UserProfile", Environ$("UserProfile")))
        End If
        On Error GoTo Sortie
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

    derniereLigne = wsPBS.Cells(wsPBS.rows.count, DATA_ASL_type.DTR_PBS).End(xlUp).Row
    If derniereLigne <= DATA_ASL_HeaderRow Then Exit Sub

    ' Projet courant : alimente la cascade Archive (exclusion + libelle source)
    sProjetCourant = ""
    On Error Resume Next
    sProjetCourant = CStr(UserForm.Project_List.Value)
    On Error GoTo 0

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
                li.ToolTipText = "Projets connus (" & nProjets & ") :" & vbLf & sListeProjets
            Else
                li.ToolTipText = "Aucun projet connu pour ce DTR"
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

        ' Prix SPARE (Frs -> Cat D&IS -> FCS calcule -> Archive)
        resSpare = CalculerPrixSpare(wsPBS, ligne, dtr, sProjetCourant)
        Call AjouterSubItemPrix(li, resSpare(0))
        Call AjouterSubItemSource(li, CStr(resSpare(1)))

        ' Prix REPAIR (Frs -> Cat D&IS -> Archive)
        resRepair = CalculerPrixRepair(wsPBS, ligne, dtr, sProjetCourant)
        Call AjouterSubItemPrix(li, resRepair(0))
        Call AjouterSubItemSource(li, CStr(resRepair(1)))

SuivanteLigne:
    Next ligne

    Debug.Print "SIRUS_RemplirListViewGeneration : " & nbAjoutees & _
                " ligne(s) (Lookup=" & SIRUS_LookupDTR_Active & ")"
    Debug.Print "  Colonnes prix mappees (0 = introuvable) :"
    Debug.Print "    CoutSpare_AlstomFR_FRS = col " & DATA_ASL_type.CoutSpare_AlstomFR_FRS
    Debug.Print "    PrixSpareCatalogue     = col " & DATA_ASL_type.PrixSpareCatalogue
    Debug.Print "    PrixSpareCalcule       = col " & DATA_ASL_type.PrixSpareCalcule
    Debug.Print "    PrixSpare_ProjetArchive= col " & DATA_ASL_type.PrixSpare_ProjetArchive
    Debug.Print "    PrixRepair_Alstom_FRS  = col " & DATA_ASL_type.PrixRepair_Alstom_FRS
    Debug.Print "    PrixRepairCatalogue    = col " & DATA_ASL_type.PrixRepairCatalogue
    Debug.Print "    PrixRepair_ProjetArchive=col " & DATA_ASL_type.PrixRepair_ProjetArchive

Sortie:
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then
        Debug.Print "ERREUR SIRUS_RemplirListViewGeneration : " & Err.Description
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
        Debug.Print "[" & nomTable & "] ChargerSetDTR_BdD : " & Err.Description
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
' FONCTION: CalculerPrixSpare
' Applique la cascade Spare : Fournisseur -> Cat D&IS -> Cat FCS calcule
' -> Archive. Retourne Array(prix, source) ou prix peut etre Empty si
' aucune source n'est connue.
'
'   Fournisseur connu          -> (prix, "Fournisseur")
'   D&IS connu                 -> (prix, RefSourceSpare ou "Catalogue D&IS")
'   FCS calcule connu (col AB) -> (prix, RefSourceSpare ou "Catalogue FCS")
'   Archive connu              -> (prix, "Projet : <projet alpha 1er>")
'   Aucun                      -> (Empty, "")
'
' Note : la col AC (RefSourceSpare) est rempli par le Refresh selon une
' cascade D&IS -> FCS, donc on peut la lire pour les 2 branches D&IS et
' FCS sans avoir a discriminer ici. Si AA (vDIS) est vide mais AB (vFCS)
' non, AC contiendra la ref FCS.
' =========================================================================
Private Function CalculerPrixSpare(ByVal wsPBS As Worksheet, _
                                    ByVal ligne As Long, _
                                    ByVal dtr As String, _
                                    ByVal sProjetCourant As String) As Variant
    Dim vFrs As Variant, vDIS As Variant, vFCS As Variant, vArchive As Variant
    Dim sRefSpare As String
    Dim sProjetArchive As String

    vFrs = LireCellule(wsPBS, ligne, DATA_ASL_type.CoutSpare_AlstomFR_FRS)
    vDIS = LireCellule(wsPBS, ligne, DATA_ASL_type.PrixSpareCatalogue)
    vFCS = LireCellule(wsPBS, ligne, DATA_ASL_type.PrixSpareCalcule)
    vArchive = LireCellule(wsPBS, ligne, DATA_ASL_type.PrixSpare_ProjetArchive)
    sRefSpare = Trim$(NzSafeStr(LireCellule(wsPBS, ligne, DATA_ASL_type.RefSourceSpare)))

    If EstNumeriqueConnu(vFrs) Then
        CalculerPrixSpare = Array(CDbl(vFrs), "Fournisseur")
    ElseIf EstNumeriqueConnu(vDIS) Then
        CalculerPrixSpare = Array(CDbl(vDIS), IIf(LenB(sRefSpare) > 0, sRefSpare, "Catalogue D&IS"))
    ElseIf EstNumeriqueConnu(vFCS) Then
        CalculerPrixSpare = Array(CDbl(vFCS), IIf(LenB(sRefSpare) > 0, sRefSpare, "Catalogue FCS"))
    ElseIf EstNumeriqueConnu(vArchive) Then
        sProjetArchive = ProjetArchiveAlphaPremier(dtr, sProjetCourant)
        If LenB(sProjetArchive) > 0 Then
            CalculerPrixSpare = Array(CDbl(vArchive), "Projet : " & sProjetArchive)
        Else
            CalculerPrixSpare = Array(CDbl(vArchive), "Projet archive")
        End If
    Else
        CalculerPrixSpare = Array(Empty, "")
    End If
End Function


' =========================================================================
' FONCTION: CalculerPrixRepair
' Idem CalculerPrixSpare sans le maillon FCS (pas de prix Repair FCS dans
' la BdD).
'
'   Fournisseur connu -> (prix, "Fournisseur")
'   D&IS connu        -> (prix, RefSourceRepair ou "Catalogue D&IS")
'   Archive connu     -> (prix, "Projet : <projet alpha 1er>")
'   Aucun             -> (Empty, "")
' =========================================================================
Private Function CalculerPrixRepair(ByVal wsPBS As Worksheet, _
                                     ByVal ligne As Long, _
                                     ByVal dtr As String, _
                                     ByVal sProjetCourant As String) As Variant
    Dim vFrs As Variant, vDIS As Variant, vArchive As Variant
    Dim sRefDIS As String
    Dim sProjetArchive As String

    vFrs = LireCellule(wsPBS, ligne, DATA_ASL_type.PrixRepair_Alstom_FRS)
    vDIS = LireCellule(wsPBS, ligne, DATA_ASL_type.PrixRepairCatalogue)
    vArchive = LireCellule(wsPBS, ligne, DATA_ASL_type.PrixRepair_ProjetArchive)
    sRefDIS = Trim$(NzSafeStr(LireCellule(wsPBS, ligne, DATA_ASL_type.RefSourceRepair)))

    If EstNumeriqueConnu(vFrs) Then
        CalculerPrixRepair = Array(CDbl(vFrs), "Fournisseur")
    ElseIf EstNumeriqueConnu(vDIS) Then
        CalculerPrixRepair = Array(CDbl(vDIS), IIf(LenB(sRefDIS) > 0, sRefDIS, "Catalogue D&IS"))
    ElseIf EstNumeriqueConnu(vArchive) Then
        sProjetArchive = ProjetArchiveAlphaPremier(dtr, sProjetCourant)
        If LenB(sProjetArchive) > 0 Then
            CalculerPrixRepair = Array(CDbl(vArchive), "Projet : " & sProjetArchive)
        Else
            CalculerPrixRepair = Array(CDbl(vArchive), "Projet archive")
        End If
    Else
        CalculerPrixRepair = Array(Empty, "")
    End If
End Function


' =========================================================================
' FONCTION: ProjetArchiveAlphaPremier
' Pour un DTR donne, parcourt dictProjetASL_Rows (cache module-level
' alimente par SIRUS_RemplirListViewGeneration) et retourne le nom du
' projet alphabetiquement premier, projet courant exclu. "" si aucun
' projet connu.
' =========================================================================
Private Function ProjetArchiveAlphaPremier(ByVal dtr As String, _
                                            ByVal sProjetCourant As String) As String
    Dim rows As Collection
    Dim row As Object
    Dim sProjet As String
    Dim sBest As String

    If dictProjetASL_Rows Is Nothing Then Exit Function
    If Not dictProjetASL_Rows.Exists(dtr) Then Exit Function
    Set rows = dictProjetASL_Rows(dtr)

    For Each row In rows
        If row.Exists("PROJET_PROJETASL") Then
            sProjet = Trim$(NzSafeStr(row("PROJET_PROJETASL")))
            If LenB(sProjet) > 0 And UCase$(sProjet) <> UCase$(sProjetCourant) Then
                If LenB(sBest) = 0 Then
                    sBest = sProjet
                ElseIf StrComp(sProjet, sBest, vbTextCompare) < 0 Then
                    sBest = sProjet
                End If
            End If
        End If
    Next row

    ProjetArchiveAlphaPremier = sBest
End Function


' =========================================================================
' SUB: AjouterSubItemPrix
' Affiche un prix format #,##0.00, ou "-" en rouge gras si vide / non
' numerique (cas ou aucune source de la cascade n'est connue).
' =========================================================================
Private Sub AjouterSubItemPrix(li As MSComctlLib.ListItem, ByVal vPrix As Variant)
    Dim lsi As MSComctlLib.ListSubItem
    If EstNumeriqueConnu(vPrix) Then
        Set lsi = li.ListSubItems.Add(, , Format$(CDbl(vPrix), "#,##0.00"))
    Else
        Set lsi = li.ListSubItems.Add(, , "-")
        lsi.ForeColor = RGB(192, 0, 0)
        lsi.Bold = True
    End If
End Sub


' =========================================================================
' SUB: AjouterSubItemSource
' Colore la source selon son origine. Detection par mot-cle dans le libelle
' (qui peut etre une ref catalogue complete genre "D&IS 2026-2027_v03") :
'   Fournisseur          -> vert  (prio 1)
'   Projet : <nom>       -> orange(prio archive ; teste avant FCS pour eviter
'                                  qu'un nom de projet contenant "FCS" matche)
'   contient "FCS"       -> gris  (prio 3 Spare via col AB calcule)
'   contient "D&IS"      -> bleu  (prio 2 catalogue D&IS)
'   autre non vide       -> bleu  (catalogue D&IS par defaut)
'   ""                   -> rouge gras (aucune source)
' =========================================================================
Private Sub AjouterSubItemSource(li As MSComctlLib.ListItem, ByVal sSource As String)
    Dim lsi As MSComctlLib.ListSubItem
    Set lsi = li.ListSubItems.Add(, , sSource)

    If LenB(sSource) = 0 Then
        lsi.ForeColor = RGB(192, 0, 0)
        lsi.Bold = True
    ElseIf sSource = "Fournisseur" Then
        lsi.ForeColor = RGB(0, 128, 0)
    ElseIf Left$(sSource, 7) = "Projet " Then
        lsi.ForeColor = RGB(255, 140, 0)
    ElseIf InStr(1, sSource, "FCS", vbTextCompare) > 0 Then
        lsi.ForeColor = RGB(100, 100, 100)
    Else
        ' Catalogue D&IS (libelle Ref ou fallback)
        lsi.ForeColor = RGB(0, 0, 192)
    End If
End Sub


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
    Debug.Print "  CoutSpare_AlstomFR_FRS = " & DATA_ASL_type.CoutSpare_AlstomFR_FRS
    Debug.Print "  PrixSpareCatalogue     = " & DATA_ASL_type.PrixSpareCatalogue
    Debug.Print "  PrixSpareCalcule       = " & DATA_ASL_type.PrixSpareCalcule
    Debug.Print "  RefSourceSpare         = " & DATA_ASL_type.RefSourceSpare
    Debug.Print "  PrixSpare_ProjetArchive= " & DATA_ASL_type.PrixSpare_ProjetArchive
    Debug.Print "  PrixRepair_Alstom_FRS  = " & DATA_ASL_type.PrixRepair_Alstom_FRS
    Debug.Print "  PrixRepairCatalogue    = " & DATA_ASL_type.PrixRepairCatalogue
    Debug.Print "  RefSourceRepair        = " & DATA_ASL_type.RefSourceRepair
    Debug.Print "  PrixRepair_ProjetArchive=" & DATA_ASL_type.PrixRepair_ProjetArchive

End Sub


' =========================================================================
' SUB: Test_CalculerPrix
' Tests unitaires sur CalculerPrixSpare / CalculerPrixRepair via un mock
' de feuille temporaire. Ecrit 4 lignes (1 par cas de cascade) et verifie
' le label Source retourne. Affiche les resultats dans la fenetre
' Immediate.
' =========================================================================
Public Sub Test_CalculerPrix()
    Dim wb As Workbook, ws As Worksheet
    Dim r As Variant
    Dim sNomTmp As String

    sNomTmp = "_TestCalculerPrix"
    Set wb = ActiveWorkbook
    On Error Resume Next
    Application.DisplayAlerts = False
    wb.Worksheets(sNomTmp).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    Set ws = wb.Worksheets.Add
    ws.Name = sNomTmp

    ' Cadrage minimal : on aligne juste les colonnes utiles aux cascades.
    ' Lignes 1 = en-tete, lignes 2-5 = cas de test.
    ws.Cells(1, 1).Value = "DTR"
    ws.Cells(1, 2).Value = "Frs Spare"
    ws.Cells(1, 3).Value = "Cat DIS Spare"
    ws.Cells(1, 4).Value = "Cat FCS Calc"
    ws.Cells(1, 5).Value = "Archive Spare"
    ws.Cells(1, 6).Value = "Ref DIS"

    ws.Cells(2, 1).Value = "DTR-FRS"
    ws.Cells(2, 2).Value = 100
    ws.Cells(3, 1).Value = "DTR-DIS"
    ws.Cells(3, 3).Value = 110
    ws.Cells(3, 6).Value = "D&IS 2026-2027_v03"
    ws.Cells(4, 1).Value = "DTR-FCS"
    ws.Cells(4, 4).Value = 120
    ws.Cells(5, 1).Value = "DTR-VIDE"

    ' Mock du type DATA_ASL pour utiliser les cascades
    Dim sauv As DATA_ASL_Col
    sauv = DATA_ASL_type
    DATA_ASL_type.CoutSpare_AlstomFR_FRS = 2
    DATA_ASL_type.PrixSpareCatalogue = 3
    DATA_ASL_type.PrixSpareCalcule = 4
    DATA_ASL_type.PrixSpare_ProjetArchive = 5
    DATA_ASL_type.RefSourceSpare = 6

    r = CalculerPrixSpare(ws, 2, "DTR-FRS", "PROJ_X")
    Debug.Print "[Frs]  prix=" & r(0) & " | source=" & r(1)
    r = CalculerPrixSpare(ws, 3, "DTR-DIS", "PROJ_X")
    Debug.Print "[DIS]  prix=" & r(0) & " | source=" & r(1)
    r = CalculerPrixSpare(ws, 4, "DTR-FCS", "PROJ_X")
    Debug.Print "[FCS]  prix=" & r(0) & " | source=" & r(1)
    r = CalculerPrixSpare(ws, 5, "DTR-VIDE", "PROJ_X")
    Debug.Print "[Vide] prix=" & r(0) & " | source=" & r(1)

    DATA_ASL_type = sauv
    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = True
End Sub


