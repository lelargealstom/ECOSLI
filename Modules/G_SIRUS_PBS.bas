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

' Cache session : SnapShotPrix est charge une seule fois par session Excel
' (cout ~quelques secondes a cause du TypeGuessRows=0 qui scanne tout).
' Reset via SIRUS_ResetCacheSnapShotPrix (a appeler apres un refresh BdD
' ou changement de projet si necessaire).
Private m_dictSnapShotPrixCache As Object


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
        MsgBox "La feuille 'DATA_ASL' est introuvable dans le classeur '" & wb.Name & "'." & vbCrLf & "Ouvre un classeur projet contenant la PBS avant de lancer ECOSLI.", vbExclamation, "ECOSLI - PBS"
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

    ' Sur les libelles qui ont evolue entre l ancien template et le nouveau,
    ' on passe les 2 candidats a TrouverColonne (nouveau en prio, ancien
    ' en fallback). Le template actuel est en transition : seuls les 4 prix
    ' Catalogue/FCS ont ete renommes, le reste a encore les libelles longs.
    With DATA_ASL_type
        ' --- Identification --------------------------------------------------
        .DTR_PBS = TrouverColonne(ws, "DTR Reference (issue de PBS)", True, manqueCritique, manqueOptionnel)
        .Designation_PBS = TrouverColonne(ws, "DESIGNATION (issue de PBS)", False, manqueCritique, manqueOptionnel)
        .Designation_BdD = TrouverColonne(ws, Array("DESIGNATION (issue BdD)", "DESIGNATION (issue de BdD)"), False, manqueCritique, manqueOptionnel)
        .Designation_Providor = TrouverColonne(ws, Array("DESIGNATION (Providor)", "DESIGNATION (mise en forme pour copier dans Providor : construite a partir de celle de la PBS et de la designation issue de BdD (ORCHESTRA)))"), False, manqueCritique, manqueOptionnel)
        .SousSysteme = TrouverColonne(ws, Array("SOUS SYSTEME", "SOUS SYSTEME (copies de la PBS et classifies selon generiques)"), False, manqueCritique, manqueOptionnel)
        .Qte = TrouverColonne(ws, "Qte", False, manqueCritique, manqueOptionnel)
        .NodeType = TrouverColonne(ws, "Node typed (LRU/SRU...)", False, manqueCritique, manqueOptionnel)
        .HardSoft = TrouverColonne(ws, "HARD/SOFT", False, manqueCritique, manqueOptionnel)
        .NTI = TrouverColonne(ws, "NTI", False, manqueCritique, manqueOptionnel)
        .RefPlan = TrouverColonne(ws, Array("Ref plan ensemble", "ref du plan d'ensemble (issue de la BdD)"), False, manqueCritique, manqueOptionnel)

        ' --- RAM -------------------------------------------------------------
        .MTBF_Gen = TrouverColonne(ws, Array("MTBF generique (offre+PGR)", "MTBF generique (offre et PGR)"), False, manqueCritique, manqueOptionnel)
        .TxDef_Gen = TrouverColonne(ws, Array("Tx defaillance generique", "Tx deffaillance generique"), False, manqueCritique, manqueOptionnel)
        .DureeTacheCorrPrev_Gen_HARD = TrouverColonne(ws, Array("Duree tache correctif+preventif HARD generique", "Duree de la tache corrective et preventive generique (offre et PGR) HARDWARE"), False, manqueCritique, manqueOptionnel)
        .MTBF_Intr = TrouverColonne(ws, Array("MTBF intrinseque RAM Report specifique", "MTBF intrinseque (RAM report specifique projet (CGR))"), False, manqueCritique, manqueOptionnel)
        .MTBF_Fonct = TrouverColonne(ws, Array("MTBF fonctionnel RAM Report specifique", "MTBF fonctionnel (RAM report specifique projet (CGR))"), False, manqueCritique, manqueOptionnel)
        .TxDef_Fonct = TrouverColonne(ws, Array("Tx defaillance fonctionnel", "Tx deffaillance fonctionnel"), False, manqueCritique, manqueOptionnel)
        .DureeTacheCorrPrev_Ana_HARD = TrouverColonne(ws, Array("Duree tache correctif+preventif HARD analytique", "Duree de la tache corrective et preventive analytique (CGR) HARDWARE"), False, manqueCritique, manqueOptionnel)
        .UniteMaintCorrective = TrouverColonne(ws, Array("UNITE maintenance corrective", "UNITE de la maintenance corrective"), False, manqueCritique, manqueOptionnel)

        ' --- Duree de vie / taches preventives -----------------------------
        ' DureeVie_Ana_* mappe la colonne "DVL <type> specifique" du nouveau
        ' template (= ce qui etait appele "analytique" / "RAM report specifique
        ' projet (CGR)" dans l'ancien). Saisie ASL.
        .DureeVie_Gen_HARD = TrouverColonne(ws, Array("DVL HARD generique", "DUREE DE VIE LIMITEE generique (offre et PGR) HARDWARE"), False, manqueCritique, manqueOptionnel)
        .DureeVie_Ana_HARD = TrouverColonne(ws, Array("DVL HARD specifique", "DUREE DE VIE LIMITEE (RAM report specifique projet (CGR)) HARDWARE"), False, manqueCritique, manqueOptionnel)
        .DureeVie_Gen_SOFT = TrouverColonne(ws, Array("DVL SOFT generique", "DUREE DE VIE LIMITEE generique (offre et PGR) SOFTWARE"), False, manqueCritique, manqueOptionnel)
        .DureeTachePrev_Gen_SOFT = TrouverColonne(ws, Array("Duree tache preventif SOFT generique", "Duree de la tache preventive generique SOFTWARE (offre et PGR)"), False, manqueCritique, manqueOptionnel)
        .DureeVie_Ana_SOFT = TrouverColonne(ws, Array("DVL SOFT specifique", "DUREE DE VIE LIMITEE (RAM report specifique projet (CGR)) SOFTWARE"), False, manqueCritique, manqueOptionnel)
        .DureeTachePrev_Ana_SOFT = TrouverColonne(ws, Array("Duree tache preventif SOFT analytique", "Duree de la tache preventive analytique SOFTWARE (CGR)"), False, manqueCritique, manqueOptionnel)
        .UniteMaintPreventive = TrouverColonne(ws, Array("UNITE maintenance preventive", "UNITE de la maintenance preventive"), False, manqueCritique, manqueOptionnel)

        ' --- Couts et prix --------------------------------------------------
        ' Critiques : les 4 prix necessaires aux cascades Spare/Repair de la
        ' ListView_Generation. Cascade Spare : Fournisseur -> Cat D&IS ->
        ' FCS calcule -> Archive. Cascade Repair : Fournisseur -> Cat D&IS ->
        ' Archive.
        ' Les 4 renames PrixFCS / PrixSpareCatalogue / PrixSpareCalcule /
        ' PrixRepairCatalogue ont deja ete appliques au template par Thomas.
        ' Pas de fallback necessaire ici.
        .PrixFCS = TrouverColonne(ws, "Prix FCS", False, manqueCritique, manqueOptionnel)
        .PrixSpareCatalogue = TrouverColonne(ws, "Prix spare catalogue", True, manqueCritique, manqueOptionnel)
        .PrixSpareCalcule = TrouverColonne(ws, "Prix Spare calcule", True, manqueCritique, manqueOptionnel)
        .RefSourceSpare = TrouverColonne(ws, "Ref Source Spare", False, manqueCritique, manqueOptionnel)
        .LeadTimeSpareCalend = TrouverColonne(ws, Array("Lead Time SPARE (jours calendaires)", "Lead Time SPARE (en jours calendaires)"), False, manqueCritique, manqueOptionnel)
        .PrixRepairCatalogue = TrouverColonne(ws, "Prix repair catalogue", True, manqueCritique, manqueOptionnel)
        .RefSourceRepair = TrouverColonne(ws, "Ref Source Repair", False, manqueCritique, manqueOptionnel)
        .LeadTimeRepairCalend = TrouverColonne(ws, Array("Lead Time REPAIR (jours calendaires)", "Lead Time REPAIR (en jours calendaires)"), False, manqueCritique, manqueOptionnel)
        .CoutSpare_FRS_NotISP = TrouverColonne(ws, Array("Cout spare FOURNISSEUR Not ISP", "Cout spare FOURNISSEUR (Not ISP)"), False, manqueCritique, manqueOptionnel)
        .LeadTimeSpare_FRS = TrouverColonne(ws, "Lead Time SPARE FOURNISSEUR", False, manqueCritique, manqueOptionnel)
        .CoutRepair_FRS_NotISP = TrouverColonne(ws, Array("Cout Repair FOURNISSEUR Not ISP", "Cout Repair FOURNISSEUR (Not ISP)"), False, manqueCritique, manqueOptionnel)
        .LeadTimeRepair_FRS = TrouverColonne(ws, "Lead Time REPAIR FOURNISSEUR", False, manqueCritique, manqueOptionnel)
        .CoutSpare_AlstomFR_FRS = TrouverColonne(ws, "Cout spare ALSTOM FRANCE FOURNISSEUR", True, manqueCritique, manqueOptionnel)
        .PrixRepair_Alstom_FRS = TrouverColonne(ws, "Prix repair ALSTOM FOURNISSEUR", True, manqueCritique, manqueOptionnel)
        .PrixSpare_ProjetArchive = TrouverColonne(ws, Array("Prix Spare projet archive", "Prix Spare sur un projet archive"), False, manqueCritique, manqueOptionnel)
        .PrixRepair_ProjetArchive = TrouverColonne(ws, Array("Prix Repair projet archive", "Prix Repair sur un projet archive"), False, manqueCritique, manqueOptionnel)
    End With

    ' Colonnes optionnelles introuvables : silencieux (non bloquant, code gere
    ' l'absence). Decommente pour diagnostic si un cadrage part en vrille.
    ' If Len(manqueOptionnel) > 0 Then
    '     Debug.Print "[DATA_ASL] Colonnes optionnelles introuvables : " & manqueOptionnel
    ' End If

    If Len(manqueCritique) > 0 Then
        MsgBox "Cadrage DATA_ASL impossible." & vbCrLf & vbCrLf & "Colonnes critiques introuvables sur la ligne " & DATA_ASL_HeaderRow & " :" & vbCrLf & manqueCritique, vbCritical, "ECOSLI - PBS"
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
Private Function TrouverColonne(ByVal ws As Worksheet, ByVal libelleOrArray As Variant, ByVal critique As Boolean, ByRef manqueCritique As String, ByRef manqueOptionnel As String) As Long
    ' Accepte libelleOrArray = String (1 libelle) OU Array(...) de String
    ' (plusieurs libelles candidats, essayes en ordre). Permet de gerer la
    ' coexistence entre l ancien template (libelles longs) et le nouveau
    ' (libelles courts) sans devoir choisir l un ou l autre.

    Dim libelles As Variant
    Dim sLibelle As Variant
    Dim cel As Range
    Dim sPremier As String

    If IsArray(libelleOrArray) Then
        libelles = libelleOrArray
    Else
        libelles = Array(CStr(libelleOrArray))
    End If
    sPremier = CStr(libelles(LBound(libelles)))

    For Each sLibelle In libelles
        Set cel = Nothing
        On Error Resume Next
        ' Match exact d abord (xlWhole)
        Set cel = ws.rows(DATA_ASL_HeaderRow).Find(What:=CStr(sLibelle), LookIn:=xlValues, lookat:=xlWhole, MatchCase:=False)
        ' Fallback xlPart - tolere espaces / sauts de ligne / suffixes
        If cel Is Nothing Then
            Set cel = ws.rows(DATA_ASL_HeaderRow).Find(What:=CStr(sLibelle), LookIn:=xlValues, lookat:=xlPart, MatchCase:=False)
        End If
        On Error GoTo 0
        If Not cel Is Nothing Then
            TrouverColonne = cel.Column
            Exit Function
        End If
    Next sLibelle

    ' Aucun candidat trouve
    TrouverColonne = 0
    If critique Then
        manqueCritique = manqueCritique & "  - " & sPremier & vbCrLf
    Else
        manqueOptionnel = manqueOptionnel & sPremier & " ; "
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
    Set cel = ws.Cells.Find(What:=libelleAncre, LookIn:=xlValues, lookat:=xlWhole, MatchCase:=False, SearchOrder:=xlByRows, SearchDirection:=xlNext)
    If cel Is Nothing Then
        Set cel = ws.Cells.Find(What:=libelleAncre, LookIn:=xlValues, lookat:=xlPart, MatchCase:=False, SearchOrder:=xlByRows, SearchDirection:=xlNext)
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

        ' Enrichissement des rows Produit avec les 7 champs virtuels
        ' provenant de Produit_SnapShotPrix (refonte 2026-07-07 : les
        ' prix / dates / MoQ / lead times sont maintenant historises hors
        ' de Produit$). Le config Produit garde les 7 ColonneBdD comme
        ' champs virtuels ; c'est ici qu'on injecte la valeur SnapShotPrix
        ' la plus recente (Prix non nul, MAX Date Source) dans chaque row
        ' Produit avant appel a CalculerMeQDTR_Optim.
        Dim dictSnapShotPrix As Object
        Set dictSnapShotPrix = ChargerSnapShotPrixDernierParDTR()
        Call EnrichirProduitAvecSnapShot(dictProduit_Rows, dictSnapShotPrix)
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

    Debug.Print "SIRUS_RemplirListViewGeneration : " & nbAjoutees & " ligne(s) (Lookup=" & SIRUS_LookupDTR_Active & ")"
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
Private Function ChargerRowsTableBdD_Optim(ByVal nomTable As String, ByVal colonneDTR As String, ByVal configCol As Collection, Optional ByVal sExtraCol As String = "") As Object
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
' FONCTION: ChargerSnapShotPrixDernierParDTR
' =========================================================================
' Charge Produit_SnapShotPrix$ et renvoie, par DTR, les valeurs a injecter
' comme champs virtuels dans la row Produit (7 ColonneBdD).
'
' Regle metier (2026-07-07) : pour chaque couple (DTR, Type in {Spare, Repair}),
' on retient la ligne dont [Prix] est non nul et dont [Date Source] est la
' plus recente. Si aucune ligne avec Prix non nul n'existe pour ce (DTR,Type),
' aucun champ virtuel n'est cree -> le row Produit reste vide sur ce champ
' -> CalculerMeQDTR_Optim compte le champ comme non rempli.
'
' Retour : Dictionary {DTR (trim) -> Dictionary {ColonneBdD_virtuelle -> valeur}}
' Cle compareMode = vbTextCompare pour aligner avec dictProduit_Rows.
'
' Cast [DTR] & '' AS DTR_str : evite le bug ACE OLE DB avec DTR mixed-type
' (numerique + alphanumerique). Cf. Test_DryRunRafraichirPrixCatalogue.
Public Sub SIRUS_ResetCacheSnapShotPrix()
    ' Vide le cache SnapShotPrix. A appeler si Produit_SnapShotPrix a ete
    ' modifiee entre 2 refresh (nouvelle ligne, nouveau catalogue...). Declaration
    ' du cache m_dictSnapShotPrixCache dans la section General en tete de module.
    Set m_dictSnapShotPrixCache = Nothing
End Sub

Public Function ChargerSnapShotPrixDernierParDTR() As Object
    ' Refonte 2026-07-08 : lecture directe de la feuille Produit_SnapShotPrix
    ' (BdD ouverte read-only si necessaire) + cache session.
    '
    ' Historique du choix :
    '   - v1 : ADODB direct -> bug ACE DTR mixed-type (DTR alphanumeriques
    '     reviennent NULL car ACE type la colonne Numeric sur les 8 premieres
    '     lignes qui sont numeriques).
    '   - v2 : ADODB + IMEX=1 + TypeGuessRows=0 -> testé chez Thomas, ACE ne
    '     respecte pas TypeGuessRows sur cette version -> 32 DTR remontent
    '     (uniquement les numeriques), les alphanumeriques manquent.
    '   - v3 (retenue) : lecture directe de la feuille (bulk array). Cost :
    '     ouvre la BdD si elle n'est pas deja ouverte (Refresh ou manuel).
    '     Cache session -> penalite payee au 1er acces uniquement.
    Dim dict As Object, dictBest As Object
    Dim wbBdD As Workbook, wsBdD As Worksheet
    Dim wasOpen As Boolean
    Dim arr As Variant
    Dim derLigne As Long, ligne As Long
    Dim colDTR As Long, colType As Long, colPrix As Long, colLT As Long
    Dim colOnReqPrix As Long, colOnReqLT As Long, colMoQ As Long, colRefCat As Long, colDate As Long
    Dim maxCol As Long
    Dim sDTR As String, sType As String
    Dim vDate As Variant, vPrix As Variant, vLT As Variant, vMoQ As Variant, vRefCat As Variant
    Dim vOnReqPrix As Variant, vOnReqLT As Variant
    Dim sKey As String, bPrixConnu As Boolean, bOnRequest As Boolean
    Dim conserver As Boolean

    ' Cache hit : renvoie directement.
    If Not m_dictSnapShotPrixCache Is Nothing Then
        Set ChargerSnapShotPrixDernierParDTR = m_dictSnapShotPrixCache
        Exit Function
    End If

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare
    Set dictBest = CreateObject("Scripting.Dictionary")
    dictBest.CompareMode = vbTextCompare

    On Error GoTo Sortie
    If Not OuvrirFeuilleSnapShotPrix(wbBdD, wsBdD, wasOpen) Then
        Debug.Print "[Produit_SnapShotPrix] Impossible d'ouvrir la feuille (BdD introuvable ou verouillee)."
        GoTo Sortie
    End If

    ' Mapper les colonnes par libelle en ligne 1.
    colDTR = TrouverColParLibelleSirus(wsBdD, "DTR")
    colType = TrouverColParLibelleSirus(wsBdD, "Type")
    colPrix = TrouverColParLibelleSirus(wsBdD, "Prix")
    colLT = TrouverColParLibelleSirus(wsBdD, "LeadTime")
    colOnReqPrix = TrouverColParLibelleSirus(wsBdD, "OnRequestPrix")
    colOnReqLT = TrouverColParLibelleSirus(wsBdD, "OnRequestLT")
    colMoQ = TrouverColParLibelleSirus(wsBdD, "MoQ")
    colRefCat = TrouverColParLibelleSirus(wsBdD, "Ref Catalogue")
    colDate = TrouverColParLibelleSirus(wsBdD, "Date Source")
    If colDTR = 0 Or colType = 0 Or colPrix = 0 Or colDate = 0 Then
        Debug.Print "[Produit_SnapShotPrix] Colonnes essentielles manquantes (DTR/Type/Prix/Date Source)."
        GoTo CleanupBdD
    End If

    derLigne = wsBdD.Cells(wsBdD.rows.count, colDTR).End(xlUp).Row
    If derLigne < 2 Then GoTo CleanupBdD

    maxCol = Application.WorksheetFunction.Max(colDTR, colType, colPrix, colLT, colOnReqPrix, colOnReqLT, colMoQ, colRefCat, colDate)
    arr = wsBdD.Range(wsBdD.Cells(2, 1), wsBdD.Cells(derLigne, maxCol)).Value

    ' 1er passage : pour chaque (DTR,Type) on garde la ligne candidat la
    ' plus recente (MAX Date Source). Ligne candidate = Prix non nul OU
    ' OnRequestPrix renseigne.
    For ligne = 1 To UBound(arr, 1)
        sDTR = Trim$(NzSafeStr(arr(ligne, colDTR)))
        If Right$(sDTR, 2) = ".0" Then sDTR = Left$(sDTR, Len(sDTR) - 2)

        If LenB(sDTR) > 0 Then
            sType = UCase$(Trim$(NzSafeStr(arr(ligne, colType))))
            vPrix = arr(ligne, colPrix)
            vDate = arr(ligne, colDate)
            vLT = arr(ligne, colLT)
            If colOnReqPrix > 0 Then vOnReqPrix = arr(ligne, colOnReqPrix) Else vOnReqPrix = Empty
            If colOnReqLT > 0 Then vOnReqLT = arr(ligne, colOnReqLT) Else vOnReqLT = Empty
            If colMoQ > 0 Then vMoQ = arr(ligne, colMoQ) Else vMoQ = Empty
            vRefCat = arr(ligne, colRefCat)

            bPrixConnu = EstNumeriqueNonNul(vPrix)
            bOnRequest = (LenB(Trim$(NzSafeStr(vOnReqPrix))) > 0)

            If (sType = "SPARE" Or sType = "REPAIR") And IsDate(vDate) Then
                If bPrixConnu Or bOnRequest Then
                    sKey = sDTR & "|" & sType
                    conserver = True
                    If dictBest.Exists(sKey) Then
                        Dim prev As Variant
                        prev = dictBest(sKey)
                        If CDate(prev(0)) >= CDate(vDate) Then conserver = False
                    End If
                    If conserver Then
                        ' Array : 0=Date, 1=Prix, 2=LT, 3=MoQ, 4=RefCat,
                        '         5=OnReqPrix, 6=OnReqLT.
                        dictBest(sKey) = Array(vDate, vPrix, vLT, vMoQ, vRefCat, vOnReqPrix, vOnReqLT)
                    End If
                End If
            End If
        End If
    Next ligne

    ' 2nd passage : injecte les valeurs retenues dans le dict resultat.
    ' Prix : ecrit uniquement si numerique non nul (OnRequest seul -> champ
    ' virtuel Prix reste absent -> MeQ perd 1 point).
    Dim key As Variant
    For Each key In dictBest.Keys
        Dim parts() As String
        parts = Split(CStr(key), "|")
        Dim dtrKey As String, typKey As String
        dtrKey = parts(0)
        typKey = parts(1)
        Dim r As Variant
        r = dictBest(key)

        If Not dict.Exists(dtrKey) Then
            Dim d As Object
            Set d = CreateObject("Scripting.Dictionary")
            d.CompareMode = vbTextCompare
            dict.Add dtrKey, d
        End If

        If typKey = "SPARE" Then
            If EstNumeriqueNonNul(r(1)) Then dict(dtrKey)("Prix SPARE catalogue a date_PRODUIT") = r(1)
            dict(dtrKey)("Dates de validite des prix SPARE connus sur Projet (ASL)_PRODUIT") = r(4)
            dict(dtrKey)("MoQ spare catalogue_PRODUIT") = r(3)
            dict(dtrKey)("Lead time SPARE catalogue (jours) a date_PRODUIT") = r(2)
        ElseIf typKey = "REPAIR" Then
            If EstNumeriqueNonNul(r(1)) Then dict(dtrKey)("Prix REPAIR catalogue a date_PRODUIT") = r(1)
            dict(dtrKey)("Dates de validite des prix REPAIR connus sur Projet (ASL)_PRODUIT") = r(4)
            dict(dtrKey)("Lead time REPAIR (jours) manuel_PRODUIT") = r(2)
        End If
    Next key

CleanupBdD:
    Call FermerBdDSiOuverte(wbBdD, wasOpen)

Sortie:
    If Err.Number <> 0 Then
        Debug.Print "[Produit_SnapShotPrix] ChargerSnapShotPrixDernierParDTR : " & Err.description
        Call FermerBdDSiOuverte(wbBdD, wasOpen)
    End If

    ' Cache le resultat si la charge a reussi. Sinon m_dictSnapShotPrixCache
    ' reste Nothing -> nouvelle tentative au prochain appel.
    If dict.count > 0 Then Set m_dictSnapShotPrixCache = dict
    Set ChargerSnapShotPrixDernierParDTR = dict
End Function

' Helpers prives pour la lecture directe de Produit_SnapShotPrix (bypass ADODB).
' Detection workbook deja ouvert : compare par nom de fichier (pas FullName)
' pour supporter le cas OneDrive/SharePoint ou Office 365 redirige le chemin
' local vers une URL sharepoint.com -> FullName != chemin local.
Private Function OuvrirFeuilleSnapShotPrix(ByRef wbBdD As Workbook, ByRef wsBdD As Worksheet, ByRef wasOpen As Boolean) As Boolean
    Dim sCheminBdD As String
    Dim sNomFichier As String
    Dim wb As Workbook
    Dim posSlash As Long

    sCheminBdD = ExtraireCheminBdDSirus()
    If LenB(sCheminBdD) = 0 Then Exit Function

    ' Extrait le nom de fichier seul (ex "BDD excel original V4.1.xlsb").
    posSlash = InStrRev(sCheminBdD, "\")
    If posSlash > 0 Then
        sNomFichier = Mid$(sCheminBdD, posSlash + 1)
    Else
        sNomFichier = sCheminBdD
    End If

    wasOpen = False
    For Each wb In Application.Workbooks
        If StrComp(wb.Name, sNomFichier, vbTextCompare) = 0 Then
            Set wbBdD = wb
            wasOpen = True
            Exit For
        End If
    Next wb
    If wbBdD Is Nothing Then
        ' Tentative avec le chemin local. Si Office 365 redirige vers
        ' l'URL SharePoint et l'echec, on remontre l'erreur au caller
        ' (via On Error du caller) : la BdD doit alors etre ouverte
        ' manuellement par l'utilisateur avant relance.
        Set wbBdD = Application.Workbooks.Open(sCheminBdD, UpdateLinks:=0, ReadOnly:=True, Notify:=False, IgnoreReadOnlyRecommended:=True)
    End If
    If wbBdD Is Nothing Then Exit Function

    On Error Resume Next
    Set wsBdD = wbBdD.Worksheets("Produit_SnapShotPrix")
    On Error GoTo 0
    If wsBdD Is Nothing Then
        If Not wasOpen Then wbBdD.Close SaveChanges:=False
        Exit Function
    End If

    OuvrirFeuilleSnapShotPrix = True
End Function

Private Sub FermerBdDSiOuverte(ByRef wbBdD As Workbook, ByVal wasOpen As Boolean)
    If Not wasOpen And Not wbBdD Is Nothing Then
        On Error Resume Next
        wbBdD.Close SaveChanges:=False
        On Error GoTo 0
    End If
End Sub

Private Function ExtraireCheminBdDSirus() As String
    Dim cnx As String
    Dim posDS As Long, posEnd As Long
    cnx = Replace(CONNEXION_EXCEL_BDD_TRAVAIL, "UserProfile", Environ$("UserProfile"))
    posDS = InStr(1, cnx, "Data Source=", vbTextCompare)
    If posDS = 0 Then Exit Function
    posDS = posDS + Len("Data Source=")
    posEnd = InStr(posDS, cnx, ";")
    If posEnd = 0 Then posEnd = Len(cnx) + 1
    ExtraireCheminBdDSirus = Mid$(cnx, posDS, posEnd - posDS)
End Function

Private Function TrouverColParLibelleSirus(ByVal ws As Worksheet, ByVal sLibelle As String) As Long
    Dim cell As Range
    On Error Resume Next
    Set cell = ws.Rows(1).Find(What:=sLibelle, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    If Not cell Is Nothing Then TrouverColParLibelleSirus = cell.Column
    On Error GoTo 0
End Function

' Helper local : True si v est numerique et non nul (ni Empty ni Null ni 0).
' Duplique le helper prive de G_PBS_Refresh (pas d'export public a maintenir).
Private Function EstNumeriqueNonNul(ByVal v As Variant) As Boolean
    If IsEmpty(v) Then Exit Function
    If IsNull(v) Then Exit Function
    If IsError(v) Then Exit Function
    If Not IsNumeric(v) Then Exit Function
    If CDbl(v) = 0 Then Exit Function
    EstNumeriqueNonNul = True
End Function

' =========================================================================
' SUB: EnrichirProduitAvecSnapShot
' =========================================================================
' Injecte les 7 champs virtuels (issus de Produit_SnapShotPrix) dans chaque
' row du dict Produit. Si le DTR n'a pas de ligne SnapShotPrix retenue, les
' champs restent absents du row -> CalculerMeQDTR_Optim les compte comme
' vides -> MeQ tombe (comportement attendu : incite a alimenter SnapShotPrix).
' Ecrit toujours (overwrite) : SnapShotPrix est la source de verite.
Private Sub EnrichirProduitAvecSnapShot(ByVal dictProduit As Object, ByVal dictSnapShot As Object)
    Dim dtrKey As Variant
    Dim rows As Collection
    Dim row As Object
    Dim colVirt As Variant

    If dictProduit Is Nothing Or dictSnapShot Is Nothing Then Exit Sub

    For Each dtrKey In dictProduit.Keys
        If dictSnapShot.Exists(CStr(dtrKey)) Then
            Set rows = dictProduit(dtrKey)
            For Each row In rows
                For Each colVirt In dictSnapShot(CStr(dtrKey)).Keys
                    row(CStr(colVirt)) = dictSnapShot(CStr(dtrKey))(CStr(colVirt))
                Next colVirt
            Next row
        End If
    Next dtrKey
End Sub


' =========================================================================
' FONCTION: ProjetsConnusPourDTR
' Pour la colonne ProjetASL : compte les projets distincts dans Projet_ASL$
' ou ce DTR existe, et construit une chaine "ProjA, ProjB, ProjC" pour
' affichage en tooltip de ligne.
' Renvoie nProjets (ByRef) et sListe (ByRef).
' =========================================================================
Private Sub ProjetsConnusPourDTR(ByVal dictTable As Object, ByVal dtr As String, ByRef nProjets As Long, ByRef sListe As String)
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
Private Function CalculerMeQDTR_Optim(ByVal dictTable As Object, ByVal dtr As String, ByVal sTable As String, ByVal configCol As Collection) As Double
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
Private Sub AjouterSubItemColore(li As MSComctlLib.ListItem, ByVal sTexte As String, ByVal bColoreOuiNon As Boolean, Optional ByVal lColorOverride As Long = -1)
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
' True si la valeur est exploitable comme prix : non vide, non erreur,
' numerique, ET different de 0.
'
' 0 est explicitement EXCLU (modif 2026-06-28) : un prix Spare/Repair a 0
' represente metier "prix inconnu". Sans cette exclusion, la cascade
' Frs -> D&IS -> FCS -> Archive s arrete prematurement quand une cellule
' source vaut 0 (saisie zero, ou ecriture defensive) au lieu de continuer
' jusqu a une source vraiment connue.
' =========================================================================
Private Function EstNumeriqueConnu(ByVal v As Variant) As Boolean
    If IsError(v) Then Exit Function
    If IsMissing(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    If IsNull(v) Then Exit Function
    If Trim$(CStr(v)) = "" Then Exit Function
    If Not IsNumeric(v) Then Exit Function
    If CDbl(v) = 0 Then Exit Function
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
Private Function CalculerPrixSpare(ByVal wsPBS As Worksheet, ByVal ligne As Long, ByVal dtr As String, ByVal sProjetCourant As String) As Variant
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
Private Function CalculerPrixRepair(ByVal wsPBS As Worksheet, ByVal ligne As Long, ByVal dtr As String, ByVal sProjetCourant As String) As Variant
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
Private Function ProjetArchiveAlphaPremier(ByVal dtr As String, ByVal sProjetCourant As String) As String
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


' =========================================================================
' Test_SIRUS_DebugMeQProduit_PremierDTR
' =========================================================================
' Debug diagnostic pour le 1er DTR de DATA_ASL du ActiveWorkbook.
' Affiche dans la fenetre Immediate :
'   * Toutes les lignes brutes Produit_SnapShotPrix pour ce DTR
'   * La ligne retenue (Prix non nul OU OnRequest, MAX Date Source) par Type
'   * Le row Produit brut (avant merge)
'   * Le row Produit apres merge SnapShotPrix
'   * Pour chacun des 19 champs MeQ : Obligatoire (Oui/Cond/Non) + Rempli + Valeur
'   * Total nbObl / nbRempli / TauxMeQ
Public Sub Test_SIRUS_DebugMeQProduit_PremierDTR()
    Dim wb As Workbook, wsPBS As Worksheet
    Dim dtr As String
    Dim cnx As String
    Dim cfgProduit As Collection
    Dim dictProduit As Object, dictSnap As Object
    Dim row As Object
    Dim def As SIRUS_Champ_Config
    Dim champs As Variant, cle As Variant
    Dim nbObl As Long, nbRempli As Long
    Dim s As String, val As String
    Dim oblig As Boolean, rempli As Boolean, isSaisieDansConfig As Boolean

    Set wb = ActiveWorkbook
    If wb Is Nothing Then Debug.Print "Pas de ActiveWorkbook" : Exit Sub
    If Not VerifierFeuilleDATA_ASL(wb, wsPBS) Then Debug.Print "Pas de DATA_ASL" : Exit Sub
    If Not MapColonnesDATA_ASL(wsPBS) Then Debug.Print "MapColonnes echec" : Exit Sub
    If DATA_ASL_type.DTR_PBS = 0 Then Debug.Print "Colonne DTR introuvable" : Exit Sub

    dtr = Trim$(CStr(wsPBS.Cells(DATA_ASL_HeaderRow + 1, DATA_ASL_type.DTR_PBS).Value))
    If LenB(dtr) = 0 Then Debug.Print "1er DTR vide" : Exit Sub

    Debug.Print String(70, "=")
    Debug.Print "Debug MeQ Produit pour DTR = '" & dtr & "'"
    Debug.Print String(70, "=")

    ' Ouvrir la connexion BdD si besoin
    cnx = Replace(CONNEXION_EXCEL_BDD_TRAVAIL, "UserProfile", Environ$("UserProfile"))
    On Error Resume Next
    If Base Is Nothing Then
        Call Open_CONN(cnx)
    ElseIf Base.State <> 1 Then
        Call Open_CONN(cnx)
    End If
    On Error GoTo 0

    ' -- 1. Lignes brutes Produit_SnapShotPrix pour ce DTR ---
    Debug.Print vbLf & "-- Lignes brutes Produit_SnapShotPrix pour ce DTR --"
    Call DumpSnapShotBrutPourDTR(dtr)

    ' -- 2. Ligne retenue par ChargerSnapShotPrixDernierParDTR ---
    Set dictSnap = ChargerSnapShotPrixDernierParDTR()
    Debug.Print vbLf & "-- Ligne retenue (dernier par Type, Prix non nul OU OnRequest) --"
    If dictSnap.Exists(dtr) Then
        Dim cKey As Variant
        For Each cKey In dictSnap(dtr).Keys
            Debug.Print "  " & CStr(cKey) & " = " & NzSafeStr(dictSnap(dtr)(cKey))
        Next cKey
    Else
        Debug.Print "  (aucune donnee SnapShotPrix retenue pour ce DTR)"
    End If

    ' -- 3. Row Produit brut ---
    Set cfgProduit = SIRUS_LireConfig("Produit")
    Set dictProduit = ChargerRowsTableBdD_Optim("Produit$", "DTR", cfgProduit)

    Debug.Print vbLf & "-- Row Produit brut (avant merge SnapShotPrix) --"
    If Not dictProduit.Exists(dtr) Then
        Debug.Print "  !! DTR '" & dtr & "' absent de Produit$. Voici 5 exemples de cles :"
        Dim iEx As Long
        For Each cle In dictProduit.Keys
            iEx = iEx + 1 : If iEx > 5 Then Exit For
            Debug.Print "    '" & cle & "'"
        Next cle
        Exit Sub
    End If
    Dim rowsProd As Collection
    Set rowsProd = dictProduit(dtr)
    Debug.Print "  Produit$ row count : " & rowsProd.count
    Set row = rowsProd(1)
    Dim k As Variant
    For Each k In row.Keys
        Debug.Print "  " & CStr(k) & " = " & NzSafeStr(row(k))
    Next k

    ' -- 4. Row apres merge ---
    Call EnrichirProduitAvecSnapShot(dictProduit, dictSnap)
    Set row = dictProduit(dtr)(1)

    Debug.Print vbLf & "-- 19 champs MeQ (Obligatoire=Ob | Rempli=V) --"
    Dim sColNT As String
    sColNT = "Node type Produit (LRU/SRU" & ChrW(&H2026) & ")_PRODUIT"
    champs = Array( _
        "DTR", _
        "Designation_PRODUIT", _
        sColNT, _
        "Fournisseur_PRODUIT", _
        "Reference fournisseur_PRODUIT", _
        "MotifNonApplicabilite_PRODUIT", _
        "Motif_NonApplicabilitePreventif_PRODUIT", _
        "Drawing number", _
        "MTBF intrinseque Produit (h) issu de la datasheet _PRODUIT", _
        "Duree de vie generique Produit HARDWARE (annee)_PRODUIT", _
        "Duree de vie generique Produit (annee) SOFTWARE _PRODUIT", _
        "Reference MUL maintenance_PRODUIT", _
        "Prix SPARE catalogue a date_PRODUIT", _
        "Dates de validite des prix SPARE connus sur Projet (ASL)_PRODUIT", _
        "MoQ spare catalogue_PRODUIT", _
        "Lead time SPARE catalogue (jours) a date_PRODUIT", _
        "Prix REPAIR catalogue a date_PRODUIT", _
        "Dates de validite des prix REPAIR connus sur Projet (ASL)_PRODUIT", _
        "Lead time REPAIR (jours) manuel_PRODUIT")

    For Each cle In champs
        s = CStr(cle)

        ' Trouver Obligatoire via config
        oblig = False
        isSaisieDansConfig = False
        For Each def In cfgProduit
            If def.ColonneBdD = s Then
                isSaisieDansConfig = (def.Source = "Saisie")
                If def.Obligatoire = "Oui" Then
                    oblig = True
                ElseIf def.Obligatoire = "Conditionnel" Then
                    oblig = SIRUS_RegleObligatoire("Produit", s, row)
                End If
                Exit For
            End If
        Next def

        ' Valeur dans le row (merged)
        val = ""
        If row.Exists(s) Then val = NzSafeStr(row(s))
        rempli = (LenB(Trim$(val)) > 0)

        If oblig Then nbObl = nbObl + 1
        If oblig And rempli Then nbRempli = nbRempli + 1

        Debug.Print "  " & IIf(oblig, "[Ob]", "[  ]") & " " & _
                    IIf(oblig, IIf(rempli, "[V]", "[ ]"), "[.]") & " " & _
                    Left$(s & String(65, " "), 65) & " = '" & val & "'"
    Next cle

    Debug.Print String(70, "-")
    Debug.Print "nbObligatoires = " & nbObl & " | nbRemplis = " & nbRempli
    If nbObl > 0 Then
        Debug.Print "Taux MeQ = " & Format(nbRempli / nbObl, "0.00%")
    Else
        Debug.Print "Taux MeQ = 100% (aucune obligation)"
    End If
End Sub

' Dump les lignes brutes de Produit_SnapShotPrix filtrees sur un DTR.
' Lecture directe (bypass ADODB) pour eviter le bug ACE DTR mixed-type.
' Si aucune ligne ne matche, affiche 10 exemples de DTR presents en base.
Private Sub DumpSnapShotBrutPourDTR(ByVal dtrCible As String)
    Dim wbBdD As Workbook, wsBdD As Worksheet
    Dim wasOpen As Boolean
    Dim arr As Variant
    Dim derLigne As Long, ligne As Long
    Dim colDTR As Long, colType As Long, colPrix As Long, colLT As Long
    Dim colOnReqPrix As Long, colOnReqLT As Long, colMoQ As Long, colRefCat As Long, colDate As Long
    Dim maxCol As Long
    Dim sDTR As String
    Dim n As Long, nTotal As Long
    Dim sExemples As String

    On Error GoTo Sortie
    If Not OuvrirFeuilleSnapShotPrix(wbBdD, wsBdD, wasOpen) Then
        Debug.Print "  (impossible d'ouvrir Produit_SnapShotPrix)" : GoTo Sortie
    End If

    colDTR = TrouverColParLibelleSirus(wsBdD, "DTR")
    colType = TrouverColParLibelleSirus(wsBdD, "Type")
    colPrix = TrouverColParLibelleSirus(wsBdD, "Prix")
    colLT = TrouverColParLibelleSirus(wsBdD, "LeadTime")
    colOnReqPrix = TrouverColParLibelleSirus(wsBdD, "OnRequestPrix")
    colOnReqLT = TrouverColParLibelleSirus(wsBdD, "OnRequestLT")
    colMoQ = TrouverColParLibelleSirus(wsBdD, "MoQ")
    colRefCat = TrouverColParLibelleSirus(wsBdD, "Ref Catalogue")
    colDate = TrouverColParLibelleSirus(wsBdD, "Date Source")
    If colDTR = 0 Then Debug.Print "  (colonne DTR introuvable)" : GoTo Cleanup

    derLigne = wsBdD.Cells(wsBdD.rows.count, colDTR).End(xlUp).Row
    If derLigne < 2 Then Debug.Print "  (feuille vide)" : GoTo Cleanup

    maxCol = Application.WorksheetFunction.Max(colDTR, colType, colPrix, colLT, colOnReqPrix, colOnReqLT, colMoQ, colRefCat, colDate)
    arr = wsBdD.Range(wsBdD.Cells(2, 1), wsBdD.Cells(derLigne, maxCol)).Value

    For ligne = 1 To UBound(arr, 1)
        sDTR = Trim$(NzSafeStr(arr(ligne, colDTR)))
        If Right$(sDTR, 2) = ".0" Then sDTR = Left$(sDTR, Len(sDTR) - 2)
        nTotal = nTotal + 1
        If nTotal <= 10 Then sExemples = sExemples & "'" & sDTR & "' ; "

        If StrComp(sDTR, dtrCible, vbTextCompare) = 0 Then
            n = n + 1
            Debug.Print "  [" & NzSafeStr(arr(ligne, colType)) & _
                        "] Date=" & NzSafeStr(arr(ligne, colDate)) & _
                        " | Prix=" & NzSafeStr(arr(ligne, colPrix)) & _
                        IIf(colOnReqPrix > 0, " | OnReqPrix=" & NzSafeStr(arr(ligne, colOnReqPrix)), "") & _
                        " | LT=" & NzSafeStr(arr(ligne, colLT)) & _
                        IIf(colOnReqLT > 0, " | OnReqLT=" & NzSafeStr(arr(ligne, colOnReqLT)), "") & _
                        IIf(colMoQ > 0, " | MoQ=" & NzSafeStr(arr(ligne, colMoQ)), "") & _
                        " | RefCat='" & NzSafeStr(arr(ligne, colRefCat)) & "'"
        End If
    Next ligne

    If n = 0 Then
        Debug.Print "  (aucune ligne pour ce DTR)"
        Debug.Print "  Total lignes SnapShotPrix parcourues : " & nTotal
        Debug.Print "  DTR cible recherche : '" & dtrCible & "'"
        Debug.Print "  10 premiers DTR presents en base : " & sExemples
    End If

Cleanup:
    Call FermerBdDSiOuverte(wbBdD, wasOpen)
    Exit Sub

Sortie:
    Debug.Print "  DumpSnapShotBrutPourDTR erreur : " & Err.description
    Call FermerBdDSiOuverte(wbBdD, wasOpen)
End Sub


' =========================================================================
' Test_SIRUS_DiagPreFillProduit
' =========================================================================
' Diagnostic pour le pre-remplissage SnapShotPrix cote SIRUS_Produit_Modifier.
' Affiche dans la fenetre Immediate :
'   1. Etat du cache SnapShotPrix (populated / vide)
'   2. Contenu du config Produit : les 7 champs virtuels avec leur Source
'   3. Pour un DTR donne, ce qui serait ecrit dans la feuille par
'      PreRemplirDepuisBdD (branche Source = "SnapShotPrix")
'
' Usage : Test_SIRUS_DiagPreFillProduit "DTR0000120004"
Public Sub Test_SIRUS_DiagPreFillProduit(ByVal dtr As String)
    Debug.Print String(70, "=")
    Debug.Print "Diagnostic pre-fill SnapShotPrix pour DTR = '" & dtr & "'"
    Debug.Print String(70, "=")

    ' -- 1. Etat du cache ------------------------------------
    Debug.Print vbLf & "-- 1. Etat cache m_dictSnapShotPrixCache --"
    If m_dictSnapShotPrixCache Is Nothing Then
        Debug.Print "  Cache : VIDE (Nothing). Sera charge au 1er appel."
    Else
        Debug.Print "  Cache : POPULATED. " & m_dictSnapShotPrixCache.count & " DTR distincts en cache."
        If m_dictSnapShotPrixCache.count = 0 Then
            Debug.Print "  ATTENTION : cache populated MAIS vide (0 DTR) -> il faut ResetCache et retenter."
        End If
    End If

    ' -- 2. Config Produit : les 7 champs virtuels ------------
    Debug.Print vbLf & "-- 2. Config Produit : champs Source='SnapShotPrix' --"
    Dim cfg As Collection
    Set cfg = SIRUS_LireConfig("Produit")
    If cfg Is Nothing Or cfg.count = 0 Then
        Debug.Print "  ATTENTION : config Produit vide. Verifier feuille SIRUS_Config."
        Exit Sub
    End If
    Debug.Print "  Nb champs total dans config Produit : " & cfg.count

    Dim def As SIRUS_Champ_Config
    Dim nbSnap As Long
    For Each def In cfg
        If def.Source = "SnapShotPrix" Then
            nbSnap = nbSnap + 1
            Debug.Print "  [" & nbSnap & "] Champ='" & def.Champ & "' | ColonneBdD='" & def.ColonneBdD & "' | Obligatoire=" & def.Obligatoire
        End If
    Next def
    Debug.Print "  -> " & nbSnap & " champs Source='SnapShotPrix' trouves (attendu : 7)."
    If nbSnap = 0 Then
        Debug.Print "  ATTENTION : aucun champ Source='SnapShotPrix'. La feuille SIRUS_Config n'a pas ete regeneree apres bump version."
        Debug.Print "  Verifier SIRUS_CONFIG_VERSION et forcer un rebuild via SIRUS_InitConfigSheet."
    End If

    ' -- 3. Force le chargement + lookup pour le DTR ---------
    Debug.Print vbLf & "-- 3. Lookup dictSnapShotPrix pour DTR='" & dtr & "' --"
    Dim dictSnap As Object
    Set dictSnap = ChargerSnapShotPrixDernierParDTR()
    If dictSnap Is Nothing Then
        Debug.Print "  ATTENTION : ChargerSnapShotPrixDernierParDTR a renvoye Nothing (erreur ADODB ?)."
        Exit Sub
    End If
    Debug.Print "  Nb DTR total dans le dict : " & dictSnap.count
    If dictSnap.count = 0 Then
        Debug.Print "  ATTENTION : dict vide. TypeGuessRows=0 n'a peut-etre pas fonctionne."
        Debug.Print "  Verifier avec DumpSnapShotBrutPourDTR ou lancer Test_CompterPatternsDTR."
        Exit Sub
    End If

    If Not dictSnap.Exists(dtr) Then
        Debug.Print "  DTR '" & dtr & "' ABSENT du dict SnapShotPrix (pas de snapshot retenu)."
        Debug.Print "  Verifier : le DTR a-t-il des lignes Prix non nul ou OnRequestPrix dans Produit_SnapShotPrix ?"
        Dim k As Variant, iEx As Long
        Debug.Print "  10 premiers DTR presents en dict :"
        For Each k In dictSnap.Keys
            iEx = iEx + 1 : If iEx > 10 Then Exit For
            Debug.Print "    '" & k & "'"
        Next k
        Exit Sub
    End If

    Debug.Print "  DTR '" & dtr & "' PRESENT. Valeurs qui seraient ecrites en cellules :"
    Dim colKey As Variant
    For Each colKey In dictSnap(dtr).Keys
        Debug.Print "    " & CStr(colKey) & " = '" & NzSafeStr(dictSnap(dtr)(CStr(colKey))) & "'"
    Next colKey

    ' -- 4. Verification croisee : chaque champ config match dict ? --
    Debug.Print vbLf & "-- 4. Match config <-> dict pour ce DTR --"
    For Each def In cfg
        If def.Source = "SnapShotPrix" Then
            If dictSnap(dtr).Exists(def.ColonneBdD) Then
                Debug.Print "  [OK] '" & def.ColonneBdD & "' -> '" & NzSafeStr(dictSnap(dtr)(def.ColonneBdD)) & "'"
            Else
                Debug.Print "  [--] '" & def.ColonneBdD & "' -> ABSENT du dict (cellule restera vide)."
            End If
        End If
    Next def
End Sub


