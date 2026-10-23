Option Explicit

' Index des colonnes dans les Recordsets renvoyes par RecupererTachesDTR_BdD
' et RecupererInfoATM_DTR (B_ATM_Commun). L'ordre est celui du SELECT.
' GetRows() renvoie un tableau (champ, ligne) 0-based.
Private Const RC_TACHES_DTR             As Long = 0
Private Const RC_TACHES_LANGUE          As Long = 1
Private Const RC_TACHES_RANG            As Long = 2
Private Const RC_TACHES_HARD_SOFT       As Long = 3
Private Const RC_TACHES_CATEGORIE       As Long = 4
Private Const RC_TACHES_NIV_QUALIF      As Long = 5
Private Const RC_TACHES_DESCRIPTION     As Long = 6
Private Const RC_TACHES_CRITERES_DIM    As Long = 7
Private Const RC_TACHES_OUTILLAGE_SPE   As Long = 8
Private Const RC_TACHES_OUTILLAGE_STD   As Long = 9
Private Const RC_TACHES_PIECE_RECHANGE  As Long = 10
Private Const RC_TACHES_NB_OPERATEUR    As Long = 11
Private Const RC_TACHES_TPS_ESTIME      As Long = 12

Private Const RC_INFO_DTR               As Long = 0
Private Const RC_INFO_DESCRIPTIF_FR     As Long = 1
Private Const RC_INFO_LOCALISATION_FR   As Long = 2
Private Const RC_INFO_PLAN              As Long = 3
Private Const RC_INFO_COMMENTAIRE_FR    As Long = 4
Private Const RC_INFO_VERSION           As Long = 5
Private Const RC_INFO_DESCRIPTIF_EN     As Long = 6
Private Const RC_INFO_LOCALISATION_EN   As Long = 7
Private Const RC_INFO_COMMENTAIRE_EN    As Long = 8
Private Const RC_INFO_REPARABILITE      As Long = 9
Private Const RC_INFO_HARD_SOFT         As Long = 10
Private Const RC_INFO_POIDS             As Long = 11
Private Const RC_INFO_TPS_GEN           As Long = 12
Private Const RC_INFO_TPS_GEN_SOFT      As Long = 13
Private Const RC_INFO_TPS_ANA_CUMUL     As Long = 14
Private Const RC_INFO_NB_LANGUE         As Long = 15

' Mapping {indexDTR -> indexLigne dans Tableau1 du Sommaire pour la fiche HARD}
' Pre-calcule par PrecalculerLignesSommaire avant la boucle DTR principale.
' Si le DTR est BOTH, la ligne SOFT vaut mapLigneHARD(indexDTR) + 1
' (les deux fiches du DTR sont consecutives dans le sommaire).
Private mapLigneHARD() As Long
Private nbLignesSommaireTotal As Long

' Mapping {DTR -> type BdD} ou type BdD in {"HARD", "SOFT", "BOTH", ""}
' Pre-calcule a partir de Taches_tab (lui-meme charge par RecupererDonneesBDD_ATM).
' Remplace la lecture historique de tHARD_SOFT (colonne DATA_ASL) : on ne se
' base plus sur la PBS pour decider quel template generer, mais sur la
' presence reelle des taches HARD/SOFT en BdD pour le DTR.
'   ""    = DTR sans aucune tache en BdD (genere HARD vide "A faire")
'   HARD  = taches HARD uniquement -> 1 fiche HARD
'   SOFT  = taches SOFT uniquement -> 1 fiche SOFT
'   BOTH  = taches HARD ET SOFT    -> 2 fiches consecutives
Private dictTypeDTR As Object

Public Sub ATM_V1()
'Procedure de generation de fiches ATM V1

'Methode :
' 1. Selection d'un repertoire projet par l'utilisateur
' 2. Copie du template dans le repertoire choisie
' 3. Ouverture de ce template dans le repertoire
' 4. Recuperation des donnees Info et Taches ATM par la feuile complete
' 5. Copie feuille template de chaque DTR de la PBS : HARD / SOFT ou HARD & SOFT
' 6. Remplissage de chaque feuille avec donnees taches et infos

'Procedure de generation de fiches ATM V1.0 (Version Refactorisee)
'Genere UN SEUL fichier Excel multi-feuilles avec Index + N fiches DTR

    Dim chemin As String
    Dim Template As Workbook
    Dim i As Integer
    
    ' Initialisation
    userpath = Environ$("UserProfile")
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set REQSEL = CreateObject("ADODB.Recordset")
    
    ' Selection repertoire destination
    If Not SelectRepertoire(chemin) Then
        Exit Sub
    End If
    
    ' Copier et ouvrir template
    Set Template = CopierTemplateV1(chemin)
    If Template Is Nothing Then Exit Sub
    
    ' Recuperation donnees BDD (reutilise fonction ATM V2)
    Call RecupererDonneesBDD_ATM(1)

    ' Cadrage tableaux : initialise Taches_type / Info_type (numeros de colonnes)
    Call Cadrage_Taches_ATM(Taches_tab)
    Call Cadrage_Infos_ATM(Infos_tab)

    ' Indexation DTR -> ligne en O(1). Doit etre APRES les Cadrage_* sinon
    ' Taches_type.dtr / Info_type.dtr valent 0 et l'index est cassé.
    Call IndexerTachesEtInfos

    ' Recuperation PBS
    Set DS1 = Generateur.Worksheets("DATA_ASL")
    
    ' Identifier feuilles template
    Dim index As Worksheet, CoverPage As Worksheet
    Dim Template_HARD As Worksheet, Template_SOFT As Worksheet
    
    If langue = "FR" Then
        Set index = Template.Worksheets("Sommaire")
        Set CoverPage = Template.Worksheets("Couverture")
        Set Template_HARD = Template.Worksheets("ATM_HARD")
        Set Template_SOFT = Template.Worksheets("ATM_SOFT")
    Else
        Set index = Template.Worksheets("Index")
        Set CoverPage = Template.Worksheets("CoverPage")
        Set Template_HARD = Template.Worksheets("MTA_HARD")
        Set Template_SOFT = Template.Worksheets("MTA_SOFT")
    End If
    
    ' Pre-calcul des types BdD par DTR (HARD/SOFT/BOTH/"") en parcourant
    ' Taches_tab. Remplace la dependance a tHARD_SOFT (colonne DATA_ASL).
    Call PrecalculerTypesDTR_BdD

    ' Pre-calcul du mapping {indexDTR -> ligne HARD dans Tableau1} pour que
    ' les DTR BOTH aient leurs deux fiches consecutives dans le sommaire.
    Call PrecalculerLignesSommaire

    ' Remplir Index
    Call RemplirIndexV1(index, Template)

    ' UX : initialiser la barre de progression + StatusBar + curseur sablier.
    Call Progress_Demarrer(UBound(tDTR), "ATM V1")

    ' Boucle sur chaque DTR
    For i = 1 To UBound(tDTR)

        ' UX : mise a jour barre + StatusBar avant le traitement de ce DTR.
        Call Progress_Etape(i, "", CStr(tDTR(i, 1)))

        ' Skip supports/frames
        If InStr(1, tSS(i, 1), "support") > 0 Or InStr(1, tSS(i, 1), "frame") > 0 Then
            ' Statut "NA" en col F sur la ligne HARD du DTR (mapping pre-calcule).
            index.Range("F" & (mapLigneHARD(i) + 5)) = "NA"
            GoTo NextDTR
        End If

        ' Generer fiche(s) ATM pour ce DTR
        Call GenererFicheATM_V1(Template, index, Template_HARD, Template_SOFT, i)

NextDTR:
    Next i

    ' UX : etape de finalisation (regroupements, mise en page) visible
    ' dans la StatusBar avant que la barre disparaisse.
    Call Progress_Etape(UBound(tDTR), "Finalisation...", "")

    ' Finalisation
    Call FinaliserFichierV1(Template, index, CoverPage, Template_HARD, Template_SOFT)

    ' UX : restaurer curseur normal, StatusBar, masquer la barre.
    Call Progress_Terminer

    ' Message final
    If langue = "FR" Then
        MsgBox "Fiches ATM creees. Veuillez fermer cette fenetre pour consulter le fichier genere.", vbInformation
    Else
        MsgBox "MTA created. Please close this window to consult the new file generated.", vbInformation
    End If

End Sub
' =========================================================================
' FONCTION: COPIER TEMPLATE V1
' =========================================================================
Private Function CopierTemplateV1(ByVal chemin As String) As Workbook
    Dim nomFichier As String
    Dim templatePath As String
    
    ' Construire nom fichier avec date
    If langue = "FR" Then
        nomFichier = Month(Date) & "_" & Year(Date) & "_" & Projet_Abrev & "_ATM_" & Revision & ".xlsx"
        templatePath = userpath & TEMPLATE_V1_FR_PATH
    Else
        nomFichier = Year(Date) & "_" & Month(Date) & "_" & Projet_Abrev & "_MTA_" & Revision & ".xlsx"
        templatePath = userpath & TEMPLATE_V1_EN_PATH
    End If
    
    ' Copier template
    On Error GoTo ErrHandler
    Application.DisplayAlerts = False
    fso.CopyFile templatePath, chemin & "\" & nomFichier
    Application.DisplayAlerts = True
    
    ' Ouvrir fichier
    Set CopierTemplateV1 = Workbooks.Open(chemin & "\" & nomFichier)
    Exit Function
    
ErrHandler:
    MsgBox "Erreur lors de la copie du template: " & Err.description, vbCritical
    Set CopierTemplateV1 = Nothing
End Function
' =========================================================================
' PROCEDURE: REMPLIR INDEX V1
' =========================================================================
Private Sub RemplirIndexV1(ByRef index As Worksheet, ByRef Template As Workbook)
    Dim TabIndex As ListObject
    Dim i As Integer
    Dim ligneHARD As Long, ligneSOFT As Long
    Dim designation As String
    Dim colDes As String, colSS As String, colTitre As String

    Set TabIndex = index.Range("Tableau1").ListObject

    ' Resize a la taille totale pre-calculee (HARD + SOFT consecutifs pour BOTH).
    ' nbLignesSommaireTotal est alimente par PrecalculerLignesSommaire (appele
    ' avant cette routine dans ATM_V1).
    TabIndex.Resize Range("$B$5:$F$" & 5 + nbLignesSommaireTotal)

    ' Constantes selon langue
    If langue = "FR" Then
        colDes = "Designation"
        colSS = "Sous syst" & ChrW(232) & "me"
        colTitre = "ATM Projet "
    Else
        colDes = "Designation"
        colSS = "Sub system"
        colTitre = "MTA Project "
    End If

    ' Boucle DTR : pose une ligne HARD pour chaque DTR + une ligne SOFT
    ' supplementaire et consecutive si BOTH.
    For i = 1 To UBound(tDTR)
        ligneHARD = mapLigneHARD(i)

        ' Designation : MEF prioritaire, sinon PBS
        If tDES_MEF(i, 1) = "NA" Then
            designation = CStr(tDES_PBS(i, 1))
        Else
            designation = CStr(tDES_MEF(i, 1))
        End If

        ' Ligne HARD du DTR (toujours posee, meme si le DTR sera SOFT seul :
        ' le lien hypertexte ira vers la feuille SOFT - cf. GenererFicheATM_V1).
        TabIndex.ListColumns("DTR").DataBodyRange(ligneHARD).Value = tDTR(i, 1)
        TabIndex.ListColumns(colDes).DataBodyRange(ligneHARD).Value = designation
        TabIndex.ListColumns(colSS).DataBodyRange(ligneHARD).Value = tSS(i, 1)

        ' Ligne SOFT consecutive si DTR = BOTH (selon les taches en BdD,
        ' pas selon DATA_ASL).
        If GetTypeDTR(CStr(tDTR(i, 1))) = "BOTH" Then
            ligneSOFT = ligneHARD + 1
            TabIndex.ListColumns("DTR").DataBodyRange(ligneSOFT).Value = tDTR(i, 1)
            ' Pas de suffixe "(SOFT)" : la designation est la meme que HARD.
            ' L'utilisateur distingue via le lien hypertexte (nom de la feuille).
            TabIndex.ListColumns(colDes).DataBodyRange(ligneSOFT).Value = designation
            TabIndex.ListColumns(colSS).DataBodyRange(ligneSOFT).Value = tSS(i, 1)
        End If
    Next i

    ' Titre
    index.Cells(3, 2) = colTitre & Projet_Abrev & " Rev " & Revision
End Sub

' =========================================================================
' PROCEDURE: PRECALCULER LIGNES SOMMAIRE
' =========================================================================
' Alimente le tableau module mapLigneHARD : pour chaque DTR de la PBS,
' renvoie l'index de sa ligne HARD dans Tableau1 du Sommaire. Les DTR BOTH
' occupent 2 lignes consecutives (HARD puis SOFT) pour les avoir cote a
' cote dans le sommaire au lieu d'avoir toutes les SOFT en fin de tableau.
'
' Aussi : nbLignesSommaireTotal = somme des lignes (1 par DTR + 1 pour BOTH).
Private Sub PrecalculerLignesSommaire()
    Dim i As Long
    Dim ligneCourante As Long
    Dim t As String

    ReDim mapLigneHARD(1 To UBound(tDTR))
    ligneCourante = 1
    For i = 1 To UBound(tDTR)
        mapLigneHARD(i) = ligneCourante
        t = GetTypeDTR(CStr(tDTR(i, 1)))
        If t = "BOTH" Then
            ligneCourante = ligneCourante + 2  ' HARD + SOFT consecutives
        Else
            ligneCourante = ligneCourante + 1  ' HARD seul, SOFT seul, ou "" (HARD vide "A faire")
        End If
    Next i
    nbLignesSommaireTotal = ligneCourante - 1
End Sub

' =========================================================================
' PROCEDURE: PRECALCULER TYPES DTR BDD
' =========================================================================
' Parcourt Taches_tab (charge par RecupererDonneesBDD_ATM) et construit le
' dictionnaire dictTypeDTR = {DTR -> "HARD" / "SOFT" / "BOTH" / ""}.
'
' Algorithme :
'   - 1re occurrence d'un DTR : type = HARD ou SOFT (selon la tache trouvee)
'   - Occurrence suivante avec un type different : type = "BOTH"
'   - DTR absent du dictionnaire = pas de tache en BdD (-> "" via GetTypeDTR)
Private Sub PrecalculerTypesDTR_BdD()
    Dim j As Long
    Dim dtr As String
    Dim hs As String

    Set dictTypeDTR = CreateObject("Scripting.Dictionary")
    dictTypeDTR.CompareMode = 1  ' BinaryCompare (DTR sensible casse)

    On Error GoTo Sortie
    For j = LBound(Taches_tab, 1) To UBound(Taches_tab, 1)
        dtr = CStr(Taches_tab(j, Taches_type.dtr))
        hs = UCase$(Trim$(CStr(Taches_tab(j, Taches_type.HARD_SOFT))))
        If LenB(dtr) > 0 And (hs = "HARD" Or hs = "SOFT") Then
            If dictTypeDTR.Exists(dtr) Then
                If dictTypeDTR(dtr) <> hs Then dictTypeDTR(dtr) = "BOTH"
            Else
                dictTypeDTR(dtr) = hs
            End If
        End If
    Next j
    Exit Sub
Sortie:
    Debug.Print "PrecalculerTypesDTR_BdD : erreur " & Err.description
End Sub

' =========================================================================
' FONCTION: GET TYPE DTR (depuis dictTypeDTR)
' =========================================================================
' Renvoie "HARD", "SOFT", "BOTH" ou "" selon ce qui a ete pre-calcule.
' "" = DTR sans tache en BdD pour cette langue ==> fiche HARD vide + "A faire".
Private Function GetTypeDTR(ByVal dtr As String) As String
    If dictTypeDTR Is Nothing Then
        GetTypeDTR = ""
    ElseIf dictTypeDTR.Exists(dtr) Then
        GetTypeDTR = CStr(dictTypeDTR(dtr))
    Else
        GetTypeDTR = ""
    End If
End Function
' =========================================================================
' PROCEDURE: GENERER FICHE ATM V1 POUR UN DTR
' =========================================================================
Private Sub GenererFicheATM_V1(ByRef Template As Workbook, ByRef index As Worksheet, _
                               ByRef Template_HARD As Worksheet, ByRef Template_SOFT As Worksheet, _
                               ByVal indexDTR As Integer)
    Dim dEst As Worksheet
    Dim system As String
    Dim cpt As Integer
    Dim rsTaches As Object, rsInfo As Object
    Dim arrTaches As Variant, arrInfo As Variant
    Dim hasTaches As Boolean, hasInfo As Boolean
    Dim dtrCourant As String

    system = tSS(indexDTR, 1)
    cpt = 0
    dtrCourant = CStr(tDTR(indexDTR, 1))

    ' --- Recuperation des taches + infos du DTR (1 SELECT chacun) ----------
    ' Cas "DTR absent en BdD" : Recordset.EOF = True, on marque "A faire".
    Set rsTaches = RecupererTachesDTR_BdD(dtrCourant, langue, 1)
    hasTaches = Not (rsTaches Is Nothing)
    If hasTaches Then hasTaches = Not rsTaches.EOF
    If hasTaches Then
        arrTaches = rsTaches.GetRows()       ' (champ, ligne) 0-based
        rsTaches.Close
    End If
    Set rsTaches = Nothing

    Set rsInfo = RecupererInfoATM_DTR(dtrCourant)
    hasInfo = Not (rsInfo Is Nothing)
    If hasInfo Then hasInfo = Not rsInfo.EOF
    If hasInfo Then
        arrInfo = rsInfo.GetRows()
        rsInfo.Close
    End If
    Set rsInfo = Nothing

    ' --- Type DTR depuis BdD (HARD/SOFT/BOTH/"") -----------
    ' Decide quelles fiches generer en fonction des taches reellement
    ' presentes en base, plus selon la colonne HARD/SOFT de DATA_ASL.
    Dim typeDTR As String
    typeDTR = GetTypeDTR(dtrCourant)

    ' --- Generer la feuille SOFT si DTR a des taches SOFT (SOFT ou BOTH) ----
    Dim destSOFT As Worksheet
    Set destSOFT = Nothing
    If typeDTR = "SOFT" Or typeDTR = "BOTH" Then
        Template_SOFT.Copy After:=Template.Worksheets(Template.Worksheets.count)
        Application.CutCopyMode = False
        Set destSOFT = ActiveSheet
        ' Compter doublons en se basant sur le nom complet (baseName + _SOFT)
        ' pour ne pas confondre avec la fiche HARD du meme DTR.
        cpt = CompterDoublonsFeuilles(Template, index, indexDTR, system, "SOFT")
        Call NommerFeuilleV1(destSOFT, index, indexDTR, system, cpt, "SOFT")
        If hasTaches Then
            Call RemplirTachesV1(destSOFT, arrTaches, "SOFT")
            Call PoserNiveauQualif(destSOFT, arrTaches, "SOFT")
            Call RemplirInfosV1(destSOFT, indexDTR, arrInfo, hasInfo, "SOFT")
        End If
        ' Description sur une seule ligne en SOFT (pas de wrap auto).
        destSOFT.Columns("C").WrapText = False
    End If

    ' --- Generer la feuille HARD si DTR a des taches HARD ou si type "" -----
    ' Cas "" : aucune tache en BdD -> on genere quand meme une fiche HARD
    ' vide et on marque "A faire" dans le sommaire (continuite metier).
    Set dEst = Nothing
    If typeDTR = "" Or typeDTR = "HARD" Or typeDTR = "BOTH" Then
        Template_HARD.Copy After:=Template.Worksheets(Template.Worksheets.count)
        Application.CutCopyMode = False
        Set dEst = ActiveSheet
        cpt = CompterDoublonsFeuilles(Template, index, indexDTR, system, "HARD")
        Call NommerFeuilleV1(dEst, index, indexDTR, system, cpt, "HARD")
        If hasTaches Then
            Call RemplirTachesV1(dEst, arrTaches, "HARD")
            Call PoserNiveauQualif(dEst, arrTaches, "HARD")
            Call RemplirInfosV1(dEst, indexDTR, arrInfo, hasInfo, "HARD")
        End If
    End If

    ' Pour le lien hypertexte : on prefere HARD si dispo, sinon SOFT.
    If dEst Is Nothing Then Set dEst = destSOFT

    ' Si DTR pas dans la BdD pour cette langue : marquer "A faire"
    If Not hasTaches Then Call MarquerStatutIndex(index, indexDTR, "A faire")

AjouterLien:
    ' Creer lien hypertexte (vers la 1ere feuille creee si dispo)
    If Not dEst Is Nothing Then Call CreerLienHypertexte(index, indexDTR, dEst)

    ' --- Cas BOTH : poser le lien hypertexte sur la ligne SOFT pre-allouee
    '    par RemplirIndexV1 (les 2 lignes du DTR sont deja consecutives :
    '    HARD = mapLigneHARD(indexDTR), SOFT = mapLigneHARD(indexDTR) + 1).
    If typeDTR = "BOTH" And Not destSOFT Is Nothing Then
        Dim hasTachesSOFT As Boolean
        hasTachesSOFT = False
        If hasTaches Then hasTachesSOFT = ATachesPourType(arrTaches, "SOFT")
        Call PoserLienSommaireSOFT(index, indexDTR, destSOFT, hasTachesSOFT)
    End If
End Sub

' =========================================================================
' FONCTION: A TACHES POUR TYPE
' =========================================================================
' Renvoie True si le tableau de taches contient au moins une tache du
' type HARD ou SOFT specifie. Utile pour determiner si la feuille SOFT
' du DTR doit etre marquee "A faire" (taches SOFT absentes en BdD).
Private Function ATachesPourType(ByRef arrTaches As Variant, ByVal typeHS As String) As Boolean
    Dim j As Long
    On Error GoTo Sortie
    For j = LBound(arrTaches, 2) To UBound(arrTaches, 2)
        If CStr(arrTaches(RC_TACHES_HARD_SOFT, j)) = typeHS Then
            ATachesPourType = True
            Exit Function
        End If
    Next j
Sortie:
    ATachesPourType = False
End Function

' =========================================================================
' PROCEDURE: POSER LIEN SOMMAIRE SOFT (cas BOTH)
' =========================================================================
' Pour un DTR BOTH, la ligne SOFT a deja ete posee par RemplirIndexV1 a
' la position consecutive mapLigneHARD(indexDTR) + 1 (DTR/Designation/Sous
' systeme). On a juste a poser le lien hypertexte vers la feuille SOFT et
' le statut "A faire" si pas de taches SOFT en BdD.
Private Sub PoserLienSommaireSOFT(ByRef index As Worksheet, _
                                    ByVal indexDTR As Integer, _
                                    ByRef destSOFT As Worksheet, _
                                    ByVal hasTachesSOFT As Boolean)
    Dim TabIndex As ListObject
    Dim colLien As String
    Dim ligneSOFT As Long

    Set TabIndex = index.Range("Tableau1").ListObject
    ligneSOFT = mapLigneHARD(indexDTR) + 1

    If langue = "FR" Then
        colLien = "Lien Hypertexte"
    Else
        colLien = "Hyperlink"
    End If

    ' Lien hypertexte vers la feuille SOFT
    index.Hyperlinks.Add _
        Anchor:=TabIndex.ListColumns(colLien).DataBodyRange(ligneSOFT), _
        Address:="", _
        SubAddress:="'" & destSOFT.Name & "'!A1", _
        TextToDisplay:=destSOFT.Name

    ' Statut "A faire" si pas de taches SOFT en BdD pour ce DTR (col 5)
    If Not hasTachesSOFT Then
        If langue = "FR" Then
            TabIndex.DataBodyRange(ligneSOFT, 5).Value = "A faire"
        Else
            TabIndex.DataBodyRange(ligneSOFT, 5).Value = "To do"
        End If
    End If
End Sub

' =========================================================================
' PoserNiveauQualif : ecrit le niveau qualif de la 1ere tache du bon type
' dans la cellule I7 de la feuille. Tolere l'absence de tache (no-op).
' =========================================================================
Private Sub PoserNiveauQualif(ByRef dEst As Worksheet, ByRef arrTaches As Variant, ByVal typeHS As String)
    Dim j As Long
    For j = LBound(arrTaches, 2) To UBound(arrTaches, 2)
        If CStr(arrTaches(RC_TACHES_HARD_SOFT, j)) = typeHS Then
            dEst.Range("I7") = arrTaches(RC_TACHES_NIV_QUALIF, j)
            Exit Sub
        End If
    Next j
End Sub

' =========================================================================
' FONCTION: COMPTER DOUBLONS FEUILLES
' =========================================================================
Private Function CompterDoublonsFeuilles(ByRef Template As Workbook, ByRef index As Worksheet, _
                                         ByVal indexDTR As Integer, ByVal system As String, _
                                         ByVal typeHS As String) As Integer
    ' Compte le nombre de feuilles existantes qui matchent EXACTEMENT
    ' baseName + suffixe (_HARD ou _SOFT). Distingue donc les fiches HARD
    ' et SOFT du meme DTR (fix bug DTR_xxx (2) sans suffixe).
    Dim ws As Worksheet
    Dim cpt As Integer
    Dim baseName As String
    Dim suffixe As String
    Dim nomComplet As String
    Dim nomBase As String

    baseName = index.Range("B" & (mapLigneHARD(indexDTR) + 5)) & "_" & Trim(Left(system, 4))
    Select Case typeHS
        Case "SOFT": suffixe = "_SOFT"
        Case Else:   suffixe = "_HARD"
    End Select
    nomComplet = baseName & suffixe

    cpt = 0
    For Each ws In Template.Worksheets
        ' Match exact OU match prefixe + parenthese de numerotation
        ' Ex : DTR_TCMS_HARD matche DTR_TCMS_HARD et DTR_TCMS_HARD (1)
        nomBase = Left$(ws.Name, Len(nomComplet))
        If nomBase = nomComplet Then cpt = cpt + 1
    Next ws

    CompterDoublonsFeuilles = cpt
End Function

' =========================================================================
' PROCEDURE: NOMMER FEUILLE V1
' =========================================================================
Private Sub NommerFeuilleV1(ByRef dEst As Worksheet, ByRef index As Worksheet, _
                            ByVal indexDTR As Integer, ByVal system As String, _
                            ByVal cpt As Integer, ByVal typeHS As String)
    Dim baseName As String
    Dim suffixe As String
    Dim nomComplet As String

    ' baseName = DTR_<premiers 4 chars du sous systeme>
    baseName = index.Range("B" & (mapLigneHARD(indexDTR) + 5)) & "_" & Trim(Left(system, 4))

    ' Suffixe selon le type de feuille demande.
    Select Case typeHS
        Case "SOFT": suffixe = "_SOFT"
        Case Else:   suffixe = "_HARD"     ' HARD, BOTH, "" -> HARD par defaut
    End Select

    nomComplet = baseName & suffixe

    If cpt = 0 Then
        ' Pas de doublon en PBS pour ce type : nom sans numerotation.
        dEst.Name = nomComplet
    Else
        ' Doublon en PBS DATA_ASL pour ce type : ajouter " (cpt)".
        ' Le suffixe _HARD/_SOFT est preserve (fix bug 2026-06-10).
        ' Ex : 1er DTR doublon -> "DTR_TCMS_HARD (1)"
        '      2e               -> "DTR_TCMS_HARD (2)" etc.
        dEst.Name = nomComplet & " (" & cpt & ")"
    End If
End Sub
' =========================================================================
' PROCEDURE: REMPLIR TACHES V1 (depuis tableau local rs.GetRows())
' =========================================================================
' arrTaches est un tableau 2D (champ, ligne) 0-based renvoye par
' RecupererTachesDTR_BdD(...).GetRows(). Pre-filtre SQL : WHERE DTR + Langue,
' donc on ne filtre plus que sur HARD/SOFT cote VBA.
' typeHARDSOFT : "HARD" ou "SOFT" -> filtre exact. Tout autre valeur ->
' on filtre par "HARD" par defaut (comportement legacy).
Private Sub RemplirTachesV1(ByRef dEst As Worksheet, ByRef arrTaches As Variant, _
                            ByVal typeHARDSOFT As String)
    Dim j As Long
    Dim nvx As Integer
    Dim rang As String
    Dim rangChiffre As String
    Dim filtre As String
    Dim rangAffiche As String
    Dim maxRang3 As Integer

    ' Filtre effectif : "SOFT" pour la feuille SOFT, "HARD" sinon.
    If typeHARDSOFT = "SOFT" Then filtre = "SOFT" Else filtre = "HARD"

    ' --- Mode SOFT : template n'a que 4 marqueurs (chapitres 1..4) ---
    ' Mapping rang BdD -> chapitre + renumeration (validation Thomas 2026-06-10) :
    '   1-x BdD -> chapitre 1 SOFT (rang inchange)
    '   2-x BdD -> chapitre 2 SOFT (rang inchange)
    '   3-x BdD -> chapitre 3 SOFT (rang inchange)
    '   4-x BdD -> chapitre 3 SOFT, renumerote en "3-(maxRang3 + x)"
    '   5-x BdD -> chapitre 4 SOFT, renumerote en "4-x"
    '
    ' Pre-passe pour calculer maxRang3 = plus grand x dans les rangs 3-x
    ' des taches SOFT du DTR. Necessaire avant la passe d'insertion pour
    ' connaitre l'offset a appliquer aux rangs 4-x.
    maxRang3 = 0
    If typeHARDSOFT = "SOFT" Then
        For j = LBound(arrTaches, 2) To UBound(arrTaches, 2)
            If CStr(arrTaches(RC_TACHES_HARD_SOFT, j)) = "SOFT" Then
                rang = CStr(arrTaches(RC_TACHES_RANG, j))
                If Len(rang) >= 3 And Left(rang, 2) = "3-" Then
                    On Error Resume Next
                    nvx = CInt(Mid$(rang, 3))
                    If Err.Number = 0 And nvx > maxRang3 Then maxRang3 = nvx
                    Err.Clear
                    On Error GoTo 0
                End If
            End If
        Next j
    End If

    ' Parcourir taches de la derniere a la premiere (insertion avant)
    For j = UBound(arrTaches, 2) To LBound(arrTaches, 2) Step -1

        ' Filtre HARD/SOFT selon le type de feuille demande
        If CStr(arrTaches(RC_TACHES_HARD_SOFT, j)) <> filtre Then GoTo NextTache

        rang = CStr(arrTaches(RC_TACHES_RANG, j))

        ' Verifier format rang (skip si >3 caracteres)
        If Len(rang) > 3 Then GoTo NextTache

        ' Extraire niveau (ex: "5-2" -> 2)
        nvx = CInt(Right(rang, Len(rang) - InStr(1, rang, "-")))
        rangChiffre = Left(rang, 1)
        rangAffiche = rang   ' par defaut : rang BdD tel quel

        ' --- Mode SOFT : remapping chapitre + renumeration -----
        If typeHARDSOFT = "SOFT" Then
            Select Case rangChiffre
                Case "4"
                    ' Tache rang 4-x BdD -> chapitre 3 SOFT, renumerotee
                    rangChiffre = "3"
                    nvx = maxRang3 + nvx
                    rangAffiche = "3-" & nvx
                Case "5"
                    ' Tache rang 5-x BdD -> chapitre 4 SOFT (template SOFT
                    ' n'a pas de 5e chapitre), renumerotee en "4-x".
                    rangChiffre = "4"
                    rangAffiche = "4-" & nvx
            End Select
        End If

        ' Inserer tache selon rang
        Call InsererTacheV1(dEst, arrTaches, j, rangChiffre, nvx, rangAffiche)

NextTache:
    Next j
End Sub

' =========================================================================
' PROCEDURE: INSERER TACHE V1 (depuis tableau local)
' =========================================================================
Private Sub InsererTacheV1(ByRef dEst As Worksheet, ByRef arrTaches As Variant, _
                           ByVal indexTache As Long, _
                           ByVal rangChiffre As String, ByVal niveau As Integer, _
                           Optional ByVal rangAffiche As String = "")
    Dim ligneBase As Integer
    Dim ligneInsert As Integer
    Dim temps As Double

    ' Determiner ligne de base selon rang
    Select Case rangChiffre
        Case "1": ligneBase = LIGNE_RANG_1
        Case "2": ligneBase = LIGNE_RANG_2
        Case "3": ligneBase = LIGNE_RANG_3
        Case "4": ligneBase = LIGNE_RANG_4
        Case "5": ligneBase = LIGNE_RANG_5
        Case Else: Exit Sub
    End Select

    If niveau > 1 Then
        ' Inserer nouvelle ligne en heritant directement du format de la
        ' ligne au-dessus (Excel native, pas de Clipboard). 1 op au lieu
        ' de 3 (Insert + Range.Copy + CutCopyMode) -> beaucoup plus rapide
        ' et plus d'erreur Presse-Papier en cascade.
        ligneInsert = ligneBase + 1
        dEst.Range("B" & ligneInsert).EntireRow.Insert _
            Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Else
        ' Premiere tache du rang (x-1)
        ligneInsert = ligneBase
    End If

    ' Forcer format texte sur col B (rang) pour eviter conversion "3-6"
    ' -> date par Excel. NumberFormat=@ sur la cellule cible avant ecriture.
    dEst.Range("B" & ligneInsert).NumberFormat = "@"

    ' Remplir donnees depuis le tableau local
    With dEst
        ' Rang : si rangAffiche fourni (mode SOFT renumere), il prime.
        ' Sinon valeur brute BdD.
        If LenB(rangAffiche) > 0 Then
            .Range("B" & ligneInsert) = rangAffiche
        Else
            .Range("B" & ligneInsert) = arrTaches(RC_TACHES_RANG, indexTache)
        End If

        .Range("C" & ligneInsert) = arrTaches(RC_TACHES_DESCRIPTION, indexTache)
        .Range("E" & ligneInsert) = arrTaches(RC_TACHES_CRITERES_DIM, indexTache)
        .Range("F" & ligneInsert) = arrTaches(RC_TACHES_OUTILLAGE_SPE, indexTache)
        .Range("G" & ligneInsert) = arrTaches(RC_TACHES_OUTILLAGE_STD, indexTache)
        .Range("H" & ligneInsert) = arrTaches(RC_TACHES_PIECE_RECHANGE, indexTache)
        .Range("I" & ligneInsert) = arrTaches(RC_TACHES_NB_OPERATEUR, indexTache)

        ' Temps avec gestion NULL
        temps = GetTempsFromBDD(arrTaches(RC_TACHES_TPS_ESTIME, indexTache))
        If temps > 0 Then
            .Range("J" & ligneInsert) = temps
            .Range("J" & ligneInsert).NumberFormat = "0.000"
        Else
            .Range("J" & ligneInsert) = ""
        End If
    End With
End Sub

' =========================================================================
' PROCEDURE: REMPLIR INFOS V1
' =========================================================================
Private Sub RemplirInfosV1(ByRef dEst As Worksheet, ByVal indexDTR As Integer, _
                            ByRef arrInfo As Variant, ByVal hasInfo As Boolean, _
                            ByVal typeHS As String)
    Dim j As Long
    Dim filtre As String

    ' Donnees toujours presentes (depuis la PBS)
    dEst.Range("D4") = tDES_MEF(indexDTR, 1)
    dEst.Range("D6") = tDTR(indexDTR, 1)
    dEst.Range("D7") = Code_Projet
    dEst.Range("D8") = tSS(indexDTR, 1)

    ' Remplir donnees RAM selon gate (avant/apres CGR) + type fiche
    ' (SOFT : pas de MTBF, juste la DVL SOFT en I5).
    Call RemplirDonneesRAM_V1(dEst, indexDTR, typeHS)

    ' Si pas d'Info ATM en BdD, on s'arrete la (le DTR existe en PBS mais
    ' pas encore d'Info ATM en base - pas critique).
    If Not hasInfo Then Exit Sub

    ' Filtre Info ATM selon le type de feuille
    If typeHS = "SOFT" Then filtre = "SOFT" Else filtre = "HARD"

    For j = LBound(arrInfo, 2) To UBound(arrInfo, 2)
        If CStr(arrInfo(RC_INFO_HARD_SOFT, j)) = filtre Then

            dEst.Range("D9") = arrInfo(RC_INFO_VERSION, j)
            dEst.Range("I6") = arrInfo(RC_INFO_POIDS, j)

            If langue = "EN" Then
                dEst.Range("D5") = arrInfo(RC_INFO_DESCRIPTIF_EN, j)
                If CStr(arrInfo(RC_INFO_REPARABILITE, j)) = "REPARABLE" Then
                    dEst.Range("I9") = "REPAIRABLE"
                Else
                    dEst.Range("I9") = "SPARE"
                End If
            Else
                dEst.Range("D5") = arrInfo(RC_INFO_DESCRIPTIF_FR, j)
                dEst.Range("I8") = arrInfo(RC_INFO_PLAN, j)
                dEst.Range("I9") = arrInfo(RC_INFO_REPARABILITE, j)
            End If

            Exit For
        End If
    Next j
End Sub

' =========================================================================
' PROCEDURE: REMPLIR DONNEES RAM V1
' =========================================================================
' Pose le MTBF et la DVL (Duree de Vie Limitee = periodicite) en cellules
' adjacentes aux libelles "MTBF" / "Periodicite" du template (avec fallback
' positions I4 / I5 pour compat ancien template).
'
' Regles metier :
'   - typeHS = "SOFT" : PAS de MTBF pose (toujours laisse vide en SOFT) ;
'                       DVL = DureeVie_SOFT (col U avant CGR / col W apres CGR).
'   - typeHS = "HARD" : MTBF + DVL HARD (col P avant CGR / col R apres CGR).
'   - Detection libelle scalable : libelle "MTBF" -> I4, "Periodicite" -> I5
'     dans la zone d'entete (A1:K12), avec fallback positions historiques.
Private Sub RemplirDonneesRAM_V1(ByRef dEst As Worksheet, _
                                  ByVal indexDTR As Integer, _
                                  ByVal typeHS As String)
    Dim valMTBF As Variant
    Dim valDVL As Variant

    ' --- MTBF : seulement pour HARD (logique metier, SOFT n'en a pas) ---
    If typeHS <> "SOFT" Then
        valMTBF = CalculerMTBF(indexDTR)
        If Not IsEmpty(valMTBF) Then
            Call PoserValeurApresLibelle(dEst, _
                Array("MTBF"), valMTBF, "I4")
        End If
    End If

    ' --- DVL (Duree de vie limitee = periodicite) : pour HARD ET SOFT ---
    valDVL = CalculerDVL(indexDTR, typeHS)
    If Not IsEmpty(valDVL) Then
        Call PoserValeurApresLibelle(dEst, _
            Array("P" & ChrW(233) & "riodicit", "Periodicit", "Periodicity"), _
            valDVL, "I5")
    End If
End Sub

' =========================================================================
' FONCTION: CALCULER MTBF
' =========================================================================
' Renvoie la valeur MTBF a poser selon le mode :
'   - Avant CGR    : generique (tMTBF_gen)
'   - Apres CGR    : fonctionnel si complet (tMTBF_fct), sinon hierarchie
'                    fonctionnel -> intrinseque -> generique.
Private Function CalculerMTBF(ByVal indexDTR As Integer) As Variant
    If AvantCGR Then
        CalculerMTBF = tMTBF_gen(indexDTR, 1)
    Else
        If RAMReportDATA Then
            CalculerMTBF = tMTBF_fct(indexDTR, 1)
        Else
            If tMTBF_fct(indexDTR, 1) <> "" And IsNumeric(tMTBF_fct(indexDTR, 1)) Then
                CalculerMTBF = tMTBF_fct(indexDTR, 1)
            ElseIf tMTBF_Intr(indexDTR, 1) <> "" And IsNumeric(tMTBF_Intr(indexDTR, 1)) Then
                CalculerMTBF = tMTBF_Intr(indexDTR, 1)
            ElseIf tMTBF_gen(indexDTR, 1) <> "" And IsNumeric(tMTBF_gen(indexDTR, 1)) Then
                CalculerMTBF = tMTBF_gen(indexDTR, 1)
            Else
                CalculerMTBF = Empty
            End If
        End If
    End If
End Function

' =========================================================================
' FONCTION: CALCULER DVL (DUREE DE VIE LIMITEE / PERIODICITE)
' =========================================================================
' Renvoie la DVL a poser selon le type (HARD/SOFT) et le mode (avant/apres
' CGR + RAM Report).
'
' DATA_ASL : selon le cadrage existant, les colonnes utilisees sont :
'   HARD avant CGR : tDureevie_Gen_HARD (col P)
'   HARD apres CGR : tDureevie_Ana_HARD (col R)
'   SOFT avant CGR : tDureevie_Gen_SOFT (col U)
'   SOFT apres CGR : tDureevie_Ana_SOFT (col W)
' Le cadrage est centralise dans RecupererDonneesPBS (B_ATM_Main), on lit
' juste les tableaux globaux ici.
Private Function CalculerDVL(ByVal indexDTR As Integer, ByVal typeHS As String) As Variant
    Dim estSOFT As Boolean
    estSOFT = (typeHS = "SOFT")

    If AvantCGR Then
        ' Avant CGR : generique
        If estSOFT Then
            CalculerDVL = tDureevie_Gen_SOFT(indexDTR, 1)
        Else
            CalculerDVL = tDureevie_Gen_HARD(indexDTR, 1)
        End If
    Else
        ' Apres CGR : analytique en priorite, fallback generique
        If RAMReportDATA Then
            If estSOFT Then
                CalculerDVL = tDureevie_Ana_SOFT(indexDTR, 1)
            Else
                CalculerDVL = tDureevie_Ana_HARD(indexDTR, 1)
            End If
        Else
            ' RAM Report partiel : analytique si rempli, sinon generique
            If estSOFT Then
                If tDureevie_Ana_SOFT(indexDTR, 1) <> "" And IsNumeric(tDureevie_Ana_SOFT(indexDTR, 1)) Then
                    CalculerDVL = tDureevie_Ana_SOFT(indexDTR, 1)
                Else
                    CalculerDVL = tDureevie_Gen_SOFT(indexDTR, 1)
                End If
            Else
                If tDureevie_Ana_HARD(indexDTR, 1) <> "" And IsNumeric(tDureevie_Ana_HARD(indexDTR, 1)) Then
                    CalculerDVL = tDureevie_Ana_HARD(indexDTR, 1)
                Else
                    CalculerDVL = tDureevie_Gen_HARD(indexDTR, 1)
                End If
            End If
        End If
    End If
End Function

' =========================================================================
' PROCEDURE: POSER VALEUR APRES LIBELLE
' =========================================================================
' Cherche un libelle dans la zone d'entete (A1:K12), pose la valeur dans
' la cellule a droite du libelle trouve. Si aucun libelle ne matche,
' fallback sur la cellule par defaut (compat ancien template VBN).
' Rend la generation independante des coordonnees fixes du template.
Private Sub PoserValeurApresLibelle(ByRef dEst As Worksheet, _
                                      ByRef libelles As Variant, _
                                      ByVal valeur As Variant, _
                                      ByVal celluleFallback As String)
    Dim cellLibelle As Range
    Dim mc As Variant
    Dim plage As Range

    Set plage = dEst.Range("A1:K12")

    For Each mc In libelles
        Set cellLibelle = Nothing
        On Error Resume Next
        Set cellLibelle = plage.Find(What:=CStr(mc), LookIn:=xlValues, _
                                       lookat:=xlPart, MatchCase:=False)
        On Error GoTo 0
        If Not cellLibelle Is Nothing Then
            dEst.Cells(cellLibelle.row, cellLibelle.Column + 1).Value = valeur
            Exit Sub
        End If
    Next mc

    ' Fallback positions historiques (template VBN actuel)
    dEst.Range(celluleFallback).Value = valeur
End Sub

' =========================================================================
' PROCEDURE: MARQUER STATUT INDEX
' =========================================================================
Private Sub MarquerStatutIndex(ByRef index As Worksheet, ByVal indexDTR As Integer, ByVal statut As String)
    ' Statut en col F sur la ligne HARD du DTR (Tableau1 commence en row 6 :
    ' ligne 5 = entete, ligne 5 + ligneTableau = ligne reelle worksheet).
    Dim ligneWs As Long
    ligneWs = mapLigneHARD(indexDTR) + 5
    If langue = "FR" Then
        index.Cells(ligneWs, 6) = statut
    Else
        ' Traduction anglaise
        Select Case statut
            Case "A faire": index.Cells(ligneWs, 6) = "To do"
            Case "A remplir": index.Cells(ligneWs, 6) = "To fill"
            Case Else: index.Cells(ligneWs, 6) = statut
        End Select
    End If
End Sub

' =========================================================================
' PROCEDURE: CREER LIEN HYPERTEXTE
' =========================================================================
Private Sub CreerLienHypertexte(ByRef index As Worksheet, ByVal indexDTR As Integer, ByRef dEst As Worksheet)
    Dim colonneNom As String
    Dim ligneTab As Long

    If langue = "FR" Then
        colonneNom = "Lien Hypertexte"
    Else
        colonneNom = "Hyperlink"
    End If

    ' Ligne dans Tableau1 = ligne HARD du DTR (mapping pre-calcule)
    ligneTab = mapLigneHARD(indexDTR)

    With index
        .Hyperlinks.Add _
            Anchor:=.Range("Tableau1").ListObject.ListColumns(colonneNom).DataBodyRange(ligneTab), _
            Address:="", _
            SubAddress:="'" & dEst.Name & "'!A1", _
            TextToDisplay:=dEst.Name
    End With
End Sub

' =========================================================================
' PROCEDURE: FINALISER FICHIER V1
' =========================================================================
Private Sub FinaliserFichierV1(ByRef Template As Workbook, ByRef index As Worksheet, _
                                ByRef CoverPage As Worksheet, ByRef Template_HARD As Worksheet, _
                                ByRef Template_SOFT As Worksheet)
    Dim TabIndex As ListObject
    
    ' Remplir couverture
    CoverPage.Range("C13") = Code_Projet
    CoverPage.Range("B17") = Revision
    CoverPage.Range("C17") = Auteur
    CoverPage.Range("F11") = Nv_Maint_Projet
    
    ' Renommer colonne Index
    Set TabIndex = index.Range("Tableau1").ListObject
    If langue = "FR" Then
        TabIndex.ListColumns(5).Name = "A faire?"
    Else
        TabIndex.ListColumns(5).Name = "To do?"
    End If
    
    ' Supprimer templates
    Application.DisplayAlerts = False
    Template_SOFT.Delete
    Template_HARD.Delete
    Application.DisplayAlerts = True
    
    ' Fermer connexion BDD
    Close_CONN (Replace(CONNEXION_EXCEL_BDD_TRAVAIL, "UserProfile", userpath))
End Sub



