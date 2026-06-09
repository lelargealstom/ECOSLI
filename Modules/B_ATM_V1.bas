Attribute VB_Name = "B_ATM_V1"
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
            index.Range("F" & i + 5) = "NA"
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
    
    Set TabIndex = index.Range("Tableau1").ListObject
    TabIndex.Resize Range("$B$5:$F$" & 5 + UBound(tDTR))
    
    ' Remplir tableau Index
    If langue = "FR" Then
        For i = 1 To UBound(tDTR)
            TabIndex.ListColumns("DTR").DataBodyRange(i).Value = tDTR(i, 1)
            
            ' Designation (MEF prioritaire)
            If tDES_MEF(i, 1) = "NA" Then
                TabIndex.ListColumns("Designation").DataBodyRange(i).Value = tDES_PBS(i, 1)
            Else
                TabIndex.ListColumns("Designation").DataBodyRange(i).Value = tDES_MEF(i, 1)
            End If
            
            TabIndex.ListColumns("Sous système").DataBodyRange(i).Value = tSS(i, 1)
        Next i
        
        ' Titre
        index.Cells(3, 2) = "ATM Projet " & Projet_Abrev & " Rev " & Revision
    Else
        For i = 1 To UBound(tDTR)
            TabIndex.ListColumns("DTR").DataBodyRange(i).Value = tDTR(i, 1)
            
            If tDES_MEF(i, 1) = "NA" Then
                TabIndex.ListColumns("Designation").DataBodyRange(i).Value = tDES_PBS(i, 1)
            Else
                TabIndex.ListColumns("Designation").DataBodyRange(i).Value = tDES_MEF(i, 1)
            End If
            
            TabIndex.ListColumns("Sub system").DataBodyRange(i).Value = tSS(i, 1)
        Next i
        
        ' Titre
        index.Cells(3, 2) = "MTA Project " & Projet_Abrev & " Rev " & Revision
    End If
End Sub
' =========================================================================
' PROCEDURE: GENERER FICHE ATM V1 POUR UN DTR
' =========================================================================
Private Sub GenererFicheATM_V1(ByRef Template As Workbook, ByRef index As Worksheet, _
                               ByRef Template_HARD As Worksheet, ByRef Template_SOFT As Worksheet, _
                               ByVal indexDTR As Integer)
    Dim dest As Worksheet
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

    ' Verifier doublons noms feuilles
    cpt = CompterDoublonsFeuilles(Template, index, indexDTR, system)

    ' Copier template selon type HARD/SOFT
    Select Case tHARD_SOFT(indexDTR, 1)
        Case "", "HARD"
            Template_HARD.Copy After:=Template.Worksheets(Template.Worksheets.count)
        Case "SOFT"
            Template_SOFT.Copy After:=Template.Worksheets(Template.Worksheets.count)
        Case "BOTH"
            Template_SOFT.Copy After:=Template.Worksheets(Template.Worksheets.count)
            ActiveSheet.Name = index.Range("B" & indexDTR + 5) & "_" & Trim(Left(system, 4)) & "_SOFT"
            Template_HARD.Copy After:=Template.Worksheets(Template.Worksheets.count)
    End Select

    ' Liberer le presse-papier apres Sheets.Copy (anti-erreur Clipboard).
    Application.CutCopyMode = False

    ' Nommer la feuille
    Set dest = ActiveSheet
    Call NommerFeuilleV1(dest, index, indexDTR, system, cpt)

    ' Si DTR pas dans la BdD pour cette langue : marquer "A faire"
    If Not hasTaches Then
        Call MarquerStatutIndex(index, indexDTR, "A faire")
        GoTo AjouterLien
    End If

    ' Remplir taches depuis le tableau local
    Call RemplirTachesV1(dest, arrTaches, tHARD_SOFT(indexDTR, 1))

    ' Remplir niveau qualif (1ere tache : col 5 = Niveau qualif)
    dest.Range("I7") = arrTaches(RC_TACHES_NIV_QUALIF, LBound(arrTaches, 2))

    ' Remplir infos generales
    Call RemplirInfosV1(dest, indexDTR, arrInfo, hasInfo)

AjouterLien:
    ' Creer lien hypertexte
    Call CreerLienHypertexte(index, indexDTR, dest)
End Sub

' =========================================================================
' FONCTION: COMPTER DOUBLONS FEUILLES
' =========================================================================
Private Function CompterDoublonsFeuilles(ByRef Template As Workbook, ByRef index As Worksheet, _
                                         ByVal indexDTR As Integer, ByVal system As String) As Integer
    Dim ws As Worksheet
    Dim system_temp As String
    Dim cpt As Integer
    
    cpt = 0
    For Each ws In Template.Worksheets
        system_temp = ws.Range("D8")
        If system_temp = "" Then system_temp = system
        
        If Trim(Left(ws.Name, Len(ws.Name) - 5)) = index.Range("B" & indexDTR + 5) & "_" & Trim(Left(system, 4)) Then
            cpt = cpt + 1
        End If
    Next
    
    CompterDoublonsFeuilles = cpt
End Function

' =========================================================================
' PROCEDURE: NOMMER FEUILLE V1
' =========================================================================
Private Sub NommerFeuilleV1(ByRef dest As Worksheet, ByRef index As Worksheet, _
                            ByVal indexDTR As Integer, ByVal system As String, ByVal cpt As Integer)
    Dim baseName As String
    Dim k As Integer
    Dim estnum As String
    
    baseName = index.Range("B" & indexDTR + 5) & "_" & Trim(Left(system, 4))
    
    If cpt = 0 Then
        ' Pas de doublon
        If tHARD_SOFT(indexDTR, 1) = "" Or tHARD_SOFT(indexDTR, 1) = "HARD" Then
            dest.Name = baseName & "_HARD"
        ElseIf tHARD_SOFT(indexDTR, 1) = "SOFT" Then
            dest.Name = baseName & "_SOFT"
        ElseIf tHARD_SOFT(indexDTR, 1) = "BOTH" Then
            dest.Name = baseName & "_HARD"
        End If
    Else
        ' Il y a des doublons
        For k = Worksheets.count To 1 Step -1
            If InStr(1, Worksheets(k).Name, baseName, vbTextCompare) = 1 Then
                estnum = Mid(Worksheets(k).Name, Len(Worksheets(k).Name) - 1, 1)
                If IsNumeric(estnum) Then
                    dest.Name = baseName & " (" & estnum + 1 & ")"
                    Exit For
                Else
                    dest.Name = baseName & " (" & cpt + 1 & ")"
                    Exit For
                End If
            End If
        Next k
    End If
End Sub
' =========================================================================
' PROCEDURE: REMPLIR TACHES V1 (depuis tableau local rs.GetRows())
' =========================================================================
' arrTaches est un tableau 2D (champ, ligne) 0-based renvoye par
' RecupererTachesDTR_BdD(...).GetRows(). Pre-filtre SQL : WHERE DTR + Langue,
' donc on ne filtre plus que sur HARD/SOFT cote VBA.
Private Sub RemplirTachesV1(ByRef dest As Worksheet, ByRef arrTaches As Variant, _
                            ByVal typeHARDSOFT As String)
    Dim j As Long
    Dim nvx As Integer
    Dim rang As String
    Dim rangChiffre As String

    ' Parcourir taches de la derniere a la premiere (insertion avant)
    For j = UBound(arrTaches, 2) To LBound(arrTaches, 2) Step -1

        ' Filtre HARD : seules les taches HARD sont placees dans une feuille
        ' V1 (la feuille SOFT est gere a part par le flow appelant).
        If CStr(arrTaches(RC_TACHES_HARD_SOFT, j)) <> "HARD" Then GoTo NextTache

        rang = CStr(arrTaches(RC_TACHES_RANG, j))

        ' Verifier format rang (skip si >3 caracteres)
        If Len(rang) > 3 Then GoTo NextTache

        ' Extraire niveau (ex: "5-2" -> 2)
        nvx = CInt(Right(rang, Len(rang) - InStr(1, rang, "-")))
        rangChiffre = Left(rang, 1)

        ' Inserer tache selon rang
        Call InsererTacheV1(dest, arrTaches, j, rangChiffre, nvx)

NextTache:
    Next j
End Sub

' =========================================================================
' PROCEDURE: INSERER TACHE V1 (depuis tableau local)
' =========================================================================
Private Sub InsererTacheV1(ByRef dest As Worksheet, ByRef arrTaches As Variant, _
                           ByVal indexTache As Long, _
                           ByVal rangChiffre As String, ByVal niveau As Integer)
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
        dest.Range("B" & ligneInsert).EntireRow.Insert _
            Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Else
        ' Premiere tache du rang (x-1)
        ligneInsert = ligneBase
    End If

    ' Remplir donnees depuis le tableau local
    With dest
        .Range("B" & ligneInsert) = arrTaches(RC_TACHES_RANG, indexTache)
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
Private Sub RemplirInfosV1(ByRef dest As Worksheet, ByVal indexDTR As Integer, _
                            ByRef arrInfo As Variant, ByVal hasInfo As Boolean)
    Dim j As Long

    ' Donnees toujours presentes (depuis la PBS)
    dest.Range("D4") = tDES_MEF(indexDTR, 1)
    dest.Range("D6") = tDTR(indexDTR, 1)
    dest.Range("D7") = Code_Projet
    dest.Range("D8") = tSS(indexDTR, 1)

    ' Remplir donnees RAM selon gate
    Call RemplirDonneesRAM_V1(dest, indexDTR)

    ' Si pas d'Info ATM en BdD, on s'arrete la (le DTR existe en PBS mais
    ' pas encore d'Info ATM en base - pas critique).
    If Not hasInfo Then Exit Sub

    ' Chercher la 1ere ligne HARD du DTR dans arrInfo
    For j = LBound(arrInfo, 2) To UBound(arrInfo, 2)
        If CStr(arrInfo(RC_INFO_HARD_SOFT, j)) = "HARD" Then

            dest.Range("D9") = arrInfo(RC_INFO_VERSION, j)
            dest.Range("I6") = arrInfo(RC_INFO_POIDS, j)

            If langue = "EN" Then
                dest.Range("D5") = arrInfo(RC_INFO_DESCRIPTIF_EN, j)
                If CStr(arrInfo(RC_INFO_REPARABILITE, j)) = "REPARABLE" Then
                    dest.Range("I9") = "REPAIRABLE"
                Else
                    dest.Range("I9") = "SPARE"
                End If
            Else
                dest.Range("D5") = arrInfo(RC_INFO_DESCRIPTIF_FR, j)
                dest.Range("I8") = arrInfo(RC_INFO_PLAN, j)
                dest.Range("I9") = arrInfo(RC_INFO_REPARABILITE, j)
            End If

            Exit For
        End If
    Next j
End Sub

' =========================================================================
' PROCEDURE: REMPLIR DONNEES RAM V1
' =========================================================================
Private Sub RemplirDonneesRAM_V1(ByRef dest As Worksheet, ByVal indexDTR As Integer)
    
    If AvantCGR Then
        ' Avant CGR: donnees generiques uniquement
        dest.Range("I4") = tMTBF_gen(indexDTR, 1)
        
        Select Case tHARD_SOFT(indexDTR, 1)
            Case "HARD", ""
                dest.Range("I5") = tDureevie_Gen_HARD(indexDTR, 1)
            Case "SOFT"
                dest.Range("I5") = tDureevie_Gen_SOFT(indexDTR, 1)
        End Select
    Else
        ' Apres CGR
        If RAMReportDATA Then
            ' Toutes donnees RAM Report presentes
            dest.Range("I4") = tMTBF_fct(indexDTR, 1)
            
            Select Case tHARD_SOFT(indexDTR, 1)
                Case "HARD", "", "BOTH"
                    dest.Range("I5") = tDureevie_Ana_HARD(indexDTR, 1)
                Case "SOFT"
                    dest.Range("I5") = tDureevie_Ana_SOFT(indexDTR, 1)
            End Select
        Else
            ' Donnees RAM Report partielles: hierarchie
            ' MTBF: fonctionnel ? intrinseque ? generique
            If tMTBF_fct(indexDTR, 1) <> "" And IsNumeric(tMTBF_fct(indexDTR, 1)) Then
                dest.Range("I4") = tMTBF_fct(indexDTR, 1)
            ElseIf tMTBF_Intr(indexDTR, 1) <> "" And IsNumeric(tMTBF_Intr(indexDTR, 1)) Then
                dest.Range("I4") = tMTBF_Intr(indexDTR, 1)
            ElseIf tMTBF_gen(indexDTR, 1) <> "" And IsNumeric(tMTBF_gen(indexDTR, 1)) Then
                dest.Range("I4") = tMTBF_gen(indexDTR, 1)
            End If
            
            ' Duree vie: analytique ? generique
            Select Case tHARD_SOFT(indexDTR, 1)
                Case "HARD", "", "BOTH"
                    If tDureevie_Ana_HARD(indexDTR, 1) <> "" And IsNumeric(tDureevie_Ana_HARD(indexDTR, 1)) Then
                        dest.Range("I5") = tDureevie_Ana_HARD(indexDTR, 1)
                    Else
                        dest.Range("I5") = tDureevie_Gen_HARD(indexDTR, 1)
                    End If
                Case "SOFT"
                    If tDureevie_Ana_SOFT(indexDTR, 1) <> "" And IsNumeric(tDureevie_Ana_SOFT(indexDTR, 1)) Then
                        dest.Range("I5") = tDureevie_Ana_SOFT(indexDTR, 1)
                    Else
                        dest.Range("I5") = tDureevie_Gen_SOFT(indexDTR, 1)
                    End If
            End Select
        End If
    End If
End Sub

' =========================================================================
' PROCEDURE: MARQUER STATUT INDEX
' =========================================================================
Private Sub MarquerStatutIndex(ByRef index As Worksheet, ByVal indexDTR As Integer, ByVal statut As String)
    If langue = "FR" Then
        index.Cells(indexDTR + 5, 6) = statut
    Else
        ' Traduction anglaise
        Select Case statut
            Case "A faire": index.Cells(indexDTR + 5, 6) = "To do"
            Case "A remplir": index.Cells(indexDTR + 5, 6) = "To fill"
            Case Else: index.Cells(indexDTR + 5, 6) = statut
        End Select
    End If
End Sub

' =========================================================================
' PROCEDURE: CREER LIEN HYPERTEXTE
' =========================================================================
Private Sub CreerLienHypertexte(ByRef index As Worksheet, ByVal indexDTR As Integer, ByRef dest As Worksheet)
    Dim colonneNom As String
    
    If langue = "FR" Then
        colonneNom = "Lien Hypertexte"
    Else
        colonneNom = "Hyperlink"
    End If
    
    With index
        .Hyperlinks.Add _
            Anchor:=.Range("Tableau1").ListObject.ListColumns(colonneNom).DataBodyRange(indexDTR), _
            Address:="", _
            SubAddress:="'" & dest.Name & "'!A1", _
            TextToDisplay:=dest.Name
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
    Close_CONN (Replace(CONNEXION_EXCEL_BDD_PROD, "UserProfile", userpath))
End Sub

