Attribute VB_Name = "B_ATM_V1"
Public Sub ATM_V1()
'Procédure de génération de fiches ATM V1

'Méthode :
' 1. Sélection d'un répertoire projet par l'utilisateur
' 2. Copie du template dans le répertoire choisie
' 3. Ouverture de ce template dans le répertoire
' 4. Récupération des données Info et Taches ATM par la feuile complète
' 5. Copie feuille template de chaque DTR de la PBS : HARD / SOFT ou HARD & SOFT
' 6. Remplissage de chaque feuille avec données taches et infos

'Procédure de génération de fiches ATM V1.0 (Version Refactorisée)
'Génère UN SEUL fichier Excel multi-feuilles avec Index + N fiches DTR

    Dim chemin As String
    Dim Template As Workbook
    Dim i As Integer
    
    ' Initialisation
    userpath = Environ$("UserProfile")
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set REQSEL = CreateObject("ADODB.Recordset")
    
    ' Sélection répertoire destination
    If Not SelectRepertoire(chemin) Then
        Exit Sub
    End If
    
    ' Copier et ouvrir template
    Set Template = CopierTemplateV1(chemin)
    If Template Is Nothing Then Exit Sub
    
    ' Récupération données BDD (réutilise fonction ATM V2)
    Call RecupererDonneesBDD_ATM(1)
    
    ' Cadrage tableaux
    Call Cadrage_Taches_ATM(Taches_tab)
    Call Cadrage_Infos_ATM(Infos_tab)
    
    ' Récupération PBS
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
    
    ' Boucle sur chaque DTR
    For i = 1 To UBound(tDTR)
        
        ' Skip supports/frames
        If InStr(1, tSS(i, 1), "support") > 0 Or InStr(1, tSS(i, 1), "frame") > 0 Then
            index.Range("F" & i + 5) = "NA"
            GoTo NextDTR
        End If
        
        ' Générer fiche(s) ATM pour ce DTR
        Call GenererFicheATM_V1(Template, index, Template_HARD, Template_SOFT, i)
        
NextDTR:
        ' Barre de progression
        Call barprogress(i, UBound(tDTR))
    Next i
    
    ' Finalisation
    Call FinaliserFichierV1(Template, index, CoverPage, Template_HARD, Template_SOFT)
    
    ' Message final
    If langue = "FR" Then
        MsgBox "Fiches ATM créées. Veuillez fermer cette fenêtre pour consulter le fichier généré.", vbInformation
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
' PROCÉDURE: REMPLIR INDEX V1
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
            
            ' Désignation (MEF prioritaire)
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
' PROCÉDURE: GÉNÉRER FICHE ATM V1 POUR UN DTR
' =========================================================================
Private Sub GenererFicheATM_V1(ByRef Template As Workbook, ByRef index As Worksheet, _
                               ByRef Template_HARD As Worksheet, ByRef Template_SOFT As Worksheet, _
                               ByVal indexDTR As Integer)
    Dim dest As Worksheet
    Dim system As String
    Dim test_taches As Boolean, test_infos As Boolean
    Dim PMB As Integer, PMH As Integer
    Dim cpt As Integer
    
    system = tSS(indexDTR, 1)
    cpt = 0
    
    ' Vérifier doublons noms feuilles
    cpt = CompterDoublonsFeuilles(Template, index, indexDTR, system)
    
    ' Copier template selon type HARD/SOFT
    Select Case tHARD_SOFT(indexDTR, 1)
        Case "", "HARD"
            Template_HARD.Copy After:=Template.Worksheets(Template.Worksheets.count)
            
        Case "SOFT"
            Template_SOFT.Copy After:=Template.Worksheets(Template.Worksheets.count)
            
        Case "BOTH"
            ' Créer les deux feuilles
            Template_SOFT.Copy After:=Template.Worksheets(Template.Worksheets.count)
            ActiveSheet.Name = index.Range("B" & indexDTR + 5) & "_" & Trim(Left(system, 4)) & "_SOFT"
            Template_HARD.Copy After:=Template.Worksheets(Template.Worksheets.count)
    End Select
    
    ' Nommer la feuille
    Set dest = ActiveSheet
    Call NommerFeuilleV1(dest, index, indexDTR, system, cpt)
    
    ' Rechercher plages dans BDD
    If Not RechercherPlagesDTR_V1(tDTR(indexDTR, 1), PMB, PMH, test_taches, test_infos) Then
        ' DTR non trouvé
        Call MarquerStatutIndex(index, indexDTR, "A faire")
        GoTo AjouterLien
    End If
    
    ' Remplir tâches
    Call RemplirTachesV1(dest, PMB, PMH, tHARD_SOFT(indexDTR, 1))
    
    ' Remplir niveau qualif
    dest.Range("I7") = Taches_tab(PMB, Taches_type.Niveau_Qualif)
    
    ' Remplir infos générales
    Call RemplirInfosV1(dest, indexDTR, test_infos)
    
AjouterLien:
    ' Créer lien hypertexte
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
' PROCÉDURE: NOMMER FEUILLE V1
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
' FONCTION: RECHERCHER PLAGES DTR
' =========================================================================
Public Function RechercherPlagesDTR_V1(ByVal dtr As String, _
                                        ByRef PMB As Integer, ByRef PMH As Integer, _
                                        ByRef test_taches As Boolean, ByRef test_infos As Boolean) As Boolean
    Dim j As Integer, l As Integer
    
    test_taches = False
    test_infos = False
    
    ' Rechercher première ligne tâches (PMB)
    For j = 2 To UBound(Taches_tab, 1)
        If Taches_tab(j, Taches_type.dtr) = dtr And Taches_tab(j, Taches_type.langue) = langue Then
            PMB = j
            test_taches = True
            Exit For
        End If
    Next j
    
    If Not test_taches Then
        RechercherPlagesDTR_V1 = False
        Exit Function
    End If
    
    ' Rechercher dernière ligne tâches (PMH)
    Do While j < UBound(Taches_tab, 1) - 1 And Taches_tab(j + 1, Taches_type.dtr) = Taches_tab(j, Taches_type.dtr)
        j = j + 1
    Loop
    PMH = j
    
    ' Rechercher dans Infos_ATM
    For l = 2 To UBound(Infos_tab, 1)
        If Infos_tab(l, Info_type.dtr) = dtr Then
            test_infos = True
            Exit For
        End If
    Next l
    
    RechercherPlagesDTR_V1 = True
End Function
' =========================================================================
' PROCÉDURE: REMPLIR TÂCHES V1 (VERSION SIMPLIFIÉE)
' =========================================================================
Private Sub RemplirTachesV1(ByRef dest As Worksheet, ByVal PMB As Integer, ByVal PMH As Integer, _
                            ByVal typeHARDSOFT As String)
    Dim j As Integer
    Dim nvx As Integer
    Dim rang As String
    Dim rangChiffre As String
    
    ' Parcourir tâches de la dernière à la première (insertion avant)
    For j = PMH To PMB Step -1
        
        ' Vérifier langue et type HARD
        If Taches_tab(j, Taches_type.langue) <> langue Then GoTo NextTache
        If Taches_tab(j, Taches_type.HARD_SOFT) <> "HARD" Then GoTo NextTache
        
        rang = Taches_tab(j, Taches_type.rang)
        
        ' Vérifier format rang (skip si >3 caractères)
        If Len(rang) > 3 Then GoTo NextTache
        
        ' Extraire niveau (ex: "5-2" ? 2)
        nvx = CInt(Right(rang, Len(rang) - InStr(1, rang, "-")))
        rangChiffre = Left(rang, 1)
        
        ' Insérer tâche selon rang
        Call InsererTacheV1(dest, j, rangChiffre, nvx)
        
NextTache:
    Next j
End Sub

' =========================================================================
' PROCÉDURE: INSÉRER TÂCHE V1
' =========================================================================
Private Sub InsererTacheV1(ByRef dest As Worksheet, ByVal indexTache As Integer, _
                           ByVal rangChiffre As String, ByVal niveau As Integer)
    Dim ligneBase As Integer
    Dim ligneInsert As Integer
    Dim temps As Double
    
    ' Déterminer ligne de base selon rang
    Select Case rangChiffre
        Case "1": ligneBase = LIGNE_RANG_1
        Case "2": ligneBase = LIGNE_RANG_2
        Case "3": ligneBase = LIGNE_RANG_3
        Case "4": ligneBase = LIGNE_RANG_4
        Case "5": ligneBase = LIGNE_RANG_5
        Case Else: Exit Sub
    End Select
    
    If niveau > 1 Then
        ' Insérer nouvelle ligne
        ligneInsert = ligneBase + 1
        dest.Range("B" & ligneInsert).EntireRow.Insert
        
        ' Copier format ligne au-dessus
        dest.Range("B" & ligneBase & ":J" & ligneBase).Copy dest.Range("B" & ligneInsert)
    Else
        ' Première tâche du rang (x-1)
        ligneInsert = ligneBase
    End If
    
    ' Remplir données
    With dest
        .Range("B" & ligneInsert) = Taches_tab(indexTache, Taches_type.rang)
        .Range("C" & ligneInsert) = Taches_tab(indexTache, Taches_type.description)
        .Range("E" & ligneInsert) = Taches_tab(indexTache, Taches_type.Critere_Dim)
        .Range("F" & ligneInsert) = Taches_tab(indexTache, Taches_type.OutillageSpe)
        .Range("G" & ligneInsert) = Taches_tab(indexTache, Taches_type.OutillageStd)
        .Range("H" & ligneInsert) = Taches_tab(indexTache, Taches_type.PieceRechangeCons)
        .Range("I" & ligneInsert) = Taches_tab(indexTache, Taches_type.NbOperateur)
        
        ' Temps avec gestion NULL
        temps = GetTempsFromBDD(Taches_tab(indexTache, Taches_type.TpsEstime))
        If temps > 0 Then
            .Range("J" & ligneInsert) = temps
            .Range("J" & ligneInsert).NumberFormat = "0.000"
        Else
            .Range("J" & ligneInsert) = ""
        End If
    End With
End Sub

' =========================================================================
' PROCÉDURE: REMPLIR INFOS V1
' =========================================================================
Private Sub RemplirInfosV1(ByRef dest As Worksheet, ByVal indexDTR As Integer, ByVal test_infos As Boolean)
    Dim j As Integer
    
    ' Données toujours présentes
    dest.Range("D4") = tDES_MEF(indexDTR, 1)
    dest.Range("D6") = tDTR(indexDTR, 1)
    dest.Range("D7") = Code_Projet
    dest.Range("D8") = tSS(indexDTR, 1)
    
    ' Remplir données RAM selon gate
    Call RemplirDonneesRAM_V1(dest, indexDTR)
    
    ' Rechercher dans Infos_ATM
    If Not test_infos Then Exit Sub
    
    For j = 2 To UBound(Infos_tab, 1)
        If Infos_tab(j, Info_type.dtr) = tDTR(indexDTR, 1) And Infos_tab(j, Info_type.HARD_SOFT) = "HARD" Then
            
            dest.Range("D9") = Infos_tab(j, Info_type.version)
            dest.Range("I6") = Infos_tab(j, Info_type.Poids)
            
            If langue = "EN" Then
                dest.Range("D5") = Infos_tab(j, Info_type.DescriptifEN)
                If Infos_tab(j, Info_type.Repa) = "REPARABLE" Then
                    dest.Range("I9") = "REPAIRABLE"
                Else
                    dest.Range("I9") = "SPARE"
                End If
            Else
                dest.Range("D5") = Infos_tab(j, Info_type.DescriptifFR)
                dest.Range("I8") = Infos_tab(j, Info_type.Plan)
                dest.Range("I9") = Infos_tab(j, Info_type.Repa)
            End If
            
            Exit For
        End If
    Next j
End Sub

' =========================================================================
' PROCÉDURE: REMPLIR DONNÉES RAM V1
' =========================================================================
Private Sub RemplirDonneesRAM_V1(ByRef dest As Worksheet, ByVal indexDTR As Integer)
    
    If AvantCGR Then
        ' Avant CGR: données génériques uniquement
        dest.Range("I4") = tMTBF_gen(indexDTR, 1)
        
        Select Case tHARD_SOFT(indexDTR, 1)
            Case "HARD", ""
                dest.Range("I5") = tDureevie_Gen_HARD(indexDTR, 1)
            Case "SOFT"
                dest.Range("I5") = tDureevie_Gen_SOFT(indexDTR, 1)
        End Select
    Else
        ' Après CGR
        If RAMReportDATA Then
            ' Toutes données RAM Report présentes
            dest.Range("I4") = tMTBF_fct(indexDTR, 1)
            
            Select Case tHARD_SOFT(indexDTR, 1)
                Case "HARD", "", "BOTH"
                    dest.Range("I5") = tDureevie_Ana_HARD(indexDTR, 1)
                Case "SOFT"
                    dest.Range("I5") = tDureevie_Ana_SOFT(indexDTR, 1)
            End Select
        Else
            ' Données RAM Report partielles: hiérarchie
            ' MTBF: fonctionnel ? intrinsèque ? générique
            If tMTBF_fct(indexDTR, 1) <> "" And IsNumeric(tMTBF_fct(indexDTR, 1)) Then
                dest.Range("I4") = tMTBF_fct(indexDTR, 1)
            ElseIf tMTBF_Intr(indexDTR, 1) <> "" And IsNumeric(tMTBF_Intr(indexDTR, 1)) Then
                dest.Range("I4") = tMTBF_Intr(indexDTR, 1)
            ElseIf tMTBF_gen(indexDTR, 1) <> "" And IsNumeric(tMTBF_gen(indexDTR, 1)) Then
                dest.Range("I4") = tMTBF_gen(indexDTR, 1)
            End If
            
            ' Durée vie: analytique ? générique
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
' PROCÉDURE: MARQUER STATUT INDEX
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
' PROCÉDURE: CRÉER LIEN HYPERTEXTE
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
' PROCÉDURE: FINALISER FICHIER V1
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

