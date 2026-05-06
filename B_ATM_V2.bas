Attribute VB_Name = "B_ATM_V2"
Public Sub ATM_V2()
'Procédure de génération de fiches ATM 2.0
'Méthode :
' 1. Sélection d'un répertoire projet par l'utilisateur
' 2. Pour chaque DTR Copie du template 2.0 dans le répertoire choisie
' 3. Ouverture de ce template dans le répertoire
' 4. Récupération des données Info et Taches ATM par la feuile complète
' 5. Remplissage de chaque feuille avec données taches et infos
' 6. Même opération pour chaque DTR de la PBS (x classeurs pour x DTR)
'Procédure de génération de fiches ATM 2.0 (Version Refactorisée)
'Génère 1 fichier Excel par DTR avec insertion dynamique par chapitres
    
    Dim Tbl() As String
    Dim bibli As String

    
    'Initialisation parametres user
    userpath = Environ$("UserProfile")
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set REQSEL = CreateObject("ADODB.Recordset")
    
    ' Sélection répertoire destination
    If Not SelectRepertoire(chemin) Then
        TurnOnFunctionality
        Exit Sub
    End If
    
    ' Récupération données BDD
    Call RecupererDonneesBDD_ATM(2)
    
    'Cadrage des deux onglets importants de la BDD
    Call Cadrage_Taches_ATM(Taches_tab)
    Call Cadrage_Infos_ATM(Infos_tab)
    
    'Récupération de la PBS
    Set DS1 = Generateur.Worksheets("DATA_ASL")
     
    ' Énumération bibliothèque illustrée
    bibli = userpath & BIBLIO_PATH
    Tbl = EnumFichiers(bibli)
    
    ' Boucle sur chaque DTR
    For i = 1 To UBound(tDTR)
        
        ' Vérifier si fichier illustré existe
        If RechercherFichierIllustre(Tbl, tDTR(i, 1), bibli, chemin) Then
            GoTo NextDTR
        End If
        
        ' Vérifier présence BDD
        Presence_BDD = IsInArray(CStr(tDTR(i, 1)), Taches_tab)
        
        If Presence_BDD = -1 Then
            ' DTR connu dans BDD ? Génération avec données
            Call GenererFicheATM(tDTR(i, 1), tDES_PBS(i, 1), chemin, i)
        Else
            ' DTR inconnu dans BDD ? Copie template vide
            Call CopierTemplateVide(tDTR(i, 1), tDES_PBS(i, 1), chemin)
        End If
        
NextDTR:
        ' Barre de progression
        Call barprogress(i, UBound(tDTR))
    Next i
    
    ' Fermeture connexion
    Close_CONN (Replace(CONNEXION_EXCEL_BDD_PROD, "UserProfile", userpath))
    
    ' Message final
    TurnOnFunctionality
    If langue = "FR" Then
        MsgBox "ATM 2.0 créées. Veuillez fermer cette fenêtre pour consulter les fichiers générés.", vbInformation
    Else
        MsgBox "MTA 2.0 created. Please close this window to consult the new files generated.", vbInformation
    End If


End Sub


' =========================================================================
' FONCTION: RECHERCHER FICHIER ILLUSTRÉ
' =========================================================================
Private Function RechercherFichierIllustre(ByRef Tbl() As String, ByVal dtr As String, _
                                           ByVal bibli As String, ByVal chemin As String) As Boolean
    Dim j As Integer
    
    For j = 1 To UBound(Tbl)
        If Left(Tbl(j), 16) = dtr & "_" & langue Then
            ' Fichier illustré trouvé ? Copie directe
            fso.CopyFile bibli & "\" & Tbl(j), chemin & "\" & Tbl(j)
            RechercherFichierIllustre = True
            Exit Function
        End If
    Next j
    
    RechercherFichierIllustre = False
End Function
' =========================================================================
' PROCÉDURE: COPIER TEMPLATE VIDE
' =========================================================================
Private Sub CopierTemplateVide(ByVal dtr As String, ByVal Designation As String, ByVal chemin As String)
    Dim nomFichier As String
    Dim templatePath As String
    
    ' Construire nom fichier
    nomFichier = "(vide)" & dtr & "_" & langue & "_" & Replace(Left(Designation, 10), "/", "_") & ".xlsx"
    
    ' Sélectionner template
    If langue = "FR" Then
        templatePath = userpath & TEMPLATE_V2_FR_PATH
    Else
        templatePath = userpath & TEMPLATE_V2_EN_PATH
    End If
    
    ' Copier template
    fso.CopyFile templatePath, chemin & "\" & nomFichier
End Sub
' =========================================================================
' PROCÉDURE PRINCIPALE: GÉNÉRER FICHE ATM
' =========================================================================
Private Sub GenererFicheATM(ByVal dtr As String, ByVal Designation As String, _
                            ByVal chemin As String, ByVal indexDTR As Integer)
    Dim nomFichier As String
    Dim templatePath As String
    Dim dest As Worksheet
    Dim CoverPage As Worksheet
    Dim PMB As Integer, PMH As Integer, PMJ As Integer
    Dim test_taches As Boolean, test_infos As Boolean
    
    ' Construire nom fichier
    nomFichier = "(non illustré)" & dtr & "_" & langue & "_" & Replace(Left(Designation, 10), "/", "_") & ".xlsx"
    
    ' Sélectionner template
    If langue = "FR" Then
        templatePath = userpath & TEMPLATE_V2_FR_PATH
    Else
        templatePath = userpath & TEMPLATE_V2_EN_PATH
    End If
    
    ' Copier et ouvrir template
    fso.CopyFile templatePath, chemin & "\" & nomFichier
    Workbooks.Open chemin & "\" & nomFichier
    
    ' Identifier feuilles
    If langue = "FR" Then
        Set dest = ActiveWorkbook.Sheets("DESCRIPTION_TACHE")
    Else
        Set dest = ActiveWorkbook.Sheets("TASK_DESCRIPTION")
    End If
    Set CoverPage = ActiveWorkbook.Sheets(1)
    
    ' Rechercher plages DTR dans Taches_tab et Infos_tab
    If Not RechercherPlagesDTR_V2(dtr, PMB, PMH, PMJ, test_taches, test_infos) Then
        ' DTR connu mais pas dans la bonne langue ? Sauvegarder comme vide
        ActiveWorkbook.SaveAs chemin & "\(vide)" & dtr & "_" & langue & "_" & Left(Designation, 10) & ".xlsx"
        ActiveWorkbook.Close
        Exit Sub
    End If
    
    ' Activer feuille destination
    dest.Activate
    
    ' REMPLISSAGE PAR CHAPITRES (VERSION SIMPLIFIÉE)
    Call RemplirTousChapitres(dest, PMB, PMH)
    
    ' Ajouter Chapitre E (Notes intégrateur)
    Call AjouterChapitreE(dest)
    
    ' Remplir page de couverture
    Call RemplirCouverture(CoverPage, dtr, PMJ, test_infos)
    
    ' Sauvegarder et fermer
    ActiveWorkbook.Save
    ActiveWorkbook.Close
End Sub
Private Function RechercherPlagesDTR_V2(ByVal dtr As String, _
                                        ByRef PMB As Integer, ByRef PMH As Integer, ByRef PMJ As Integer, _
                                        ByRef test_taches As Boolean, ByRef test_infos As Boolean) As Boolean

Dim j As Integer, l As Integer
    
    test_taches = False
    test_infos = False
    PMB = 0
    PMH = 0
    PMJ = 0
    
    ' =====================================================================
    ' RECHERCHE DANS TACHES_ATM
    ' =====================================================================
    ' Rechercher première ligne tâches (PMB) avec DTR + Langue
    For j = 2 To UBound(Taches_tab, 1)
        If Taches_tab(j, Taches_type.dtr) = dtr And _
           Taches_tab(j, Taches_type.langue) = langue Then
            PMB = j
            test_taches = True
            Exit For
        End If
    Next j
    
    ' Si langue non trouvée, retourner False
    If Not test_taches Then
        RechercherPlagesDTR_V2 = False
        Exit Function
    End If
    
    ' Rechercher dernière ligne tâches (PMH)
    ' V2 vérifie AUSSI la langue dans la continuation (plus strict que V1)
    Do While j < UBound(Taches_tab, 1)
        If Taches_tab(j + 1, Taches_type.dtr) = Taches_tab(j, Taches_type.dtr) And _
           Taches_tab(j + 1, Taches_type.langue) = langue Then
            j = j + 1
        Else
            Exit Do
        End If
    Loop
    PMH = j
    
    ' =====================================================================
    ' RECHERCHE DANS INFOS_ATM
    ' =====================================================================
    ' Important pour V2: PMJ utilisé dans RemplirCouverture (F30)
    For l = 2 To UBound(Infos_tab, 1)
        If Infos_tab(l, Info_type.dtr) = dtr Then
            PMJ = l
            test_infos = True
            Exit For
        End If
    Next l
    
    RechercherPlagesDTR_V2 = True
End Function
' =========================================================================
' PROCÉDURE PRINCIPALE: REMPLIR TOUS CHAPITRES (VERSION SIMPLIFIÉE)
' =========================================================================
Private Sub RemplirTousChapitres(ByRef dest As Worksheet, ByVal PMB As Integer, ByVal PMH As Integer)
    Dim j As Integer
    Dim rang As String
    Dim rangPremierChiffre As String
    Dim compteurSousChapitres As Integer
    Dim chapitres(1 To 4) As ChapitreInfo
    
    ' Initialiser structure chapitres
    Call InitialiserChapitres(dest, chapitres)
    
    compteurSousChapitres = 1
    
    ' Parcourir toutes les tâches du DTR
    For j = PMB To PMH
        
        ' Vérifier langue
        If Taches_tab(j, Taches_type.langue) <> langue Then
            GoTo NextTache
        End If
        
        rang = Taches_tab(j, Taches_type.rang)
        rangPremierChiffre = Left(rang, 1)
        
        ' Router vers le bon chapitre selon le rang
        Select Case rangPremierChiffre
            Case "1"
                Call InsererTacheChapitre(dest, chapitres, 1, j, True)  ' Chapitre A
                
            Case "2"
                Call InsererTacheChapitre(dest, chapitres, 2, j, True) ' Chapitre B
                
            Case "3", "4"
                Call InsererTacheChapitreC(dest, chapitres, 3, j, compteurSousChapitres) ' Chapitre C
                
            Case "5"
                Call InsererTacheChapitre(dest, chapitres, 4, j, True)  ' Chapitre D
        End Select
        
NextTache:
    Next j
    
    ' Écrire les sommes de temps dans chaque chapitre
    Call EcrireSommesTemps(dest, chapitres)
End Sub
' =========================================================================
' PROCÉDURE: INITIALISER CHAPITRES
' =========================================================================
Private Sub InitialiserChapitres(ByRef dest As Worksheet, ByRef chapitres() As ChapitreInfo)
    
    ' Chapitre A
    chapitres(1).Marqueur = "A"
    chapitres(1).ligneDebut = dest.Range("B:B").Find(What:="A", LookIn:=xlValues, lookat:=xlWhole).Row
    chapitres(1).LigneProchainChap = dest.Range("B:B").Find(What:="B", LookIn:=xlValues, lookat:=xlWhole).Row
    chapitres(1).TotalTemps = 0
    chapitres(1).NombreInsertions = 0
    
    ' Chapitre B
    chapitres(2).Marqueur = "B"
    chapitres(2).ligneDebut = dest.Range("B:B").Find(What:="B", LookIn:=xlValues, lookat:=xlWhole).Row
    chapitres(2).LigneProchainChap = dest.Range("B:B").Find(What:="C", LookIn:=xlValues, lookat:=xlWhole).Row
    chapitres(2).TotalTemps = 0
    chapitres(2).NombreInsertions = 0
    
    ' Chapitre C
    chapitres(3).Marqueur = "C"
    chapitres(3).ligneDebut = dest.Range("B:B").Find(What:="C", LookIn:=xlValues, lookat:=xlWhole).Row
    chapitres(3).LigneProchainChap = dest.Range("B:B").Find(What:="D", LookIn:=xlValues, lookat:=xlWhole).Row
    chapitres(3).TotalTemps = 0
    chapitres(3).NombreInsertions = 0
    
    ' Chapitre D
    chapitres(4).Marqueur = "D"
    chapitres(4).ligneDebut = dest.Range("B:B").Find(What:="D", LookIn:=xlValues, lookat:=xlWhole).Row
    chapitres(4).LigneProchainChap = -1 ' Dernier chapitre
    chapitres(4).TotalTemps = 0
    chapitres(4).NombreInsertions = 0
End Sub

' =========================================================================
' PROCÉDURE: INSÉRER TÂCHE CHAPITRE (A, B, D)
' =========================================================================
Private Sub InsererTacheChapitre(ByRef dest As Worksheet, ByRef chapitres() As ChapitreInfo, _
                                 ByVal numChapitre As Integer, ByVal indexTache As Integer, _
                                 ByVal cumulerTemps As Boolean)
    Dim ligvide As Long  ' IMPORTANT: Long au lieu de Integer pour éviter overflow
    Dim temps As Double
    Dim k As Integer
    
    ' Déterminer ligne d'insertion
    If chapitres(numChapitre).LigneProchainChap > 0 Then
        ' Cas des Chapitres A, B, C (ont un chapitre suivant)
        If dest.Range("B" & (chapitres(numChapitre).LigneProchainChap - 1)).Value <> "" Then
            ' Insérer ligne avant le prochain chapitre
            dest.Range("B" & chapitres(numChapitre).LigneProchainChap).EntireRow.Insert
            ligvide = chapitres(numChapitre).LigneProchainChap
            
            ' Mettre à jour TOUS les chapitres suivants après insertion
            For k = numChapitre To UBound(chapitres)
                If chapitres(k).LigneProchainChap > 0 Then
                    chapitres(k).LigneProchainChap = chapitres(k).LigneProchainChap + 1
                End If
            Next k
            
            For k = numChapitre + 1 To UBound(chapitres)
                chapitres(k).ligneDebut = chapitres(k).ligneDebut + 1
            Next k
        Else
            ' Utiliser la ligne vide existante
            ligvide = chapitres(numChapitre).LigneProchainChap - 1
        End If
    Else
        ' *** CORRECTION CHAPITRE D ***
        ' Chapitre D (dernier chapitre): chercher la première ligne vide
        ' Ne PAS utiliser End(xlDown) car il descend jusqu'à ligne 1048576 si pas de données
        
        ligvide = chapitres(numChapitre).ligneDebut + 1
        
        ' Parcourir vers le bas pour trouver la première ligne vide
        Do While dest.Range("B" & ligvide).Value <> ""
            ligvide = ligvide + 1
            ' Sécurité: limite à 500 lignes après le début du chapitre
            If ligvide > chapitres(numChapitre).ligneDebut + 500 Then
                Exit Do
            End If
        Loop
    End If
    
    ' Remplir la ligne
    With dest
        .Range("B" & ligvide) = Right(Taches_tab(indexTache, Taches_type.rang), 1)
        .Range("C" & ligvide) = Taches_tab(indexTache, Taches_type.categorie)
        .Range("D" & ligvide) = Taches_tab(indexTache, Taches_type.description)
        
        ' Gestion temps avec fonction utilitaire
        temps = GetTempsFromBDD(Taches_tab(indexTache, Taches_type.TpsEstime))
        If temps > 0 Then
            .Range("L" & ligvide) = temps
            .Range("L" & ligvide).NumberFormat = "0.000"
        Else
            .Range("L" & ligvide) = ""
        End If
    End With
    
    ' Cumuler temps si demandé (Chapitres A et D) et si temps valide
    If cumulerTemps And temps > 0 Then
        chapitres(numChapitre).TotalTemps = chapitres(numChapitre).TotalTemps + temps
    End If
    
    chapitres(numChapitre).NombreInsertions = chapitres(numChapitre).NombreInsertions + 1
End Sub

' =========================================================================
' PROCÉDURE: INSÉRER TÂCHE CHAPITRE C (avec sous-chapitres)
' =========================================================================
Private Sub InsererTacheChapitreC(ByRef dest As Worksheet, ByRef chapitres() As ChapitreInfo, _
                                  ByVal numChapitre As Integer, ByVal indexTache As Integer, _
                                  ByRef compteurSousChapitres As Integer)
    Dim ligvide As Integer
    Dim rang As String
    Dim premierMot As String
    Dim description As String
    Dim posEspace As Integer
    Dim temps As Double
    Dim k As Integer
    
    rang = Taches_tab(indexTache, Taches_type.rang)
    description = Taches_tab(indexTache, Taches_type.description)
    
    ' Déterminer ligne d'insertion
    If dest.Range("B" & (chapitres(numChapitre).LigneProchainChap - 1)).Value <> "" Then
        dest.Range("B" & chapitres(numChapitre).LigneProchainChap).EntireRow.Insert
        ligvide = chapitres(numChapitre).LigneProchainChap
        
        ' *** CORRECTION CRITIQUE ***
        ' Mettre à jour TOUS les chapitres suivants après insertion
        For k = numChapitre To UBound(chapitres)
            If chapitres(k).LigneProchainChap > 0 Then
                chapitres(k).LigneProchainChap = chapitres(k).LigneProchainChap + 1
            End If
        Next k
        
        For k = numChapitre + 1 To UBound(chapitres)
            chapitres(k).ligneDebut = chapitres(k).ligneDebut + 1
        Next k
    Else
        ligvide = chapitres(numChapitre).LigneProchainChap - 1
    End If
    
    ' Si premier item d'un sous-chapitre (rang x.1), créer l'en-tête du sous-chapitre
    If Right(rang, 1) = "1" Then
        ' Ligne sous-chapitre
        With dest
            .Range("B" & ligvide) = "B" & compteurSousChapitres
            If langue = "FR" Then
                .Range("C" & ligvide) = "SOUS-CHAPITRE"
                If Left(rang, 1) = "3" Then
                    .Range("D" & ligvide) = "DEPOSE"
                Else ' "4"
                    .Range("D" & ligvide) = "POSE"
                End If
            Else
                .Range("C" & ligvide) = "SUB-CHAPTER"
                If Left(rang, 1) = "3" Then
                    .Range("D" & ligvide) = "REMOVE"
                Else ' "4"
                    .Range("D" & ligvide) = "REFIT"
                End If
            End If
        End With
        
        ' Insérer ligne suivante pour la tâche
        dest.Range("B" & (ligvide + 1)).EntireRow.Insert
        ligvide = ligvide + 1
        
        ' *** CORRECTION CRITIQUE ***
        ' Mettre à jour chapitres suivants après 2ème insertion
        For k = numChapitre To UBound(chapitres)
            If chapitres(k).LigneProchainChap > 0 Then
                chapitres(k).LigneProchainChap = chapitres(k).LigneProchainChap + 1
            End If
        Next k
        
        For k = numChapitre + 1 To UBound(chapitres)
            chapitres(k).ligneDebut = chapitres(k).ligneDebut + 1
        Next k
        
        compteurSousChapitres = compteurSousChapitres + 1
    End If
    
    ' Extraire premier mot de la description
    posEspace = InStr(1, description, " ")
    If posEspace = 0 Then
        premierMot = description
    Else
        premierMot = Left(description, posEspace - 1)
    End If
    
    ' Remplir la ligne de tâche
    With dest
        .Range("B" & ligvide) = Right(rang, 1)
        .Range("C" & ligvide) = premierMot
        .Range("D" & ligvide) = description
        
        ' Opérateurs
        If langue = "FR" Then
            .Range("G" & ligvide) = "1 x ELECTRICIEN"
        Else
            .Range("G" & ligvide) = "1 x ELECTRICIAN"
        End If
        
        ' Pièces de rechange
        If Taches_tab(indexTache, Taches_type.PieceRechangeCons) <> "" Then
            .Range("I" & ligvide) = "1 x " & Taches_tab(indexTache, Taches_type.PieceRechangeCons)
        End If
        
        ' Temps - Gestion avec fonction utilitaire
        temps = GetTempsFromBDD(Taches_tab(indexTache, Taches_type.TpsEstime))
        If temps > 0 Then
            .Range("L" & ligvide) = temps
            .Range("L" & ligvide).NumberFormat = "0.000"
        Else
            .Range("L" & ligvide) = ""
        End If
    End With
    
    ' Cumuler temps seulement si valide
    If temps > 0 Then
        chapitres(numChapitre).TotalTemps = chapitres(numChapitre).TotalTemps + temps
    End If
    
    chapitres(numChapitre).NombreInsertions = chapitres(numChapitre).NombreInsertions + 1
End Sub

' =========================================================================
' PROCÉDURE: ÉCRIRE SOMMES TEMPS
' =========================================================================
Private Sub EcrireSommesTemps(ByRef dest As Worksheet, ByRef chapitres() As ChapitreInfo)
    Dim i As Integer
    
    ' Écrire le total de temps dans la colonne L de chaque ligne de chapitre
    For i = 1 To 4
        If chapitres(i).TotalTemps > 0 Then
            dest.Range("L" & chapitres(i).ligneDebut) = chapitres(i).TotalTemps
            dest.Range("L" & chapitres(i).ligneDebut).NumberFormat = "0.000"
        End If
    Next i
End Sub

' =========================================================================
' PROCÉDURE: AJOUTER CHAPITRE E
' =========================================================================
Private Sub AjouterChapitreE(ByRef dest As Worksheet)
    Dim ligvide As Integer
    
    ' Trouver dernière ligne remplie
    ligvide = dest.Range("B10000").End(xlUp).Row + 1
    
    ' Copier format de la ligne des totaux (ligne 2)
    dest.Range("B2:L2").Copy dest.Range("B" & ligvide)
    
    ' Marqueur Chapitre E
    dest.Range("B" & ligvide) = "E"
    
    ' Ligne Note
    ligvide = dest.Range("B10000").End(xlUp).Row + 1
    dest.Range("B" & ligvide) = 1
    dest.Range("C" & ligvide) = "Note"
    
    If langue = "FR" Then
        dest.Range("D" & ligvide) = "A definir par integrateur"
    Else
        dest.Range("D" & ligvide) = "To be defined by integrator"
    End If
End Sub

' =========================================================================
' PROCÉDURE: REMPLIR COUVERTURE
' =========================================================================
Private Sub RemplirCouverture(ByRef CoverPage As Worksheet, ByVal dtr As String, _
                              ByVal PMJ As Integer, ByVal test_infos As Boolean)
    
    ' Projet
    CoverPage.Range("F26") = Projet_Abrev
    
    ' DTR
    CoverPage.Range("F28") = dtr
    
    ' Descriptif (selon langue)
    If Not test_infos Then
        CoverPage.Range("F30") = ""
    Else
        If langue = "FR" Then
            CoverPage.Range("F30") = Infos_tab(PMJ, Info_type.DescriptifFR)
        Else
            CoverPage.Range("F30") = Infos_tab(PMJ, Info_type.DescriptifEN)
        End If
    End If
End Sub


