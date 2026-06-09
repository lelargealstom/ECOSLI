Attribute VB_Name = "B_ATM_V2"
' Cache du template local : copie unique depuis SharePoint en debut d'ATM_V2.
' Toutes les copies par DTR partent ensuite de cette copie locale (disque
' local ~10x plus rapide que OneDrive synchronise).
Private templateLocalV2 As String

Public Sub ATM_V2()
'Procedure de generation de fiches ATM 2.0
'Methode :
' 1. Selection d'un repertoire projet par l'utilisateur
' 2. Pour chaque DTR Copie du template 2.0 dans le repertoire choisie
' 3. Ouverture de ce template dans le repertoire
' 4. Recuperation des donnees Info et Taches ATM par la feuile complete
' 5. Remplissage de chaque feuille avec donnees taches et infos
' 6. Meme operation pour chaque DTR de la PBS (x classeurs pour x DTR)
'Procedure de generation de fiches ATM 2.0 (Version Refactorisee)
'Genere 1 fichier Excel par DTR avec insertion dynamique par chapitres
    
    Dim Tbl() As String
    Dim bibli As String

    
    'Initialisation parametres user
    userpath = Environ$("UserProfile")
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set REQSEL = CreateObject("ADODB.Recordset")
    
    ' Selection repertoire destination
    If Not SelectRepertoire(chemin) Then
        TurnOnFunctionality
        Exit Sub
    End If
    
    ' Recuperation donnees BDD
    Call RecupererDonneesBDD_ATM(2)
    
    'Cadrage des deux onglets importants de la BDD
    Call Cadrage_Taches_ATM(Taches_tab)
    Call Cadrage_Infos_ATM(Infos_tab)
    
    'Recuperation de la PBS
    Set DS1 = Generateur.Worksheets("DATA_ASL")
     
    ' enumeration bibliotheque illustree
    bibli = userpath & BIBLIO_PATH
    Tbl = EnumFichiers(bibli)

    ' Copie locale unique du template (SharePoint -> %TEMP%) pour eviter
    ' N copies depuis OneDrive synchronise (lent). Reutilisee a chaque
    ' DTR via templateLocalV2 dans GenererFicheATM / CopierTemplateVide.
    templateLocalV2 = Environ$("TEMP") & "\ECOSLI_template_ATM_V2_" & langue & ".xlsx"
    On Error Resume Next
    If langue = "FR" Then
        fso.CopyFile userpath & TEMPLATE_V2_FR_PATH, templateLocalV2, True
    Else
        fso.CopyFile userpath & TEMPLATE_V2_EN_PATH, templateLocalV2, True
    End If
    On Error GoTo 0

    ' UX : initialiser la barre de progression + StatusBar + curseur sablier.
    Call Progress_Demarrer(UBound(tDTR), "ATM V2")

    ' Boucle sur chaque DTR
    For i = 1 To UBound(tDTR)

        ' UX : mise a jour barre + StatusBar avant le traitement de ce DTR.
        Call Progress_Etape(i, "", CStr(tDTR(i, 1)))

        ' Verifier si fichier illustre existe
        If RechercherFichierIllustre(Tbl, tDTR(i, 1), bibli, chemin) Then
            GoTo NextDTR
        End If

        ' Verifier presence BDD
        Presence_BDD = IsInArray(CStr(tDTR(i, 1)), Taches_tab)

        If Presence_BDD = -1 Then
            ' DTR connu dans BDD ? Generation avec donnees
            Call GenererFicheATM(tDTR(i, 1), tDES_PBS(i, 1), chemin, i)
        Else
            ' DTR inconnu dans BDD ? Copie template vide
            Call CopierTemplateVide(tDTR(i, 1), tDES_PBS(i, 1), chemin)
        End If

NextDTR:
    Next i

    ' UX : restaurer curseur normal, StatusBar, masquer la barre.
    Call Progress_Terminer

    ' Fermeture connexion
    Close_CONN (Replace(CONNEXION_EXCEL_BDD_PROD, "UserProfile", userpath))

    ' Message final
    TurnOnFunctionality
    If langue = "FR" Then
        MsgBox "ATM 2.0 creees. Veuillez fermer cette fenetre pour consulter les fichiers generes.", vbInformation
    Else
        MsgBox "MTA 2.0 created. Please close this window to consult the new files generated.", vbInformation
    End If


End Sub


' =========================================================================
' FONCTION: RECHERCHER FICHIER ILLUSTRE
' =========================================================================
Private Function RechercherFichierIllustre(ByRef Tbl() As String, ByVal dtr As String, _
                                           ByVal bibli As String, ByVal chemin As String) As Boolean
    Dim j As Integer
    
    For j = 1 To UBound(Tbl)
        If Left(Tbl(j), 16) = dtr & "_" & langue Then
            ' Fichier illustre trouve ? Copie directe
            fso.CopyFile bibli & "\" & Tbl(j), chemin & "\" & Tbl(j)
            RechercherFichierIllustre = True
            Exit Function
        End If
    Next j
    
    RechercherFichierIllustre = False
End Function
' =========================================================================
' PROCEDURE: COPIER TEMPLATE VIDE
' =========================================================================
Private Sub CopierTemplateVide(ByVal dtr As String, ByVal Designation As String, ByVal chemin As String)
    Dim nomFichier As String

    ' Construire nom fichier
    nomFichier = "(vide)" & dtr & "_" & langue & "_" & Replace(Left(Designation, 10), "/", "_") & ".xlsx"

    ' Copier depuis le cache local (templateLocalV2 initialise dans ATM_V2).
    fso.CopyFile templateLocalV2, chemin & "\" & nomFichier
End Sub
' =========================================================================
' PROCEDURE PRINCIPALE: GENERER FICHE ATM
' =========================================================================
Private Sub GenererFicheATM(ByVal dtr As String, ByVal Designation As String, _
                            ByVal chemin As String, ByVal indexDTR As Integer)
    Dim nomFichier As String
    Dim dest As Worksheet
    Dim CoverPage As Worksheet
    Dim PMB As Integer, PMH As Integer, PMJ As Integer
    Dim test_taches As Boolean, test_infos As Boolean

    ' Construire nom fichier
    nomFichier = "(non illustre)" & dtr & "_" & langue & "_" & Replace(Left(Designation, 10), "/", "_") & ".xlsx"

    ' Copier depuis le cache local (10x plus rapide que SharePoint) puis ouvrir.
    fso.CopyFile templateLocalV2, chemin & "\" & nomFichier
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
    
    ' REMPLISSAGE PAR CHAPITRES (VERSION SIMPLIFIEE)
    Call RemplirTousChapitres(dest, PMB, PMH)
    
    ' Ajouter Chapitre E (Notes integrateur)
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
    ' Rechercher premiere ligne taches (PMB) avec DTR + Langue
    For j = 2 To UBound(Taches_tab, 1)
        If Taches_tab(j, Taches_type.dtr) = dtr And _
           Taches_tab(j, Taches_type.langue) = langue Then
            PMB = j
            test_taches = True
            Exit For
        End If
    Next j

    ' Si langue non trouvee, retourner False
    If Not test_taches Then
        RechercherPlagesDTR_V2 = False
        Exit Function
    End If

    ' Rechercher derniere ligne taches (PMH)
    ' V2 verifie AUSSI la langue dans la continuation (plus strict que V1)
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
    ' Important pour V2: PMJ utilise dans RemplirCouverture (F30)
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
' PROCEDURE PRINCIPALE: REMPLIR TOUS CHAPITRES (VERSION SIMPLIFIEE)
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
    
    ' Parcourir toutes les taches du DTR
    For j = PMB To PMH
        
        ' Verifier langue
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
    
    ' ecrire les sommes de temps dans chaque chapitre
    Call EcrireSommesTemps(dest, chapitres)
End Sub
' =========================================================================
' PROCEDURE: INITIALISER CHAPITRES
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
' PROCEDURE: INSERER TACHE CHAPITRE (A, B, D)
' =========================================================================
Private Sub InsererTacheChapitre(ByRef dest As Worksheet, ByRef chapitres() As ChapitreInfo, _
                                 ByVal numChapitre As Integer, ByVal indexTache As Integer, _
                                 ByVal cumulerTemps As Boolean)
    Dim ligvide As Long  ' IMPORTANT: Long au lieu de Integer pour eviter overflow
    Dim temps As Double
    Dim k As Integer
    
    ' Determiner ligne d'insertion
    If chapitres(numChapitre).LigneProchainChap > 0 Then
        ' Cas des Chapitres A, B, C (ont un chapitre suivant)
        If dest.Range("B" & (chapitres(numChapitre).LigneProchainChap - 1)).Value <> "" Then
            ' Inserer ligne avant le prochain chapitre
            dest.Range("B" & chapitres(numChapitre).LigneProchainChap).EntireRow.Insert
            ligvide = chapitres(numChapitre).LigneProchainChap
            
            ' Mettre a jour TOUS les chapitres suivants apres insertion
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
        ' Chapitre D (dernier chapitre): chercher la premiere ligne vide
        ' Ne PAS utiliser End(xlDown) car il descend jusqu'e ligne 1048576 si pas de donnees
        
        ligvide = chapitres(numChapitre).ligneDebut + 1
        
        ' Parcourir vers le bas pour trouver la premiere ligne vide
        Do While dest.Range("B" & ligvide).Value <> ""
            ligvide = ligvide + 1
            ' Securite: limite a 500 lignes apres le debut du chapitre
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
    
    ' Cumuler temps si demande (Chapitres A et D) et si temps valide
    If cumulerTemps And temps > 0 Then
        chapitres(numChapitre).TotalTemps = chapitres(numChapitre).TotalTemps + temps
    End If
    
    chapitres(numChapitre).NombreInsertions = chapitres(numChapitre).NombreInsertions + 1
End Sub

' =========================================================================
' PROCEDURE: INSERER TACHE CHAPITRE C (avec sous-chapitres)
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
    
    ' Determiner ligne d'insertion
    If dest.Range("B" & (chapitres(numChapitre).LigneProchainChap - 1)).Value <> "" Then
        dest.Range("B" & chapitres(numChapitre).LigneProchainChap).EntireRow.Insert
        ligvide = chapitres(numChapitre).LigneProchainChap
        
        ' *** CORRECTION CRITIQUE ***
        ' Mettre a jour TOUS les chapitres suivants apres insertion
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
    
    ' Si premier item d'un sous-chapitre (rang x.1), creer l'en-tete du sous-chapitre
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
        
        ' Inserer ligne suivante pour la tache
        dest.Range("B" & (ligvide + 1)).EntireRow.Insert
        ligvide = ligvide + 1
        
        ' *** CORRECTION CRITIQUE ***
        ' Mettre a jour chapitres suivants apres 2eme insertion
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
    
    ' Remplir la ligne de tache
    With dest
        .Range("B" & ligvide) = Right(rang, 1)
        .Range("C" & ligvide) = premierMot
        .Range("D" & ligvide) = description
        
        ' Operateurs
        If langue = "FR" Then
            .Range("G" & ligvide) = "1 x ELECTRICIEN"
        Else
            .Range("G" & ligvide) = "1 x ELECTRICIAN"
        End If
        
        ' Pieces de rechange
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
' PROCEDURE: ECRIRE SOMMES TEMPS
' =========================================================================
Private Sub EcrireSommesTemps(ByRef dest As Worksheet, ByRef chapitres() As ChapitreInfo)
    Dim i As Integer
    
    ' ecrire le total de temps dans la colonne L de chaque ligne de chapitre
    For i = 1 To 4
        If chapitres(i).TotalTemps > 0 Then
            dest.Range("L" & chapitres(i).ligneDebut) = chapitres(i).TotalTemps
            dest.Range("L" & chapitres(i).ligneDebut).NumberFormat = "0.000"
        End If
    Next i
End Sub

' =========================================================================
' PROCEDURE: AJOUTER CHAPITRE E
' =========================================================================
Private Sub AjouterChapitreE(ByRef dest As Worksheet)
    Dim ligvide As Integer
    
    ' Trouver derniere ligne remplie
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
' PROCEDURE: REMPLIR COUVERTURE
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


