' Cache du template local : copie unique depuis SharePoint en debut d'ATM_V2.
' Toutes les copies par DTR partent ensuite de cette copie locale (disque
' local ~10x plus rapide que OneDrive synchronise).
Private templateLocalV2 As String

' Cache des positions des marqueurs de chapitre dans le template.
' Le template etant constant, les 4 Find sur "A"/"B"/"C"/"D" donnent toujours
' les memes lignes. On les calcule UNE seule fois au 1er DTR genere puis on
' reutilise les valeurs pour tous les DTR suivants (gain : ~8 Find par DTR).
Private cacheChapInitialise As Boolean
Private cacheChapLigneDebut(1 To 4) As Long
Private cacheChapLigneProchain(1 To 4) As Long

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

    ' Reset cache positions chapitres (recalcul au 1er DTR genere de cette run)
    cacheChapInitialise = False

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

    'Cadrage des deux onglets importants de la BDD : initialise
    ' Taches_type / Info_type (numeros de colonnes).
    Call Cadrage_Taches_ATM(Taches_tab)
    Call Cadrage_Infos_ATM(Infos_tab)

    ' Indexation DTR -> ligne en O(1). Doit etre APRES les Cadrage_*
    ' sinon Taches_type.dtr / Info_type.dtr valent 0 et l'index est cassé.
    Call IndexerTachesEtInfos
    
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
    Close_CONN (Replace(CONNEXION_EXCEL_BDD_TRAVAIL, "UserProfile", userpath))

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
Private Sub CopierTemplateVide(ByVal dtr As String, ByVal designation As String, ByVal chemin As String)
    Dim nomFichier As String

    ' Construire nom fichier
    nomFichier = "(vide)" & dtr & "_" & langue & "_" & Replace(Left(designation, 10), "/", "_") & ".xlsx"

    ' Copier depuis le cache local (templateLocalV2 initialise dans ATM_V2).
    fso.CopyFile templateLocalV2, chemin & "\" & nomFichier
End Sub
' =========================================================================
' PROCEDURE PRINCIPALE: GENERER FICHE ATM
' =========================================================================
Private Sub GenererFicheATM(ByVal dtr As String, ByVal designation As String, _
                            ByVal chemin As String, ByVal indexDTR As Integer)
    Dim nomFichier As String
    Dim dEst As Worksheet
    Dim CoverPage As Worksheet
    Dim PMB As Integer, PMH As Integer, PMJ As Integer
    Dim test_taches As Boolean, test_infos As Boolean

    ' Construire nom fichier
    nomFichier = "(non illustre)" & dtr & "_" & langue & "_" & Replace(Left(designation, 10), "/", "_") & ".xlsx"

    ' Copier depuis le cache local (10x plus rapide que SharePoint) puis ouvrir.
    fso.CopyFile templateLocalV2, chemin & "\" & nomFichier
    Workbooks.Open chemin & "\" & nomFichier

    ' Workbooks.Open peut restaurer Calculation au mode du classeur ouvert
    ' (Excel aligne le mode sur le 1er classeur). On force manuel ici pour
    ' eviter le recalcul a chaque Insert/ecriture de cellule pendant la
    ' generation. TurnOnFunctionality (fin de ATM_V2) restaure xlCalculationAutomatic.
    Application.Calculation = xlCalculationManual

    ' Identifier feuilles
    If langue = "FR" Then
        Set dEst = ActiveWorkbook.Sheets("DESCRIPTION_TACHE")
    Else
        Set dEst = ActiveWorkbook.Sheets("TASK_DESCRIPTION")
    End If
    Set CoverPage = ActiveWorkbook.Sheets(1)
    
    ' Rechercher plages DTR dans Taches_tab et Infos_tab
    If Not RechercherPlagesDTR_V2(dtr, PMB, PMH, PMJ, test_taches, test_infos) Then
        ' DTR connu mais pas dans la bonne langue ? Sauvegarder comme vide
        ActiveWorkbook.SaveAs chemin & "\(vide)" & dtr & "_" & langue & "_" & Left(designation, 10) & ".xlsx"
        ActiveWorkbook.Close
        Exit Sub
    End If
    
    ' Activer feuille destination
    dEst.Activate
    
    ' REMPLISSAGE PAR CHAPITRES (VERSION SIMPLIFIEE)
    Call RemplirTousChapitres(dEst, PMB, PMH)
    
    ' Ajouter Chapitre E (Notes integrateur)
    Call AjouterChapitreE(dEst)
    
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
Dim cle As String

    test_taches = False
    test_infos = False
    PMB = 0
    PMH = 0
    PMJ = 0

    ' --- PMB : tentative dico O(1) --------------------------
    cle = dtr & "|" & langue
    If Not dictTachesPMB Is Nothing Then
        If dictTachesPMB.Exists(cle) Then
            PMB = dictTachesPMB(cle)
            test_taches = True
        End If
    End If

    ' --- PMB : fallback lineaire (filet de securite) --------
    If Not test_taches Then
        For j = 2 To UBound(Taches_tab, 1)
            If Taches_tab(j, Taches_type.dtr) = dtr And _
               Taches_tab(j, Taches_type.langue) = langue Then
                PMB = j
                test_taches = True
                Debug.Print "RechercherPlagesDTR_V2 [" & dtr & "] : dico miss, fallback lineaire OK PMB=" & PMB
                Exit For
            End If
        Next j
    End If

    ' Si langue non trouvee, retourner False
    If Not test_taches Then
        RechercherPlagesDTR_V2 = False
        Exit Function
    End If

    ' --- PMH : continuation depuis PMB ----------------------
    ' V2 verifie AUSSI la langue dans la continuation (plus strict que V1)
    j = PMB
    Do While j < UBound(Taches_tab, 1)
        If Taches_tab(j + 1, Taches_type.dtr) = Taches_tab(j, Taches_type.dtr) And _
           Taches_tab(j + 1, Taches_type.langue) = langue Then
            j = j + 1
        Else
            Exit Do
        End If
    Loop
    PMH = j

    ' --- PMJ : tentative dico (V2 a besoin de la ligne, pas juste un bool) ----
    If Not dictInfosPMJ Is Nothing Then
        If dictInfosPMJ.Exists(dtr) Then
            PMJ = dictInfosPMJ(dtr)
            test_infos = True
        End If
    End If

    ' --- PMJ : fallback lineaire ----------------------------
    If Not test_infos Then
        For l = 2 To UBound(Infos_tab, 1)
            If Infos_tab(l, Info_type.dtr) = dtr Then
                PMJ = l
                test_infos = True
                Exit For
            End If
        Next l
    End If

    RechercherPlagesDTR_V2 = True
End Function
' =========================================================================
' PROCEDURE PRINCIPALE: REMPLIR TOUS CHAPITRES (VERSION SIMPLIFIEE)
' =========================================================================
Private Sub RemplirTousChapitres(ByRef dEst As Worksheet, ByVal PMB As Integer, ByVal PMH As Integer)
    Dim j As Integer
    Dim rang As String
    Dim rangPremierChiffre As String
    Dim compteurSousChapitres As Integer
    Dim chapitres(1 To 4) As ChapitreInfo
    
    ' Initialiser structure chapitres
    Call InitialiserChapitres(dEst, chapitres)
    
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
                Call InsererTacheChapitre(dEst, chapitres, 1, j, True)  ' Chapitre A
                
            Case "2"
                Call InsererTacheChapitre(dEst, chapitres, 2, j, True) ' Chapitre B
                
            Case "3", "4"
                Call InsererTacheChapitreC(dEst, chapitres, 3, j, compteurSousChapitres) ' Chapitre C
                
            Case "5"
                Call InsererTacheChapitre(dEst, chapitres, 4, j, True)  ' Chapitre D
        End Select
        
NextTache:
    Next j
    
    ' ecrire les sommes de temps dans chaque chapitre
    Call EcrireSommesTemps(dEst, chapitres)
End Sub
' =========================================================================
' PROCEDURE: INITIALISER CHAPITRES
' =========================================================================
' Utilise le cache cacheChapLigneDebut / cacheChapLigneProchain alimente au
' premier DTR genere (8 Find sur le template effectues une seule fois pour
' toute la run). Pour les DTR suivants : remplissage direct sans Find.
Private Sub InitialiserChapitres(ByRef dEst As Worksheet, ByRef chapitres() As ChapitreInfo)
    Dim i As Integer
    Dim marqueurs As Variant

    If Not cacheChapInitialise Then
        ' Premier DTR de la run : 8 Find sur le template (le seul moment)
        cacheChapLigneDebut(1) = dEst.Range("B:B").Find(What:="A", LookIn:=xlValues, lookat:=xlWhole).row
        cacheChapLigneProchain(1) = dEst.Range("B:B").Find(What:="B", LookIn:=xlValues, lookat:=xlWhole).row
        cacheChapLigneDebut(2) = cacheChapLigneProchain(1)
        cacheChapLigneProchain(2) = dEst.Range("B:B").Find(What:="C", LookIn:=xlValues, lookat:=xlWhole).row
        cacheChapLigneDebut(3) = cacheChapLigneProchain(2)
        cacheChapLigneProchain(3) = dEst.Range("B:B").Find(What:="D", LookIn:=xlValues, lookat:=xlWhole).row
        cacheChapLigneDebut(4) = cacheChapLigneProchain(3)
        cacheChapLigneProchain(4) = -1 ' Dernier chapitre
        cacheChapInitialise = True
    End If

    marqueurs = Array("", "A", "B", "C", "D")  ' index 1..4
    For i = 1 To 4
        chapitres(i).Marqueur = marqueurs(i)
        chapitres(i).ligneDebut = cacheChapLigneDebut(i)
        chapitres(i).LigneProchainChap = cacheChapLigneProchain(i)
        chapitres(i).TotalTemps = 0
        chapitres(i).NombreInsertions = 0
    Next i
End Sub

' =========================================================================
' PROCEDURE: INSERER TACHE CHAPITRE (A, B, D)
' =========================================================================
Private Sub InsererTacheChapitre(ByRef dEst As Worksheet, ByRef chapitres() As ChapitreInfo, _
                                 ByVal numChapitre As Integer, ByVal indexTache As Integer, _
                                 ByVal cumulerTemps As Boolean)
    Dim ligvide As Long  ' IMPORTANT: Long au lieu de Integer pour eviter overflow
    Dim temps As Double
    Dim k As Integer
    
    ' Determiner ligne d'insertion
    If chapitres(numChapitre).LigneProchainChap > 0 Then
        ' Cas des Chapitres A, B, C (ont un chapitre suivant)
        If dEst.Range("B" & (chapitres(numChapitre).LigneProchainChap - 1)).Value <> "" Then
            ' Inserer ligne avant le prochain chapitre
            dEst.Range("B" & chapitres(numChapitre).LigneProchainChap).EntireRow.Insert
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
        ' Chapitre D (dernier chapitre) : premiere ligne vide apres la
        ' derniere ligne remplie de la colonne B. End(xlUp) depuis la
        ' derniere ligne de la feuille est O(1) cote Excel, contrairement
        ' a une boucle Do While qui lit cellule par cellule.
        ' Borne basse : ligneDebut (marqueur D), donc +1 minimum.
        ligvide = dEst.Cells(dEst.rows.count, "B").End(xlUp).row + 1
        If ligvide < chapitres(numChapitre).ligneDebut + 1 Then
            ligvide = chapitres(numChapitre).ligneDebut + 1
        End If
    End If
    
    ' Remplir la ligne
    With dEst
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
Private Sub InsererTacheChapitreC(ByRef dEst As Worksheet, ByRef chapitres() As ChapitreInfo, _
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
    If dEst.Range("B" & (chapitres(numChapitre).LigneProchainChap - 1)).Value <> "" Then
        dEst.Range("B" & chapitres(numChapitre).LigneProchainChap).EntireRow.Insert
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
        With dEst
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
        dEst.Range("B" & (ligvide + 1)).EntireRow.Insert
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
    With dEst
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
Private Sub EcrireSommesTemps(ByRef dEst As Worksheet, ByRef chapitres() As ChapitreInfo)
    Dim i As Integer
    
    ' ecrire le total de temps dans la colonne L de chaque ligne de chapitre
    For i = 1 To 4
        If chapitres(i).TotalTemps > 0 Then
            dEst.Range("L" & chapitres(i).ligneDebut) = chapitres(i).TotalTemps
            dEst.Range("L" & chapitres(i).ligneDebut).NumberFormat = "0.000"
        End If
    Next i
End Sub

' =========================================================================
' PROCEDURE: AJOUTER CHAPITRE E
' =========================================================================
Private Sub AjouterChapitreE(ByRef dEst As Worksheet)
    Dim ligvide As Integer
    
    ' Trouver derniere ligne remplie
    ligvide = dEst.Range("B10000").End(xlUp).row + 1
    
    ' Copier format de la ligne des totaux (ligne 2)
    dEst.Range("B2:L2").Copy dEst.Range("B" & ligvide)
    
    ' Marqueur Chapitre E
    dEst.Range("B" & ligvide) = "E"
    
    ' Ligne Note
    ligvide = dEst.Range("B10000").End(xlUp).row + 1
    dEst.Range("B" & ligvide) = 1
    dEst.Range("C" & ligvide) = "Note"
    
    If langue = "FR" Then
        dEst.Range("D" & ligvide) = "A definir par integrateur"
    Else
        dEst.Range("D" & ligvide) = "To be defined by integrator"
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

