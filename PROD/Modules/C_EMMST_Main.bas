'--------------------------------------------------------------------------------------------------------
'
'                                       Generateur Livrables SLI
'                                       Module B_EMMST
'                                       Developpement : LE LARGE Thomas (LGM/Alstom)
'                                       Module permettant de generer des fiches EMMST
'
'
'-----------------------------------------------------------------------------------------------------------
Option Compare Text
Option Explicit

' =========================================================================
' DECLARATIONS MODULE : CACHE COMMENTAIRES EMMST
' =========================================================================
' Depuis 2026-11-09 - contournement de la troncature 255 chars d'ACE OLE DB
' sur les 10 colonnes commentaires (Emballage/Marquage/Manutention/Stockage/
' Transport, FR + EN). Cache lazy session Dictionary(DTR -> Array(10)),
' peuple au 1er appel de Main_EMMST, reutilise aux appels suivants.
' Cf. Docs/CHANGELOG.md session 2026-11-09 pour l'historique diagnostic.
' =========================================================================

' Cache session - Dictionary(DTR -> Array(0 To 9) Variant). Ordre FR puis EN :
'   0 Emballage FR, 1 Marquage FR, 2 Manutention FR, 3 Stockage FR,
'   4 Transport FR, 5 Emballage EN, 6 Marquage EN, 7 Manutention EN,
'   8 Stockage EN, 9 Transport EN.
Private m_dictCommentaires As Object

' Noms des 10 colonnes commentaires - ordre FR puis EN, coherent avec les
' indices 0-9 du cache (source BdD <-> injection EMMST_tab).
Private Const N_EMBALLAGE_FR As String = "CommentaireEmballage_EMMST"
Private Const N_MARQUAGE_FR As String = "CommentaireMarquage_EMMST"
Private Const N_MANUTENTION_FR As String = "CommentaireManutention_EMMST"
Private Const N_STOCKAGE_FR As String = "CommentaireStockage_EMMST"
Private Const N_TRANSPORT_FR As String = "CommentaireTransport_EMMST"
Private Const N_EMBALLAGE_EN As String = "CommentaireEmballageEN_EMMST"
Private Const N_MARQUAGE_EN As String = "CommentaireMarquageEN_EMMST"
Private Const N_MANUTENTION_EN As String = "CommentaireManutentionEN_EMMST"
Private Const N_STOCKAGE_EN As String = "CommentaireStockageEN_EMMST"
Private Const N_TRANSPORT_EN As String = "CommentaireTransportEN_EMMST"

' =========================================================================
' PROCEDURE PRINCIPALE : MAIN_EMMST
' =========================================================================
Public Sub Main_EMMST()
' Point d'entree principal pour la generation des EMMST
' Orchestre le workflow complet :
'   1. Validation des prerequis
'   2. Recuperation des donnees utilisateur
'   3. Recuperation des donnees PBS
'   4. Selection repertoire destination
'   5. Copie et ouverture template
'   6. Recuperation donnees BDD
'   7. Generation des fiches EMMST
'   8. Finalisation

    Dim chemin As String
    Dim Template As Workbook
    
    ' Desactiver fonctionnalites Excel pour performance
    TurnOffFunctionality
    
    ' Validation et recuperation donnees UserForm
    If Not InitialiserContexteEMMST Then
        TurnOnFunctionality
        Exit Sub
    End If
    
    ' Recuperation donnees PBS depuis DATA_ASL
    Call RecupererDonneesPBS_EMMST
    
    ' Selection repertoire destination
    If Not SelectRepertoire(chemin) Then
        TurnOnFunctionality
        Exit Sub
    End If
    
    ' Copier et ouvrir template
    Set Template = CopierTemplateEMMST(chemin)
    If Template Is Nothing Then
        TurnOnFunctionality
        Exit Sub
    End If
    
    ' Recuperation donnees BDD EMMST
    Call RecupererDonneesBDD_EMMST
    
    ' Cadrage tableau EMMST
    Call Cadrage_EMMST(EMMST_tab)
    
    ' Generer toutes les fiches EMMST
    Call GenererFichesEMMST(Template, chemin)
    
    ' Finalisation
    Call FinaliserFichierEMMST(Template)
    
    ' Message final
    If langue = "FR" Then
        MsgBox "EMMST creees", vbInformation
    Else
        MsgBox "PHLST created", vbInformation
    End If
    
    ' Reactiver fonctionnalites Excel
    TurnOnFunctionality

End Sub

' =========================================================================
' FONCTION: INITIALISER CONTEXTE EMMST
' =========================================================================
Private Function InitialiserContexteEMMST() As Boolean
' Recupere et valide les parametres du UserForm
' Retourne True si tout est OK, False sinon

    
    ' Definir classeur generateur
    Set Generateur = ActiveWorkbook
    Set DS1 = Generateur.Worksheets("DATA_ASL")
    Set REQSEL = CreateObject("ADODB.Recordset")
    
    ' Recuperer parametres systeme
    userpath = Environ$("UserProfile")
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Recuperer parametres projet
    Code_Projet = UserForm.Project_List.Value
    Projet_Abrev = Abrev_Projet(Code_Projet)
    Revision = UserForm.Revision_List.Value
    gate = UserForm.Gate_List.Value
    
    ' Validation gate
    If gate = "" Then
        MsgBox "Aucune gate n'est selectionnee. Merci de selectionner la gate avant de generer les fiches EMMST", _
               vbInformation, "Gate"
        InitialiserContexteEMMST = False
        Exit Function
    End If
    
    ' Determiner langue
    If UserForm.Langue_FR.Value = True Then
        langue = "FR"
    ElseIf UserForm.Langue_EN.Value = True Then
        langue = "EN"
    Else
        MsgBox "Veuillez selectionner une langue (FR ou EN)", vbExclamation
        InitialiserContexteEMMST = False
        Exit Function
    End If
    
    ' Validation basique
    If Code_Projet = "" Then
        MsgBox "Veuillez selectionner un projet", vbExclamation
        InitialiserContexteEMMST = False
        Exit Function
    End If
    
    If Revision = "" Then
        MsgBox "Veuillez saisir une revision", vbExclamation
        InitialiserContexteEMMST = False
        Exit Function
    End If
    
    InitialiserContexteEMMST = True
End Function

' =========================================================================
' PROCEDURE: RECUPeRER DONNEES PBS EMMST
' =========================================================================
Private Sub RecupererDonneesPBS_EMMST()
' Recupere les colonnes necessaires du tableau Providor pour EMMST
    
    tDTR = DS1.Range("Providor[DTR Reference (issue de PBS)]")
    tDES_PBS = DS1.Range("Providor[DESIGNATION (issue de PBS)]")
    tDES_MEF = DS1.Range("Providor[DESIGNATION (mise en forme pour copier dans Providor : construite à partir de celle de la PBS et de la désignation issue de BdD (ORCHESTRA)))]")
    tSS = DS1.Range("Providor[SOUS SYSTÈME (copiés de la PBS et classifiés selon génériques)]")
    
End Sub

' =========================================================================
' FONCTION: SELECTION REPERTOIRE (Reutilise fonction commune)
' =========================================================================
Private Function SelectRepertoire(ByRef chemin As String) As Boolean
' Dialogue selection repertoire destination
    
    Dim Repertoire As FileDialog
    
    MsgBox "Veuillez selectionner le repertoire ou vous souhaitez deposer les fiches EMMST pour le projet : " & Projet_Abrev, vbOKOnly
    
    Set Repertoire = Application.FileDialog(msoFileDialogFolderPicker)
    Repertoire.Show
    
    If Repertoire.SelectedItems.count = 0 Then
        SelectRepertoire = False
        Exit Function
    End If
    
    chemin = Repertoire.SelectedItems(1)
    SelectRepertoire = True
End Function

' =========================================================================
' FONCTION: COPIER TEMPLATE EMMST
' =========================================================================
Private Function CopierTemplateEMMST(ByVal chemin As String) As Workbook
' Copie le template EMMST selon la langue et l'ouvre
' Retourne le Workbook ouvert ou Nothing si erreur

    Dim nomFichier As String
    Dim templatePath As String
    
    
    ' Construire nom fichier et chemin template selon langue
    If langue = "FR" Then
        nomFichier = Projet_Abrev & "_EMMST_" & Revision & ".xlsx"
        templatePath = userpath & "\Alstom\DIS Service VBN Pole ISR - Qualité\005-03 Templates\02_Phase Projet\02. ASL\Fiches EMMST\Template EMMST-FR_V02.xlsx"
    Else
        nomFichier = Projet_Abrev & "_PHLST_" & Revision & ".xlsx"
        templatePath = userpath & "\Alstom\DIS Service VBN Pole ISR - Qualité\005-03 Templates\02_Phase Projet\02. ASL\Fiches EMMST\Template PHLST-ENG_V02.xlsx"
    End If
    
    ' Copier template
    Application.DisplayAlerts = False
    fso.CopyFile templatePath, chemin & "\" & nomFichier
    Application.DisplayAlerts = True
    
    ' Ouvrir fichier
    Set CopierTemplateEMMST = Workbooks.Open(chemin & "\" & nomFichier)
End Function

' =========================================================================
' PROCEDURE: RECUPeRER DONNEES BDD EMMST
' =========================================================================
Private Sub RecupererDonneesBDD_EMMST()
' Recupere les donnees EMMST et projet depuis la BDD
    
    Dim STRSEL As String
    
    ' =====================================================================
    ' REQUeTE 1: Table EMMST (48 champs)
    ' =====================================================================
    STRSEL = "SELECT [DTR_EMMST], [DimensionNuLongueur_EMMST], [DimensionNuLargeur_EMMST], [DimensionNuHauteur_EMMST], [DimensionNuPoids], " & _
             "[DimensionConditionneLongueur_EMMST], [DimensionConditionneLargeur_EMMST], [DimensionConditionneHauteur_EMMST], [DimensionConditionnePoids_EMMST], " & _
             "[DatePeremption_EMMST], [UniteConditionnement_EMMST], [CommentaireEmballage_EMMST], [CommentaireMarquage_EMMST], [CommentaireManutention_EMMST], " & _
             "[CommentaireStockage_EMMST], [CommentaireTransport_EMMST], [TypeOutil_EMMST], [DureePeremption_EMMST], [Reutilisable_EMMST], [EtiquetteEmballage_EMMST], " & _
             "[MarquageEquipement_EMMST], [ManutentionEquipement_EMMST], [OutillageSpecifiqueManutention_EMMST], [PrecautionManutention_EMMST], [PositionStockage_EMMST], " & _
             "[StockageExterieur_EMMST], [StockageSousAbri_EMMST], [StockageAirLibre_EMMST], [StockageHorsPoussiere_EMMST], [StockageHorsEau_EMMST], [StockageInterieur_EMMST], " & _
             "[StockageChaud_EMMST], [StockageFroid_EMMST], [StockageHorsLumiere_EMMST], [Gerbable_EMMST], [MoyenStockage_EMMST], [PrecautionStockage_EMMST], [MaintenanceStockage_EMMST], " & _
             "[MoyenTransport_EMMST], [OutillageSpecifiqueTransport_EMMST], [PrescriptionTransport_EMMST], [PrecautionTransport_EMMST], [Version_EMMST], [DatePeremptionEN_EMMST], " & _
             "[UniteConditionneEN_EMMST], [CommentaireEmballageEN_EMMST], [CommentaireMarquageEN_EMMST], [CommentaireManutentionEN_EMMST], [CommentaireStockageEN_EMMST], [CommentaireTransportEN_EMMST] " & _
             "FROM [EMMST$]"
    
    Call open_REQSEL(STRSEL)
    EMMST_tab = RecordsetToArray(REQSEL)

    ' Contourne la troncature ACE 255 chars sur les 10 colonnes commentaires
    ' (Emballage/Marquage/Manutention/Stockage/Transport FR+EN) : surcharge
    ' dans EMMST_tab les valeurs tronquees par ADO avec les valeurs completes
    ' lues en Excel natif ReadOnly invisible. Cache session Dictionary - cout
    ' 14 s au 1er appel, 0 s aux suivants. Cf. D_EMMST_CacheCommentaires +
    ' Docs/CHANGELOG.md session 2026-11-09.
    Call ChargerEtInjecterCommentairesEMMST

    ' =====================================================================
    ' REQUeTE 2: RASL Projet
    ' =====================================================================
    STRSEL = "SELECT [RASL_PROJET SLI] FROM [Projet_SLI$] WHERE [Nom projet_PROJET SLI] = '" & Code_Projet & "'"
    Call open_REQSEL(STRSEL)
    
    If IsNull(REQSEL.Fields(0).Value) Then
        Auteur = ""
    Else
        Auteur = REQSEL.Fields(0).Value
    End If
    
    ' =====================================================================
    ' REQUeTE 3: Niveau Maintenance
    ' =====================================================================
    STRSEL = "SELECT [Niv de maintenance_PROJET SLI] FROM [Projet_SLI$] WHERE [Nom projet_PROJET SLI] = '" & Code_Projet & "'"
    Call open_REQSEL(STRSEL)
    
    If IsNull(REQSEL.Fields(0).Value) Then
        Nv_Maint_Projet = ""
    Else
        Nv_Maint_Projet = REQSEL.Fields(0).Value
    End If
    
End Sub

' =========================================================================
' PROCEDURE: GENERER FICHES EMMST
' =========================================================================
Private Sub GenererFichesEMMST(ByRef Template As Workbook, ByVal chemin As String)
' Genere toutes les fiches EMMST pour chaque DTR
    
    Dim index As Worksheet
    Dim CoverPage As Worksheet
    Dim Template_EMMST As Worksheet
    Dim TabIndex As ListObject
    Dim i As Integer
    
    ' Identifier feuilles selon langue
    If langue = "FR" Then
        Set index = Template.Worksheets("Sommaire")
        Set CoverPage = Template.Worksheets("Couverture")
        Set Template_EMMST = Template.Worksheets("EMMST_VBN_FR")
    Else
        Set index = Template.Worksheets("Index")
        Set CoverPage = Template.Worksheets("CoverPage")
        Set Template_EMMST = Template.Worksheets("PHLST")
    End If
    
    ' Activer template
    Template.Activate
    
    ' Remplir Index
    Call RemplirIndexEMMST(index, Template)
    
    ' Boucle sur chaque DTR
    For i = 1 To UBound(tDTR)
        
        ' Skip supports/frames
        If InStr(1, tSS(i, 1), "support") > 0 Or InStr(1, tSS(i, 1), "frame") > 0 Then
            index.Range("F" & i + 5) = "NA"
            GoTo NextDTR
        End If
        
        ' Generer fiche EMMST pour ce DTR
        Call GenererFicheEMMST(Template, index, Template_EMMST, i)
        
NextDTR:
        ' Barre de progression
        Call barprogress(i, UBound(tDTR))
    Next i
    
    ' Finaliser
    Call RemplirCouvertureEMMST(CoverPage)
    
End Sub

' =========================================================================
' PROCEDURE: REMPLIR INDEX EMMST
' =========================================================================
Private Sub RemplirIndexEMMST(ByRef index As Worksheet, ByRef Template As Workbook)
' Remplit le tableau Index avec les donnees PBS
    
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
    End If
    
End Sub

' =========================================================================
' PROCEDURE: GENERER FICHE EMMST (pour un DTR)
' =========================================================================
Private Sub GenererFicheEMMST(ByRef Template As Workbook, ByRef index As Worksheet, _
                              ByRef Template_EMMST As Worksheet, ByVal indexDTR As Integer)
' Genere une fiche EMMST pour un DTR specifique
    
    Dim dEst As Worksheet
    Dim system As String
    Dim cpt As Integer
    Dim ligne As Integer
    Dim test As Boolean
    
    system = tSS(indexDTR, 1)
    cpt = 0
    
    ' Verifier doublons noms feuilles
    cpt = CompterDoublonsFeuillesEMMST(Template, index, indexDTR, system)
    
    ' Copier template
    Template_EMMST.Copy After:=Template.Worksheets(Template.Worksheets.count)
    Set dEst = ActiveSheet
    
    ' Nommer la feuille
    Call NommerFeuilleEMMST(dEst, index, indexDTR, system, cpt)
    
    ' Rechercher DTR dans BDD
    If Not RechercherDTR_EMMST(tDTR(indexDTR, 1), ligne, test) Then
        ' DTR non trouve
        Call MarquerStatutIndexEMMST(index, indexDTR, "A creer")
        Call RemplirInfosMinimalesEMMST(dEst, index, indexDTR)
        GoTo AjouterLien
    End If
    
    ' Verifier completude donnees
    Call VerifierCompletudeEMMST(ligne, index, indexDTR)
    
    ' Remplir fiche EMMST complete
    Call RemplirFicheEMMST(dEst, index, indexDTR, ligne)
    
AjouterLien:
    ' Creer lien hypertexte
    Call CreerLienHypertexteEMMST(index, indexDTR, dEst)
    
End Sub

' =========================================================================
' FONCTION: COMPTER DOUBLONS FEUILLES EMMST
' =========================================================================
Private Function CompterDoublonsFeuillesEMMST(ByRef Template As Workbook, ByRef index As Worksheet, _
                                               ByVal indexDTR As Integer, ByVal system As String) As Integer
' Compte le nombre de feuilles avec nom similaire (gestion doublons)
    
    Dim ws As Worksheet
    Dim system_temp As String
    Dim cpt As Integer
    
    cpt = 0
    For Each ws In Template.Worksheets
        system_temp = ws.Range("D" & indexDTR + 5)
        If system_temp = "" Then system_temp = system
        
        If Trim(Left(ws.Name, 20 + Len(system_temp))) = index.Range("B" & indexDTR + 5) & "_" & Trim(Left(system, 4)) Then
            cpt = cpt + 1
        End If
    Next ws
    
    CompterDoublonsFeuillesEMMST = cpt
End Function

' =========================================================================
' PROCEDURE: NOMMER FEUILLE EMMST
' =========================================================================
Private Sub NommerFeuilleEMMST(ByRef dEst As Worksheet, ByRef index As Worksheet, _
                               ByVal indexDTR As Integer, ByVal system As String, ByVal cpt As Integer)
' Nomme la feuille EMMST en gerant les doublons
    
    Dim baseName As String
    Dim k As Integer
    Dim estnum As String
    
    baseName = index.Range("B" & indexDTR + 5) & "_" & Trim(Left(system, 4))
    
    If cpt = 0 Then
        ' Pas de doublon
        dEst.Name = baseName
    Else
        ' Il y a des doublons
        For k = Worksheets.count To 1 Step -1
            If InStr(1, Worksheets(k).Name, baseName, vbTextCompare) = 1 Then
                estnum = Mid(Worksheets(k).Name, Len(Worksheets(k).Name) - 1, 1)
                If IsNumeric(estnum) Then
                    dEst.Name = baseName & " (" & estnum + 1 & ")"
                    Exit For
                Else
                    dEst.Name = baseName & " (" & cpt + 1 & ")"
                    Exit For
                End If
            End If
        Next k
    End If
    
End Sub

' =========================================================================
' FONCTION: RECHERCHER DTR EMMST
' =========================================================================
Private Function RechercherDTR_EMMST(ByVal dtr As String, ByRef ligne As Integer, _
                                     ByRef test As Boolean) As Boolean
' Recherche un DTR dans le tableau EMMST_tab
' Retourne True si trouve, False sinon
    
    Dim j As Integer
    
    test = False
    ligne = 0
    
    For j = 2 To UBound(EMMST_tab)
        If EMMST_tab(j, EMMST_type.dtr) = dtr Then
            ligne = j
            test = True
            Exit For
        End If
    Next j
    
    RechercherDTR_EMMST = test
End Function

' =========================================================================
' PROCEDURE: VERIFIER COMPLeTUDE EMMST
' =========================================================================
Private Sub VerifierCompletudeEMMST(ByVal ligne As Integer, ByRef index As Worksheet, _
                                    ByVal indexDTR As Integer)
' Verifie si la fiche EMMST a beaucoup de donnees manquantes
' Si plus de 20 champs NULL ? marquer "A remplir"
    
    Dim h As Integer
    Dim incr As Integer
    
    incr = 0
    For h = 2 To UBound(EMMST_tab, 2)
        If IsNull(EMMST_tab(ligne, h)) Then
            incr = incr + 1
        End If
    Next h
    
    ' Si plus de 20 champs vides ? "A remplir"
    If incr > 20 Then
        Call MarquerStatutIndexEMMST(index, indexDTR, "A remplir")
    End If
    
End Sub

' =========================================================================
' PROCEDURE: REMPLIR FICHE EMMST COMPLeTE
' =========================================================================
Private Sub RemplirFicheEMMST(ByRef dEst As Worksheet, ByRef index As Worksheet, _
                              ByVal indexDTR As Integer, ByVal ligne As Integer)
' Remplit une fiche EMMST complete avec toutes les donnees
' Utilise la classe EMMST_Ranges pour acceder aux cellules
    
    Dim r As New EMMST_Ranges
    Dim STRSEL As String
    
    ' Initialiser les ranges de destination
    Call r.Init(dEst)
    
    ' Recuperer fournisseur depuis table Produit
    STRSEL = "SELECT [Fournisseur_PRODUIT], [Reference fournisseur_PRODUIT] FROM [Produit$] WHERE [DTR] = '" & tDTR(indexDTR, 1) & "'"
    Call open_REQSEL(STRSEL)
    
    If IsNull(REQSEL.Fields(0).Value) Then
        r.supplier_name.Value = ""
    Else
        r.supplier_name.Value = REQSEL.Fields(0).Value
    End If
    
    If IsNull(REQSEL.Fields(1).Value) Then
        r.supplier_id.Value = ""
    Else
        r.supplier_id.Value = REQSEL.Fields(1).Value
    End If
    
    ' Remplir tous les champs (appel fonction dediee)
    Call RemplirChampsGeneraux(r, index, indexDTR, ligne)
    Call RemplirChampsDimensions(r, ligne)
    Call RemplirChampsEmballage(r, ligne)
    Call RemplirChampsMarquage(r, ligne)
    Call RemplirChampsManutention(r, ligne)
    Call RemplirChampsStockage(r, ligne)
    Call RemplirChampsTransport(r, ligne)
    Call RemplirCommentaires(r, ligne)

End Sub


' =========================================================================
' SECTION : CACHE COMMENTAIRES EMMST (depuis 2026-11-09)
' =========================================================================
' Procedures de contournement de la troncature 255 chars d'ACE OLE DB sur
' les 10 colonnes commentaires EMMST. Appele depuis RecupererDonneesBDD_EMMST
' juste apres EMMST_tab = RecordsetToArray(REQSEL). Cf. Docs/CHANGELOG.md
' session 2026-11-09 pour l'historique diagnostic (Piste 1 ADO KO, bench
' v1 scenarios A/B, bench v2 A1/A2/A3, Plan B1 adopte).
' =========================================================================


' =========================================================================
' PROCEDURE: ChargerEtInjecterCommentairesEMMST
' =========================================================================
' Point d'entree unique appele depuis RecupererDonneesBDD_EMMST juste apres
' EMMST_tab = RecordsetToArray(REQSEL).
'
' Au 1er appel de la session : charge le cache (14 s typiques).
' Aux appels suivants : reutilise le Dictionary memoire (0 s).
'
' Ecrase dans EMMST_tab les 10 colonnes commentaires tronquees par ADO avec
' les valeurs completes. Les DTR absents du cache (nouveaux ou BdD desync)
' sont laisses intacts (valeur tronquee ADO conservee - degradation gracieuse).
Public Sub ChargerEtInjecterCommentairesEMMST()

    Dim colDTR As Long
    Dim colsBdD As Variant
    Dim colsTab() As Long
    Dim iCol As Long
    Dim iLigne As Long
    Dim dtr As String
    Dim arrComm As Variant

    If m_dictCommentaires Is Nothing Then
        Call ChargerCacheCommentaires
    End If

    colsBdD = Array(N_EMBALLAGE_FR, N_MARQUAGE_FR, N_MANUTENTION_FR, _
                    N_STOCKAGE_FR, N_TRANSPORT_FR, _
                    N_EMBALLAGE_EN, N_MARQUAGE_EN, N_MANUTENTION_EN, _
                    N_STOCKAGE_EN, N_TRANSPORT_EN)

    colDTR = Arraypos(EMMST_tab, "DTR_EMMST")
    If colDTR = -1 Then
        Err.Raise vbObjectError + 540, "ChargerEtInjecterCommentairesEMMST", _
                  "Colonne DTR_EMMST introuvable dans EMMST_tab."
    End If

    ReDim colsTab(LBound(colsBdD) To UBound(colsBdD))
    For iCol = LBound(colsBdD) To UBound(colsBdD)
        colsTab(iCol) = Arraypos(EMMST_tab, CStr(colsBdD(iCol)))
        If colsTab(iCol) = -1 Then
            Err.Raise vbObjectError + 541, "ChargerEtInjecterCommentairesEMMST", _
                      "Colonne introuvable dans EMMST_tab : " & colsBdD(iCol)
        End If
    Next iCol

    ' Parcours des lignes de donnees (ligne 0 = en-tetes dans RecordsetToArray)
    For iLigne = 1 To UBound(EMMST_tab, 1)
        dtr = CStr(EMMST_tab(iLigne, colDTR))
        If m_dictCommentaires.Exists(dtr) Then
            arrComm = m_dictCommentaires(dtr)
            For iCol = LBound(colsBdD) To UBound(colsBdD)
                EMMST_tab(iLigne, colsTab(iCol)) = arrComm(iCol)
            Next iCol
        End If
    Next iLigne

End Sub


' =========================================================================
' PROCEDURE: ChargerCacheCommentaires
' =========================================================================
' 1er appel uniquement : copie BdD TRAVAIL vers %TEMP%, ouvre en Excel
' invisible ReadOnly, lit les 10 colonnes commentaires + DTR via Range.Value,
' ferme sans sauvegarder, supprime la copie. 5 etapes via F_Progression.
'
' Decoupage de cout (bench v2 scenario A3 sur poste Alstom, 998 DTR) :
'   - FileCopy local       : 0.07 s  (elimine latence OneDrive sur close)
'   - Workbooks.Open       : 12.86 s (incompressible, parsing .xlsb)
'   - Range.Value x 10     : 0.02 s  (lecture bloc, 10 colonnes)
'   - wb.Close             : 1.14 s
'   - Restauration + Kill  : 0.22 s
'   - TOTAL                : 14.40 s
'
' Application.ScreenUpdating / EnableEvents / Calculation / DisplayAlerts
' sont deja positionnes par TurnOffFunctionality appele dans Main_EMMST
' avant RecupererDonneesBDD_EMMST. On ne touche pas a l'etat Application
' ici (restauration sous controle de TurnOnFunctionality en fin de Main).
Private Sub ChargerCacheCommentaires()

    Dim cheminBdD As String
    Dim cheminTravail As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim colDTR As Long
    Dim colsBdD As Variant
    Dim colsIdx() As Long
    Dim iCol As Long
    Dim derniereLigne As Long
    Dim dtrs As Variant
    Dim tabCom0 As Variant, tabCom1 As Variant, tabCom2 As Variant
    Dim tabCom3 As Variant, tabCom4 As Variant, tabCom5 As Variant
    Dim tabCom6 As Variant, tabCom7 As Variant, tabCom8 As Variant
    Dim tabCom9 As Variant
    Dim tabsParCol As Variant
    Dim iLigne As Long
    Dim dtr As String
    Dim arrComm(0 To 9) As Variant

    Set m_dictCommentaires = CreateObject("Scripting.Dictionary")
    m_dictCommentaires.CompareMode = vbTextCompare

    colsBdD = Array(N_EMBALLAGE_FR, N_MARQUAGE_FR, N_MANUTENTION_FR, _
                    N_STOCKAGE_FR, N_TRANSPORT_FR, _
                    N_EMBALLAGE_EN, N_MARQUAGE_EN, N_MANUTENTION_EN, _
                    N_STOCKAGE_EN, N_TRANSPORT_EN)

    Call Progress_Demarrer(5, "Chargement commentaires EMMST")

    ' --- Etape 1 : copie locale %TEMP% -----------------------------------
    Call Progress_Etape(1, "Copie BdD en local", "")
    cheminBdD = ExtraireDataSourceDepuisConnStr( _
        Replace(CONNEXION_EXCEL_BDD_TRAVAIL, "UserProfile", Environ$("UserProfile")))
    cheminTravail = Environ$("TEMP") & "\ECOSLI_bdd_cache_" & _
                    Format(Now, "yyyymmdd_hhnnss") & ".xlsb"
    FileCopy cheminBdD, cheminTravail

    ' --- Etape 2 : ouverture BdD invisible ReadOnly ----------------------
    Call Progress_Etape(2, "Ouverture BdD invisible", "")
    Set wb = Workbooks.Open(Filename:=cheminTravail, UpdateLinks:=0, _
                             ReadOnly:=True, Notify:=False)
    Set ws = wb.Worksheets("EMMST")

    colDTR = TrouverColonneBdD_EMMST(ws, "DTR_EMMST")
    ReDim colsIdx(LBound(colsBdD) To UBound(colsBdD))
    For iCol = LBound(colsBdD) To UBound(colsBdD)
        colsIdx(iCol) = TrouverColonneBdD_EMMST(ws, CStr(colsBdD(iCol)))
    Next iCol

    ' --- Etape 3 : lecture bloc Range.Value (10 colonnes + DTR) ----------
    Call Progress_Etape(3, "Lecture commentaires", "")
    With ws
        derniereLigne = .Cells(.Rows.count, colDTR).End(xlUp).Row
        dtrs = .Range(.Cells(2, colDTR), .Cells(derniereLigne, colDTR)).Value

        tabCom0 = .Range(.Cells(2, colsIdx(0)), .Cells(derniereLigne, colsIdx(0))).Value
        tabCom1 = .Range(.Cells(2, colsIdx(1)), .Cells(derniereLigne, colsIdx(1))).Value
        tabCom2 = .Range(.Cells(2, colsIdx(2)), .Cells(derniereLigne, colsIdx(2))).Value
        tabCom3 = .Range(.Cells(2, colsIdx(3)), .Cells(derniereLigne, colsIdx(3))).Value
        tabCom4 = .Range(.Cells(2, colsIdx(4)), .Cells(derniereLigne, colsIdx(4))).Value
        tabCom5 = .Range(.Cells(2, colsIdx(5)), .Cells(derniereLigne, colsIdx(5))).Value
        tabCom6 = .Range(.Cells(2, colsIdx(6)), .Cells(derniereLigne, colsIdx(6))).Value
        tabCom7 = .Range(.Cells(2, colsIdx(7)), .Cells(derniereLigne, colsIdx(7))).Value
        tabCom8 = .Range(.Cells(2, colsIdx(8)), .Cells(derniereLigne, colsIdx(8))).Value
        tabCom9 = .Range(.Cells(2, colsIdx(9)), .Cells(derniereLigne, colsIdx(9))).Value
    End With
    tabsParCol = Array(tabCom0, tabCom1, tabCom2, tabCom3, tabCom4, _
                       tabCom5, tabCom6, tabCom7, tabCom8, tabCom9)

    ' --- Etape 4 : indexation Dictionary ---------------------------------
    Call Progress_Etape(4, "Indexation DTR", "")
    For iLigne = 1 To UBound(dtrs, 1)
        dtr = CStr(dtrs(iLigne, 1))
        If LenB(dtr) > 0 Then
            For iCol = 0 To 9
                arrComm(iCol) = tabsParCol(iCol)(iLigne, 1)
            Next iCol
            ' Assignation d'un Array a un Dictionary = copie par valeur en
            ' VBA. La modification de arrComm a l'iteration suivante ne
            ' modifie pas la valeur deja stockee dans le Dictionary.
            m_dictCommentaires(dtr) = arrComm
        End If
    Next iLigne

    ' --- Etape 5 : fermeture propre + Kill copie -------------------------
    Call Progress_Etape(5, "Fermeture BdD", "")
    wb.Close SaveChanges:=False
    Set wb = Nothing
    Set ws = Nothing
    Kill cheminTravail

    Call Progress_Terminer

End Sub


' =========================================================================
' FONCTION: TrouverColonneBdD_EMMST
' =========================================================================
' Cherche nomColonne en ligne 1 de ws et retourne l'index de colonne.
' Pattern Find convention 2 : xlValues + xlWhole, fallback xlPart.
' Suffixe _EMMST pour eviter toute collision Private-to-module.
Private Function TrouverColonneBdD_EMMST(ByRef ws As Worksheet, _
                                          ByVal nomColonne As String) As Long
    Dim trouve As Range

    Set trouve = ws.Rows(1).Find(What:=nomColonne, LookIn:=xlValues, _
                                  LookAt:=xlWhole, MatchCase:=False)
    If trouve Is Nothing Then
        Set trouve = ws.Rows(1).Find(What:=nomColonne, LookIn:=xlValues, _
                                      LookAt:=xlPart, MatchCase:=False)
    End If

    If trouve Is Nothing Then
        Err.Raise vbObjectError + 542, "TrouverColonneBdD_EMMST", _
                  "Colonne introuvable dans la feuille EMMST : " & nomColonne
    End If

    TrouverColonneBdD_EMMST = trouve.Column
End Function


' =========================================================================
' FONCTION: ExtraireDataSourceDepuisConnStr
' =========================================================================
' Extrait la valeur de Data Source= dans une chaine de connexion ADO.
' Forme attendue : Provider=...;Data Source=<chemin>;Extended Properties='...';
Private Function ExtraireDataSourceDepuisConnStr(ByVal connStr As String) As String
    Dim posStart As Long
    Dim posFin As Long
    Dim debut As Long

    posStart = InStr(1, connStr, "Data Source=", vbTextCompare)
    If posStart = 0 Then
        Err.Raise vbObjectError + 543, "ExtraireDataSourceDepuisConnStr", _
                  "Chaine sans 'Data Source='."
    End If

    debut = posStart + Len("Data Source=")
    posFin = InStr(debut, connStr, ";")
    If posFin = 0 Then posFin = Len(connStr) + 1

    ExtraireDataSourceDepuisConnStr = Mid$(connStr, debut, posFin - debut)
End Function

