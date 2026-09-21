Attribute VB_Name = "C_EMMST_Main"
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

    On Error GoTo ErrHandler
    
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
    Exit Function
    
ErrHandler:
    MsgBox "Erreur lors de l'initialisation EMMST : " & Err.description, vbCritical
    InitialiserContexteEMMST = False
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
    
    On Error GoTo ErrHandler
    
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
    
    Exit Function
    
ErrHandler:
    MsgBox "Erreur lors de la copie du template EMMST : " & Err.description, vbCritical
    Set CopierTemplateEMMST = Nothing
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




