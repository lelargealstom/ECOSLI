Attribute VB_Name = "D_Providor_Main"
'--------------------------------------------------------------------------------------------------------
'
'                                       Generateur Livrables SLI
'                                       Module D_Providor_Main (Version Refactorisée)
'                                       Developpement : LE LARGE Thomas (LGM/Alstom)
'                                       Module permettant de générer un Providor LCC
'                                       Refactorisation : Février 2026
'
'-----------------------------------------------------------------------------------------------------------
Option Compare Text
Option Explicit

' =========================================================================
' PROCÉDURE PRINCIPALE : MAIN_PROVIDOR
' =========================================================================
Public Sub Main_Providor(test As Boolean)
' Point d'entrée principal pour la génération du Providor
' Orchestre le workflow complet :
'   1. Validation des prérequis
'   2. Récupération des données PBS (14 tableaux)
'   3. Validation données RAM Report selon gate
'   4. Sélection répertoire + copie template
'   5. Génération Providor (BS, ML, Corr, Prev, GP, LCC)
'   6. Finalisation

    Dim chemin As String
    Dim wb As Workbook
    Dim fichierTest As String
    
    ' Désactiver fonctionnalités Excel pour performance
    TurnOffFunctionality
    
    ' Validation et récupération données UserForm
    If Not InitialiserContexteProvidor(test) Then
        TurnOnFunctionality
        Exit Sub
    End If
    
    ' Récupération données PBS depuis DATA_ASL
    Call RecupererDonneesPBS_Providor(test)

    If Not (test) Then
        ' Validation données RAM Report selon gate
        If Not ValiderDonneesRAMProvidor Then
            TurnOnFunctionality
            Exit Sub
        End If
    End If
    
    ' Sélection répertoire destination
    If Not SelectRepertoireProvidor(chemin, test) Then
        TurnOnFunctionality
        Exit Sub
    End If
    
    ' Copier et ouvrir template Providor
    Set wb = CopierTemplateProvidor(chemin, test)
    If wb Is Nothing Then
        TurnOnFunctionality
        Exit Sub
    End If
    
    
    
    ' Générer le Providor complet
    Call GenererProvidor(wb, test)
    
    ' Réactiver fonctionnalités Excel
    TurnOnFunctionality
    
    If test Then
        Call D_Providor_TNR.CollecterMetriquesTest(wb)
        Call D_Providor_TNR.AfficherResultatsTest
        Call D_Providor_TNR.NettoyerFichierTest(wb)
    Else
        ' Activer feuille LCC
        wb.Worksheets("LCC Calculation").Activate
    End If
    

End Sub
' =========================================================================
' FONCTION: INITIALISER CONTEXTE PROVIDOR
' =========================================================================
Private Function InitialiserContexteProvidor(test As Boolean) As Boolean
' Récupère et valide les paramètres du UserForm
' Retourne True si tout est OK, False sinon

    On Error GoTo ErrHandler
    
    ' Créer objets
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Récupérer paramètres système
    userpath = Environ$("UserProfile")
    
    'Definir les donnees projets avant generation
    If Not test Then
        Set Generateur = ActiveWorkbook
        Set DS1 = Generateur.Worksheets("DATA_ASL")

        Code_Projet = UserForm.Project_List.Value
        Projet_Abrev = Abrev_Projet(Code_Projet)
        Revision = UserForm.Revision_List.Value
        gate = UserForm.Gate_List.Value
        BoolMaintPrev = UserForm.CheckBoxPreventiveProvidor.Value

        If UserForm.Langue_FR.Value = True Then
            langue = "FR"
        ElseIf UserForm.Langue_EN.Value = True Then
            langue = "EN"
        Else
            MsgBox "Veuillez sélectionner une langue (FR ou EN)", vbExclamation
            InitialiserContexteProvidor = False
            Exit Function
        End If

        If Code_Projet = "" Then
            MsgBox "Veuillez sélectionner un projet", vbExclamation
            InitialiserContexteProvidor = False
            Exit Function
        End If

        If Revision = "" Then
            MsgBox "Veuillez saisir une révision", vbExclamation
            InitialiserContexteProvidor = False
            Exit Function
        End If

        If gate = "" Then
            MsgBox "Aucune gate n'est sélectionnée.", vbInformation, "Gate"
            InitialiserContexteProvidor = False
            Exit Function
        End If

    Else
        ' --- Mode test : valeurs fixes, pas de UserForm ---
        Set Generateur = ActiveWorkbook

        ' DS1 pointé pour éviter erreurs dans fonctions annexes
        On Error Resume Next
        Set DS1 = Generateur.Worksheets("DATA_ASL")
        On Error GoTo ErrHandler

        Code_Projet = "ETALON"
        Projet_Abrev = "ETALON"
        Revision = "A"
        gate = UserForm.Gate_List.Value           ' Gate récupérée depuis UserForm
        langue = "FR"
        BoolMaintPrev = True                      ' Tester aussi la préventive
    End If
    
    InitialiserContexteProvidor = True
    Exit Function
    
ErrHandler:
    MsgBox "Erreur lors de l'initialisation Providor : " & Err.description, vbCritical
    InitialiserContexteProvidor = False
    
End Function

' =========================================================================
' PROCÉDURE: RÉCUPÉRER DONNÉES PBS PROVIDOR
' =========================================================================
Private Sub RecupererDonneesPBS_Providor(test As Boolean)
' Récupère les 14 colonnes nécessaires du tableau Providor
    
    If Not (test) Then
        ' Identification
        tDTR = DS1.Range("Providor[DTR Reference (issue de PBS)]")
        tDES_PBS = DS1.Range("Providor[DESIGNATION (issue de PBS)]")
        tDES_BDD = DS1.Range("Providor[DESIGNATION (issue de BdD)]")
        tDES_MEF = DS1.Range("Providor[DESIGNATION (mise en forme pour copier dans Providor : construite à partir de celle de la PBS et de la désignation issue de BdD (ORCHESTRA)))]")
        tSS = DS1.Range("Providor[SOUS SYSTÈME (copiés de la PBS et classifiés selon génériques)]")
        
        ' Quantités et types
        tQTE = DS1.Range("Providor[Qte]")
        tNTI = DS1.Range("Providor[NTI]")
        tNodeType = DS1.Range("Providor[Node typed (LRU/SRU…)]")
        
        ' Prix
        tPrix_Cat_Spare = DS1.Range("Providor[ISP TIS FRANCE (FCS)]")
        tPrix_Cat_Repair = DS1.Range("Providor[Prix repa (ISP Catalogue Répa)]")
        tPrix_Manuel_Spare = DS1.Range("Providor[Prix Manuel Spare]")
        tPrix_Manuel_Repair = DS1.Range("Providor[Prix Manuel Repair]")
        tPrix_SpareProjet_Archive = DS1.Range("Providor[Prix Spare sur un projet archive]")
        tPrix_RepairProjet_Archive = DS1.Range("Providor[Prix Repair sur un projet archive]")
        tPrix_FRS_Spare = DS1.Range("Providor[Cout spare ALSTOM FRANCE FOURNISSEUR]")
        tPrix_FRS_Repair = DS1.Range("Providor[Prix repair ALSTOM FOURNISSEUR]")
        
        ' Temps corrective/préventive
        tTps_CorrPrev_Gen_HARD = DS1.Range("Providor[Durée de la tâche corrective et préventive générique (offre et PGR) HARDWARE]")
        tTps_CorrPrev_Ana_HARD = DS1.Range("Providor[Durée de la tâche corrective et préventive analytique (CGR) HARDWARE]")
        tTps_Prev_Gen_SOFT = DS1.Range("Providor[Durée de la tâche préventive générique SOFTWARE (offre et PGR)]")
        tTps_Prev_Ana_SOFT = DS1.Range("Providor[Durée de la tâche préventive analytique SOFTWARE (CGR)]")
        
        ' Durée de vie
        tDureevie_Gen_HARD = DS1.Range("Providor[DUREE DE VIE LIMITEE générique (offre et PGR) HARDWARE]")
        tDureevie_Gen_SOFT = DS1.Range("Providor[DUREE DE VIE LIMITEE générique (offre et PGR) SOFTWARE]")
        tDureevie_Ana_HARD = DS1.Range("Providor[DUREE DE VIE LIMITEE (RAM report specifique projet (CGR)) HARDWARE]")
        tDureevie_Ana_SOFT = DS1.Range("Providor[DUREE DE VIE LIMITEE (RAM report specifique projet (CGR)) SOFTWARE]")
        
        ' MTBF
        tMTBF_gen = DS1.Range("Providor[MTBF générique (offre et PGR)]")
        tMTBF_Intr = DS1.Range("Providor[MTBF intrinsèque (RAM report specifique projet (CGR))]")
        tMTBF_fct = DS1.Range("Providor[MTBF fonctionnel (RAM report specifique projet (CGR))]")
    Else
        Call InitPBS_Etalon
        Call InjecterDonneesEtalon
    End If
End Sub

' =========================================================================
' FONCTION: VALIDER DONNÉES RAM PROVIDOR
' =========================================================================
Private Function ValiderDonneesRAMProvidor() As Boolean
' Vérifie si les données du RAM Report sont requises et présentes
' Retourne True pour continuer, False pour arrêter

    Dim MTBF_intr_pres As Boolean, MTBF_fonct_pres As Boolean
    Dim Tps_CorrPrev_Ana_HARD_pres As Boolean, Tps_CorrPrev_Ana_SOFT_pres As Boolean
    Dim Dureevie_Ana_HARD_pres As Boolean, Dureevie_Ana_SOFT_pres As Boolean
    
    ' Déterminer si on est avant ou après CGR
    Select Case gate
        Case "PGR", "GO ORDER", "BA", "IGR"
            ' Avant CGR : pas de vérification RAM Report
            AvantCGR = True
            RAMReportDATA = False
            ValiderDonneesRAMProvidor = True
            
        Case "CGR", "FEI", "IQA"
            ' Après CGR : vérification obligatoire
            AvantCGR = False
            
            ' Vérifier présence données
            MTBF_intr_pres = IsInArray_1D("", tMTBF_Intr)
            MTBF_fonct_pres = IsInArray_1D("", tMTBF_fct)
            Tps_CorrPrev_Ana_HARD_pres = IsInArray_1D("", tTps_CorrPrev_Ana_HARD)
            Tps_CorrPrev_Ana_SOFT_pres = IsInArray_1D("", tTps_Prev_Ana_SOFT)
            Dureevie_Ana_HARD_pres = IsInArray_1D("", tDureevie_Ana_HARD)
            Dureevie_Ana_SOFT_pres = IsInArray_1D("", tDureevie_Ana_SOFT)
            
            If MTBF_fonct_pres Or MTBF_intr_pres Or Tps_CorrPrev_Ana_HARD_pres Or _
               Tps_CorrPrev_Ana_SOFT_pres Or Dureevie_Ana_HARD_pres Or Dureevie_Ana_SOFT_pres Then
                ' Données incomplètes : demander confirmation
                If MsgBox("Toutes les données spécifiques du RAM Report ne sont pas renseignées (obligatoire dès CGR)" & vbCrLf & _
                          "Souhaitez-vous continuer avec les données intrinsèques ?" & vbCrLf & vbCrLf & _
                          "Les données non renseignées seront remplacées par les données intrinsèques si présentes, et par les valeurs génériques si rien n'est renseigné", _
                          vbYesNo + vbQuestion) = vbNo Then
                    MsgBox "Arrêt de la procédure, vous pouvez renseigner les données du RAM Report (obligatoire à CGR)", vbInformation
                    ValiderDonneesRAMProvidor = False
                    Exit Function
                End If
                RAMReportDATA = False
            Else
                ' Toutes les données présentes
                MsgBox "Toutes les données spécifiques du RAM Report sont renseignées." & vbCrLf & _
                       "La procédure peut continuer", vbInformation
                RAMReportDATA = True
            End If
            
            ValiderDonneesRAMProvidor = True
            
        Case Else
            ' Gate inconnue : considérer comme avant CGR
            AvantCGR = True
            RAMReportDATA = False
            ValiderDonneesRAMProvidor = True
    End Select
    
End Function

' =========================================================================
' FONCTION: SÉLECTION RÉPERTOIRE PROVIDOR
' =========================================================================
Private Function SelectRepertoireProvidor(ByRef chemin As String, test As Boolean) As Boolean
' Dialogue sélection répertoire + gestion fichier existant
    
    Dim Repertoire As FileDialog
    Dim nomFichier As String
    Dim fichierComplet As String
    Dim wb As Workbook
    Dim resp As VbMsgBoxResult, respSave As VbMsgBoxResult
    
    If Not (test) Then
        ' Dialogue sélection répertoire
        MsgBox "Veuillez sélectionner le répertoire où vous souhaitez déposer le Providor pour le projet : " & Projet_Abrev, vbOKOnly
        Set Repertoire = Application.FileDialog(msoFileDialogFolderPicker)
        Repertoire.Show
        
        If Repertoire.SelectedItems.count = 0 Then
            SelectRepertoireProvidor = False
            Exit Function
        End If
        
        chemin = Repertoire.SelectedItems(1)
        nomFichier = Projet_Abrev & "_" & Month(Date) & "_" & Year(Date) & "_Providor.xlsm"
        fichierComplet = chemin & "\" & nomFichier
        
        ' Vérifier si fichier déjà ouvert
        For Each wb In Application.Workbooks
            If StrComp(wb.Name, nomFichier, vbTextCompare) = 0 Then
                resp = MsgBox("Un fichier nommé '" & nomFichier & "' est déjà ouvert." & vbCrLf & _
                              "Voulez-vous le fermer pour continuer ?", vbYesNo + vbQuestion, "Classeur déjà ouvert")
                If resp = vbYes Then
                    If Not wb.Saved Then
                        respSave = MsgBox("Le classeur ouvert contient des modifications non enregistrées." & vbCrLf & _
                                          "Voulez-vous enregistrer avant la fermeture ?", _
                                          vbYesNoCancel + vbExclamation, "Enregistrer les modifications ?")
                        Select Case respSave
                            Case vbYes: wb.Close SaveChanges:=True
                            Case vbNo:  wb.Close SaveChanges:=False
                            Case vbCancel
                                MsgBox "Opération annulée.", vbInformation
                                SelectRepertoireProvidor = False
                                Exit Function
                        End Select
                    Else
                        wb.Close SaveChanges:=False
                    End If
                Else
                    MsgBox "Opération annulée. Le fichier n'a pas été copié.", vbInformation
                    SelectRepertoireProvidor = False
                    Exit Function
                End If
                Exit For
            End If
        Next wb
        
        ' Vérifier si fichier existe déjà dans répertoire
        If Dir(fichierComplet) <> "" Then
            resp = MsgBox("Le fichier '" & nomFichier & "' existe déjà dans ce répertoire." & vbCrLf & _
                         "Souhaitez-vous le remplacer ?", vbYesNo + vbQuestion, "Fichier existant détecté")
            If resp = vbNo Then
                MsgBox "Opération annulée. Le fichier n'a pas été copié.", vbInformation
                SelectRepertoireProvidor = False
                Exit Function
            End If
        End If
    Else
    
        fichierComplet = Environ("TEMP") & "\Providor_Etalon.xlsm"
        
    End If
    
    SelectRepertoireProvidor = True

End Function

' =========================================================================
' FONCTION: COPIER TEMPLATE PROVIDOR
' =========================================================================
Private Function CopierTemplateProvidor(ByVal chemin As String, ByVal test As Boolean) As Workbook
' Copie le template Providor et l'ouvre
' Retourne le Workbook ouvert ou Nothing si erreur

    Dim nomFichier As String
    Dim fichierComplet As String
    
    On Error GoTo ErrHandler
    
    If Not (test) Then
    
        ' Construire nom fichier
        nomFichier = Projet_Abrev & "_" & Month(Date) & "_" & Year(Date) & "_Providor.xlsb"
        fichierComplet = chemin & "\" & nomFichier
        
        ' Copier template
        Application.DisplayAlerts = False
        fso.CopyFile userpath & PROVIDOR_TEMPLATE_PATH, fichierComplet
        Application.DisplayAlerts = True
        
        ' Ouvrir fichier
        Set CopierTemplateProvidor = Workbooks.Open(fichierComplet)
        
        Exit Function
    Else
        
        fichierComplet = Environ("TEMP") & "\Providor_Etalon.xlsb"
        
        ' Copier template
        Application.DisplayAlerts = False
        fso.CopyFile userpath & PROVIDOR_TEMPLATE_PATH, fichierComplet
        Application.DisplayAlerts = True
        
        ' Ouvrir fichier
        Set CopierTemplateProvidor = Workbooks.Open(fichierComplet)
        
        Exit Function
    
    End If
    
ErrHandler:
    MsgBox "Erreur lors de la copie du template Providor : " & Err.description, vbCritical
    Set CopierTemplateProvidor = Nothing
End Function

' =========================================================================
' PROCÉDURE: GÉNÉRER PROVIDOR
' =========================================================================
Private Sub GenererProvidor(ByRef wb As Workbook, test As Boolean)
' Orchestre la génération complète du Providor
' Appelle toutes les fonctions de remplissage dans l'ordre

    ' Définir feuilles du Providor
    Set GP = wb.Worksheets("General Parameters")
    Set BS = wb.Worksheets("Breakdown structure")
    Set ML = wb.Worksheets("Material list")
    Set CM = wb.Worksheets("Reliability & Corr. maintenance")
    Set PM = wb.Worksheets("Preventive maintenance")
    Set SL = wb.Worksheets("Skill List")
    Set TL = wb.Worksheets("Tool List")
    Set LCC = wb.Worksheets("LCC Calculation")
    
    'Etape 4 : Initialisation des colonnes du Providor
    Call InitialiserCadrageColonnes  ' O_Providor_CadrageColonnes
    
    If Not ValiderCadrageColonnes() Then
        MsgBox "Échec initialisation cadrage colonnes. Veuillez faire un ticket ATMAP", vbCritical
        GoTo CleanExit
    End If
    
    ' Remplissage dans l'ordre
    Call remplissage_BreakdownStructure          ' H_Providor_BreakdownStructure
    Call remplissage_MaterialList                ' H_Providor_BreakdownStructure
    Call remplissage_corr                        ' I_Providor_Maintenance
    
    ' Remplissage préventive (si activé)
    If BoolMaintPrev Then
        Call remplissage_prev                    ' I_Providor_Maintenance
    End If
    
    Call remplissage_skillAndcontrolSheets       ' J_Providor_Finalisation
    Call remplissage_GeneralParameters           ' J_Providor_Finalisation
    
    TurnOnFunctionality
    
    ' Exécution calcul LCC
    Call ExecuterCalculLCC(wb)                   ' J_Providor_Finalisation
    
CleanExit:
    TurnOnFunctionality
    Exit Sub
    
End Sub

