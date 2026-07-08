'--------------------------------------------------------------------------------------------------------
'
'                                       Generateur Livrables SLI
'                                       Module D_Providor_Main (Version Refactorisee)
'                                       Developpement : LE LARGE Thomas (LGM/Alstom)
'                                       Module permettant de generer un Providor LCC
'                                       Refactorisation : Fevrier 2026
'
'-----------------------------------------------------------------------------------------------------------
Option Compare Text
Option Explicit

' =========================================================================
' PROCEDURE PRINCIPALE : MAIN_PROVIDOR
' =========================================================================
Public Sub Main_Providor(test As Boolean)
' Point d'entree principal pour la generation du Providor
' Orchestre le workflow complet :
'   1. Validation des prerequis
'   2. Recuperation des donnees PBS (14 tableaux)
'   3. Validation donnees RAM Report selon gate
'   4. Selection repertoire + copie template
'   5. Generation Providor (BS, ML, Corr, Prev, GP, LCC)
'   6. Finalisation

    Dim chemin As String
    Dim wb As Workbook
    Dim fichierTest As String
    
    ' Desactiver fonctionnalites Excel pour performance
    TurnOffFunctionality
    
    ' Validation et recuperation donnees UserForm
    If Not InitialiserContexteProvidor(test) Then
        TurnOnFunctionality
        Exit Sub
    End If
    
    ' Recuperation donnees PBS depuis DATA_ASL
    Call RecupererDonneesPBS_Providor(test)

    If Not (test) Then
        ' Validation donnees RAM Report selon gate
        If Not ValiderDonneesRAMProvidor Then
            TurnOnFunctionality
            Exit Sub
        End If
    End If
    
    ' Selection repertoire destination
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
    
    
    
    ' Generer le Providor complet
    Call GenererProvidor(wb, test)
    
    ' Reactiver fonctionnalites Excel
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
' Recupere et valide les parametres du UserForm
' Retourne True si tout est OK, False sinon

    On Error GoTo ErrHandler
    
    ' Creer objets
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Recuperer parametres systeme
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
            MsgBox "Veuillez selectionner une langue (FR ou EN)", vbExclamation
            InitialiserContexteProvidor = False
            Exit Function
        End If

        If Code_Projet = "" Then
            MsgBox "Veuillez selectionner un projet", vbExclamation
            InitialiserContexteProvidor = False
            Exit Function
        End If

        If Revision = "" Then
            MsgBox "Veuillez saisir une revision", vbExclamation
            InitialiserContexteProvidor = False
            Exit Function
        End If

        If gate = "" Then
            MsgBox "Aucune gate n'est selectionnee.", vbInformation, "Gate"
            InitialiserContexteProvidor = False
            Exit Function
        End If

    Else
        ' --- Mode test : valeurs fixes, pas de UserForm ---
        Set Generateur = ActiveWorkbook

        ' DS1 pointe pour eviter erreurs dans fonctions annexes
        On Error Resume Next
        Set DS1 = Generateur.Worksheets("DATA_ASL")
        On Error GoTo ErrHandler

        Code_Projet = "ETALON"
        Projet_Abrev = "ETALON"
        Revision = "A"
        gate = UserForm.Gate_List.Value           ' Gate recuperee depuis UserForm
        langue = "FR"
        BoolMaintPrev = True                      ' Tester aussi la preventive
    End If
    
    InitialiserContexteProvidor = True
    Exit Function
    
ErrHandler:
    MsgBox "Erreur lors de l'initialisation Providor : " & Err.description, vbCritical
    InitialiserContexteProvidor = False
    
End Function

' =========================================================================
' PROCEDURE: RECUPERER DONNEES PBS PROVIDOR
' =========================================================================
Private Sub RecupererDonneesPBS_Providor(test As Boolean)
' Recupere les 14 colonnes necessaires du tableau Providor
    
    If Not (test) Then
        ' Identification
        tDTR = DS1.Range("Providor[DTR Reference (issue de PBS)]")
        tDES_PBS = DS1.Range("Providor[DESIGNATION (issue de PBS)]")
        tDES_BDD = DS1.Range("Providor[DESIGNATION (issue BdD)]")
        tDES_MEF = DS1.Range("Providor[DESIGNATION (Providor)]")
        tSS = DS1.Range("Providor[SOUS SYSTEME]")
        
        ' Quantites et types
        tQTE = DS1.Range("Providor[Qte]")
        tNTI = DS1.Range("Providor[NTI]")
        tNodeType = DS1.Range("Providor[Node typed (LRU/SRU...)]")
        
        ' Prix
        tPrix_Cat_Spare = DS1.Range("Providor[Prix spare catalogue]")
        tPrix_Cat_Repair = DS1.Range("Providor[Prix repair catalogue]")
        tPrix_SpareProjet_Archive = DS1.Range("Providor[Prix Spare projet archive]")
        tPrix_RepairProjet_Archive = DS1.Range("Providor[Prix Repair projet archive]")
        tPrix_FRS_Spare = DS1.Range("Providor[Cout spare ALSTOM FRANCE FOURNISSEUR]")
        tPrix_FRS_Repair = DS1.Range("Providor[Prix repair ALSTOM FOURNISSEUR]")
        
        ' Temps corrective/preventive
        tTps_CorrPrev_Gen_HARD = DS1.Range("Providor[Duree tache correctif+preventif HARD generique]")
        tTps_CorrPrev_Ana_HARD = DS1.Range("Providor[Duree tache correctif+preventif HARD analytique]")
        tTps_Prev_Gen_SOFT = DS1.Range("Providor[Duree tache preventif SOFT generique]")
        tTps_Prev_Ana_SOFT = DS1.Range("Providor[Duree tache preventif SOFT analytique]")
        
        ' Duree de vie
        tDureevie_Gen_HARD = DS1.Range("Providor[DVL HARD generique]")
        tDureevie_Gen_SOFT = DS1.Range("Providor[DVL SOFT generique]")
        tDureevie_Ana_HARD = DS1.Range("Providor[DVL HARD specifique]")
        tDureevie_Ana_SOFT = DS1.Range("Providor[DVL SOFT specifique]")
        
        ' MTBF
        tMTBF_gen = DS1.Range("Providor[MTBF generique (offre+PGR)]")
        tMTBF_Intr = DS1.Range("Providor[MTBF intrinseque RAM Report specifique]")
        tMTBF_fct = DS1.Range("Providor[MTBF fonctionnel RAM Report specifique]")
    Else
        Call InitPBS_Etalon
        Call InjecterDonneesEtalon
    End If
End Sub

' =========================================================================
' FONCTION: VALIDER DONNEES RAM PROVIDOR
' =========================================================================
Private Function ValiderDonneesRAMProvidor() As Boolean
' Verifie si les donnees du RAM Report sont requises et presentes
' Retourne True pour continuer, False pour arreter

    Dim MTBF_intr_pres As Boolean, MTBF_fonct_pres As Boolean
    Dim Tps_CorrPrev_Ana_HARD_pres As Boolean, Tps_CorrPrev_Ana_SOFT_pres As Boolean
    Dim Dureevie_Ana_HARD_pres As Boolean, Dureevie_Ana_SOFT_pres As Boolean
    
    ' Determiner si on est avant ou apres CGR
    Select Case gate
        Case "PGR", "GO ORDER", "BA", "IGR"
            ' Avant CGR : pas de verification RAM Report
            AvantCGR = True
            RAMReportDATA = False
            ValiderDonneesRAMProvidor = True
            
        Case "CGR", "FEI", "IQA"
            ' Apres CGR : verification obligatoire
            AvantCGR = False
            
            ' Verifier presence donnees
            MTBF_intr_pres = IsInArray_1D("", tMTBF_Intr)
            MTBF_fonct_pres = IsInArray_1D("", tMTBF_fct)
            Tps_CorrPrev_Ana_HARD_pres = IsInArray_1D("", tTps_CorrPrev_Ana_HARD)
            Tps_CorrPrev_Ana_SOFT_pres = IsInArray_1D("", tTps_Prev_Ana_SOFT)
            Dureevie_Ana_HARD_pres = IsInArray_1D("", tDureevie_Ana_HARD)
            Dureevie_Ana_SOFT_pres = IsInArray_1D("", tDureevie_Ana_SOFT)
            
            If MTBF_fonct_pres Or MTBF_intr_pres Or Tps_CorrPrev_Ana_HARD_pres Or Tps_CorrPrev_Ana_SOFT_pres Or Dureevie_Ana_HARD_pres Or Dureevie_Ana_SOFT_pres Then
                ' Donnees incompletes : demander confirmation
                If MsgBox("Toutes les donnees specifiques du RAM Report ne sont pas renseignees (obligatoire des CGR)" & vbCrLf & "Souhaitez-vous continuer avec les donnees intrinseques ?" & vbCrLf & vbCrLf & "Les donnees non renseignees seront remplacees par les donnees intrinseques si presentes, et par les valeurs generiques si rien n'est renseigne", vbYesNo + vbQuestion) = vbNo Then
                    MsgBox "Arret de la procedure, vous pouvez renseigner les donnees du RAM Report (obligatoire a CGR)", vbInformation
                    ValiderDonneesRAMProvidor = False
                    Exit Function
                End If
                RAMReportDATA = False
            Else
                ' Toutes les donnees presentes
                MsgBox "Toutes les donnees specifiques du RAM Report sont renseignees." & vbCrLf & "La procedure peut continuer", vbInformation
                RAMReportDATA = True
            End If
            
            ValiderDonneesRAMProvidor = True
            
        Case Else
            ' Gate inconnue : considerer comme avant CGR
            AvantCGR = True
            RAMReportDATA = False
            ValiderDonneesRAMProvidor = True
    End Select
    
End Function

' =========================================================================
' FONCTION: SELECTION REPERTOIRE PROVIDOR
' =========================================================================
Private Function SelectRepertoireProvidor(ByRef chemin As String, test As Boolean) As Boolean
' Dialogue selection repertoire + gestion fichier existant
    
    Dim Repertoire As FileDialog
    Dim nomFichier As String
    Dim fichierComplet As String
    Dim wb As Workbook
    Dim resp As VbMsgBoxResult, respSave As VbMsgBoxResult
    
    If Not (test) Then
        ' Dialogue selection repertoire
        MsgBox "Veuillez selectionner le repertoire ou vous souhaitez deposer le Providor pour le projet : " & Projet_Abrev, vbOKOnly
        Set Repertoire = Application.FileDialog(msoFileDialogFolderPicker)
        Repertoire.Show
        
        If Repertoire.SelectedItems.count = 0 Then
            SelectRepertoireProvidor = False
            Exit Function
        End If
        
        chemin = Repertoire.SelectedItems(1)
        nomFichier = Projet_Abrev & "_" & Month(Date) & "_" & Year(Date) & "_Providor.xlsm"
        fichierComplet = chemin & "\" & nomFichier
        
        ' Verifier si fichier deja ouvert
        For Each wb In Application.Workbooks
            If StrComp(wb.Name, nomFichier, vbTextCompare) = 0 Then
                resp = MsgBox("Un fichier nomme '" & nomFichier & "' est deja ouvert." & vbCrLf & "Voulez-vous le fermer pour continuer ?", vbYesNo + vbQuestion, "Classeur deja ouvert")
                If resp = vbYes Then
                    If Not wb.Saved Then
                        respSave = MsgBox("Le classeur ouvert contient des modifications non enregistrees." & vbCrLf & "Voulez-vous enregistrer avant la fermeture ?", vbYesNoCancel + vbExclamation, "Enregistrer les modifications ?")
                        Select Case respSave
                            Case vbYes: wb.Close SaveChanges:=True
                            Case vbNo:  wb.Close SaveChanges:=False
                            Case vbCancel
                                MsgBox "Operation annulee.", vbInformation
                                SelectRepertoireProvidor = False
                                Exit Function
                        End Select
                    Else
                        wb.Close SaveChanges:=False
                    End If
                Else
                    MsgBox "Operation annulee. Le fichier n'a pas ete copie.", vbInformation
                    SelectRepertoireProvidor = False
                    Exit Function
                End If
                Exit For
            End If
        Next wb
        
        ' Verifier si fichier existe deja dans repertoire
        If Dir(fichierComplet) <> "" Then
            resp = MsgBox("Le fichier '" & nomFichier & "' existe deja dans ce repertoire." & vbCrLf & "Souhaitez-vous le remplacer ?", vbYesNo + vbQuestion, "Fichier existant detecte")
            If resp = vbNo Then
                MsgBox "Operation annulee. Le fichier n'a pas ete copie.", vbInformation
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
' PROCEDURE: GENERER PROVIDOR
' =========================================================================
Private Sub GenererProvidor(ByRef wb As Workbook, test As Boolean)
' Orchestre la generation complete du Providor
' Appelle toutes les fonctions de remplissage dans l'ordre

    ' Definir feuilles du Providor
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
        MsgBox "Echec initialisation cadrage colonnes. Veuillez faire un ticket ATMAP", vbCritical
        GoTo CleanExit
    End If
    
    ' Remplissage dans l'ordre
    Call remplissage_BreakdownStructure          ' H_Providor_BreakdownStructure
    Call remplissage_MaterialList                ' H_Providor_BreakdownStructure
    Call remplissage_corr                        ' I_Providor_Maintenance
    
    ' Remplissage preventive (si active)
    If BoolMaintPrev Then
        Call remplissage_prev                    ' I_Providor_Maintenance
    End If
    
    Call remplissage_skillAndcontrolSheets       ' J_Providor_Finalisation
    Call remplissage_GeneralParameters           ' J_Providor_Finalisation
    
    TurnOnFunctionality
    
    ' Execution calcul LCC
    Call ExecuterCalculLCC(wb)                   ' J_Providor_Finalisation
    
CleanExit:
    TurnOnFunctionality
    Exit Sub
    
End Sub

