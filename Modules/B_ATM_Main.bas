Attribute VB_Name = "B_ATM_Main"
'--------------------------------------------------------------------------------------------------------
'
'                                       Generateur Livrables SLI
'                                       Module A_ATM
'                                       Developpement : LE LARGE Thomas (LGM/Alstom)
'                                       Module permettant de generer des fiches ATM au format V1 et 2.0
'
'
'-----------------------------------------------------------------------------------------------------------
Option Compare Text
Option Explicit


' =========================================================================
' PROCEDURE PRINCIPALE : MAIN_ATM
' =========================================================================
Public Sub Main_ATM()
' Point d'entree principal pour la generation des ATM
' Orchestre le workflow complet :
'   1. Validation des prerequis
'   2. Recuperation des donnees utilisateur
'   3. Recuperation des donnees PBS
'   4. Validation donnees RAM Report (si CGR+)
'   5. Appel de la generation (V1 ou V2)

    ' Desactiver fonctionnalites Excel pour performance
    TurnOffFunctionality
    
    ' Validation et recuperation donnees UserForm
    If Not InitialiserContexteGeneration Then
        TurnOnFunctionality
        Exit Sub
    End If
    
    ' Recuperation donnees PBS depuis DATA_ASL
    Call RecupererDonneesPBS
    
    ' Validation donnees RAM Report selon gate
    If Not ValiderDonneesRAMReport Then
        TurnOnFunctionality
        Exit Sub
    End If
    
    ' Determiner version ATM a generer
    If Not DeterminerVersionATM Then
        TurnOnFunctionality
        Exit Sub
    End If
    
    ' Determiner langue
    If Not DeterminerLangue Then
        TurnOnFunctionality
        Exit Sub
    End If
    
    ' Generer ATM selon version
    Call ExecuterGenerationATM
    
    ' Reactiver fonctionnalites Excel
    TurnOnFunctionality

End Sub
' =========================================================================
' FONCTION: INITIALISER CONTEXTE GENERATION
' =========================================================================
Private Function InitialiserContexteGeneration() As Boolean
' Recupere et valide les parametres du UserForm
' Retourne True si tout est OK, False sinon

    On Error GoTo ErrHandler
    
    ' Definir classeur generateur
    Set Generateur = ActiveWorkbook
    Set DS1 = Generateur.Worksheets("DATA_ASL")
    
    ' Recuperer parametres projet
    Code_Projet = UserForm.Project_List.Value
    Projet_Abrev = Abrev_Projet(Code_Projet)
    Revision = UserForm.Revision_List.Value
    gate = UserForm.Gate_List.Value
    
    ' Validation basique
    If Code_Projet = "" Then
        MsgBox "Veuillez selectionner un projet", vbExclamation
        InitialiserContexteGeneration = False
        Exit Function
    End If
    
    If Revision = "" Then
        MsgBox "Veuillez saisir une revision", vbExclamation
        InitialiserContexteGeneration = False
        Exit Function
    End If
    
    If gate = "" Then
        MsgBox "Veuillez selectionner une gate", vbExclamation
        InitialiserContexteGeneration = False
        Exit Function
    End If
    
    InitialiserContexteGeneration = True
    Exit Function
    
ErrHandler:
    MsgBox "Erreur lors de l'initialisation : " & Err.description, vbCritical
    InitialiserContexteGeneration = False
End Function

' =========================================================================
' PROCEDURE: RECUPERER DONNEES PBS
' =========================================================================
Private Sub RecupererDonneesPBS()
' Recupere toutes les colonnes necessaires du tableau Providor
' Stocke dans variables globales tableaux (tDTR, tDES_PBS, etc.)

    ' Colonnes de base
    tDTR = DS1.Range("Providor[DTR Reference (issue de PBS)]")
    tDES_PBS = DS1.Range("Providor[DESIGNATION (issue de PBS)]")
    tDES_BDD = DS1.Range("Providor[DESIGNATION (issue BdD)]")
    tDES_MEF = DS1.Range("Providor[DESIGNATION (Providor)]")
    tSS = DS1.Range("Providor[SOUS SYSTEME]")
    
    ' Colonnes HARD/SOFT et periodicite
    tHARD_SOFT = DS1.Range("Providor[HARD/SOFT]")
    tPERIODICITE = DS1.Range("Providor[DVL HARD generique]")
    
    ' Colonnes RAM - MTBF
    tMTBF_gen = DS1.Range("Providor[MTBF generique (offre+PGR)]")
    tMTBF_fct = DS1.Range("Providor[MTBF fonctionnel RAM Report specifique]")
    tMTBF_Intr = DS1.Range("Providor[MTBF intrinseque RAM Report specifique]")
    
    ' Colonnes RAM - Duree de vie HARDWARE
    tDureevie_Gen_HARD = DS1.Range("Providor[DVL HARD generique]")
    tDureevie_Ana_HARD = DS1.Range("Providor[DVL HARD specifique]")
    
    ' Colonnes RAM - Duree de vie SOFTWARE
    tDureevie_Gen_SOFT = DS1.Range("Providor[DVL SOFT generique]")
    tDureevie_Ana_SOFT = DS1.Range("Providor[DVL SOFT specifique]")
    
End Sub

' =========================================================================
' FONCTION: VALIDER DONNEES RAM REPORT
' =========================================================================
Private Function ValiderDonneesRAMReport() As Boolean
' Verifie si les donnees du RAM Report sont requises et presentes
' Retourne True pour continuer, False pour arreter

    Dim MTBF_pres As Boolean
    Dim Duree_vie_HARD_pres As Boolean
    Dim Duree_vie_SOFT_pres As Boolean
    
    ' Determiner si on est avant ou apres CGR
    Select Case gate
        Case "PGR", "GO ORDER", "BA", "IGR"
            ' Avant CGR : pas de verification RAM Report
            AvantCGR = True
            RAMReportDATA = False
            ValiderDonneesRAMReport = True
            
        Case "CGR", "FEI", "IQA"
            ' Apres CGR : verification obligatoire
            AvantCGR = False
            
            ' Verifier presence donnees
            MTBF_pres = IsInArray_1D("", tMTBF_fct)
            Duree_vie_HARD_pres = IsInArray_1D("", tDureevie_Ana_HARD)
            Duree_vie_SOFT_pres = IsInArray_1D("", tDureevie_Ana_SOFT)
            
            If MTBF_pres Or Duree_vie_HARD_pres Or Duree_vie_SOFT_pres Then
                ' Donnees incompletes : demander confirmation
                If Not DemanderConfirmationRAMPartiel Then
                    ValiderDonneesRAMReport = False
                    Exit Function
                End If
                RAMReportDATA = False
            Else
                ' Toutes les donnees presentes
                MsgBox "Toutes les donnees specifiques du RAM Report sont renseignees." & vbCrLf & "La procedure peut continuer", vbInformation
                RAMReportDATA = True
            End If
            
            ValiderDonneesRAMReport = True
            
        Case Else
            ' Gate inconnue : considerer comme avant CGR
            AvantCGR = True
            RAMReportDATA = False
            ValiderDonneesRAMReport = True
    End Select
    
End Function

' =========================================================================
' FONCTION: DEMANDER CONFIRMATION RAM PARTIEL
' =========================================================================
Private Function DemanderConfirmationRAMPartiel() As Boolean
' Affiche boite de dialogue pour donnees RAM incompletes
' Retourne True si l'utilisateur souhaite continuer, False sinon

    Dim reponse As VbMsgBoxResult
    
    reponse = MsgBox( "Toutes les donnees specifiques du RAM Report ne sont pas renseignees (obligatoire des CGR)" & vbCrLf & "Souhaitez-vous continuer avec les donnees intrinseques ?" & vbCrLf & vbCrLf & "Les donnees non renseignees seront remplacees par les donnees intrinseques si presentes, et par les valeurs generiques si rien n'est renseigne", vbYesNo + vbQuestion, "Donnees RAM Report incompletes")
    
    If reponse = vbYes Then
        DemanderConfirmationRAMPartiel = True
    Else
        MsgBox "Arret de la procedure, vous pouvez renseigner les donnees du RAM Report (obligatoire a CGR)", vbInformation
        DemanderConfirmationRAMPartiel = False
    End If
    
End Function

' =========================================================================
' FONCTION: DETERMINER VERSION ATM
' =========================================================================
Private Function DeterminerVersionATM() As Boolean
' Recupere la version ATM selectionnee dans le UserForm
' Retourne True si version valide, False sinon

    If UserForm.V1_Option.Value = True Then
        Version_ATM = 1
        DeterminerVersionATM = True
        
    ElseIf UserForm.V2_Option.Value = True Then
        Version_ATM = 2
        DeterminerVersionATM = True
        
    Else
        MsgBox "Veuillez selectionner une version ATM (V1 ou V2.0)", vbExclamation
        DeterminerVersionATM = False
    End If
    
End Function

' =========================================================================
' FONCTION: DETERMINER LANGUE
' =========================================================================
Private Function DeterminerLangue() As Boolean
' Recupere la langue selectionnee dans le UserForm
' Retourne True si langue valide, False sinon

    If UserForm.Langue_FR.Value = True Then
        langue = "FR"
        DeterminerLangue = True
        
    ElseIf UserForm.Langue_EN.Value = True Then
        langue = "EN"
        DeterminerLangue = True
        
    Else
        MsgBox "Veuillez selectionner une langue (FR ou EN)", vbExclamation
        DeterminerLangue = False
    End If
    
End Function

' =========================================================================
' PROCEDURE: EXECUTER GENERATION ATM
' =========================================================================
Private Sub ExecuterGenerationATM()
' Appelle la procedure de generation appropriee selon la version
    
    Select Case Version_ATM
        Case 1
            ' Generation ATM V1 (fichier multi-feuilles)
            Call ATM_V1
            
        Case 2
            ' Generation ATM V2 (multi-fichiers)
            Call ATM_V2
            
        Case Else
            MsgBox "Version ATM non reconnue : " & Version_ATM, vbCritical
            Exit Sub
    End Select
    
    Exit Sub
    
End Sub


