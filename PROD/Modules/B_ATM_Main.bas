Attribute VB_Name = "B_ATM_Main"
'--------------------------------------------------------------------------------------------------------
'
'                                       Generateur Livrables SLI
'                                       Module A_ATM
'                                       Developpement : LE LARGE Thomas (LGM/Alstom)
'                                       Module permettant de générer des fiches ATM au format V1 et 2.0
'
'
'-----------------------------------------------------------------------------------------------------------
Option Compare Text
Option Explicit


' =========================================================================
' PROCÉDURE PRINCIPALE : MAIN_ATM
' =========================================================================
Public Sub Main_ATM()
' Point d'entrée principal pour la génération des ATM
' Orchestre le workflow complet :
'   1. Validation des prérequis
'   2. Récupération des données utilisateur
'   3. Récupération des données PBS
'   4. Validation données RAM Report (si CGR+)
'   5. Appel de la génération (V1 ou V2)

    ' Désactiver fonctionnalités Excel pour performance
    TurnOffFunctionality
    
    ' Validation et récupération données UserForm
    If Not InitialiserContexteGeneration Then
        TurnOnFunctionality
        Exit Sub
    End If
    
    ' Récupération données PBS depuis DATA_ASL
    Call RecupererDonneesPBS
    
    ' Validation données RAM Report selon gate
    If Not ValiderDonneesRAMReport Then
        TurnOnFunctionality
        Exit Sub
    End If
    
    ' Déterminer version ATM à générer
    If Not DeterminerVersionATM Then
        TurnOnFunctionality
        Exit Sub
    End If
    
    ' Déterminer langue
    If Not DeterminerLangue Then
        TurnOnFunctionality
        Exit Sub
    End If
    
    ' Générer ATM selon version
    Call ExecuterGenerationATM
    
    ' Réactiver fonctionnalités Excel
    TurnOnFunctionality

End Sub
' =========================================================================
' FONCTION: INITIALISER CONTEXTE GÉNÉRATION
' =========================================================================
Private Function InitialiserContexteGeneration() As Boolean
' Récupère et valide les paramètres du UserForm
' Retourne True si tout est OK, False sinon

    On Error GoTo ErrHandler
    
    ' Définir classeur générateur
    Set Generateur = ActiveWorkbook
    Set DS1 = Generateur.Worksheets("DATA_ASL")
    
    ' Récupérer paramètres projet
    Code_Projet = UserForm.Project_List.Value
    Projet_Abrev = Abrev_Projet(Code_Projet)
    Revision = UserForm.Revision_List.Value
    gate = UserForm.Gate_List.Value
    
    ' Validation basique
    If Code_Projet = "" Then
        MsgBox "Veuillez sélectionner un projet", vbExclamation
        InitialiserContexteGeneration = False
        Exit Function
    End If
    
    If Revision = "" Then
        MsgBox "Veuillez saisir une révision", vbExclamation
        InitialiserContexteGeneration = False
        Exit Function
    End If
    
    If gate = "" Then
        MsgBox "Veuillez sélectionner une gate", vbExclamation
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
' PROCÉDURE: RÉCUPÉRER DONNÉES PBS
' =========================================================================
Private Sub RecupererDonneesPBS()
' Récupère toutes les colonnes nécessaires du tableau Providor
' Stocke dans variables globales tableaux (tDTR, tDES_PBS, etc.)

    ' Colonnes de base
    tDTR = DS1.Range("Providor[DTR Reference (issue de PBS)]")
    tDES_PBS = DS1.Range("Providor[DESIGNATION (issue de PBS)]")
    tDES_BDD = DS1.Range("Providor[DESIGNATION (issue de BdD)]")
    tDES_MEF = DS1.Range("Providor[DESIGNATION (mise en forme pour copier dans Providor : construite à partir de celle de la PBS et de la désignation issue de BdD (ORCHESTRA)))]")
    tSS = DS1.Range("Providor[SOUS SYSTÈME (copiés de la PBS et classifiés selon génériques)]")
    
    ' Colonnes HARD/SOFT et périodicité
    tHARD_SOFT = DS1.Range("Providor[HARD/SOFT]")
    tPERIODICITE = DS1.Range("Providor[DUREE DE VIE LIMITEE générique (offre et PGR) HARDWARE]")
    
    ' Colonnes RAM - MTBF
    tMTBF_gen = DS1.Range("Providor[MTBF générique (offre et PGR)]")
    tMTBF_fct = DS1.Range("Providor[MTBF fonctionnel (RAM report specifique projet (CGR))]")
    tMTBF_Intr = DS1.Range("Providor[MTBF intrinsèque (RAM report specifique projet (CGR))]")
    
    ' Colonnes RAM - Durée de vie HARDWARE
    tDureevie_Gen_HARD = DS1.Range("Providor[DUREE DE VIE LIMITEE générique (offre et PGR) HARDWARE]")
    tDureevie_Ana_HARD = DS1.Range("Providor[DUREE DE VIE LIMITEE (RAM report specifique projet (CGR)) HARDWARE]")
    
    ' Colonnes RAM - Durée de vie SOFTWARE
    tDureevie_Gen_SOFT = DS1.Range("Providor[DUREE DE VIE LIMITEE générique (offre et PGR) SOFTWARE]")
    tDureevie_Ana_SOFT = DS1.Range("Providor[DUREE DE VIE LIMITEE (RAM report specifique projet (CGR)) SOFTWARE]")
    
End Sub

' =========================================================================
' FONCTION: VALIDER DONNÉES RAM REPORT
' =========================================================================
Private Function ValiderDonneesRAMReport() As Boolean
' Vérifie si les données du RAM Report sont requises et présentes
' Retourne True pour continuer, False pour arrêter

    Dim MTBF_pres As Boolean
    Dim Duree_vie_HARD_pres As Boolean
    Dim Duree_vie_SOFT_pres As Boolean
    
    ' Déterminer si on est avant ou après CGR
    Select Case gate
        Case "PGR", "GO ORDER", "BA", "IGR"
            ' Avant CGR : pas de vérification RAM Report
            AvantCGR = True
            RAMReportDATA = False
            ValiderDonneesRAMReport = True
            
        Case "CGR", "FEI", "IQA"
            ' Après CGR : vérification obligatoire
            AvantCGR = False
            
            ' Vérifier présence données
            MTBF_pres = IsInArray_1D("", tMTBF_fct)
            Duree_vie_HARD_pres = IsInArray_1D("", tDureevie_Ana_HARD)
            Duree_vie_SOFT_pres = IsInArray_1D("", tDureevie_Ana_SOFT)
            
            If MTBF_pres Or Duree_vie_HARD_pres Or Duree_vie_SOFT_pres Then
                ' Données incomplètes : demander confirmation
                If Not DemanderConfirmationRAMPartiel Then
                    ValiderDonneesRAMReport = False
                    Exit Function
                End If
                RAMReportDATA = False
            Else
                ' Toutes les données présentes
                MsgBox "Toutes les données spécifiques du RAM Report sont renseignées." & vbCrLf & _
                       "La procédure peut continuer", vbInformation
                RAMReportDATA = True
            End If
            
            ValiderDonneesRAMReport = True
            
        Case Else
            ' Gate inconnue : considérer comme avant CGR
            AvantCGR = True
            RAMReportDATA = False
            ValiderDonneesRAMReport = True
    End Select
    
End Function

' =========================================================================
' FONCTION: DEMANDER CONFIRMATION RAM PARTIEL
' =========================================================================
Private Function DemanderConfirmationRAMPartiel() As Boolean
' Affiche boîte de dialogue pour données RAM incomplètes
' Retourne True si l'utilisateur souhaite continuer, False sinon

    Dim reponse As VbMsgBoxResult
    
    reponse = MsgBox( _
        "Toutes les données spécifiques du RAM Report ne sont pas renseignées (obligatoire dès CGR)" & vbCrLf & _
        "Souhaitez-vous continuer avec les données intrinsèques ?" & vbCrLf & vbCrLf & _
        "Les données non renseignées seront remplacées par les données intrinsèques si présentes, " & _
        "et par les valeurs génériques si rien n'est renseigné", _
        vbYesNo + vbQuestion, _
        "Données RAM Report incomplètes")
    
    If reponse = vbYes Then
        DemanderConfirmationRAMPartiel = True
    Else
        MsgBox "Arrêt de la procédure, vous pouvez renseigner les données du RAM Report (obligatoire à CGR)", _
               vbInformation
        DemanderConfirmationRAMPartiel = False
    End If
    
End Function

' =========================================================================
' FONCTION: DÉTERMINER VERSION ATM
' =========================================================================
Private Function DeterminerVersionATM() As Boolean
' Récupère la version ATM sélectionnée dans le UserForm
' Retourne True si version valide, False sinon

    If UserForm.V1_Option.Value = True Then
        Version_ATM = 1
        DeterminerVersionATM = True
        
    ElseIf UserForm.V2_Option.Value = True Then
        Version_ATM = 2
        DeterminerVersionATM = True
        
    Else
        MsgBox "Veuillez sélectionner une version ATM (V1 ou V2.0)", vbExclamation
        DeterminerVersionATM = False
    End If
    
End Function

' =========================================================================
' FONCTION: DÉTERMINER LANGUE
' =========================================================================
Private Function DeterminerLangue() As Boolean
' Récupère la langue sélectionnée dans le UserForm
' Retourne True si langue valide, False sinon

    If UserForm.Langue_FR.Value = True Then
        langue = "FR"
        DeterminerLangue = True
        
    ElseIf UserForm.Langue_EN.Value = True Then
        langue = "EN"
        DeterminerLangue = True
        
    Else
        MsgBox "Veuillez sélectionner une langue (FR ou EN)", vbExclamation
        DeterminerLangue = False
    End If
    
End Function

' =========================================================================
' PROCÉDURE: EXÉCUTER GÉNÉRATION ATM
' =========================================================================
Private Sub ExecuterGenerationATM()
' Appelle la procédure de génération appropriée selon la version
    
    Select Case Version_ATM
        Case 1
            ' Génération ATM V1 (fichier multi-feuilles)
            Call ATM_V1
            
        Case 2
            ' Génération ATM V2 (multi-fichiers)
            Call ATM_V2
            
        Case Else
            MsgBox "Version ATM non reconnue : " & Version_ATM, vbCritical
            Exit Sub
    End Select
    
    Exit Sub
    
End Sub


