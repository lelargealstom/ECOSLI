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

        Code_Projet = "ETALON"
        Projet_Abrev = "ETALON"
        Revision = "A"
        gate = UserForm.Gate_List.Value           ' Gate récupérée depuis UserForm
        langue = "FR"
        BoolMaintPrev = True                      ' Tester aussi la préventive
    End If
    
    InitialiserContexteProvidor = True
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

        ' HARD/SOFT (necessaire pour PM : calcul des occurrences SOFT annulees)
        tHARD_SOFT = DS1.Range("Providor[HARD/SOFT]")

        ' Duree de circulation du projet (ans) necessaire pour PM
        Call RecupererDureeCirculationProjet
    Else
        Call InitPBS_Etalon
        Call InjecterDonneesEtalon
    End If

    ' Construit tID_PBS : un identifiant incremental 1..N pour chaque ligne du
    ' tableau Providor. Resout les DTR dupliques (meme DTR peut apparaitre sur
    ' plusieurs lignes PBS avec valeurs differentes). Utilise ensuite par la BS,
    ' la ML et la CM pour tracer la ligne PBS d'origine de chaque entree.
    Dim nbLignesPBS As Long
    If IsArray(tDTR) Then
        nbLignesPBS = UBound(tDTR, 1)
    Else
        nbLignesPBS = 0
    End If

    Dim idxPBS As Long
    If nbLignesPBS > 0 Then
        ReDim tID_PBS(1 To nbLignesPBS, 1 To 1)
        For idxPBS = 1 To nbLignesPBS
            tID_PBS(idxPBS, 1) = idxPBS
        Next idxPBS
    End If

    ' Initialise les 3 mappings volatiles Ligne feuille -> ID_PBS. Vides au
    ' debut du pipeline, peuples lors des copies BS, BS->ML, BS->CM.
    Set dictBSRowToIDPBS = CreateObject("Scripting.Dictionary")
    Set dictMLRowToIDPBS = CreateObject("Scripting.Dictionary")
    Set dictCMRowToIDPBS = CreateObject("Scripting.Dictionary")
End Sub

' =========================================================================
' PROCEDURE: RECUPERER DUREE CIRCULATION PROJET
' =========================================================================
Private Sub RecupererDureeCirculationProjet()
' Recupere la duree de circulation du projet (ans) depuis Projet_SLI.
' Stocke le resultat dans la variable globale DureeCirculation.
' La connexion BdD est deja ouverte par Btn_Providor_Click avant Main_Providor.
    Dim sql As String
    Set REQSEL = CreateObject("ADODB.Recordset")
    sql = "SELECT [Duree de circulation_PROJET SLI] FROM [Projet_SLI$] " & _
          "WHERE [Nom projet_PROJET SLI] = '" & Replace(Code_Projet, "'", "''") & "'"
    Call open_REQSEL(sql)
    If Not REQSEL.EOF Then
        DureeCirculation = REQSEL.Fields(0).Value
    Else
        DureeCirculation = 0
    End If
    REQSEL.Close
End Sub

' =========================================================================
' FONCTION: VALIDER DONNÉES RAM PROVIDOR
' =========================================================================
Private Function ValiderDonneesRAMProvidor() As Boolean
' Positionne AvantCGR et RAMReportDATA selon la Gate.
' A partir de CGR : appelle AfficherAlerteRAMReportCGR (MsgBox informatif OK, non bloquant).
' Retourne toujours True (plus de dialogue bloquant Oui/Non).

    Select Case gate
        Case "PGR", "GO ORDER", "BA", "IGR"
            AvantCGR = True
            RAMReportDATA = False

        Case "CGR", "FEI", "IQA"
            AvantCGR = False
            Call AfficherAlerteRAMReportCGR

        Case Else
            AvantCGR = True
            RAMReportDATA = False
    End Select

    ValiderDonneesRAMProvidor = True
End Function

' =========================================================================
' PROCEDURE: AFFICHER ALERTE RAM REPORT CGR
' =========================================================================
Private Sub AfficherAlerteRAMReportCGR()
' Construit et affiche un MsgBox informatif OK avant la generation du Providor
' quand la Gate est a CGR ou au-dela. Positionne RAMReportDATA en consequence.
' 3 sections :
'   - Durees de vie limitee analytique HARDWARE/SOFTWARE : annonce du fallback generique
'     si la colonne analytique est entierement vide (meme critere que le fallback
'     applique en pratique dans remplissage_prev).
'   - Durees de tache analytique Corrective HARDWARE / Preventive SOFTWARE :
'     nombre de DTR sans valeur. Si > 0 : suggere de renseigner les taches ATM en BdD.
'   - MTBF : nombre de DTR sans MTBF fonctionnel. Rappel de la hierarchie
'     appliquee par DTR dans RemplirMTBF_Corr.

    Dim nbTotal As Long
    Dim nbSansTpsCorr As Long
    Dim nbSansTpsPrev As Long
    Dim nbSansMTBFfct As Long
    Dim fallbackDVL_HARD As Boolean
    Dim fallbackDVL_SOFT As Boolean
    Dim i As Long
    Dim msg As String

    nbTotal = UBound(tDTR, 1) - LBound(tDTR, 1) + 1

    nbSansTpsCorr = 0
    nbSansTpsPrev = 0
    nbSansMTBFfct = 0
    For i = LBound(tDTR, 1) To UBound(tDTR, 1)
        If Not EstValeurRenseignee(tTps_CorrPrev_Ana_HARD(i, 1)) Then nbSansTpsCorr = nbSansTpsCorr + 1
        If Not EstValeurRenseignee(tTps_Prev_Ana_SOFT(i, 1)) Then nbSansTpsPrev = nbSansTpsPrev + 1
        If Not EstValeurRenseignee(tMTBF_fct(i, 1)) Then nbSansMTBFfct = nbSansMTBFfct + 1
    Next i

    fallbackDVL_HARD = Not AuMoinsUneValeurNumerique(tDureevie_Ana_HARD)
    fallbackDVL_SOFT = Not AuMoinsUneValeurNumerique(tDureevie_Ana_SOFT)

    If fallbackDVL_HARD Or fallbackDVL_SOFT Or nbSansTpsCorr > 0 Or nbSansTpsPrev > 0 Or nbSansMTBFfct > 0 Then
        RAMReportDATA = False
    Else
        RAMReportDATA = True
    End If

    msg = "Attention : certaines donnees RAM Report sont incompletes a CGR." & vbCrLf
    msg = msg & "La generation du Providor continue avec les fallbacks ci-dessous." & vbCrLf & vbCrLf

    msg = msg & "== Durees de vie limitee ==" & vbCrLf
    If fallbackDVL_HARD Then
        msg = msg & "- HARDWARE : colonne analytique vide, bascule sur la colonne generique (Offre et PGR)." & vbCrLf
    Else
        msg = msg & "- HARDWARE : OK (colonne analytique utilisee)." & vbCrLf
    End If
    If fallbackDVL_SOFT Then
        msg = msg & "- SOFTWARE : colonne analytique vide, bascule sur la colonne generique (Offre et PGR)." & vbCrLf
    Else
        msg = msg & "- SOFTWARE : OK (colonne analytique utilisee)." & vbCrLf
    End If
    msg = msg & vbCrLf

    msg = msg & "== Durees de tache analytique ==" & vbCrLf
    msg = msg & "- Corrective HARDWARE : " & nbSansTpsCorr & " DTR sur " & nbTotal & " sans valeur."
    If nbSansTpsCorr > 0 Then msg = msg & " -> renseigner les taches ATM en BdD."
    msg = msg & vbCrLf
    msg = msg & "- Preventive SOFTWARE : " & nbSansTpsPrev & " DTR sur " & nbTotal & " sans valeur."
    If nbSansTpsPrev > 0 Then msg = msg & " -> renseigner les taches ATM en BdD."
    msg = msg & vbCrLf & vbCrLf

    msg = msg & "== MTBF ==" & vbCrLf
    msg = msg & "- " & nbSansMTBFfct & " DTR sur " & nbTotal & " sans MTBF fonctionnel." & vbCrLf
    msg = msg & "- Hierarchie appliquee par DTR : 1) MTBF fonctionnel  2) MTBF intrinseque  3) vide + orange"

    MsgBox msg, vbInformation, "Providor - Donnees RAM Report a CGR"

End Sub

' =========================================================================
' FONCTION: EST VALEUR RENSEIGNEE
' =========================================================================
Private Function EstValeurRenseignee(ByVal v As Variant) As Boolean
' Vrai si v est une valeur numerique non nulle (ni vide, ni Null, ni 0, ni chaine vide).
    If IsEmpty(v) Then Exit Function
    If IsNull(v) Then Exit Function
    If Not IsNumeric(v) Then Exit Function
    If Len(CStr(v)) = 0 Then Exit Function
    If CDbl(v) = 0 Then Exit Function
    EstValeurRenseignee = True
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
    Set CS = wb.Worksheets("Control Sheet")
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
    Call remplissage_ControlSheet                ' D_Providor_ControlSheet
    
    TurnOnFunctionality
    
    ' Exécution calcul LCC
    Call ExecuterCalculLCC(wb)                   ' J_Providor_Finalisation
    
CleanExit:
    TurnOnFunctionality
    Exit Sub
    
End Sub

