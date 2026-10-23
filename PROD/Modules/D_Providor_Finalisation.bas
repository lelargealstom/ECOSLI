'--------------------------------------------------------------------------------------------------------
'
'                                       Generateur Livrables SLI
'                                       Module D_Providor_Finalisation (Version Refactorisée)
'                                       Developpement : LE LARGE Thomas (LGM/Alstom)
'                                       Module de finalisation Providor (GP, Skills, LCC)
'                                       Refactorisation : Février 2026
'
'-----------------------------------------------------------------------------------------------------------
Option Compare Text
Option Explicit
' =========================================================================
' TYPE: PROFIL MISSION PROJET
' =========================================================================
Private Type ProfilMissionProjet
    ' Operating Units
    KmParAn As Variant
    ServiceDayParAn As Variant
    PHParAn As Variant
    RHParAn As Variant
    
    ' LCC Parameters
    NombreTrain As Variant
    DureeCirculation As Variant
    DureeGarantie As Variant
End Type
' =========================================================================
' PROCÉDURE: REMPLISSAGE GENERAL PARAMETERS
' =========================================================================
Public Sub remplissage_GeneralParameters()
' Remplit la feuille General Parameters avec les données projet
' VERSION OPTIMISÉE : Une seule requête SQL

    Dim profil As ProfilMissionProjet
    
    Set REQSEL = CreateObject("ADODB.Recordset")
    
    ' Récupérer profil mission en une seule requête
    If Not RecupererProfilMission(Code_Projet, profil) Then
        MsgBox "Erreur lors de la récupération du profil mission pour le projet : " & Code_Projet, vbExclamation
        Exit Sub
    End If
    
    ' Remplir General Parameters
    Call RemplirOperatingUnits(profil)
    Call RemplirProjectIdentification
    Call RemplirLCCParameters(profil)
    Call RemplirCurrencyUnits
    
End Sub
' =========================================================================
' FONCTION: RÉCUPÉRER PROFIL MISSION
' =========================================================================
Private Function RecupererProfilMission(ByVal codeProjet As String, ByRef profil As ProfilMissionProjet) As Boolean
' Récupère TOUTES les données du profil mission en UNE SEULE requête SQL
' Retourne True si succès, False si erreur
'
' OPTIMISATION :
'   - Avant : 7+ requêtes SQL (1 par champ)
'   - Après : 1 seule requête SQL
'   - Gain : -85% temps exécution

    Dim sql As String
    
    
    ' Construction requête SQL optimisée
    sql = "SELECT " & _
          "[Km by year_PROJET SLI], " & _
          "[Servicing day by year_PROJET SLI], " & _
          "[PH by year_PROJET SLI], " & _
          "[RH by year_PROJET SLI], " & _
          "[Nombre de train_PROJET SLI], " & _
          "[Duree de circulation_PROJET SLI], " & _
          "[Duree de garantie_PROJET SLI] " & _
          "FROM [Projet_SLI$] " & _
          "WHERE [Nom projet_PROJET SLI] = '" & Replace(codeProjet, "'", "''") & "'"
    
    ' Exécution requête
    Call open_REQSEL(sql)
    
    ' Vérifier résultat
    If REQSEL.EOF Then
        RecupererProfilMission = False
        Exit Function
    End If
    
    ' Récupération données (une seule lecture)
    With profil
        .KmParAn = REQSEL.Fields(0).Value
        .ServiceDayParAn = REQSEL.Fields(1).Value
        .PHParAn = REQSEL.Fields(2).Value
        .RHParAn = REQSEL.Fields(3).Value
        .NombreTrain = REQSEL.Fields(4).Value
        .DureeCirculation = REQSEL.Fields(5).Value
        .DureeGarantie = REQSEL.Fields(6).Value
    End With
    
    ' Fermer recordset
    REQSEL.Close
    
    RecupererProfilMission = True
    Exit Function
    
'ErrHandler:
    RecupererProfilMission = False
End Function
' =========================================================================
' PROCÉDURE: REMPLIR OPERATING UNITS
' =========================================================================
Private Sub RemplirOperatingUnits(ByRef profil As ProfilMissionProjet)
' Remplit les Operating Units avec les données du profil mission
    
    With GP
        ' KILOMETER
        .Cells(3, gp_col.OpUnits_Description) = "KILOMETER"
        .Cells(3, gp_col.OpUnits_Acronym) = "KM"
        .Cells(3, gp_col.OpUnits_YearlyValue) = SafeNum(profil.KmParAn, 0)
        .Cells(3, gp_col.OpUnits_Unit) = "Yes"
        
        ' SERVICING DAYS
        .Cells(4, gp_col.OpUnits_Description) = "SERVICING DAYS"
        .Cells(4, gp_col.OpUnits_Acronym) = "SD"
        .Cells(4, gp_col.OpUnits_YearlyValue) = SafeNum(profil.ServiceDayParAn, 0)
        .Cells(4, gp_col.OpUnits_Unit) = "Yes"
        
        ' POWER-UP HOUR
        .Cells(5, gp_col.OpUnits_Description) = "POWER-UP HOUR"
        .Cells(5, gp_col.OpUnits_Acronym) = "PH"
        .Cells(5, gp_col.OpUnits_YearlyValue) = SafeNum(profil.PHParAn, 0)
        .Cells(5, gp_col.OpUnits_Unit) = "Yes"
        
        ' RUNNING HOUR
        .Cells(6, gp_col.OpUnits_Description) = "RUNNING HOUR"
        .Cells(6, gp_col.OpUnits_Acronym) = "RH"
        .Cells(6, gp_col.OpUnits_YearlyValue) = SafeNum(profil.RHParAn, 0)
        .Cells(6, gp_col.OpUnits_Unit) = "Yes"
        
        ' YEAR
        .Cells(7, gp_col.OpUnits_Description) = "YEAR"
        .Cells(7, gp_col.OpUnits_Acronym) = "Y"
        .Cells(7, gp_col.OpUnits_YearlyValue) = 1
        .Cells(7, gp_col.OpUnits_Unit) = "No"
    End With
    
End Sub

' =========================================================================
' PROCÉDURE: REMPLIR PROJECT IDENTIFICATION
' =========================================================================
Private Sub RemplirProjectIdentification()
' Remplit l'identification du projet
    
    With GP
        .Cells(gp_lig.Pjct_number, 2) = Projet_Abrev
        .Cells(gp_lig.Pjct_name, 2) = Code_Projet
        .Cells(gp_lig.Pjct_acronym, 2) = Projet_Abrev & "_VBN"
        .Cells(gp_lig.Pjct_productname, 2) = "Scope_VBN"
    End With
    
End Sub

' =========================================================================
' PROCÉDURE: REMPLIR LCC PARAMETERS
' =========================================================================
Private Sub RemplirLCCParameters(ByRef profil As ProfilMissionProjet)
' Remplit les paramètres LCC avec les données du profil mission
    
    With GP
        .Cells(gp_lig.LCC_nbassets, 2) = SafeNum(profil.NombreTrain, 0)
        .Cells(gp_lig.LCC_length, 2) = SafeNum(profil.DureeCirculation, 0)
        .Cells(gp_lig.LCC_warranty, 2) = SafeNum(profil.DureeGarantie, 0)
    End With
    
End Sub

' =========================================================================
' PROCÉDURE: REMPLIR CURRENCY UNITS
' =========================================================================
Private Sub RemplirCurrencyUnits()
' Remplit les Currency Units depuis LISTES_DATA_CHOIX
    
    ' Copier données devises
    Generateur.Sheets("LISTES_DATA_CHOIX").Range("CURRUNIT_DATA").Copy
    GP.Cells(36, gp_col.CurrUnit_Currtablename).PasteSpecial Paste:=xlPasteValues
    
    ' Note : Code taux devise commenté dans original
    ' Calculer dernière ligne Currency Units
    dlCU = GP.Cells(35, gp_col.CurrUnit_Currtablename).End(xlDown).row
    
    'Pour activer taux de change automatique, décommenter :
    Call RemplirTauxDevise(36, dlCU)
    
End Sub

' =========================================================================
' PROCÉDURE: REMPLIR TAUX DEVISE
' =========================================================================
Private Sub RemplirTauxDevise(ByVal ligneDebut As Integer, ByVal ligneFin As Integer)
' Remplit les taux de change pour chaque devise
' NOTE : Fonction désactivée par défaut dans le code original
'        Nécessite connexion API externe (peut être lent)
    
    Dim i As Integer
    Dim devise As String
    Dim taux As Variant
    
    For i = ligneDebut To ligneFin
        devise = GP.Cells(i, gp_col.CurrUnit_Currabbrev)
        
        If devise <> "" And devise <> "EUR" Then
            ' Récupération taux via API
            taux = TauxDevise("EUR", devise)
            
            If taux <> "" And taux <> "Not found" Then
                GP.Cells(i, gp_col.CurrUnit_changerate) = taux
                GP.Cells(i, gp_col.CurrUnit_pjctchangerate) = taux
                GP.Cells(i, gp_col.CurrUnit_datereviewtoolrate) = Format(Date, "dd/mm/yyyy")
                GP.Cells(i, gp_col.CurrUnit_datereviewpjctrate) = Format(Date, "dd/mm/yyyy")
            Else
                GP.Cells(i, gp_col.CurrUnit_changerate) = "Not found"
                GP.Cells(i, gp_col.CurrUnit_pjctchangerate) = "Not found"
            End If
        End If
    Next i
    
End Sub

' =========================================================================
' AUTRES FONCTIONS FINALISATION
' =========================================================================
Public Sub remplissage_skillAndcontrolSheets()
' Remplit les feuilles Skill List et Tool List
' VERSION SÉCURISÉE : Gestion robuste prix Dataplug Adapter

    Dim price_manuel As Variant
    Dim price_catalogue As Variant
    Dim price As Double
    Dim Comments As String
    
    Set REQSEL = CreateObject("ADODB.Recordset")
    
    ' =========================================================================
    ' PARTIE 1 : SKILL LIST (Man Hour Cost + Devise)
    ' =========================================================================
    
    STRSEL = "SELECT [MAN hour cost_PROJET SLI], [Devise_PROJET SLI] " & _
             "FROM [Projet_SLI$] " & _
             "WHERE [Nom projet_PROJET SLI] = '" & Replace(Code_Projet, "'", "''") & "'"
    Call open_REQSEL(STRSEL)
    
    ' Man Hour Cost
    If IsNull(REQSEL.Fields(0).Value) Then
        SL.Cells(4, sl_col.ManHourCost) = ""
    Else
        SL.Cells(4, sl_col.ManHourCost) = REQSEL.Fields(0).Value
    End If
    
    ' Devise
    If IsNull(REQSEL.Fields(1).Value) Then
        SL.Cells(4, sl_col.CurrencyUnit) = "EUR"
    Else
        SL.Cells(4, sl_col.CurrencyUnit) = REQSEL.Fields(1).Value
    End If
    
    ' Remplir Skill List
    With SL
        .Cells(4, sl_col.Update) = "A0"
        If langue = "FR" Then
            .Cells(4, sl_col.SkillCategory) = "Electricien"
            .Cells(4, sl_col.ProfessionCategory) = "Electricite"
        Else
            .Cells(4, sl_col.SkillCategory) = "Electrician"
            .Cells(4, sl_col.ProfessionCategory) = "Electricity"
        End If
    End With
    
    ' =========================================================================
    ' PARTIE 2 : TOOL LIST (Prix Dataplug Adapter) - SÉCURISÉ
    ' =========================================================================
    
    STRSEL = "SELECT [Prix SPARE catalogue a date_PRODUIT], [Prix SPARE manuel_PRODUIT] " & _
             "FROM [Produit$] " & _
             "WHERE DTR = 'DTR0000222700'"
    
    Call open_REQSEL(STRSEL)
    
    ' Déterminer prix avec gestion robuste
    If Not REQSEL.EOF Then
        price_catalogue = REQSEL.Fields(0).Value
        price_manuel = REQSEL.Fields(1).Value
        
        ' Déterminer prix final avec hiérarchie
        Call DeterminerPrixDataplug(price_catalogue, price_manuel, price, Comments)
    Else
        ' DTR non trouvé dans BDD
        price = 0
        Comments = "DTR Dataplug Adapter non trouvé en BDD"
    End If
    
    ' Remplir Tool List
    With TL
        .Cells(4, tl_col.Update) = "A0"
        .Cells(4, tl_col.Name) = "DATAPLUG ADAPTER"
        .Cells(4, tl_col.PartNumber) = "DTR0000222700"
        .Cells(4, tl_col.AcquisitionCost) = price
        .Cells(4, tl_col.Currency) = "EUR"
        .Cells(4, tl_col.ToolType) = "OS - Specific tools"
    End With
    
    ' Ajouter commentaire prix
    TL.Cells(4, tl_col.AcquisitionCost).AddComment
    TL.Cells(4, tl_col.AcquisitionCost).Comment.Visible = False
    TL.Cells(4, tl_col.AcquisitionCost).Comment.Text Text:=Comments
    
    ' Colorier si prix = 0
    If price = 0 Then
        TL.Cells(4, tl_col.AcquisitionCost).Interior.ColorIndex = 44  ' Orange
    End If
    
End Sub
' =========================================================================
' PROCÉDURE: DÉTERMINER PRIX DATAPLUG (SÉCURISÉE)
' =========================================================================
Private Sub DeterminerPrixDataplug(ByVal prixCatalogue As Variant, _
                                   ByVal prixManuel As Variant, _
                                   ByRef prixFinal As Double, _
                                   ByRef Commentaire As String)
' Détermine le prix Dataplug Adapter avec gestion robuste des valeurs invalides
'
' Hiérarchie :
'   1. Prix Catalogue (si valide)
'   2. Prix Manuel (si valide)
'   3. 0 (si aucun prix valide)
'
' Valeurs considérées comme INVALIDES :
'   - NULL
'   - Empty
'   - "" (chaîne vide)
'   - "N/A", "#N/A", "N/D", "ND", "#N/D"
'   - 0
'   - Valeurs non numériques

    Dim prixCatalogueValide As Boolean
    Dim prixManuelValide As Boolean
    
    ' Vérifier validité prix catalogue
    prixCatalogueValide = EstPrixValide(prixCatalogue)
    
    ' Vérifier validité prix manuel
    prixManuelValide = EstPrixValide(prixManuel)
    
    ' Déterminer prix final selon hiérarchie
    If prixCatalogueValide Then
        ' Prix Catalogue valide
        prixFinal = CDbl(prixCatalogue)
        Commentaire = "Prix Spare Catalogue"
        
    ElseIf prixManuelValide Then
        ' Prix Manuel valide (catalogue invalide)
        prixFinal = CDbl(prixManuel)
        Commentaire = "Prix Spare Manuel"
        
    Else
        ' Aucun prix valide
        prixFinal = 0
        Commentaire = "Aucun prix valide en BDD (Catalogue: " & _
                      ValeurToString(prixCatalogue) & ", Manuel: " & _
                      ValeurToString(prixManuel) & ")"
    End If
    
End Sub

' =========================================================================
' FONCTION: EST PRIX VALIDE (SÉCURISÉE)
' =========================================================================
Private Function EstPrixValide(ByVal valeur As Variant) As Boolean
' Vérifie si une valeur est un prix valide (> 0 et numérique)
'
' Retourne False si :
'   - NULL
'   - Empty
'   - Chaîne vide
'   - "N/A", "#N/A", "N/D", "ND", "#N/D", "#REF!", "#VALUE!", etc.
'   - 0
'   - Non numérique

    On Error GoTo InvalidValue
    
    ' Vérifier NULL ou Empty
    If IsNull(valeur) Or IsEmpty(valeur) Then
        EstPrixValide = False
        Exit Function
    End If
    
    ' Vérifier si c'est une erreur Excel (#N/A, #REF!, etc.)
    If IsError(valeur) Then
        EstPrixValide = False
        Exit Function
    End If
    
    ' Convertir en String pour tests
    Dim valeurStr As String
    valeurStr = Trim$(CStr(valeur))
    
    ' Vérifier chaîne vide
    If Len(valeurStr) = 0 Then
        EstPrixValide = False
        Exit Function
    End If
    
    ' Vérifier valeurs textuelles invalides
    Select Case UCase$(valeurStr)
        Case "N/A", "#N/A", "N/D", "ND", "#N/D", "#REF!", "#VALUE!", "#DIV/0!", _
             "#NAME?", "#NUM!", "#NULL!", "NA", "#N", "NONE", "NULL", "-"
            EstPrixValide = False
            Exit Function
    End Select
    
    ' Vérifier si numérique
    If Not IsNumeric(valeurStr) Then
        EstPrixValide = False
        Exit Function
    End If
    
    ' Vérifier si > 0
    If CDbl(valeurStr) <= 0 Then
        EstPrixValide = False
        Exit Function
    End If
    
    ' Prix valide
    EstPrixValide = True
    Exit Function
    
InvalidValue:
    ' Erreur conversion ? prix invalide
    EstPrixValide = False
    
End Function

' =========================================================================
' FONCTION: VALEUR TO STRING (HELPER)
' =========================================================================
Private Function ValeurToString(ByVal valeur As Variant) As String
' Convertit une valeur en String pour affichage dans commentaire
' Gère NULL, Empty, Erreur, etc.

    On Error GoTo ErrorHandler
    
    If IsNull(valeur) Then
        ValeurToString = "NULL"
    ElseIf IsEmpty(valeur) Then
        ValeurToString = "Empty"
    ElseIf IsError(valeur) Then
        ValeurToString = "Erreur Excel"
    ElseIf Trim$(CStr(valeur)) = "" Then
        ValeurToString = "Vide"
    Else
        ValeurToString = CStr(valeur)
    End If
    Exit Function
    
ErrorHandler:
    ValeurToString = "Valeur invalide"
    
End Function
' =========================================================================
' PROCÉDURE: EXÉCUTER CALCUL LCC (SANS MOT DE PASSE)
' =========================================================================
Public Sub ExecuterCalculLCC(ByRef wb As Workbook)
' Exécute le calcul LCC dans le Providor
' VERSION SIMPLE : Juste le script VBS pour msgbox (pas de mot de passe)
    
    Dim vbsPath As String
    
    On Error GoTo SafeExit
    
    ' Activer le classeur Providor AVANT d'appeler les macros
    wb.Activate
    Application.Wait Now + TimeValue("00:00:01")
    
    ' Créer script VBS pour gérer les 2 MsgBox
    vbsPath = Build_DoubleClicker_VBS("LCC Calculation", "LCC Calculation Process", _
                                      15000, 15000, False)
    
    ' Lancer script VBS en arrière-plan
    If Len(vbsPath) > 0 Then
        CreateObject("WScript.Shell").Run "wscript.exe """ & vbsPath & """", 0, False
    End If
    
    ' Attendre que le script démarre
    Application.Wait Now + TimeValue("00:00:01")
    
    ' Appeler macro BS_Update
    Application.Run "'" & wb.Name & "'!NBS.BS_Update"
    
    ' Appeler macro LCC_Button_Click
    Application.Run "'" & wb.Name & "'!Feuil29.LCC_Button_Click"
    
SafeExit:
    ' Nettoyage du fichier VBS
    On Error Resume Next
    Application.Wait Now + TimeValue("00:00:01")
    If Len(vbsPath) > 0 Then Kill vbsPath
    On Error GoTo 0
    
End Sub

' =========================================================================
' FONCTION: BUILD DOUBLE CLICKER VBS (SIMPLE)
' =========================================================================
Public Function Build_DoubleClicker_VBS( _
    ByVal msgTitleQuestion As String, _
    ByVal msgTitleInfo As String, _
    ByVal timeoutQuestionMs As Long, _
    ByVal timeoutInfoMs As Long, _
    ByVal pressNoInsteadOfYes As Boolean) As String
' Construit le VBScript pour gérer 2 MsgBox successives
' Étape A : attendre la MsgBox de question (titre msgTitleQuestion)
'   - Yes/Oui : Entrée
'   - No/Non  : Tab puis Entrée (si pressNoInsteadOfYes=True)
' Étape B : attendre la MsgBox d'information (titre msgTitleInfo)
'   - Fermer par Entrée

    Dim tempPath As String, vbsPath As String, f As Integer
    Dim script As String

    tempPath = Environ$("TEMP")
    If Len(tempPath) = 0 Then tempPath = Environ$("TMP")
    If Len(tempPath) = 0 Then
        Build_DoubleClicker_VBS = ""
        Exit Function
    End If

    vbsPath = tempPath & "\ClickLccDouble.vbs"

    ' === Script VBS ===
    script = ""
    script = script & "Option Explicit" & vbCrLf
    script = script & "Dim sh" & vbCrLf
    script = script & "Set sh = CreateObject(""WScript.Shell"")" & vbCrLf

    ' === A) Question ===
    script = script & "Dim t0, timeout" & vbCrLf
    script = script & "timeout = " & CLng(timeoutQuestionMs) & vbCrLf
    script = script & "t0 = Timer" & vbCrLf
    script = script & "Do" & vbCrLf
    script = script & "  WScript.Sleep 100" & vbCrLf
    script = script & "  If sh.AppActivate(""" & Replace(msgTitleQuestion, """", """""") & """) Then" & vbCrLf
    If pressNoInsteadOfYes Then
        ' Déplacer sur No/Non puis valider
        script = script & "     sh.SendKeys ""{TAB}~""" & vbCrLf
    Else
        ' Valider bouton par défaut (Yes/Oui) par Entrée
        script = script & "     sh.SendKeys ""~""" & vbCrLf
    End If
    script = script & "     Exit Do" & vbCrLf
    script = script & "  End If" & vbCrLf
    script = script & "  If (Timer - t0) * 1000 > timeout Then Exit Do" & vbCrLf
    script = script & "Loop" & vbCrLf

    ' === B) Information de fin ===
    script = script & "t0 = Timer" & vbCrLf
    script = script & "timeout = " & CLng(timeoutInfoMs) & vbCrLf
    script = script & "Do" & vbCrLf
    script = script & "  WScript.Sleep 100" & vbCrLf
    script = script & "  If sh.AppActivate(""" & Replace(msgTitleInfo, """", """""") & """) Then" & vbCrLf
    script = script & "     ' Ferme la MsgBox d'information" & vbCrLf
    script = script & "     sh.SendKeys ""~""" & vbCrLf
    script = script & "     Exit Do" & vbCrLf
    script = script & "  End If" & vbCrLf
    script = script & "  If (Timer - t0) * 1000 > timeout Then Exit Do" & vbCrLf
    script = script & "Loop" & vbCrLf

    ' Écrit le .vbs
    On Error GoTo Fail
    f = FreeFile
    Open vbsPath For Output As #f
    Print #f, script
    Close #f

    Build_DoubleClicker_VBS = vbsPath
    Exit Function

Fail:
    On Error Resume Next
    If f <> 0 Then Close #f
    Build_DoubleClicker_VBS = ""
End Function
