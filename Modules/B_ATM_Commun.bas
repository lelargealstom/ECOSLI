Attribute VB_Name = "B_ATM_Commun"
Public Sub RecupererDonneesBDD_ATM(ByVal versionATM As Integer)
' Récupère les données de la BDD pour ATM V1 ou V2
'
' Paramètres:
'   versionATM: 1 pour ATM V1, 2 pour ATM V2
'
' Différences:
'   - V1: Requête Taches_ATM$ SANS filtre HARD/SOFT
'   - V2: Requête Taches_ATM$ AVEC filtre WHERE [HARD/SOFT]='HARD'
'
' Variables globales remplies:
'   - Infos_tab (tableau 2D, 17 colonnes)
'   - Taches_tab (tableau 2D, 15 colonnes)
'   - Auteur (String)
'   - Nv_Maint_Projet (String)

    Dim STRSEL As String
    
    ' =====================================================================
    ' REQUÊTE 1: Info_ATM$ (17 champs) - IDENTIQUE V1 ET V2
    ' =====================================================================
    STRSEL = "SELECT " & _
         "[DTR_INFOATM], " & _
         "[DescriptifATMFR_INFOATM], " & _
         "[LocalisationATM_INFOATM], " & _
         "[PLANATM_INFOATM], " & _
         "[CommentaireATM_INFOATM], " & _
         "[Version_INFOATM], " & _
         "[DescriptifATMEN_INFOATM], " & _
         "[LocalisationATMEN_INFOATM], " & _
         "[CommentairesATMEN_INFOATM], " & _
         "[Reparabilite_INFOATM], " & _
         "[HARD/SOFT], " & _
         "[Poids_INFOATM], " & _
         "[Temps générique de la tâche_INFOATM], " & _
         "[Temps générique de la tâche SOFTWARE_INFOATM], " & _
         "[Temps analytique cumulé de la tache_INFOATM], " & _
         "[NbLanguePopulation_INFOATM] " & _
         "FROM [Info_ATM$]"
    
    Call open_REQSEL(STRSEL)
    Infos_tab = RecordsetToArray(REQSEL)
    
    ' =====================================================================
    ' REQUÊTE 2: Taches_ATM$ (15 champs) - DIFFÉRENCE V1/V2
    ' =====================================================================
    ' V1: PAS de filtre HARD/SOFT (gère les deux types dans le code)
    ' V2: FILTRE WHERE [HARD/SOFT]='HARD' uniquement
    
    If versionATM = 1 Then
        ' ATM V1: Récupérer TOUTES les tâches (HARD et SOFT)
        STRSEL = "SELECT [DTR_TACHEATM], [Langue_TACHEATM], [Rang_TACHEATM], [HARD/SOFT], " & _
                 "[Categorie_TACHEATM], [Niveau qualif_TACHEATM], [Description_TACHEATM], [CritèresDimensionnels_TACHEATM], " & _
                 "[OutillageSpécifique_TACHEATM], [OutillageStandard_TACHEATM], [PiècedeRechangeConsommable_TACHEATM], [NbOpérateurs_TACHEATM], " & _
                 "[TempsEstiméHeures_TACHEATM] " & _
                 "FROM [Taches_ATM$]"
    ElseIf versionATM = 2 Then
        ' ATM V2: Filtrer HARD uniquement
        STRSEL = "SELECT [DTR_TACHEATM], [Langue_TACHEATM], [Rang_TACHEATM], [HARD/SOFT], " & _
                 "[Categorie_TACHEATM], [Niveau qualif_TACHEATM], [Description_TACHEATM], [CritèresDimensionnels_TACHEATM], " & _
                 "[OutillageSpécifique_TACHEATM], [OutillageStandard_TACHEATM], [PiècedeRechangeConsommable_TACHEATM], [NbOpérateurs_TACHEATM], " & _
                 "[TempsEstiméHeures_TACHEATM] " & _
                 "FROM [Taches_ATM$] WHERE [HARD/SOFT] = 'HARD'"
    Else
        ' Version inconnue
        MsgBox "Version ATM inconnue: " & versionATM, vbCritical
        Exit Sub
    End If
    
    Call open_REQSEL(STRSEL)
    Taches_tab = RecordsetToArray(REQSEL)
    
    ' =====================================================================
    ' REQUÊTE 3: RASL Projet - IDENTIQUE V1 ET V2
    ' =====================================================================
    STRSEL = "SELECT [RASL_PROJET SLI] FROM [Projet_SLI$] WHERE [Nom projet_PROJET SLI] = '" & Code_Projet & "'"
    Call open_REQSEL(STRSEL)
    
    If IsNull(REQSEL.Fields(0).Value) Then
        Auteur = ""
    Else
        Auteur = REQSEL.Fields(0).Value
    End If
    
    ' =====================================================================
    ' REQUÊTE 4: Niveau Maintenance - IDENTIQUE V1 ET V2
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
' FONCTION: SÉLECTION RÉPERTOIRE
' =========================================================================
Public Function SelectRepertoire(ByRef chemin As String) As Boolean
    Dim Repertoire As FileDialog
    
    MsgBox "Veuillez sélectionner le répertoire où vous souhaitez déposer les fichiers ATM pour le projet : " & Projet_Abrev, vbOKOnly
    
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
' FONCTION UTILITAIRE: RÉCUPÉRER TEMPS AVEC GESTION NULL
' =========================================================================
Public Function GetTempsFromBDD(ByVal valeurBDD As Variant) As Double
' Retourne le temps en Double ou 0 si NULL/vide
' Permet de centraliser la gestion des valeurs NULL de la BDD
    
    If IsNull(valeurBDD) Then
        GetTempsFromBDD = 0
    ElseIf IsEmpty(valeurBDD) Then
        GetTempsFromBDD = 0
    ElseIf valeurBDD = "" Then
        GetTempsFromBDD = 0
    Else
        GetTempsFromBDD = CDbl(valeurBDD)
    End If
End Function
