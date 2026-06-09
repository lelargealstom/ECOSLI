Attribute VB_Name = "B_ATM_Commun"
Public Sub RecupererDonneesBDD_ATM(ByVal versionATM As Integer)
' R�cup�re les donn�es de la BDD pour ATM V1 ou V2
'
' Param�tres:
'   versionATM: 1 pour ATM V1, 2 pour ATM V2
'
' Diff�rences:
'   - V1: Requ�te Taches_ATM$ SANS filtre HARD/SOFT
'   - V2: Requ�te Taches_ATM$ AVEC filtre WHERE [HARD/SOFT]='HARD'
'
' Variables globales remplies:
'   - Infos_tab (tableau 2D, 17 colonnes)
'   - Taches_tab (tableau 2D, 15 colonnes)
'   - Auteur (String)
'   - Nv_Maint_Projet (String)

    Dim STRSEL As String
    
    ' =====================================================================
    ' REQU�TE 1: Info_ATM$ (17 champs) - IDENTIQUE V1 ET V2
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
         "[Temps g�n�rique de la t�che_INFOATM], " & _
         "[Temps g�n�rique de la t�che SOFTWARE_INFOATM], " & _
         "[Temps analytique cumul� de la tache_INFOATM], " & _
         "[NbLanguePopulation_INFOATM] " & _
         "FROM [Info_ATM$]"
    
    Call open_REQSEL(STRSEL)
    Infos_tab = RecordsetToArray(REQSEL)
    
    ' =====================================================================
    ' REQU�TE 2: Taches_ATM$ (15 champs) - DIFF�RENCE V1/V2
    ' =====================================================================
    ' V1: PAS de filtre HARD/SOFT (g�re les deux types dans le code)
    ' V2: FILTRE WHERE [HARD/SOFT]='HARD' uniquement
    
    If versionATM = 1 Then
        ' ATM V1: R�cup�rer TOUTES les t�ches (HARD et SOFT)
        STRSEL = "SELECT [DTR_TACHEATM], [Langue_TACHEATM], [Rang_TACHEATM], [HARD/SOFT], " & _
                 "[Categorie_TACHEATM], [Niveau qualif_TACHEATM], [Description_TACHEATM], [Crit�resDimensionnels_TACHEATM], " & _
                 "[OutillageSp�cifique_TACHEATM], [OutillageStandard_TACHEATM], [Pi�cedeRechangeConsommable_TACHEATM], [NbOp�rateurs_TACHEATM], " & _
                 "[TempsEstim�Heures_TACHEATM] " & _
                 "FROM [Taches_ATM$]"
    ElseIf versionATM = 2 Then
        ' ATM V2: Filtrer HARD uniquement
        STRSEL = "SELECT [DTR_TACHEATM], [Langue_TACHEATM], [Rang_TACHEATM], [HARD/SOFT], " & _
                 "[Categorie_TACHEATM], [Niveau qualif_TACHEATM], [Description_TACHEATM], [Crit�resDimensionnels_TACHEATM], " & _
                 "[OutillageSp�cifique_TACHEATM], [OutillageStandard_TACHEATM], [Pi�cedeRechangeConsommable_TACHEATM], [NbOp�rateurs_TACHEATM], " & _
                 "[TempsEstim�Heures_TACHEATM] " & _
                 "FROM [Taches_ATM$] WHERE [HARD/SOFT] = 'HARD'"
    Else
        ' Version inconnue
        MsgBox "Version ATM inconnue: " & versionATM, vbCritical
        Exit Sub
    End If
    
    Call open_REQSEL(STRSEL)
    Taches_tab = RecordsetToArray(REQSEL)

    ' =====================================================================
    ' INDEX DTR -> LIGNE (acces O(1) en boucle, gain ~5-20s sur 50 DTR)
    ' =====================================================================
    Call IndexerTachesEtInfos

    ' =====================================================================
    ' REQU�TE 3: RASL Projet - IDENTIQUE V1 ET V2
    ' =====================================================================
    STRSEL = "SELECT [RASL_PROJET SLI] FROM [Projet_SLI$] WHERE [Nom projet_PROJET SLI] = '" & Code_Projet & "'"
    Call open_REQSEL(STRSEL)
    
    If IsNull(REQSEL.Fields(0).Value) Then
        Auteur = ""
    Else
        Auteur = REQSEL.Fields(0).Value
    End If
    
    ' =====================================================================
    ' REQU�TE 4: Niveau Maintenance - IDENTIQUE V1 ET V2
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
' PROCEDURE: IndexerTachesEtInfos
' Construit les Dictionaries dictTachesPMB et dictInfosPMJ pour acceder
' aux lignes de Taches_tab / Infos_tab en O(1) au lieu de O(N) dans
' RechercherPlagesDTR_V1/_V2.
' Appele apres chaque RecupererDonneesBDD_ATM (les tableaux changent).
' =========================================================================
Public Sub IndexerTachesEtInfos()
    Dim j As Long
    Dim cle As String

    ' Reset
    Set dictTachesPMB = CreateObject("Scripting.Dictionary")
    dictTachesPMB.CompareMode = vbTextCompare
    Set dictInfosPMJ = CreateObject("Scripting.Dictionary")
    dictInfosPMJ.CompareMode = vbTextCompare

    ' Index Taches_tab : cle = DTR|Langue -> PMB (1ere occurrence).
    ' Les lignes suivantes du meme DTR restent contigues dans le tableau ;
    ' RechercherPlagesDTR_V1/_V2 trouve PMH a partir de PMB par continuation.
    If IsArray(Taches_tab) Then
        For j = LBound(Taches_tab, 1) + 1 To UBound(Taches_tab, 1)
            cle = CStr(Taches_tab(j, Taches_type.dtr)) & "|" & _
                  CStr(Taches_tab(j, Taches_type.langue))
            If Not dictTachesPMB.Exists(cle) Then dictTachesPMB.Add cle, j
        Next j
    End If

    ' Index Infos_tab : cle = DTR -> PMJ (1ere occurrence).
    If IsArray(Infos_tab) Then
        For j = LBound(Infos_tab, 1) + 1 To UBound(Infos_tab, 1)
            cle = CStr(Infos_tab(j, Info_type.dtr))
            If Not dictInfosPMJ.Exists(cle) Then dictInfosPMJ.Add cle, j
        Next j
    End If

    Debug.Print "IndexerTachesEtInfos : " & dictTachesPMB.count & _
                " entrees Taches, " & dictInfosPMJ.count & " entrees Infos."
End Sub

' =========================================================================
' FONCTION: S�LECTION R�PERTOIRE
' =========================================================================
Public Function SelectRepertoire(ByRef chemin As String) As Boolean
    Dim Repertoire As FileDialog
    
    MsgBox "Veuillez s�lectionner le r�pertoire o� vous souhaitez d�poser les fichiers ATM pour le projet : " & Projet_Abrev, vbOKOnly
    
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
' FONCTION UTILITAIRE: R�CUP�RER TEMPS AVEC GESTION NULL
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
