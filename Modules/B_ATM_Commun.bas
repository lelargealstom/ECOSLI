Attribute VB_Name = "B_ATM_Commun"
Public Sub RecupererDonneesBDD_ATM(ByVal versionATM As Integer)
' Recupere les donnees de la BDD pour ATM V1 ou V2
'
' Parametres:
'   versionATM: 1 pour ATM V1, 2 pour ATM V2
'
' Differences:
'   - V1: Requete Taches_ATM$ SANS filtre HARD/SOFT
'   - V2: Requete Taches_ATM$ AVEC filtre WHERE [HARD/SOFT]='HARD'
'
' Variables globales remplies:
'   - Infos_tab (tableau 2D, 17 colonnes)
'   - Taches_tab (tableau 2D, 15 colonnes)
'   - Auteur (String)
'   - Nv_Maint_Projet (String)

    Dim STRSEL As String
    
    ' =====================================================================
    ' REQUETE 1: Info_ATM$ (17 champs) - IDENTIQUE V1 ET V2
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
    ' REQUETE 2: Taches_ATM$ (15 champs) - DIFFERENCES V1/V2
    ' =====================================================================
    ' V1: PAS de filtre HARD/SOFT (gère les deux types dans le code)
    ' V2: FILTRE WHERE [HARD/SOFT]='HARD' uniquement
    
    If versionATM = 1 Then
        ' ATM V1: RRecuperer TOUTES les taches (HARD et SOFT)
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

    ' Note : IndexerTachesEtInfos n'est PAS appelee ici. Elle doit l'etre
    ' apres Cadrage_Taches_ATM et Cadrage_Infos_ATM (qui initialisent
    ' Taches_type et Info_type) - donc depuis ATM_V1 et ATM_V2 directement.

    ' =====================================================================
    ' REQUETE 3: RASL Projet - IDENTIQUE V1 ET V2
    ' =====================================================================
    STRSEL = "SELECT [RASL_PROJET SLI] FROM [Projet_SLI$] WHERE [Nom projet_PROJET SLI] = '" & Code_Projet & "'"
    Call open_REQSEL(STRSEL)
    
    If IsNull(REQSEL.Fields(0).Value) Then
        Auteur = ""
    Else
        Auteur = REQSEL.Fields(0).Value
    End If
    
    ' =====================================================================
    ' REQUETE 4: Niveau Maintenance - IDENTIQUE V1 ET V2
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
' Construit dictTachesPMB et dictInfosPMJ pour acceder aux lignes en O(1)
' dans RechercherPlagesDTR_V1/V2. Trace les bornes du tableau et les 3
' premieres cles indexees pour diagnostic (Debug.Print).
' =========================================================================
Public Sub IndexerTachesEtInfos()
    Dim j As Long
    Dim cle As String
    Dim lb As Long, ub As Long
    Dim nbTraces As Long

    Set dictTachesPMB = CreateObject("Scripting.Dictionary")
    dictTachesPMB.CompareMode = vbTextCompare
    Set dictInfosPMJ = CreateObject("Scripting.Dictionary")
    dictInfosPMJ.CompareMode = vbTextCompare

    ' Index Taches_tab : cle = DTR|Langue -> PMB.
    If IsArray(Taches_tab) Then
        lb = LBound(Taches_tab, 1)
        ub = UBound(Taches_tab, 1)
        Debug.Print "IndexerTachesEtInfos : Taches_tab LBound=" & lb & " UBound=" & ub & _
                    " (Taches_type.dtr=" & Taches_type.dtr & " Taches_type.langue=" & Taches_type.langue & ")"
        nbTraces = 0
        For j = lb + 1 To ub
            cle = CStr(Taches_tab(j, Taches_type.dtr)) & "|" & _
                  CStr(Taches_tab(j, Taches_type.langue))
            If Not dictTachesPMB.Exists(cle) Then
                dictTachesPMB.Add cle, j
                If nbTraces < 3 Then
                    Debug.Print "  Taches[" & j & "] cle = |" & cle & "|"
                    nbTraces = nbTraces + 1
                End If
            End If
        Next j
    Else
        Debug.Print "IndexerTachesEtInfos : Taches_tab n'est pas un Array, skip."
    End If

    ' Index Infos_tab : cle = DTR -> PMJ.
    If IsArray(Infos_tab) Then
        lb = LBound(Infos_tab, 1)
        ub = UBound(Infos_tab, 1)
        Debug.Print "IndexerTachesEtInfos : Infos_tab LBound=" & lb & " UBound=" & ub & _
                    " (Info_type.dtr=" & Info_type.dtr & ")"
        nbTraces = 0
        For j = lb + 1 To ub
            cle = CStr(Infos_tab(j, Info_type.dtr))
            If Not dictInfosPMJ.Exists(cle) Then
                dictInfosPMJ.Add cle, j
                If nbTraces < 3 Then
                    Debug.Print "  Infos[" & j & "] cle = |" & cle & "|"
                    nbTraces = nbTraces + 1
                End If
            End If
        Next j
    Else
        Debug.Print "IndexerTachesEtInfos : Infos_tab n'est pas un Array, skip."
    End If

    Debug.Print "IndexerTachesEtInfos : " & dictTachesPMB.count & _
                " entrees Taches, " & dictInfosPMJ.count & " entrees Infos."
End Sub

' =========================================================================
' FONCTION: RecupererTachesDTR_BdD
' Renvoie un Recordset des taches du DTR pour la langue ATM donnee.
'   - versionATM = 1 : toutes les taches (HARD + SOFT)
'   - versionATM = 2 : taches HARD uniquement (filtre supplementaire)
' Recordset vide (rs.EOF = True) si DTR absent en BdD ou langue non
' disponible pour ce DTR : l'appelant decide quoi faire.
'
' Pourquoi pas open_REQSEL (REQSEL global) : on a besoin d'avoir 2
' Recordsets ouverts en parallele (taches + infos) pour le meme DTR
' sans que le second ecrase le premier. CreateObject + Base = nouvelle
' instance independante.
' =========================================================================
Public Function RecupererTachesDTR_BdD(ByVal dtr As String, _
                                        ByVal langueATM As String, _
                                        ByVal versionATM As Integer) As Object
    Dim rs As Object
    Dim sSQL As String

    sSQL = "SELECT [DTR_TACHEATM], [Langue_TACHEATM], [Rang_TACHEATM], [HARD/SOFT], " & _
           "[Categorie_TACHEATM], [Niveau qualif_TACHEATM], [Description_TACHEATM], " & _
           "[CritèresDimensionnels_TACHEATM], [OutillageSpécifique_TACHEATM], " & _
           "[OutillageStandard_TACHEATM], [PiècedeRechangeConsommable_TACHEATM], " & _
           "[NbOpérateurs_TACHEATM], [TempsEstiméHeures_TACHEATM] " & _
           "FROM [Taches_ATM$] " & _
           "WHERE [DTR_TACHEATM] = '" & Replace(dtr, "'", "''") & "' " & _
           "AND [Langue_TACHEATM] = '" & Replace(langueATM, "'", "''") & "'"

    If versionATM = 2 Then sSQL = sSQL & " AND [HARD/SOFT] = 'HARD'"

    Set rs = CreateObject("ADODB.Recordset")
    rs.CursorLocation = 3                ' adUseClient
    rs.Open sSQL, Base, 3, 1, 1          ' adOpenStatic, adLockReadOnly, adCmdText

    Set RecupererTachesDTR_BdD = rs
End Function

' =========================================================================
' FONCTION: RecupererInfoATM_DTR
' Renvoie un Recordset de l'info ATM du DTR (1 ligne attendue).
' Recordset vide si DTR absent en BdD : l'appelant decide.
' =========================================================================
Public Function RecupererInfoATM_DTR(ByVal dtr As String) As Object
    Dim rs As Object
    Dim sSQL As String

    sSQL = "SELECT [DTR_INFOATM], [DescriptifATMFR_INFOATM], [LocalisationATM_INFOATM], " & _
           "[PLANATM_INFOATM], [CommentaireATM_INFOATM], [Version_INFOATM], " & _
           "[DescriptifATMEN_INFOATM], [LocalisationATMEN_INFOATM], [CommentairesATMEN_INFOATM], " & _
           "[Reparabilite_INFOATM], [HARD/SOFT], [Poids_INFOATM], " & _
           "[Temps générique de la tâche_INFOATM], [Temps générique de la tâche SOFTWARE_INFOATM], " & _
           "[Temps analytique cumulé de la tache_INFOATM], [NbLanguePopulation_INFOATM] " & _
           "FROM [Info_ATM$] " & _
           "WHERE [DTR_INFOATM] = '" & Replace(dtr, "'", "''") & "'"

    Set rs = CreateObject("ADODB.Recordset")
    rs.CursorLocation = 3
    rs.Open sSQL, Base, 3, 1, 1

    Set RecupererInfoATM_DTR = rs
End Function

' =========================================================================
' FONCTION: SLECTION RPERTOIRE
' =========================================================================
Public Function SelectRepertoire(ByRef chemin As String) As Boolean
    Dim Repertoire As FileDialog
    
    MsgBox "Veuillez selectionner le repertoire ou vous souhaitez deposer les fichiers ATM pour le projet : " & Projet_Abrev, vbOKOnly
    
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
' FONCTION UTILITAIRE: RECUPERER TEMPS AVEC GESTION NULL
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