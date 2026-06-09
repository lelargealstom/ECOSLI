Attribute VB_Name = "G_SIRUS_Generation"
'--------------------------------------------------------------------------------------------------------
'
'                                       Generateur Livrables SLI
'                                       Module G_SIRUS_Generation
'                                       Developpement : LE LARGE Thomas (LGM/Alstom)
'                                       SIRUS - Phase 1 : generateur de feuille d'analyse
'                                       (creation vierge ou modification pre-remplie depuis BdD).
'                                       Juin 2026
'
'--------------------------------------------------------------------------------------------------------
Option Explicit
Option Compare Text

' =========================================================================
' CONSTANTES
' =========================================================================
Public Const SIRUS_OP_CREER    As String = "Creer"
Public Const SIRUS_OP_MODIFIER As String = "Modifier"

Private Const SIRUS_LIGNE_ENTETE_HUMAIN  As Long = 1
Private Const SIRUS_LIGNE_ENTETE_BDD     As Long = 2
Private Const SIRUS_LIGNE_DONNEES_DEBUT  As Long = 3
Private Const SIRUS_NB_LIGNES_VALIDATION As Long = 1000  ' lignes couvertes par les listes deroulantes

Private Const COULEUR_ENTETE             As Long = 14606303   ' bleu tres clair RGB(220,230,241)
Private Const COULEUR_AUTO_HERITE        As Long = 14211288   ' gris clair RGB(216,216,216)
Private Const COULEUR_ENTETE_BDD         As Long = 9868950    ' gris fonce RGB(150,150,150)

' =========================================================================
' Connexion BdD SIRUS
' =========================================================================
' SIRUS lit (et a terme ecrit indirectement via ATMAP) la BdD de TRAVAIL,
' pas la BdD de PROD. La chaine est construite en remplacant 'UserProfile'
' par le vrai chemin utilisateur via Environ$.
Private Function ChaineConnexionSIRUS() As String
    ChaineConnexionSIRUS = Replace(CONNEXION_EXCEL_BDD_TRAVAIL, _
                                    "UserProfile", Environ$("UserProfile"))
End Function

' =========================================================================
' SIRUS_GenererFeuilleAnalyse
' =========================================================================
' Genere dans ActiveWorkbook (= classeur projet de l'utilisateur) la feuille
' "SIRUS_<Table>_<Operation>" pour le couple (table, operation) demande.
'
' Parametres :
'   sTable    : "Produit" / "InfoATM" / "TachesATM" / "EMMST" / "ProjetSLI" / "ProjetASL"
'   sOperation: "Creer" ou "Modifier" (utiliser les constantes SIRUS_OP_*).
'   vCles     : tableau de cles a pre-remplir (DTR pour la plupart, Nom
'               projet pour ProjetSLI). Ignore si sOperation = "Creer".
'
' Comportement :
'   - Si la feuille existe deja : demande confirmation a l'utilisateur.
'   - Ligne 1 = libelle Champ humain, ligne 2 = ColonneBdD masquee.
'   - Listes deroulantes (xlValidateList) sur les TypeChamp = "Liste" ou "Booleen".
'   - Champs Source = "Auto" ou "Herite" : fond gris (lecture seule visuelle ;
'     verrouillage reel reporte Phase 2).
'   - Operation "Modifier" : ouvre ADODB, SELECT par cle, remplit ligne par ligne.
'     Cle absente en base : colonne Statut = "Cle non trouvee".
'
Public Sub SIRUS_GenererFeuilleAnalyse(ByVal sTable As String, _
                                       ByVal sOperation As String, _
                                       Optional ByRef vCles As Variant)

    Dim wb As Workbook
    Dim wsCible As Worksheet
    Dim nomFeuille As String
    Dim col As Collection
    Dim resp As VbMsgBoxResult

    On Error GoTo Sortie

    ' -- 1. Validation arguments ----------------------------
    If LenB(sTable) = 0 Then
        MsgBox "SIRUS : nom de table manquant.", vbCritical
        Exit Sub
    End If
    If sOperation <> SIRUS_OP_CREER And sOperation <> SIRUS_OP_MODIFIER Then
        MsgBox "SIRUS : operation invalide ('" & sOperation & "'). " & _
               "Doit etre 'Creer' ou 'Modifier'.", vbCritical
        Exit Sub
    End If

    ' -- 2. Verifier le classeur projet actif ----------------
    ' Note : on ne bloque PAS sur "wb Is ThisWorkbook" car en mode dev,
    ' Thomas edite les macros dans ECOSLI.xlsm (le template projet lui-meme,
    ' qui contient DATA_ASL). On valide plutot sur la presence de la feuille
    ' DATA_ASL (signal qu'on est sur un classeur projet legitime). Si elle
    ' manque, on previent l'utilisateur et il decide.
    Set wb = ActiveWorkbook
    If wb Is Nothing Then
        MsgBox "Aucun classeur actif. Ouvrir d'abord le fichier projet.", vbExclamation
        Exit Sub
    End If
    If Not FeuilleExiste(wb, "DATA_ASL") Then
        resp = MsgBox("Le classeur actif '" & wb.Name & "' ne contient pas de feuille DATA_ASL " & _
                      "(habituellement le marqueur d'un classeur projet SLI)." & vbCrLf & vbCrLf & _
                      "Voulez-vous quand meme generer la feuille SIRUS dans ce classeur ?", _
                      vbYesNo + vbQuestion + vbDefaultButton2, _
                      "SIRUS - Classeur cible inhabituel")
        If resp <> vbYes Then Exit Sub
    End If

    ' -- 3. Lire le config -----------------------------------
    Set col = SIRUS_LireConfig(sTable)
    If col.Count = 0 Then
        MsgBox "Aucun champ defini dans SIRUS_Config pour la table '" & sTable & "'. " & _
               "Lancer SIRUS_InitConfigSheet d'abord.", vbCritical
        Exit Sub
    End If

    ' -- 4. Nom de la feuille + confirmation si existante ---
    nomFeuille = "SIRUS_" & sTable & "_" & sOperation
    If Len(nomFeuille) > 31 Then nomFeuille = Left$(nomFeuille, 31)  ' limite Excel

    If FeuilleExiste(wb, nomFeuille) Then
        resp = MsgBox("La feuille '" & nomFeuille & "' existe deja dans le classeur '" & _
                      wb.Name & "'." & vbCrLf & vbCrLf & _
                      "Voulez-vous l'ecraser ?", _
                      vbYesNo + vbQuestion + vbDefaultButton2, _
                      "SIRUS - Generation de feuille")
        If resp <> vbYes Then Exit Sub
        Application.DisplayAlerts = False
        wb.Worksheets(nomFeuille).Delete
        Application.DisplayAlerts = True
    End If

    ' -- 5. Creer la feuille --------------------------------
    Set wsCible = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    wsCible.Name = nomFeuille

    ' -- 6. Entetes (humain L1 + ColonneBdD L2 masquee) -----
    Call PoserEntetes(wsCible, col)

    ' -- 7. Pre-remplissage si Modifier ---------------------
    If sOperation = SIRUS_OP_MODIFIER Then
        If IsMissing(vCles) Then
            MsgBox "Operation 'Modifier' : il faut fournir des cles (DTR ou Nom projet) " & _
                   "a pre-remplir. La feuille a ete creee vierge.", vbExclamation
        Else
            Call PreRemplirDepuisBdD(wsCible, sTable, col, vCles)
        End If
    End If

    ' -- 8. Listes deroulantes ------------------------------
    Call AppliquerListes(wsCible, col)

    ' -- 9. Griser les champs Auto / Herite -----------------
    Call GriserChampsAutoHerite(wsCible, col)

    ' -- 10. Aller a la cellule de saisie -------------------
    wsCible.Activate
    wsCible.Cells(SIRUS_LIGNE_DONNEES_DEBUT, 1).Select
    ActiveWindow.FreezePanes = False
    wsCible.Cells(SIRUS_LIGNE_DONNEES_DEBUT, 1).Select
    ActiveWindow.FreezePanes = True

    MsgBox "Feuille '" & nomFeuille & "' generee avec succes." & vbCrLf & _
           col.Count & " colonnes + colonne Statut.", vbInformation, "SIRUS"
    Exit Sub

Sortie:
    MsgBox "Erreur dans SIRUS_GenererFeuilleAnalyse : " & Err.Description, vbCritical
End Sub

' =========================================================================
' PoserEntetes
' =========================================================================
' Ligne 1 : libelles humains (Champ du config) + colonne "Statut" en bout.
' Ligne 2 : ColonneBdD techniques (masquee). Sert d'index BdD <-> colonne
'           pour le moteur d'analyse Phase 2 et la PJ ATMAP Phase 5.
Private Sub PoserEntetes(ws As Worksheet, col As Collection)

    Dim def As SIRUS_Champ_Config
    Dim i As Long
    Dim nbCol As Long

    i = 1
    For Each def In col
        ws.Cells(SIRUS_LIGNE_ENTETE_HUMAIN, i).Value = def.Champ
        ws.Cells(SIRUS_LIGNE_ENTETE_BDD, i).Value = def.ColonneBdD
        i = i + 1
    Next def
    nbCol = col.Count

    ' Statut en derniere colonne
    ws.Cells(SIRUS_LIGNE_ENTETE_HUMAIN, nbCol + 1).Value = "Statut"
    ws.Cells(SIRUS_LIGNE_ENTETE_BDD, nbCol + 1).Value = "_statut"

    ' Formatage ligne 1 (humain)
    With ws.Range(ws.Cells(SIRUS_LIGNE_ENTETE_HUMAIN, 1), _
                  ws.Cells(SIRUS_LIGNE_ENTETE_HUMAIN, nbCol + 1))
        .Font.Bold = True
        .Interior.Color = COULEUR_ENTETE
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With

    ' Formatage ligne 2 (BdD) + masquage
    With ws.Range(ws.Cells(SIRUS_LIGNE_ENTETE_BDD, 1), _
                  ws.Cells(SIRUS_LIGNE_ENTETE_BDD, nbCol + 1))
        .Font.Color = COULEUR_ENTETE_BDD
        .Font.Size = 8
        .Font.Italic = True
    End With
    ws.Rows(SIRUS_LIGNE_ENTETE_BDD).Hidden = True

    ' Largeur auto
    ws.Range(ws.Cells(SIRUS_LIGNE_ENTETE_HUMAIN, 1), _
             ws.Cells(SIRUS_LIGNE_ENTETE_HUMAIN, nbCol + 1)).EntireColumn.AutoFit

    ' Hauteur entete + figer
    ws.Rows(SIRUS_LIGNE_ENTETE_HUMAIN).RowHeight = 30
End Sub

' =========================================================================
' AppliquerListes
' =========================================================================
' Pour chaque champ TypeChamp = "Liste" ou "Booleen", ajoute une
' Validation.Add xlValidateList sur SIRUS_NB_LIGNES_VALIDATION lignes.
' Source de la liste : SIRUS_GetListe via SIRUS_ListeEnChaine.
Private Sub AppliquerListes(ws As Worksheet, col As Collection)

    Dim def As SIRUS_Champ_Config
    Dim i As Long
    Dim sFormula As String
    Dim plage As Range

    i = 1
    For Each def In col
        If def.TypeChamp = "Liste" Or def.TypeChamp = "Booleen" Then
            sFormula = SIRUS_ListeEnChaine(def.ListeSource)
            If LenB(sFormula) > 0 Then
                Set plage = ws.Range( _
                    ws.Cells(SIRUS_LIGNE_DONNEES_DEBUT, i), _
                    ws.Cells(SIRUS_LIGNE_DONNEES_DEBUT + SIRUS_NB_LIGNES_VALIDATION - 1, i))
                On Error Resume Next
                plage.Validation.Delete
                plage.Validation.Add Type:=xlValidateList, _
                                     AlertStyle:=xlValidAlertStop, _
                                     Operator:=xlBetween, _
                                     Formula1:=sFormula
                On Error GoTo 0
            Else
                Debug.Print "AppliquerListes : liste '" & def.ListeSource & _
                            "' vide ou inconnue pour le champ '" & def.Champ & "'."
            End If
        End If
        i = i + 1
    Next def
End Sub

' =========================================================================
' GriserChampsAutoHerite
' =========================================================================
' Fond gris clair sur les colonnes Source = "Auto" ou "Herite", pour
' signaler visuellement que ce sont des champs en lecture seule
' (heritage ou calcul automatique).
' Verrouillage cellule reel : reporte Phase 2 (avec protection feuille).
Private Sub GriserChampsAutoHerite(ws As Worksheet, col As Collection)

    Dim def As SIRUS_Champ_Config
    Dim i As Long
    Dim plage As Range

    i = 1
    For Each def In col
        If def.Source = "Auto" Or def.Source = "Herite" Then
            Set plage = ws.Range( _
                ws.Cells(SIRUS_LIGNE_DONNEES_DEBUT, i), _
                ws.Cells(SIRUS_LIGNE_DONNEES_DEBUT + SIRUS_NB_LIGNES_VALIDATION - 1, i))
            plage.Interior.Color = COULEUR_AUTO_HERITE
        End If
        i = i + 1
    Next def
End Sub

' =========================================================================
' PreRemplirDepuisBdD
' =========================================================================
' Pour Modifier : execute un SELECT par cle, remplit la feuille ligne par
' ligne. Cle = ColonneBdD du 1er champ du config (DTR pour la plupart des
' tables, Nom projet pour ProjetSLI).
' Champs avec ColonneBdD vide (ex. MTBF intrinseque info Produit dans
' ProjetASL) : ignores pour le SELECT, laisses vides cote feuille.
' Cle absente en BdD : ligne marquee dans la colonne Statut.
Private Sub PreRemplirDepuisBdD(ws As Worksheet, ByVal sTable As String, _
                                 col As Collection, ByRef vCles As Variant)

    Dim def As SIRUS_Champ_Config
    Dim cleColBdD As String
    Dim sNomTableSQL As String
    Dim sSelectCols As String
    Dim sSQL As String
    Dim i As Long
    Dim ligne As Long
    Dim cle As Variant
    Dim trouve As Boolean
    Dim nbOk As Long, nbKo As Long

    ' -- Cle = 1er champ du config -------------------------
    cleColBdD = ""
    For Each def In col
        cleColBdD = def.ColonneBdD
        Exit For
    Next def

    If LenB(cleColBdD) = 0 Then
        MsgBox "Cle BdD introuvable pour la table '" & sTable & _
               "' (1er champ du config sans ColonneBdD).", vbCritical
        Exit Sub
    End If

    ' -- Nom de la table SQL -------------------------------
    sNomTableSQL = NomTableSQL(sTable)
    If LenB(sNomTableSQL) = 0 Then
        MsgBox "Table SQL inconnue pour '" & sTable & "'.", vbCritical
        Exit Sub
    End If

    ' -- Construire la liste des colonnes a selectionner ---
    sSelectCols = ""
    For Each def In col
        If LenB(def.ColonneBdD) > 0 Then
            If LenB(sSelectCols) > 0 Then sSelectCols = sSelectCols & ", "
            sSelectCols = sSelectCols & "[" & def.ColonneBdD & "]"
        End If
    Next def

    ' -- Ouvrir la connexion -------------------------------
    On Error GoTo SortieBdD
    Call Open_CONN(ChaineConnexionSIRUS)

    ' -- Boucle par cle ------------------------------------
    ligne = SIRUS_LIGNE_DONNEES_DEBUT
    nbOk = 0
    nbKo = 0

    For Each cle In vCles
        sSQL = "SELECT " & sSelectCols & _
               " FROM [" & sNomTableSQL & "]" & _
               " WHERE [" & cleColBdD & "] = '" & Replace(CStr(cle), "'", "''") & "'"
        Debug.Print "PreRemplirDepuisBdD SQL : " & sSQL
        Call open_REQSEL(sSQL)

        trouve = False
        If Not (REQSEL Is Nothing) Then
            If Not REQSEL.EOF Then
                ' Remplir la ligne champ par champ
                i = 1
                For Each def In col
                    If LenB(def.ColonneBdD) > 0 Then
                        On Error Resume Next
                        ws.Cells(ligne, i).Value = REQSEL.Fields(def.ColonneBdD).Value
                        On Error GoTo SortieBdD
                    End If
                    i = i + 1
                Next def
                trouve = True
                nbOk = nbOk + 1
            End If
        End If

        If Not trouve Then
            ' Cle non trouvee : on remplit au moins la cle + un Statut explicite
            ws.Cells(ligne, 1).Value = cle
            ws.Cells(ligne, col.Count + 1).Value = "Cle non trouvee en BdD"
            ws.Cells(ligne, col.Count + 1).Font.Color = vbRed
            nbKo = nbKo + 1
        End If

        ligne = ligne + 1
    Next cle

    Call Close_CONN(ChaineConnexionSIRUS)
    Debug.Print "PreRemplirDepuisBdD : " & nbOk & " trouve(s), " & nbKo & " absent(es)."
    Exit Sub

SortieBdD:
    On Error Resume Next
    Call Close_CONN(ChaineConnexionSIRUS)
    On Error GoTo 0
    MsgBox "Erreur lors du pre-remplissage BdD : " & Err.Description, vbCritical
End Sub

' =========================================================================
' NomTableSQL
' =========================================================================
' Mapping table logique config -> nom de table SQL Excel (entoure de [...$]).
Private Function NomTableSQL(ByVal sTable As String) As String
    Select Case sTable
        Case "Produit":    NomTableSQL = "Produit$"
        Case "InfoATM":    NomTableSQL = "Info_ATM$"
        Case "TachesATM":  NomTableSQL = "Taches_ATM$"
        Case "EMMST":      NomTableSQL = "EMMST$"
        Case "ProjetSLI":  NomTableSQL = "Projet_SLI$"
        Case "ProjetASL":  NomTableSQL = "Projet_ASL$"
        Case Else:         NomTableSQL = ""
    End Select
End Function

' =========================================================================
' FeuilleExiste
' =========================================================================
Private Function FeuilleExiste(wb As Workbook, ByVal sNom As String) As Boolean
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If ws.Name = sNom Then
            FeuilleExiste = True
            Exit Function
        End If
    Next ws
    FeuilleExiste = False
End Function

' =========================================================================
' PROCEDURES DE TEST (executables depuis VBE)
' =========================================================================
' Important : pour ces tests, le classeur projet (.xlsx/.xlsm) doit etre
' ouvert ET actif (ActiveWorkbook != .xlam) AVANT de lancer le test.

Public Sub Test_Generer_Produit_Creer()
    Call SIRUS_GenererFeuilleAnalyse(SIRUS_TBL_PRODUIT, SIRUS_OP_CREER)
End Sub

Public Sub Test_Generer_Produit_Modifier()
    Dim cles As Variant
    ' Remplacer par des DTR reels presents en BdD pour tester le pre-remplissage.
    cles = Array("DTR_TEST_001", "DTR_TEST_002")
    Call SIRUS_GenererFeuilleAnalyse(SIRUS_TBL_PRODUIT, SIRUS_OP_MODIFIER, cles)
End Sub

Public Sub Test_Generer_ProjetSLI_Creer()
    Call SIRUS_GenererFeuilleAnalyse(SIRUS_TBL_PROJETSLI, SIRUS_OP_CREER)
End Sub

Public Sub Test_Generer_EMMST_Creer()
    Call SIRUS_GenererFeuilleAnalyse(SIRUS_TBL_EMMST, SIRUS_OP_CREER)
End Sub

' =========================================================================
' Test_DiagnostiquerColonnesBdDFeuille
' =========================================================================
' Genere une feuille "SIRUS_Diag_Colonnes" dans ActiveWorkbook avec tous
' les noms reels des colonnes BdD pour les 6 tables. Pas de troncature
' (contrairement a la fenetre Immediate de VBE).
'
' Layout :
'   col A : Table SQL (ex. "Produit$")
'   col B : Index colonne (1, 2, ...)
'   col C : Nom BdD exact
'   col D : Longueur du nom (utile pour reperer les espaces de bout)
'
' Procedure a privilegier pour resoudre les ecarts identifies par
' Test_DiagnostiquerEcartsConfig_Toutes.
Public Sub Test_DiagnostiquerColonnesBdDFeuille()

    Const FEUILLE As String = "SIRUS_Diag_Colonnes"

    Dim tables As Variant
    Dim t As Variant
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim sSQL As String
    Dim i As Long
    Dim ligne As Long
    Dim sNom As String

    On Error GoTo Sortie

    ' -- 1. ActiveWorkbook (classeur projet ou dev ECOSLI.xlsm) ----
    ' Pas de blocage sur ThisWorkbook : la feuille de diag peut etre
    ' generee dans n'importe quel classeur visible (incluant ECOSLI.xlsm
    ' en mode dev).
    Set wb = ActiveWorkbook
    If wb Is Nothing Then
        MsgBox "Aucun classeur actif.", vbExclamation
        Exit Sub
    End If

    ' -- 2. (Re)creer la feuille ---------------------------
    On Error Resume Next
    Application.DisplayAlerts = False
    wb.Worksheets(FEUILLE).Delete
    Application.DisplayAlerts = True
    On Error GoTo Sortie

    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = FEUILLE

    ' -- 3. Entetes ----------------------------------------
    ws.Cells(1, 1).Value = "Table"
    ws.Cells(1, 2).Value = "Index"
    ws.Cells(1, 3).Value = "Nom BdD"
    ws.Cells(1, 4).Value = "Longueur"
    With ws.Range("A1:D1")
        .Font.Bold = True
        .Interior.Color = COULEUR_ENTETE
        .HorizontalAlignment = xlCenter
    End With

    ' -- 4. Connexion BdD + parcours des tables ------------
    tables = Array("Produit$", "Info_ATM$", "Taches_ATM$", "EMMST$", _
                   "Projet_SLI$", "Projet_ASL$")

    Call Open_CONN(ChaineConnexionSIRUS)
    ligne = 2

    For Each t In tables
        sSQL = "SELECT * FROM [" & t & "] WHERE 1=0"
        Call open_REQSEL(sSQL)
        If Not (REQSEL Is Nothing) Then
            For i = 0 To REQSEL.Fields.Count - 1
                sNom = REQSEL.Fields(i).Name
                ws.Cells(ligne, 1).Value = t
                ws.Cells(ligne, 2).Value = i + 1
                ' Forcer ecriture en texte (eviter conversion auto Excel
                ' qui pourrait transformer un nom commencant par '=' etc.)
                ws.Cells(ligne, 3).NumberFormat = "@"
                ws.Cells(ligne, 3).Value = sNom
                ws.Cells(ligne, 4).Value = Len(sNom)
                ligne = ligne + 1
            Next i
        End If
    Next t

    Call Close_CONN(ChaineConnexionSIRUS)

    ' -- 5. Mise en forme ----------------------------------
    ws.Columns("A:B").AutoFit
    ws.Columns("C").ColumnWidth = 80
    ws.Columns("D").ColumnWidth = 10
    ws.Activate
    ws.Range("A1").Select
    ActiveWindow.FreezePanes = False
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True

    MsgBox "Feuille '" & FEUILLE & "' generee avec " & (ligne - 2) & _
           " colonnes BdD au total." & vbCrLf & _
           "Filtrer la colonne A pour isoler une table.", vbInformation, "SIRUS Diag"
    Exit Sub

Sortie:
    On Error Resume Next
    Call Close_CONN(ChaineConnexionSIRUS)
    On Error GoTo 0
    MsgBox "Erreur Test_DiagnostiquerColonnesBdDFeuille : " & Err.Description, vbCritical
End Sub

' =========================================================================
' Test_DiagnostiquerColonnesBdD
' =========================================================================
' Pour chaque table BdD, fait un SELECT * WHERE 1=0 (structure sans
' donnees) et trace dans la fenetre Immediate les noms reels des
' colonnes BdD. Permet de detecter les ecarts avec les ColonneBdD du
' config (accents, espaces, caractere ellipsis '...', casse...).
'
' Usage : lancer depuis VBE, regarder la fenetre Immediate, comparer
' avec le contenu de SIRUS_Config (ligne 2 des feuilles SIRUS_*).
Public Sub Test_DiagnostiquerColonnesBdD()

    Dim tables As Variant
    Dim t As Variant
    Dim sSQL As String
    Dim i As Long
    Dim sNom As String

    tables = Array("Produit$", "Info_ATM$", "Taches_ATM$", "EMMST$", _
                   "Projet_SLI$", "Projet_ASL$")

    On Error GoTo Sortie

    Debug.Print "=== Test_DiagnostiquerColonnesBdD ==="
    Call Open_CONN(ChaineConnexionSIRUS)

    For Each t In tables
        sSQL = "SELECT * FROM [" & t & "] WHERE 1=0"
        Debug.Print ""
        Debug.Print "--- Table : [" & t & "] ---"
        Call open_REQSEL(sSQL)
        If Not (REQSEL Is Nothing) Then
            For i = 0 To REQSEL.Fields.Count - 1
                sNom = REQSEL.Fields(i).Name
                ' Encadre par '|' pour voir les espaces de tete/queue
                Debug.Print "  [" & Format$(i + 1, "00") & "] |" & sNom & "|"
            Next i
        End If
    Next t

    Call Close_CONN(ChaineConnexionSIRUS)
    Debug.Print ""
    Debug.Print "=== Fin diagnostic ==="
    Exit Sub

Sortie:
    On Error Resume Next
    Call Close_CONN(ChaineConnexionSIRUS)
    On Error GoTo 0
    Debug.Print "Erreur Test_DiagnostiquerColonnesBdD : " & Err.Description
End Sub

' =========================================================================
' Test_DiagnostiquerEcartsConfig
' =========================================================================
' Pour la table demandee, compare les ColonneBdD du config a ceux
' reellement presents en BdD et signale les noms manquants.
' Reutilise SIRUS_LireConfig + un SELECT * WHERE 1=0.
Public Sub Test_DiagnostiquerEcartsConfig_Produit()
    Call DiagnostiquerEcartsTable(SIRUS_TBL_PRODUIT, "Produit$")
End Sub

Public Sub Test_DiagnostiquerEcartsConfig_Toutes()
    Call DiagnostiquerEcartsTable(SIRUS_TBL_PRODUIT, "Produit$")
    Call DiagnostiquerEcartsTable(SIRUS_TBL_INFOATM, "Info_ATM$")
    Call DiagnostiquerEcartsTable(SIRUS_TBL_TACHESATM, "Taches_ATM$")
    Call DiagnostiquerEcartsTable(SIRUS_TBL_EMMST, "EMMST$")
    Call DiagnostiquerEcartsTable(SIRUS_TBL_PROJETSLI, "Projet_SLI$")
    Call DiagnostiquerEcartsTable(SIRUS_TBL_PROJETASL, "Projet_ASL$")
End Sub

Private Sub DiagnostiquerEcartsTable(ByVal sTable As String, ByVal sTableSQL As String)

    Dim col As Collection
    Dim def As SIRUS_Champ_Config
    Dim noms As Object  ' Dictionary des noms BdD reels
    Dim sSQL As String
    Dim i As Long
    Dim nbTrouve As Long, nbAbsent As Long

    On Error GoTo Sortie

    Set noms = CreateObject("Scripting.Dictionary")
    noms.CompareMode = 1  ' BinaryCompare (sensible a la casse)

    ' Charger les vrais noms BdD
    Call Open_CONN(ChaineConnexionSIRUS)
    sSQL = "SELECT * FROM [" & sTableSQL & "] WHERE 1=0"
    Call open_REQSEL(sSQL)
    If Not (REQSEL Is Nothing) Then
        For i = 0 To REQSEL.Fields.Count - 1
            noms(REQSEL.Fields(i).Name) = True
        Next i
    End If
    Call Close_CONN(ChaineConnexionSIRUS)

    ' Comparer aux ColonneBdD du config
    Set col = SIRUS_LireConfig(sTable)
    Debug.Print ""
    Debug.Print "=== Ecarts config <-> BdD pour " & sTable & " ([" & sTableSQL & "]) ==="

    For Each def In col
        If LenB(def.ColonneBdD) > 0 Then
            If noms.Exists(def.ColonneBdD) Then
                nbTrouve = nbTrouve + 1
            Else
                nbAbsent = nbAbsent + 1
                Debug.Print "  ABSENT en BdD : |" & def.ColonneBdD & "|  (champ '" & def.Champ & "')"
            End If
        End If
    Next def

    Debug.Print "  -> " & nbTrouve & " trouve(s), " & nbAbsent & " absent(s)."
    Exit Sub

Sortie:
    On Error Resume Next
    Call Close_CONN(ChaineConnexionSIRUS)
    On Error GoTo 0
    Debug.Print "Erreur DiagnostiquerEcartsTable(" & sTable & ") : " & Err.Description
End Sub
