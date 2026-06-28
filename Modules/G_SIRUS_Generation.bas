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
Public Const SIRUS_OP_CREER      As String = "Creer"
Public Const SIRUS_OP_MODIFIER   As String = "Modifier"
Public Const SIRUS_OP_EXTRACTION As String = "Extraction"

Private Const SIRUS_LIGNE_ENTETE_HUMAIN  As Long = 1
Private Const SIRUS_LIGNE_ENTETE_BDD     As Long = 2
Private Const SIRUS_LIGNE_DONNEES_DEBUT  As Long = 3
Private Const SIRUS_NB_LIGNES_VALIDATION As Long = 1000  ' lignes couvertes par les listes deroulantes

Private Const COULEUR_ENTETE             As Long = 14606303   ' bleu tres clair RGB(220,230,241)
Private Const COULEUR_AUTO_HERITE        As Long = 14211288   ' gris clair RGB(216,216,216)
Private Const COULEUR_ENTETE_BDD         As Long = 9868950    ' gris fonce RGB(150,150,150)

' Mot de passe de protection des feuilles SIRUS generees (Phase 3.1).
' UserInterfaceOnly:=True permet a VBA (moteur d'analyse Phase 2, ATMAP
' Phase 5) d'ecrire dans toutes les cellules, l'utilisateur ne peut
' modifier QUE les cellules Saisie (Auto/Herite + entetes + Statut verrouilles).
' A centraliser en Phase 6.2 (MOT_DE_PASSE_PROTECTION_FEUILLES dans params).
Private Const SIRUS_PROTECTION_PWD As String = "SIRUS"

' =========================================================================
' Limite ACE OLE DB : noms de champ Excel tronques a 64 caracteres
' =========================================================================
' Le driver Microsoft.ACE.OLEDB.12.0 (et son ancetre Jet) tronque les
' noms de colonnes Excel a 64 caracteres quand exposes via OLE DB. Donc
' un SELECT [nom > 64 chars] echoue avec "Aucune valeur donnee pour un
' ou plusieurs des parametres requis" (Jet interprete le nom inconnu
' comme un parametre).
'
' Strategie SIRUS : on garde le nom complet dans le config (utile pour
' lisibilite humaine et eventuellement pour la PJ ATMAP cote Data Owner)
' et on tronque automatiquement dans les SELECT / INSERT via cette
' fonction.
Public Const SIRUS_BDD_MAX_NOM_COL As Long = 64

Public Function TronquerColonneBdD(ByVal sNom As String) As String
    If Len(sNom) > SIRUS_BDD_MAX_NOM_COL Then
        TronquerColonneBdD = Left$(sNom, SIRUS_BDD_MAX_NOM_COL)
    Else
        TronquerColonneBdD = sNom
    End If
End Function

' =========================================================================
' Connexion BdD SIRUS
' =========================================================================
' SIRUS lit (et a terme ecrit indirectement via ATMAP) la BdD de TRAVAIL,
' pas la BdD de PROD. La chaine est construite en remplacant 'UserProfile'
' par le vrai chemin utilisateur via Environ$.
Private Function ChaineConnexionSIRUS() As String
    ChaineConnexionSIRUS = Replace(CONNEXION_EXCEL_BDD_TRAVAIL, "UserProfile", Environ$("UserProfile"))
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
Public Sub SIRUS_GenererFeuilleAnalyse(ByVal sTable As String, ByVal sOperation As String, Optional ByRef vCles As Variant)

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
    If sOperation <> SIRUS_OP_CREER And sOperation <> SIRUS_OP_MODIFIER And sOperation <> SIRUS_OP_EXTRACTION Then
        MsgBox "SIRUS : operation invalide ('" & sOperation & "'). Doit etre 'Creer', 'Modifier' ou 'Extraction'.", vbCritical
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
        resp = MsgBox("Le classeur actif '" & wb.Name & "' ne contient pas de feuille DATA_ASL (habituellement le marqueur d'un classeur projet SLI)." & vbCrLf & vbCrLf & "Voulez-vous quand meme generer la feuille SIRUS dans ce classeur ?", vbYesNo + vbQuestion + vbDefaultButton2, "SIRUS - Classeur cible inhabituel")
        If resp <> vbYes Then Exit Sub
    End If

    ' -- 3. Lire le config -----------------------------------
    Set col = SIRUS_LireConfig(sTable)
    If col.Count = 0 Then
        MsgBox "Aucun champ defini dans SIRUS_Config pour la table '" & sTable & "'. Lancer SIRUS_InitConfigSheet d'abord.", vbCritical
        Exit Sub
    End If

    ' -- 4. Nom de la feuille + confirmation si existante ---
    nomFeuille = "SIRUS_" & sTable & "_" & sOperation
    If Len(nomFeuille) > 31 Then nomFeuille = Left$(nomFeuille, 31)  ' limite Excel

    If FeuilleExiste(wb, nomFeuille) Then
        resp = MsgBox("La feuille '" & nomFeuille & "' existe deja dans le classeur '" & wb.Name & "'." & vbCrLf & vbCrLf & "Voulez-vous l'ecraser ?", vbYesNo + vbQuestion + vbDefaultButton2, "SIRUS - Generation de feuille")
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
            MsgBox "Operation 'Modifier' : il faut fournir des cles (DTR ou Nom projet) a pre-remplir. La feuille a ete creee vierge.", vbExclamation
        Else
            Call PreRemplirDepuisBdD(wsCible, sTable, col, vCles)
        End If
    End If

    ' -- 8. Listes deroulantes ------------------------------
    Call AppliquerListes(wsCible, col)

    ' -- 9. Griser les champs Auto / Herite -----------------
    Call GriserChampsAutoHerite(wsCible, col)

    ' -- 9 bis. Verrouiller cellules Auto/Herite + entetes + Statut
    '          et activer la protection feuille (Phase 3.1).
    Call VerrouillerEtProteger(wsCible, col)

    ' -- 10. Alignement vertical centre sur toute la zone de donnees ----
    '    Donne un rendu propre meme quand le WrapText elargit certaines
    '    cellules (notamment la colonne Commentaires apres l analyse).
    On Error Resume Next
    wsCible.UsedRange.VerticalAlignment = xlCenter
    On Error GoTo 0

    ' -- 11. Freeze panes selon la table + activation cellule -----
    '    Fige lignes 1-2 (entetes) + colonnes a gauche jusqu a la cle metier
    '    (incluse) via SIRUS_ColonneClePourTable :
    '       ProjetASL : Statut + Projet + Sous Systeme + DTR (col 1-4)
    '       Produit   : Statut + Famille Generique + DTR     (col 1-3)
    '       autres    : Statut + DTR / Nom projet            (col 1-2)
    Dim colFreeze As Long
    colFreeze = SIRUS_ColonneClePourTable(sTable) + 1
    wsCible.Activate
    ActiveWindow.FreezePanes = False
    wsCible.Cells(SIRUS_LIGNE_DONNEES_DEBUT, colFreeze).Select
    ActiveWindow.FreezePanes = True
    ' Curseur sur 1re cellule de saisie (col 2, juste apres Statut)
    wsCible.Cells(SIRUS_LIGNE_DONNEES_DEBUT, 2).Select

    MsgBox "Feuille '" & nomFeuille & "' generee avec succes." & vbCrLf & "Statut (col 1) + " & col.Count & " champs + Commentaires (col " & (col.Count + 2) & ").", vbInformation, "SIRUS"
    Exit Sub

Sortie:
    MsgBox "Erreur dans SIRUS_GenererFeuilleAnalyse : " & Err.Description, vbCritical
End Sub

' =========================================================================
' PoserEntetes
' =========================================================================
' Layout (refonte 2026-06-10) :
'   colonne 1            : "Statut"        (ColonneBdD = "_statut")
'   colonnes 2..(nbCol+1): champs du config (ordre = Ordre croissant)
'   colonne (nbCol+2)    : "Commentaires"  (ColonneBdD = "_commentaires")
'
' Ligne 1 : libelles humains. Ligne 2 : ColonneBdD techniques (masquee).
Private Sub PoserEntetes(ws As Worksheet, col As Collection)

    Dim def As SIRUS_Champ_Config
    Dim i As Long
    Dim nbCol As Long
    Dim colStatut As Long, colCommentaires As Long

    nbCol = col.Count
    colStatut = 1
    colCommentaires = nbCol + 2

    ' --- Statut en colonne 1 ---
    ws.Cells(SIRUS_LIGNE_ENTETE_HUMAIN, colStatut).Value = "Statut"
    ws.Cells(SIRUS_LIGNE_ENTETE_BDD, colStatut).Value = "_statut"

    ' --- Champs du config en colonnes 2..(nbCol+1) ---
    '    Marquage visuel des champs obligatoires :
    '       " *"   en rouge pour Obligatoire = "Oui"
    '       " (*)" en rouge pour Obligatoire = "Conditionnel"
    '    Permet a l'utilisateur de reperer en un coup d'oeil les champs
    '    a remplir avant de pousser la demande ATMAP.
    i = 2
    For Each def In col
        Call PoserLibelleEntete(ws.Cells(SIRUS_LIGNE_ENTETE_HUMAIN, i), def)
        ws.Cells(SIRUS_LIGNE_ENTETE_BDD, i).Value = def.ColonneBdD
        i = i + 1
    Next def

    ' --- Commentaires en derniere colonne ---
    ws.Cells(SIRUS_LIGNE_ENTETE_HUMAIN, colCommentaires).Value = "Commentaires"
    ws.Cells(SIRUS_LIGNE_ENTETE_BDD, colCommentaires).Value = "_commentaires"

    ' Formatage ligne 1 (humain)
    With ws.Range(ws.Cells(SIRUS_LIGNE_ENTETE_HUMAIN, 1), ws.Cells(SIRUS_LIGNE_ENTETE_HUMAIN, colCommentaires))
        .Font.Bold = True
        .Interior.Color = COULEUR_ENTETE
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With

    ' Formatage ligne 2 (BdD) + masquage
    With ws.Range(ws.Cells(SIRUS_LIGNE_ENTETE_BDD, 1), ws.Cells(SIRUS_LIGNE_ENTETE_BDD, colCommentaires))
        .Font.Color = COULEUR_ENTETE_BDD
        .Font.Size = 8
        .Font.Italic = True
    End With
    ws.Rows(SIRUS_LIGNE_ENTETE_BDD).Hidden = True

    ' Largeur auto (basee sur le contenu de l'entete) avec bornes :
    ' AutoFit pur avec WrapText donne des largeurs minimales (Excel essaie
    ' d'ajuster en hauteur). On force un min/max raisonnable pour la
    ' lisibilite : 14 (champs courts) <= largeur <= 35 (libelles longs).
    Const LARG_MIN As Double = 14
    Const LARG_MAX As Double = 35
    Dim cIdx As Long
    ws.Range(ws.Cells(SIRUS_LIGNE_ENTETE_HUMAIN, 1), ws.Cells(SIRUS_LIGNE_ENTETE_HUMAIN, colCommentaires)).EntireColumn.AutoFit
    For cIdx = 1 To colCommentaires
        If ws.Columns(cIdx).ColumnWidth < LARG_MIN Then
            ws.Columns(cIdx).ColumnWidth = LARG_MIN
        ElseIf ws.Columns(cIdx).ColumnWidth > LARG_MAX Then
            ws.Columns(cIdx).ColumnWidth = LARG_MAX
        End If
    Next cIdx

    ' Hauteur entete + figer. Hauteur 45 pour laisser le wrap text afficher
    ' les libelles sur 2 lignes sans coupure.
    ws.Rows(SIRUS_LIGNE_ENTETE_HUMAIN).RowHeight = 45
End Sub

' =========================================================================
' PoserLibelleEntete (helper interne)
' =========================================================================
' Ecrit le libelle Champ dans la cellule d'entete + suffixe rouge pour les
' champs obligatoires :
'   - Obligatoire = "Oui"          ==> suffixe " *"   en rouge
'   - Obligatoire = "Conditionnel" ==> suffixe " (*)" en rouge
'   - Obligatoire = "Non"          ==> rien
' Le libelle principal reste dans la couleur d'entete (noir gras).
Private Sub PoserLibelleEntete(ByVal cellEntete As Range, ByVal def As SIRUS_Champ_Config)
    Dim sLibelle As String
    Dim sSuffixe As String
    Dim posSuffixe As Long

    sLibelle = def.Champ
    Select Case def.Obligatoire
        Case "Oui"
            sSuffixe = " *"
        Case "Conditionnel"
            sSuffixe = " (*)"
        Case Else
            sSuffixe = ""
    End Select

    cellEntete.Value = sLibelle & sSuffixe

    ' Mettre le suffixe en rouge gras pour le distinguer du libelle
    If LenB(sSuffixe) > 0 Then
        posSuffixe = Len(sLibelle) + 1
        With cellEntete.Characters(Start:=posSuffixe, Length:=Len(sSuffixe)).Font
            .Color = RGB(192, 0, 0)
            .Bold = True
        End With
    End If
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

    ' Champs decales en col 2..(nbCol+1) suite a Statut en col 1.
    i = 2
    For Each def In col
        If def.TypeChamp = "Liste" Or def.TypeChamp = "Booleen" Then
            sFormula = SIRUS_ListeEnChaine(def.ListeSource)
            If LenB(sFormula) > 0 Then
                Set plage = ws.Range( ws.Cells(SIRUS_LIGNE_DONNEES_DEBUT, i), ws.Cells(SIRUS_LIGNE_DONNEES_DEBUT + SIRUS_NB_LIGNES_VALIDATION - 1, i))
                On Error Resume Next
                plage.Validation.Delete
                plage.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=sFormula
                On Error GoTo 0
            Else
                Debug.Print "AppliquerListes : liste '" & def.ListeSource & "' vide ou inconnue pour le champ '" & def.Champ & "'."
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

    ' Champs decales en col 2..(nbCol+1) suite a Statut en col 1.
    i = 2
    For Each def In col
        If def.Source = "Auto" Or def.Source = "Herite" Then
            Set plage = ws.Range( ws.Cells(SIRUS_LIGNE_DONNEES_DEBUT, i), ws.Cells(SIRUS_LIGNE_DONNEES_DEBUT + SIRUS_NB_LIGNES_VALIDATION - 1, i))
            plage.Interior.Color = COULEUR_AUTO_HERITE
        End If
        i = i + 1
    Next def
End Sub

' =========================================================================
' VerrouillerEtProteger (Phase 3.1)
' =========================================================================
' Verrouille les cellules non-editables par l'utilisateur (entetes + colonne
' Statut + colonnes Source = "Auto"/"Herite") puis active la protection de
' la feuille avec UserInterfaceOnly:=True (VBA peut toujours ecrire partout).
'
' Strategie : deverrouiller TOUTES les cellules d'abord (par defaut Excel
' met .Locked = True partout), puis re-verrouiller specifiquement les zones
' qui ne doivent pas etre editees. Plus simple a maintenir que de tracker
' les Saisie individuellement.
'
' Note : UserInterfaceOnly n'est pas persiste apres Save/Reopen. Si le
' classeur est ferme puis rouvert, VBA devra reappeler Protect pour
' restaurer le drapeau. A traiter en Phase 6.2 si besoin (Workbook_Open).
Private Sub VerrouillerEtProteger(ws As Worksheet, col As Collection)

    Dim def As SIRUS_Champ_Config
    Dim i As Long
    Dim nbCol As Long
    Dim plage As Range

    nbCol = col.Count

    ' 1. Deverrouiller toutes les cellules (.Locked = True par defaut sinon)
    ws.Cells.Locked = False

    ' 2. Verrouiller les 2 lignes d'entete (humain + ColonneBdD masquee)
    ws.Rows(SIRUS_LIGNE_ENTETE_HUMAIN).Locked = True
    ws.Rows(SIRUS_LIGNE_ENTETE_BDD).Locked = True

    ' 3. Verrouiller la colonne 1 (Statut) - ecrite par le moteur d'analyse
    '    Phase 2, pas par l'utilisateur.
    ws.Columns(1).Locked = True

    ' 4. Colonne Commentaires (en bout : nbCol + 2) : LAISSEE DEVERROUILLEE
    '    pour la saisie libre utilisateur (note metier, justification).

    ' 5. Verrouiller les colonnes Auto / Herite dans la zone de saisie
    '    (decalees de +1 a cause de la col Statut).
    i = 2
    For Each def In col
        If def.Source = "Auto" Or def.Source = "Herite" Then
            Set plage = ws.Range( ws.Cells(SIRUS_LIGNE_DONNEES_DEBUT, i), ws.Cells(SIRUS_LIGNE_DONNEES_DEBUT + SIRUS_NB_LIGNES_VALIDATION - 1, i))
            plage.Locked = True
        End If
        i = i + 1
    Next def

    ' 5. Activer la protection.
    '    - UserInterfaceOnly:=True : VBA peut ecrire dans toutes les cellules
    '      (moteur d'analyse, ATMAP), seul l'utilisateur est bloque sur les
    '      cellules verrouillees.
    '    - DrawingObjects:=False : autorise VBA a manipuler les objets dessin
    '      (commentaires de cellule, formes). Necessaire pour que
    '      EcrireStatutCellule puisse poser un commentaire de cellule avec
    '      le detail des alertes (sinon AddComment leve "Erreur 1004 :
    '      Impossible de definir la propriete Comment, la feuille est protegee").
    '    - AllowFormattingColumns/Rows : utile pour AutoFit, masquage de lignes.
    ws.Protect Password:=SIRUS_PROTECTION_PWD, DrawingObjects:=False, UserInterfaceOnly:=True, AllowFormattingCells:=False, AllowFormattingColumns:=True, AllowFormattingRows:=True
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
Private Sub PreRemplirDepuisBdD(ws As Worksheet, ByVal sTable As String, col As Collection, ByRef vCles As Variant)

    Dim def As SIRUS_Champ_Config
    Dim cleColBdD As String
    Dim sNomTableSQL As String
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
        MsgBox "Cle BdD introuvable pour la table '" & sTable & "' (1er champ du config sans ColonneBdD).", vbCritical
        Exit Sub
    End If

    ' -- Nom de la table SQL -------------------------------
    sNomTableSQL = NomTableSQL(sTable)
    If LenB(sNomTableSQL) = 0 Then
        MsgBox "Table SQL inconnue pour '" & sTable & "'.", vbCritical
        Exit Sub
    End If

    ' -- Ouvrir la connexion -------------------------------
    On Error GoTo SortieBdD
    Call Open_CONN(ChaineConnexionSIRUS)

    ' -- Boucle par cle ------------------------------------
    ligne = SIRUS_LIGNE_DONNEES_DEBUT
    nbOk = 0
    nbKo = 0

    ' Diagnostic : trace les colonnes du config absentes en BdD (pour aide
    ' au cadrage des Champs_BDD.xlsx). Trace une fois par run.
    Dim dictAbsentes As Object
    Set dictAbsentes = CreateObject("Scripting.Dictionary")
    dictAbsentes.CompareMode = 1

    ' SELECT * : on recupere toutes les colonnes BdD existantes puis on
    ' lit champ par champ avec tolerance aux colonnes manquantes (cas des
    ' ??_*_a_definir et des colonnes BdD non encore creees comme
    ' "Prix spare ..._PROJETASL" qui peut etre absente cote BdD TRAVAIL).
    ' Plus robuste qu un SELECT explicite qui plante en bloc des qu'une
    ' seule colonne demandee n'existe pas.
    For Each cle In vCles
        sSQL = "SELECT * FROM [" & sNomTableSQL & "] WHERE [" & TronquerColonneBdD(cleColBdD) & "] = '" & Replace(CStr(cle), "'", "''") & "'"
        Debug.Print "PreRemplirDepuisBdD SQL : " & sSQL
        Call open_REQSEL(sSQL)

        trouve = False
        If Not (REQSEL Is Nothing) Then
            ' Boucle Do While (et non If) car une cle peut renvoyer
            ' N lignes (ProjetASL : 1 nom de projet -> N DTR du projet).
            ' Chaque ligne du Recordset est ecrite sur une ligne distincte
            ' de la feuille SIRUS.
            Do While Not REQSEL.EOF
                ' Remplir la ligne champ par champ. Les champs du config
                ' sont en col 2..(nbCol+1) suite a Statut en col 1.
                i = 2
                For Each def In col
                    If LenB(def.ColonneBdD) > 0 Then
                        Dim sColTrim As String
                        Dim vVal As Variant
                        sColTrim = TronquerColonneBdD(def.ColonneBdD)
                        On Error Resume Next
                        Err.Clear
                        vVal = REQSEL.Fields(sColTrim).Value
                        If Err.Number <> 0 Then
                            ' Colonne absente cote BdD : on laisse la
                            ' cellule vide et on trace dans Debug.Print
                            ' (1 fois par colonne).
                            If Not dictAbsentes.Exists(sColTrim) Then
                                dictAbsentes.Add sColTrim, Empty
                                Debug.Print "PreRemplirDepuisBdD : colonne BdD absente '" & sColTrim & "' (champ '" & def.Champ & "')"
                            End If
                            Err.Clear
                        Else
                            If Not IsNull(vVal) Then
                                ws.Cells(ligne, i).Value = vVal
                            End If
                        End If
                        On Error GoTo SortieBdD
                    End If
                    i = i + 1
                Next def
                trouve = True
                nbOk = nbOk + 1
                ligne = ligne + 1
                REQSEL.MoveNext
            Loop
        End If

        If Not trouve Then
            ' Cle non trouvee :
            '   - col 1 (Statut) : marquer "Cle non trouvee" en rouge
            '   - col 2 (1er champ config = cle) : reporter la valeur de cle
            ws.Cells(ligne, 1).Value = "Cle non trouvee en BdD"
            ws.Cells(ligne, 1).Font.Color = vbRed
            ws.Cells(ligne, 2).Value = cle
            nbKo = nbKo + 1
            ligne = ligne + 1
        End If
    Next cle

    Call Close_CONN(ChaineConnexionSIRUS)
    Debug.Print "PreRemplirDepuisBdD : " & nbOk & " trouve(s), " & nbKo & " absent(es), " & dictAbsentes.count & " colonne(s) BdD absente(s)."
    Exit Sub

SortieBdD:
    Dim sErr As String
    sErr = Err.description
    On Error Resume Next
    Call Close_CONN(ChaineConnexionSIRUS)
    On Error GoTo 0
    MsgBox "Erreur lors du pre-remplissage BdD :" & vbCrLf & vbCrLf & sErr & vbCrLf & vbCrLf & "SQL en cours :" & vbCrLf & sSQL, vbCritical, "SIRUS"
End Sub

' =========================================================================
' SIRUS_EnrichirProjetASL_DepuisDATA_ASL
' =========================================================================
' Pour ProjetASL (Creer ou Modifier) : enrichit / complete la feuille
' SIRUS_ProjetASL_<Op> avec les donnees de la PBS DATA_ASL du classeur
' projet actif. Champs croises (spec Thomas 2026-06-11) :
'   Nom Projet      <- sProjet (= Project_List.Value, identique pour
'                       toutes les lignes)
'   Sous Systeme    <- DATA_ASL [SOUS SYSTEME]
'   DTR             <- DATA_ASL [DTR Reference (issue de PBS)]
'   Designation PBS <- DATA_ASL [DESIGNATION (issue de PBS)]
'   Qte             <- DATA_ASL [Qte]
'   Node Type       <- DATA_ASL [Node typed (LRU/SRU...)]
'
' Strategie :
'   1. Charger les lignes deja presentes dans la feuille SIRUS (cas Modifier
'      apres pre-remplissage BdD) -> dict DTR_PBS -> ligne SIRUS.
'   2. Iterer DATA_ASL. Pour chaque DTR :
'        - present en SIRUS -> overwrite des 6 champs (DATA_ASL fait foi
'          pour la photo PBS courante).
'        - absent en SIRUS -> ajouter une nouvelle ligne en fin (cas DTR
'          PBS qui n existe pas encore en BdD ProjetASL).
'   3. Les autres colonnes (MTBF fct, Prix archive, etc.) restent ce que
'      le pre-remplissage BdD a pose (ou vides pour les nouveaux DTR).
'
' Robustesse :
'   - Si DATA_ASL absente / non cadrable : MsgBox + Exit (l utilisateur
'     n a pas de PBS active, on ne touche pas).
'   - Cols absentes en DATA_ASL (manqueOptionnel) : skip silencieux du
'     champ correspondant.
Public Sub SIRUS_EnrichirProjetASL_DepuisDATA_ASL(ByVal sOperation As String, ByVal sProjet As String)
    Dim wb As Workbook
    Dim wsCible As Worksheet
    Dim wsPBS As Worksheet
    Dim nomFeuille As String

    Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Sub

    ' -- 1. Localiser la feuille SIRUS_ProjetASL_<Op> ------
    nomFeuille = "SIRUS_ProjetASL_" & sOperation
    If Len(nomFeuille) > 31 Then nomFeuille = Left$(nomFeuille, 31)
    On Error Resume Next
    Set wsCible = wb.Worksheets(nomFeuille)
    On Error GoTo 0
    If wsCible Is Nothing Then
        Debug.Print "SIRUS_EnrichirProjetASL_DepuisDATA_ASL : feuille '" & nomFeuille & "' introuvable."
        Exit Sub
    End If

    ' -- 2. Cadrer DATA_ASL --------------------------------
    If Not VerifierFeuilleDATA_ASL(wb, wsPBS) Then Exit Sub
    If Not MapColonnesDATA_ASL(wsPBS) Then Exit Sub
    If DATA_ASL_type.DTR_PBS = 0 Then Exit Sub

    ' -- 3. Trouver les cols de la feuille SIRUS par ColonneBdD (ligne 2)
    Dim colNomProjet As Long, colSousSysteme As Long, colDTR As Long
    Dim colDesignation As Long, colQte As Long, colNodeType As Long
    Dim colMTBFIntr As Long, colMTBFFonct As Long
    Dim colDVL_HARD As Long, colDVL_SOFT As Long
    Dim colPrixSpareArch As Long, colLeadSpareArch As Long
    Dim colPrixRepairArch As Long, colLeadRepairArch As Long

    colNomProjet = TrouverColParColonneBdD(wsCible, "PROJET_PROJETASL")
    colSousSysteme = TrouverColParColonneBdD(wsCible, "Sous Systeme_PROJETASL")
    colDTR = TrouverColParColonneBdD(wsCible, "DTR_PROJETASL")
    colDesignation = TrouverColParColonneBdD(wsCible, "Designation_PROJETASL")
    colQte = TrouverColParColonneBdD(wsCible, "Q PBS_PROJETASL")
    colNodeType = TrouverColParColonneBdD(wsCible, "Node typed PROJET_PROJETASL")
    colMTBFIntr = TrouverColParColonneBdD(wsCible, "MTBF intrinseque Projet (h) issu du RAM report projet_PROJETASL")
    colMTBFFonct = TrouverColParColonneBdD(wsCible, "MTBF fonctionnel (h) issu du RAM report projet_PROJETASL")
    colDVL_HARD = TrouverColParColonneBdD(wsCible, "DUREE DE VIE LIMITEE (RAM report specifique projet (CGR)) HARDWARE_PROJET ASL")
    colDVL_SOFT = TrouverColParColonneBdD(wsCible, "DUREE DE VIE LIMITEE (RAM report specifique projet (CGR)) SOFTWARE_PROJET ASL")
    colPrixSpareArch = TrouverColParColonneBdD(wsCible, "Prix spare (DEVISE du projet (SLI))_PROJETASL")
    colLeadSpareArch = TrouverColParColonneBdD(wsCible, "Lead time SPARE (jours)_PROJETASL")
    colPrixRepairArch = TrouverColParColonneBdD(wsCible, "Prix repair (DEVISE du projet (SLI))_PROJETASL")
    colLeadRepairArch = TrouverColParColonneBdD(wsCible, "Lead time REPAIR (jours)_PROJETASL")

    If colDTR = 0 Then
        Debug.Print "SIRUS_EnrichirProjetASL_DepuisDATA_ASL : col DTR introuvable."
        Exit Sub
    End If

    ' -- 4. Index des lignes SIRUS existantes (DTR -> ligne) ---
    Dim dictSIRUS As Object
    Set dictSIRUS = CreateObject("Scripting.Dictionary")
    dictSIRUS.CompareMode = 1

    Dim ligneMaxSIRUS As Long, ligne As Long
    ligneMaxSIRUS = wsCible.Cells(wsCible.rows.count, colDTR).End(xlUp).Row
    For ligne = SIRUS_LIGNE_DONNEES_DEBUT To ligneMaxSIRUS
        Dim sDTRcourant As String
        sDTRcourant = Trim$(CStr(wsCible.Cells(ligne, colDTR).Value))
        If LenB(sDTRcourant) > 0 Then
            If Not dictSIRUS.Exists(sDTRcourant) Then dictSIRUS.Add sDTRcourant, ligne
        End If
    Next ligne

    ' -- 5. Parcourir DATA_ASL et croiser -------------------
    Dim derniereLignePBS As Long, lignePBS As Long
    derniereLignePBS = wsPBS.Cells(wsPBS.rows.count, DATA_ASL_type.DTR_PBS).End(xlUp).Row

    Dim ligneAjout As Long
    If ligneMaxSIRUS < SIRUS_LIGNE_DONNEES_DEBUT Then
        ligneAjout = SIRUS_LIGNE_DONNEES_DEBUT
    Else
        ligneAjout = ligneMaxSIRUS + 1
    End If

    Dim nbOverwrite As Long, nbAjout As Long
    nbOverwrite = 0
    nbAjout = 0

    For lignePBS = DATA_ASL_HeaderRow + 1 To derniereLignePBS
        Dim dtrPBS As String
        dtrPBS = Trim$(CStr(wsPBS.Cells(lignePBS, DATA_ASL_type.DTR_PBS).Value))
        If LenB(dtrPBS) = 0 Then GoTo NextPBS

        Dim ligneCible As Long
        If dictSIRUS.Exists(dtrPBS) Then
            ligneCible = dictSIRUS(dtrPBS)
            nbOverwrite = nbOverwrite + 1
        Else
            ligneCible = ligneAjout
            ligneAjout = ligneAjout + 1
            nbAjout = nbAjout + 1
        End If

        ' Nom Projet (toujours = sProjet)
        If colNomProjet > 0 Then wsCible.Cells(ligneCible, colNomProjet).Value = sProjet
        ' Sous Systeme
        If colSousSysteme > 0 And DATA_ASL_type.SousSysteme > 0 Then wsCible.Cells(ligneCible, colSousSysteme).Value = wsPBS.Cells(lignePBS, DATA_ASL_type.SousSysteme).Value
        ' DTR
        wsCible.Cells(ligneCible, colDTR).Value = dtrPBS
        ' Designation PBS
        If colDesignation > 0 And DATA_ASL_type.Designation_PBS > 0 Then wsCible.Cells(ligneCible, colDesignation).Value = wsPBS.Cells(lignePBS, DATA_ASL_type.Designation_PBS).Value
        ' Qte
        If colQte > 0 And DATA_ASL_type.Qte > 0 Then wsCible.Cells(ligneCible, colQte).Value = wsPBS.Cells(lignePBS, DATA_ASL_type.Qte).Value
        ' Node Type
        If colNodeType > 0 And DATA_ASL_type.NodeType > 0 Then wsCible.Cells(ligneCible, colNodeType).Value = wsPBS.Cells(lignePBS, DATA_ASL_type.NodeType).Value

        ' MTBF intrinseque Projet (RAM report specifique projet)
        If colMTBFIntr > 0 And DATA_ASL_type.MTBF_Intr > 0 Then wsCible.Cells(ligneCible, colMTBFIntr).Value = wsPBS.Cells(lignePBS, DATA_ASL_type.MTBF_Intr).Value

        ' MTBF fonctionnel (RAM report specifique projet)
        If colMTBFFonct > 0 And DATA_ASL_type.MTBF_Fonct > 0 Then wsCible.Cells(ligneCible, colMTBFFonct).Value = wsPBS.Cells(lignePBS, DATA_ASL_type.MTBF_Fonct).Value

        ' DVL fonctionnel HARD/SOFT (analytique = specifique projet CGR)
        If colDVL_HARD > 0 And DATA_ASL_type.DureeVie_Ana_HARD > 0 Then wsCible.Cells(ligneCible, colDVL_HARD).Value = wsPBS.Cells(lignePBS, DATA_ASL_type.DureeVie_Ana_HARD).Value
        If colDVL_SOFT > 0 And DATA_ASL_type.DureeVie_Ana_SOFT > 0 Then wsCible.Cells(ligneCible, colDVL_SOFT).Value = wsPBS.Cells(lignePBS, DATA_ASL_type.DureeVie_Ana_SOFT).Value

        ' Lead time Spare/Repair archive (en jours calendaires)
        If colLeadSpareArch > 0 And DATA_ASL_type.LeadTimeSpareCalend > 0 Then wsCible.Cells(ligneCible, colLeadSpareArch).Value = wsPBS.Cells(lignePBS, DATA_ASL_type.LeadTimeSpareCalend).Value
        If colLeadRepairArch > 0 And DATA_ASL_type.LeadTimeRepairCalend > 0 Then wsCible.Cells(ligneCible, colLeadRepairArch).Value = wsPBS.Cells(lignePBS, DATA_ASL_type.LeadTimeRepairCalend).Value

        ' Prix Spare archive : cascade Fournisseur -> Cat D&IS -> Cat FCS
        ' calcule (col AB) -> Archive. Une valeur valide = IsNumeric AND <> 0.
        If colPrixSpareArch > 0 Then
            Dim vPrixSpare As Variant
            vPrixSpare = CascadePrixSpareDATA_ASL(wsPBS, lignePBS)
            If Not IsEmpty(vPrixSpare) Then wsCible.Cells(ligneCible, colPrixSpareArch).Value = vPrixSpare
        End If

        ' Prix Repair archive : cascade Fournisseur -> Cat D&IS -> Archive
        ' (pas de FCS Repair dans la BdD).
        If colPrixRepairArch > 0 Then
            Dim vPrixRepair As Variant
            vPrixRepair = CascadePrixRepairDATA_ASL(wsPBS, lignePBS)
            If Not IsEmpty(vPrixRepair) Then wsCible.Cells(ligneCible, colPrixRepairArch).Value = vPrixRepair
        End If

NextPBS:
    Next lignePBS

    ' Re-appliquer l alignement vertical centre (les nouvelles lignes en
    ' avaient besoin).
    On Error Resume Next
    wsCible.UsedRange.VerticalAlignment = xlCenter
    On Error GoTo 0

    Debug.Print "SIRUS_EnrichirProjetASL_DepuisDATA_ASL : " & nbOverwrite & " overwrite, " & nbAjout & " ajout(s) depuis DATA_ASL."
End Sub

' Cascade prix Spare DATA_ASL : retourne la 1re valeur non-zero numerique
' selon la priorite : Fournisseur (col Alstom FRS) -> Catalogue D&IS
' (col AA) -> Catalogue FCS calcule (col AB = FCS x 1,0557) -> Archive
' (col AN). Une cellule est "connue" si IsNumeric ET <> 0. Renvoie Empty
' si aucune source ne donne de valeur connue (le pre-fill BdD n est alors
' pas ecrase).
Private Function CascadePrixSpareDATA_ASL(ByVal wsPBS As Worksheet, ByVal lignePBS As Long) As Variant
    Dim v As Variant
    Dim cols As Variant
    Dim i As Long

    cols = Array(DATA_ASL_type.CoutSpare_AlstomFR_FRS, DATA_ASL_type.PrixSpareCatalogue, DATA_ASL_type.PrixSpareCalcule, DATA_ASL_type.PrixSpare_ProjetArchive)

    For i = LBound(cols) To UBound(cols)
        If cols(i) > 0 Then
            v = wsPBS.Cells(lignePBS, cols(i)).Value
            If IsNumeric(v) Then
                If CDbl(v) <> 0 Then
                    CascadePrixSpareDATA_ASL = CDbl(v)
                    Exit Function
                End If
            End If
        End If
    Next i

    CascadePrixSpareDATA_ASL = Empty
End Function

' Cascade prix Repair DATA_ASL : Fournisseur -> Catalogue D&IS -> Archive.
' Pas de FCS Repair dans la BdD. Memes regles que CascadePrixSpareDATA_ASL.
Private Function CascadePrixRepairDATA_ASL(ByVal wsPBS As Worksheet, ByVal lignePBS As Long) As Variant
    Dim v As Variant
    Dim cols As Variant
    Dim i As Long

    cols = Array(DATA_ASL_type.PrixRepair_Alstom_FRS, DATA_ASL_type.PrixRepairCatalogue, DATA_ASL_type.PrixRepair_ProjetArchive)

    For i = LBound(cols) To UBound(cols)
        If cols(i) > 0 Then
            v = wsPBS.Cells(lignePBS, cols(i)).Value
            If IsNumeric(v) Then
                If CDbl(v) <> 0 Then
                    CascadePrixRepairDATA_ASL = CDbl(v)
                    Exit Function
                End If
            End If
        End If
    Next i

    CascadePrixRepairDATA_ASL = Empty
End Function

' Cherche la colonne d une feuille SIRUS par son ColonneBdD (entete ligne 2).
' Renvoie 0 si introuvable. Permet une lookup tolerante a l ordre du config.
Private Function TrouverColParColonneBdD(ByVal ws As Worksheet, ByVal sColBdD As String) As Long
    Dim c As Long, lastCol As Long
    On Error Resume Next
    lastCol = ws.Cells(2, ws.Columns.count).End(xlToLeft).Column
    On Error GoTo 0
    If lastCol < 1 Then lastCol = 1
    For c = 1 To lastCol
        If CStr(ws.Cells(2, c).Value) = sColBdD Then
            TrouverColParColonneBdD = c
            Exit Function
        End If
    Next c
    TrouverColParColonneBdD = 0
End Function

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
    tables = Array("Produit$", "Info_ATM$", "Taches_ATM$", "EMMST$", "Projet_SLI$", "Projet_ASL$")

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

    MsgBox "Feuille '" & FEUILLE & "' generee avec " & (ligne - 2) & " colonnes BdD au total." & vbCrLf & "Filtrer la colonne A pour isoler une table.", vbInformation, "SIRUS Diag"
    Exit Sub

Sortie:
    On Error Resume Next
    Call Close_CONN(ChaineConnexionSIRUS)
    On Error GoTo 0
    MsgBox "Erreur Test_DiagnostiquerColonnesBdDFeuille : " & Err.Description, vbCritical
End Sub

' =========================================================================
' Test_DiagnostiquerEcartsAvecCandidats
' =========================================================================
' Pour chaque ColonneBdD du config absent en BdD, identifie le nom BdD
' le plus proche (plus long prefixe commun) et localise la 1ere position
' ou les caracteres different, avec leurs codes Unicode (AscW).
'
' Resultat dans une feuille "SIRUS_Diag_Diff" :
'   A : Table
'   B : Champ logique
'   C : ColonneBdD du config
'   D : Longueur cfg
'   E : Candidat BdD le plus proche (vrai nom)
'   F : Longueur candidat
'   G : Position 1ere difference (1-based)
'   H : Code Unicode caractere cfg (ex. 39 = ' ASCII, 8217 = ' typographique)
'   I : Code Unicode caractere BdD a cette position
'
' Permet d'identifier les ecarts subtils (apostrophe typographique vs
' ASCII, espace insecable U+00A0 vs espace, NBSP, etc.).
Public Sub Test_DiagnostiquerEcartsAvecCandidats()

    Const FEUILLE As String = "SIRUS_Diag_Diff"

    Dim tables As Variant
    Dim t As Variant
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim ligne As Long
    Dim col As Collection
    Dim def As SIRUS_Champ_Config
    Dim noms As Object
    Dim sSQL As String
    Dim i As Long
    Dim candidat As String
    Dim posDiff As Long

    On Error GoTo Sortie

    Set wb = ActiveWorkbook
    If wb Is Nothing Then
        MsgBox "Aucun classeur actif.", vbExclamation
        Exit Sub
    End If

    ' (Re)creer la feuille
    On Error Resume Next
    Application.DisplayAlerts = False
    wb.Worksheets(FEUILLE).Delete
    Application.DisplayAlerts = True
    On Error GoTo Sortie

    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = FEUILLE

    ' En-tetes
    ws.Cells(1, 1).Value = "Table"
    ws.Cells(1, 2).Value = "Champ logique"
    ws.Cells(1, 3).Value = "ColonneBdD config"
    ws.Cells(1, 4).Value = "Len cfg"
    ws.Cells(1, 5).Value = "Candidat BdD"
    ws.Cells(1, 6).Value = "Len BdD"
    ws.Cells(1, 7).Value = "Pos diff"
    ws.Cells(1, 8).Value = "Unicode cfg"
    ws.Cells(1, 9).Value = "Unicode BdD"
    With ws.Range("A1:I1")
        .Font.Bold = True
        .Interior.Color = COULEUR_ENTETE
    End With

    tables = Array(Array("Produit", "Produit$"), Array("InfoATM", "Info_ATM$"), Array("TachesATM", "Taches_ATM$"), Array("EMMST", "EMMST$"), Array("ProjetSLI", "Projet_SLI$"), Array("ProjetASL", "Projet_ASL$"))

    Call Open_CONN(ChaineConnexionSIRUS)
    ligne = 2

    For Each t In tables
        Set noms = CreateObject("Scripting.Dictionary")
        noms.CompareMode = 1  ' BinaryCompare (sensible casse)

        sSQL = "SELECT * FROM [" & t(1) & "] WHERE 1=0"
        Call open_REQSEL(sSQL)
        If Not (REQSEL Is Nothing) Then
            For i = 0 To REQSEL.Fields.Count - 1
                noms(REQSEL.Fields(i).Name) = True
            Next i
        End If

        Set col = SIRUS_LireConfig(CStr(t(0)))
        For Each def In col
            If LenB(def.ColonneBdD) > 0 And Not noms.Exists(TronquerColonneBdD(def.ColonneBdD)) Then
                candidat = TrouverCandidatBdD(def.ColonneBdD, noms)
                posDiff = PositionPremiereDifference(def.ColonneBdD, candidat)

                ws.Cells(ligne, 1).Value = CStr(t(0))
                ws.Cells(ligne, 2).Value = def.Champ
                ws.Cells(ligne, 3).NumberFormat = "@"
                ws.Cells(ligne, 3).Value = def.ColonneBdD
                ws.Cells(ligne, 4).Value = Len(def.ColonneBdD)
                ws.Cells(ligne, 5).NumberFormat = "@"
                ws.Cells(ligne, 5).Value = candidat
                ws.Cells(ligne, 6).Value = Len(candidat)

                If posDiff > 0 Then
                    ws.Cells(ligne, 7).Value = posDiff
                    If posDiff <= Len(def.ColonneBdD) Then
                        ws.Cells(ligne, 8).Value = AscW(Mid$(def.ColonneBdD, posDiff, 1))
                    Else
                        ws.Cells(ligne, 8).Value = "(absent cfg)"
                    End If
                    If posDiff <= Len(candidat) Then
                        ws.Cells(ligne, 9).Value = AscW(Mid$(candidat, posDiff, 1))
                    Else
                        ws.Cells(ligne, 9).Value = "(absent BdD)"
                    End If
                Else
                    ws.Cells(ligne, 7).Value = "= identique ?!"
                End If
                ligne = ligne + 1
            End If
        Next def
    Next t

    Call Close_CONN(ChaineConnexionSIRUS)

    ws.Columns("A:I").AutoFit
    ws.Columns("C").ColumnWidth = 60
    ws.Columns("E").ColumnWidth = 60
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True

    MsgBox "Feuille '" & FEUILLE & "' generee. " & (ligne - 2) & " ecart(s) identifie(s)." & vbCrLf & vbCrLf & "Codes Unicode courants a surveiller :" & vbCrLf & "  39  = ' apostrophe ASCII" & vbCrLf & "  8217 = ' apostrophe typographique" & vbCrLf & "  32  = espace normal" & vbCrLf & "  160 = espace insecable (NBSP)" & vbCrLf & "  233 = e accent aigu", vbInformation, "SIRUS Diag Diff"
    Exit Sub

Sortie:
    On Error Resume Next
    Call Close_CONN(ChaineConnexionSIRUS)
    On Error GoTo 0
    MsgBox "Erreur Test_DiagnostiquerEcartsAvecCandidats : " & Err.Description, vbCritical
End Sub

' Helpers (privates) pour la comparaison de chaines.
Private Function TrouverCandidatBdD(ByVal sCfg As String, noms As Object) As String
    ' Renvoie le nom BdD avec le plus long prefixe commun avec sCfg.
    Dim k As Variant
    Dim meilleur As String
    Dim meilleurLen As Long
    Dim n As Long
    For Each k In noms.Keys
        n = LongueurPrefixeCommun(sCfg, CStr(k))
        If n > meilleurLen Then
            meilleurLen = n
            meilleur = CStr(k)
        End If
    Next k
    TrouverCandidatBdD = meilleur
End Function

Private Function LongueurPrefixeCommun(ByVal a As String, ByVal b As String) As Long
    Dim i As Long, n As Long
    n = Len(a)
    If Len(b) < n Then n = Len(b)
    For i = 1 To n
        If Mid$(a, i, 1) <> Mid$(b, i, 1) Then
            LongueurPrefixeCommun = i - 1
            Exit Function
        End If
    Next i
    LongueurPrefixeCommun = n
End Function

Private Function PositionPremiereDifference(ByVal a As String, ByVal b As String) As Long
    Dim i As Long, n As Long
    n = Len(a)
    If Len(b) < n Then n = Len(b)
    For i = 1 To n
        If Mid$(a, i, 1) <> Mid$(b, i, 1) Then
            PositionPremiereDifference = i
            Exit Function
        End If
    Next i
    If Len(a) <> Len(b) Then
        PositionPremiereDifference = n + 1
    Else
        PositionPremiereDifference = 0  ' identiques
    End If
End Function

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

    tables = Array("Produit$", "Info_ATM$", "Taches_ATM$", "EMMST$", "Projet_SLI$", "Projet_ASL$")

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

    ' Comparer aux ColonneBdD du config (apres troncature a 64 chars
    ' pour s'aligner sur la limite ACE OLE DB).
    Set col = SIRUS_LireConfig(sTable)
    Debug.Print ""
    Debug.Print "=== Ecarts config <-> BdD pour " & sTable & " ([" & sTableSQL & "]) ==="

    For Each def In col
        If LenB(def.ColonneBdD) > 0 Then
            If noms.Exists(TronquerColonneBdD(def.ColonneBdD)) Then
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
