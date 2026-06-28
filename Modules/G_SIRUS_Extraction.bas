Attribute VB_Name = "G_SIRUS_Extraction"
'--------------------------------------------------------------------------------------------------------
'
'                                       Generateur Livrables SLI
'                                       Module G_SIRUS_Extraction
'                                       Developpement : LE LARGE Thomas (LGM/Alstom)
'                                       SIRUS - Phase 4 entree A (etape D) :
'                                       deverse les feuilles brutes d'extraction ATM/EMMST
'                                       dans les feuilles SIRUS au layout standard.
'                                       Juin 2026
'
'--------------------------------------------------------------------------------------------------------
Option Explicit
Option Compare Text

' =========================================================================
' SIRUS_DeverserExtractionDansSIRUS
' =========================================================================
' Point d'entree publique. Lance Main_Extraction_Params (E_Extraction_Main)
' sur le fichier livrable, puis copie les valeurs des feuilles brutes
' (Info_ATM/Taches_ATM/EMMST) vers les feuilles SIRUS au layout standard
' (Statut col 1 | champs col 2..N+1 | Commentaires col N+2).
'
' Prerequis : les feuilles SIRUS cibles doivent deja exister (creees en
' amont par SIRUS_GenererFeuilleAnalyse(table, SIRUS_OP_EXTRACTION)).
'
' Parametres :
'   sTable    : "ATM" ou "EMMST"
'   sFichier  : chemin complet du fichier livrable
'   sLangue   : "FR" ou "EN"
'   sVersion  : "V1" ou "V2" (informatif, pas utilise par l'extraction
'               actuellement - les fiches DTR ont la meme structure)
Public Sub SIRUS_DeverserExtractionDansSIRUS(ByVal sTable As String, ByVal sFichier As String, ByVal sLangue As String, Optional ByVal sVersion As String = "")

    Dim wbExtr As Workbook
    Dim wbProjet As Workbook
    Dim sLivrable As String

    On Error GoTo Sortie

    Set wbProjet = ActiveWorkbook
    If wbProjet Is Nothing Then
        MsgBox "Aucun classeur projet actif.", vbCritical, "SIRUS Extraction"
        Exit Sub
    End If

    ' ATM (table SIRUS) -> ATM (livrable Extraction). EMMST identique.
    sLivrable = sTable

    ' Lancer l'extraction sur le fichier livrable (sans MsgBox final, on
    ' enchaine sur le deversement). Retourne le wb livrable ouvert avec
    ' les feuilles Info_ATM/Taches_ATM ou EMMST creees.
    Set wbExtr = Main_Extraction_Params(sFichier, sLivrable, sLangue, False)
    If wbExtr Is Nothing Then
        ' Erreur deja signalee par Main_Extraction_Params, on reactive juste
        ' le classeur projet et on sort.
        wbProjet.Activate
        Exit Sub
    End If

    ' Deverser feuilles brutes -> feuilles SIRUS
    Select Case sTable
        Case "ATM"
            Call DeverserFeuilleVersSIRUS(wbExtr, "Info_ATM", wbProjet, "SIRUS_InfoATM_Extraction")
            Call DeverserFeuilleVersSIRUS(wbExtr, "Taches_ATM", wbProjet, "SIRUS_TachesATM_Extraction")
        Case "EMMST"
            Call DeverserFeuilleVersSIRUS(wbExtr, "EMMST", wbProjet, "SIRUS_EMMST_Extraction")
        Case Else
            MsgBox "SIRUS Extraction : table inconnue '" & sTable & "'.", vbCritical
    End Select

    ' Fermer le fichier livrable extrait sans sauvegarder (les feuilles
    ' brutes Info_ATM/Taches_ATM ont ete ajoutees dans wbExtr, mais on ne
    ' veut pas polluer le fichier source).
    wbExtr.Close SaveChanges:=False

    wbProjet.Activate
    MsgBox "Extraction terminee : donnees deversees dans les feuilles SIRUS_*_Extraction.", vbInformation, "SIRUS Extraction"
    Exit Sub

Sortie:
    On Error Resume Next
    If Not wbExtr Is Nothing Then wbExtr.Close SaveChanges:=False
    On Error GoTo 0
    MsgBox "Erreur SIRUS_DeverserExtractionDansSIRUS : " & Err.Description, vbCritical
End Sub

' =========================================================================
' DeverserFeuilleVersSIRUS (helper interne)
' =========================================================================
' Copie les valeurs d'une feuille brute (issue d'extraction) vers une
' feuille SIRUS au layout standard, en mappant les colonnes par leur
' nom BdD.
'
' Layout brute (Extraction_Fiches_ATM/EMMST) :
'   ligne 1 = en-tete BdD (vrais noms)
'   ligne 2..N = donnees
'
' Layout SIRUS :
'   ligne 1 = libelle humain  (col 1 = "Statut", col N+2 = "Commentaires")
'   ligne 2 = ColonneBdD masquee
'   ligne 3..N = donnees
'
' Strategie : pour chaque colonne source (ligne 1 brute), chercher la
' meme valeur dans la ligne 2 SIRUS et copier la colonne au bon endroit.
' Les colonnes brutes sans correspondance SIRUS sont ignorees (avec
' un Debug.Print informatif).
Private Sub DeverserFeuilleVersSIRUS(ByVal wbSrc As Workbook, ByVal sSrcFeuille As String, ByVal wbDest As Workbook, ByVal sDestFeuille As String)
    Dim wsSrc As Worksheet
    Dim wsDest As Worksheet
    Dim mapColBdD As Object       ' Dictionary {NomNormalise -> col index SIRUS}
    Dim mapColBdDOriginal As Object ' Dictionary {NomNormalise -> NomBrut config}
    Dim i As Long, j As Long
    Dim sNomBdD As String, sNomNormalise As String
    Dim nbLignesSrc As Long, nbColsSrc As Long
    Dim ligneSrc As Long, ligneDest As Long
    Dim cellBdD As String
    Dim nbColonnesMappees As Long, nbColonnesIgnorees As Long
    Dim arrSrc As Variant
    Dim valeur As Variant
    Dim colsIgnorees As String

    On Error GoTo Sortie

    ' Localiser les feuilles
    Set wsSrc = Nothing
    Set wsDest = Nothing
    On Error Resume Next
    Set wsSrc = wbSrc.Worksheets(sSrcFeuille)
    Set wsDest = wbDest.Worksheets(sDestFeuille)
    On Error GoTo Sortie

    If wsSrc Is Nothing Then
        MsgBox "SIRUS Extraction : feuille source '" & sSrcFeuille & "' absente du livrable apres extraction." & vbCrLf & "Feuilles disponibles : " & FeuillesDispo(wbSrc), vbCritical, "SIRUS"
        Exit Sub
    End If
    If wsDest Is Nothing Then
        MsgBox "SIRUS Extraction : feuille SIRUS '" & sDestFeuille & "' absente du projet. La generer avant le deversement.", vbCritical, "SIRUS"
        Exit Sub
    End If

    ' Construire le mapping ColonneBdD -> col SIRUS en lisant la ligne 2
    ' de la feuille destination. Cle = nom normalise (trim + sans accents
    ' corrompus) pour tolerer les ecarts mineurs entre le format
    ' Champs_BDD.xlsx (config SIRUS) et le format extraction (E_Extraction_*).
    Set mapColBdD = CreateObject("Scripting.Dictionary")
    Set mapColBdDOriginal = CreateObject("Scripting.Dictionary")
    Dim nbColsDest As Long
    nbColsDest = wsDest.Cells(2, wsDest.Columns.count).End(xlToLeft).Column
    For i = 1 To nbColsDest
        cellBdD = CStr(wsDest.Cells(2, i).Value)
        If LenB(cellBdD) > 0 Then
            sNomNormalise = NormaliserNomBdD(cellBdD)
            mapColBdD(sNomNormalise) = i
            mapColBdDOriginal(sNomNormalise) = cellBdD
        End If
    Next i

    ' Forcer NumberFormat = "@" (texte) sur les colonnes SIRUS dont le
    ' TypeChamp du config = "Texte" (notamment Rang). Empeche la conversion
    ' Excel "1-1" / "5-2" en date au moment de l'ecriture.
    Call ForcerFormatTexte(wsDest, sDestFeuille, mapColBdD)

    ' Lire les en-tetes + donnees brutes en bloc pour reduire les acces COM
    nbColsSrc = wsSrc.Cells(1, wsSrc.Columns.count).End(xlToLeft).Column
    nbLignesSrc = wsSrc.Cells(wsSrc.rows.count, 1).End(xlUp).Row

    If nbLignesSrc < 2 Then
        MsgBox "SIRUS Extraction : feuille '" & sSrcFeuille & "' vide (pas de ligne de donnees apres l'entete).", vbExclamation, "SIRUS"
        Exit Sub
    End If

    arrSrc = wsSrc.Range(wsSrc.Cells(1, 1), wsSrc.Cells(nbLignesSrc, nbColsSrc)).Value

    ' Pour chaque colonne source, mapper vers SIRUS et copier les donnees
    nbColonnesMappees = 0
    nbColonnesIgnorees = 0
    colsIgnorees = ""
    ligneDest = 3  ' SIRUS_LIGNE_DONNEES_DEBUT
    For j = 1 To nbColsSrc
        sNomBdD = CStr(arrSrc(1, j))
        sNomNormalise = NormaliserNomBdD(sNomBdD)
        If mapColBdD.Exists(sNomNormalise) Then
            Dim colDest As Long
            colDest = CLng(mapColBdD(sNomNormalise))
            ' Copier les donnees lignes 2..nbLignesSrc de wsSrc vers
            ' lignes 3..(nbLignesSrc + 1) de wsDest (decalage +1 lie a
            ' la ligne 2 d'entete BdD SIRUS).
            For ligneSrc = 2 To nbLignesSrc
                valeur = arrSrc(ligneSrc, j)
                wsDest.Cells(ligneDest + ligneSrc - 2, colDest).Value = valeur
            Next ligneSrc
            nbColonnesMappees = nbColonnesMappees + 1
        Else
            nbColonnesIgnorees = nbColonnesIgnorees + 1
            colsIgnorees = colsIgnorees & "  - '" & sNomBdD & "'" & vbCrLf
        End If
    Next j

    ' Recap visible : MsgBox avec le bilan + liste des colonnes ignorees
    ' pour permettre a l'utilisateur de diagnostiquer les ecarts entre
    ' config SIRUS et format extraction.
    Dim sMsg As String
    sMsg = "Feuille SIRUS : " & sDestFeuille & vbCrLf & vbCrLf
    sMsg = sMsg & nbColonnesMappees & " colonne(s) mappee(s) sur " & nbColsSrc & " brutes" & vbCrLf
    sMsg = sMsg & (nbLignesSrc - 1) & " ligne(s) de donnees deversee(s)" & vbCrLf
    If nbColonnesIgnorees > 0 Then
        sMsg = sMsg & vbCrLf & "Colonnes ignorees (absentes du config SIRUS) :" & vbCrLf
        sMsg = sMsg & colsIgnorees
        sMsg = sMsg & vbCrLf & "Ces colonnes ne sont pas perdues : elles existent" & vbCrLf
        sMsg = sMsg & "dans le fichier livrable mais pas dans le config SIRUS." & vbCrLf
        sMsg = sMsg & "Pour les inclure : enrichir Peupler" & Replace(sDestFeuille, "SIRUS_", "") & "."
    End If
    Debug.Print "DeverserFeuilleVersSIRUS [" & sDestFeuille & "] : " & nbColonnesMappees & " mappee(s), " & nbColonnesIgnorees & " ignoree(s), " & (nbLignesSrc - 1) & " ligne(s) deversee(s)."
    MsgBox sMsg, vbInformation, "SIRUS - Deversement"

    ' Refresh statut col 1 : afficher "A analyser" en attendant l'analyse
    ' (sinon les lignes deversees ont col 1 vide ce qui fait croire qu'elles
    ' sont vides pour SIRUS_AnalyserFeuille).
    Dim cellSt As Range
    For ligneSrc = 2 To nbLignesSrc
        Set cellSt = wsDest.Cells(ligneDest + ligneSrc - 2, 1)
        cellSt.Value = "A analyser"
        cellSt.Font.Color = RGB(80, 80, 80)
    Next ligneSrc
    Exit Sub

Sortie:
    MsgBox "Erreur DeverserFeuilleVersSIRUS [" & sDestFeuille & "] : " & Err.Description, vbCritical, "SIRUS"
End Sub

' =========================================================================
' ForcerFormatTexte (helper interne)
' =========================================================================
' Pour chaque champ TypeChamp = "Texte" du config SIRUS de la table
' correspondante, force NumberFormat = "@" sur la zone de saisie de la
' colonne SIRUS. Evite la conversion automatique Excel de valeurs
' comme "1-1", "5-2" en date lors du deversement.
'
' Strategie : nom feuille = "SIRUS_<Table>_<Op>" -> on extrait <Table>
' et on lit le config via SIRUS_LireConfig(<Table>).
Private Sub ForcerFormatTexte(ByVal wsDest As Worksheet, ByVal sDestFeuille As String, ByRef mapColBdD As Object)
    Dim parts() As String
    Dim sTable As String
    Dim col As Collection
    Dim def As SIRUS_Champ_Config
    Dim sNomNormalise As String
    Dim colIdx As Long

    ' Extraire le nom de table depuis "SIRUS_<Table>_<Op>"
    parts = Split(sDestFeuille, "_")
    If UBound(parts) < 2 Then Exit Sub
    sTable = parts(1)

    On Error Resume Next
    Set col = SIRUS_LireConfig(sTable)
    On Error GoTo 0
    If col Is Nothing Then Exit Sub

    For Each def In col
        If def.TypeChamp = "Texte" And LenB(def.ColonneBdD) > 0 Then
            sNomNormalise = NormaliserNomBdD(def.ColonneBdD)
            If mapColBdD.Exists(sNomNormalise) Then
                colIdx = CLng(mapColBdD(sNomNormalise))
                ' Forcer texte sur la zone de saisie (ligne 3 jusqu'au bas
                ' raisonnable - on prend 5000 lignes pour couvrir les
                ' extractions volumineuses).
                wsDest.Range(wsDest.Cells(3, colIdx), wsDest.Cells(5000, colIdx)).NumberFormat = "@"
            End If
        End If
    Next def
End Sub

' =========================================================================
' NormaliserNomBdD (helper interne)
' =========================================================================
' Normalise un nom de colonne BdD pour matching tolerant entre config
' SIRUS et format extraction. Strategie :
'   - Trim des espaces avant/apres
'   - Suppression des espaces multiples
'   - Lowercase (matching insensible a la casse)
' Note : Option Compare Text est deja actif au niveau module, donc les
' comparaisons via Dictionary sont deja insensibles a la casse. Cette
' normalisation enleve en plus les variations d'espaces.
Private Function NormaliserNomBdD(ByVal sNom As String) As String
    Dim s As String
    s = Trim$(CStr(sNom))
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    NormaliserNomBdD = LCase$(s)
End Function

' =========================================================================
' FeuillesDispo (helper interne)
' =========================================================================
' Renvoie la liste des feuilles d'un classeur (pour message d'erreur
' diagnostic quand une feuille attendue n'est pas trouvee).
Private Function FeuillesDispo(ByVal wb As Workbook) As String
    Dim ws As Worksheet
    Dim s As String
    For Each ws In wb.Worksheets
        If LenB(s) > 0 Then s = s & ", "
        s = s & ws.Name
    Next ws
    FeuillesDispo = s
End Function
