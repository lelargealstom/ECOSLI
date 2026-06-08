Attribute VB_Name = "G_SIRUS_Config"
'--------------------------------------------------------------------------------------------------------
'
'                                       Generateur Livrables SLI
'                                       Module G_SIRUS_Config
'                                       Developpement : LE LARGE Thomas (LGM/Alstom)
'                                       SIRUS - Phase 0 : feuille de configuration pilote
'                                       (definitions de champ par table x operation)
'                                       cf. brief SIRUS v0.2 sections 6 et 8.
'                                       Juin 2026
'
'--------------------------------------------------------------------------------------------------------
Option Explicit
Option Compare Text

' =========================================================================
' CONSTANTES
' =========================================================================
Public Const SIRUS_CONFIG_SHEET    As String = "SIRUS_Config"
Public Const SIRUS_CONFIG_LISTOBJ  As String = "tblSIRUSConfig"

' Noms de table (valeurs de la colonne "Table" du config)
Public Const SIRUS_TBL_PRODUIT     As String = "Produit"
Public Const SIRUS_TBL_INFOATM     As String = "InfoATM"
Public Const SIRUS_TBL_TACHESATM   As String = "TachesATM"
Public Const SIRUS_TBL_EMMST       As String = "EMMST"
Public Const SIRUS_TBL_PROJETSLI   As String = "ProjetSLI"
Public Const SIRUS_TBL_PROJETASL   As String = "ProjetASL"

' =========================================================================
' DEFINITION D'UN CHAMP DU CONFIG
' =========================================================================
' Cf. classe SIRUS_Champ_Config.cls (13 proprietes alignees sur le brief
' v0.2 section 8). On a choisi une classe plutot qu'un Type UDT car VBA
' interdit le stockage d'un UDT dans une Collection (impossible a passer
' en Variant), donc le For Each plante.

' =========================================================================
' INITIALISATION DE LA FEUILLE SIRUS_Config
' =========================================================================
' Cree la feuille SIRUS_Config (veryHidden) avec son ListObject si absente,
' et la peuple pour la table Produit (pilote Phase 0).
' Idempotent : si la feuille existe deja et bForcer = False, ne fait rien.
' bForcer = True ==> suppression + recreation (utile pour iteration en dev).
Public Sub SIRUS_InitConfigSheet(Optional ByVal bForcer As Boolean = False)

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim entetes As Variant
    Dim deja As Boolean

    On Error GoTo Sortie

    ' La config vit dans le complement ECOSLI.xlam (ThisWorkbook = .xlam).
    ' Convention : ActiveWorkbook est reserve a la PBS du classeur projet ;
    ' la config interne du complement est bien ThisWorkbook.
    Set wb = ThisWorkbook

    ' -- 1. Verifier si la feuille existe deja ----------------
    deja = False
    For Each ws In wb.Worksheets
        If ws.Name = SIRUS_CONFIG_SHEET Then
            deja = True
            Exit For
        End If
    Next ws

    If deja And Not bForcer Then
        Debug.Print "SIRUS_InitConfigSheet : la feuille " & SIRUS_CONFIG_SHEET & _
                    " existe deja. Appeler avec bForcer:=True pour reinitialiser."
        Exit Sub
    End If

    ' -- 2. Creer (ou re-creer) la feuille --------------------
    If deja Then
        Application.DisplayAlerts = False
        wb.Worksheets(SIRUS_CONFIG_SHEET).Visible = xlSheetVisible
        wb.Worksheets(SIRUS_CONFIG_SHEET).Delete
        Application.DisplayAlerts = True
    End If

    Set ws = wb.Worksheets.Add
    ws.Name = SIRUS_CONFIG_SHEET

    ' -- 3. Ecrire l'entete (13 colonnes, brief section 8) ----
    entetes = Array("Table", "Champ", "ColonneBdD", "Ordre", "TypeChamp", _
                    "ListeSource", "Obligatoire", "ConditionObligatoire", _
                    "Source", "SourceHeritage", "PoidsMeQ", "RegleAlerte", "Commentaire")
    ws.Range("A1").Resize(1, UBound(entetes) - LBound(entetes) + 1).Value = entetes

    ' -- 4. Creer le ListObject -------------------------------
    Set lo = ws.ListObjects.Add(xlSrcRange, _
                                 ws.Range("A1").Resize(2, UBound(entetes) - LBound(entetes) + 1), _
                                 , xlYes)
    lo.Name = SIRUS_CONFIG_LISTOBJ
    ' Vider la ligne creee automatiquement par xlYes
    If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.ClearContents

    ' -- 5. Peupler pour Produit (pilote Phase 0) ------------
    Call PeuplerProduit(lo)

    ' Si une ligne vide a survecu en debut, on la supprime
    If lo.ListRows.Count > 0 Then
        If LenB(CStr(lo.ListRows(1).Range(1, 1).Value)) = 0 Then
            lo.ListRows(1).Delete
        End If
    End If

    ' -- 6. Cacher la feuille (veryHidden) -------------------
    ws.Visible = xlSheetVeryHidden

    Debug.Print "SIRUS_InitConfigSheet : feuille " & SIRUS_CONFIG_SHEET & _
                " creee et peuplee Produit (" & lo.ListRows.Count & " champs)."
    Exit Sub

Sortie:
    MsgBox "Erreur dans SIRUS_InitConfigSheet : " & Err.Description, vbCritical
End Sub

' =========================================================================
' PEUPLEMENT DU CONFIG POUR LA TABLE PRODUIT
' =========================================================================
' Champs definis a partir du brief v0.2 section 6.1.
' Obligatoires simples : DTR, Designation manuelle, Node Type, Fournisseur,
' Reference fournisseur, Numero plan, MoQ Spare.
' Conditionnels : MTBF intrinseque / Motif NA (au moins un des deux),
' Prix/Date validite Spare et Repair selon Node Type et presence prix manuel.
Private Sub PeuplerProduit(lo As ListObject)

    Call AjouterChamp(lo, "Produit", "DTR", "DTR", 10, _
                      "Texte", "", "Oui", "", "Saisie", "", 3, "", _
                      "Identifiant unique du produit (cle).")

    Call AjouterChamp(lo, "Produit", "Designation manuelle", "DesignationManuelle", 20, _
                      "Texte", "", "Oui", "", "Saisie", "", 2, "", "")

    Call AjouterChamp(lo, "Produit", "Node Type", "NodeType", 30, _
                      "Liste", "NodeType", "Oui", "", "Saisie", "", 2, "", "")

    Call AjouterChamp(lo, "Produit", "Fournisseur", "Fournisseur", 40, _
                      "Texte", "", "Oui", "", "Saisie", "", 1, "", "")

    Call AjouterChamp(lo, "Produit", "Reference fournisseur", "ReferenceFournisseur", 50, _
                      "Texte", "", "Oui", "", "Saisie", "", 1, "", "")

    Call AjouterChamp(lo, "Produit", "Numero plan", "NumeroPlan", 60, _
                      "Texte", "", "Oui", "", "Saisie", "", 1, "", "")

    Call AjouterChamp(lo, "Produit", "MoQ Spare", "MoQSpare", 70, _
                      "Nombre", "", "Oui", "", "Saisie", "", 1, "", _
                      "Minimum Order Quantity Spare.")

    Call AjouterChamp(lo, "Produit", "MTBF intrinseque", "MTBFIntrinseque", 80, _
                      "Nombre", "", "Conditionnel", "MotifNA=""""", "Saisie", "", 2, "", _
                      "Obligatoire sauf si Motif NA renseigne.")

    Call AjouterChamp(lo, "Produit", "Motif NA", "MotifNA", 90, _
                      "Texte", "", "Conditionnel", "MTBFIntrinseque=""""", "Saisie", "", 1, "", _
                      "Obligatoire sauf si MTBF intrinseque renseigne.")

    Call AjouterChamp(lo, "Produit", "Prix Spare manuel", "PrixSpareManuel", 100, _
                      "Nombre", "", "Conditionnel", "NodeType IN (LRU,SRU)", "Saisie", "", 1, "", _
                      "Obligatoire pour LRU/SRU ; au moins 1 prix Spare sur 3 pour LRUNR/SRUNR/KIT/VISSERIE/ASSEMBLY.")

    Call AjouterChamp(lo, "Produit", "Prix Repair manuel", "PrixRepairManuel", 110, _
                      "Nombre", "", "Conditionnel", "NodeType IN (LRU,SRU)", "Saisie", "", 1, "", _
                      "Obligatoire pour LRU/SRU.")

    Call AjouterChamp(lo, "Produit", "Date validite Spare", "DateValiditeSpare", 120, _
                      "Date", "", "Conditionnel", "PrixSpareManuel<>""""", "Saisie", "", 1, "", _
                      "Attendue des qu'un prix Spare manuel est saisi.")

    Call AjouterChamp(lo, "Produit", "Date validite Repair", "DateValiditeRepair", 130, _
                      "Date", "", "Conditionnel", "PrixRepairManuel<>""""", "Saisie", "", 1, "", _
                      "Attendue des qu'un prix Repair manuel est saisi.")

    Call AjouterChamp(lo, "Produit", "Designation BdD", "DesignationBdD", 200, _
                      "Texte", "", "Non", "", "Auto", "", 0, "", _
                      "Lecture seule, recuperee de la BdD (affichage gris).")
End Sub

' =========================================================================
' AJOUT D'UNE LIGNE AU LISTOBJECT DU CONFIG
' =========================================================================
Private Sub AjouterChamp(lo As ListObject, _
                         ByVal sTable As String, _
                         ByVal sChamp As String, _
                         ByVal sColonneBdD As String, _
                         ByVal lOrdre As Long, _
                         ByVal sTypeChamp As String, _
                         ByVal sListeSource As String, _
                         ByVal sObligatoire As String, _
                         ByVal sConditionObligatoire As String, _
                         ByVal sSource As String, _
                         ByVal sSourceHeritage As String, _
                         ByVal dPoidsMeQ As Double, _
                         ByVal sRegleAlerte As String, _
                         ByVal sCommentaire As String)

    Dim r As ListRow
    Set r = lo.ListRows.Add

    r.Range(1, 1).Value = sTable
    r.Range(1, 2).Value = sChamp
    r.Range(1, 3).Value = sColonneBdD
    r.Range(1, 4).Value = lOrdre
    r.Range(1, 5).Value = sTypeChamp
    r.Range(1, 6).Value = sListeSource
    r.Range(1, 7).Value = sObligatoire
    r.Range(1, 8).Value = sConditionObligatoire
    r.Range(1, 9).Value = sSource
    r.Range(1, 10).Value = sSourceHeritage
    r.Range(1, 11).Value = dPoidsMeQ
    r.Range(1, 12).Value = sRegleAlerte
    r.Range(1, 13).Value = sCommentaire
End Sub

' =========================================================================
' LECTURE DU CONFIG POUR UNE TABLE DONNEE
' =========================================================================
' Renvoie une Collection de Champ_Config pour la table demandee, triee par
' Ordre croissant. Si la feuille n'existe pas ou si la table est inconnue,
' renvoie une Collection vide et trace un message dans la fenetre Immediate.
' Aucune erreur bloquante pour l'utilisateur (lecture diagnostique).
Public Function SIRUS_LireConfig(ByVal sTable As String) As Collection

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim r As ListRow
    Dim col As Collection
    Dim def As SIRUS_Champ_Config
    Dim deja As Boolean
    Dim i As Long
    Dim inserted As Boolean

    Set col = New Collection

    On Error GoTo Sortie

    Set wb = ThisWorkbook

    ' -- Verifier presence de la feuille --------------------
    deja = False
    For Each ws In wb.Worksheets
        If ws.Name = SIRUS_CONFIG_SHEET Then
            deja = True
            Exit For
        End If
    Next ws

    If Not deja Then
        Debug.Print "SIRUS_LireConfig : feuille " & SIRUS_CONFIG_SHEET & _
                    " introuvable. Lancer SIRUS_InitConfigSheet d'abord."
        Set SIRUS_LireConfig = col
        Exit Function
    End If

    Set ws = wb.Worksheets(SIRUS_CONFIG_SHEET)
    Set lo = ws.ListObjects(SIRUS_CONFIG_LISTOBJ)

    If lo.ListRows.Count = 0 Then
        Set SIRUS_LireConfig = col
        Exit Function
    End If

    ' -- Parcours des lignes et filtrage par table ----------
    ' Insertion triee par Ordre via Collection.Add ... Before
    ' Important : une NOUVELLE instance par ligne (sinon toutes les
    ' entrees de la Collection pointent vers le meme objet).
    For Each r In lo.ListRows
        If CStr(r.Range(1, 1).Value) = sTable Then
            Set def = New SIRUS_Champ_Config
            def.Table = CStr(r.Range(1, 1).Value)
            def.Champ = CStr(r.Range(1, 2).Value)
            def.ColonneBdD = CStr(r.Range(1, 3).Value)
            def.Ordre = NzLng(r.Range(1, 4).Value)
            def.TypeChamp = CStr(r.Range(1, 5).Value)
            def.ListeSource = CStr(r.Range(1, 6).Value)
            def.Obligatoire = CStr(r.Range(1, 7).Value)
            def.ConditionObligatoire = CStr(r.Range(1, 8).Value)
            def.Source = CStr(r.Range(1, 9).Value)
            def.SourceHeritage = CStr(r.Range(1, 10).Value)
            def.PoidsMeQ = NzDbl(r.Range(1, 11).Value)
            def.RegleAlerte = CStr(r.Range(1, 12).Value)
            def.Commentaire = CStr(r.Range(1, 13).Value)

            inserted = False
            For i = 1 To col.Count
                If col(i).Ordre > def.Ordre Then
                    col.Add def, , i
                    inserted = True
                    Exit For
                End If
            Next i
            If Not inserted Then col.Add def
        End If
    Next r

    Set SIRUS_LireConfig = col
    Exit Function

Sortie:
    Debug.Print "SIRUS_LireConfig : erreur " & Err.Description
    Set SIRUS_LireConfig = col
End Function

' =========================================================================
' HELPERS PRIVES
' =========================================================================
Private Function NzLng(v As Variant) As Long
    If IsNumeric(v) Then
        NzLng = CLng(v)
    Else
        NzLng = 0
    End If
End Function

Private Function NzDbl(v As Variant) As Double
    If IsNumeric(v) Then
        NzDbl = CDbl(v)
    Else
        NzDbl = 0
    End If
End Function

' =========================================================================
' PROCEDURES DE TEST (executables depuis VBE)
' =========================================================================

' Test_InitConfigSheet : (re)cree la feuille SIRUS_Config et la peuple Produit.
' Lancer depuis VBE apres import du module.
Public Sub Test_InitConfigSheet()
    Call SIRUS_InitConfigSheet(bForcer:=True)
End Sub

' Test_LireConfig : lit le config pour la table Produit et trace les
' definitions dans la fenetre Immediate.
Public Sub Test_LireConfig()
    Dim col As Collection
    Dim def As SIRUS_Champ_Config
    Dim s As String

    Set col = SIRUS_LireConfig(SIRUS_TBL_PRODUIT)

    If col.Count = 0 Then
        Debug.Print "Test_LireConfig : aucun champ pour la table Produit."
        Exit Sub
    End If

    Debug.Print "Test_LireConfig : " & col.Count & " champs lus pour Produit."
    For Each def In col
        s = "  Ordre=" & def.Ordre & " | " & def.Champ & _
            " (BdD=" & def.ColonneBdD & ", Type=" & def.TypeChamp & _
            ", Oblig=" & def.Obligatoire & ", Source=" & def.Source & _
            ", MeQ=" & def.PoidsMeQ & ")"
        If LenB(def.ConditionObligatoire) > 0 Then
            s = s & " [cond: " & def.ConditionObligatoire & "]"
        End If
        Debug.Print s
    Next def
End Sub
