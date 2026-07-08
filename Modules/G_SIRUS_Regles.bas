'--------------------------------------------------------------------------------------------------------
'
'                                       Generateur Livrables SLI
'                                       Module G_SIRUS_Regles
'                                       Developpement : LE LARGE Thomas (LGM/Alstom)
'                                       SIRUS - Phase 2 : regles en dur (D1)
'                                       Pas de mini-langage : Select Case par (Table, ColonneBdD).
'                                       Juin 2026
'
'--------------------------------------------------------------------------------------------------------
Option Explicit
Option Compare Text

' =========================================================================
' CACHES PRIVES
' =========================================================================
' Cache des projets connus en ProjetSLI (BdD TRAVAIL). Chargement paresseux
' au 1er appel, reset via SIRUS_ResetCacheProjetsSLI (a appeler apres un
' refresh BdD ou re-import config). Cle = Nom projet en UCase trim.
Private projetsSLI_Cache As Object   ' Scripting.Dictionary

' =========================================================================
' API publique
' =========================================================================
'
' Les regles SIRUS sont evaluees a partir de la ligne courante en cours
' d'analyse, materialisee par un Dictionary {ColonneBdD -> valeur} (passe
' en parametre 'ligne'). Cela permet aux regles d'inspecter les autres
' champs de la meme ligne (ex. NodeType pour conditionner l'obligation
' d'un Prix SPARE manuel).
'
' Deux fonctions :
'   - SIRUS_RegleObligatoire : True si le champ DOIT etre rempli sur cette
'     ligne (regle conditionnelle - les "Obligatoire = Oui" du config sont
'     deja gerees en amont par G_SIRUS_Analyse).
'   - SIRUS_RegleAlerte : "" si pas d'alerte, sinon libelle a afficher dans
'     la colonne Statut. Sert aux coherences metier (ex. MTBF intrinseque
'     vide et MotifNonApplicabilite vide).
'
' Les regles sont rangees dans des Select Case par (Table | ColonneBdD).
' Pour ajouter une regle :
'   1. Identifier la table et la ColonneBdD exacte (cf. SIRUS_Config).
'   2. Ajouter un Case "<Table>|<ColonneBdD>" dans la bonne fonction.
'   3. Documenter dans un commentaire la regle metier verrouillee.

Public Function SIRUS_RegleObligatoire(ByVal sTable As String, ByVal sColonneBdD As String, ByRef ligne As Object) As Boolean
    ' Renvoie True si le champ est obligatoire pour cette ligne.
    ' ligne : Dictionary {ColonneBdD -> valeur} de la ligne en cours.

    ' ----------------- Produit -----------------
    ' Logique factorisee sur les 19 champs MeQ v0.2 (refonte 2026-07-07) :
    '   Base 3 champs toujours obligatoires -> DTR, Designation, Node type.
    '   Si MotifNonApplicabilite_PRODUIT = "NA" -> rien d'autre n'est obligatoire.
    '   Sinon (Motif = "OK") : 16 champs restants obligatoires avec exemptions :
    '     * DVL HARDWARE oblig ssi MotifPreventif in {HARD, BOTH}
    '     * DVL SOFTWARE oblig ssi MotifPreventif in {SOFT, BOTH}
    '     * 3 champs REPAIR non obligatoires si NodeType in {LRUNR, SRUNR}
    If sTable = "Produit" Then
        SIRUS_RegleObligatoire = EstObligatoire_Produit(sColonneBdD, ligne)
        Exit Function
    End If

    Select Case sTable & "|" & sColonneBdD

        ' ----------------- TachesATM -----------------
        Case "TachesATM|Description_TACHEATM"
            ' Description obligatoire pour les taches de rang 3.* (Depose)
            ' et 4.* (Pose) qui sont les operations elementaires detaillees.
            Dim rang As String
            rang = CStr(GetVal(ligne, "Rang_TACHEATM"))
            SIRUS_RegleObligatoire = (Left$(rang, 2) = "3." Or Left$(rang, 2) = "4.")

        Case "TachesATM|CriteresDimensionnels_TACHEATM"
            ' Criteres dimensionnels (Critere Nm) obligatoires si la
            ' description mentionne "Torque Tighten" (couple de serrage).
            Dim descTorque As String
            descTorque = CStr(GetVal(ligne, "Description_TACHEATM"))
            SIRUS_RegleObligatoire = (InStr(1, descTorque, "Torque Tighten", vbTextCompare) > 0)

        Case "TachesATM|OutillageStandard_TACHEATM", "TachesATM|OutillageSpecifique_TACHEATM"
            ' Outillage standard/specifique obligatoire pour les taches
            ' de rang 2.1 (Outils prerequis Depose) et 5.1 (Outils
            ' prerequis Pose) sur les composants >= SRU (LRU exclus :
            ' demontage complet du LRU, pas d'outillage detaille).
            Dim rangO As String, ntO As String
            rangO = CStr(GetVal(ligne, "Rang_TACHEATM"))
            ntO = UCase$(Trim$(CStr(GetVal(ligne, "Node type Produit (LRU/SRU...)_PRODUIT"))))
            SIRUS_RegleObligatoire = ((rangO = "2.1" Or rangO = "5.1") And (ntO = "SRU" Or ntO = "SRUNR" Or ntO = "ASSEMBLY"))

        ' ----------------- InfoATM -----------------
        ' TODO Phase 2.1 : Plan, Reparabilite, Poids, Temps generique
        '   (en grande partie deja Obligatoire=Oui en config).

        ' ----------------- EMMST -----------------
        ' TODO Phase 2.1 : Date peremption obligatoire si Reutilisable=Non,
        '   Outillage specifique transport obligatoire si Poids > seuil, etc.

        ' ----------------- ProjetSLI -----------------
        ' TODO Phase 2.1 : interlocuteurs DL obligatoires selon scope, etc.

        ' ----------------- ProjetASL -----------------
        Case "ProjetASL|MTBF fonctionnel (h) issu du RAM report projet_PROJETASL"
            ' MTBF fonctionnel obligatoire post-CGR (controle metier deja
            ' present dans Main_ATM via gate). En SIRUS on retourne True
            ' systematiquement : le moteur doit toujours pousser pour la
            ' completude (l'utilisateur peut commenter dans le ticket
            ' ATMAP si justification).
            SIRUS_RegleObligatoire = True

        Case "ProjetASL|Prix repair (DEVISE du projet (SLI))_PROJETASL", "ProjetASL|Lead time REPAIR (jours)_PROJETASL"
            ' Prix Repair + Lead time Repair : obligatoires uniquement si
            ' le DTR est un LRU ou un SRU (les autres NodeType n ont pas
            ' de cycle reparation - KIT, VISSERIE, ASSEMBLY, etc.).
            ' Source : regle metier ProjetASL validee Thomas 2026-06-11.
            Dim ntPjAsl As String
            ntPjAsl = UCase$(Trim$(CStr(GetVal(ligne, "Node typed PROJET_PROJETASL"))))
            SIRUS_RegleObligatoire = (ntPjAsl = "LRU" Or ntPjAsl = "SRU")

        Case Else
            ' Defaut : pas de regle conditionnelle definie -> False.
            ' (Les champs Obligatoire=Oui du config sont traites en amont
            '  par SIRUS_AnalyserFeuille, pas par cette fonction.)
            SIRUS_RegleObligatoire = False
    End Select
End Function

Public Function SIRUS_RegleAlerte(ByVal sTable As String, ByVal sColonneBdD As String, ByRef ligne As Object) As String
    ' Renvoie "" si pas d'alerte, sinon le libelle a afficher dans Statut.

    Select Case sTable & "|" & sColonneBdD

        ' ----------------- Produit -----------------
        ' Note 2026-06-10 : l'ancienne alerte "prix sans date de validite"
        ' a ete retiree apres refonte (champs Date Source = "Non" obligatoire,
        ' libre saisie). A reintroduire au besoin sur les nouveaux noms BdD
        ' une fois "??_DateSourceSpare/Repair_a_definir" remplaces.

        ' ----------------- TachesATM -----------------
        Case "TachesATM|NbOperateurs_TACHEATM"
            ' Alerte si > 4 operateurs : valeur metier rare, a confirmer.
            Dim nbOp As String
            nbOp = CStr(GetVal(ligne, "NbOperateurs_TACHEATM"))
            If IsNumeric(nbOp) Then
                If CDbl(nbOp) > 4 Then
                    SIRUS_RegleAlerte = "Nb operateurs > 4 : valeur a confirmer"
                End If
            End If

        ' ----------------- ProjetASL -----------------
        Case "ProjetASL|MTBF fonctionnel (h) issu du RAM report projet_PROJETASL"
            ' Alerte si MTBF fonctionnel < MTBF intrinseque du Produit
            ' (le RAM report projet doit etre >= valeur datasheet).
            Dim mtbfF As String, mtbfI As String
            mtbfF = CStr(GetVal(ligne, "MTBF fonctionnel (h) issu du RAM report projet_PROJETASL"))
            mtbfI = CStr(GetVal(ligne, "MTBF intrinseque Produit (h) issu de la datasheet _PRODUIT"))
            If IsNumeric(mtbfF) And IsNumeric(mtbfI) Then
                If CDbl(mtbfF) < CDbl(mtbfI) Then
                    SIRUS_RegleAlerte = "MTBF fonctionnel < MTBF intrinseque : verifier RAM report"
                End If
            End If

        Case "ProjetASL|PROJET_PROJETASL"
            ' Alerte si le projet ProjetASL n existe pas dans ProjetSLI.
            ' Critere bloquant a l envoi : impossible de creer un ProjetASL
            ' pour un code projet inconnu en SLI (cf. SIRUS_UI_BtnEnvoyerClick).
            ' Le blocage final est cote envoi, l alerte ici sert au radar UI.
            Dim sProjAsl As String, dictProj As Object
            sProjAsl = Trim$(CStr(GetVal(ligne, "PROJET_PROJETASL")))
            If LenB(sProjAsl) > 0 Then
                Set dictProj = SIRUS_ProjetsSLI()
                If Not dictProj Is Nothing Then
                    If Not dictProj.Exists(UCase$(sProjAsl)) Then
                        SIRUS_RegleAlerte = "Projet inconnu dans ProjetSLI : import impossible"
                    End If
                End If
            End If

        Case Else
            SIRUS_RegleAlerte = ""
    End Select
End Function

' =========================================================================
' SIRUS_MeQLigne : calcul unifie du MeQ pour une ligne
' =========================================================================
' Source unique de verite pour le calcul MeQ. Utilise par :
'   - G_SIRUS_PBS.CalculerMeQDTR_Optim (ListView Generation)
'   - G_SIRUS_Analyse.SIRUS_AnalyserFeuille (analyse feuille SIRUS)
'
' Formule : MeQ = nbRempli / nbObl
'   nbObl = nombre de champs obligatoires POUR CETTE LIGNE (variable selon
'           les regles conditionnelles ex. matrice Produit v0.2 : 3/14..19)
'   nbRempli = nombre de ces obligatoires effectivement remplis
'
' Comptage 1-champ = 1-unite (PoidsMeQ ignore : la ponderation n'est pas
' active pour l'instant, decision Thomas Q5 2026-07-07).
'
' Skip :
'   - champs sans ColonneBdD (infos pures, heritages du config)
'   - placeholders ??_*_a_definir
'   - Source in {Auto, Herite} (rempli automatiquement ou herite -> pas
'     concerne par la saisie utilisateur)
'   - Source = "SnapShotPrix" est INCLUS dans le comptage (source virtuelle
'     mais compte comme un champ MeQ).
'
' Retour : Dictionary {
'   nbObl       As Long
'   nbRempli    As Long
'   taux        As Double  (nbRempli / nbObl, 1.0 si nbObl = 0)
'   manquants   As Collection (de SIRUS_Champ_Config, les obligatoires vides)
' }
Public Function SIRUS_MeQLigne(ByRef ligneDict As Object, ByVal sTable As String, _
                                ByVal configCol As Collection) As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    result("nbObl") = 0&
    result("nbRempli") = 0&
    result("taux") = 1#

    Dim manquants As Collection
    Set manquants = New Collection

    If configCol Is Nothing Or ligneDict Is Nothing Then
        Set result("manquants") = manquants
        Set SIRUS_MeQLigne = result
        Exit Function
    End If

    Dim def As SIRUS_Champ_Config
    Dim nbObl As Long, nbRempli As Long
    Dim bOblig As Boolean
    Dim sVal As String

    For Each def In configCol
        If LenB(def.ColonneBdD) = 0 Then GoTo NextDef
        If Left$(def.ColonneBdD, 3) = "??_" Then GoTo NextDef
        If def.Source = "Auto" Or def.Source = "Herite" Then GoTo NextDef

        bOblig = False
        Select Case def.Obligatoire
            Case "Oui":          bOblig = True
            Case "Conditionnel": bOblig = SIRUS_RegleObligatoire(sTable, def.ColonneBdD, ligneDict)
        End Select

        If Not bOblig Then GoTo NextDef

        nbObl = nbObl + 1

        sVal = CStr(GetVal(ligneDict, def.ColonneBdD))
        If LenB(Trim$(sVal)) > 0 Then
            nbRempli = nbRempli + 1
        Else
            manquants.Add def
        End If
NextDef:
    Next def

    result("nbObl") = nbObl
    result("nbRempli") = nbRempli
    If nbObl > 0 Then
        result("taux") = nbRempli / nbObl
    Else
        result("taux") = 1#
    End If
    Set result("manquants") = manquants
    Set SIRUS_MeQLigne = result
End Function

' =========================================================================
' SIRUS_ProjetsSLI : cache des projets connus en BdD ProjetSLI
' =========================================================================
' Chargement paresseux au 1er appel. Renvoie un Dictionary {NomProjetUCase
' -> Empty}. Utilise par SIRUS_RegleAlerte (ProjetASL) et par
' SIRUS_UI_BtnEnvoyerClick (blocage envoi). Reset via la sub publique
' SIRUS_ResetCacheProjetsSLI.
Public Function SIRUS_ProjetsSLI() As Object
    If projetsSLI_Cache Is Nothing Then
        Set projetsSLI_Cache = ChargerProjetsSLI()
    End If
    Set SIRUS_ProjetsSLI = projetsSLI_Cache
End Function

Public Sub SIRUS_ResetCacheProjetsSLI()
    Set projetsSLI_Cache = Nothing
End Sub

Private Function ChargerProjetsSLI() As Object
    ' Charge tous les Nom projet de Projet_SLI$ (BdD TRAVAIL) dans un
    ' Dictionary. Retourne Nothing si la connexion echoue (l alerte
    ' projet inconnu sera alors silencieusement skip - on ne casse pas
    ' l analyse).
    '
    ' Pourquoi Projet_SLI (et pas Projet_ASL) ?
    ' ----------------------------------------
    ' La validation metier est : "un ProjetASL ne peut exister que pour
    ' un projet deja declare au niveau SLI". Projet_SLI est la table-mere
    ' (referentiel projet), Projet_ASL est la table-fille (ligne par DTR
    ' rattachee a un projet). C est une integrite referentielle : avant
    ' d ajouter une ligne dans la table-fille, on verifie que la cle existe
    ' dans la table-mere. Donc on liste les projets connus en SLI pour
    ' bloquer toute Creation ProjetASL referencant un projet inconnu.
    Dim dict As Object
    Dim cnx As String
    Dim sql As String
    Dim bOpened As Boolean

    On Error GoTo Sortie

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1   ' TextCompare (case insensitive)

    cnx = Replace(CONNEXION_EXCEL_BDD_TRAVAIL, "UserProfile", Environ$("UserProfile"))
    Call Open_CONN(cnx)
    bOpened = True

    sql = "SELECT [Nom projet_PROJET SLI] FROM [Projet_SLI$] WHERE [Nom projet_PROJET SLI] IS NOT NULL"
    Call open_REQSEL(sql)

    Do While Not REQSEL.EOF
        Dim sNom As String
        sNom = UCase$(Trim$(CStr(REQSEL.Fields(0).Value)))
        If LenB(sNom) > 0 Then
            If Not dict.Exists(sNom) Then dict.Add sNom, Empty
        End If
        REQSEL.MoveNext
    Loop

    Call Close_CONN(cnx)
    Set ChargerProjetsSLI = dict
    Exit Function

Sortie:
    Debug.Print "ChargerProjetsSLI erreur : " & Err.description
    If bOpened Then
        On Error Resume Next
        Call Close_CONN(cnx)
        On Error GoTo 0
    End If
    Set ChargerProjetsSLI = Nothing
End Function

' =========================================================================
' HELPERS PRIVES
' =========================================================================

Private Function GetVal(ByRef ligne As Object, ByVal sColonneBdD As String) As Variant
    ' Lecture safe d'une valeur dans le Dictionary ligne. Renvoie "" si
    ' la colonne n'existe pas (cas des champs ColonneBdD vide du config).
    If ligne Is Nothing Then
        GetVal = ""
    ElseIf ligne.Exists(sColonneBdD) Then
        GetVal = ligne(sColonneBdD)
        If IsNull(GetVal) Then GetVal = ""
    Else
        GetVal = ""
    End If
End Function

Private Function EstNodeTypeIn(ByRef ligne As Object, ByVal sColonneBdD As String, ByRef tValeurs As Variant) As Boolean
    ' Test si la valeur du champ NodeType de la ligne est dans tValeurs.
    Dim v As String
    Dim i As Long
    v = UCase$(Trim$(CStr(GetVal(ligne, sColonneBdD))))
    For i = LBound(tValeurs) To UBound(tValeurs)
        If v = UCase$(CStr(tValeurs(i))) Then
            EstNodeTypeIn = True
            Exit Function
        End If
    Next i
    EstNodeTypeIn = False
End Function

' =========================================================================
' EstObligatoire_Produit
' =========================================================================
' Applique la matrice MeQ v0.2 sur les 19 champs Produit :
'   - Base (3) : DTR, Designation, Node type -> toujours True.
'   - Si MotifNonApplicabilite = "NA" : rien d'autre n'est obligatoire.
'   - Sinon : les 16 autres champs sont obligatoires, avec 3 exemptions :
'       * DVL HARDWARE ssi MotifPreventif in {HARD, BOTH}
'       * DVL SOFTWARE ssi MotifPreventif in {SOFT, BOTH}
'       * 3 champs REPAIR : non obligatoires si NodeType in {LRUNR, SRUNR}
'   - Autres champs Produit hors 19 (Famille Gen, Ref RAM, Ref datasheet,
'     Ref MUL install/soft/interface) : jamais obligatoires ici (le config
'     les met deja Obligatoire=Non, donc en pratique on n'entre pas ici).
Private Function EstObligatoire_Produit(ByVal sColonneBdD As String, ByRef ligne As Object) As Boolean
    ' Noms BdD avec caracteres non-ASCII construits via ChrW pour eviter
    ' toute normalisation par l'editeur VBA / re-import du .bas :
    '   NodeType : "...LRU/SRU" + U+2026 (HORIZONTAL ELLIPSIS) + ")..."
    '   Motif_Non... : underscore apres "Motif" (nom BdD 2026-07-07).
    Dim sColNodeType As String
    sColNodeType = "Node type Produit (LRU/SRU" & ChrW(&H2026) & ")_PRODUIT"
    Const COL_MOTIFNA   As String = "MotifNonApplicabilite_PRODUIT"
    Const COL_MOTIFPREV As String = "Motif_NonApplicabilitePreventif_PRODUIT"

    ' Base : toujours obligatoire (meme si MotifNA = NA).
    Select Case sColonneBdD
        Case "DTR", "Designation_PRODUIT", sColNodeType
            EstObligatoire_Produit = True
            Exit Function
    End Select

    ' Lecture des 3 champs qui pilotent les regles.
    Dim motifNA As String, motifPrev As String, nodeType As String
    motifNA = UCase$(Trim$(CStr(GetVal(ligne, COL_MOTIFNA))))
    motifPrev = UCase$(Trim$(CStr(GetVal(ligne, COL_MOTIFPREV))))
    nodeType = UCase$(Trim$(CStr(GetVal(ligne, sColNodeType))))

    ' Si MotifNA = "NA" : plus rien n'est obligatoire (seule la base compte).
    If motifNA = "NA" Then
        EstObligatoire_Produit = False
        Exit Function
    End If

    ' DVL HARDWARE : oblig ssi MotifPreventif in {HARD, BOTH}.
    If sColonneBdD = "Duree de vie generique Produit HARDWARE (annee)_PRODUIT" Then
        EstObligatoire_Produit = (motifPrev = "HARD" Or motifPrev = "BOTH")
        Exit Function
    End If

    ' DVL SOFTWARE : oblig ssi MotifPreventif in {SOFT, BOTH}.
    If sColonneBdD = "Duree de vie generique Produit (annee) SOFTWARE _PRODUIT" Then
        EstObligatoire_Produit = (motifPrev = "SOFT" Or motifPrev = "BOTH")
        Exit Function
    End If

    ' 3 champs REPAIR : non obligatoires si NodeType in {LRUNR, SRUNR}.
    Select Case sColonneBdD
        Case "Prix REPAIR catalogue a date_PRODUIT", _
             "Dates de validite des prix REPAIR connus sur Projet (ASL)_PRODUIT", _
             "Lead time REPAIR (jours) manuel_PRODUIT"
            EstObligatoire_Produit = Not (nodeType = "LRUNR" Or nodeType = "SRUNR")
            Exit Function
    End Select

    ' 11 autres champs metier des 19 : obligatoires si MotifNA != NA.
    Select Case sColonneBdD
        Case "Fournisseur_PRODUIT", _
             "Reference fournisseur_PRODUIT", _
             COL_MOTIFNA, _
             COL_MOTIFPREV, _
             "Drawing number", _
             "MTBF intrinseque Produit (h) issu de la datasheet _PRODUIT", _
             "Reference MUL maintenance_PRODUIT", _
             "Prix SPARE catalogue a date_PRODUIT", _
             "Dates de validite des prix SPARE connus sur Projet (ASL)_PRODUIT", _
             "MoQ spare catalogue_PRODUIT", _
             "Lead time SPARE catalogue (jours) a date_PRODUIT"
            EstObligatoire_Produit = True
            Exit Function
    End Select

    ' Champs hors 19 (Famille Generique, Ref RAM Report, Ref datasheet,
    ' Ref MUL install/soft/interface) : jamais obligatoires.
    EstObligatoire_Produit = False
End Function

' =========================================================================
' TESTS
' =========================================================================

Public Sub Test_Regle_Produit_MatriceMeQ()
    ' Test : matrice MeQ v0.2 sur les 19 champs Produit.
    ' Chaque cas construit un row minimal, iterate sur les 19 champs, compte
    ' nbObl et compare a la valeur attendue.
    Debug.Print "=== Test_Regle_Produit_MatriceMeQ ==="
    Call VerifierNbObl_Produit("NA",   "*",    "LRU",   3, "MotifNA=NA -> 3")
    Call VerifierNbObl_Produit("OK",   "HARD", "LRU",  18, "HARD + non-NR -> 18")
    Call VerifierNbObl_Produit("OK",   "HARD", "LRUNR", 15, "HARD + LRUNR -> 15")
    Call VerifierNbObl_Produit("OK",   "SOFT", "SRU",  18, "SOFT + non-NR -> 18")
    Call VerifierNbObl_Produit("OK",   "SOFT", "SRUNR", 15, "SOFT + SRUNR -> 15")
    Call VerifierNbObl_Produit("OK",   "BOTH", "LRU",  19, "BOTH + non-NR -> 19")
    Call VerifierNbObl_Produit("OK",   "BOTH", "LRUNR", 16, "BOTH + LRUNR -> 16")
    Call VerifierNbObl_Produit("OK",   "NA",   "SRU",  17, "NA (prev) + non-NR -> 17")
    Call VerifierNbObl_Produit("OK",   "NA",   "SRUNR", 14, "NA (prev) + SRUNR -> 14")
End Sub

Private Sub VerifierNbObl_Produit(ByVal sMotifNA As String, ByVal sMotifPrev As String, ByVal sNodeType As String, ByVal nbAttendu As Long, ByVal sLibelle As String)
    ' Iterate sur les 19 ColonneBdD et compte combien retournent True.
    Dim sColNT As String
    sColNT = "Node type Produit (LRU/SRU" & ChrW(&H2026) & ")_PRODUIT"

    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("MotifNonApplicabilite_PRODUIT") = sMotifNA
    d("Motif_NonApplicabilitePreventif_PRODUIT") = sMotifPrev
    d(sColNT) = sNodeType

    Dim champs As Variant
    champs = Array( _
        "DTR", _
        "Designation_PRODUIT", _
        sColNT, _
        "Fournisseur_PRODUIT", _
        "Reference fournisseur_PRODUIT", _
        "MotifNonApplicabilite_PRODUIT", _
        "Motif_NonApplicabilitePreventif_PRODUIT", _
        "Drawing number", _
        "MTBF intrinseque Produit (h) issu de la datasheet _PRODUIT", _
        "Duree de vie generique Produit HARDWARE (annee)_PRODUIT", _
        "Duree de vie generique Produit (annee) SOFTWARE _PRODUIT", _
        "Reference MUL maintenance_PRODUIT", _
        "Prix SPARE catalogue a date_PRODUIT", _
        "Dates de validite des prix SPARE connus sur Projet (ASL)_PRODUIT", _
        "MoQ spare catalogue_PRODUIT", _
        "Lead time SPARE catalogue (jours) a date_PRODUIT", _
        "Prix REPAIR catalogue a date_PRODUIT", _
        "Dates de validite des prix REPAIR connus sur Projet (ASL)_PRODUIT", _
        "Lead time REPAIR (jours) manuel_PRODUIT")

    Dim nb As Long, i As Long
    nb = 0
    For i = LBound(champs) To UBound(champs)
        If SIRUS_RegleObligatoire("Produit", CStr(champs(i)), d) Then nb = nb + 1
    Next i

    Dim sStatut As String
    If nb = nbAttendu Then sStatut = "OK" Else sStatut = "KO"
    Debug.Print "  [" & sStatut & "] " & sLibelle & " : attendu=" & nbAttendu & " obtenu=" & nb
End Sub

Public Sub Test_Regle_TachesATM_DescriptionRang3()
    ' Test : Description obligatoire si Rang commence par 3 ou 4
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("Rang_TACHEATM") = "3.1"
    Debug.Print "Rang 3.1 -> Description oblig ? " & SIRUS_RegleObligatoire("TachesATM", "Description_TACHEATM", d)
    ' Attendu : True

    d("Rang_TACHEATM") = "2.1"
    Debug.Print "Rang 2.1 -> Description oblig ? " & SIRUS_RegleObligatoire("TachesATM", "Description_TACHEATM", d)
    ' Attendu : False
End Sub

Public Sub Test_Regle_TachesATM_TorqueTighten()
    ' Test : Criteres Nm obligatoire si Description contient "Torque Tighten"
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("Description_TACHEATM") = "Torque Tighten the M8 screws to 25 Nm"
    Debug.Print "Torque Tighten -> Criteres Nm oblig ? " & SIRUS_RegleObligatoire("TachesATM", "CriteresDimensionnels_TACHEATM", d)
    ' Attendu : True

    d("Description_TACHEATM") = "Remove the bracket"
    Debug.Print "Pas de Torque -> Criteres Nm oblig ? " & SIRUS_RegleObligatoire("TachesATM", "CriteresDimensionnels_TACHEATM", d)
    ' Attendu : False
End Sub

Public Sub Test_Regle_Alerte_MTBFInferieur()
    ' Test : alerte si MTBF fonctionnel < MTBF intrinseque
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("MTBF fonctionnel (h) issu du RAM report projet_PROJETASL") = "5000"
    d("MTBF intrinseque Produit (h) issu de la datasheet _PRODUIT") = "8000"
    Debug.Print "MTBF fct=5000 < intrins=8000 -> alerte : '" & SIRUS_RegleAlerte("ProjetASL", "MTBF fonctionnel (h) issu du RAM report projet_PROJETASL", d) & "'"
    ' Attendu : libelle d'alerte

    d("MTBF fonctionnel (h) issu du RAM report projet_PROJETASL") = "10000"
    Debug.Print "MTBF fct=10000 >= intrins=8000 -> alerte : '" & SIRUS_RegleAlerte("ProjetASL", "MTBF fonctionnel (h) issu du RAM report projet_PROJETASL", d) & "'"
    ' Attendu : "" (vide)
End Sub

Public Sub Test_Regle_Toutes()
    Call Test_Regle_Produit_MatriceMeQ
    Call Test_Regle_TachesATM_DescriptionRang3
    Call Test_Regle_TachesATM_TorqueTighten
    Call Test_Regle_Alerte_MTBFInferieur
End Sub
