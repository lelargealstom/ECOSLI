Attribute VB_Name = "G_SIRUS_Regles"
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

Public Function SIRUS_RegleObligatoire(ByVal sTable As String, _
                                       ByVal sColonneBdD As String, _
                                       ByRef ligne As Object) As Boolean
    ' Renvoie True si le champ est obligatoire pour cette ligne.
    ' ligne : Dictionary {ColonneBdD -> valeur} de la ligne en cours.

    Select Case sTable & "|" & sColonneBdD

        ' ----------------- Produit -----------------
        ' Note 2026-06-10 : Prix Spare/Repair manuel + dates source ne sont
        ' PLUS obligatoires conditionnellement (saisie libre).
        ' Cf. PeuplerProduit -> Obligatoire = "Non".

        Case "Produit|MTBF intrins" & ChrW(232) & "que Produit (h) issu de la datasheet _PRODUIT"
            ' MTBF intrinseque obligatoire si NodeType LRU/SRU ET pas de
            ' MotifNonApplicabilite renseigne (cas typique : LRU sans
            ' donnees RAM disponibles a date).
            If EstNodeTypeIn(ligne, "Node type Produit (LRU/SRU" & ChrW(8230) & ")_PRODUIT", _
                             Array("LRU", "SRU")) Then
                SIRUS_RegleObligatoire = (LenB(GetVal(ligne, "MotifNonApplicabilit" & ChrW(233) & "_PRODUIT")) = 0)
            End If

        Case "Produit|MotifNonApplicabilit" & ChrW(233) & "_PRODUIT"
            ' MotifNonApplicabilite obligatoire si MTBF intrinseque vide
            ' (justifie pourquoi on ne renseigne pas la valeur RAM).
            SIRUS_RegleObligatoire = _
                (LenB(GetVal(ligne, "MTBF intrins" & ChrW(232) & "que Produit (h) issu de la datasheet _PRODUIT")) = 0)

        ' ----------------- TachesATM -----------------
        Case "TachesATM|Description_TACHEATM"
            ' Description obligatoire pour les taches de rang 3.* (Depose)
            ' et 4.* (Pose) qui sont les operations elementaires detaillees.
            Dim rang As String
            rang = CStr(GetVal(ligne, "Rang_TACHEATM"))
            SIRUS_RegleObligatoire = (Left$(rang, 2) = "3." Or Left$(rang, 2) = "4.")

        Case "TachesATM|Crit" & ChrW(232) & "resDimensionnels_TACHEATM"
            ' Criteres dimensionnels (Critere Nm) obligatoires si la
            ' description mentionne "Torque Tighten" (couple de serrage).
            Dim descTorque As String
            descTorque = CStr(GetVal(ligne, "Description_TACHEATM"))
            SIRUS_RegleObligatoire = (InStr(1, descTorque, "Torque Tighten", vbTextCompare) > 0)

        Case "TachesATM|OutillageStandard_TACHEATM", "TachesATM|Outillage" & ChrW(233) & "cifique_TACHEATM"
            ' Outillage standard/specifique obligatoire pour les taches
            ' de rang 2.1 (Outils prerequis Depose) et 5.1 (Outils
            ' prerequis Pose) sur les composants >= SRU (LRU exclus :
            ' demontage complet du LRU, pas d'outillage detaille).
            Dim rangO As String, ntO As String
            rangO = CStr(GetVal(ligne, "Rang_TACHEATM"))
            ntO = UCase$(Trim$(CStr(GetVal(ligne, "Node type Produit (LRU/SRU" & ChrW(8230) & ")_PRODUIT"))))
            SIRUS_RegleObligatoire = ((rangO = "2.1" Or rangO = "5.1") And _
                                     (ntO = "SRU" Or ntO = "SRUNR" Or ntO = "ASSEMBLY"))

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

        Case Else
            ' Defaut : pas de regle conditionnelle definie -> False.
            ' (Les champs Obligatoire=Oui du config sont traites en amont
            '  par SIRUS_AnalyserFeuille, pas par cette fonction.)
            SIRUS_RegleObligatoire = False
    End Select
End Function

Public Function SIRUS_RegleAlerte(ByVal sTable As String, _
                                  ByVal sColonneBdD As String, _
                                  ByRef ligne As Object) As String
    ' Renvoie "" si pas d'alerte, sinon le libelle a afficher dans Statut.

    Select Case sTable & "|" & sColonneBdD

        ' ----------------- Produit -----------------
        ' Note 2026-06-10 : l'ancienne alerte "prix sans date de validite"
        ' a ete retiree apres refonte (champs Date Source = "Non" obligatoire,
        ' libre saisie). A reintroduire au besoin sur les nouveaux noms BdD
        ' une fois "??_DateSourceSpare/Repair_a_definir" remplaces.

        ' ----------------- TachesATM -----------------
        Case "TachesATM|NbOp" & ChrW(233) & "rateurs_TACHEATM"
            ' Alerte si > 4 operateurs : valeur metier rare, a confirmer.
            Dim nbOp As String
            nbOp = CStr(GetVal(ligne, "NbOp" & ChrW(233) & "rateurs_TACHEATM"))
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
            mtbfI = CStr(GetVal(ligne, "MTBF intrins" & ChrW(232) & "que Produit (h) issu de la datasheet _PRODUIT"))
            If IsNumeric(mtbfF) And IsNumeric(mtbfI) Then
                If CDbl(mtbfF) < CDbl(mtbfI) Then
                    SIRUS_RegleAlerte = "MTBF fonctionnel < MTBF intrinseque : verifier RAM report"
                End If
            End If

        Case Else
            SIRUS_RegleAlerte = ""
    End Select
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

Private Function EstNodeTypeIn(ByRef ligne As Object, _
                                ByVal sColonneBdD As String, _
                                ByRef tValeurs As Variant) As Boolean
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
' TESTS
' =========================================================================

Public Sub Test_Regle_Produit_MTBFConditionnel()
    ' Test : MTBF intrinseque obligatoire si NodeType LRU/SRU ET MotifNA vide
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("Node type Produit (LRU/SRU" & ChrW(8230) & ")_PRODUIT") = "LRU"
    d("MotifNonApplicabilit" & ChrW(233) & "_PRODUIT") = ""
    Debug.Print "LRU + MotifNA vide -> MTBF oblig ? " & _
        SIRUS_RegleObligatoire("Produit", "MTBF intrins" & ChrW(232) & "que Produit (h) issu de la datasheet _PRODUIT", d)
    ' Attendu : True

    d("MotifNonApplicabilit" & ChrW(233) & "_PRODUIT") = "NA MECA"
    Debug.Print "LRU + MotifNA renseigne -> MTBF oblig ? " & _
        SIRUS_RegleObligatoire("Produit", "MTBF intrins" & ChrW(232) & "que Produit (h) issu de la datasheet _PRODUIT", d)
    ' Attendu : False
End Sub

Public Sub Test_Regle_TachesATM_DescriptionRang3()
    ' Test : Description obligatoire si Rang commence par 3 ou 4
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("Rang_TACHEATM") = "3.1"
    Debug.Print "Rang 3.1 -> Description oblig ? " & _
        SIRUS_RegleObligatoire("TachesATM", "Description_TACHEATM", d)
    ' Attendu : True

    d("Rang_TACHEATM") = "2.1"
    Debug.Print "Rang 2.1 -> Description oblig ? " & _
        SIRUS_RegleObligatoire("TachesATM", "Description_TACHEATM", d)
    ' Attendu : False
End Sub

Public Sub Test_Regle_TachesATM_TorqueTighten()
    ' Test : Criteres Nm obligatoire si Description contient "Torque Tighten"
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("Description_TACHEATM") = "Torque Tighten the M8 screws to 25 Nm"
    Debug.Print "Torque Tighten -> Criteres Nm oblig ? " & _
        SIRUS_RegleObligatoire("TachesATM", "Crit" & ChrW(232) & "resDimensionnels_TACHEATM", d)
    ' Attendu : True

    d("Description_TACHEATM") = "Remove the bracket"
    Debug.Print "Pas de Torque -> Criteres Nm oblig ? " & _
        SIRUS_RegleObligatoire("TachesATM", "Crit" & ChrW(232) & "resDimensionnels_TACHEATM", d)
    ' Attendu : False
End Sub

Public Sub Test_Regle_Alerte_MTBFInferieur()
    ' Test : alerte si MTBF fonctionnel < MTBF intrinseque
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("MTBF fonctionnel (h) issu du RAM report projet_PROJETASL") = "5000"
    d("MTBF intrins" & ChrW(232) & "que Produit (h) issu de la datasheet _PRODUIT") = "8000"
    Debug.Print "MTBF fct=5000 < intrins=8000 -> alerte : '" & _
        SIRUS_RegleAlerte("ProjetASL", "MTBF fonctionnel (h) issu du RAM report projet_PROJETASL", d) & "'"
    ' Attendu : libelle d'alerte

    d("MTBF fonctionnel (h) issu du RAM report projet_PROJETASL") = "10000"
    Debug.Print "MTBF fct=10000 >= intrins=8000 -> alerte : '" & _
        SIRUS_RegleAlerte("ProjetASL", "MTBF fonctionnel (h) issu du RAM report projet_PROJETASL", d) & "'"
    ' Attendu : "" (vide)
End Sub

Public Sub Test_Regle_Toutes()
    Call Test_Regle_Produit_MTBFConditionnel
    Call Test_Regle_TachesATM_DescriptionRang3
    Call Test_Regle_TachesATM_TorqueTighten
    Call Test_Regle_Alerte_MTBFInferieur
End Sub
