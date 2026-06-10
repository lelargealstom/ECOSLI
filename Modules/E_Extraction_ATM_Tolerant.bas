Attribute VB_Name = "E_Extraction_ATM_Tolerant"
'--------------------------------------------------------------------------------------------------------
'
'                                       Generateur Livrables SLI
'                                       Module E_Extraction_ATM_Tolerant
'                                       Developpement : LE LARGE Thomas (LGM/Alstom)
'                                       Extraction ATM tolerante - detection par libelle
'                                       au lieu de positions fixes (cf. EXTRACTION_LOGIQUE_METIER.md).
'                                       L'ancien E_Extraction_ATM reste fonctionnel comme fallback.
'                                       Juin 2026
'
'--------------------------------------------------------------------------------------------------------
Option Compare Text
Option Explicit

' Variables module : recordsets pour verifier presence DTR en BdD
Private REQSEL_INFO As Object
Private REQSEL_TACHES As Object

' Zone d'entete par defaut pour chercher les libelles cartouche
Private Const ZONE_CARTOUCHE As String = "A1:N12"

' =========================================================================
' Extraction_Fiches_ATM_Tolerant : point d'entree publique
' =========================================================================
' Meme signature que Extraction_Fiches_ATM (ancien) pour drop-in replacement.
' Strategie tolerante : detection des champs cartouche par libelle + detection
' des colonnes du tableau de taches par header.
Public Sub Extraction_Fiches_ATM_Tolerant(ByRef wb As Workbook, _
                                            ByVal langue As String, _
                                            ByVal nbFiches As Integer)

    Dim ws As Worksheet
    Dim indexFiche As Integer

    On Error GoTo ErrorHandler

    ' Recordsets BdD (memes que ancien algo)
    Call InitialiserRecordsetsATM_T

    ' Creer les feuilles Info_ATM et Taches_ATM dans wb (memes en-tetes
    ' que l'ancien algo - format BdD attendu par SIRUS)
    Call CreerFeuillesExtraction_T(wb)

    ' Parcourir chaque fiche DTR
    indexFiche = 0
    For Each ws In wb.Worksheets
        If ws.Name Like "*DTR*" Then
            indexFiche = indexFiche + 1
            Call ExtraireInfoATM_T(ws, langue)
            Call ExtraireTachesATM_T(ws, langue)
            On Error Resume Next
            Call barprogress(indexFiche, nbFiches)
            On Error GoTo ErrorHandler
        End If
    Next ws

    ' Mise en forme finale (wrap text off)
    Call FormaterFeuilles_T(wb)

    On Error Resume Next
    REQSEL_INFO.Close
    REQSEL_TACHES.Close
    Exit Sub

ErrorHandler:
    MsgBox "Erreur extraction ATM tolerante : " & Err.Description & vbCrLf & _
           "Feuille en cours : " & IIf(ws Is Nothing, "n/a", ws.Name), vbCritical
    On Error Resume Next
    REQSEL_INFO.Close
    REQSEL_TACHES.Close
End Sub

' =========================================================================
' InitialiserRecordsetsATM_T (copie de l'algo existant)
' =========================================================================
Private Sub InitialiserRecordsetsATM_T()
    Dim sqlInfo As String, sqlTaches As String

    Set REQSEL_INFO = CreateObject("ADODB.Recordset")
    Set REQSEL_TACHES = CreateObject("ADODB.Recordset")

    sqlInfo = "SELECT [DTR_INFOATM] FROM [Info_ATM$]"
    sqlTaches = "SELECT [DTR_TACHEATM] FROM [Taches_ATM$]"

    Call open_REQSEL1(sqlInfo, "REQSEL_INFO")
    Call open_REQSEL1(sqlTaches, "REQSEL_TACHES")
End Sub

' =========================================================================
' CreerFeuillesExtraction_T : memes en-tetes que l'ancien
' =========================================================================
Private Sub CreerFeuillesExtraction_T(ByRef wb As Workbook)
    Dim wsInfo As Worksheet, wsTaches As Worksheet

    ' Supprimer si deja presentes (cas relance)
    On Error Resume Next
    Application.DisplayAlerts = False
    wb.Sheets("Info_ATM").Delete
    wb.Sheets("Taches_ATM").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    ' Feuille Info_ATM
    wb.Sheets.Add Before:=wb.Worksheets(1)
    Set wsInfo = wb.ActiveSheet
    wsInfo.Name = "Info_ATM"
    With wsInfo
        .Range("A1").Value = "DTR_INFOATM"
        .Range("B1").Value = "DescriptifATMFR_INFOATM"
        .Range("C1").Value = "DesignationATM_INFOATM"
        .Range("D1").Value = "LocalisationATM_INFOATM"
        .Range("E1").Value = "PLANATM_INFOATM"
        .Range("F1").Value = "CommentaireATM_INFOATM"
        .Range("G1").Value = "Version_INFOATM"
        .Range("H1").Value = "DescriptifATMEN_INFOATM"
        .Range("I1").Value = "LocalisationATMEN_INFOATM"
        .Range("J1").Value = "CommentairesATMEN_INFOATM"
        .Range("K1").Value = "Reparabilite_INFOATM"
        .Range("L1").Value = "HARD/SOFT"
        .Range("M1").Value = "Poids_INFOATM"
        .Range("N1").Value = "Temps g" & ChrW(233) & "n" & ChrW(233) & "rique de la t" & ChrW(226) & "che HARDWARE_INFOATM"
        .Range("O1").Value = "Temps g" & ChrW(233) & "n" & ChrW(233) & "rique de la t" & ChrW(226) & "che SOFTWARE_INFOATM"
        .Range("P1").Value = "Temps analytique cumul" & ChrW(233) & " de la tache HARDWARE_INFOATM"
        .Range("Q1").Value = "Temps analytique cumul" & ChrW(233) & " de la tache SOFTWARE_INFOATM"
        .Range("R1").Value = "Connu en BDD ?"
    End With

    ' Feuille Taches_ATM
    wb.Sheets.Add Before:=wb.Worksheets(1)
    Set wsTaches = wb.ActiveSheet
    wsTaches.Name = "Taches_ATM"
    ' Forcer format Texte sur la colonne Rang (C) AVANT toute ecriture
    ' pour empecher la conversion Excel de "1-1" / "5-2" en date.
    ' Idem pour les colonnes de description / criteres qui peuvent contenir
    ' du texte ambigus.
    wsTaches.Columns("C").NumberFormat = "@"
    wsTaches.Columns("G").NumberFormat = "@"
    wsTaches.Columns("H").NumberFormat = "@"
    With wsTaches
        .Range("A1").Value = "DTR_TACHEATM"
        .Range("B1").Value = "Langue_TACHEATM"
        .Range("C1").Value = "Rang_TACHEATM"
        .Range("D1").Value = "HARD/SOFT"
        .Range("E1").Value = "Categorie_TACHEATM"
        .Range("F1").Value = "Niveau qualif_TACHEATM"
        .Range("G1").Value = "Description_TACHEATM"
        .Range("H1").Value = "Crit" & ChrW(232) & "resDimensionnels_TACHEATM"
        .Range("I1").Value = "Outillage" & ChrW(233) & "cifique_TACHEATM"
        .Range("J1").Value = "OutillageStandard_TACHEATM"
        .Range("K1").Value = "Pi" & ChrW(232) & "cedeRechangeConsommable_TACHEATM"
        .Range("L1").Value = "NbOp" & ChrW(233) & "rateurs_TACHEATM"
        .Range("M1").Value = "TempsEstim" & ChrW(233) & "Heures_TACHEATM"
        .Range("N1").Value = "Temps analytique cumul" & ChrW(233) & " de la tache HARD_TACHEATM"
        .Range("O1").Value = "Temps analytique cumul" & ChrW(233) & " de la tache SOFT_TACHEATM"
        .Range("P1").Value = "Pictogramme pour ATM2.0_TACHEATM"
        .Range("Q1").Value = "Texte / commentaire picto2_TACHEATM"
        .Range("R1").Value = "Connu en BDD ?"
    End With
End Sub

' =========================================================================
' ExtraireInfoATM_T : extraction tolerante du cartouche
' =========================================================================
' Detecte les positions D4/D5/D10/I6/I7/I8/I9 par libelle :
'   col C : "Equipement", "Descriptif", "Commentaire"   -> valeur col D
'   col H : "Poids", "Niveau", "plan", "Reparable"       -> valeur col I
' Fallback : positions fixes si libelle introuvable (compat ancien VBN).
Private Sub ExtraireInfoATM_T(ByRef ws As Worksheet, ByVal langue As String)
    Dim wsInfo As Worksheet
    Dim derniereLigne As Long
    Dim dtr As String, hardSoft As String
    Dim tempsTotal As Double
    Dim dtrPresent As Boolean

    Dim valDesignation As String, valDescriptifFR As String, valCommentaire As String
    Dim valPlan As String, valPoids As String, valReparabilite As String

    Set wsInfo = ws.Parent.Sheets("Info_ATM")
    derniereLigne = wsInfo.Range("A100000").End(xlUp).Row

    ' Identifiant DTR depuis nom de feuille (logique metier inchangee)
    dtr = Left(ws.Name, 13)
    hardSoft = DeterminerHardSoft_T(ws.Name)
    tempsTotal = RecupererTempsTotal_T(ws, langue)
    dtrPresent = RechercheValeurdansRS(REQSEL_INFO, dtr, "DTR_INFOATM")

    ' Detection des champs cartouche par libelle (zone A1:N12)
    valDesignation = LireValeurApresLibelle(ws, ZONE_CARTOUCHE, _
                        Array("Equipement", "Designation", "D" & ChrW(233) & "signation"))
    valDescriptifFR = LireValeurApresLibelle(ws, ZONE_CARTOUCHE, _
                        Array("Descriptif", "Description"))
    valCommentaire = LireValeurApresLibelle(ws, ZONE_CARTOUCHE, _
                        Array("Commentaire"))
    valPlan = LireValeurApresLibelle(ws, ZONE_CARTOUCHE, _
                        Array("Num" & ChrW(233) & "ro de plan", "Plan", "Drawing"))
    valPoids = LireValeurApresLibelle(ws, ZONE_CARTOUCHE, _
                        Array("Poids", "Weight"))
    valReparabilite = LireValeurApresLibelle(ws, ZONE_CARTOUCHE, _
                        Array("R" & ChrW(233) & "parable", "Reparable", "Repairability"))

    ' Fallback positions historiques si rien trouve (compat ancien VBN)
    If LenB(valDesignation) = 0 Then valDesignation = CStr(ws.Range("D4").Value)
    If LenB(valDescriptifFR) = 0 Then valDescriptifFR = CStr(ws.Range("D5").Value)
    If LenB(valCommentaire) = 0 Then valCommentaire = CStr(ws.Range("D10").Value)
    If LenB(valPlan) = 0 Then valPlan = CStr(ws.Range("I8").Value)
    If LenB(valPoids) = 0 Then valPoids = CStr(ws.Range("I6").Value)
    If LenB(valReparabilite) = 0 Then valReparabilite = CStr(ws.Range("I9").Value)

    ' Ecrire ligne Info_ATM (memes colonnes que ancien algo)
    With wsInfo
        .Cells(derniereLigne + 1, 1).Value = dtr                          ' DTR
        .Cells(derniereLigne + 1, 2).Value = valDescriptifFR              ' Descriptif FR
        .Cells(derniereLigne + 1, 3).Value = valDesignation               ' Designation
        .Cells(derniereLigne + 1, 4).Value = ""                           ' Localisation
        .Cells(derniereLigne + 1, 5).Value = valPlan                      ' Plan
        .Cells(derniereLigne + 1, 6).Value = valCommentaire               ' Commentaire FR
        .Cells(derniereLigne + 1, 7).Value = "A1"                         ' Version (dur)

        ' Descriptif EN : prefixe "REPLACE " si FR non vide (logique metier)
        If LenB(valDescriptifFR) > 0 Then
            .Cells(derniereLigne + 1, 8).Value = "REPLACE " & valDescriptifFR
        End If

        .Cells(derniereLigne + 1, 9).Value = ""                           ' Localisation EN
        .Cells(derniereLigne + 1, 10).Value = valCommentaire              ' Commentaires EN
        .Cells(derniereLigne + 1, 11).Value = valReparabilite             ' Reparabilite
        .Cells(derniereLigne + 1, 12).Value = hardSoft                    ' HARD/SOFT
        .Cells(derniereLigne + 1, 13).Value = valPoids                    ' Poids

        ' Temps generique selon HARD/SOFT
        Select Case hardSoft
            Case "HARD"
                .Cells(derniereLigne + 1, 14).Value = tempsTotal
            Case "SOFT"
                .Cells(derniereLigne + 1, 15).Value = tempsTotal
            Case Else
                .Cells(derniereLigne + 1, 14).Value = tempsTotal
        End Select

        .Cells(derniereLigne + 1, 18).Value = IIf(dtrPresent, "Connu", "Inconnu")
    End With
End Sub

' =========================================================================
' ExtraireTachesATM_T : extraction tolerante des taches
' =========================================================================
' 1. Trouver la ligne d'en-tete du tableau (libelle "Nb" en col B + autres
'    libelles colonnes a cote). Si introuvable -> ligne 12 par defaut.
' 2. Mapper chaque colonne BdD (Description/Outillage/Temps...) a son col index
'    par libelle. Fallback positions fixes (B=Rang,C=Description,E=Crit...).
' 3. Parcourir de ligneHeader+1 jusqu'a "Temps total :" ou ligne avec col B vide
'    sur 3 lignes consecutives (fin du tableau).
' 4. Distinguer ligne categorie (col B = 1 char) vs ligne tache (col B = "x-y"),
'    meme logique que l'ancien algo.
Private Sub ExtraireTachesATM_T(ByRef ws As Worksheet, ByVal langue As String)
    Dim wsTaches As Worksheet
    Dim dtr As String, hardSoft As String, nvMaint As String
    Dim tempsTotal As Double
    Dim dtrPresent As Boolean
    Dim categorieTache As String
    Dim ligneHeader As Long, ligneDebut As Long, ligneFin As Long
    Dim i As Long
    Dim mapCols As Object

    Set wsTaches = ws.Parent.Sheets("Taches_ATM")

    dtr = Left(ws.Name, 13)
    hardSoft = DeterminerHardSoft_T(ws.Name)
    nvMaint = LireValeurApresLibelle(ws, ZONE_CARTOUCHE, _
                Array("Niveau maintenance", "Niv maintenance", "Maintenance level"))
    If LenB(nvMaint) = 0 Then nvMaint = CStr(ws.Range("I7").Value)
    tempsTotal = RecupererTempsTotal_T(ws, langue)
    dtrPresent = RechercheValeurdansRS(REQSEL_TACHES, dtr, "DTR_TACHEATM")

    ' 1. Detection de la ligne d'en-tete du tableau
    ligneHeader = TrouverLigneHeaderTaches(ws)
    If ligneHeader = 0 Then ligneHeader = 12  ' fallback ancien template

    ligneDebut = ligneHeader + 1
    ligneFin = ws.Range("B10000").End(xlUp).Row

    ' 2. Mapping colonnes du tableau par libelle d'en-tete
    Set mapCols = MapperColonnesTaches(ws, ligneHeader)

    ' 3. Verification feuille vide tolerante (test sur les 30 premieres lignes
    '    apres header : si aucune n'a un format "x-y" en col B, fiche vide)
    If EstFeuilleVide_T(ws, ligneDebut) Then Exit Sub

    ' 4. Parcourir taches
    For i = ligneDebut To ligneFin
        Dim valB As String
        valB = CStr(ws.Cells(i, 2).Value)

        If Len(valB) = 0 Then
            ' Ligne vide -> passer
        ElseIf Len(valB) = 1 Then
            ' Categorie de chapitre (A/B/C/D ou 1/2/3/...)
            categorieTache = CStr(ws.Cells(i, 3).Value)
        Else
            ' Tache : extraire via mapping
            Dim derniereLigneTaches As Long
            derniereLigneTaches = wsTaches.Range("A100000").End(xlUp).Row
            Call AjouterTacheATM_T(wsTaches, derniereLigneTaches + 1, ws, i, _
                                    dtr, langue, hardSoft, nvMaint, tempsTotal, _
                                    categorieTache, dtrPresent, mapCols)
        End If
    Next i
End Sub

' =========================================================================
' AjouterTacheATM_T : ecrit une ligne de tache dans Taches_ATM
' =========================================================================
Private Sub AjouterTacheATM_T(ByRef wsTaches As Worksheet, _
                                ByVal ligne As Long, _
                                ByRef wsDTR As Worksheet, _
                                ByVal ligneDTR As Long, _
                                ByVal dtr As String, _
                                ByVal langue As String, _
                                ByVal hardSoft As String, _
                                ByVal nvMaint As String, _
                                ByVal tempsTotal As Double, _
                                ByVal categorie As String, _
                                ByVal dtrPresent As Boolean, _
                                ByRef mapCols As Object)
    With wsTaches
        Dim rangBrut As Variant
        rangBrut = wsDTR.Cells(ligneDTR, ColParDefaut(mapCols, "Rang", 2)).Value
        ' Lecture par .Text + CStr pour preserver la chaine telle qu'affichee
        ' (eviter conversion auto Excel "1-1" -> date)
        Dim rangTexte As String
        rangTexte = CStr(wsDTR.Cells(ligneDTR, ColParDefaut(mapCols, "Rang", 2)).Text)
        If LenB(rangTexte) = 0 Then rangTexte = CStr(rangBrut)

        .Cells(ligne, 1).Value = dtr
        .Cells(ligne, 2).Value = langue
        ' Colonne C est en NumberFormat = "@" (force a la creation de la
        ' feuille Taches_ATM) : valeur ecrite comme texte sans risque.
        .Cells(ligne, 3).Value = rangTexte
        .Cells(ligne, 4).Value = hardSoft
        .Cells(ligne, 5).Value = categorie
        .Cells(ligne, 6).Value = nvMaint
        .Cells(ligne, 7).Value = wsDTR.Cells(ligneDTR, ColParDefaut(mapCols, "Description", 3)).Value
        .Cells(ligne, 8).Value = wsDTR.Cells(ligneDTR, ColParDefaut(mapCols, "Criteres", 5)).Value
        .Cells(ligne, 9).Value = wsDTR.Cells(ligneDTR, ColParDefaut(mapCols, "OutillageSpec", 6)).Value
        .Cells(ligne, 10).Value = wsDTR.Cells(ligneDTR, ColParDefaut(mapCols, "OutillageStd", 7)).Value
        .Cells(ligne, 11).Value = wsDTR.Cells(ligneDTR, ColParDefaut(mapCols, "PieceRechange", 8)).Value
        .Cells(ligne, 12).Value = wsDTR.Cells(ligneDTR, ColParDefaut(mapCols, "NbOperateurs", 9)).Value
        .Cells(ligne, 13).Value = wsDTR.Cells(ligneDTR, ColParDefaut(mapCols, "Temps", 10)).Value

        ' Temps analytique cumule selon HARD/SOFT
        Select Case hardSoft
            Case "HARD"
                .Cells(ligne, 14).Value = tempsTotal
            Case "SOFT"
                .Cells(ligne, 15).Value = tempsTotal
            Case Else
                .Cells(ligne, 14).Value = tempsTotal
        End Select

        .Cells(ligne, 16).Value = ""
        .Cells(ligne, 17).Value = ""
        .Cells(ligne, 18).Value = IIf(dtrPresent, "Connu", "Inconnu")
    End With
End Sub

' =========================================================================
' Helpers metier (logique inchangee de l'ancien algo)
' =========================================================================

Private Function DeterminerHardSoft_T(ByVal nomFeuille As String) As String
    If Right(nomFeuille, 4) = "HARD" Then
        DeterminerHardSoft_T = "HARD"
    ElseIf Right(nomFeuille, 4) = "SOFT" Then
        DeterminerHardSoft_T = "SOFT"
    Else
        DeterminerHardSoft_T = ""
    End If
End Function

Private Function RecupererTempsTotal_T(ByRef ws As Worksheet, ByVal langue As String) As Double
    Dim cellTemps As Range
    Dim texteFR As String, texteEN As String

    texteFR = "Temps total"
    texteEN = "Total time"

    On Error Resume Next
    Set cellTemps = ws.Cells.Find(What:=texteFR, LookIn:=xlValues, lookat:=xlPart)
    If cellTemps Is Nothing Then
        Set cellTemps = ws.Cells.Find(What:=texteEN, LookIn:=xlValues, lookat:=xlPart)
    End If
    On Error GoTo 0

    If Not cellTemps Is Nothing Then
        ' Le temps est en col J (10) sur la meme ligne (logique metier conservee)
        RecupererTempsTotal_T = SafeNum(ws.Cells(cellTemps.Row, 10).Value, 0)
    Else
        RecupererTempsTotal_T = 0
    End If
End Function

' =========================================================================
' EstFeuilleVide_T : tolerant - cherche un format de rang "x-y" en col B
' dans les 30 lignes apres le header. Aucun -> fiche vide.
' =========================================================================
Private Function EstFeuilleVide_T(ByRef ws As Worksheet, ByVal ligneDebut As Long) As Boolean
    Dim i As Long
    Dim valB As String

    For i = ligneDebut To ligneDebut + 30
        valB = CStr(ws.Cells(i, 2).Value)
        ' Format tache : au moins 2 caracteres avec un separateur (- ou .) au milieu
        If Len(valB) >= 3 And (InStr(valB, "-") > 0 Or InStr(valB, ".") > 0) Then
            EstFeuilleVide_T = False
            Exit Function
        End If
    Next i
    EstFeuilleVide_T = True
End Function

' =========================================================================
' TrouverLigneHeaderTaches : cherche la ligne contenant "Description" et
' "Outillage" en col proche (caracteristique de l'en-tete tableau).
' Renvoie 0 si introuvable.
' =========================================================================
Private Function TrouverLigneHeaderTaches(ByRef ws As Worksheet) As Long
    Dim i As Long, j As Long
    Dim aDescription As Boolean, aOutillage As Boolean

    ' Cherche entre les lignes 10 et 25 (zone typique du header)
    For i = 10 To 25
        aDescription = False
        aOutillage = False
        For j = 1 To 12
            Dim val As String
            val = CStr(ws.Cells(i, j).Value)
            If InStr(1, val, "Description", vbTextCompare) > 0 Then aDescription = True
            If InStr(1, val, "Outillage", vbTextCompare) > 0 Then aOutillage = True
        Next j
        If aDescription And aOutillage Then
            TrouverLigneHeaderTaches = i
            Exit Function
        End If
    Next i

    TrouverLigneHeaderTaches = 0
End Function

' =========================================================================
' MapperColonnesTaches : construit Dictionary {role -> col index} en parsant
' la ligne d'en-tete du tableau.
' Roles : Rang, Description, Criteres, OutillageSpec, OutillageStd,
'         PieceRechange, NbOperateurs, Temps.
' =========================================================================
Private Function MapperColonnesTaches(ByRef ws As Worksheet, ByVal ligneHeader As Long) As Object
    Dim d As Object
    Dim j As Long
    Dim val As String

    Set d = CreateObject("Scripting.Dictionary")
    If ligneHeader = 0 Then
        Set MapperColonnesTaches = d
        Exit Function
    End If

    For j = 1 To 15
        val = CStr(ws.Cells(ligneHeader, j).Value)
        If LenB(val) > 0 Then
            ' Mapping roles selon mots-cles dans le libelle d'en-tete
            If InStr(1, val, "Nb", vbTextCompare) > 0 _
               And InStr(1, val, "op" & ChrW(233) & "rateur", vbTextCompare) > 0 Then
                d("NbOperateurs") = j
            ElseIf val = "Nb" Or LCase$(val) = "rang" Or LCase$(val) = "n" Then
                d("Rang") = j
            ElseIf InStr(1, val, "Description", vbTextCompare) > 0 Then
                d("Description") = j
            ElseIf InStr(1, val, "Crit" & ChrW(232) & "re", vbTextCompare) > 0 _
                   Or InStr(1, val, "Critere", vbTextCompare) > 0 Then
                d("Criteres") = j
            ElseIf InStr(1, val, "Outillage", vbTextCompare) > 0 _
                   And InStr(1, val, "sp" & ChrW(233) & "cifique", vbTextCompare) > 0 Then
                d("OutillageSpec") = j
            ElseIf InStr(1, val, "Outillage", vbTextCompare) > 0 _
                   And InStr(1, val, "standard", vbTextCompare) > 0 Then
                d("OutillageStd") = j
            ElseIf InStr(1, val, "Pi" & ChrW(232) & "ce", vbTextCompare) > 0 _
                   Or InStr(1, val, "Piece", vbTextCompare) > 0 _
                   Or InStr(1, val, "rechange", vbTextCompare) > 0 Then
                d("PieceRechange") = j
            ElseIf InStr(1, val, "Temps", vbTextCompare) > 0 _
                   Or InStr(1, val, "Time", vbTextCompare) > 0 Then
                d("Temps") = j
            End If
        End If
    Next j

    ' Le Rang est par convention en col B (2) si non trouve
    If Not d.Exists("Rang") Then d("Rang") = 2

    Set MapperColonnesTaches = d
End Function

Private Function ColParDefaut(ByRef mapCols As Object, _
                                ByVal cle As String, _
                                ByVal defaut As Long) As Long
    If mapCols Is Nothing Then
        ColParDefaut = defaut
    ElseIf mapCols.Exists(cle) Then
        ColParDefaut = CLng(mapCols(cle))
    Else
        ColParDefaut = defaut
    End If
End Function

' =========================================================================
' LireValeurApresLibelle : cherche un libelle dans une zone, renvoie la
' valeur de la cellule a droite (offset +1 col) ou en dessous si droite vide.
' Pattern classique des cartouches "Label : / Valeur" tableau Excel.
' =========================================================================
Private Function LireValeurApresLibelle(ByRef ws As Worksheet, _
                                          ByVal zone As String, _
                                          ByRef libelles As Variant) As String
    Dim cell As Range
    Dim cellMatch As Range
    Dim mc As Variant
    Dim plage As Range

    Set plage = ws.Range(zone)

    For Each mc In libelles
        Set cellMatch = Nothing
        On Error Resume Next
        Set cellMatch = plage.Find(What:=CStr(mc), LookIn:=xlValues, lookat:=xlPart, MatchCase:=False)
        On Error GoTo 0
        If Not cellMatch Is Nothing Then
            ' Tenter cellule a droite d'abord
            Dim valDroite As String, valDessous As String
            valDroite = CStr(ws.Cells(cellMatch.Row, cellMatch.Column + 1).Value)
            If LenB(Trim$(valDroite)) > 0 Then
                LireValeurApresLibelle = valDroite
                Exit Function
            End If
            ' Sinon cellule dessous
            valDessous = CStr(ws.Cells(cellMatch.Row + 1, cellMatch.Column).Value)
            If LenB(Trim$(valDessous)) > 0 Then
                LireValeurApresLibelle = valDessous
                Exit Function
            End If
        End If
    Next mc

    LireValeurApresLibelle = ""
End Function

' =========================================================================
' FormaterFeuilles_T : memes finitions que l'ancien
' =========================================================================
Private Sub FormaterFeuilles_T(ByRef wb As Workbook)
    Dim wsInfo As Worksheet, wsTaches As Worksheet
    On Error Resume Next
    Set wsInfo = wb.Sheets("Info_ATM")
    Set wsTaches = wb.Sheets("Taches_ATM")
    On Error GoTo 0

    If Not wsInfo Is Nothing Then wsInfo.Cells.WrapText = False
    If Not wsTaches Is Nothing Then wsTaches.Cells.WrapText = False

    If Not wsInfo Is Nothing Then
        wsInfo.Activate
        wsInfo.Range("A1").Select
    End If
End Sub
