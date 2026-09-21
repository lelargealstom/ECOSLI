'--------------------------------------------------------------------------------------------------------
'
'                                       Generateur Livrables SLI
'                                       Module M_Extraction_ATM (Version Refactorisee)
'                                       Developpement : LE LARGE Thomas (LGM/Alstom)
'                                       Module d'extraction fiches ATM vers format BDD
'
'-----------------------------------------------------------------------------------------------------------
Option Compare Text
Option Explicit

' Variables module
Private REQSEL_INFO As Object
Private REQSEL_TACHES As Object

' =========================================================================
' PROCEDURE PRINCIPALE: EXTRACTION FICHES ATM
' =========================================================================
Public Sub Extraction_Fiches_ATM(ByRef wb As Workbook, ByVal langue As String, ByVal nbFiches As Integer)
' Extrait toutes les donnees ATM d'un fichier et les structure au format BDD
    
    Dim ws As Worksheet
    Dim indexFiche As Integer

    ' ======================================================================
    ' ETAPE 1 : Initialisation Recordsets BDD
    ' ======================================================================
    Call InitialiserRecordsetsATM
    
    ' ======================================================================
    ' ETAPE 2 : Creation feuilles extraction
    ' ======================================================================
    Call CreerFeuillesExtractionATM(wb)
    
    ' ======================================================================
    ' ETAPE 3 : Extraction donnees de chaque fiche DTR
    ' ======================================================================
    indexFiche = 0
    
    For Each ws In wb.Worksheets
        If ws.Name Like "*DTR*" Then
            indexFiche = indexFiche + 1
            
            ' Extraire Info ATM (cartouche)
            Call ExtraireInfoATM(ws, langue)
            
            ' Extraire Taches ATM
            Call ExtraireTachesATM(ws, langue)
            
            ' Mettre a jour barre progression
            Call barprogress(indexFiche, nbFiches)
        End If
    Next ws
    
    ' ======================================================================
    ' ETAPE 4 : Mise en forme finale
    ' ======================================================================
    Call FormaterFeuillesExtractionATM(wb)
    
    ' Fermer recordsets
    On Error Resume Next
    REQSEL_INFO.Close
    REQSEL_TACHES.Close
    On Error GoTo 0

End Sub

' =========================================================================
' PROCEDURE: INITIALISER RECORDSETS ATM
' =========================================================================
Private Sub InitialiserRecordsetsATM()
' Cree les recordsets pour verifier existence DTR en BDD
    
    Dim sqlInfo As String
    Dim sqlTaches As String
    
    Set REQSEL_INFO = CreateObject("ADODB.Recordset")
    Set REQSEL_TACHES = CreateObject("ADODB.Recordset")
    
    ' Requete Info ATM
    sqlInfo = "SELECT [DTR_INFOATM], [DescriptifATMFR_INFOATM], [DesignationATM_INFOATM], [LocalisationATM_INFOATM], [PLANATM_INFOATM], [CommentaireATM_INFOATM], [Version_INFOATM], [DescriptifATMEN_INFOATM], [LocalisationATMEN_INFOATM], [CommentairesATMEN_INFOATM], [Reparabilite_INFOATM], [HARD/SOFT], [Poids_INFOATM], [Temps generique de la tache_INFOATM], [Temps generique de la tache SOFTWARE_INFOATM], [Temps analytique cumule de la tache HARDWARE_INFOATM], [Temps analytique cumule de la tache SOFTWARE_INFOATM] FROM [Info_ATM$]"
    
    ' Requete Taches ATM
    sqlTaches = "SELECT [DTR_TACHEATM], [Langue_TACHEATM], [Rang_TACHEATM], [HARD/SOFT], [Categorie_TACHEATM], [Niveau qualif_TACHEATM], [Description_TACHEATM], [CriteresDimensionnels_TACHEATM], [OutillageSpecifique_TACHEATM], [OutillageStandard_TACHEATM], [PiecedeRechangeConsommable_TACHEATM], [NbOperateurs_TACHEATM], [TempsEstimeHeures_TACHEATM], [Temps analytique cumule de la tache HARD_TACHEATM], [Temps analytique cumule de la tache SOFT_TACHEATM] FROM [Taches_ATM$]"
    
    ' Ouvrir recordsets
    Call open_REQSEL1(sqlInfo, "REQSEL_INFO")
    Call open_REQSEL1(sqlTaches, "REQSEL_TACHES")
    
End Sub

' =========================================================================
' PROCEDURE: CREER FEUILLES EXTRACTION ATM
' =========================================================================
Private Sub CreerFeuillesExtractionATM(ByRef wb As Workbook)
' Cree les feuilles "Info_ATM" et "Taches_ATM" avec en-tetes
    
    Dim wsInfo As Worksheet
    Dim wsTaches As Worksheet
    
    ' ======================================================================
    ' Feuille Info_ATM
    ' ======================================================================
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
        .Range("N1").Value = "Temps generique de la tache HARDWARE_INFOATM"
        .Range("O1").Value = "Temps generique de la tache SOFTWARE_INFOATM"
        .Range("P1").Value = "Temps analytique cumule de la tache HARDWARE_INFOATM"
        .Range("Q1").Value = "Temps analytique cumule de la tache SOFTWARE_INFOATM"
        .Range("R1").Value = "Connu en BDD ?"
    End With
    
    ' ======================================================================
    ' Feuille Taches_ATM
    ' ======================================================================
    wb.Sheets.Add Before:=wb.Worksheets(1)
    Set wsTaches = wb.ActiveSheet
    wsTaches.Name = "Taches_ATM"
    
    With wsTaches
        .Range("A1").Value = "DTR_TACHEATM"
        .Range("B1").Value = "Langue_TACHEATM"
        .Range("C1").Value = "Rang_TACHEATM"
        .Range("D1").Value = "HARD/SOFT"
        .Range("E1").Value = "Categorie_TACHEATM"
        .Range("F1").Value = "Niveau qualif_TACHEATM"
        .Range("G1").Value = "Description_TACHEATM"
        .Range("H1").Value = "CriteresDimensionnels_TACHEATM"
        .Range("I1").Value = "OutillageSpecifique_TACHEATM"
        .Range("J1").Value = "OutillageStandard_TACHEATM"
        .Range("K1").Value = "PiecedeRechangeConsommable_TACHEATM"
        .Range("L1").Value = "NbOperateurs_TACHEATM"
        .Range("M1").Value = "TempsEstimeHeures_TACHEATM"
        .Range("N1").Value = "Temps analytique cumule de la tache HARD_TACHEATM"
        .Range("O1").Value = "Temps analytique cumule de la tache SOFT_TACHEATM"
        .Range("P1").Value = "Pictogramme pour ATM2.0_TACHEATM"
        .Range("Q1").Value = "Texte / commentaire picto2_TACHEATM"
        .Range("R1").Value = "Connu en BDD ?"
    End With
    
End Sub

' =========================================================================
' PROCEDURE: EXTRAIRE INFO ATM
' =========================================================================
Private Sub ExtraireInfoATM(ByRef ws As Worksheet, ByVal langue As String)
' Extrait les donnees du cartouche ATM d'une feuille vers Info_ATM
    
    Dim wsInfo As Worksheet
    Dim derniereLigne As Long
    Dim dtr As String
    Dim hardSoft As String
    Dim nvMaint As String
    Dim tempsTotal As Double
    Dim dtrPresent As Boolean
    
    Set wsInfo = ws.Parent.Sheets("Info_ATM")
    derniereLigne = wsInfo.Range("A100000").End(xlUp).Row
    
    ' Recuperer donnees de base
    dtr = Left(ws.Name, 13)
    hardSoft = DeterminerHardSoft(ws.Name)
    nvMaint = ws.Range("I7").Value
    tempsTotal = RecupererTempsTotal(ws, langue)
    
    ' Verifier presence en BDD
    dtrPresent = RechercheValeurdansRS(REQSEL_INFO, dtr, "DTR_INFOATM")
    
    ' Remplir ligne Info_ATM
    With wsInfo
        .Cells(derniereLigne + 1, 1).Value = dtr                                        ' DTR
        .Cells(derniereLigne + 1, 2).Value = ws.Range("D5").Value                       ' Descriptif FR
        .Cells(derniereLigne + 1, 3).Value = ws.Range("D4").Value                       ' Designation
        .Cells(derniereLigne + 1, 4).Value = ""                                         ' Localisation
        .Cells(derniereLigne + 1, 5).Value = ws.Range("I8").Value                       ' Plan
        .Cells(derniereLigne + 1, 6).Value = ws.Range("D10").Value                      ' Commentaires FR
        .Cells(derniereLigne + 1, 7).Value = "A1"                                       ' Version
        
        ' Descriptif EN (prefixe REPLACE si necessaire)
        If Not IsError(ws.Range("D5").Value) And ws.Range("D5").Value <> "" Then
            .Cells(derniereLigne + 1, 8).Value = "REPLACE " & ws.Range("D5").Value
        End If
        
        .Cells(derniereLigne + 1, 9).Value = ""                                         ' Localisation EN
        .Cells(derniereLigne + 1, 10).Value = ws.Range("D10").Value                     ' Commentaires EN
        .Cells(derniereLigne + 1, 11).Value = ws.Range("I9").Value                      ' Reparabilite
        .Cells(derniereLigne + 1, 12).Value = hardSoft                                  ' HARD/SOFT
        .Cells(derniereLigne + 1, 13).Value = ws.Range("I6").Value                      ' Poids
        
        ' Temps generique selon HARD/SOFT
        Select Case hardSoft
            Case "HARD"
                .Cells(derniereLigne + 1, 14).Value = tempsTotal                        ' Temps HARD
            Case "SOFT"
                .Cells(derniereLigne + 1, 15).Value = tempsTotal                        ' Temps SOFT
            Case Else
                .Cells(derniereLigne + 1, 14).Value = tempsTotal                        ' Temps par defaut
        End Select
        
        ' Marqueur BDD
        If dtrPresent Then
            .Cells(derniereLigne + 1, 18).Value = "Connu"
        Else
            .Cells(derniereLigne + 1, 18).Value = "Inconnu"
        End If
    End With
    
End Sub

' =========================================================================
' PROCEDURE: EXTRAIRE TACHES ATM
' =========================================================================
Private Sub ExtraireTachesATM(ByRef ws As Worksheet, ByVal langue As String)
' Extrait toutes les taches ATM d'une feuille vers Taches_ATM
    
    Dim wsTaches As Worksheet
    Dim derniereLigne As Long
    Dim dtr As String
    Dim hardSoft As String
    Dim nvMaint As String
    Dim tempsTotal As Double
    Dim dtrPresent As Boolean
    Dim categorieTache As String
    Dim ligneDebut As Long, ligneFin As Long
    Dim i As Long
    
    Set wsTaches = ws.Parent.Sheets("Taches_ATM")
    
    ' Recuperer donnees de base
    dtr = Left(ws.Name, 13)
    hardSoft = DeterminerHardSoft(ws.Name)
    nvMaint = ws.Range("I7").Value
    tempsTotal = RecupererTempsTotal(ws, langue)
    dtrPresent = RechercheValeurdansRS(REQSEL_TACHES, dtr, "DTR_TACHEATM")
    
    ' Verifier si feuille contient des taches
    If EstFeuilleVide(ws) Then Exit Sub
    
    ' Determiner plage de taches
    ligneDebut = 13
    ligneFin = ws.Range("B10000").End(xlUp).Row
    
    ' Parcourir toutes les lignes de taches
    For i = ligneDebut To ligneFin
        
        ' Ignorer lignes vides
        If Len(ws.Cells(i, 2).Value) = 0 Then
            ' Ligne vide ? passer
            
        ' Ligne categorie (1 seul caractere)
        ElseIf Len(ws.Cells(i, 2).Value) = 1 Then
            categorieTache = ws.Cells(i, 3).Value
            
        ' Ligne tache (rang avec format x-y)
        Else
            derniereLigne = wsTaches.Range("A100000").End(xlUp).Row
            
            Call AjouterTacheATM(wsTaches, derniereLigne + 1, ws, i, dtr, langue, hardSoft, nvMaint, tempsTotal, categorieTache, dtrPresent)
        End If
        
    Next i
    
End Sub

' =========================================================================
' PROCEDURE: AJOUTER TACHE ATM
' =========================================================================
Private Sub AjouterTacheATM(ByRef wsTaches As Worksheet, ByVal ligne As Long, ByRef wsDTR As Worksheet, ByVal ligneDTR As Long, ByVal dtr As String, ByVal langue As String, ByVal hardSoft As String, ByVal nvMaint As String, ByVal tempsTotal As Double, ByVal categorie As String, ByVal dtrPresent As Boolean)
' Ajoute une ligne de tache dans la feuille Taches_ATM
    
    With wsTaches
        .Cells(ligne, 1).Value = dtr                                    ' DTR
        .Cells(ligne, 2).Value = langue                                 ' Langue
        .Cells(ligne, 3).Value = "'" & wsDTR.Cells(ligneDTR, 2).Value   ' Rang (avec ')
        .Cells(ligne, 4).Value = hardSoft                               ' HARD/SOFT
        .Cells(ligne, 5).Value = categorie                              ' Categorie
        .Cells(ligne, 6).Value = nvMaint                                ' Niveau qualif
        .Cells(ligne, 7).Value = wsDTR.Cells(ligneDTR, 3).Value         ' Description
        .Cells(ligne, 8).Value = wsDTR.Cells(ligneDTR, 5).Value         ' Criteres dim
        .Cells(ligne, 9).Value = wsDTR.Cells(ligneDTR, 6).Value         ' Outillage spe
        .Cells(ligne, 10).Value = wsDTR.Cells(ligneDTR, 7).Value        ' Outillage std
        .Cells(ligne, 11).Value = wsDTR.Cells(ligneDTR, 8).Value        ' Pieces rechange
        .Cells(ligne, 12).Value = wsDTR.Cells(ligneDTR, 9).Value        ' Nb operateurs
        .Cells(ligne, 13).Value = wsDTR.Cells(ligneDTR, 10).Value       ' Temps estime
        
        ' Temps analytique cumule selon HARD/SOFT
        Select Case hardSoft
            Case "HARD"
                .Cells(ligne, 14).Value = tempsTotal                    ' Temps HARD
            Case "SOFT"
                .Cells(ligne, 15).Value = tempsTotal                    ' Temps SOFT
            Case Else
                .Cells(ligne, 14).Value = tempsTotal                    ' Temps par defaut
        End Select
        
        .Cells(ligne, 16).Value = ""                                    ' Pictogramme ATM 2.0
        .Cells(ligne, 17).Value = ""                                    ' Texte picto
        
        ' Marqueur BDD
        If dtrPresent Then
            .Cells(ligne, 18).Value = "Connu"
        Else
            .Cells(ligne, 18).Value = "Inconnu"
        End If
    End With
    
End Sub

' =========================================================================
' FONCTION: DETERMINER HARD/SOFT
' =========================================================================
Private Function DeterminerHardSoft(ByVal nomFeuille As String) As String
' Determine si la fiche est HARD ou SOFT selon le nom de feuille
    
    If Right(nomFeuille, 4) = "HARD" Then
        DeterminerHardSoft = "HARD"
    ElseIf Right(nomFeuille, 4) = "SOFT" Then
        DeterminerHardSoft = "SOFT"
    Else
        DeterminerHardSoft = ""
    End If
    
End Function

' =========================================================================
' FONCTION: RECUPERER TEMPS TOTAL
' =========================================================================
Private Function RecupererTempsTotal(ByRef ws As Worksheet, ByVal langue As String) As Double
' Recupere le temps total depuis la cellule "Temps total :" ou "Total time :"
    
    Dim celluleTemps As Range
    Dim texteRecherche As String

    ' Determiner texte a rechercher
    If langue = "FR" Then
        texteRecherche = "Temps total :"
    Else
        texteRecherche = "Total time :"
    End If
    
    ' Rechercher cellule
    Set celluleTemps = ws.Cells.Find(What:=texteRecherche, LookIn:=xlValues, lookat:=xlWhole)
    
    If Not celluleTemps Is Nothing Then
        RecupererTempsTotal = SafeNum(ws.Cells(celluleTemps.Row, 10).Value, 0)
    Else
        RecupererTempsTotal = 0
    End If

End Function

' =========================================================================
' FONCTION: EST FEUILLE VIDE
' =========================================================================
Private Function EstFeuilleVide(ByRef ws As Worksheet) As Boolean
' Verifie si la feuille ATM ne contient aucune tache
    
    If ws.Cells(14, 3).Value = "" And ws.Cells(17, 3).Value = "" And ws.Cells(20, 3).Value = "" And ws.Cells(24, 3).Value = "" Then
        EstFeuilleVide = True
    Else
        EstFeuilleVide = False
    End If
    
End Function

' =========================================================================
' PROCEDURE: FORMATER FEUILLES EXTRACTION ATM
' =========================================================================
Private Sub FormaterFeuillesExtractionATM(ByRef wb As Workbook)
' Desactive le wrap text sur les feuilles Info_ATM et Taches_ATM
    
    Dim wsInfo As Worksheet
    Dim wsTaches As Worksheet
    
    Set wsInfo = wb.Sheets("Info_ATM")
    Set wsTaches = wb.Sheets("Taches_ATM")
    
    ' Desactiver wrap text Info_ATM
    wsInfo.Activate
    wsInfo.Cells.Select
    With Selection
        .WrapText = False
    End With
    
    ' Desactiver wrap text Taches_ATM
    wsTaches.Activate
    wsTaches.Cells.Select
    With Selection
        .WrapText = False
    End With
    
    ' Retourner sur Info_ATM
    wsInfo.Activate
    wsInfo.Range("A1").Select
    
End Sub


