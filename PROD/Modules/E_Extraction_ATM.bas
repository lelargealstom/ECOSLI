Attribute VB_Name = "E_Extraction_ATM"
'--------------------------------------------------------------------------------------------------------
'
'                                       Generateur Livrables SLI
'                                       Module M_Extraction_ATM (Version Refactorisée)
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
' PROCÉDURE PRINCIPALE: EXTRACTION FICHES ATM
' =========================================================================
Public Sub Extraction_Fiches_ATM(ByRef wb As Workbook, ByVal langue As String, ByVal nbFiches As Integer)
' Extrait toutes les données ATM d'un fichier et les structure au format BDD
    
    Dim ws As Worksheet
    Dim indexFiche As Integer
    
    On Error GoTo ErrorHandler
    
    ' ======================================================================
    ' ÉTAPE 1 : Initialisation Recordsets BDD
    ' ======================================================================
    Call InitialiserRecordsetsATM
    
    ' ======================================================================
    ' ÉTAPE 2 : Création feuilles extraction
    ' ======================================================================
    Call CreerFeuillesExtractionATM(wb)
    
    ' ======================================================================
    ' ÉTAPE 3 : Extraction données de chaque fiche DTR
    ' ======================================================================
    indexFiche = 0
    
    For Each ws In wb.Worksheets
        If ws.Name Like "*DTR*" Then
            indexFiche = indexFiche + 1
            
            ' Extraire Info ATM (cartouche)
            Call ExtraireInfoATM(ws, langue)
            
            ' Extraire Tâches ATM
            Call ExtraireTachesATM(ws, langue)
            
            ' Mettre à jour barre progression
            Call barprogress(indexFiche, nbFiches)
        End If
    Next ws
    
    ' ======================================================================
    ' ÉTAPE 4 : Mise en forme finale
    ' ======================================================================
    Call FormaterFeuillesExtractionATM(wb)
    
    ' Fermer recordsets
    On Error Resume Next
    REQSEL_INFO.Close
    REQSEL_TACHES.Close
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Erreur extraction ATM : " & Err.description, vbCritical
    On Error Resume Next
    REQSEL_INFO.Close
    REQSEL_TACHES.Close
End Sub

' =========================================================================
' PROCÉDURE: INITIALISER RECORDSETS ATM
' =========================================================================
Private Sub InitialiserRecordsetsATM()
' Crée les recordsets pour vérifier existence DTR en BDD
    
    Dim sqlInfo As String
    Dim sqlTaches As String
    
    Set REQSEL_INFO = CreateObject("ADODB.Recordset")
    Set REQSEL_TACHES = CreateObject("ADODB.Recordset")
    
    ' Requête Info ATM
    sqlInfo = "SELECT [DTR_INFOATM], [DescriptifATMFR_INFOATM], [DesignationATM_INFOATM], " & _
              "[LocalisationATM_INFOATM], [PLANATM_INFOATM], [CommentaireATM_INFOATM], " & _
              "[Version_INFOATM], [DescriptifATMEN_INFOATM], [LocalisationATMEN_INFOATM], " & _
              "[CommentairesATMEN_INFOATM], [Reparabilite_INFOATM], [HARD/SOFT], [Poids_INFOATM], " & _
              "[Temps generique de la tâche_INFOATM], [Temps generique de la tache SOFTWARE_INFOATM], " & _
              "[Temps analytique cumule de la tache HARDWARE_INFOATM], " & _
              "[Temps analytique cumule de la tache SOFTWARE_INFOATM] " & _
              "FROM [Info_ATM$]"
    
    ' Requête Tâches ATM
    sqlTaches = "SELECT [DTR_TACHEATM], [Langue_TACHEATM], [Rang_TACHEATM], [HARD/SOFT], " & _
                "[Categorie_TACHEATM], [Niveau qualif_TACHEATM], [Description_TACHEATM], " & _
                "[CriteresDimensionnels_TACHEATM], [OutillageSpecifique_TACHEATM], " & _
                "[OutillageStandard_TACHEATM], [PiecedeRechangeConsommable_TACHEATM], " & _
                "[NbOperateurs_TACHEATM], [TempsEstimeHeures_TACHEATM], " & _
                "[Temps analytique cumule de la tache HARD_TACHEATM], " & _
                "[Temps analytique cumule de la tache SOFT_TACHEATM] " & _
                "FROM [Taches_ATM$]"
    
    ' Ouvrir recordsets
    Call open_REQSEL1(sqlInfo, "REQSEL_INFO")
    Call open_REQSEL1(sqlTaches, "REQSEL_TACHES")
    
End Sub

' =========================================================================
' PROCÉDURE: CRÉER FEUILLES EXTRACTION ATM
' =========================================================================
Private Sub CreerFeuillesExtractionATM(ByRef wb As Workbook)
' Crée les feuilles "Info_ATM" et "Taches_ATM" avec en-têtes
    
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
        .Range("N1").Value = "Temps générique de la tâche HARDWARE_INFOATM"
        .Range("O1").Value = "Temps générique de la tâche SOFTWARE_INFOATM"
        .Range("P1").Value = "Temps analytique cumulé de la tache HARDWARE_INFOATM"
        .Range("Q1").Value = "Temps analytique cumulé de la tache SOFTWARE_INFOATM"
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
        .Range("H1").Value = "CritèresDimensionnels_TACHEATM"
        .Range("I1").Value = "OutillageSpécifique_TACHEATM"
        .Range("J1").Value = "OutillageStandard_TACHEATM"
        .Range("K1").Value = "PiècedeRechangeConsommable_TACHEATM"
        .Range("L1").Value = "NbOpérateurs_TACHEATM"
        .Range("M1").Value = "TempsEstiméHeures_TACHEATM"
        .Range("N1").Value = "Temps analytique cumulé de la tache HARD_TACHEATM"
        .Range("O1").Value = "Temps analytique cumulé de la tache SOFT_TACHEATM"
        .Range("P1").Value = "Pictogramme pour ATM2.0_TACHEATM"
        .Range("Q1").Value = "Texte / commentaire picto2_TACHEATM"
        .Range("R1").Value = "Connu en BDD ?"
    End With
    
End Sub

' =========================================================================
' PROCÉDURE: EXTRAIRE INFO ATM
' =========================================================================
Private Sub ExtraireInfoATM(ByRef ws As Worksheet, ByVal langue As String)
' Extrait les données du cartouche ATM d'une feuille vers Info_ATM
    
    Dim wsInfo As Worksheet
    Dim derniereLigne As Long
    Dim dtr As String
    Dim HardSoft As String
    Dim nvMaint As String
    Dim tempsTotal As Double
    Dim dtrPresent As Boolean
    
    Set wsInfo = ws.Parent.Sheets("Info_ATM")
    derniereLigne = wsInfo.Range("A100000").End(xlUp).row
    
    ' Récupérer données de base
    dtr = Left(ws.Name, 13)
    HardSoft = DeterminerHardSoft(ws.Name)
    nvMaint = ws.Range("I7").Value
    tempsTotal = RecupererTempsTotal(ws, langue)
    
    ' Vérifier présence en BDD
    dtrPresent = RechercheValeurdansRS(REQSEL_INFO, dtr, "DTR_INFOATM")
    
    ' Remplir ligne Info_ATM
    With wsInfo
        .Cells(derniereLigne + 1, 1).Value = dtr                                        ' DTR
        .Cells(derniereLigne + 1, 2).Value = ws.Range("D5").Value                       ' Descriptif FR
        .Cells(derniereLigne + 1, 3).Value = ws.Range("D4").Value                       ' Désignation
        .Cells(derniereLigne + 1, 4).Value = ""                                         ' Localisation
        .Cells(derniereLigne + 1, 5).Value = ws.Range("I8").Value                       ' Plan
        .Cells(derniereLigne + 1, 6).Value = ws.Range("D10").Value                      ' Commentaires FR
        .Cells(derniereLigne + 1, 7).Value = "A1"                                       ' Version
        
        ' Descriptif EN (préfixe REPLACE si nécessaire)
        If Not IsError(ws.Range("D5").Value) And ws.Range("D5").Value <> "" Then
            .Cells(derniereLigne + 1, 8).Value = "REPLACE " & ws.Range("D5").Value
        End If
        
        .Cells(derniereLigne + 1, 9).Value = ""                                         ' Localisation EN
        .Cells(derniereLigne + 1, 10).Value = ws.Range("D10").Value                     ' Commentaires EN
        .Cells(derniereLigne + 1, 11).Value = ws.Range("I9").Value                      ' Réparabilité
        .Cells(derniereLigne + 1, 12).Value = HardSoft                                  ' HARD/SOFT
        .Cells(derniereLigne + 1, 13).Value = ws.Range("I6").Value                      ' Poids
        
        ' Temps générique selon HARD/SOFT
        Select Case HardSoft
            Case "HARD"
                .Cells(derniereLigne + 1, 14).Value = tempsTotal                        ' Temps HARD
            Case "SOFT"
                .Cells(derniereLigne + 1, 15).Value = tempsTotal                        ' Temps SOFT
            Case Else
                .Cells(derniereLigne + 1, 14).Value = tempsTotal                        ' Temps par défaut
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
' PROCÉDURE: EXTRAIRE TÂCHES ATM
' =========================================================================
Private Sub ExtraireTachesATM(ByRef ws As Worksheet, ByVal langue As String)
' Extrait toutes les tâches ATM d'une feuille vers Taches_ATM
    
    Dim wsTaches As Worksheet
    Dim derniereLigne As Long
    Dim dtr As String
    Dim HardSoft As String
    Dim nvMaint As String
    Dim tempsTotal As Double
    Dim dtrPresent As Boolean
    Dim categorieTache As String
    Dim ligneDebut As Long, ligneFin As Long
    Dim i As Long
    
    Set wsTaches = ws.Parent.Sheets("Taches_ATM")
    
    ' Récupérer données de base
    dtr = Left(ws.Name, 13)
    HardSoft = DeterminerHardSoft(ws.Name)
    nvMaint = ws.Range("I7").Value
    tempsTotal = RecupererTempsTotal(ws, langue)
    dtrPresent = RechercheValeurdansRS(REQSEL_TACHES, dtr, "DTR_TACHEATM")
    
    ' Vérifier si feuille contient des tâches
    If EstFeuilleVide(ws) Then Exit Sub
    
    ' Déterminer plage de tâches
    ligneDebut = 13
    ligneFin = ws.Range("B10000").End(xlUp).row
    
    ' Parcourir toutes les lignes de tâches
    For i = ligneDebut To ligneFin
        
        ' Ignorer lignes vides
        If Len(ws.Cells(i, 2).Value) = 0 Then
            ' Ligne vide ? passer
            
        ' Ligne catégorie (1 seul caractère)
        ElseIf Len(ws.Cells(i, 2).Value) = 1 Then
            categorieTache = ws.Cells(i, 3).Value
            
        ' Ligne tâche (rang avec format x-y)
        Else
            derniereLigne = wsTaches.Range("A100000").End(xlUp).row
            
            Call AjouterTacheATM(wsTaches, derniereLigne + 1, ws, i, _
                                 dtr, langue, HardSoft, nvMaint, tempsTotal, _
                                 categorieTache, dtrPresent)
        End If
        
    Next i
    
End Sub

' =========================================================================
' PROCÉDURE: AJOUTER TÂCHE ATM
' =========================================================================
Private Sub AjouterTacheATM(ByRef wsTaches As Worksheet, _
                            ByVal ligne As Long, _
                            ByRef wsDTR As Worksheet, _
                            ByVal ligneDTR As Long, _
                            ByVal dtr As String, _
                            ByVal langue As String, _
                            ByVal HardSoft As String, _
                            ByVal nvMaint As String, _
                            ByVal tempsTotal As Double, _
                            ByVal categorie As String, _
                            ByVal dtrPresent As Boolean)
' Ajoute une ligne de tâche dans la feuille Taches_ATM
    
    With wsTaches
        .Cells(ligne, 1).Value = dtr                                    ' DTR
        .Cells(ligne, 2).Value = langue                                 ' Langue
        .Cells(ligne, 3).Value = "'" & wsDTR.Cells(ligneDTR, 2).Value   ' Rang (avec ')
        .Cells(ligne, 4).Value = HardSoft                               ' HARD/SOFT
        .Cells(ligne, 5).Value = categorie                              ' Catégorie
        .Cells(ligne, 6).Value = nvMaint                                ' Niveau qualif
        .Cells(ligne, 7).Value = wsDTR.Cells(ligneDTR, 3).Value         ' Description
        .Cells(ligne, 8).Value = wsDTR.Cells(ligneDTR, 5).Value         ' Critères dim
        .Cells(ligne, 9).Value = wsDTR.Cells(ligneDTR, 6).Value         ' Outillage spé
        .Cells(ligne, 10).Value = wsDTR.Cells(ligneDTR, 7).Value        ' Outillage std
        .Cells(ligne, 11).Value = wsDTR.Cells(ligneDTR, 8).Value        ' Pièces rechange
        .Cells(ligne, 12).Value = wsDTR.Cells(ligneDTR, 9).Value        ' Nb opérateurs
        .Cells(ligne, 13).Value = wsDTR.Cells(ligneDTR, 10).Value       ' Temps estimé
        
        ' Temps analytique cumulé selon HARD/SOFT
        Select Case HardSoft
            Case "HARD"
                .Cells(ligne, 14).Value = tempsTotal                    ' Temps HARD
            Case "SOFT"
                .Cells(ligne, 15).Value = tempsTotal                    ' Temps SOFT
            Case Else
                .Cells(ligne, 14).Value = tempsTotal                    ' Temps par défaut
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
' FONCTION: DÉTERMINER HARD/SOFT
' =========================================================================
Private Function DeterminerHardSoft(ByVal nomFeuille As String) As String
' Détermine si la fiche est HARD ou SOFT selon le nom de feuille
    
    If Right(nomFeuille, 4) = "HARD" Then
        DeterminerHardSoft = "HARD"
    ElseIf Right(nomFeuille, 4) = "SOFT" Then
        DeterminerHardSoft = "SOFT"
    Else
        DeterminerHardSoft = ""
    End If
    
End Function

' =========================================================================
' FONCTION: RÉCUPÉRER TEMPS TOTAL
' =========================================================================
Private Function RecupererTempsTotal(ByRef ws As Worksheet, ByVal langue As String) As Double
' Récupère le temps total depuis la cellule "Temps total :" ou "Total time :"
    
    Dim celluleTemps As Range
    Dim texteRecherche As String
    
    On Error GoTo ErrorHandler
    
    ' Déterminer texte à rechercher
    If langue = "FR" Then
        texteRecherche = "Temps total :"
    Else
        texteRecherche = "Total time :"
    End If
    
    ' Rechercher cellule
    Set celluleTemps = ws.Cells.Find(What:=texteRecherche, LookIn:=xlValues, lookat:=xlWhole)
    
    If Not celluleTemps Is Nothing Then
        RecupererTempsTotal = SafeNum(ws.Cells(celluleTemps.row, 10).Value, 0)
    Else
        RecupererTempsTotal = 0
    End If
    
    Exit Function
    
ErrorHandler:
    RecupererTempsTotal = 0
    
End Function

' =========================================================================
' FONCTION: EST FEUILLE VIDE
' =========================================================================
Private Function EstFeuilleVide(ByRef ws As Worksheet) As Boolean
' Vérifie si la feuille ATM ne contient aucune tâche
    
    If ws.Cells(14, 3).Value = "" And _
       ws.Cells(17, 3).Value = "" And _
       ws.Cells(20, 3).Value = "" And _
       ws.Cells(24, 3).Value = "" Then
        EstFeuilleVide = True
    Else
        EstFeuilleVide = False
    End If
    
End Function

' =========================================================================
' PROCÉDURE: FORMATER FEUILLES EXTRACTION ATM
' =========================================================================
Private Sub FormaterFeuillesExtractionATM(ByRef wb As Workbook)
' Désactive le wrap text sur les feuilles Info_ATM et Taches_ATM
    
    Dim wsInfo As Worksheet
    Dim wsTaches As Worksheet
    
    Set wsInfo = wb.Sheets("Info_ATM")
    Set wsTaches = wb.Sheets("Taches_ATM")
    
    ' Désactiver wrap text Info_ATM
    wsInfo.Activate
    wsInfo.Cells.Select
    With Selection
        .WrapText = False
    End With
    
    ' Désactiver wrap text Taches_ATM
    wsTaches.Activate
    wsTaches.Cells.Select
    With Selection
        .WrapText = False
    End With
    
    ' Retourner sur Info_ATM
    wsInfo.Activate
    wsInfo.Range("A1").Select
    
End Sub


