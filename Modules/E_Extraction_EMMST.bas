Attribute VB_Name = "E_Extraction_EMMST"
'--------------------------------------------------------------------------------------------------------
'
'                                       Generateur Livrables SLI
'                                       Module N_Extraction_EMMST (Version Refactorisee)
'                                       Developpement : LE LARGE Thomas (LGM/Alstom)
'                                       Module d'extraction fiches EMMST vers format BDD
'
'-----------------------------------------------------------------------------------------------------------
Option Compare Text
Option Explicit

' Variables module
Private REQSEL_EMMST As Object

' =========================================================================
' PROCEDURE PRINCIPALE: EXTRACTION FICHES EMMST
' =========================================================================
Public Sub Extraction_Fiches_EMMST(ByRef wb As Workbook, ByVal langue As String, ByVal nbFiches As Integer)
' Extrait toutes les donnees EMMST d'un fichier et les structure au format BDD
    
    Dim ws As Worksheet
    Dim indexFiche As Integer
    
    On Error GoTo ErrorHandler
    
    ' ======================================================================
    ' ETAPE 1 : Initialisation Recordset BDD
    ' ======================================================================
    Call InitialiserRecordsetEMMST
    
    ' ======================================================================
    ' ETAPE 2 : Creation feuille extraction
    ' ======================================================================
    Call CreerFeuilleExtractionEMMST(wb)
    
    ' ======================================================================
    ' ETAPE 3 : Extraction donnees de chaque fiche DTR
    ' ======================================================================
    indexFiche = 0
    
    For Each ws In wb.Worksheets
        If ws.Name Like "*DTR*" Then
            indexFiche = indexFiche + 1
            
            ' Extraire donnees EMMST
            Call ExtraireEMMST(ws, langue)
            
            ' Mettre a jour barre progression
            Call barprogress(indexFiche, nbFiches)
        End If
    Next ws
    
    ' ======================================================================
    ' ETAPE 4 : Mise en forme finale
    ' ======================================================================
    Call FormaterFeuilleExtractionEMMST(wb)
    
    ' Fermer recordset
    On Error Resume Next
    REQSEL_EMMST.Close
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Erreur extraction EMMST : " & Err.description, vbCritical
    On Error Resume Next
    REQSEL_EMMST.Close
End Sub

' =========================================================================
' PROCEDURE: INITIALISER RECORDSET EMMST
' =========================================================================
Private Sub InitialiserRecordsetEMMST()
' Cree le recordset pour verifier existence DTR en BDD
    
    Dim sql As String
    
    Set REQSEL_EMMST = CreateObject("ADODB.Recordset")
    
    ' Requete SQL complete (50 colonnes)
    sql = "SELECT [DTR_EMMST], [DimensionNuLongueur_EMMST], [DimensionNuLargeur_EMMST], [DimensionNuHauteur_EMMST], [DimensionNuPoids], [DimensionConditionneLongueur_EMMST], [DimensionConditionneLargeur_EMMST], [DimensionConditionneHauteur_EMMST], [DimensionConditionnePoids_EMMST], [DatePeremption_EMMST], [UniteConditionnement_EMMST], [CommentaireEmballage_EMMST], [CommentaireMarquage_EMMST], [CommentaireManutention_EMMST], [CommentaireStockage_EMMST], [CommentaireTransport_EMMST], [TypeOutil_EMMST], [DureePeremption_EMMST], [Reutilisable_EMMST], [EtiquetteEmballage_EMMST], [MarquageEquipement_EMMST], [ManutentionEquipement_EMMST], [OutillageSpecifiqueManutention_EMMST], [PrecautionManutention_EMMST], [PositionStockage_EMMST], [StockageExterieur_EMMST], [StockageSousAbri_EMMST], [StockageAirLibre_EMMST], [StockageHorsPoussiere_EMMST], [StockageHorsEau_EMMST], [StockageInterieur_EMMST], [StockageChaud_EMMST], [StockageFroid_EMMST], [StockageHorsLumiere_EMMST], [Gerbable_EMMST], [MoyenStockage_EMMST], [PrecautionStockage_EMMST], [MaintenanceStockage_EMMST], [MoyenTransport_EMMST], [OutillageSpecifiqueTransport_EMMST], [PrescriptionTransport_EMMST], [PrecautionTransport_EMMST], [Version_EMMST], [DatePeremptionEN_EMMST], [UniteConditionneEN_EMMST], [CommentaireEmballageEN_EMMST], [CommentaireMarquageEN_EMMST], [CommentaireManutentionEN_EMMST], [CommentaireStockageEN_EMMST], [CommentaireTransportEN_EMMST] FROM [EMMST$]"
    
    Call open_REQSEL(sql)
    Set REQSEL_EMMST = REQSEL
    
End Sub

' =========================================================================
' PROCEDURE: CREER FEUILLE EXTRACTION EMMST
' =========================================================================
Private Sub CreerFeuilleExtractionEMMST(ByRef wb As Workbook)
' Cree la feuille "EMMST" avec tous les en-tetes (51 colonnes)
    
    Dim wsEMMST As Worksheet
    
    wb.Sheets.Add Before:=wb.Worksheets(1)
    Set wsEMMST = wb.ActiveSheet
    wsEMMST.Name = "EMMST"
    
    ' En-tetes (51 colonnes)
    With wsEMMST
        .Range("A1").Value = "DTR_EMMST"
        .Range("B1").Value = "DimensionNuLongueur_EMMST"
        .Range("C1").Value = "DimensionNuLargeur_EMMST"
        .Range("D1").Value = "DimensionNuHauteur_EMMST"
        .Range("E1").Value = "DimensionNuPoids"
        .Range("F1").Value = "DimensionConditionneLongueur_EMMST"
        .Range("G1").Value = "DimensionConditionneLargeur_EMMST"
        .Range("H1").Value = "DimensionConditionneHauteur_EMMST"
        .Range("I1").Value = "DimensionConditionnePoids_EMMST"
        .Range("J1").Value = "DatePeremption_EMMST"
        .Range("K1").Value = "UniteConditionnement_EMMST"
        .Range("L1").Value = "CommentaireEmballage_EMMST"
        .Range("M1").Value = "CommentaireMarquage_EMMST"
        .Range("N1").Value = "CommentaireManutention_EMMST"
        .Range("O1").Value = "CommentaireStockage_EMMST"
        .Range("P1").Value = "CommentaireTransport_EMMST"
        .Range("Q1").Value = "TypeOutil_EMMST"
        .Range("R1").Value = "DureePeremption_EMMST"
        .Range("S1").Value = "Reutilisable_EMMST"
        .Range("T1").Value = "EtiquetteEmballage_EMMST"
        .Range("U1").Value = "MarquageEquipement_EMMST"
        .Range("V1").Value = "ManutentionEquipement_EMMST"
        .Range("W1").Value = "OutillageSpecifiqueManutention_EMMST"
        .Range("X1").Value = "PrecautionManutention_EMMST"
        .Range("Y1").Value = "PositionStockage_EMMST"
        .Range("Z1").Value = "StockageExterieur_EMMST"
        .Range("AA1").Value = "StockageSousAbri_EMMST"
        .Range("AB1").Value = "StockageAirLibre_EMMST"
        .Range("AC1").Value = "StockageHorsPoussiere_EMMST"
        .Range("AD1").Value = "StockageHorsEau_EMMST"
        .Range("AE1").Value = "StockageInterieur_EMMST"
        .Range("AF1").Value = "StockageChaud_EMMST"
        .Range("AG1").Value = "StockageFroid_EMMST"
        .Range("AH1").Value = "StockageHorsLumiere_EMMST"
        .Range("AI1").Value = "Gerbable_EMMST"
        .Range("AJ1").Value = "MoyenStockage_EMMST"
        .Range("AK1").Value = "PrecautionStockage_EMMST"
        .Range("AL1").Value = "MaintenanceStockage_EMMST"
        .Range("AM1").Value = "MoyenTransport_EMMST"
        .Range("AN1").Value = "OutillageSpecifiqueTransport_EMMST"
        .Range("AO1").Value = "PrescriptionTransport_EMMST"
        .Range("AP1").Value = "PrecautionTransport_EMMST"
        .Range("AQ1").Value = "Version_EMMST"
        .Range("AR1").Value = "DatePeremptionEN_EMMST"
        .Range("AS1").Value = "UniteConditionneEN_EMMST"
        .Range("AT1").Value = "CommentaireEmballageEN_EMMST"
        .Range("AU1").Value = "CommentaireMarquageEN_EMMST"
        .Range("AV1").Value = "CommentaireManutentionEN_EMMST"
        .Range("AW1").Value = "CommentaireStockageEN_EMMST"
        .Range("AX1").Value = "CommentaireTransportEN_EMMST"
        .Range("AY1").Value = "Connu en BDD ?"
    End With
    
End Sub

' =========================================================================
' PROCEDURE: EXTRAIRE EMMST
' =========================================================================
Private Sub ExtraireEMMST(ByRef ws As Worksheet, ByVal langue As String)
' Extrait toutes les donnees EMMST d'une feuille vers la feuille EMMST
    
    Dim wsEMMST As Worksheet
    Dim derniereLigne As Long
    Dim dtr As String
    Dim dtrPresent As Boolean
    
    Set wsEMMST = ws.Parent.Sheets("EMMST")
    derniereLigne = wsEMMST.Range("A100000").End(xlUp).Row
    
    ' Recuperer DTR
    dtr = ws.Range("H6").Value
    
    ' Verifier presence en BDD
    dtrPresent = RechercheValeurdansRS(REQSEL_EMMST, dtr, "DTR_EMMST")
    
    ' Remplir donnees par categorie
    Call RemplirDonneesBase(wsEMMST, derniereLigne + 1, ws, dtr)
    Call RemplirCommentaires(wsEMMST, derniereLigne + 1, ws, langue)
    Call RemplirTypeOutil(wsEMMST, derniereLigne + 1, ws)
    Call RemplirReutilisable(wsEMMST, derniereLigne + 1, ws)
    Call RemplirEtiquettes(wsEMMST, derniereLigne + 1, ws)
    Call RemplirManutention(wsEMMST, derniereLigne + 1, ws)
    Call RemplirStockage(wsEMMST, derniereLigne + 1, ws)
    Call RemplirTransport(wsEMMST, derniereLigne + 1, ws)
    Call RemplirVersion(wsEMMST, derniereLigne + 1, ws, langue)
    
    ' Marqueur BDD
    If dtrPresent Then
        wsEMMST.Cells(derniereLigne + 1, 51).Value = "Connu"
    Else
        wsEMMST.Cells(derniereLigne + 1, 51).Value = "Inconnu"
    End If
    
End Sub

' =========================================================================
' PROCEDURE: REMPLIR DONNEES BASE
' =========================================================================
Private Sub RemplirDonneesBase(ByRef wsEMMST As Worksheet, ByVal ligne As Long, ByRef wsDTR As Worksheet, ByVal dtr As String)
' Remplit les colonnes 1-11 : DTR, Dimensions, Date, Unite
    
    With wsEMMST
        .Cells(ligne, 1).Value = dtr                            ' DTR
        .Cells(ligne, 2).Value = wsDTR.Range("F10").Value       ' Dim Nu Longueur
        .Cells(ligne, 3).Value = wsDTR.Range("H10").Value       ' Dim Nu Largeur
        .Cells(ligne, 4).Value = wsDTR.Range("J10").Value       ' Dim Nu Hauteur
        .Cells(ligne, 5).Value = wsDTR.Range("L10").Value       ' Dim Nu Poids
        .Cells(ligne, 6).Value = wsDTR.Range("F12").Value       ' Dim Cond Longueur
        .Cells(ligne, 7).Value = wsDTR.Range("H12").Value       ' Dim Cond Largeur
        .Cells(ligne, 8).Value = wsDTR.Range("J12").Value       ' Dim Cond Hauteur
        .Cells(ligne, 9).Value = wsDTR.Range("L12").Value       ' Dim Cond Poids
        .Cells(ligne, 10).Value = wsDTR.Range("J14").Value      ' Date Peremption
        .Cells(ligne, 11).Value = wsDTR.Range("E16").Value      ' Unite Conditionnement
    End With
    
End Sub

' =========================================================================
' PROCEDURE: REMPLIR COMMENTAIRES
' =========================================================================
Private Sub RemplirCommentaires(ByRef wsEMMST As Worksheet, ByVal ligne As Long, ByRef wsDTR As Worksheet, ByVal langue As String)
' Remplit les commentaires selon la langue (FR ou EN)
    
    If langue = "FR" Then
        ' Commentaires FR (colonnes 12-16)
        wsEMMST.Cells(ligne, 12).Value = wsDTR.Range("G15").Value      ' Comm Emballage FR
        wsEMMST.Cells(ligne, 13).Value = wsDTR.Range("G20").Value      ' Comm Marquage FR
        wsEMMST.Cells(ligne, 14).Value = wsDTR.Range("G26").Value      ' Comm Manutention FR
        wsEMMST.Cells(ligne, 15).Value = wsDTR.Range("G40").Value      ' Comm Stockage FR
        wsEMMST.Cells(ligne, 16).Value = wsDTR.Range("G50").Value      ' Comm Transport FR
    Else
        ' Commentaires EN (colonnes 46-50)
        wsEMMST.Cells(ligne, 46).Value = wsDTR.Range("G15").Value      ' Comm Emballage EN
        wsEMMST.Cells(ligne, 47).Value = wsDTR.Range("G20").Value      ' Comm Marquage EN
        wsEMMST.Cells(ligne, 48).Value = wsDTR.Range("G26").Value      ' Comm Manutention EN
        wsEMMST.Cells(ligne, 49).Value = wsDTR.Range("G40").Value      ' Comm Stockage EN
        wsEMMST.Cells(ligne, 50).Value = wsDTR.Range("G50").Value      ' Comm Transport EN
    End If
    
End Sub

' =========================================================================
' PROCEDURE: REMPLIR TYPE OUTIL
' =========================================================================
Private Sub RemplirTypeOutil(ByRef wsEMMST As Worksheet, ByVal ligne As Long, ByRef wsDTR As Worksheet)
' Colonne 17 : Type Outil
    
    If wsDTR.Range("E4").Value = "X" Then
        wsEMMST.Cells(ligne, 17).Value = "Rechange"
    Else
        wsEMMST.Cells(ligne, 17).Value = ""
    End If
    
    ' Colonne 18 : Duree peremption (vide)
    wsEMMST.Cells(ligne, 18).Value = ""
    
End Sub

' =========================================================================
' PROCEDURE: REMPLIR REUTILISABLE
' =========================================================================
Private Sub RemplirReutilisable(ByRef wsEMMST As Worksheet, ByVal ligne As Long, ByRef wsDTR As Worksheet)
' Colonne 19 : Reutilisable (0 = Non, -1 = Oui)
    
    wsEMMST.Cells(ligne, 19).Value = ConvertirCocheBoolean(wsDTR.Range("F18").Value, wsDTR.Range("H18").Value)
    
End Sub

' =========================================================================
' PROCEDURE: REMPLIR ETIQUETTES
' =========================================================================
Private Sub RemplirEtiquettes(ByRef wsEMMST As Worksheet, ByVal ligne As Long, ByRef wsDTR As Worksheet)
' Colonnes 20-21 : Etiquette Emballage, Marquage Equipement
    
    ' Colonne 20 : Etiquette Emballage
    wsEMMST.Cells(ligne, 20).Value = ConvertirCocheBoolean(wsDTR.Range("E22").Value, wsDTR.Range("F22").Value)
    
    ' Colonne 21 : Marquage Equipement
    wsEMMST.Cells(ligne, 21).Value = ConvertirCocheBoolean(wsDTR.Range("E24").Value, wsDTR.Range("F24").Value)
    
End Sub

' =========================================================================
' PROCEDURE: REMPLIR MANUTENTION
' =========================================================================
Private Sub RemplirManutention(ByRef wsEMMST As Worksheet, ByVal ligne As Long, ByRef wsDTR As Worksheet)
' Colonnes 22-24 : Manutention Equipement, Outillage Spe, Precautions
    
    ' Colonne 22 : Manutention Equipement
    wsEMMST.Cells(ligne, 22).Value = ConvertirCocheBoolean(wsDTR.Range("E28").Value, wsDTR.Range("F28").Value)
    
    ' Colonne 23 : Outillage Specifique Manutention
    wsEMMST.Cells(ligne, 23).Value = ConvertirCocheBoolean(wsDTR.Range("E30").Value, wsDTR.Range("F30").Value)
    
    ' Colonne 24 : Precaution Manutention
    wsEMMST.Cells(ligne, 24).Value = ConvertirCocheBoolean(wsDTR.Range("E32").Value, wsDTR.Range("F32").Value)
    
End Sub
' =========================================================================
' PROCEDURE: REMPLIR STOCKAGE
' =========================================================================
Private Sub RemplirStockage(ByRef wsEMMST As Worksheet, ByVal ligne As Long, ByRef wsDTR As Worksheet)
' Colonnes 25-38 : Stockage (Position, Exterieur/Interieur, Conditions, etc.)
    
    ' Colonne 25 : Position Stockage
    wsEMMST.Cells(ligne, 25).Value = DeterminerPositionStockage(wsDTR)
    
    ' Colonnes 26-34 : Stockage Exterieur OU Interieur
    If wsDTR.Range("E36").Value = "X" Then
        ' Stockage EXTERIEUR
        Call RemplirStockageExterieur(wsEMMST, ligne, wsDTR)
    ElseIf wsDTR.Range("E38").Value = "X" Then
        ' Stockage INTERIEUR
        Call RemplirStockageInterieur(wsEMMST, ligne, wsDTR)
    End If
    
    ' Colonnes 35-38 : Gerbable, Moyen, Precautions, Maintenance
    wsEMMST.Cells(ligne, 35).Value = ConvertirCocheBoolean(wsDTR.Range("E42").Value, wsDTR.Range("F42").Value)
    
    wsEMMST.Cells(ligne, 36).Value = ConvertirCocheBoolean(wsDTR.Range("E44").Value, wsDTR.Range("F44").Value)
    
    wsEMMST.Cells(ligne, 37).Value = ConvertirCocheBoolean(wsDTR.Range("E46").Value, wsDTR.Range("F46").Value)
    
    wsEMMST.Cells(ligne, 38).Value = ConvertirCocheBoolean(wsDTR.Range("E48").Value, wsDTR.Range("F48").Value)
    
End Sub

' =========================================================================
' FONCTION: DETERMINER POSITION STOCKAGE
' =========================================================================
Private Function DeterminerPositionStockage(ByRef wsDTR As Worksheet) As String
' Colonne 25 : Retourne "Horizontal", "Vertical" ou "Indifferent"
    
    If wsDTR.Range("F34").Value = "X" Then
        DeterminerPositionStockage = "Horizontal"
    ElseIf wsDTR.Range("H34").Value = "X" Then
        DeterminerPositionStockage = "Vertical"
    ElseIf wsDTR.Range("J34").Value = "X" Then
        DeterminerPositionStockage = "Indifferent"
    Else
        DeterminerPositionStockage = ""
    End If
    
End Function

' =========================================================================
' PROCEDURE: REMPLIR STOCKAGE EXTERIEUR
' =========================================================================
Private Sub RemplirStockageExterieur(ByRef wsEMMST As Worksheet, ByVal ligne As Long, ByRef wsDTR As Worksheet)
' Colonnes 26-30 : Stockage Exterieur et sous-conditions
    
    With wsEMMST
        .Cells(ligne, 26).Value = -1        ' Exterieur = -1 (True)
        .Cells(ligne, 31).Value = 0         ' Interieur = 0 (False)
        
        ' Sous Abri
        .Cells(ligne, 27).Value = ConvertirCocheNumeric(wsDTR.Range("G36").Value)
        
        ' Air Libre
        .Cells(ligne, 28).Value = ConvertirCocheNumeric(wsDTR.Range("I36").Value)
        
        ' Hors Poussiere
        .Cells(ligne, 29).Value = ConvertirCocheNumeric(wsDTR.Range("K36").Value)
        
        ' Hors Eau
        .Cells(ligne, 30).Value = ConvertirCocheNumeric(wsDTR.Range("M36").Value)
    End With
    
End Sub

' =========================================================================
' PROCEDURE: REMPLIR STOCKAGE INTERIEUR
' =========================================================================
Private Sub RemplirStockageInterieur(ByRef wsEMMST As Worksheet, ByVal ligne As Long, ByRef wsDTR As Worksheet)
' Colonnes 26, 29-34 : Stockage Interieur et sous-conditions
    
    With wsEMMST
        .Cells(ligne, 26).Value = 0         ' Exterieur = 0 (False)
        .Cells(ligne, 31).Value = -1        ' Interieur = -1 (True)
        
        ' Chaud
        .Cells(ligne, 32).Value = ConvertirCocheNumeric(wsDTR.Range("G38").Value)
        
        ' Froid
        .Cells(ligne, 33).Value = ConvertirCocheNumeric(wsDTR.Range("I38").Value)
        
        ' Hors Poussiere
        .Cells(ligne, 29).Value = ConvertirCocheNumeric(wsDTR.Range("K38").Value)
        
        ' Hors Eau
        .Cells(ligne, 30).Value = ConvertirCocheNumeric(wsDTR.Range("M38").Value)
        
        ' Hors Lumiere
        .Cells(ligne, 34).Value = ConvertirCocheNumeric(wsDTR.Range("O38").Value)
    End With
    
End Sub

' =========================================================================
' PROCEDURE: REMPLIR TRANSPORT
' =========================================================================
Private Sub RemplirTransport(ByRef wsEMMST As Worksheet, ByVal ligne As Long, ByRef wsDTR As Worksheet)
' Colonnes 39-42 : Transport (Moyen, Outillage, Prescription, Precautions)
    
    ' Colonne 39 : Moyen Transport
    wsEMMST.Cells(ligne, 39).Value = ConvertirCocheBoolean(wsDTR.Range("E52").Value, wsDTR.Range("F52").Value)
    
    ' Colonne 40 : Outillage Specifique Transport
    wsEMMST.Cells(ligne, 40).Value = ConvertirCocheBoolean(wsDTR.Range("E54").Value, wsDTR.Range("F54").Value)
    
    ' Colonne 41 : Prescription Transport
    wsEMMST.Cells(ligne, 41).Value = ConvertirCocheBoolean(wsDTR.Range("E56").Value, wsDTR.Range("F56").Value)
    
    ' Colonne 42 : Precaution Transport
    wsEMMST.Cells(ligne, 42).Value = ConvertirCocheBoolean(wsDTR.Range("E58").Value, wsDTR.Range("F58").Value)
    
End Sub

' =========================================================================
' PROCEDURE: REMPLIR VERSION
' =========================================================================
Private Sub RemplirVersion(ByRef wsEMMST As Worksheet, ByVal ligne As Long, ByRef wsDTR As Worksheet, ByVal langue As String)
' Colonnes 43-45 : Version, Date Peremption EN, Unite EN
    
    With wsEMMST
        .Cells(ligne, 43).Value = ""                                ' Version (vide)
        .Cells(ligne, 44).Value = wsDTR.Range("J14").Value          ' Date Peremption EN
        .Cells(ligne, 45).Value = wsDTR.Range("E16").Value          ' Unite Conditionnement EN
    End With
    
End Sub

' =========================================================================
' FONCTION: CONVERTIR COCHE BOOLEAN
' =========================================================================
Private Function ConvertirCocheBoolean(ByVal cocheNon As Variant, ByVal cocheOui As Variant) As Variant
' Convertit 2 coches (Non/Oui) en valeur BDD : 0 = Non, -1 = Oui, "" = Vide
    
    If cocheNon = "X" Then
        ConvertirCocheBoolean = 0           ' Non
    ElseIf cocheOui = "X" Then
        ConvertirCocheBoolean = -1          ' Oui
    Else
        ConvertirCocheBoolean = ""          ' Vide
    End If
    
End Function

' =========================================================================
' FONCTION: CONVERTIR COCHE NUMERIC
' =========================================================================
Private Function ConvertirCocheNumeric(ByVal coche As Variant) As Integer
' Convertit une coche en valeur BDD : -1 si "X", 0 sinon
    
    If coche = "X" Then
        ConvertirCocheNumeric = -1
    Else
        ConvertirCocheNumeric = 0
    End If
    
End Function

' =========================================================================
' PROCEDURE: FORMATER FEUILLE EXTRACTION EMMST
' =========================================================================
Private Sub FormaterFeuilleExtractionEMMST(ByRef wb As Workbook)
' Desactive le wrap text sur la feuille EMMST
    
    Dim wsEMMST As Worksheet
    Dim derniereLigne As Long
    
    Set wsEMMST = wb.Sheets("EMMST")
    derniereLigne = wsEMMST.Range("A100000").End(xlUp).Row
    
    ' Desactiver wrap text
    wsEMMST.Activate
    wsEMMST.Range("A1:AY" & derniereLigne).Select
    With Selection
        .WrapText = False
    End With
    
    ' Retourner sur A1
    wsEMMST.Range("A1").Select
    
End Sub


