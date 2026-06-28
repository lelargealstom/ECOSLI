Attribute VB_Name = "D_Providor_TNR"
'--------------------------------------------------------------------------------------------------------
'
'                                       Generateur Livrables SLI
'                                       Module D_Providor_TNR (Version Refactorisee)
'                                       Developpement : LE LARGE Thomas (LGM/Alstom)
'                                       Module contenant la PBS etalon pour test de non-regression Providor
'                                       Refactorisation : Fevrier 2026
'
'-----------------------------------------------------------------------------------------------------------
Option Explicit

' =========================================================================
' STRUCTURE ETALON
' =========================================================================
Public Type T_PBS_Etalon
    dtr          As String
    Designation  As String
    SS           As String
    Qte          As Long
    MTBFIntr     As Double
    MTBFFct      As Double
    DVLHARD      As Double
    DVLSOFT      As Double
End Type

Public PBS_Etalon()          As T_PBS_Etalon
Public Const NB_LIGNES_ETALON As Long = 48

' =========================================================================
' INITIALISATION COMPLETE
' =========================================================================
Public Sub InitPBS_Etalon()
    ReDim PBS_Etalon(1 To NB_LIGNES_ETALON) 'base 1
    Call Init_DTR
    Call Init_Designation
    Call Init_SS
    Call Init_Qte
    ' MTBFIntr et MTBFFct vides pour l'instant ? valeurs a 0 par defaut
    ' Call _Init_MTBFIntr
    ' Call _Init_MTBFFct
    ' Call _Init_DVLHARD
    ' Call _Init_DVLSOFT
End Sub

Private Sub Init_DTR()
    Dim v As Variant
    Dim i As Long

    v = Array( "DTR0000120004", "DTR0000221263", "DTR0000228338", "DTR0000253129", "DTR0000299092", "DTR0000382566", "DTR0000382822", "DTR0000421031", "DTR0000423752", "DTR0000436717", "DTR0000436719", "DTR0000436720", "DTR0000445292", "DTR0000459726", "DTR0000460596", "DTR0000460597", "DTR0000460598", "DTR0000460600", "DTR0000464916", "DTR0000494988", "DTR0000499404", "DTR0000520477", "DTR0000540751", "DTR0000560752", "DTR0000560755", "DTR0000560758", "DTR0000564075", "DTR0000564077", "DTR0000565035", "DTR0000569001", "DTR0000569002", "DTR0000569003", "DTR0000581177", "DTR0000583207", "DTR0000587713", "DTR0000593059", "DTR0000593207", "DTR2000000973", "DTR2000014281", "DTR2000014282", "DTR2000014283", "DTR2000015440", "DTR2000021122", "DTR2000029651", "DTR2000038851", "DTR2000038852", "DTR2000056364", "DTR2000061970")

    ' v est base 0 (0 To 47), PBS_Etalon est base 1 (1 To 48)
    ' ? decalage de +1 sur l'index
    For i = 1 To NB_LIGNES_ETALON
        PBS_Etalon(i).dtr = v(i - 1)
    Next i
End Sub

' =========================================================================
' DESIGNATION
' =========================================================================
Private Sub Init_Designation()
    Dim v1 As Variant, v2 As Variant, v3 As Variant
    Dim v() As String
    Dim i As Long

    v1 = Array( "DATAPLUG DMI", "100MPS FIXED ETHERNET REPEATER", "DATAPLUG HC", "ANTENNA SENCITY RAIL", "POWER OVER ETHERNET POE 6 PORTS", "GIGABIT ETHERNET SWITCH", "GIGABIT ETHERNET SWITCH POE HV", "GES DATAPLUG", "TRAIN ROUTER SWITCH V1.1", "NETBOX V2 BASE / NETBOX (WIRELESS)", "NETBOX V2 MPU / XPU", "NETBOX V2 MPU 32GB / XPU 32GB", "CARTE CPU G40A COMPACT PCI SERIAL", "BRIO XT R201 SIL2", "BRIO XT R101", "BRIO XT R102")

    v2 = Array( "BRIO XT R104", "BRIO XT R105", "FMI CARD TECHNICAL SPECIFICATION", "ANTENA SWA-0759/360/6/0/MIMO", "RAILWAY ETHERNET SWITCH", "13 PORT SENCITY RAIL MULTI ANTENNA", "SECURITY GATEWAY SGTW", "DMIACE2_121_XGA_TK32", "DMIACE2_104_XGA_T", "DMIACE2_104_XGA_TK32", "PSU BOARD OF NETBOX_V2", "POWER & I/O BOARD FOR NETBOX V2", "NETBOX V2 EVO (WIRELESS)", "IRMA6 POE WITHOUT I/O CONNECTOR", "IRMA6 ETH WITHOUT I/O CONNECTOR", "IRMA6 R2 PASSENGERS COUNTING SENSOR")

    v3 = Array( "IP-TS-8GBE 8PORT ETBN", "GIGABIT ETHERNET SWITCH POE HV", "MSATA BOARD FOR NETBOX V2", "SENCITY RAIL MIMO ANTENNA 5 PORT", "ECP-USB", "CAPACITOR CABLING ADAPTER KIT", "STPU EQUIPPED BLOCK", "IMLI EQUIPPED BLOCK", "TSA EQUIPPED BLOCK", "AC4 EQUIPPED FAN RACK", "CRSv2 evo", "AC4L-01-L0L00", "BDOV EQUIPPED BLOCK", "BDI EQUIPPED BLOCK", "AC4-05-262200 CONFIGURED RACK", "Gigabit Cut Shield Cabling Adapter")

    ReDim v(1 To NB_LIGNES_ETALON)
    For i = 1 To 16: v(i) = v1(i - 1): Next i
    For i = 1 To 16: v(i + 16) = v2(i - 1): Next i
    For i = 1 To 16: v(i + 32) = v3(i - 1): Next i

    For i = 1 To NB_LIGNES_ETALON
        PBS_Etalon(i).Designation = v(i)
    Next i
End Sub

Private Sub Init_SS()
    Dim TCMS As String, TRAC As String, PACI As String
    TCMS = "TCMS - T2G - OMS"
    TRAC = "Traction - BMS"
    PACI = "PACIS - CCTV - Rearview"

    Dim v(1 To NB_LIGNES_ETALON) As String
    Dim i As Long

    For i = 1 To 20: v(i) = TCMS: Next i
    For i = 21 To 29: v(i) = TRAC: Next i
    For i = 30 To 48: v(i) = PACI: Next i

    For i = 1 To NB_LIGNES_ETALON
        PBS_Etalon(i).SS = v(i)
    Next i
End Sub

' =========================================================================
' QUANTITE (toutes a 1)
' =========================================================================
Private Sub Init_Qte()
    Dim i As Long
    For i = 1 To NB_LIGNES_ETALON
        PBS_Etalon(i).Qte = 1
    Next i
End Sub

' =========================================================================
' MTBF INTRINSEQUE  ? a completer
' =========================================================================
Private Sub Init_MTBFIntr()
    Dim i As Long
    For i = 1 To NB_LIGNES_ETALON
        PBS_Etalon(i).MTBFIntr = 0#
    Next i
End Sub

' =========================================================================
' MTBF FONCTIONNEL  ? a completer
' =========================================================================
Private Sub Init_MTBFFct()
    Dim i As Long
    For i = 1 To NB_LIGNES_ETALON
        PBS_Etalon(i).MTBFFct = 0#
    Next i
End Sub

' =========================================================================
' DVL HARDWARE  ? a completer
' =========================================================================
Private Sub Init_DVLHARD()
    Dim i As Long
    For i = 1 To NB_LIGNES_ETALON
        PBS_Etalon(i).DVLHARD = 0#
    Next i
End Sub

' =========================================================================
' DVL SOFTWARE  ? a completer
' =========================================================================
Private Sub Init_DVLSOFT()
    Dim i As Long
    For i = 1 To NB_LIGNES_ETALON
        PBS_Etalon(i).DVLSOFT = 0#
    Next i
End Sub

' =========================================================================
' POINT D'ENTREE : TEST DE NON-REGRESSION PROVIDOR
' =========================================================================
Public Sub RunTest_NonRegression_Providor()
    'On initialise la PBS Etalon
    Call InitPBS_Etalon
    
    'On injecte les donnees de l'etalon dans les tableaux du code Providor
    Call InjecterDonneesEtalon
    
    MsgBox "Donnees etalon injectees (" & NB_LIGNES_ETALON & " lignes)." & vbCrLf & "Pret pour le test de generation Providor.", vbInformation, "Test Non-Regression"
           
    Call Main_Providor(True)
    
End Sub

' =========================================================================
' INJECTION dans les tableaux globaux de G_Providor_Main
' =========================================================================
Public Sub InjecterDonneesEtalon()
    Dim n As Long
    n = NB_LIGNES_ETALON

    ReDim tDTR(1 To n, 1 To 1)
    ReDim tDES_PBS(1 To n, 1 To 1)
    ReDim tSS(1 To n, 1 To 1)
    ReDim tQTE(1 To n, 1 To 1)
    ReDim tMTBF_Intr(1 To n, 1 To 1)
    ReDim tMTBF_fct(1 To n, 1 To 1)
    ReDim tDureevie_Ana_HARD(1 To n, 1 To 1)
    ReDim tDureevie_Ana_SOFT(1 To n, 1 To 1)

    Dim i As Long
    For i = 1 To n
        tDTR(i, 1) = PBS_Etalon(i).dtr
        tDES_PBS(i, 1) = PBS_Etalon(i).Designation
        tSS(i, 1) = PBS_Etalon(i).SS
        tQTE(i, 1) = PBS_Etalon(i).Qte
        tMTBF_Intr(i, 1) = PBS_Etalon(i).MTBFIntr
        tMTBF_fct(i, 1) = PBS_Etalon(i).MTBFFct
        tDureevie_Ana_HARD(i, 1) = PBS_Etalon(i).DVLHARD
        tDureevie_Ana_SOFT(i, 1) = PBS_Etalon(i).DVLSOFT
    Next i
End Sub

' =========================================================================
' UTILITAIRES
' =========================================================================
Public Function PBS_Etalon_Count() As Long
    On Error Resume Next
    PBS_Etalon_Count = UBound(PBS_Etalon) + 1
    On Error GoTo 0
End Function

Public Sub Debug_AfficherEtalon()
    Dim i As Long
    Dim n As Long
    n = PBS_Etalon_Count
    If n = 0 Then
        Debug.Print "PBS_Etalon non initialise - appeler InitPBS_Etalon() d'abord"
        Exit Sub
    End If
    Debug.Print String(80, "=")
    Debug.Print "PBS Etalon - " & n & " lignes"
    Debug.Print String(80, "=")
    For i = 1 To n
        Debug.Print Format(i, "00") & " | " & PBS_Etalon(i).dtr & " | " & Left(PBS_Etalon(i).Designation, 35) & " | " & Left(PBS_Etalon(i).SS, 25) & " | Qte=" & PBS_Etalon(i).Qte & " | MTBFi=" & PBS_Etalon(i).MTBFIntr & " | MTBFf=" & PBS_Etalon(i).MTBFFct
    Next i
End Sub
Public Sub CollecterMetriquesTest(ByRef wb As Workbook)

    On Error Resume Next

    With ResultatTest
        .Succes = True

        ' --- Nombre de lignes par feuille (a partir de la ligne donnees) ---
        .NbLignes_BS = CompterLignesFeuille(wb.Worksheets("Breakdown structure"), 3, 5)
        .NbLignes_ML = CompterLignesFeuille(wb.Worksheets("Material list"), 3, 4)
        .NbLignes_CM = CompterLignesFeuille(wb.Worksheets("Reliability & Corr. maintenance"), 1, 4)
        .NbLignes_PM = CompterLignesFeuille(wb.Worksheets("Preventive maintenance"), 4, 4)
        .NbLignes_SL = CompterLignesFeuille(wb.Worksheets("Skill List"), 3, 4)
        .NbLignes_TL = CompterLignesFeuille(wb.Worksheets("Tool List"), 3, 4)

        ' --- LCC Final : feuille LCC Calculation, cellule P18 ---
        Dim lccVal As Variant
        lccVal = wb.Worksheets("LCC Calculation").Range("P18").Value

        If IsNumeric(lccVal) And Not IsEmpty(lccVal) Then
            .LCC_Final = lccVal
        Else
            .LCC_Final = "N/D"
            .MessageErreur = "LCC P18 non numerique ou vide"
        End If
    End With

    On Error GoTo 0
End Sub
' =========================================================================
' FONCTION: COMPTER LIGNES FEUILLE
' Retourne le nombre de lignes de donnees a partir de ligneDebut
' en se basant sur la colonne colRef
' =========================================================================
Public Function CompterLignesFeuille(ByRef ws As Worksheet, ByVal colRef As Long, ByVal ligneDebut As Long) As Long
    Dim derniereLigne As Long
    On Error Resume Next
    derniereLigne = ws.Cells(ws.rows.count, colRef).End(xlUp).Row
    If derniereLigne < ligneDebut Then
        CompterLignesFeuille = 0
    Else
        CompterLignesFeuille = derniereLigne - ligneDebut + 1
    End If
    On Error GoTo 0
End Function
Public Sub AfficherResultatsTest()

    Dim msg As String
    Dim lccStr As String

    If IsNumeric(ResultatTest.LCC_Final) Then
        lccStr = Format(CDbl(ResultatTest.LCC_Final), "#,##0.00 ?")
    Else
        lccStr = CStr(ResultatTest.LCC_Final)
    End If

    msg = "=== RESULTATS TEST NON-REGRESSION PROVIDOR ===" & vbCrLf & vbCrLf & "Projet etalon : " & Code_Projet & vbCrLf & "Gate          : " & gate & vbCrLf & vbCrLf & "--- Lignes generees par feuille ---" & vbCrLf & "Breakdown Structure    : " & ResultatTest.NbLignes_BS & vbCrLf & "Material List          : " & ResultatTest.NbLignes_ML & vbCrLf & "Corrective Maintenance : " & ResultatTest.NbLignes_CM & vbCrLf & "Preventive Maintenance : " & ResultatTest.NbLignes_PM & vbCrLf & "Skill List             : " & ResultatTest.NbLignes_SL & vbCrLf & "Tool List              : " & ResultatTest.NbLignes_TL & vbCrLf & vbCrLf & "--- LCC Final (P18) ---" & vbCrLf & lccStr

    If ResultatTest.MessageErreur <> "" Then
        msg = msg & vbCrLf & vbCrLf & "? Avertissement : " & ResultatTest.MessageErreur
    End If

    MsgBox msg, vbInformation, "Test Non-Regression Providor"

End Sub
' =========================================================================
' PROCEDURE: NETTOYER FICHIER TEST
' Ferme sans sauvegarder et supprime le fichier temporaire
' =========================================================================
Public Sub NettoyerFichierTest(ByRef wb As Workbook)

    Dim cheminFichier As String

    On Error Resume Next

    cheminFichier = wb.FullName

    ' Fermer sans sauvegarder
    wb.Close SaveChanges:=False

    ' Supprimer le fichier temporaire
    If Dir(cheminFichier) <> "" Then
        Kill cheminFichier
    End If

    On Error GoTo 0

End Sub


