Option Explicit

Sub ExportAllModules()
    Dim vbComp As Object
    Dim exportPath, userpath As String
    
    userpath = Environ$("UserProfile")
    
    ' Dossier d'export (a adapter)
    exportPath = userpath & "\Alstom\Svcs_DIS_Fr_VBN - Documents\01_Metiers\01_ISR\007 - Outils metier ISR\01. BDD SLI\01_Outils MACRO EXCEL\VBA_Export\"
    
    ' Cree le dossier s'il n'existe pas
    If Dir(exportPath, vbDirectory) = "" Then MkDir exportPath
    
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        
        Select Case vbComp.Type
            Case 1 ' Module standard (.bas)
                vbComp.Export exportPath & vbComp.Name & ".bas"
                
            Case 2 ' Classe (.cls)
                vbComp.Export exportPath & vbComp.Name & ".cls"
                
            Case 3 ' UserForm (.frm)
                vbComp.Export exportPath & vbComp.Name & ".frm"
        End Select
        
    Next vbComp
    
    MsgBox "Export termine !"
End Sub

' --- Helpers ---
Private Function IsNumberKnown(v As Variant) As Boolean
    If IsError(v) Then Exit Function
    If IsMissing(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    If Trim$(CStr(v)) = "" Then Exit Function
    If Not IsNumeric(v) Then Exit Function
    IsNumberKnown = True
End Function

Private Function BuildKnownArray(v1 As Variant, v2 As Variant, v3 As Variant, ByRef idxMap() As Long) As Double()
    Dim tmp(1 To 3) As Double, idx(1 To 3) As Long, i As Long, n As Long
    n = 0
    If IsNumberKnown(v1) Then n = n + 1: tmp(n) = CDbl(v1): idx(n) = 1
    If IsNumberKnown(v2) Then n = n + 1: tmp(n) = CDbl(v2): idx(n) = 2
    If IsNumberKnown(v3) Then n = n + 1: tmp(n) = CDbl(v3): idx(n) = 3

    Dim arr() As Double
    If n > 0 Then
        ReDim arr(1 To n)
        ReDim idxMap(1 To n)
        For i = 1 To n
            arr(i) = tmp(i)
            idxMap(i) = idx(i)
        Next
    Else
        ReDim arr(1 To 1)
        ReDim idxMap(1 To 1)
        arr(1) = 0#
        idxMap(1) = 0
    End If
    BuildKnownArray = arr
End Function

Private Function MedianOf(arr() As Double) As Double
    Dim n As Long, temp() As Double, i As Long, j As Long, t As Double
    n = UBound(arr) - LBound(arr) + 1
    ReDim temp(1 To n)
    For i = 1 To n: temp(i) = arr(i): Next

    ' Tri simple (n = 3 ici)
    For i = 1 To n - 1
        For j = i + 1 To n
            If temp(j) < temp(i) Then t = temp(i): temp(i) = temp(j): temp(j) = t
        Next j
    Next i

    If n Mod 2 = 1 Then
        MedianOf = temp((n + 1) \ 2)
    Else
        MedianOf = (temp(n \ 2) + temp(n \ 2 + 1)) / 2
    End If
End Function

Private Function ColumnName(i As Long) As String
    Select Case i
        Case 1: ColumnName = "Prix Manuel Spare"
        Case 2: ColumnName = "Cout spare ALSTOM FRANCE FOURNISSEUR"
        Case 3: ColumnName = "ISP TIS FRANCE (FCS)"
        Case Else: ColumnName = ""
    End Select
End Function

' --- Fonctions publiques ---

' Nombre de prix renseignes (0 a 3)
Public Function Prix_NombreConnus(Manual As Variant, Alstom As Variant, ISP As Variant) As Long
    Dim idx() As Long, vals() As Double
    vals = BuildKnownArray(Manual, Alstom, ISP, idx)
    If idx(1) = 0 Then
        Prix_NombreConnus = 0
    Else
        Prix_NombreConnus = UBound(vals)
    End If
End Function

' Moyenne uniquement des prix connus (NA si aucun)
Public Function Prix_MoyenneConnus(Manual As Variant, Alstom As Variant, ISP As Variant) As Variant
    Dim idx() As Long, vals() As Double, n As Long, i As Long, s As Double
    vals = BuildKnownArray(Manual, Alstom, ISP, idx)
    If idx(1) = 0 Then
        Prix_MoyenneConnus = CVErr(xlErrNA)
        Exit Function
    End If
    n = UBound(vals)
    For i = 1 To n: s = s + vals(i): Next
    Prix_MoyenneConnus = s / n
End Function

' Vrai si au moins un prix s'ecarte de plus de "seuil" (25% par defaut) de la mediane
Public Function Prix_DifferenceSuperieureSeuil(Manual As Variant, Alstom As Variant, ISP As Variant, Optional seuil As Double = 0.25) As Boolean
    Dim idx() As Long, vals() As Double, n As Long, med As Double, i As Long, diff As Double
    vals = BuildKnownArray(Manual, Alstom, ISP, idx)
    If idx(1) = 0 Or UBound(vals) < 2 Then
        Prix_DifferenceSuperieureSeuil = False
        Exit Function
    End If

    med = MedianOf(vals)
    If med = 0 Then
        ' Si mediane = 0, tout prix non nul est considere "tres different"
        For i = 1 To UBound(vals)
            If vals(i) <> 0 Then Prix_DifferenceSuperieureSeuil = True: Exit Function
        Next
        Prix_DifferenceSuperieureSeuil = False
        Exit Function
    End If

    For i = 1 To UBound(vals)
        diff = Abs(vals(i) - med) / Abs(med)
        If diff > seuil Then Prix_DifferenceSuperieureSeuil = True: Exit Function
    Next
    Prix_DifferenceSuperieureSeuil = False
End Function

' Renvoie le NOM de la colonne la plus differente (> seuil) ou vide si rien a signaler
Public Function Prix_ColonneDifferente(Manual As Variant, Alstom As Variant, ISP As Variant, Optional seuil As Double = 0.25) As String
    Dim idx() As Long, vals() As Double, n As Long, med As Double, i As Long
    Dim diff As Double, bestIdx As Long, bestDiff As Double

    vals = BuildKnownArray(Manual, Alstom, ISP, idx)
    If idx(1) = 0 Or UBound(vals) < 2 Then
        Prix_ColonneDifferente = ""
        Exit Function
    End If

    med = MedianOf(vals)
    bestDiff = -1#
    bestIdx = 0

    If med = 0 Then
        For i = 1 To UBound(vals)
            diff = IIf(vals(i) = 0, 0, 1E+99)
            If diff > bestDiff Then bestDiff = diff: bestIdx = idx(i)
        Next
        If bestDiff > seuil Then
            Prix_ColonneDifferente = ColumnName(bestIdx)
        Else
            Prix_ColonneDifferente = ""
        End If
        Exit Function
    End If

    For i = 1 To UBound(vals)
        diff = Abs(vals(i) - med) / Abs(med)
        If diff > bestDiff Then
            bestDiff = diff
            bestIdx = idx(i)
        End If
    Next

    If bestDiff > seuil Then
        Prix_ColonneDifferente = ColumnName(bestIdx)
    Else
        Prix_ColonneDifferente = ""
    End If
End Function

' (Optionnel) Resume texte lisible
Public Function Prix_Resume(Manual As Variant, Alstom As Variant, ISP As Variant, Optional seuil As Double = 0.25) As String
    Dim nb As Long, avg As Variant, flag As Boolean, col As String, avgText As String
    nb = Prix_NombreConnus(Manual, Alstom, ISP)
    avg = Prix_MoyenneConnus(Manual, Alstom, ISP)
    flag = Prix_DifferenceSuperieureSeuil(Manual, Alstom, ISP, seuil)
    col = Prix_ColonneDifferente(Manual, Alstom, ISP, seuil)

    If IsError(avg) Then avgText = "NA" Else avgText = Format(avg, "0.00")
    Prix_Resume = "Connus=" & nb & "/3 ; Moyenne=" & avgText & " ; Ecart > " & Replace(Format(seuil, "0.0%"), ",0", "") & " ? " & IIf(flag, "Oui (" & col & ")", "Non")
End Function


