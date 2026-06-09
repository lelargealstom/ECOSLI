Attribute VB_Name = "G_SIRUS_ATMAP"
Option Explicit

' ============================================================
' CONSTANTES
' ============================================================
Private Const SITE_URL         As String = "https://alstomgroup.sharepoint.com/sites/Svcs_DIS_Fr_VBN"
Private Const LIST_NAME        As String = "ATMAP_TICKETS"
Private Const LIST_ENTITY_TYPE As String = "SP.Data.ATMAP_x005f_TICKETSListItem"

' ============================================================
' POINT D'ENTR�E � TEST
' ============================================================

Sub Test_Creer_Ticket_ATMAP()

    Dim digest   As String
    Dim userID   As Long
    Dim newID    As Long
    Dim result   As String
    Dim pjResult As String

    ' -- �tape 1 � Digest CSRF ------------------------------
    digest = GetFormDigest()
    If digest = "" Then
        MsgBox "�chec r�cup�ration digest.", vbCritical, "ATMAP Test"
        Exit Sub
    End If

    ' -- �tape 2 � Utilisateur connect� ---------------------
    userID = GetCurrentUserID(digest)
    If userID = -1 Then
        MsgBox "Impossible de r�cup�rer l'utilisateur connect�.", vbExclamation, "ATMAP Test"
        Exit Sub
    End If

    ' -- �tape 3 � Donn�es d'import -------------------------
    Dim lignesImport As New Collection
    lignesImport.Add Array( _
        "DTR-001-A", "Calculateur traction", "LRU", "Bombardier", _
        "BC-2241-TRC", "DRW-4412", "45000", "15", "10", "EUR")
    lignesImport.Add Array( _
        "DTR-002-B", "Carte alimentation 24V", "SRU", "Alstom", _
        "ALT-PWR-024", "DRW-0981", "22000", "12", "", "EUR")

    ' -- �tape 4 � Cr�ation ticket --------------------------
    result = CreateTicketREST( _
        digest:=digest, _
        userID:=userID, _
        outilID:=43, _
        numeroTicket:="VBN_Test", _
        sujet:="IMP_001 � Import Produit", _
        statut:="Nouveau", _
        description:="Import de " & lignesImport.count & _
                     " DTR en table Produit � DTR-001-A, DTR-002-B", _
        attentes:="Int�gration en BdD SLI � Table Produit", _
        typedemande:="Import", _
        priorite:="High", _
        ticketIDOut:=newID _
    )

    MsgBox result, vbInformation, "ATMAP Test � R�sultat"

    If newID = -1 Then Exit Sub

    ' -- �tape 5 � Pi�ce jointe Excel -----------------------
    pjResult = AttachExcelToTicket( _
        ticketID:=newID, _
        digest:=digest, _
        tableName:="Produit", _
        numeroTicket:="VBN_Test", _
        dataRows:=lignesImport _
    )

    MsgBox pjResult, vbInformation, "ATMAP Test � Pi�ce jointe"

End Sub

' ============================================================
' CR�ATION TICKET COMPLET � UNE SEULE REQU�TE REST POST
' ============================================================

Private Function CreateTicketREST(digest As String, _
                                   userID As Long, _
                                   outilID As Long, _
                                   numeroTicket As String, _
                                   sujet As String, _
                                   statut As String, _
                                   description As String, _
                                   attentes As String, _
                                   typedemande As String, _
                                   priorite As String, _
                                   ticketIDOut As Long) As String

    Dim xhr   As Object
    Dim url   As String
    Dim body  As String
    Dim resp  As String
    Dim today As String

    ticketIDOut = -1

    ' MSXML2 valid� � ne pas changer pour WinHttp ici
    Set xhr = CreateObject("MSXML2.XMLHTTP.6.0")

    url = SITE_URL & "/_api/web/lists/getbytitle('" & LIST_NAME & "')/items"
    today = Format(Now(), "yyyy-mm-dd") & "T00:00:00Z"

    ' Pas de Title � champ non utilis� dans ATMAP
    body = "{""__metadata"":{""type"":""" & LIST_ENTITY_TYPE & """}," & _
           """Numero_ticket"":""" & EscJson(numeroTicket) & """," & _
           """Sujet"":""" & EscJson(sujet) & """," & _
           """Statut"":""" & EscJson(statut) & """," & _
           """Description"":""" & EscJson(description) & """," & _
           """Attentes"":""" & EscJson(attentes) & """," & _
           """Typedemande"":""" & EscJson(typedemande) & """," & _
           """Priorit_x00e9_"":""" & EscJson(priorite) & """," & _
           """Datecr_x00e9_ation"":""" & today & """," & _
           """DemandeurId"":" & userID & "," & _
           """Assign_x00e9__x00e0_Id"":" & userID & "," & _
           """NomoutilId"":" & outilID & _
           "}"

    On Error GoTo ErrCreate

    xhr.Open "POST", url, False
    xhr.setRequestHeader "Accept", "application/json;odata=verbose"
    xhr.setRequestHeader "Content-Type", "application/json;odata=verbose"
    xhr.setRequestHeader "X-RequestDigest", digest
    xhr.Send body

    resp = xhr.responseText

    If xhr.Status = 201 Then
        ticketIDOut = CLng(ParseJsonLong(resp, """Id"":"))
        CreateTicketREST = "Ticket cr�� � ID : " & ticketIDOut & vbCrLf & _
                           "Num�ro      : " & numeroTicket & vbCrLf & _
                           "DemandeurId : " & userID & vbCrLf & _
                           "NomoutilId  : " & outilID
    Else
        CreateTicketREST = "Erreur HTTP " & xhr.Status & vbCrLf & Left(resp, 500)
    End If
    Exit Function

ErrCreate:
    CreateTicketREST = "Erreur : " & Err.Number & " � " & Err.description
    MsgBox CreateTicketREST, vbCritical, "ATMAP � CreateTicketREST"

End Function

' ============================================================
' CR�ER ET ATTACHER UN FICHIER EXCEL AU TICKET
' ============================================================

Private Function AttachExcelToTicket(ticketID As Long, _
                                      digest As String, _
                                      tableName As String, _
                                      numeroTicket As String, _
                                      dataRows As Collection) As String

    Dim xlApp      As Object
    Dim xlWb       As Object
    Dim xlWs       As Object
    Dim xlPath     As String
    Dim xlName     As String
    Dim fileData   As Variant
    Dim xhr        As Object
    Dim headers    As Variant
    Dim col        As Long
    Dim rowIdx     As Long
    Dim uploadStatus As Long
    Dim stream As Object
    
    

    AttachExcelToTicket = ""

    ' -- 1. Chemin temporaire --------------------------------
    xlName = "SIRUS_" & tableName & "_" & Format(Now(), "yyyymmdd_hhmmss") & ".xlsx"
    xlPath = Environ("TEMP") & "\" & xlName

    On Error GoTo ErrExcel

    ' -- 2. Cr�er le classeur --------------------------------
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    xlApp.DisplayAlerts = False

    Set xlWb = xlApp.Workbooks.Add
    Set xlWs = xlWb.Sheets(1)
    xlWs.Name = tableName

    headers = GetHeadersForTable(tableName)

    ' -- 3. Ligne titre fusionn�e ----------------------------
    xlWs.rows(1).RowHeight = 28
    With xlWs.Range(xlWs.Cells(1, 1), xlWs.Cells(1, UBound(headers) + 1))
        .Merge
        .Value = "SIRUS � Import � " & tableName & _
                 "  |  " & numeroTicket & _
                 "  |  " & Format(Now(), "dd/mm/yyyy hh:mm")
        .Font.Bold = True
        .Font.Size = 10
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 56, 100)
        .HorizontalAlignment = -4108
        .VerticalAlignment = -4108
    End With

    ' -- 4. En-t�tes colonnes --------------------------------
    xlWs.rows(2).RowHeight = 20
    For col = 1 To UBound(headers) + 1
        With xlWs.Cells(2, col)
            .Value = headers(col - 1)
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(46, 117, 182)
            .Borders.LineStyle = 1
            .Borders.Color = RGB(180, 180, 180)
            .HorizontalAlignment = -4108
        End With
    Next col

    ' -- 5. Donn�es ------------------------------------------
    rowIdx = 3
    If Not dataRows Is Nothing Then
        Dim rowData As Variant
        For Each rowData In dataRows
            If IsArray(rowData) Then
                xlWs.rows(rowIdx).RowHeight = 18
                For col = 1 To UBound(rowData) + 1
                    With xlWs.Cells(rowIdx, col)
                        .Value = rowData(col - 1)
                        .Borders.LineStyle = 1
                        .Borders.Color = RGB(180, 180, 180)
                        If rowIdx Mod 2 = 0 Then
                            .Interior.Color = RGB(214, 228, 241)
                        End If
                    End With
                Next col
                rowIdx = rowIdx + 1
            End If
        Next rowData
    End If

    ' -- 6. Tableau structur� --------------------------------
    If rowIdx > 3 Then
        Dim lo As Object
        Set lo = xlWs.ListObjects.Add(1, _
            xlWs.Range(xlWs.Cells(2, 1), _
                       xlWs.Cells(rowIdx - 1, UBound(headers) + 1)), , 1)
        lo.Name = "tbl_Import_" & tableName
        lo.TableStyle = "TableStyleMedium2"
    End If

    ' -- 7. Figer titres + en-t�tes --------------------------
    xlWs.Cells(3, 1).Select
    xlApp.ActiveWindow.FreezePanes = True
    xlWs.Columns.AutoFit

    ' -- 8. Sauvegarder et fermer ----------------------------
    xlWb.SaveAs Filename:=xlPath, FileFormat:=51
    xlWb.Close False
    xlApp.Quit
    Set xlWs = Nothing: Set xlWb = Nothing: Set xlApp = Nothing

    ' -- 9. Lire en binaire ----------------------------------
    fileData = ReadFileBinary(xlPath)
    If IsEmpty(fileData) Then
        AttachExcelToTicket = "Erreur lecture fichier Excel."
        Exit Function
    End If
    
    ' -- 10. Uploader via REST � WinHttp pour les PJ ---------
    ' WinHttp obligatoire ici : SetAutoLogonPolicy + envoi binaire Byte()
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1  ' adTypeBinary
    stream.Open
    stream.Write fileData
    stream.Position = 0
    
    Set xhr = CreateObject("MSXML2.XMLHTTP.6.0")
    xhr.Open "POST", SITE_URL & "/_api/web/lists/getbytitle('" & LIST_NAME & "')" & _
             "/items(" & ticketID & ")/AttachmentFiles/add(FileName='" & xlName & "')", False
    xhr.setRequestHeader "Accept", "application/json;odata=verbose"
    xhr.setRequestHeader "Content-Type", "application/octet-stream"
    xhr.setRequestHeader "X-RequestDigest", digest
    xhr.Send stream

    uploadStatus = xhr.Status

    stream.Close
    Set stream = Nothing

    ' -- 11. Supprimer le temporaire -------------------------
    On Error Resume Next
    Kill xlPath
    On Error GoTo 0

    If uploadStatus = 200 Then
        AttachExcelToTicket = "PJ Excel OK � " & xlName & vbCrLf & _
                              dataRows.count & " ligne(s) export�e(s)"
    Else
        AttachExcelToTicket = "PJ HTTP " & uploadStatus & vbCrLf & _
                              Left(xhr.responseText, 300)
    End If
    Exit Function

ErrExcel:
    AttachExcelToTicket = "Erreur Excel : " & Err.Number & " � " & Err.description
    MsgBox AttachExcelToTicket, vbCritical, "ATMAP � AttachExcelToTicket"
    On Error Resume Next
    If Not xlWb Is Nothing Then xlWb.Close False
    If Not xlApp Is Nothing Then xlApp.Quit
    On Error GoTo 0

End Function

' ============================================================
' EN-T�TES PAR TABLE
' ============================================================

Private Function GetHeadersForTable(tableName As String) As Variant
    Select Case UCase(tableName)
        Case "PRODUIT"
            GetHeadersForTable = Array( _
                "DTR", "D�signation", "Node Type", "Fournisseur", _
                "R�f. Fournisseur", "Drawing Number", "MTBF Intrins�que (h)", _
                "DVL HARD (ans)", "DVL SOFT (ans)", "Devise")
        Case "INFO_ATM", "INFOATM"
            GetHeadersForTable = Array( _
                "DTR", "Descriptif FR", "Descriptif EN", "HARD/SOFT", _
                "Indice", "Plan ATM", "Commentaire FR", "Commentaire EN", _
                "R�parabilit�", "Poids (kg)", "DVL g�n�rique (ans)", _
                "DVL sp�. projet (ans)", "Tps g�n�rique (h)", "Tps analytique cumul� (h)")
        Case "TACHES_ATM", "TACHESATM"
            GetHeadersForTable = Array( _
                "DTR", "Langue", "Rang", "HARD/SOFT", "Cat�gorie", _
                "Niveau Qualif", "Description", "Crit�res dim.", _
                "Outillage sp�cifique", "Outillage standard", _
                "Pi�ces rech./conso.", "Nb Op�rateurs", _
                "Tps estim� (h)", "Tps analytique cumul� (h)", _
                "Pictogramme", "Commentaire picto")
        Case "EMMST"
            GetHeadersForTable = Array( _
                "DTR", "Dim. Nu L (mm)", "Dim. Nu l (mm)", "Dim. Nu H (mm)", "Poids Nu (kg)", _
                "Dim. Cond. L (mm)", "Dim. Cond. l (mm)", "Dim. Cond. H (mm)", "Poids Cond. (kg)", _
                "Date P�remption", "Unit� Cond.", "Dur�e P�remption", _
                "Commentaire Emballage FR", "Commentaire Emballage EN", _
                "Commentaire Stockage FR", "Commentaire Stockage EN", _
                "Commentaire Transport FR", "Commentaire Transport EN", "Version")
        Case "PROJET_ASL", "PROJETASL"
            GetHeadersForTable = Array( _
                "Projet", "DTR", "Sous-Syst�me", "D�signation", "Quantit�", _
                "MTBF Fonctionnel (h)", "DVL analytique HARD (ans)", _
                "DVL analytique SOFT (ans)", "ISP TIS (archiv�)")
        Case Else
            GetHeadersForTable = Array("Colonne 1", "Colonne 2", "Colonne 3")
    End Select
End Function

' ============================================================
' LIRE UN FICHIER EN BINAIRE
' ============================================================

Private Function ReadFileBinary(filePath As String) As Variant

    Dim fileNum  As Integer
    Dim fileSize As Long
    Dim data()   As Byte

    On Error GoTo ErrRead

    fileSize = FileLen(filePath)
    If fileSize = 0 Then Exit Function

    ReDim data(0 To fileSize - 1)
    fileNum = FreeFile
    Open filePath For Binary Access Read As #fileNum
    Get #fileNum, , data
    Close #fileNum

    ReadFileBinary = data
    Exit Function

ErrRead:
    MsgBox "Erreur ReadFileBinary : " & Err.Number & " � " & Err.description, vbCritical, "ATMAP"
    ReadFileBinary = Empty

End Function

' ============================================================
' DIGEST CSRF
' ============================================================

Private Function GetFormDigest() As String

    Dim xhr  As Object
    Dim resp As String

    GetFormDigest = ""
    Set xhr = CreateObject("MSXML2.XMLHTTP.6.0")

    On Error GoTo ErrDigest

    xhr.Open "POST", SITE_URL & "/_api/contextinfo", False
    xhr.setRequestHeader "Accept", "application/json;odata=verbose"
    xhr.setRequestHeader "Content-Length", "0"
    xhr.Send

    resp = xhr.responseText
    GetFormDigest = ParseJsonString(resp, """FormDigestValue"":""")
    Exit Function

ErrDigest:
    MsgBox "Erreur GetFormDigest : " & Err.Number & " � " & Err.description, vbCritical, "ATMAP"

End Function

' ============================================================
' UTILISATEUR CONNECT�
' ============================================================

Private Function GetCurrentUserID(digest As String) As Long

    Dim xhr  As Object
    Dim resp As String
    Dim pos1 As Long
    Dim posC As Long
    Dim posB As Long
    Dim pos2 As Long

    GetCurrentUserID = -1
    Set xhr = CreateObject("MSXML2.XMLHTTP.6.0")

    On Error GoTo ErrUser

    xhr.Open "GET", SITE_URL & "/_api/web/currentuser?$select=Id", False
    xhr.setRequestHeader "Accept", "application/json;odata=verbose"
    xhr.Send

    resp = xhr.responseText
    Debug.Print "CurrentUser : " & resp

    pos1 = InStr(resp, """Id"":") + Len("""Id"":")
    If pos1 <= Len("""Id"":") Then Exit Function

    posC = InStr(pos1, resp, ","): If posC = 0 Then posC = Len(resp) + 1
    posB = InStr(pos1, resp, "}"): If posB = 0 Then posB = Len(resp) + 1
    pos2 = IIf(posC < posB, posC, posB)

    If pos2 > pos1 Then
        GetCurrentUserID = CLng(Trim(Mid(resp, pos1, pos2 - pos1)))
    End If

    Debug.Print "Current user ID : " & GetCurrentUserID
    Exit Function

ErrUser:
    MsgBox "Erreur GetCurrentUserID : " & Err.Number & " � " & Err.description, vbCritical, "ATMAP"
    GetCurrentUserID = -1

End Function

' ============================================================
' UTILITAIRES � DIAGNOSTIC
' ============================================================

Sub GetFieldInternalNames()

    Dim xhr   As Object
    Dim resp  As String
    Dim ws    As Worksheet
    Dim pos   As Long
    Dim p1    As Long, p2 As Long
    Dim ligne As Long

    Set xhr = CreateObject("MSXML2.XMLHTTP.6.0")
    xhr.Open "GET", SITE_URL & "/_api/web/lists/getbytitle('" & LIST_NAME & "')" & _
             "/fields?$select=Title,InternalName,TypeAsString&$filter=Hidden eq false", False
    xhr.setRequestHeader "Accept", "application/json;odata=verbose"
    xhr.Send
    resp = xhr.responseText

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Debug_Fields")
    If ws Is Nothing Then Set ws = ThisWorkbook.Sheets.Add: ws.Name = "Debug_Fields"
    On Error GoTo 0

    ws.Cells.Clear
    ws.Cells(1, 1) = "Title": ws.Cells(1, 2) = "InternalName": ws.Cells(1, 3) = "TypeAsString"
    ligne = 2: pos = 1

    Do
        p1 = InStr(pos, resp, """Title"":""") + Len("""Title"":""")
        p2 = InStr(p1, resp, """")
        If p1 <= Len("""Title"":""") Or p2 <= p1 Then Exit Do
        ws.Cells(ligne, 1) = Mid(resp, p1, p2 - p1)

        p1 = InStr(pos, resp, """InternalName"":""") + Len("""InternalName"":""")
        p2 = InStr(p1, resp, """")
        ws.Cells(ligne, 2) = Mid(resp, p1, p2 - p1)

        p1 = InStr(pos, resp, """TypeAsString"":""") + Len("""TypeAsString"":""")
        p2 = InStr(p1, resp, """")
        ws.Cells(ligne, 3) = Mid(resp, p1, p2 - p1)

        ligne = ligne + 1
        pos = InStr(pos, resp, """InternalName"":""") + 1
        If pos <= 1 Then Exit Do
    Loop While pos < Len(resp) And ligne < 200

    ws.Columns("A:C").AutoFit
    ws.Activate
    MsgBox ligne - 2 & " colonnes �crites.", vbInformation, "ATMAP"

End Sub

Sub GetListItemEntityType()

    Dim xhr  As Object
    Dim resp As String

    Set xhr = CreateObject("MSXML2.XMLHTTP.6.0")
    xhr.Open "GET", SITE_URL & "/_api/web/lists/getbytitle('" & LIST_NAME & "')" & _
             "?$select=ListItemEntityTypeFullName", False
    xhr.setRequestHeader "Accept", "application/json;odata=verbose"
    xhr.Send
    resp = xhr.responseText

    MsgBox "ListItemEntityTypeFullName = " & _
           ParseJsonString(resp, """ListItemEntityTypeFullName"":"""), vbInformation, "ATMAP"

End Sub

Sub Diagnostic_Connexion()

    Dim xhr As Object
    Set xhr = CreateObject("MSXML2.XMLHTTP.6.0")
    xhr.Open "GET", SITE_URL & "/_api/web?$select=Title", False
    xhr.setRequestHeader "Accept", "application/json;odata=verbose"
    xhr.Send

    MsgBox "Status : " & xhr.Status & vbCrLf & Left(xhr.responseText, 300), _
           vbInformation, "Diagnostic connexion SharePoint"

End Sub

' ============================================================
' UTILITAIRES � PARSING JSON
' ============================================================

Private Function ParseJsonString(json As String, key As String) As String
    Dim pos1 As Long, pos2 As Long
    pos1 = InStr(json, key) + Len(key)
    pos2 = InStr(pos1, json, """")
    If pos1 > Len(key) And pos2 > pos1 Then
        ParseJsonString = Mid(json, pos1, pos2 - pos1)
    End If
End Function

Private Function ParseJsonLong(json As String, key As String) As String
    Dim pos1 As Long, posC As Long, posB As Long, pos2 As Long
    pos1 = InStr(json, key) + Len(key)
    posC = InStr(pos1, json, ","): If posC = 0 Then posC = Len(json) + 1
    posB = InStr(pos1, json, "}"): If posB = 0 Then posB = Len(json) + 1
    pos2 = IIf(posC < posB, posC, posB)
    If pos2 > pos1 Then ParseJsonLong = Trim(Mid(json, pos1, pos2 - pos1))
End Function

Private Function EscJson(s As String) As String
    EscJson = Replace(Replace(s, "\", "\\"), """", "\""")
End Function

Private Function EscSql(s As String) As String
    EscSql = Replace(s, "'", "''")
End Function



