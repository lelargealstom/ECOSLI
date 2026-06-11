Attribute VB_Name = "G_SIRUS_ATMAP"
Option Explicit

' ============================================================
' CONSTANTES
' ============================================================
Private Const SITE_URL         As String = "https://alstomgroup.sharepoint.com/sites/Svcs_DIS_Fr_VBN"
Private Const LIST_NAME        As String = "ATMAP_TICKETS"
Private Const LIST_ENTITY_TYPE As String = "SP.Data.ATMAP_x005f_TICKETSListItem"

' ============================================================
' POINT D'ENTReE a TEST
' ============================================================

Sub Test_Creer_Ticket_ATMAP()

    Dim digest   As String
    Dim userID   As Long
    Dim newID    As Long
    Dim result   As String
    Dim pjResult As String

    ' -- etape 1 a Digest CSRF ------------------------------
    digest = GetFormDigest()
    If digest = "" Then
        MsgBox "echec recuperation digest.", vbCritical, "ATMAP Test"
        Exit Sub
    End If

    ' -- etape 2 a Utilisateur connecta ---------------------
    userID = GetCurrentUserID(digest)
    If userID = -1 Then
        MsgBox "Impossible de recuperer l'utilisateur connecte.", vbExclamation, "ATMAP Test"
        Exit Sub
    End If

    ' -- etape 3 a Donnees d'import -------------------------
    Dim lignesImport As New Collection
    lignesImport.Add Array( _
        "DTR-001-A", "Calculateur traction", "LRU", "Bombardier", _
        "BC-2241-TRC", "DRW-4412", "45000", "15", "10", "EUR")
    lignesImport.Add Array( _
        "DTR-002-B", "Carte alimentation 24V", "SRU", "Alstom", _
        "ALT-PWR-024", "DRW-0981", "22000", "12", "", "EUR")

    ' -- etape 4 a Creation ticket --------------------------
    result = CreateTicketREST( _
        digest:=digest, _
        userID:=userID, _
        outilID:=43, _
        numeroTicket:="VBN_Test", _
        sujet:="IMP_001 a Import Produit", _
        statut:="Nouveau", _
        description:="Import de " & lignesImport.count & _
                     " DTR en table Produit a DTR-001-A, DTR-002-B", _
        attentes:="Integration en BdD SLI a Table Produit", _
        typedemande:="Import", _
        priorite:="High", _
        ticketIDOut:=newID _
    )

    MsgBox result, vbInformation, "ATMAP Test a Resultat"

    If newID = -1 Then Exit Sub

    ' -- etape 5 a Piece jointe Excel -----------------------
    pjResult = AttachExcelToTicket( _
        ticketID:=newID, _
        digest:=digest, _
        tableName:="Produit", _
        numeroTicket:="VBN_Test", _
        dataRows:=lignesImport _
    )

    MsgBox pjResult, vbInformation, "ATMAP Test a Piece jointe"

End Sub

' ============================================================
' CREATION TICKET COMPLET a UNE SEULE REQUeTE REST POST
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
                                   ticketIDOut As Long, _
                                   Optional numeroImport As String = "", _
                                   Optional dateEstimee As String = "") As String

    Dim xhr   As Object
    Dim url   As String
    Dim body  As String
    Dim resp  As String
    Dim today As String
    Dim champImport As String
    Dim champDateEst As String

    ticketIDOut = -1

    ' MSXML2 valida a ne pas changer pour WinHttp ici
    Set xhr = CreateObject("MSXML2.XMLHTTP.6.0")

    url = SITE_URL & "/_api/web/lists/getbytitle('" & LIST_NAME & "')/items"
    today = Format(Now(), "yyyy-mm-dd") & "T00:00:00Z"

    ' Champ Numero_Import optionnel (non emis si vide)
    If LenB(numeroImport) > 0 Then
        champImport = """Numero_Import"":""" & EscJson(numeroImport) & ""","
    End If

    ' Champ Date_estimee optionnel : ISO yyyy-mm-ddT00:00:00Z attendu.
    ' Nom interne SP confirme par Thomas : "Date_estimee" (sans accent).
    If LenB(dateEstimee) > 0 Then
        champDateEst = """Date_estimee"":""" & dateEstimee & ""","
    End If

    ' Pas de Title a champ non utilisa dans ATMAP
    body = "{""__metadata"":{""type"":""" & LIST_ENTITY_TYPE & """}," & _
           """Numero_ticket"":""" & EscJson(numeroTicket) & """," & _
           champImport & _
           """Sujet"":""" & EscJson(sujet) & """," & _
           """Statut"":""" & EscJson(statut) & """," & _
           """Description"":""" & EscJson(description) & """," & _
           """Attentes"":""" & EscJson(attentes) & """," & _
           """Typedemande"":""" & EscJson(typedemande) & """," & _
           """Priorit_x00e9_"":""" & EscJson(priorite) & """," & _
           """Datecr_x00e9_ation"":""" & today & """," & _
           champDateEst & _
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
        CreateTicketREST = "Ticket crea a ID : " & ticketIDOut & vbCrLf & _
                           "Numero      : " & numeroTicket & vbCrLf & _
                           "DemandeurId : " & userID & vbCrLf & _
                           "NomoutilId  : " & outilID
    Else
        CreateTicketREST = "Erreur HTTP " & xhr.Status & vbCrLf & Left(resp, 500)
    End If
    Exit Function

ErrCreate:
    CreateTicketREST = "Erreur : " & Err.Number & " a " & Err.description
    MsgBox CreateTicketREST, vbCritical, "ATMAP a CreateTicketREST"

End Function

' ============================================================
' CREER ET ATTACHER UN FICHIER EXCEL AU TICKET
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

    ' -- 2. Creer le classeur --------------------------------
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    xlApp.DisplayAlerts = False

    Set xlWb = xlApp.Workbooks.Add
    Set xlWs = xlWb.Sheets(1)
    xlWs.Name = tableName

    headers = GetHeadersForTable(tableName)

    ' -- 3. Ligne titre fusionnee ----------------------------
    xlWs.rows(1).RowHeight = 28
    With xlWs.Range(xlWs.Cells(1, 1), xlWs.Cells(1, UBound(headers) + 1))
        .Merge
        .Value = "SIRUS a Import a " & tableName & _
                 "  |  " & numeroTicket & _
                 "  |  " & Format(Now(), "dd/mm/yyyy hh:mm")
        .Font.Bold = True
        .Font.Size = 10
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 56, 100)
        .HorizontalAlignment = -4108
        .VerticalAlignment = -4108
    End With

    ' -- 4. En-tetes colonnes --------------------------------
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

    ' -- 5. Donnees ------------------------------------------
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

    ' -- 6. Tableau structura --------------------------------
    If rowIdx > 3 Then
        Dim lo As Object
        Set lo = xlWs.ListObjects.Add(1, _
            xlWs.Range(xlWs.Cells(2, 1), _
                       xlWs.Cells(rowIdx - 1, UBound(headers) + 1)), , 1)
        lo.Name = "tbl_Import_" & tableName
        lo.TableStyle = "TableStyleMedium2"
    End If

    ' -- 7. Figer titres + en-tetes --------------------------
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
    
    ' -- 10. Uploader via REST a WinHttp pour les PJ ---------
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
        AttachExcelToTicket = "PJ Excel OK a " & xlName & vbCrLf & _
                              dataRows.count & " ligne(s) exportee(s)"
    Else
        AttachExcelToTicket = "PJ HTTP " & uploadStatus & vbCrLf & _
                              Left(xhr.responseText, 300)
    End If
    Exit Function

ErrExcel:
    AttachExcelToTicket = "Erreur Excel : " & Err.Number & " a " & Err.description
    MsgBox AttachExcelToTicket, vbCritical, "ATMAP a AttachExcelToTicket"
    On Error Resume Next
    If Not xlWb Is Nothing Then xlWb.Close False
    If Not xlApp Is Nothing Then xlApp.Quit
    On Error GoTo 0

End Function

' =========================================================================
' =========================================================================
' SIRUS Phase 5 : API publique pour le UserForm
' =========================================================================
' Wrappers + helpers utilises par SIRUS_UI_BtnEnvoyerClick (G_SIRUS_UserForm).
' Les fonctions Private du module (CreateTicketREST, GetFormDigest,
' GetCurrentUserID) sont exposees via des wrappers publics SIRUS_*.
' =========================================================================

' Wrapper public sur GetFormDigest
Public Function SIRUS_GetFormDigest() As String
    SIRUS_GetFormDigest = GetFormDigest()
End Function

' Wrapper public sur GetCurrentUserID
Public Function SIRUS_GetCurrentUserID(ByVal digest As String) As Long
    SIRUS_GetCurrentUserID = GetCurrentUserID(digest)
End Function

' Wrapper public sur CreateTicketREST avec defaults SIRUS :
'   statut = "Nouveau"
'   typedemande = "Import"
'   outilID = 43 (ID outil ECOSLI / SIRUS dans ATMAP - a confirmer si change)
Public Function SIRUS_CreerTicketEnvoi(ByVal digest As String, _
                                         ByVal userID As Long, _
                                         ByVal numeroTicket As String, _
                                         ByVal sujet As String, _
                                         ByVal description As String, _
                                         ByVal attentes As String, _
                                         ByVal priorite As String, _
                                         ByRef ticketIDOut As Long, _
                                         Optional ByVal numeroImport As String = "", _
                                         Optional ByVal dateEstimee As String = "") As String
    SIRUS_CreerTicketEnvoi = CreateTicketREST( _
        digest:=digest, _
        userID:=userID, _
        outilID:=43, _
        numeroTicket:=numeroTicket, _
        sujet:=sujet, _
        statut:="Nouveau", _
        description:=description, _
        attentes:=attentes, _
        typedemande:="Import", _
        priorite:=priorite, _
        ticketIDOut:=ticketIDOut, _
        numeroImport:=numeroImport, _
        dateEstimee:=dateEstimee)
End Function

' =========================================================================
' SIRUS_JoursTraitementSelonPriorite
' =========================================================================
' Centralise la regle metier "priorite -> nombre de jours de traitement".
' Utilise par SIRUS_UI_PrioriteChange (Attentes) ET SIRUS_DateEstimeeISO
' (champ Date_estimee du ticket). Modifier ici si la regle evolue.
'   High   -> 1 jour
'   Medium -> 3 jours  (defaut si priorite inconnue ou vide)
'   Low    -> 5 jours
Public Function SIRUS_JoursTraitementSelonPriorite(ByVal sPriorite As String) As Long
    Select Case sPriorite
        Case "High":   SIRUS_JoursTraitementSelonPriorite = 1
        Case "Low":    SIRUS_JoursTraitementSelonPriorite = 5
        Case Else:     SIRUS_JoursTraitementSelonPriorite = 3
    End Select
End Function

' =========================================================================
' SIRUS_DateEstimeeISO
' =========================================================================
' Calcule la date estimee = aujourd'hui + N jours (selon priorite) et
' renvoie au format ISO yyyy-mm-ddT00:00:00Z attendu par SharePoint.
Public Function SIRUS_DateEstimeeISO(ByVal sPriorite As String) As String
    Dim nbJours As Long
    Dim dEst As Date
    nbJours = SIRUS_JoursTraitementSelonPriorite(sPriorite)
    dEst = DateAdd("d", nbJours, Date)
    SIRUS_DateEstimeeISO = Format(dEst, "yyyy-mm-dd") & "T00:00:00Z"
End Function

' =========================================================================
' SIRUS_NumeroVBNSuivant
' =========================================================================
' Genere le prochain Numero_ticket au format VBN_NN, suivant l'increment
' des tickets deja presents dans SharePoint ATMAP_TICKETS.
' Strategie : GET top 500 (tous types confondus), parse "VBN_<nb>",
' max +1. Si rien -> VBN_1. Le format est entier brut, pas zero-padde
' (ex: VBN_20 -> VBN_21).
Public Function SIRUS_NumeroVBNSuivant() As String
    Dim maxNum As Long
    Dim numNouveau As Long

    maxNum = ChercherMaxNumeroVBN()
    If maxNum < 1 Then
        numNouveau = 1
    Else
        numNouveau = maxNum + 1
    End If

    SIRUS_NumeroVBNSuivant = "VBN_" & numNouveau
End Function

' Helper prive : interroge SharePoint pour le max VBN_<nb> tous types
' confondus. Ignore les VBN_<non_numerique> (ex: VBN_Test).
Private Function ChercherMaxNumeroVBN() As Long
    Dim xhr As Object
    Dim url As String
    Dim resp As String
    Dim maxNum As Long
    Dim pos As Long
    Dim startPos As Long, endPos As Long
    Dim numTicket As String
    Dim suffix As String

    maxNum = 0

    On Error GoTo Sortie

    Set xhr = CreateObject("MSXML2.XMLHTTP.6.0")
    ' Recupere les 500 derniers tickets (tous types). Suffisant pour le
    ' max courant VBN_<nb>.
    url = SITE_URL & "/_api/web/lists/getbytitle('" & LIST_NAME & "')/items" & _
          "?$select=Numero_ticket&$orderby=Id desc&$top=500"

    xhr.Open "GET", url, False
    xhr.setRequestHeader "Accept", "application/json;odata=verbose"
    xhr.Send

    If xhr.Status <> 200 Then
        ChercherMaxNumeroVBN = 0
        Exit Function
    End If

    resp = xhr.responseText

    ' Parser tous les "Numero_ticket":"VBN_<...>"
    pos = 1
    Do
        startPos = InStr(pos, resp, """Numero_ticket"":""")
        If startPos = 0 Then Exit Do
        startPos = startPos + Len("""Numero_ticket"":""")
        endPos = InStr(startPos, resp, """")
        If endPos = 0 Then Exit Do

        numTicket = Mid$(resp, startPos, endPos - startPos)
        pos = endPos + 1

        ' Filtrer prefixe VBN_
        If Left$(numTicket, 4) = "VBN_" Then
            suffix = Mid$(numTicket, 5)
            If IsNumeric(suffix) Then
                If CLng(suffix) > maxNum Then maxNum = CLng(suffix)
            End If
        End If
    Loop

    ChercherMaxNumeroVBN = maxNum
    Exit Function

Sortie:
    Debug.Print "ChercherMaxNumeroVBN erreur : " & Err.Description
    ChercherMaxNumeroVBN = 0
End Function

' =========================================================================
' SIRUS_NumeroTicketSuivant
' =========================================================================
' Genere le prochain numero de ticket Import au format IMP-YYYY-00X.
' Strategie : GET sur la liste SharePoint ATMAP_TICKETS filtree
'   Typedemande = 'Import', parse les Numero_ticket pour trouver le max
'   de l annee en cours, incremente +1. Si rien -> 001. Passage d annee
'   transparent (premier ticket d une nouvelle annee = IMP-YYYY-001).
Public Function SIRUS_NumeroTicketSuivant() As String
    Dim annee As Long
    Dim maxNum As Long
    Dim numNouveau As Long

    annee = Year(Now())
    maxNum = ChercherMaxNumeroImport(annee)
    If maxNum < 1 Then
        numNouveau = 1
    Else
        numNouveau = maxNum + 1
    End If

    SIRUS_NumeroTicketSuivant = "IMP-" & annee & "-" & Format$(numNouveau, "000")
End Function

' Helper prive : interroge SharePoint pour le max du champ Numero_Import
' (format IMP-YYYY-NNN) sur les tickets Import de l'annee en cours.
' Note : depuis le split Numero_ticket (VBN_NN) / Numero_Import (IMP-...)
' du 2026-06-11, c'est bien Numero_Import qu il faut parser ici, pas
' Numero_ticket comme avant.
Private Function ChercherMaxNumeroImport(ByVal annee As Long) As Long
    Dim xhr As Object
    Dim url As String
    Dim resp As String
    Dim prefixe As String
    Dim maxNum As Long
    Dim pos As Long
    Dim startPos As Long, endPos As Long
    Dim numImport As String
    Dim suffix As String

    prefixe = "IMP-" & annee & "-"
    maxNum = 0

    On Error GoTo Sortie

    Set xhr = CreateObject("MSXML2.XMLHTTP.6.0")
    ' Recupere les 500 derniers tickets Import. Suffisant pour trouver le
    ' max de l annee en cours (au-dela les tickets sont vieux d au moins
    ' plusieurs annees).
    url = SITE_URL & "/_api/web/lists/getbytitle('" & LIST_NAME & "')/items" & _
          "?$select=Numero_Import,Typedemande&$filter=Typedemande eq 'Import'" & _
          "&$orderby=Id desc&$top=500"

    xhr.Open "GET", url, False
    xhr.setRequestHeader "Accept", "application/json;odata=verbose"
    xhr.Send

    If xhr.Status <> 200 Then
        ChercherMaxNumeroImport = 0
        Exit Function
    End If

    resp = xhr.responseText

    ' Parser tous les "Numero_Import":"IMP-YYYY-NNN"
    pos = 1
    Do
        startPos = InStr(pos, resp, """Numero_Import"":""")
        If startPos = 0 Then Exit Do
        startPos = startPos + Len("""Numero_Import"":""")
        endPos = InStr(startPos, resp, """")
        If endPos = 0 Then Exit Do

        numImport = Mid$(resp, startPos, endPos - startPos)
        pos = endPos + 1

        ' Filtrer prefixe IMP-YYYY-
        If Left$(numImport, Len(prefixe)) = prefixe Then
            suffix = Mid$(numImport, Len(prefixe) + 1)
            If IsNumeric(suffix) Then
                If CLng(suffix) > maxNum Then maxNum = CLng(suffix)
            End If
        End If
    Loop

    ChercherMaxNumeroImport = maxNum
    Exit Function

Sortie:
    Debug.Print "ChercherMaxNumeroImport erreur : " & Err.description
    ChercherMaxNumeroImport = 0
End Function

' =========================================================================
' SIRUS_ConstruirePJ
' =========================================================================
' Construit un classeur Excel dans %TEMP% en copiant les feuilles SIRUS_*
' depuis le classeur projet actif. Retourne le chemin du fichier sauve.
' Cas multi-feuilles (Extraction ATM) : InfoATM + TachesATM dans un meme
' classeur (logique : 1 livrable = 1 PJ).
Public Function SIRUS_ConstruirePJ(ByRef feuillesNoms As Variant, _
                                     ByVal sNomBase As String) As String
    Dim xlPath As String
    Dim xlName As String
    Dim wbProjet As Workbook
    Dim wbPJ As Workbook
    Dim i As Long
    Dim ws As Worksheet

    Set wbProjet = ActiveWorkbook

    xlName = sNomBase & "_" & Format$(Now(), "yyyymmdd_hhmmss") & ".xlsx"
    xlPath = Environ$("TEMP") & "\" & xlName

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

    On Error GoTo Sortie

    ' Nouveau classeur (par defaut 1 feuille a supprimer ensuite)
    Set wbPJ = Workbooks.Add

    ' Copier chaque feuille SIRUS_* depuis le projet
    For i = LBound(feuillesNoms) To UBound(feuillesNoms)
        Set ws = Nothing
        On Error Resume Next
        Set ws = wbProjet.Worksheets(CStr(feuillesNoms(i)))
        On Error GoTo Sortie
        If Not ws Is Nothing Then
            ws.Copy After:=wbPJ.Worksheets(wbPJ.Worksheets.count)
        End If
    Next i

    ' Supprimer la feuille par defaut (Sheet1) si feuilles copiees presentes
    If wbPJ.Worksheets.count > 1 Then
        wbPJ.Worksheets(1).Delete
    End If

    wbPJ.SaveAs Filename:=xlPath, FileFormat:=51  ' xlOpenXMLWorkbook (.xlsx)
    wbPJ.Close False

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    SIRUS_ConstruirePJ = xlPath
    Exit Function

Sortie:
    On Error Resume Next
    If Not wbPJ Is Nothing Then wbPJ.Close False
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    SIRUS_ConstruirePJ = ""
End Function

' =========================================================================
' SIRUS_AttacherFichierAuTicket
' =========================================================================
' Upload un fichier deja construit (chemin local) en PJ a un ticket ATMAP
' existant. Utilise le digest fourni. Lit en binaire puis POST en
' octet-stream via XMLHTTP.
Public Function SIRUS_AttacherFichierAuTicket(ByVal ticketID As Long, _
                                                ByVal digest As String, _
                                                ByVal cheminFichier As String) As String
    Dim fileData As Variant
    Dim xhr As Object
    Dim stream As Object
    Dim xlName As String
    Dim uploadStatus As Long

    On Error GoTo Sortie

    xlName = Mid$(cheminFichier, InStrRev(cheminFichier, "\") + 1)

    fileData = ReadFileBinary(cheminFichier)
    If IsEmpty(fileData) Then
        SIRUS_AttacherFichierAuTicket = "Erreur : lecture fichier impossible (" & cheminFichier & ")"
        Exit Function
    End If

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

    If uploadStatus = 200 Then
        SIRUS_AttacherFichierAuTicket = "PJ uploadee : " & xlName
    Else
        SIRUS_AttacherFichierAuTicket = "PJ HTTP " & uploadStatus & vbCrLf & _
                                          Left$(xhr.responseText, 300)
    End If
    Exit Function

Sortie:
    SIRUS_AttacherFichierAuTicket = "Erreur SIRUS_AttacherFichierAuTicket : " & Err.Description
End Function

' ============================================================
' EN-TeTES PAR TABLE
' ============================================================

Private Function GetHeadersForTable(tableName As String) As Variant
    Select Case UCase(tableName)
        Case "PRODUIT"
            GetHeadersForTable = Array( _
                "DTR", "Designation", "Node Type", "Fournisseur", _
                "Ref. Fournisseur", "Drawing Number", "MTBF Intrinseque (h)", _
                "DVL HARD (ans)", "DVL SOFT (ans)", "Devise")
        Case "INFO_ATM", "INFOATM"
            GetHeadersForTable = Array( _
                "DTR", "Descriptif FR", "Descriptif EN", "HARD/SOFT", _
                "Indice", "Plan ATM", "Commentaire FR", "Commentaire EN", _
                "Reparabilite", "Poids (kg)", "DVL generique (ans)", _
                "DVL spe. projet (ans)", "Tps generique (h)", "Tps analytique cumula (h)")
        Case "TACHES_ATM", "TACHESATM"
            GetHeadersForTable = Array( _
                "DTR", "Langue", "Rang", "HARD/SOFT", "Categorie", _
                "Niveau Qualif", "Description", "Criteres dim.", _
                "Outillage specifique", "Outillage standard", _
                "Pieces rech./conso.", "Nb Operateurs", _
                "Tps estima (h)", "Tps analytique cumula (h)", _
                "Pictogramme", "Commentaire picto")
        Case "EMMST"
            GetHeadersForTable = Array( _
                "DTR", "Dim. Nu L (mm)", "Dim. Nu l (mm)", "Dim. Nu H (mm)", "Poids Nu (kg)", _
                "Dim. Cond. L (mm)", "Dim. Cond. l (mm)", "Dim. Cond. H (mm)", "Poids Cond. (kg)", _
                "Date Peremption", "Unita Cond.", "Duree Peremption", _
                "Commentaire Emballage FR", "Commentaire Emballage EN", _
                "Commentaire Stockage FR", "Commentaire Stockage EN", _
                "Commentaire Transport FR", "Commentaire Transport EN", "Version")
        Case "PROJET_ASL", "PROJETASL"
            GetHeadersForTable = Array( _
                "Projet", "DTR", "Sous-Systeme", "Designation", "Quantite", _
                "MTBF Fonctionnel (h)", "DVL analytique HARD (ans)", _
                "DVL analytique SOFT (ans)", "ISP TIS (archive)")
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
    MsgBox "Erreur ReadFileBinary : " & Err.Number & " a " & Err.description, vbCritical, "ATMAP"
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
    MsgBox "Erreur GetFormDigest : " & Err.Number & " a " & Err.description, vbCritical, "ATMAP"

End Function

' ============================================================
' UTILISATEUR CONNECTe
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
    MsgBox "Erreur GetCurrentUserID : " & Err.Number & " a " & Err.description, vbCritical, "ATMAP"
    GetCurrentUserID = -1

End Function

' ============================================================
' UTILITAIRES a DIAGNOSTIC
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
    MsgBox ligne - 2 & " colonnes ecrites.", vbInformation, "ATMAP"

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
' UTILITAIRES a PARSING JSON
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



