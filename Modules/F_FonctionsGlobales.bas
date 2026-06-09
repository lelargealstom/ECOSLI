Attribute VB_Name = "F_FonctionsGlobales"
Option Explicit
Public Function LastRowInCol(ByVal ws As Worksheet, ByVal colIndex As Long) As Long
    ' Renvoie la derni�re ligne utilis�e de la colonne colIndex (>= 1)
    Dim lastCell As Range
    On Error Resume Next
    Set lastCell = ws.Cells(ws.rows.count, colIndex).End(xlUp)
    On Error GoTo 0
    If lastCell Is Nothing Then
        LastRowInCol = 1
    ElseIf lastCell.Row < 1 Then
        LastRowInCol = 1
    Else
        LastRowInCol = lastCell.Row
    End If
End Function
Public Function NzStr(v As Variant, Optional fallback As String = "") As String
    If IsError(v) Then NzStr = fallback Else NzStr = Trim$(CStr(v))
End Function
Public Function SafeNum(v As Variant, Optional fallback As Double = 0#) As Double
    If IsNumeric(v) Then SafeNum = CDbl(v) Else SafeNum = fallback
End Function
Public Function NodeTypeToLevel(ByVal NodeType As String) As Long
    Select Case UCase$(Trim$(NodeType))
        Case "LRU", "LRUNR", "0", "NON TROUVE", "NON TROUV�"
            NodeTypeToLevel = 3
        Case "SRU", "SRUNR"
            NodeTypeToLevel = 4
        Case "FAMILY"
            NodeTypeToLevel = 2
        Case Else
            NodeTypeToLevel = 3
    End Select
End Function

'Fonction qui v�rifie si une feuille est pr�sente ou non dans un workbooks
'renvoie un False si elle n'est pas l�, un true si elle est pr�sente
Public Function Test_feuille(FeuilleAVerifier As String) As Boolean
'fonction qui v�rifie si la "FeuilleAVerifier" existe dans le Classeur actif

    On Error GoTo SiErreur
    Dim feuille
        
        Test_feuille = False
        For Each feuille In Sheets
            If UCase(feuille.Name) = UCase(FeuilleAVerifier) Then
                Test_feuille = True
                Exit Function
            End If
        Next feuille
    Exit Function
    
SiErreur:
        Test_feuille = CVErr(xlErrNA)
End Function
Public Function Niv_Maintenance(tableau As Variant, projet As String) As String
'Fonction qui permet de r�cup�rer le niveau de maintenance d'un projet
    Dim colonne_nvmaint As Integer
    Dim colonne_projet  As Integer
    Dim i As Integer
   
    colonne_nvmaint = Arraypos(tableau, "Niv de maintenance_PROJET SLI")
    colonne_projet = Arraypos(tableau, "Nom projet_PROJET SLI")
    For i = LBound(tableau) To UBound(tableau)
        If tableau(i, colonne_projet) = projet Then
            If IsNull(tableau(i, colonne_nvmaint)) Then
                Niv_Maintenance = ""
            Else
                Niv_Maintenance = CStr(tableau(i, colonne_nvmaint))
            End If
            Exit For
        End If
    Next
    
End Function
Public Function Auteur_Projet(tableau As Variant, projet As String) As String
'Fonction qui permet de r�cup�rer le niveau de maintenance d'un projet
    Dim colonne_auteur As Integer
    Dim colonne_projet  As Integer
    Dim i As Integer
   
    colonne_auteur = Arraypos(tableau, "RASL_PROJET SLI")
    colonne_projet = Arraypos(tableau, "Nom projet_PROJET SLI")
    For i = LBound(tableau) To UBound(tableau)
        If tableau(i, colonne_projet) = projet Then
            If IsNull(tableau(i, colonne_auteur)) Then
                Auteur_Projet = ""
            Else
                Auteur_Projet = CStr(tableau(i, colonne_auteur))
            End If
            Exit For
        End If
    Next
    
End Function
'Fonction qui met � jour la barre de progression dans le UserForm
Function barprogress(bar_cpt, bar_tot)
    
    UserForm.barre.Width = bar_cpt / bar_tot * UserForm.barreimage.Width
    UserForm.barretexte.Caption = CInt((bar_cpt / bar_tot) * 100) & "%"
    DoEvents
        
End Function
'Fonction qui renvoie l'abr�viation d'un code projet
Public Function Abrev_Projet(Code_Projet As String) As String
    Dim nb As Integer
    Dim pos1 As Integer, pos2 As Integer
    Dim temp As String
    
    ' Compter le nombre de "_"
    nb = UBound(Split(Code_Projet, "_"))
    
    Select Case nb
        Case 0
            ' Aucun "_" ? renvoyer le code complet
            Abrev_Projet = Code_Projet
            
        Case 1
            ' Un seul "_" ? prendre avant et apr�s jusqu'� espace
            pos1 = InStr(Code_Projet, "_")
            temp = Left(Code_Projet, pos1 - 1) & "_" & Mid(Code_Projet, pos1 + 1)
            
            ' Si espace apr�s le deuxi�me morceau, couper
            If InStr(temp, " ") > 0 Then
                temp = Left(temp, InStr(temp, " ") - 1)
            End If
            
            Abrev_Projet = temp
            
        Case Is >= 2
            ' Deux "_" ou plus ? prendre tout avant le deuxi�me "_"
            pos2 = InStr(InStr(Code_Projet, "_") + 1, Code_Projet, "_")
            Abrev_Projet = Left(Code_Projet, pos2 - 1)
    End Select
End Function
' Convertir l'index de colonne num�rique en r�f�rence de colonne Excel
Function Index_Colonne(colIndex As Long) As String
    Dim colLetter As String
    colLetter = ""
    Do
        colLetter = Chr((colIndex - 1) Mod 26 + 65) & colLetter
        colIndex = (colIndex - 1) \ 26
    Loop While colIndex > 0
    Index_Colonne = colLetter
End Function
'Turn off automatic calculations, events and screen updating
Public Sub TurnOffFunctionality()
    With Application
        .Calculation = xlCalculationManual
        .EnableEvents = False
        .DisplayAlerts = False
        .ScreenUpdating = False
        .PrintCommunication = False     ' evite les aller-retours driver imprimante (PageSetup)
    End With
End Sub
'Turn on automatic calculations, events and screen updating
Public Sub TurnOnFunctionality()
    With Application
        .Cursor = xlDefault
        .Calculation = xlCalculationAutomatic
        .DisplayAlerts = True
        .EnableEvents = True
        .ScreenUpdating = True
        .PrintCommunication = True
    End With
End Sub
Public Function HARD_SOFT(Range As Variant, dtr As String) As String
' Fonction permettant de renseigner si le DTR recherch� est un HARD, un SOFT ou poss�de les deux dans un range
' ---------------- Variable ----------------
' range : S�lection de la colonne HARD/SOFT de la BDD dans la table 'Taches_ATM' de la BDD
' ---------------- R�sultat ----------------
' Resultat = "HARD" si le DTR ne comprend que des lignes HARD
' Resultat = "SOFT" si le DTR ne comprend que des lignes SOFT
' Resultat = "BOTH" si le DTR comprend des lignes SOFT et HARD

If IsInArray(dtr, Range) = True Then
    HARD_SOFT = "true"
Else
    HARD_SOFT = "false"
End If

End Function
Public Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
' Fonction permettant de savoir si une valeur est pr�sente dans un tableau
' Renvoie True si Valeur est dans le tableau et False si elle n'y est pas
    
    Dim i
    For i = LBound(arr) To UBound(arr)
        If arr(i, Taches_type.dtr) = stringToBeFound Then
            IsInArray = True
            Exit Function
        End If
    Next i
    IsInArray = False

End Function
Public Function IsInArray_1D(stringToBeFound As String, arr As Variant) As Boolean
' Fonction permettant de savoir si une valeur est pr�sente dans un tableau 1D
' Renvoie True si Valeur est dans le tableau et False si elle n'y est pas

    Dim i
    For i = LBound(arr) To UBound(arr)
        If arr(i, 1) = stringToBeFound Then
            IsInArray_1D = True
            Exit Function
        End If
    Next i
    IsInArray_1D = False

End Function
Public Function IsInArrayNb(stringToBeFound As String, arr As Variant) As Integer
' Fonction permettant de compter le nombre d'occurrences d'une valeur dans un tableau
' Renvoie :
'   0 si la valeur n'est pas trouv�e
'   1 si elle est trouv�e une fois
'   x si elle est trouv�e x fois
 
    Dim i
    Dim count As Integer
    count = 0
    
    For i = LBound(arr) To UBound(arr)
        If arr(i, 1) = stringToBeFound Then
            count = count + 1
        End If
    Next i
    
    'On renvoie le nombre d'occurence
    IsInArrayNb = count

End Function
Public Function Contains_str(stringToBeFound As String, arr As Variant) As Integer
' Fonction permettant de savoir si une valeur est pr�sente dans un tableau 1D
' Renvoie True si Valeur est dans le tableau et False si elle n'y est pas

    Dim i
    For i = LBound(arr) To UBound(arr)
        If InStr(1, arr(i, 1), stringToBeFound) > 0 Then
            Contains_str = i
            Exit Function
        End If
    Next i
    Contains_str = 0

End Function
Public Function UpdateDesignation(des As String) As String
'Fonction permettant de mettre en forme la d�signation d'un code avant de sauver son ATM 2.0 dans le r�pertoire projet
Dim firstspace, secondspace As Double

firstspace = InStr(1, des, " ")

If firstspace = 0 Then
    Exit Function
Else
    secondspace = InStr(firstspace + 1, des, " ")
    If secondspace = 0 Then
        UpdateDesignation = Replace(Replace(Left(des, firstspace - 1), " ", "_"), "/", "_")
        Exit Function
    Else
        UpdateDesignation = Replace(Replace(Left(des, secondspace - 1), " ", "_"), "/", "_")
    End If
End If

End Function

Public Sub ClearOneControl(ByVal c As MSForms.control, ByVal ctrlName As String)
'Proc�dure qui vide un contr�le de formulaire selon le type de contr�le
    On Error Resume Next
    Select Case TypeName(c)
        Case "TextBox": c.Text = ""
        Case "ComboBox": c.Value = "": c.ListIndex = -1
        Case "CheckBox": c.Value = False: c.BackColor = RGB(180, 180, 180)
        Case "ListBox":
            Dim i As Long
            For i = 0 To c.ListCount - 1: c.Selected(i) = False: Next
            c.ListIndex = -1
        Case "SpinButton", "ScrollBar": c.Value = c.Min
        Case "Image": Set c.Picture = Nothing
        Case Else
            If InStr(1, ctrlName, "Date", vbTextCompare) > 0 Then
                c.Value = ""
            Else
                On Error Resume Next: c.Value = Empty
            End If
    End Select
End Sub


Public Function EnumFichiers(chemin As String) As String()
' Fonction permettant de lister dans un tableau le nom de tout les fichiers � partir d'un dossier
' ---------------- Variable ----------------
' Chemin : Chemin d'acc�s du dossier (string)

Dim TableauFichiers() As String
Dim fichier As String
Dim i As Integer

'On compl�te le chemin s'il ne l'est pas
If Right(chemin, 1) <> "\" Then chemin = chemin & "\"

'On r�cup�re les fichiers
fichier = Dir(chemin)

'Boucle sur les fichiers du dossier
Do While (Len(fichier) > 0)

    i = i + 1

    ReDim Preserve TableauFichiers(1 To i)

    TableauFichiers(i) = fichier

    fichier = Dir()

Loop

'Retourne le tableau avec les noms de fichiers
EnumFichiers = TableauFichiers()

End Function
Public Sub nettoyer_donn�es()
'Proc�dure qui permet de nettoyer la liste de DTR renseign�s dans le recueil de donn�es
Dim derlig As Integer

If MsgBox("Voulez-vous effacer la PBS renseign�e ?", vbYesNo, "Suppression") = vbNo Then Exit Sub

derlig = Range("B10000").End(xlUp).Row

Range("A2:F" & derlig).ClearContents
Range("A2:F" & derlig).Borders.LineStyle = xlLineStyleNone
End Sub

Public Function Arraypos(tableau As Variant, valeur As String)
'Fonction permettant le num�ro de la colonne d'un tableau 1D
Dim i

    Arraypos = -1
    
    For i = LBound(tableau) To UBound(tableau, 2)
    
        If tableau(0, i) = valeur Then
            Arraypos = i
            Exit For
        End If
        
    Next i


End Function
Function RecordsetToArray(rs As ADODB.Recordset) As Variant
'Fonction permettant de stocker un recordset dans un tableau

    Dim arr() As Variant
    Dim rows As Long
    Dim cols As Long
    Dim i As Long
    Dim j As Long
    
    ' Aller au d�but du Recordset
    rs.MoveFirst
    
    ' Compter le nombre de lignes
    rows = 0
    Do Until rs.EOF
        rows = rows + 1
        rs.MoveNext
    Loop
    
    ' Aller au d�but du Recordset encore une fois
    rs.MoveFirst
    
    ' Nombre de colonnes
    cols = rs.Fields.count
    
    ' Redimensionner l'array pour contenir les donn�es
    ReDim arr(0 To rows, 0 To cols - 1)
    
    ' Ajouter les en-t�tes de colonnes
    For j = 0 To cols - 1
        arr(0, j) = rs.Fields(j).Name
    Next j
    
    ' Copier les donn�es dans l'array
    i = 1
    Do Until rs.EOF
        For j = 0 To cols - 1
            arr(i, j) = rs.Fields(j).Value
        Next j
        rs.MoveNext
        i = i + 1
    Loop
    
    ' Retourner l'array
    RecordsetToArray = arr
End Function
Public Function TauxDevise(AbrevDevIn As String, AbrevDevOut As String)
'Fonction permettant de r�cup�rer les taux de change entre diff�rentes devises

Dim ApiUrl
Dim sResponse As String

TauxDevise = ""
ApiUrl = "https://open.er-api.com/v6/latest/" & AbrevDevIn

With CreateObject("MSXML2.XMLHTTP")
        .Open "GET", ApiUrl, False
        .Send
        sResponse = .responseText
End With

With CreateObject("vbscript.regexp")
    .Pattern = AbrevDevOut & Chr(34) & ":([\d.]+)"
    If .test(sResponse) Then TauxDevise = .Execute(sResponse)(0).Submatches(0)
End With

End Function
Public Function Test_Connexion(chaineconnexion As String) As Variant
'Fonction permettant de tester si la connexion avec la BDD fonctionne

    Dim result(1) As Variant
    Dim conn As Object
    
    On Error GoTo ErrHandler
    
    ' Cr�ation de la connexion
    Set conn = CreateObject("ADODB.Connection")
    conn.ConnectionString = Replace(chaineconnexion, "UserProfile", Environ$("UserProfile"))
    conn.Open
    
    ' Si la connexion est ouverte
    If conn.State = 1 Then
        result(0) = "Connexion r�ussie � la base de donn�es."
        result(1) = vbInformation
    Else
        result(0) = "�chec de la connexion � la base de donn�es."
        result(1) = vbCritical
    End If
    
    ' Fermer la connexion
    conn.Close
    Set conn = Nothing
    
    Test_Connexion = result
    Exit Function

ErrHandler:
    result(0) = "Erreur lors de la connexion : " & Err.description
    result(1) = vbCritical
    Test_Connexion = result
End Function
Public Sub SynchroBDD()
'Proc�dure de synchronisation One Drive de la BDD

cheminbase = Replace(CONNEXION_EXCEL_BDD_PROD, "UserProfile", Environ$("UserProfile"))
cheminbase = Split(cheminbase, "Data Source=")(1)
cheminbase = Split(cheminbase, ";")(0)

If Not FichierDisponible(cheminbase) Then
    MsgBox "Le fichier n'est pas disponible localement. Tentative de synchronisation...", vbExclamation
    ' Ouvre et referme le fichier pour forcer OneDrive
    Workbooks.Open cheminbase, ReadOnly:=True
    ActiveWorkbook.Close False
Else
    ' Ouvre et referme le fichier pour forcer OneDrive
    Workbooks.Open cheminbase, ReadOnly:=True
    ActiveWorkbook.Close False
End If
    

End Sub
Public Function FichierDisponible(ByVal cheminFichier As String) As Boolean
    On Error Resume Next
    Dim attr As Long
    
    attr = GetAttr(cheminFichier)
    If Err.Number <> 0 Then
        FichierDisponible = False
    Else
        ' V�rifie si le fichier est marqu� comme offline (OneDrive non synchronis�)
        FichierDisponible = ((attr And vbOffline) = 0)
    End If
End Function
Public Function RechercheValeurdansRS(rs As ADODB.Recordset, valeur As String, field As String) As Boolean
' Fonction qui v�rifie la pr�sence d'une valeur dans un Recordset en utilisant Filter

On Error GoTo ErrHandler

' V�rifier si le recordset est valide et ouvert
If rs Is Nothing Or rs.State <> adStateOpen Then
    RechercheValeurdansRS = False
    Exit Function
End If

' V�rifier si le recordset contient des donn�es
If rs.EOF Then
    RechercheValeurdansRS = False
    Exit Function
End If

' Appliquer le filtre (gestion des apostrophes dans valeur)
rs.Filter = "[" & field & "] = '" & Replace(valeur, "'", "''") & "'"

' Si le recordset filtr� n'est pas vide, valeur trouv�e
If Not rs.EOF Then
    RechercheValeurdansRS = True
Else
    RechercheValeurdansRS = False
End If

' R�initialiser le filtre
rs.Filter = ""

Exit Function

ErrHandler:
    RechercheValeurdansRS = False
End Function
Public Function DictToArray2D(ByVal dict As Object) As Variant
'Fonction qui permet de cr�er un tableau � partir d'un dictionnaire
Dim n As Long
Dim arr, ks, vs As Variant

n = dict.count
If n = 0 Then
    DictToArray2D = Empty
    Exit Function
End If

ReDim arr(1 To n, 1 To 2)

ks = dict.Keys
vs = dict.Items

For i = 1 To n
    arr(i, 1) = ks(i - 1)
    arr(i, 2) = vs(i - 1)
Next i

DictToArray2D = arr

End Function

Public Function RecupProjetsList() As Variant
'Fonction permettant de r�cup�rer la liste des projets de la BDD en tableau 1D
Set REQSEL = CreateObject("ADODB.Recordset")

Dim i As Integer
Dim RecupProjetsListe() As Variant
        
    'Appel fonction requete BDD
    STRSEL = "SELECT [Nom projet_PROJET SLI] FROM [Projet_SLI$] ORDER BY [Nom projet_PROJET SLI] ASC"
    Call open_REQSEL(STRSEL)
    
    ReDim RecupProjetsListe(REQSEL.RecordCount)
    For i = 0 To REQSEL.RecordCount - 1
        RecupProjetsListe(i) = REQSEL.Fields(0).Value
        REQSEL.MoveNext
    Next i

    RecupProjetsList = RecupProjetsListe
    
End Function
Public Sub Open_CONN(chainedeconnexion As String)
'Fonction qui ouvre la connexion avec la BDD Excel

Set Base = CreateObject("ADODB.Connection")
'Set Base = New ADODB.Connection
Base.Open chainedeconnexion

    'Set Base = Nothing

        With Base
            If .State = 1 Then
                .Close
            End If
            .ConnectionString = chainedeconnexion
            .Open
            '.CommandTimeout = STANDARD_TIMEOUT
        End With
        Exit Sub

End Sub
Public Sub Close_CONN(chainedeconnexion As String)
'Fonction qui ferme la connexion � la BDD Excel
    With Base
        If .State = 1 Then
            .Close
        End If
    End With
    Set Base = Nothing
    
    Exit Sub
 
End Sub
Public Function open_REQSEL(STRSEL_text As String)
'Fonction d'ex�cution d'une requ�te SQL sur la base de donn�es

    Set REQSEL = CreateObject("ADODB.Recordset")

    REQSEL.CursorLocation = adUseClient
    REQSEL.Open STRSEL_text, Base, , , adCmdText

End Function
Public Function ExecNonQuery(ByVal sqlText As String, Optional ByVal commandTimeout As Long = 30) As Long
' Ex�cute une requ�te d'action (INSERT / UPDATE / DELETE) sur la connexion Base.
' Retourne le nombre de lignes affect�es. L�ve une erreur en cas d'�chec.
    
    Dim rowsAffected As Long
    
    If Base Is Nothing Then
        Err.Raise vbObjectError + 513, "ExecNonQuery", "La connexion 'Base' n'est pas initialis�e."
    End If
    
    On Error GoTo EH
    Base.commandTimeout = commandTimeout
    Base.Execute sqlText, rowsAffected, adCmdText
    
    ExecNonQuery = rowsAffected
    Exit Function
    
EH:
    ' Remonte l'erreur avec le SQL pour diagnostic
    Err.Raise Err.Number, "ExecNonQuery", "Erreur sur ExecNonQuery: " & Err.description & vbCrLf & "SQL: " & sqlText
End Function

Public Function open_REQSEL1(STRSEL_text As String, targetRecordset As String)
'Fonction d'ex�cution d'une requ�te SQL sur la base de donn�es
'Elle affecte le r�sultat � REQSEL1 ou REQSEL2 selon le param�tre targetRecordset

    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")

    rs.CursorLocation = adUseClient
    rs.Open STRSEL_text, Base, , , adCmdText

    ' Affectation dynamique
    Select Case targetRecordset
        Case "REQSEL1"
            Set REQSEL1 = rs
        Case "REQSEL2"
            Set REQSEL2 = rs
        Case Else
            Set REQSEL = rs ' fallback si besoin
    End Select

End Function


' Construit une requ�te UPDATE pour une feuille Excel (ACE/Jet SQL),
' avec SET g�n�r� � partir d'un array (1..n,1..2) et de valeurs lues sur un UserForm.
' - feuille          : nom de la feuille Excel (onglet) cible
' - arr              : tableau (1..n,1..2) o� [i,1]=Nom de colonne ; [i,2]=Nom du contr�le du UserForm
' - uf               : instance du UserForm (ex. Me dans le code du UserForm)
' - projectValue     : valeur de la cl� pour WHERE ([Nom projet_PROJET SLI] = projectValue)
' - iStart, iEnd     : bornes de boucle pour la partie SET (par d�faut 1..23)
Public Function BuildUpdateSqlFromArray( _
    ByVal arr As Variant, _
    ByVal uf As Object, _
    ByVal projectValue As String, _
    Optional ByVal iStart As Long = 1, _
    Optional ByVal iEnd As Long = 23) As String

    Dim i As Long
    Dim colName As String
    Dim ctrlName As String
    Dim val As Variant
    Dim setParts As Collection
    Dim sql As String
    Dim setClause As String
 
    ' V�rifs simples
    If IsEmpty(arr) Then
        BuildUpdateSqlFromArray = vbNullString
        Exit Function
    End If
    If iStart < LBound(arr, 1) Then iStart = LBound(arr, 1)
    If iEnd > UBound(arr, 1) Then iEnd = UBound(arr, 1)
    If iEnd < iStart Then
        BuildUpdateSqlFromArray = vbNullString
        Exit Function
    End If

    Set setParts = New Collection

    ' Boucle SET : de iStart � iEnd (ici 1..23)
    For i = iStart To iEnd
        colName = CStr(arr(i, 1))
        ctrlName = CStr(arr(i, 2))

        ' Si le contr�le existe dans le UserForm, on r�cup�re sa valeur ; sinon on met Null
        If ControlExists(uf, ctrlName) Then
            val = CallByName(uf, ctrlName, VbGet) ' uf.Gen_NomProjet -> renvoie l'objet contr�le
            ' Si c'est un contr�le MSForms (TextBox/ComboBox), prendre .Value
            On Error Resume Next
            val = val.Value
            On Error GoTo 0
        Else
            val = Null
        End If

        ' Construction " [Col] = <valeur SQL> "
        setParts.Add "[" & colName & "] = " & ToAceSqlLiteral(val)
    Next i

    ' Concat�ne les �l�ments du SET par ", "
    setClause = JoinCollection(setParts, ", ")

    ' WHERE sur la cl� "Nom projet_PROJET SLI"
    sql = "UPDATE [Projet_SLI$] " & _
          "SET " & setClause & " " & _
          "WHERE [Nom projet_PROJET SLI] = " & ToAceSqlLiteral(projectValue)

    BuildUpdateSqlFromArray = sql
End Function
' Construit une requ�te INSERT INTO pour une feuille Excel (ACE/Jet SQL),
' � partir d'un array (1..n,1..2) et des valeurs lues sur un UserForm.
' - feuille      : nom de l�onglet cible (ex. "Projets_SLI")
' - arr          : tableau (1..n,1..2) o� [i,1] = colonne, [i,2] = nom du contr�le du UserForm
' - uf           : instance du UserForm (ex. Me)
' - iStart,iEnd  : bornes de boucle pour l�INSERT (par d�faut 1..23)
' - skipNulls    : si True, ignore les colonnes dont la valeur est Null => INSERT partiel
Public Function BuildInsertSqlFromArray( _
    ByVal arr As Variant, _
    ByVal uf As Object, _
    Optional ByVal iStart As Long = 1, _
    Optional ByVal iEnd As Long = 23, _
    Optional ByVal skipNulls As Boolean = False) As String

    Dim i As Long
    Dim colName As String
    Dim ctrlName As String
    Dim val As Variant
    
    Dim cols As New Collection
    Dim vals As New Collection
    
    ' V�rifs & bornes
    If IsEmpty(arr) Then
        BuildInsertSqlFromArray = vbNullString
        Exit Function
    End If
    If iStart < LBound(arr, 1) Then iStart = LBound(arr, 1)
    If iEnd > UBound(arr, 1) Then iEnd = UBound(arr, 1)
    If iEnd < iStart Then
        BuildInsertSqlFromArray = vbNullString
        Exit Function
    End If

    ' Constitue la liste de colonnes et valeurs
    For i = iStart To iEnd
        colName = CStr(arr(i, 1))
        ctrlName = CStr(arr(i, 2))

        If ControlExists(uf, ctrlName) Then
            val = CallByName(uf, ctrlName, VbGet)
            On Error Resume Next
            val = val.Value   ' r�cup�re .Value pour TextBox/ComboBox/etc.
            On Error GoTo 0
        Else
            val = Null
        End If

        If Not (skipNulls And IsNull(val)) Then
            cols.Add "[" & colName & "]"
            vals.Add ToAceSqlLiteral(val)
        End If
    Next i

    If cols.count = 0 Then
        BuildInsertSqlFromArray = vbNullString
        Exit Function
    End If

    BuildInsertSqlFromArray = _
        "INSERT INTO [Projet_SLI$] (" & JoinCollection(cols, ", ") & ") " & _
        "VALUES (" & JoinCollection(vals, ", ") & ")"
End Function

' === Helpers ===

' Retourne True si un membre portant ce nom existe dans l'objet (UserForm)
Public Function ControlExists(ByVal uf As Object, ByVal ctrlName As String) As Boolean
    On Error Resume Next
    Dim tmp As Object
    Set tmp = CallByName(uf, ctrlName, VbGet)
    ControlExists = (Err.Number = 0 And Not tmp Is Nothing)
    On Error GoTo 0
End Function

' Convertit une valeur VBA en litt�ral SQL ACE/Jet :
' - Null -> NULL
' - Date -> #yyyy-mm-dd#
' - Num�rique -> 123.45
' - Texte -> 'texte' avec apostrophes �chapp�es
Public Function ToAceSqlLiteral(ByVal v As Variant) As String
    If IsNull(v) Then
        ToAceSqlLiteral = "NULL"
    ElseIf IsDate(v) Then
        ToAceSqlLiteral = "#" & Format$(CDate(v), "yyyy-mm-dd") & "#"
    ElseIf IsNumeric(v) Then
        ToAceSqlLiteral = CStr(v)
    Else
        ToAceSqlLiteral = "'" & Replace(CStr(v), "'", "''") & "'"
    End If
End Function

' Joint les items d'une Collection par un s�parateur (�quivalent � Join() pour Collections)
Public Function JoinCollection(ByVal c As Collection, ByVal sep As String) As String

Dim i As Long, s As String
For i = 1 To c.count
    If i > 1 Then s = s & sep
    s = s & CStr(c.item(i))
Next i
JoinCollection = s

End Function





