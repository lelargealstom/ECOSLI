Public Function UpdateSharePointList(dtr As String)
Set REQSEL = CreateObject("ADODB.Recordset")

'Assignation des variables
STRSEL = "SELECT [Version] FROM [Info ATMC_SLI] WHERE [DTR] = '" & dtr & "';"

'Ouverture de la connexion a la BDD
Call Open_CONN(Replace(CONNEXION_SHAREPOINT_LIST, "ListeSharePoint", INFO_ATM_List))

Call open_REQSEL(STRSEL)

If Not REQSEL.EOF Then
    UpdateSharePointList = REQSEL.Fields(0).Value
Else
    UpdateSharePointList = "Non trouve"
End If
    
' Fermer la connexion
Call Close_CONN(Replace(CONNEXION_SHAREPOINT_LIST, "ListeSharePoint", INFO_ATM_List))

End Function
Public Sub recup_fields()
    Set REQSEL = CreateObject("ADODB.Recordset")

    'Ouverture de la connexion a la BDD
    Call Open_CONN(Replace(CONNEXION_EXCEL_BDD_TRAVAIL, "UserProfile", Environ$("UserProfile")))
    
    STRSEL = "SELECT * FROM [Info_ATM$]"
    
    Call open_REQSEL(STRSEL)
    
    REQSEL.MoveFirst
    
    For i = 0 To REQSEL.Fields.count - 1
        Debug.Print REQSEL.Fields(i).Name
    Next i
    
End Sub

Sub ajoutEnregistrement()

    Dim cn As Object
    Dim fichier, UserProfile As String, feuille As String, strSQL, texte_sql As String
    Dim LaDate As Date
    Dim PrixUnit As Integer
    Dim leNom As String, lePrenom As String
    
    UserProfile = Environ$("UserProfile")
    
    ' Chemin vers le fichier Excel
    fichier = UserProfile & "\Alstom\Svcs_DIS_Fr_VBN - Documents\01_Metiers\01_ISR\007 - Outils metier ISR\01. BDD SLI\11_Livraison de l'OUTIL excel BDD\BDD excel original V03.xlsm"
    feuille = "Projet_SLI"

    ' Donnees a inserer
    LaDate = CDate("26/05/2006")
    leNom = "Le Large"
    lePrenom = "Thomas"
    PrixUnit = 40

    ' Creation de la connexion
    Set cn = CreateObject("ADODB.Connection")
    With cn
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .ConnectionString = "Data Source=" & fichier & ";Extended Properties='Excel 12.0 Xml;HDR=Yes';"
        .Open
    End With
    
    ' Requete SQL de lecture
    texte_sql = "SELECT * FROM [Projet_SLI$]"
    
    ' Execution
    cn.Execute texte_sql
    
    ' Requete SQL d'insertion
    strSQL = "INSERT INTO [" & feuille & "$] ([Date], [Nom], [Prenom], [PrixUnit]) VALUES (#" & Format(LaDate, "yyyy-mm-dd") & "#, '" & leNom & "', '" & lePrenom & "', " & PrixUnit & ")"

    ' Execution
    cn.Execute strSQL

    ' Fermeture
    cn.Close
    Set cn = Nothing

    MsgBox "Ligne inseree avec succes dans le fichier Excel."
End Sub

Sub ModifierValeurProjet()

    Dim cn As Object
    Dim strSQL As String
    Dim fichier As String
    Dim feuille As String

    ' Definir le chemin du fichier et la feuille
    fichier = "C:\Users\201057333\Alstom\Svcs_DIS_Fr_VBN - Documents\01_Metiers\01_ISR\007 - Outils metier ISR\01. BDD SLI\11_Livraison de l'OUTIL excel BDD\BDD excel original V03.xlsm"
    feuille = "Produit"

    ' Creer la connexion
    Set cn = CreateObject("ADODB.Connection")
    With cn
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .ConnectionString = "Data Source=" & fichier & ";Extended Properties='Excel 12.0 Xml;HDR=Yes';"
        .Open
    End With

    ' Requete SQL de mise a jour
    strSQL = "UPDATE [" & feuille & "$] SET [Projet / Offre_PROJET SLI] = 'Projet' WHERE [Nom projet_PROJET SLI] = 'PROJET_TEST'"

    ' Executer la requete
    cn.Execute strSQL

    ' Fermer la connexion
    cn.Close
    Set cn = Nothing

    MsgBox "Mise a jour effectuee avec succes !"

End Sub
Sub Recup_projets()

Dim req As String
Dim cn As Object
Dim rs As ADODB.Recordset
Dim feuille, fichier As String

    ' Definir le chemin du fichier et la feuille
    fichier = "C:\Users\201057333\Alstom\Svcs_DIS_Fr_VBN - Documents\01_Metiers\01_ISR\007 - Outils metier ISR\01. BDD SLI\11_Livraison de l'OUTIL excel BDD\BDD excel original V04.xlsb"
    feuille = "Projet_SLI"

    Set cn = CreateObject("ADODB.Connection")
    With cn
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .ConnectionString = "Data Source=" & fichier & ";Extended Properties='Excel 12.0 Xml;HDR=Yes';"
        .Open
    End With

    
    req = "SELECT [Nom projet_PROJET SLI] FROM [Projet_SLI$]"

    
    Set rs = New ADODB.Recordset
    rs.Open req, cn, adOpenStatic, adLockReadOnly

    If Not rs.EOF Then
        msg = "Resultat de la requete :" & vbCrLf & vbCrLf & "Nombre de projets : " & rs.RecordCount & vbCrLf & vbCrLf
        rs.MoveFirst
        For i = 1 To rs.RecordCount
            msg = msg & rs.Fields(0).Value & vbCrLf
            Debug.Print rs.Fields(0).Value
            rs.MoveNext
        Next i
    Else
        msg = "Aucune donnee trouvee pour la table projet."
    End If
    
    'tableau_test = RecordsetToArray(rs)
    
    ' Afficher le resultat
    MsgBox msg, vbInformation, "Lecture BDD"
    
    ' Fermer la connexion
    cn.Close
    Set cn = Nothing
    
End Sub



