Attribute VB_Name = "V_Nom_ADODB_Type"
'--------------------------------------------------------------------------------------------------------
'
'                                       Generateur Livrables SLI
'                                       Module G_Nom_ADODB_Type
'                                       Developpement : LE LARGE Thomas (LGM/Alstom)
'                                       Module permettant de tester la structure de la BDD requetee
'
'
'-----------------------------------------------------------------------------------------------------------
Sub ListTables()
    Dim cn As Object
    Dim rs As ADODB.Recordset
    Dim Base As String
    Dim userpath As String
    Dim wb As Workbook
    
    ' Initialize UserProfile for this example, replace with actual value
    userpath = Environ$("UserProfile")
    
    Base = userpath & "\Alstom\Svcs_DIS_Fr_VBN - Documents\01_Metiers\01_ISR\007 - Outils metier ISR\01. BDD SLI\11_Livraison de l'OUTIL excel BDD\BDD excel original V04.xlsb"
    ' Ouvrir le fichier
    'Set wb = Workbooks.Open(Base)
    Set cn = CreateObject("ADODB.Connection")
    ' Sauvegarder et fermer le fichier
    'wb.Close SaveChanges:=True ' or False si vous ne souhaitez pas enregistrer les modifications
    
    ' Pause pour s'assurer que le fichier est sauvegarde et ferme (Optionnel)
    Application.Wait (Now + TimeValue("0:00:02"))
    '--- Connection ---
    With cn
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .ConnectionString = "Data Source=" & Base & ";Extended Properties='Excel 12.0 Macro;HDR=Yes;IMEX=1'"
        .Open
    End With
    '-----------------
    
    Set rs = cn.OpenSchema(adSchemaTables)
    
    Do While Not rs.EOF
        Debug.Print rs!TABLE_NAME & " (" & rs!TABLE_TYPE & ")"
        rs.MoveNext
    Loop
    
    '--- Fermeture connexion ---
    rs.Close
    cn.Close
    Set rs = Nothing
    Set cn = Nothing
End Sub






