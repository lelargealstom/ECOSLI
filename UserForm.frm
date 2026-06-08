VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm 
   Caption         =   "UserForm"
   ClientHeight    =   7170
   ClientLeft      =   -9330.001
   ClientTop       =   -38840
   ClientWidth     =   16120
   OleObjectBlob   =   "UserForm.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "UserForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public chemin As String
Public userpath
Private Sub ATM_Option_Click()

V1_Option.Visible = True
V1_txt.Visible = True
V2_Option.Visible = True
V2_txt.Visible = True
Version_txt.Visible = True

End Sub
Public Function MsgValidation(Livrable As String) As String
    
Dim missingFields As String

' V�rification feuille
If Test_feuille("DATA_ASL") = False Then
    MsgBox "Veuillez s�lectionner la feuille correspondante � votre PBS."
    Unload UserForm
    Exit Function
End If

Select Case Livrable
    Case "Providor"
        If Project_List.Value = "" Then missingFields = missingFields & "Projet / "
        If Gate_List.Value = "" Then missingFields = missingFields & "Gate / "
        If UserForm.Langue_FR.Value = False And Langue_EN.Value = False Then missingFields = missingFields & "Langue / "
        
        If missingFields <> "" Then
            MsgValidation = "Veuillez remplir les champs obligatoires : " & Left(missingFields, Len(missingFields) - 3)
        Else
            MsgValidation = ""
        End If
        
    Case "ATM"
        If Project_List.Value = "" Then missingFields = missingFields & "Projet / "
        If Gate_List.Value = "" Then missingFields = missingFields & "Gate / "
        If V1_Option.Value = False And V2_Option.Value = False Then missingFields = missingFields & "Version / "
        If UserForm.Langue_FR.Value = False And Langue_EN.Value = False Then missingFields = missingFields & "Langue / "
            
        If missingFields <> "" Then
            MsgValidation = "Veuillez remplir les champs obligatoires : " & Left(missingFields, Len(missingFields) - 3)
        Else
            MsgValidation = ""
        End If
    
    Case "EMMST"
        If Project_List.Value = "" Then missingFields = missingFields & "Projet / "
        If Gate_List.Value = "" Then missingFields = missingFields & "Gate / "
        If UserForm.Langue_FR.Value = False And Langue_EN.Value = False Then missingFields = missingFields & "Langue / "
        
        If missingFields <> "" Then
            MsgValidation = "Veuillez remplir les champs obligatoires : " & Left(missingFields, Len(missingFields) - 3)
        Else
            MsgValidation = ""
        End If
        
End Select
    
End Function

Private Sub AddLigne_Produit_Click()
Call G_SIRIUS_UserForm.SIRIUS_AjouterLigneProduit("Cr�ation")
End Sub

Private Sub AddTache_ATM_Click()

Call SIRIUS_AjouterTacheATM

End Sub

Private Sub Analyser_Produit_Click()

'Hide / Unhide controls
UserForm.ImgWorkflow_Analyse.Visible = True
UserForm.ImgWorkflow_Saisie.Visible = False
UserForm.Workflow_MultiPage.Visible = False

End Sub

Private Sub ATMV1Extr_Click()

If ATMV1Extr.Value = True Then
    ATMV2Extr.Value = False
End If

End Sub

Private Sub ATMV2Extr_Click()

If ATMV2Extr.Value = True Then
    ATMV1Extr.Value = False
End If

End Sub

Private Sub Btn_ATM_Click()
'On v�rifie les donn�es renseign�es par l'utilisateur
msg = MsgValidation("ATM")

If msg <> "" Then
    MsgBox msg, vbExclamation, "Attention"
    Exit Sub
End If

'Si tout est ok on passe � la suite
barre.Visible = True
barretexte.Visible = True
barreimage.Visible = True
barrearriere.Visible = True

'Ouverture de la connexion � la BDD
Call Open_CONN(Replace(CONNEXION_EXCEL_BDD_PROD, "UserProfile", Environ("UserProfile")))

'Il faut appeller la macro de g�n�ration ATM
Call Main_ATM

End Sub

Private Sub Btn_ChoixFichier_Extr_Click()
'Ouverture du r�pertoire pour s�lectionner un fichier ATM
'1. V�rification si les options d'extraction sont effectu�es
'  a. Choix d'un livrable type
'  b. Choix d'une langue
'  c. Choix d'une version si ATM
'2. V�rification des feuilles existantes
'  a. Pr�sence d'un Index
'  b. Pr�sence d'un template selon la langue

' 1. V�rification Options Extraction
If UserForm.Extr_Livrables_List.Value = "" Then
    MsgBox "Veuillez s�lectionner un livrable avant de choisir un fichier."
    Exit Sub
Else
    Select Case UserForm.Extr_Livrables_List.Value
        Case "ATM"    'Langue et Version
            If UserForm.Langue_EN_Extr.Value = False And UserForm.Langue_FR_Extr.Value = False Then
                MsgBox "Veuillez s�lectionner la langue relative aux fiches ATM."
                Exit Sub
            End If
            If UserForm.ATMV1Extr.Value = False And UserForm.ATMV2Extr.Value = False Then
                MsgBox "Veuillez s�lectionner la version des fiches ATM."
                Exit Sub
            ElseIf UserForm.ATMV2Extr.Value = True Then
                MsgBox "Extraction des ATM 2.0 indisponible pour le moment"
                Exit Sub
            End If
        Case "EMMST"   'Langue
            If UserForm.Langue_EN_Extr.Value = False And UserForm.Langue_FR_Extr.Value = False Then
                MsgBox "Veuillez s�lectionner la langue relative aux fiches EMMST."
                Exit Sub
            End If
    End Select
End If

'2. R�cup�ration des param�tres
Select Case UserForm.Langue_EN_Extr.Value
    Case True
        langue = "EN"
    Case False
        langue = "FR"
End Select

'3. Choix du fichier par l'utilisateur
fichier = Application.GetOpenFilename("Fichier Excel, *.xlsx; *.xlsm; *.xlsb", , "S�lection de vos fiches " & UserForm.Extr_Livrables_List.Value)
If Not InStr(1, fichier, ".xlsx") >= 1 Or InStr(1, fichier, ".xlsm") >= 1 Then
    MsgBox "Votre fichier n'est pas compatible aux fiches "
    Exit Sub
End If

'4. R�cup�ration du fichier dans le textbox
lb_nomfichier_extr.Value = fichier

Debug.Print fichier

End Sub
Private Sub Btn_DeleteFichier_Extr_Click()

'On vide l'espace fichier
lb_nomfichier_extr.Value = ""

End Sub
Private Sub Btn_EMMST_Click()
'On v�rifie les donn�es renseign�es par l'utilisateur
msg = MsgValidation("EMMST")

If msg <> "" Then
    MsgBox msg, vbExclamation, "Attention"
    Exit Sub
End If

'Si tout est ok on passe � la suite
barre.Visible = True
barretexte.Visible = True
barreimage.Visible = True
barrearriere.Visible = True

'Ouverture de la connexion � la BDD
Call Open_CONN(Replace(CONNEXION_EXCEL_BDD_PROD, "UserProfile", Environ$("UserProfile")))

'Il faut appeller la macro de g�n�ration EMMST
Call Main_EMMST

End Sub
Private Sub Btn_Providor_Click()
'On v�rifie les donn�es renseign�es par l'utilisateur
msg = MsgValidation("Providor")

If msg <> "" Then
    MsgBox msg, vbExclamation, "Attention"
    Exit Sub
End If

'Si tout est ok on passe � la suite
barre.Visible = True
barretexte.Visible = True
barreimage.Visible = True
barrearriere.Visible = True

'Ouverture de la connexion � la BDD
Call Open_CONN(Replace(CONNEXION_EXCEL_BDD_PROD, "UserProfile", Environ$("UserProfile")))

'Il faut appeller la macro de g�n�ration Providor - Sans cas de test
Call Main_Providor(False)

End Sub
' Recherche des informations ATM � partir d'un DTR
Private Sub Btn_SearchDTR_ATM_Click()

    Dim dtr As String
    dtr = Trim(UserForm.DTRATM_txt.Value)

    If dtr = "" Then
        MsgBox "Veuillez renseigner un DTR avant de lancer la recherche.", _
               vbExclamation, "SIRIUS � Recherche DTR"
        Exit Sub
    End If

    ' -- Requ�te sur la table Produit ----------------------
    Set REQSEL = CreateObject("ADODB.Recordset")

    STRSEL = "SELECT * FROM [Produit$] WHERE [DTR] = '" & _
             Replace(dtr, "'", "''") & "'"

    Call open_REQSEL(STRSEL)

    If REQSEL.EOF Then
        MsgBox "DTR '" & dtr & "' introuvable dans la table Produit. Veuillez cr�er le DTR avant de cr�er les informations relatives aux fiches ATM", _
               vbExclamation, "SIRIUS � Recherche DTR"
        Exit Sub
    End If

    ' -- Remplir les champs --------------------------------
    Dim NodeType As String
    NodeType = ""

    ' Node Type
    If Not IsNull(REQSEL.Fields("Node type Produit (LRU/SRU�)_PRODUIT").Value) Then
        NodeType = Trim(REQSEL.Fields("Node type Produit (LRU/SRU�)_PRODUIT").Value)
    End If
    UserForm.NodeTypeATM_txt.Value = NodeType

    ' Niveau Qualif � 3 si LRU ou LRUNR, sinon 3+
    Select Case UCase(NodeType)
        Case "LRU", "LRUNR"
            UserForm.NivQualifATM_txt.Value = "3"
        Case Else
            UserForm.NivQualifATM_txt.Value = "3+"
    End Select

    ' Famille g�n�rique
    If Not IsNull(REQSEL.Fields("Famille g�n�rique_PRODUIT").Value) Then
        UserForm.FamPrdATM_txt.Value = _
            Trim(REQSEL.Fields("Famille g�n�rique_PRODUIT").Value)
    Else
        UserForm.FamPrdATM_txt.Value = ""
    End If

    ' R�f�rence Plan (Drawing Number)
    If Not IsNull(REQSEL.Fields("Drawing number").Value) Then
        UserForm.RefPlanATM_txt.Value = _
            Trim(REQSEL.Fields("Drawing number").Value)
    Else
        UserForm.RefPlanATM_txt.Value = ""
    End If

End Sub
Private Sub Btn_TNR_Click()

'On v�rifie les donn�es renseign�es par l'utilisateur
msg = MsgValidation("Providor")

If msg <> "" Then
    MsgBox msg, vbExclamation, "Attention"
    Exit Sub
End If

MsgBox "Le test de non r�gression n'est pas encore disponible"
Exit Sub
'Ouverture de la connexion � la BDD
Call Open_CONN(Replace(CONNEXION_EXCEL_BDD_PROD, "UserProfile", Environ$("UserProfile")))

'Il faut appeller la macro de g�n�ration Providor - Sans cas de test
Call Main_Providor(True)

End Sub
Private Sub BtnATMAP_Click()
    
    Dim url As String
    
    url = "https://apps.powerapps.com/play/e/default-0d993ad3-fa73-421a-b129-1fe5590103f3/a/a0fc7db8-c62b-4166-a6a9-8be62fc64486?tenantId=0d993ad3-fa73-421a-b129-1fe5590103f3&hint=60f199af-f900-4abc-8581-6d0553f5645e&sourcetime=1772119906992"
    
    ' Ouvre le lien dans le navigateur par d�faut
    ThisWorkbook.FollowHyperlink Address:=url, NewWindow:=True

End Sub

Private Sub CommandButton4_Click()

End Sub

Private Sub Connexion_BDD_Button_Click()
    Dim result As Variant
    
    result = Test_Connexion(CONNEXION_EXCEL_BDD_PROD)
    MsgBox result(0), result(1), "Test connexion BDD"
End Sub
Private Sub EMMST_Option_Click()
V1_Option.Visible = False
V1_txt.Visible = False
V2_Option.Visible = False
V2_txt.Visible = False
Version_txt.Visible = False
End Sub
Private Sub actualiser(taux As Integer)

    barre.Width = barretexte.Width * taux / 100
    barretexte = taux & "%"
    
    If taux = 100 Then Unload Me
    
    DoEvents
    
End Sub
Private Sub DelLigne_Tache_Click()

Call SIRIUS_SupprimerTacheATM

End Sub
Private Sub Extr_Livrables_List_Change()

'On affiche les versions si le choix est ATM
If Extr_Livrables_List.Value = "ATM" Then
    UserForm.Version_Frame_Extr.Visible = True
End If

End Sub
Private Sub Btn_Extraction_Click()
'On v�rifie bien que la s�lection d'un fichier est faite par l'utilisateur
If lb_nomfichier_extr.Value = "" Then
    MsgBox "Veuillez s�lectionner un fichier avant de lancer l'extraction", vbInformation
    Exit Sub
End If

'Si tout est ok on passe � la suite
barre.Visible = True
barretexte.Visible = True
barreimage.Visible = True
barrearriere.Visible = True

'On r�cup�re le nom du fichier pour l'ouvrir
Call Main_Extraction

End Sub
Private Sub FamilyProduct_List_Change()

Set REQSEL = CreateObject("ADODB.Recordset")

'Info_Projet = DataProjet_SLI(Project_List.Value)
'Requete table Projets SLI
STRSEL = "SELECT * FROM [Projet_SLI$] WHERE [Nom projet_PROJET SLI] = '" & Project_List.Value & "'"
Call open_REQSEL(STRSEL)

' Si aucun enregistrement trouv�, on sort
If REQSEL.RecordCount = 0 Then
    MsgBox "Projet introuvable dans la base.", vbExclamation
    Exit Sub
End If

' Initialisation du mapping si non fait
If Champs_Dates Is Nothing Then Call GetChampsDates

' Positionner sur le premier enregistrement
REQSEL.MoveFirst

'On remplit ensuite les dates en rapport avec le fournisseur selectionne

End Sub

Private Sub Langue_EN_Extr_Click()

If Langue_EN_Extr.Value = True Then
    Langue_FR_Extr.Value = False
End If

End Sub

Private Sub Langue_FR_Extr_Click()

If Langue_FR_Extr.Value = True Then
    Langue_EN_Extr.Value = False
End If

End Sub
Private Sub ListBox_Produit_Click()

If ListBox_Produit.ListIndex = 0 Then
    ListBox_Produit.ListIndex = -1
End If

End Sub
Private Sub ECOSLI_MultiPage_Change()
' Si la page s�lectionn� est G�n�ration, il faut v�rifier que l'utilisateur est dans un template du g�n�rateur de livrable
If ECOSLI_MultiPage.Value = 1 Then
    If Test_feuille("DATA_ASL") = False Then
        MsgBox "Impossible de g�n�rer � partir de ce fichier"
        Exit Sub
    End If
End If
        
' Si la page s�lectionn� est Information, il faut dire � l'utilisateur de s�lectionner un projet dans la liste si rien n'est s�lectionn�
If Project_List.Value = "" And ECOSLI_MultiPage.Value = 0 Then
    MsgBox "Veuillez s�lectionner un projet dans la liste afin d'afficher les informations g�n�riques", vbInformation, "Aucun projet s�lectionn�."
End If

End Sub

Private Sub ListView_Produit_BeforeLabelEdit(Cancel As Integer)



End Sub

Private Sub Project_List_Change()

Set REQSEL = CreateObject("ADODB.Recordset")

'D�s que le projet est modifi� par l'utilisateur il faut r�cup�rer toutes les informations sur le projet de la BDD
Dim chk As MSForms.CheckBox

'Info_Projet = DataProjet_SLI(Project_List.Value)
'Requete table Projets SLI
STRSEL = "SELECT * FROM [Projet_SLI$] WHERE [Nom projet_PROJET SLI] = '" & Project_List.Value & "'"
Call open_REQSEL(STRSEL)

' Si aucun enregistrement trouv�, on sort
If REQSEL.RecordCount = 0 Then
    MsgBox "Projet introuvable dans la base.", vbExclamation
    Exit Sub
End If

' Initialisation du mapping si non fait
If Champs_ProjetsSLI Is Nothing Then Call GetChampsProjetsSLI

' Positionner sur le premier enregistrement
REQSEL.MoveFirst

' Remplissage automatique des champs du formulaire
For Each key In Champs_ProjetsSLI.Keys
    ChampControl = Champs_ProjetsSLI(key)
    If Me.Controls(ChampControl) Is Nothing Then
        Debug.Print "Contr�le introuvable pour : " & key
    Else
        ' G�rer les valeurs Null
        If IsNull(REQSEL.Fields(key).Value) Then
            Me.Controls(ChampControl).Value = ""
        Else
            val = REQSEL.Fields(key).Value
            If TypeName(Me.Controls(ChampControl)) = "CheckBox" Then
                Set chk = Me.Controls(ChampControl)
                
                val = UCase(Trim(val))
                
                ' Gestion des �tats
                Select Case val
                    Case "OUI"
                        chk.Value = True
                        chk.Enabled = True
                        chk.BackColor = RGB(0, 200, 0) ' Vert
                    Case "NON"
                        chk.Value = False
                        chk.Enabled = True
                        chk.BackColor = RGB(200, 0, 0) ' Rouge
                    Case "NA", ""
                        chk.Value = False
                        chk.Enabled = True
                        chk.BackColor = RGB(180, 180, 180) ' Gris
                    Case Else
                        chk.Value = False
                        chk.Enabled = True
                        chk.BackColor = RGB(180, 180, 180) ' Gris par d�faut
                End Select
            Else
                'Pour tout les autres controles (TExtBox, Combobox...)
                If InStr(ChampControl, "Date") > 1 Then
                    Me.Controls(ChampControl).Value = Left(val, 11)
                Else
                    Me.Controls(ChampControl).Value = val
                End If
            End If
        End If
    End If
Next key

End Sub
Private Sub Refresh_BDD_Button_Click()
'Ouvre la base de donn�es et ferme direct
'Forcer la synchronisation One Drive
Call SynchroBDD

End Sub
Private Sub RangATM_list_Change()

Dim val As String
val = Trim(UserForm.RangATM_list.Value)

'Bloquer la s�lection des lignes de groupe
If InStr(val, "---") >= 1 Or val = "" Then
    UserForm.RangATM_list.Value = ""
End If

End Sub
Private Sub Test_Connexion_SIRUS_Click()

Dim result As Variant
    
result = Test_Connexion(CONNEXION_EXCEL_BDD_TRAVAIL)
MsgBox result(0), result(1), "Test connexion BDD"

End Sub
Public Sub UserForm_Initialize()
'Initialisation de la fen�tre ECOSLI

'On force l'ouverture de la base de donn�es pour que la synchronisation soit optimale

'On cr�e la connexion � la Base de donn�es en ouvrant la fen�tre
Call Open_CONN(Replace(CONNEXION_EXCEL_BDD_PROD, "UserProfile", Environ$("UserProfile")))
Dim wb As Workbook
'Dim ws As Worksheet
Dim ts As Object
Dim listprojets, tab_1, colIndex As Variant
Dim i, j, dlignenode, dlignefrs, dlignerang As Integer
Dim leftpos As Single
Dim lbl As MSForms.Label
Dim gates, revisions, livrables, family, HARD_SOFT, langue As Variant
Dim ws_list As Worksheet

Set wb = ThisWorkbook
Set ws = wb.ActiveSheet
Set ws_list = ThisWorkbook.Sheets("Liste")

dlignefrs = ws_list.Cells(ws.rows.count, "C").End(xlUp).Row
dlignenode = ws_list.Cells(ws.rows.count, "B").End(xlUp).Row
dlignerang = ws_list.Cells(ws.rows.count, "H").End(xlUp).Row

'On rend invisible les barres de progression et tableau de valeur
barre.Visible = False
barretexte.Visible = False
barreimage.Visible = False
barrearriere.Visible = False
Article_label.Visible = False
Article_List.Visible = False
Version_Frame_Extr.Visible = False


'Taille du userform
UserForm.Height = 570
UserForm.Width = 810
ECOSLI_MultiPage.MultiRow = False

'Ajout de la liste de projets avec une fonction appel�
listprojets = RecupProjetsList()
For i = LBound(listprojets) To UBound(listprojets)
    If listprojets(i) <> "" Then
        Project_List.AddItem listprojets(i)
        'Debug.Print listprojets(i) & " : " & Abrev_Projet(CStr(listprojets(i)))
    Else: Exit For
    End If
Next i

'D�finir la liste des gates dans un tableau
gates = Array("BA", "GO Order", "IGR", "PGR", "CGR", "FEI", "IQA")
family = Array("PACIS", "IOB", "APC", "DAS", "EVR", "SRS", "Lecteur Badge", "GSM Repeater")
revisions = Array( _
    "1A", "2A", "3A", "A", _
    "1B", "2B", "3B", "B", _
    "1C", "2C", "3C", "C", _
    "1D", "2D", "3D", "D", _
    "1E", "2E", "3E", "E", _
    "1F", "2F", "3F")
livrables = Array("ATM", "EMMST")
HARD_SOFT = Array("HARD", "SOFT")
langue = Array("FR", "EN")
    
' Ajouter chaque �l�ment � la liste d�roulante
For i = LBound(gates) To UBound(gates)
    Gate_List.AddItem gates(i)
Next i

For i = LBound(revisions) To UBound(revisions)
    Revision_List.AddItem revisions(i)
Next i
For i = LBound(livrables) To UBound(livrables)
    Extr_Livrables_List.AddItem livrables(i)
Next i
For i = LBound(family) To UBound(family)
    FamilyProduct_List.AddItem family(i)
Next i

'SIRUS
Call G_SIRIUS_UserForm.SIRIUS_ConfigListViewProduit
Call G_SIRIUS_UserForm.SIRIUS_ConfigListViewTaches
Call G_SIRIUS_UserForm.SIRIUS_ConfigListViewATM
Call G_SIRIUS_UserForm.SIRIUS_ConfigListViewGeneration

' SIRIUS - Alimentation PBS dans ListView_Generation a partir de DATA_ASL du classeur actif
G_SIRIUS_PBS.SIRIUS_LookupDTR_Active = True
Call G_SIRIUS_PBS.SIRIUS_RemplirListViewGeneration

' SIRIUS � �tat initial NavTab type de demande
Call SIRIUS_SetType("Cr�ation")

' Ajout des listes de valeurs
For i = 2 To dlignefrs
    FRS_list.AddItem ws_list.Cells(i, 3)
Next i

For i = 2 To dlignenode
    NodeType_list.AddItem ws_list.Cells(i, 2)
Next i

For i = LBound(HARD_SOFT) To UBound(HARD_SOFT)
    TypeATM_list.AddItem HARD_SOFT(i)
Next i

For i = LBound(langue) To UBound(langue)
    LangueATM_list.AddItem langue(i)
Next i

For i = 2 To dlignerang
    RangATM_list.AddItem ws_list.Cells(i, 8)
Next i

End Sub


