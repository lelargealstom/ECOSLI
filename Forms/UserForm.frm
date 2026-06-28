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

' =========================================================================
' API Windows : bouton "minimiser" natif sur la barre de titre du UserForm
' =========================================================================
' Par defaut, les UserForms VBA sont crees avec WS_CAPTION mais sans
' WS_MINIMIZEBOX -> pas de bouton minimize visible.
' On modifie le style de la fenetre via SetWindowLong pour ajouter
' WS_MINIMIZEBOX. Compatible Excel 32-bit (VBA6) et 64-bit (VBA7).
'
' Comportement utilisateur : clic sur le bouton ? minimise le UserForm
' dans la barre des taches Windows. Clic sur l icone Excel re-affiche.
' Indispensable pour le va-et-vient PBS / UserForm (consultation
' DATA_ASL pendant la generation Generation/Rafraichir).
#If VBA7 Then
    Private Declare PtrSafe Function FindWindow_API Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
    Private Declare PtrSafe Function GetWindowLong_API Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function SetWindowLong_API Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare PtrSafe Function DrawMenuBar_API Lib "user32" Alias "DrawMenuBar" (ByVal hWnd As LongPtr) As Long
#Else
    Private Declare Function FindWindow_API Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function GetWindowLong_API Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
    Private Declare Function SetWindowLong_API Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare Function DrawMenuBar_API Lib "user32" Alias "DrawMenuBar" (ByVal hWnd As Long) As Long
#End If

Private Const GWL_STYLE      As Long = -16
Private Const WS_MINIMIZEBOX As Long = &H20000
Private Const WS_MAXIMIZEBOX As Long = &H10000

' Active les boutons "minimiser" et "maximiser" natifs sur la barre de
' titre du UserForm. A appeler dans UserForm_Initialize. Tolerant : si
' FindWindow echoue (autre langue Office, classe renommee), les boutons
' ne sont pas ajoutes mais aucun crash.
'
' Note : le maximize sans UserForm_Resize laissera du vide a droite/bas
' (les controles ont une position absolue, ne suivent pas). Refonte
' layout adaptatif a programmer dans un second temps si necessaire.
Private Sub UI_ActiverBoutonMinimize()
    #If VBA7 Then
        Dim hWndForm As LongPtr
    #Else
        Dim hWndForm As Long
    #End If
    Dim lStyle As Long

    On Error Resume Next
    hWndForm = FindWindow_API("ThunderDFrame", UserForm.Caption)
    If hWndForm = 0 Then Exit Sub

    lStyle = GetWindowLong_API(hWndForm, GWL_STYLE)
    SetWindowLong_API hWndForm, GWL_STYLE, lStyle Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX
    DrawMenuBar_API hWndForm
    On Error GoTo 0
End Sub

Private Sub ATM_Option_Click()

V1_Option.Visible = True
V1_txt.Visible = True
V2_Option.Visible = True
V2_txt.Visible = True
Version_txt.Visible = True

End Sub
Public Function MsgValidation(Livrable As String) As String
    
Dim missingFields As String

' Verification feuille
If Test_feuille("DATA_ASL") = False Then
    MsgBox "Veuillez selectionner la feuille correspondante a votre PBS."
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
Call G_SIRUS_UserForm.SIRUS_AjouterLigneProduit("Creation")
End Sub

Private Sub AddTache_ATM_Click()

Call SIRUS_AjouterTacheATM

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
'On verifie les donnees renseignees par l'utilisateur
msg = MsgValidation("ATM")

If msg <> "" Then
    MsgBox msg, vbExclamation, "Attention"
    Exit Sub
End If

'Si tout est ok on passe a la suite
barre.Visible = True
barretexte.Visible = True
barreimage.Visible = True
barrearriere.Visible = True

'Ouverture de la connexion a la BDD
Call Open_CONN(Replace(CONNEXION_EXCEL_BDD_TRAVAIL, "UserProfile", Environ("UserProfile")))

'Il faut appeller la macro de generation ATM
Call Main_ATM

End Sub

Private Sub Btn_ChoixFichier_Extr_Click()
'Ouverture du repertoire pour selectionner un fichier ATM
'1. Verification si les options d'extraction sont effectuees
'  a. Choix d'un livrable type
'  b. Choix d'une langue
'  c. Choix d'une version si ATM
'2. Verification des feuilles existantes
'  a. Presence d'un Index
'  b. Presence d'un template selon la langue

' 1. Verification Options Extraction
If UserForm.Extr_Livrables_List.Value = "" Then
    MsgBox "Veuillez selectionner un livrable avant de choisir un fichier."
    Exit Sub
Else
    Select Case UserForm.Extr_Livrables_List.Value
        Case "ATM"    'Langue et Version
            If UserForm.Langue_EN_Extr.Value = False And UserForm.Langue_FR_Extr.Value = False Then
                MsgBox "Veuillez selectionner la langue relative aux fiches ATM."
                Exit Sub
            End If
            If UserForm.ATMV1Extr.Value = False And UserForm.ATMV2Extr.Value = False Then
                MsgBox "Veuillez selectionner la version des fiches ATM."
                Exit Sub
            ElseIf UserForm.ATMV2Extr.Value = True Then
                MsgBox "Extraction des ATM 2.0 indisponible pour le moment"
                Exit Sub
            End If
        Case "EMMST"   'Langue
            If UserForm.Langue_EN_Extr.Value = False And UserForm.Langue_FR_Extr.Value = False Then
                MsgBox "Veuillez selectionner la langue relative aux fiches EMMST."
                Exit Sub
            End If
    End Select
End If

'2. Recuperation des parametres
Select Case UserForm.Langue_EN_Extr.Value
    Case True
        langue = "EN"
    Case False
        langue = "FR"
End Select

'3. Choix du fichier par l'utilisateur
fichier = Application.GetOpenFilename("Fichier Excel, *.xlsx; *.xlsm; *.xlsb", , "Selection de vos fiches " & UserForm.Extr_Livrables_List.Value)
If Not InStr(1, fichier, ".xlsx") >= 1 Or InStr(1, fichier, ".xlsm") >= 1 Then
    MsgBox "Votre fichier n'est pas compatible aux fiches "
    Exit Sub
End If

'4. Recuperation du fichier dans le textbox
lb_nomfichier_extr.Value = fichier

Debug.Print fichier

End Sub
Private Sub Btn_DeleteFichier_Extr_Click()

'On vide l'espace fichier
lb_nomfichier_extr.Value = ""

End Sub
Private Sub Btn_EMMST_Click()
'On verifie les donnees renseignees par l'utilisateur
msg = MsgValidation("EMMST")

If msg <> "" Then
    MsgBox msg, vbExclamation, "Attention"
    Exit Sub
End If

'Si tout est ok on passe a la suite
barre.Visible = True
barretexte.Visible = True
barreimage.Visible = True
barrearriere.Visible = True

'Ouverture de la connexion a la BDD
Call Open_CONN(Replace(CONNEXION_EXCEL_BDD_TRAVAIL, "UserProfile", Environ$("UserProfile")))

'Il faut appeller la macro de generation EMMST
Call Main_EMMST

End Sub
Private Sub Btn_Providor_Click()
'On verifie les donnees renseignees par l'utilisateur
msg = MsgValidation("Providor")

If msg <> "" Then
    MsgBox msg, vbExclamation, "Attention"
    Exit Sub
End If

'Si tout est ok on passe a la suite
barre.Visible = True
barretexte.Visible = True
barreimage.Visible = True
barrearriere.Visible = True

'Ouverture de la connexion a la BDD
Call Open_CONN(Replace(CONNEXION_EXCEL_BDD_TRAVAIL, "UserProfile", Environ$("UserProfile")))

'Il faut appeller la macro de generation Providor - Sans cas de test
Call Main_Providor(False)

End Sub
' Recherche des informations ATM a partir d'un DTR
Private Sub Btn_SearchDTR_ATM_Click()

    Dim dtr As String
    dtr = Trim(UserForm.DTRATM_txt.Value)

    If dtr = "" Then
        MsgBox "Veuillez renseigner un DTR avant de lancer la recherche.", vbExclamation, "SIRUS a Recherche DTR"
        Exit Sub
    End If

    ' -- Requete sur la table Produit ----------------------
    Set REQSEL = CreateObject("ADODB.Recordset")

    STRSEL = "SELECT * FROM [Produit$] WHERE [DTR] = '" & Replace(dtr, "'", "''") & "'"

    Call open_REQSEL(STRSEL)

    If REQSEL.EOF Then
        MsgBox "DTR '" & dtr & "' introuvable dans la table Produit. Veuillez creer le DTR avant de creer les informations relatives aux fiches ATM", vbExclamation, "SIRUS a Recherche DTR"
        Exit Sub
    End If

    ' -- Remplir les champs --------------------------------
    Dim NodeType As String
    NodeType = ""

    ' Node Type
    If Not IsNull(REQSEL.Fields("Node type Produit (LRU/SRUe)_PRODUIT").Value) Then
        NodeType = Trim(REQSEL.Fields("Node type Produit (LRU/SRUe)_PRODUIT").Value)
    End If
    UserForm.NodeTypeATM_txt.Value = NodeType

    ' Niveau Qualif a 3 si LRU ou LRUNR, sinon 3+
    Select Case UCase(NodeType)
        Case "LRU", "LRUNR"
            UserForm.NivQualifATM_txt.Value = "3"
        Case Else
            UserForm.NivQualifATM_txt.Value = "3+"
    End Select

    ' Famille generique
    If Not IsNull(REQSEL.Fields("Famille generique_PRODUIT").Value) Then
        UserForm.FamPrdATM_txt.Value = Trim(REQSEL.Fields("Famille generique_PRODUIT").Value)
    Else
        UserForm.FamPrdATM_txt.Value = ""
    End If

    ' Reference Plan (Drawing Number)
    If Not IsNull(REQSEL.Fields("Drawing number").Value) Then
        UserForm.RefPlanATM_txt.Value = Trim(REQSEL.Fields("Drawing number").Value)
    Else
        UserForm.RefPlanATM_txt.Value = ""
    End If

End Sub
Private Sub Btn_TNR_Click()

'On verifie les donnees renseignees par l'utilisateur
msg = MsgValidation("Providor")

If msg <> "" Then
    MsgBox msg, vbExclamation, "Attention"
    Exit Sub
End If

MsgBox "Le test de non regression n'est pas encore disponible"
Exit Sub
'Ouverture de la connexion a la BDD
Call Open_CONN(Replace(CONNEXION_EXCEL_BDD_TRAVAIL, "UserProfile", Environ$("UserProfile")))

'Il faut appeller la macro de generation Providor - Sans cas de test
Call Main_Providor(True)

End Sub
Private Sub BtnATMAP_Click()
    
    Dim url As String
    
    url = "https://apps.powerapps.com/play/e/default-0d993ad3-fa73-421a-b129-1fe5590103f3/a/a0fc7db8-c62b-4166-a6a9-8be62fc64486?tenantId=0d993ad3-fa73-421a-b129-1fe5590103f3&hint=60f199af-f900-4abc-8581-6d0553f5645e&sourcetime=1772119906992"
    
    ' Ouvre le lien dans le navigateur par defaut
    ThisWorkbook.FollowHyperlink Address:=url, NewWindow:=True

End Sub

Private Sub CommandButton4_Click()

End Sub

Private Sub Connexion_BDD_Button_Click()
    Dim result As Variant
    
    result = Test_Connexion(CONNEXION_EXCEL_BDD_TRAVAIL)
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

Call SIRUS_SupprimerTacheATM

End Sub
Private Sub Extr_Livrables_List_Change()

'On affiche les versions si le choix est ATM
If Extr_Livrables_List.Value = "ATM" Then
    UserForm.Version_Frame_Extr.Visible = True
End If

End Sub
Private Sub Btn_Extraction_Click()
'On verifie bien que la selection d'un fichier est faite par l'utilisateur
If lb_nomfichier_extr.Value = "" Then
    MsgBox "Veuillez selectionner un fichier avant de lancer l'extraction", vbInformation
    Exit Sub
End If

'Si tout est ok on passe a la suite
barre.Visible = True
barretexte.Visible = True
barreimage.Visible = True
barrearriere.Visible = True

'On recupere le nom du fichier pour l'ouvrir
Call Main_Extraction

End Sub
Private Sub FamilyProduct_List_Change()

Set REQSEL = CreateObject("ADODB.Recordset")

'Info_Projet = DataProjet_SLI(Project_List.Value)
'Requete table Projets SLI
STRSEL = "SELECT * FROM [Projet_SLI$] WHERE [Nom projet_PROJET SLI] = '" & Project_List.Value & "'"
Call open_REQSEL(STRSEL)

' Si aucun enregistrement trouve, on sort
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
' Si la page selectionna est Generation, il faut verifier que l'utilisateur est dans un template du generateur de livrable
If ECOSLI_MultiPage.Value = 1 Then
    If Test_feuille("DATA_ASL") = False Then
        MsgBox "Impossible de generer a partir de ce fichier"
        Exit Sub
    End If
End If
        
' Si la page selectionna est Information, il faut dire a l'utilisateur de selectionner un projet dans la liste si rien n'est selectionne
If Project_List.Value = "" And ECOSLI_MultiPage.Value = 0 Then
    MsgBox "Veuillez selectionner un projet dans la liste afin d'afficher les informations generiques", vbInformation, "Aucun projet selectionne."
End If

End Sub

Private Sub ListView_Produit_BeforeLabelEdit(Cancel As Integer)



End Sub

' =========================================================================
' Bouton Rafraichir Data PBS (a cote de Project_List)
' =========================================================================
' Cable au CommandButton Btn_Rafraichir_DataPBS du UserForm (a ajouter
' manuellement dans le VBE Designer). Lance le Refresh DATA_ASL via
' ADODB sur le classeur projet actif. Le projet courant est lu depuis
' Project_List dans la sub publique SIRUS_RafraichirDATA_ASL.
Private Sub Btn_Rafraichir_DataPBS_Click()
    Call G_PBS_Refresh.SIRUS_RafraichirDATA_ASL
End Sub

Private Sub Project_List_Change()
' Auto-remplit les controles SLI a partir de la BdD lorsque l utilisateur
' choisit un projet. Indispensable avant d ouvrir un ticket ATMAP (la
' Zone Envoi SIRUS lit Project_List pour pre-remplir Sujet/Description).
'
' Robuste a Base = Nothing : les flows SIRUS appellent Close_CONN apres
' leurs requetes (G_SIRUS_Regles.ChargerProjetsSLI, G_SIRUS_Generation.
' PreRemplirDepuisBdD, etc.), ce qui ferme la connexion ouverte par
' UserForm_Initialize. On la re-ouvre ici si besoin.

Dim chk As MSForms.CheckBox
Dim val As Variant
Dim key As Variant
Dim ChampControl As String

On Error GoTo Sortie

' -- Garde 1 : projet vide -> rien a faire ----------------
If LenB(Trim$(CStr(Project_List.Value))) = 0 Then Exit Sub

' -- Garde 2 : connexion BdD ouverte ----------------------
'    Si Base a ete ferme par un Close_CONN (SIRUS et autres), on
'    re-ouvre. Open_CONN gere idempotemment l etat ferme/ouvert.
If Base Is Nothing Then
    Call Open_CONN(Replace(CONNEXION_EXCEL_BDD_TRAVAIL, "UserProfile", Environ$("UserProfile")))
ElseIf Base.State <> 1 Then  ' adStateOpen = 1
    Call Open_CONN(Replace(CONNEXION_EXCEL_BDD_TRAVAIL, "UserProfile", Environ$("UserProfile")))
End If

' -- Requete projet ---------------------------------------
STRSEL = "SELECT * FROM [Projet_SLI$] WHERE [Nom projet_PROJET SLI] = '" & Replace(CStr(Project_List.Value), "'", "''") & "'"
Call open_REQSEL(STRSEL)

If REQSEL.RecordCount = 0 Then
    MsgBox "Projet introuvable dans la base.", vbExclamation
    Exit Sub
End If

' -- Initialisation du mapping si non fait ----------------
If Champs_ProjetsSLI Is Nothing Then Call GetChampsProjetsSLI

REQSEL.MoveFirst

' -- Remplissage des controles ----------------------------
For Each key In Champs_ProjetsSLI.Keys
    ChampControl = Champs_ProjetsSLI(key)
    If Me.Controls(ChampControl) Is Nothing Then
        Debug.Print "Controle introuvable pour : " & key
    Else
        If IsNull(REQSEL.Fields(key).Value) Then
            Me.Controls(ChampControl).Value = ""
        Else
            val = REQSEL.Fields(key).Value
            If TypeName(Me.Controls(ChampControl)) = "CheckBox" Then
                Set chk = Me.Controls(ChampControl)
                val = UCase(Trim(val))
                Select Case val
                    Case "OUI"
                        chk.Value = True
                        chk.BackColor = RGB(0, 200, 0)
                    Case "NON"
                        chk.Value = False
                        chk.BackColor = RGB(200, 0, 0)
                    Case Else  ' NA, vide, autres -> gris
                        chk.Value = False
                        chk.BackColor = RGB(180, 180, 180)
                End Select
                chk.Enabled = True
            Else
                If InStr(ChampControl, "Date") > 1 Then
                    Me.Controls(ChampControl).Value = Left(val, 11)
                Else
                    Me.Controls(ChampControl).Value = val
                End If
            End If
        End If
    End If
Next key

' Synchroniser la description ATMAP si la Zone Envoi est deja visible
On Error Resume Next
Call G_SIRUS_UserForm.SIRUS_UI_ProjetListChange
On Error GoTo 0

Exit Sub

Sortie:
    MsgBox "Erreur Project_List_Change : " & Err.Description & vbCrLf & vbCrLf & "Verifier que la BdD TRAVAIL est accessible.", vbExclamation, "SIRUS"
End Sub
Private Sub Refresh_BDD_Button_Click()
'Ouvre la base de donnees et ferme direct
'Forcer la synchronisation One Drive
Call SynchroBDD

End Sub
Private Sub RangATM_list_Change()

Dim val As String
val = Trim(UserForm.RangATM_list.Value)

'Bloquer la selection des lignes de groupe
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
'Initialisation de la fenetre ECOSLI

'On force l'ouverture de la base de donnees pour que la synchronisation soit optimale

'On cree la connexion a la Base de donnees en ouvrant la fenetre
Call Open_CONN(Replace(CONNEXION_EXCEL_BDD_TRAVAIL, "UserProfile", Environ$("UserProfile")))
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
Set ws_list = ThisWorkbook.Sheets("LISTES_DATA_CHOIX")

dlignefrs = ws_list.Cells(ws_list.rows.count, "C").End(xlUp).Row
dlignenode = ws_list.Cells(ws_list.rows.count, "B").End(xlUp).Row
dlignerang = ws_list.Cells(ws_list.rows.count, "H").End(xlUp).Row

'On rend invisible les barres de progression et tableau de valeur
barre.Visible = False
barretexte.Visible = False
barreimage.Visible = False
barrearriere.Visible = False
Article_label.Visible = False
Article_List.Visible = False

'Taille du userform
UserForm.Height = 570
UserForm.Width = 810
ECOSLI_MultiPage.MultiRow = False

'Ajout de la liste de projets avec une fonction appele
listprojets = RecupProjetsList()
For i = LBound(listprojets) To UBound(listprojets)
    If listprojets(i) <> "" Then
        Project_List.AddItem listprojets(i)
        'Debug.Print listprojets(i) & " : " & Abrev_Projet(CStr(listprojets(i)))
    Else: Exit For
    End If
Next i

'Pre-selection automatique : si le classeur actif a une cellule nommee
'DATA_ASL_PROJET non vide (template ECOSLI_DATA_ASL existant rempli par
'un Refresh precedent), on pre-positionne Project_List dessus.
'ActiveWorkbook (pas ThisWorkbook qui pointerait sur le .xlam).
Dim sProjetAuto As String
sProjetAuto = ""
On Error Resume Next
sProjetAuto = Trim$(CStr(ActiveWorkbook.Names("DATA_ASL_PROJET").RefersToRange.Value))
On Error GoTo 0
If LenB(sProjetAuto) > 0 Then
    On Error Resume Next
    Project_List.Value = sProjetAuto
    On Error GoTo 0
End If

'Definir la liste des gates dans un tableau
gates = Array("BA", "GO Order", "IGR", "PGR", "CGR", "FEI", "IQA")
family = Array("PACIS", "IOB", "APC", "DAS", "EVR", "SRS", "Lecteur Badge", "GSM Repeater")
revisions = Array( "1A", "2A", "3A", "A", "1B", "2B", "3B", "B", "1C", "2C", "3C", "C", "1D", "2D", "3D", "D", "1E", "2E", "3E", "E", "1F", "2F", "3F")
livrables = Array("ATM", "EMMST")
HARD_SOFT = Array("HARD", "SOFT")
langue = Array("FR", "EN")
    
' Ajouter chaque element a la liste deroulante
For i = LBound(gates) To UBound(gates)
    Gate_List.AddItem gates(i)
Next i

For i = LBound(revisions) To UBound(revisions)
    Revision_List.AddItem revisions(i)
Next i
For i = LBound(family) To UBound(family)
    FamilyProduct_List.AddItem family(i)
Next i

'SIRUS - configuration ListViews + alimentation Generation depuis DATA_ASL + BdD
Call G_SIRUS_UserForm.SIRUS_InitGeneration

' SIRUS Phase v2 : defaults Op=Creer, Frames caches, ComboBox Priorite peuplee
Call G_SIRUS_UserForm.SIRUS_PhaseInit_v2

' UX : ajouter le bouton minimize natif sur la barre de titre du UserForm
' (API user32.SetWindowLong). UX style fenetre Windows / navigateur.
Call UI_ActiverBoutonMinimize

' SIRUS a etat initial NavTab type de demande
Call SIRUS_SetType("Creation")

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