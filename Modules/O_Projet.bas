Attribute VB_Name = "O_Projet"
Option Compare Text
Option Explicit
Public Sub projet()

' Procedure qui permet de creer un projet en BDD
' Il faut verifier si le projet existe en BDD avant de l'importer
' Developpe par Thomas LE LARGE
' Date : 09/01/2025
Dim BDD, Generateur As Workbook
Dim userpath As String
Dim Projets_SLI, Feuille_Projet As Worksheet
Dim Projet_liste As Variant
Dim new_CP As String
Dim Answer, Answer2 As String
Dim Projet_tab As ListObject

'On definit le classeur de generateur de livrables
Set Generateur = ActiveWorkbook
Set Feuille_Projet = Generateur.Sheets("DATA_SLI")

'On recupere le numero de l'utilisateur
userpath = Environ$("UserProfile")

' On ouvre la BDD avec le chemin correspondant
Application.DisplayAlerts = False
Set BDD = Workbooks.Open(userpath & "\Alstom\Svcs_DIS_Fr_VBN - Documents\01_Metiers\01_ISR\007 - Outils metier ISR\01. BDD SLI\11.OUTIL excel BDD\BDD excel original V03.xlsm", UpdateLinks:=1)
Application.DisplayAlerts = True

' On recupere les colonnes utiles pour ajouter la nouvelle ligne
Set Projets_SLI = BDD.Sheets("Projet (SLI)")
Set Projet_tab = Projets_SLI.Range("Projets_SLI").ListObject
Projet_liste = Projets_SLI.Range("Projets_SLI[Nom projet_PROJETSLI]")

' On recupere le nom du projet renseigne par l'utilisateur en haut de la feuille
new_CP = Feuille_Projet.Cells(2, 2)

' On regarde dans ce tableau si le code projet renseigne par l'utilisateur appartient a cette liste
If Contains_str(new_CP, Projet_liste) > 0 Then
    
    ' On recupere la reponse de l'utilisateur
    Answer = MsgBox("Ce projet existe deja en BD. " & Chr(10) & " Voulez-vous remplacez les donnees actuelles ?", vbYesNoCancel)
    
    Select Case Answer
        
        Case vbCancel
            Exit Sub
        
        Case vbYes
        ' Il faudrait afficher les donnees actuelles de la BDD pour l'utilisateur afin de l'informer du projet actuel
        ' Il pourra choisir de remplacer les donnees ou de maintenir les donnees actuelles
        
        Case vbNo
            MsgBox "Aucun remplacement dans la BDD pour ce projet"
            
    End Select
    
Else
    ' On demande a l'utilisateur s'il souhaite importer ce nouveau projet en BDD
    Answer2 = MsgBox("Souhaitez-vous importer ce projet en BDD ?" & Chr(10) & "Nombre de projet actuel : " & UBound(Projet_liste), vbYesNo)
    
    Select Case Answer2
        
        Case vbYes
            ' On importe les donnees rentrees par l'utilisateur dans tout les champs qu'il pouvait renseigne
            Projets_SLI.ListObjects("Projets_SLI").ListRows.Add

        Case vbNo
            ' On sort de la macro pour que l'utilisateur puisse importer toutes les donnees necessaires avant de reexecuter la macro
            MsgBox "Veuillez remplir toutes les donnees necessaires avant import en BDD"
            
        Exit Sub
        
    End Select
    
End If

End Sub
