VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Barre_Progression 
   Caption         =   "Veuillez patienter..."
   ClientHeight    =   30
   ClientLeft      =   -1310
   ClientTop       =   -5250
   ClientWidth     =   70
   OleObjectBlob   =   "Barre_Progression.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Barre_Progression"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub afficher()

    Me.Show 0

End Sub
Sub actualiser(taux As Integer)

    barre.Width = barretexte.Width * taux / 100
    barretexte = taux & "%"
    
    If taux = 100 Then Unload Me
    
    DoEvents
    
End Sub

Private Sub barretexte_Click()

End Sub
