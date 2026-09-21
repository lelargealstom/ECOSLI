VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAppEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ==========================================
' Classe: clsAppEvents
' Description: Gestionnaire d'evenements Application
' Surveille les changements dans les feuilles de calcul
' ==========================================
Option Explicit

Private WithEvents App As Application
Attribute App.VB_VarHelpID = -1

' ==========================================
' Initialisation
' ==========================================
Public Sub Initialize()
    Set App = Application
End Sub

' ==========================================
' Termination
' ==========================================
Public Sub Terminate()
    Set App = Nothing
End Sub
