Public Sub Appel_Usf()
Attribute Appel_Usf.VB_ProcData.VB_Invoke_Func = "w\n14"
'Module qui affiche le formulaire
    UserForm.Show vbModeless
End Sub


' ==========================================================================
' PROCEDURE: ActiverRaccourciCtrlW
' ==========================================================================
' A lancer manuellement via Developpeur > Macros si le raccourci Ctrl+W ne
' fonctionne plus apres un import de module (perte de l'attribut procedure
' VB_ProcData.VB_Invoke_Func).
'
' NB : normalement appelee automatiquement a chaque chargement du complement
' par A_1_ModEvents.Auto_Open. Cette version manuelle est un filet de securite
' quand on modifie W_AppelUsf.bas sans redemarrer Excel.
Public Sub ActiverRaccourciCtrlW()
    Application.OnKey "^w", "Appel_Usf"
    MsgBox "Raccourci Ctrl+W reactive pour ouvrir le UserForm ECOSLI.", vbInformation
End Sub