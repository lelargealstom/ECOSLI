Option Explicit

' =========================================================================
' Sub Appel_Usf - Point d'entree pour afficher le UserForm ECOSLI
' =========================================================================
' Cable au raccourci Ctrl+W via Application.OnKey dans
' A_1_ModEvents.Auto_Open (qui s execute au chargement de l add-in).
' On n utilise PAS l attribut VB_ProcData.VB_Invoke_Func (raccourci type
' "w\n14") car cet attribut est fragile au re-import des .bas dans le
' VBE et peut etre interprete comme erreur de syntaxe.
Public Sub Appel_Usf()
    UserForm.Show vbModeless
End Sub
