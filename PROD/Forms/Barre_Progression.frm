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
