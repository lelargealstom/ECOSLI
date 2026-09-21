Option Compare Text
Option Explicit
Public Sub Nettoyage()
' Module permettant de nettoyer la PBS importe dans le generateur de livrables
' Variables : Colonnes de la feuille 'Providor_Data'

Dim dligne, pligne As Integer
Dim Generateur As Workbook
Dim PBS As Worksheet

'On definit le classeur de generateur de livrables
Set Generateur = ActiveWorkbook
Set PBS = Generateur.Sheets("DATA_ASL")

If MsgBox("Voulez-vous effacer la PBS renseignee ?", vbYesNo, "Suppression") = vbNo Then Exit Sub

dligne = PBS.Cells(PBS.rows.count, 1).End(xlUp).Row
pligne = PBS.Cells.Find(What:="DTR Reference (issue de PBS)", LookIn:=xlValues, lookat:=xlWhole).Row + 1

Range("F" & pligne & ":G" & dligne).ClearContents       'DTR et designations
Range("J" & pligne & ":J" & dligne).ClearContents       'Sous systemes
Range("L" & pligne & ":L" & dligne).ClearContents       'Quantite
Range("T" & pligne & ":U" & dligne).ClearContents       'MTBF Intrinseque et fonctionnel
Range("Z" & pligne & ":Z" & dligne).ClearContents       'Duree vie limitee HARD
Range("AC" & pligne & ":AC" & dligne).ClearContents     'Duree vie limitee SOFT
Range("AN" & pligne & ":AQ" & dligne).ClearContents     'MTBF Intrinseque et fonctionnel

End Sub
