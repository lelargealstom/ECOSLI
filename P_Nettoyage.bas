Attribute VB_Name = "P_Nettoyage"
Option Compare Text
Option Explicit
Public Sub Nettoyage()
' Module permettant de nettoyer la PBS importé dans le générateur de livrables
' Variables : Colonnes de la feuille 'Providor_Data'

Dim dligne, pligne As Integer
Dim Generateur As Workbook
Dim PBS As Worksheet

'On définit le classeur de générateur de livrables
Set Generateur = ActiveWorkbook
Set PBS = Generateur.Sheets("DATA_ASL")

If MsgBox("Voulez-vous effacer la PBS renseignée ?", vbYesNo, "Suppression") = vbNo Then Exit Sub

dligne = PBS.Cells(PBS.rows.count, 1).End(xlUp).Row
pligne = PBS.Cells.Find(What:="DTR Reference (issue de PBS)", LookIn:=xlValues, lookat:=xlWhole).Row + 1

Range("F" & pligne & ":G" & dligne).ClearContents       'DTR et désignations
Range("J" & pligne & ":J" & dligne).ClearContents       'Sous systèmes
Range("L" & pligne & ":L" & dligne).ClearContents       'Quantité
Range("T" & pligne & ":U" & dligne).ClearContents       'MTBF Intrinsèque et fonctionnel
Range("Z" & pligne & ":Z" & dligne).ClearContents       'Duree vie limitee HARD
Range("AC" & pligne & ":AC" & dligne).ClearContents     'Duree vie limitee SOFT
Range("AN" & pligne & ":AQ" & dligne).ClearContents     'MTBF Intrinsèque et fonctionnel

End Sub
