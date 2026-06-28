Attribute VB_Name = "G_SIRUS_Listes"
'--------------------------------------------------------------------------------------------------------
'
'                                       Generateur Livrables SLI
'                                       Module G_SIRUS_Listes
'                                       Developpement : LE LARGE Thomas (LGM/Alstom)
'                                       SIRUS - Phase 0 ter : sources des listes deroulantes
'                                       pour les champs TypeChamp = "Liste" ou "Booleen"
'                                       du config SIRUS. Renvoyees par SIRUS_GetListe(cle).
'                                       Decision : listes gerees en code (pas en feuille).
'                                       Juin 2026
'
'--------------------------------------------------------------------------------------------------------
Option Explicit
Option Compare Text

' =========================================================================
' CLES DES LISTES (alignees avec la colonne ListeSource du config)
' =========================================================================
Public Const SIRUS_LIST_NODETYPE        As String = "NodeType"
Public Const SIRUS_LIST_HARDSOFT        As String = "HardSoft"
Public Const SIRUS_LIST_LANGUE          As String = "Langue"
Public Const SIRUS_LIST_BOOLEEN         As String = "Booleen"
Public Const SIRUS_LIST_TYPETRAIN       As String = "TypeTrain"
Public Const SIRUS_LIST_DEVISE          As String = "Devise"
Public Const SIRUS_LIST_NIVMAINTENANCE  As String = "NiveauMaintenance"
Public Const SIRUS_LIST_MOTIFNA         As String = "MotifNA"

' =========================================================================
' SIRUS_GetListe
' =========================================================================
' Renvoie les valeurs d'une liste en tableau 1D (LBound=0).
' Si la cle est inconnue, renvoie un tableau vide (UBound=-1) + Debug.Print.
' Aucune erreur bloquante : l'appelant peut iterer en For Each / For i (la
' boucle ne s'execute pas sur tableau vide).
Public Function SIRUS_GetListe(ByVal sNomListe As String) As Variant
    Select Case sNomListe

        Case SIRUS_LIST_NODETYPE
            SIRUS_GetListe = Array("LRU", "SRU", "LRUNR", "SRUNR", "KIT", "VISSERIE", "ASSEMBLY")

        Case SIRUS_LIST_HARDSOFT
            SIRUS_GetListe = Array("HARD", "SOFT")

        Case SIRUS_LIST_LANGUE
            SIRUS_GetListe = Array("FR", "EN")

        Case SIRUS_LIST_BOOLEEN
            SIRUS_GetListe = Array("Oui", "Non")

        Case SIRUS_LIST_TYPETRAIN
            SIRUS_GetListe = Array("CORADIA", "AVELIA", "METROPOLIS", "XTRAPOLIS", "TRAX", "INNOVIA", "CITADIS", "ADESSIA", "Voitures", "SIG")

        Case SIRUS_LIST_DEVISE
            SIRUS_GetListe = Array("EUR", "USD", "CAD", "JPY", "CHF", "GBP")

        Case SIRUS_LIST_NIVMAINTENANCE
            SIRUS_GetListe = Array("3", "3/3+", "3+")

        Case SIRUS_LIST_MOTIFNA
            SIRUS_GetListe = Array("NA", "NA MECA", "NA KIT")

        Case Else
            Debug.Print "SIRUS_GetListe : nom de liste inconnu '" & sNomListe & "'."
            SIRUS_GetListe = Array()
    End Select
End Function

' =========================================================================
' SIRUS_ListeEnChaine
' =========================================================================
' Renvoie une chaine "v1,v2,v3" prete a etre posee dans une Data Validation
' Excel (xlValidateList, Formula1:=...). Limite Excel : 255 caracteres.
Public Function SIRUS_ListeEnChaine(ByVal sNomListe As String, Optional ByVal sSep As String = ",") As String
    Dim arr As Variant
    Dim i As Long
    Dim s As String

    arr = SIRUS_GetListe(sNomListe)
    If UBound(arr) < LBound(arr) Then
        SIRUS_ListeEnChaine = ""
        Exit Function
    End If

    For i = LBound(arr) To UBound(arr)
        If LenB(s) > 0 Then s = s & sSep
        s = s & CStr(arr(i))
    Next i
    SIRUS_ListeEnChaine = s
End Function

' =========================================================================
' Test_GetListe : trace les 7 listes dans la fenetre Immediate.
' =========================================================================
Public Sub Test_GetListe()
    Dim cles As Variant
    Dim k As Variant
    Dim arr As Variant
    Dim i As Long
    Dim s As String

    cles = Array(SIRUS_LIST_NODETYPE, SIRUS_LIST_HARDSOFT, SIRUS_LIST_LANGUE, SIRUS_LIST_BOOLEEN, SIRUS_LIST_TYPETRAIN, SIRUS_LIST_DEVISE, SIRUS_LIST_NIVMAINTENANCE, SIRUS_LIST_MOTIFNA)

    Debug.Print "=== Test_GetListe ==="
    For Each k In cles
        arr = SIRUS_GetListe(CStr(k))
        s = ""
        For i = LBound(arr) To UBound(arr)
            If LenB(s) > 0 Then s = s & ", "
            s = s & CStr(arr(i))
        Next i
        Debug.Print "  " & k & " (" & (UBound(arr) - LBound(arr) + 1) & ") : " & s
    Next k
End Sub
