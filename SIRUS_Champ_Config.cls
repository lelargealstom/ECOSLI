VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SIRUS_Champ_Config"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' =========================================================================
' Classe : SIRUS_Champ_Config
' Definition d'un champ du config SIRUS = une ligne de la feuille
' SIRUS_Config. 13 proprietes alignees sur le brief v0.2 section 8.
'
' Instanciation : Set def = New SIRUS_Champ_Config
' Stockage : Collection (compatible For Each, contrairement aux UDT).
' =========================================================================
Option Explicit

Public Table As String                ' Produit / InfoATM / TachesATM / EMMST / ProjetSLI / ProjetASL
Public Champ As String                ' libelle logique affiche
Public ColonneBdD As String           ' nom interne / colonne dans la table BdD
Public Ordre As Long                  ' ordre d'affichage dans la feuille generee
Public TypeChamp As String            ' Texte / Nombre / Date / Liste / Booleen
Public ListeSource As String          ' cle vers Listes si Liste/Booleen
Public Obligatoire As String          ' Non / Oui / Conditionnel
Public ConditionObligatoire As String ' expression evaluable
Public Source As String               ' Saisie / Auto / Herite (Auto/Herite ==> grise)
Public SourceHeritage As String       ' origine si Herite (ex. Produit.NodeType)
Public PoidsMeQ As Double             ' poids du champ dans le taux de completude
Public RegleAlerte As String          ' expression de coherence
Public Commentaire As String          ' aide a la saisie
