Attribute VB_Name = "G_SIRUS_Config"
'--------------------------------------------------------------------------------------------------------
'
'                                       Generateur Livrables SLI
'                                       Module G_SIRUS_Config
'                                       Developpement : LE LARGE Thomas (LGM/Alstom)
'                                       SIRUS - Phase 0 : feuille de configuration pilote
'                                       (definitions de champ par table x operation)
'                                       cf. brief SIRUS v0.2 sections 6 et 8.
'                                       Juin 2026
'
'--------------------------------------------------------------------------------------------------------
Option Explicit
Option Compare Text

' =========================================================================
' CONSTANTES
' =========================================================================
Public Const SIRUS_CONFIG_SHEET    As String = "SIRUS_Config"
Public Const SIRUS_CONFIG_LISTOBJ  As String = "tblSIRUSConfig"

' Noms de table (valeurs de la colonne "Table" du config)
Public Const SIRUS_TBL_PRODUIT     As String = "Produit"
Public Const SIRUS_TBL_INFOATM     As String = "InfoATM"
Public Const SIRUS_TBL_TACHESATM   As String = "TachesATM"
Public Const SIRUS_TBL_EMMST       As String = "EMMST"
Public Const SIRUS_TBL_PROJETSLI   As String = "ProjetSLI"
Public Const SIRUS_TBL_PROJETASL   As String = "ProjetASL"

' =========================================================================
' SIRUS_ColonneClePourTable
' =========================================================================
' Renvoie le numero de colonne contenant la cle metier d une feuille SIRUS
' selon la table. La cle est ce qui identifie unitairement la ligne pour
' l affichage ListView Resultats et le scan ligneMax.
'   ProjetASL : col 4 (DTR ; col 2 = Projet, col 3 = Sous Systeme)
'   Produit   : col 3 (DTR ; col 2 = Famille Generique)
'   ProjetSLI : col 2 (Nom projet)
'   InfoATM, TachesATM, EMMST : col 2 (DTR)
' Source unique de verite : modifier ici si on ajoute des champs avant la cle.
Public Function SIRUS_ColonneClePourTable(ByVal sTable As String) As Long
    Select Case sTable
        Case "ProjetASL":  SIRUS_ColonneClePourTable = 4
        Case "Produit":    SIRUS_ColonneClePourTable = 3
        Case Else:         SIRUS_ColonneClePourTable = 2
    End Select
End Function

' =========================================================================
' DEFINITION D'UN CHAMP DU CONFIG
' =========================================================================
' Cf. classe SIRUS_Champ_Config.cls (13 proprietes alignees sur le brief
' v0.2 section 8). On a choisi une classe plutot qu'un Type UDT car VBA
' interdit le stockage d'un UDT dans une Collection (impossible a passer
' en Variant), donc le For Each plante.

' =========================================================================
' INITIALISATION DE LA FEUILLE SIRUS_Config
' =========================================================================
' Cree la feuille SIRUS_Config (veryHidden) avec son ListObject si absente,
' et la peuple pour la table Produit (pilote Phase 0).
' Idempotent : si la feuille existe deja et bForcer = False, ne fait rien.
' bForcer = True ==> suppression + recreation (utile pour iteration en dev).
Public Sub SIRUS_InitConfigSheet(Optional ByVal bForcer As Boolean = False)

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim entetes As Variant
    Dim deja As Boolean

    On Error GoTo Sortie

    ' La config vit dans le complement ECOSLI.xlam (ThisWorkbook = .xlam).
    ' Convention : ActiveWorkbook est reserve a la PBS du classeur projet ;
    ' la config interne du complement est bien ThisWorkbook.
    Set wb = ThisWorkbook

    ' -- 1. Verifier si la feuille existe deja ----------------
    deja = False
    For Each ws In wb.Worksheets
        If ws.Name = SIRUS_CONFIG_SHEET Then
            deja = True
            Exit For
        End If
    Next ws

    If deja And Not bForcer Then
        Debug.Print "SIRUS_InitConfigSheet : la feuille " & SIRUS_CONFIG_SHEET & " existe deja. Appeler avec bForcer:=True pour reinitialiser."
        Exit Sub
    End If

    ' -- 2. Creer (ou re-creer) la feuille --------------------
    If deja Then
        Application.DisplayAlerts = False
        wb.Worksheets(SIRUS_CONFIG_SHEET).Visible = xlSheetVisible
        wb.Worksheets(SIRUS_CONFIG_SHEET).Delete
        Application.DisplayAlerts = True
    End If

    Set ws = wb.Worksheets.Add
    ws.Name = SIRUS_CONFIG_SHEET

    ' -- 3. Ecrire l'entete (13 colonnes, brief section 8) ----
    entetes = Array("Table", "Champ", "ColonneBdD", "Ordre", "TypeChamp", "ListeSource", "Obligatoire", "ConditionDescription", "Source", "SourceHeritage", "PoidsMeQ", "RegleAlerte", "Commentaire")
    ws.Range("A1").Resize(1, UBound(entetes) - LBound(entetes) + 1).Value = entetes

    ' -- 4. Creer le ListObject -------------------------------
    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1").Resize(2, UBound(entetes) - LBound(entetes) + 1), , xlYes)
    lo.Name = SIRUS_CONFIG_LISTOBJ
    ' Vider la ligne creee automatiquement par xlYes
    If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.ClearContents

    ' -- 5. Peupler pour les 6 tables (brief v0.2 sections 6.1 a 6.5) --
    Call PeuplerProduit(lo)
    Call PeuplerInfoATM(lo)
    Call PeuplerTachesATM(lo)
    Call PeuplerEMMST(lo)
    Call PeuplerProjetSLI(lo)
    Call PeuplerProjetASL(lo)

    ' Si une ligne vide a survecu en debut, on la supprime
    If lo.ListRows.Count > 0 Then
        If LenB(CStr(lo.ListRows(1).Range(1, 1).Value)) = 0 Then
            lo.ListRows(1).Delete
        End If
    End If

    ' -- 6. Cacher la feuille (veryHidden) -------------------
    ws.Visible = xlSheetVeryHidden

    Debug.Print "SIRUS_InitConfigSheet : feuille " & SIRUS_CONFIG_SHEET & " creee et peuplee pour 6 tables (" & lo.ListRows.Count & " champs au total)."
    Exit Sub

Sortie:
    MsgBox "Erreur dans SIRUS_InitConfigSheet : " & Err.Description, vbCritical
End Sub

' =========================================================================
' PEUPLEMENT DU CONFIG POUR LA TABLE PRODUIT
' =========================================================================
' Champs definis a partir du brief v0.2 section 6.1.
' Obligatoires simples : DTR, Designation manuelle, Node Type, Fournisseur,
' Reference fournisseur, Numero plan, MoQ Spare.
' Conditionnels : MTBF intrinseque / Motif NA (au moins un des deux),
' Prix/Date validite Spare et Repair selon Node Type et presence prix manuel.
Private Sub PeuplerProduit(lo As ListObject)
    ' Refonte 2026-06-10 : nouvel ordre + 25 champs metier alignes sur le
    ' brief Thomas. Champs marques ??_*_a_definir : nom BdD a confirmer
    ' via Test_DiagnostiquerEcartsConfig_Toutes / Test_DiagnostiquerColonnesBdDFeuille.

    Call AjouterChamp(lo, "Produit", "Famille Generique", "??_FamilleGenerique_a_definir", 10, "Texte", "", "Non", "", "Saisie", "", 1, "", "Famille generique du produit (a definir).")

    Call AjouterChamp(lo, "Produit", "DTR", "DTR", 20, "Texte", "", "Oui", "", "Saisie", "", 3, "", "Identifiant unique du produit (cle).")

    Call AjouterChamp(lo, "Produit", "Designation manuelle", "Designation manuelle_PRODUIT", 30, "Texte", "", "Oui", "", "Saisie", "", 2, "", "")

    Call AjouterChamp(lo, "Produit", "Node Type", "Node type Produit (LRU/SRU...)_PRODUIT", 40, "Liste", "NodeType", "Oui", "", "Saisie", "", 2, "", "")

    Call AjouterChamp(lo, "Produit", "Ref RAM Report", "??_RefRAMReport_a_definir", 50, "Texte", "", "Non", "", "Saisie", "", 1, "", "Reference RAM Report (a definir).")

    Call AjouterChamp(lo, "Produit", "Fournisseur", "Fournisseur_PRODUIT", 60, "Texte", "", "Oui", "", "Saisie", "", 1, "", "")

    Call AjouterChamp(lo, "Produit", "Reference fournisseur", "Reference fournisseur_PRODUIT", 70, "Texte", "", "Oui", "", "Saisie", "", 1, "", "")

    Call AjouterChamp(lo, "Produit", "Motif NA", "MotifNonApplicabilite_PRODUIT", 80, "Liste", "MotifNA", "Conditionnel", "Vide si MTBF intrinseque renseigne", "Saisie", "", 1, "", "Liste : NA, NA MECA, NA KIT.")

    Call AjouterChamp(lo, "Produit", "Drawing Number", "Drawing number", 90, "Texte", "", "Oui", "", "Saisie", "", 1, "", "Nom BdD sans suffixe _PRODUIT.")

    Call AjouterChamp(lo, "Produit", "MTBF intrinseque (h)", "MTBF intrinseque Produit (h) issu de la datasheet _PRODUIT", 100, "Nombre", "", "Conditionnel", "Vide si Motif NA renseigne", "Saisie", "", 2, "", "Obligatoire sauf si Motif NA renseigne. Note : espace avant _PRODUIT dans le nom BdD.")

    Call AjouterChamp(lo, "Produit", "DVL HARD generique", "??_DureeVieLimiteeHARDGenerique_a_definir", 110, "Nombre", "", "Non", "", "Saisie", "", 1, "", "Duree de vie limitee HARD generique (a definir).")

    Call AjouterChamp(lo, "Produit", "DVL SOFT generique", "??_DureeVieLimiteeSOFTGenerique_a_definir", 120, "Nombre", "", "Non", "", "Saisie", "", 1, "", "Duree de vie limitee SOFT generique (a definir).")

    Call AjouterChamp(lo, "Produit", "Prix Spare manuel", "Prix SPARE manuel_PRODUIT", 130, "Nombre", "", "Non", "", "Saisie", "", 1, "", "Prix Spare manuel (saisie libre).")

    Call AjouterChamp(lo, "Produit", "Lead Time Spare", "??_LeadTimeSpare_a_definir", 140, "Nombre", "", "Non", "", "Saisie", "", 1, "", "Lead time Spare (jours) - a definir.")

    Call AjouterChamp(lo, "Produit", "Source Spare", "??_SourceSpare_a_definir", 150, "Texte", "", "Non", "", "Saisie", "", 1, "", "Source du prix Spare (a definir).")

    Call AjouterChamp(lo, "Produit", "Date Source Spare", "??_DateSourceSpare_a_definir", 160, "Date", "", "Non", "", "Saisie", "", 1, "", "Date de la source du prix Spare (a definir).")

    Call AjouterChamp(lo, "Produit", "Prix Repair manuel", "Prix REPAIR manuel_PRODUIT", 170, "Nombre", "", "Non", "", "Saisie", "", 1, "", "Prix Repair manuel (saisie libre).")

    Call AjouterChamp(lo, "Produit", "Lead Time Repair", "Lead time REPAIR (jours) manuel_PRODUIT", 180, "Nombre", "", "Non", "", "Saisie", "", 1, "", "Lead time Repair manuel (jours).")

    Call AjouterChamp(lo, "Produit", "Source Repair", "??_SourceRepair_a_definir", 190, "Texte", "", "Non", "", "Saisie", "", 1, "", "Source du prix Repair (a definir).")

    Call AjouterChamp(lo, "Produit", "Date Source Repair", "??_DateSourceRepair_a_definir", 200, "Date", "", "Non", "", "Saisie", "", 1, "", "Date de la source du prix Repair (a definir).")

    Call AjouterChamp(lo, "Produit", "Reference datasheet", "??_RefDatasheet_a_definir", 210, "Texte", "", "Non", "", "Saisie", "", 1, "", "Reference de la datasheet (a definir).")

    Call AjouterChamp(lo, "Produit", "Reference MUL installation", "??_RefMULInstallation_a_definir", 220, "Texte", "", "Non", "", "Saisie", "", 1, "", "Reference MUL installation (a definir).")

    Call AjouterChamp(lo, "Produit", "Reference MUL maintenance", "??_RefMULMaintenance_a_definir", 230, "Texte", "", "Non", "", "Saisie", "", 1, "", "Reference MUL maintenance (a definir).")

    Call AjouterChamp(lo, "Produit", "Reference MUL software", "??_RefMULSoftware_a_definir", 240, "Texte", "", "Non", "", "Saisie", "", 1, "", "Reference MUL software (a definir).")

    Call AjouterChamp(lo, "Produit", "Reference MUL interface", "??_RefMULInterface_a_definir", 250, "Texte", "", "Non", "", "Saisie", "", 1, "", "Reference MUL interface (a definir).")
End Sub

' =========================================================================
' PEUPLEMENT DU CONFIG POUR LA TABLE INFO ATM
' =========================================================================
' Brief v0.2 section 6.2.
' Obligatoires : DTR, HARD/SOFT.
' Auto / Herites (grises) : Descriptif FR/EN, Plan (herite Produit.NumeroPlan),
' Reparabilite, Poids (herite EMMST.Poids), Temps generique.
' Alerte : Poids absent ==> EMMST inconnu en base.
Private Sub PeuplerInfoATM(lo As ListObject)
    ' ColonneBdD = vrais noms BdD (cf. Champs_BDD.xlsx, colonne Info_ATM).

    Call AjouterChamp(lo, "InfoATM", "DTR", "DTR_INFOATM", 10, "Texte", "", "Oui", "", "Saisie", "", 3, "", "Identifiant unique ATM (cle).")

    Call AjouterChamp(lo, "InfoATM", "HARD/SOFT", "HARD/SOFT", 20, "Liste", "HardSoft", "Oui", "", "Saisie", "", 2, "HS_Info = HS_Tache", "Doit correspondre au HARD/SOFT des Taches associees. Note : nom BdD sans suffixe table.")

    Call AjouterChamp(lo, "InfoATM", "Descriptif FR", "DescriptifATMFR_INFOATM", 30, "Texte", "", "Non", "", "Auto", "", 0, "", "Lecture seule, recupere de la BdD (affichage gris).")

    Call AjouterChamp(lo, "InfoATM", "Descriptif EN", "DescriptifATMEN_INFOATM", 40, "Texte", "", "Non", "", "Auto", "", 0, "", "Lecture seule, recupere de la BdD (affichage gris).")

    Call AjouterChamp(lo, "InfoATM", "Plan", "PLANATM_INFOATM", 50, "Texte", "", "Non", "", "Herite", "Produit.Drawing number", 0, "", "Herite de Produit (affichage gris).")

    Call AjouterChamp(lo, "InfoATM", "Reparabilite", "Reparabilite_INFOATM", 60, "Texte", "", "Non", "", "Auto", "", 0, "", "Lecture seule (affichage gris).")

    Call AjouterChamp(lo, "InfoATM", "Poids", "Poids_INFOATM", 70, "Nombre", "", "Non", "", "Herite", "EMMST.DimensionNuPoids", 0, "Poids absent ==> EMMST inconnu", "Herite de EMMST. Si absent en base ==> alerte EMMST inconnu.")

    Call AjouterChamp(lo, "InfoATM", "Temps generique", "Temps generique de la tache_INFOATM", 80, "Nombre", "", "Non", "", "Auto", "", 0, "", "Lecture seule, calcule (affichage gris).")
End Sub

' =========================================================================
' PEUPLEMENT DU CONFIG POUR LA TABLE TACHES ATM
' =========================================================================
' Brief v0.2 section 6.2.
' Obligatoires : DTR, Rang, Langue, HARD/SOFT, Description (rang 3 & 4).
' Regles metier :
'   * "Torque Tighten" dans Description ==> critere "N.m" attendu
'   * Rang 2.1 / 5.1 ET NodeType >= SRU ==> Banc test attendu
'   * Coherence HARD/SOFT InfoATM <==> TachesATM
Private Sub PeuplerTachesATM(lo As ListObject)
    ' ColonneBdD = vrais noms BdD (cf. Champs_BDD.xlsx, colonne Taches_ATM).

    Call AjouterChamp(lo, "TachesATM", "DTR", "DTR_TACHEATM", 10, "Texte", "", "Oui", "", "Saisie", "", 3, "", "Cle parent (ATM de rattachement).")

    Call AjouterChamp(lo, "TachesATM", "Rang", "Rang_TACHEATM", 20, "Texte", "", "Oui", "", "Saisie", "", 2, "", "Numerotation hierarchique (ex. 2.1, 5.1, 3.2.4).")

    Call AjouterChamp(lo, "TachesATM", "Langue", "Langue_TACHEATM", 30, "Liste", "Langue", "Oui", "", "Saisie", "", 1, "", "FR ou EN.")

    Call AjouterChamp(lo, "TachesATM", "HARD/SOFT", "HARD/SOFT", 40, "Liste", "HardSoft", "Oui", "", "Saisie", "", 1, "HS_Info = HS_Tache", "Doit etre coherent avec InfoATM.HARD/SOFT du meme DTR.")

    Call AjouterChamp(lo, "TachesATM", "Description", "Description_TACHEATM", 50, "Texte", "", "Conditionnel", "Rang 3.* ou 4.*", "Saisie", "", 2, "Torque Tighten ==> critere N.m attendu", "Obligatoire pour rang 3.* et 4.*. Si contient 'Torque Tighten' ==> critere N.m attendu.")

    Call AjouterChamp(lo, "TachesATM", "Critere Nm", "CriteresDimensionnels_TACHEATM", 60, "Nombre", "", "Conditionnel", "Description contient 'Torque Tighten'", "Saisie", "", 1, "", "Couple de serrage en N.m, attendu si Description contient 'Torque Tighten'. (Champ BdD 'CriteresDimensionnels' a confirmer.)")

    ' "Banc test" du brief = en realite les outillages (BTE = Banc Test
    ' Equipement). Pas de champ BancTest dedie : on utilise les colonnes
    ' Outillage standard / specifique de la table Taches_ATM.
    Call AjouterChamp(lo, "TachesATM", "Outillage standard", "OutillageStandard_TACHEATM", 70, "Texte", "", "Conditionnel", "Rang 2.1/5.1 et NodeType >= SRU", "Saisie", "", 1, "", "Outillage standard utilise (couvre le besoin BTE du brief).")

    Call AjouterChamp(lo, "TachesATM", "Outillage specifique", "OutillageSpecifique_TACHEATM", 80, "Texte", "", "Conditionnel", "Rang 2.1/5.1 et NodeType >= SRU", "Saisie", "", 1, "", "Outillage specifique projet (couvre le besoin BTE du brief).")
End Sub

' =========================================================================
' PEUPLEMENT DU CONFIG POUR LA TABLE EMMST
' =========================================================================
' Brief v0.2 section 6.3 : "Tous champs attendus" (dimensions nu &
' conditionne, conditionnement, stockage, transport, marquage, EN).
' Feuille longue sectionnee, nombreux booleens.
' Peuplement Phase 0 : champs principaux derives du Type EMMST_Col existant
' (Z_Variables.bas). A completer Phase 1+ avec la grille complete EMMST.
Private Sub PeuplerEMMST(lo As ListObject)
    ' ColonneBdD = vrais noms BdD (cf. Champs_BDD.xlsx, colonne EMMST).
    ' 38 champs alignes sur Champs_BDD.xlsx. Brief 6.3 : feuille longue
    ' sectionnee, "tous champs attendus" (la plupart Non bloquants).

    ' --- Identification ---
    Call AjouterChamp(lo, "EMMST", "DTR", "DTR_EMMST", 10, "Texte", "", "Oui", "", "Saisie", "", 3, "", "Identifiant unique EMMST (cle).")

    ' --- Dimensions nu ---
    Call AjouterChamp(lo, "EMMST", "Longueur nu", "DimensionNuLongueur_EMMST", 100, "Nombre", "", "Oui", "", "Saisie", "", 1, "", "mm")
    Call AjouterChamp(lo, "EMMST", "Largeur nu", "DimensionNuLargeur_EMMST", 110, "Nombre", "", "Oui", "", "Saisie", "", 1, "", "mm")
    Call AjouterChamp(lo, "EMMST", "Hauteur nu", "DimensionNuHauteur_EMMST", 120, "Nombre", "", "Oui", "", "Saisie", "", 1, "", "mm")
    Call AjouterChamp(lo, "EMMST", "Poids nu", "DimensionNuPoids", 130, "Nombre", "", "Oui", "", "Saisie", "", 2, "", "kg. Reutilise par InfoATM.Poids (heritage). Note : nom BdD sans suffixe _EMMST.")

    ' --- Dimensions conditionne ---
    Call AjouterChamp(lo, "EMMST", "Longueur conditionne", "DimensionConditionneLongueur_EMMST", 200, "Nombre", "", "Oui", "", "Saisie", "", 1, "", "mm")
    Call AjouterChamp(lo, "EMMST", "Largeur conditionne", "DimensionConditionneLargeur_EMMST", 210, "Nombre", "", "Oui", "", "Saisie", "", 1, "", "mm")
    Call AjouterChamp(lo, "EMMST", "Hauteur conditionne", "DimensionConditionneHauteur_EMMST", 220, "Nombre", "", "Oui", "", "Saisie", "", 1, "", "mm")
    Call AjouterChamp(lo, "EMMST", "Poids conditionne", "DimensionConditionnePoids_EMMST", 230, "Nombre", "", "Oui", "", "Saisie", "", 1, "", "kg")

    ' --- Conditionnement ---
    Call AjouterChamp(lo, "EMMST", "Unite conditionnement", "UniteConditionnement_EMMST", 300, "Texte", "", "Non", "", "Saisie", "", 0, "", "")
    Call AjouterChamp(lo, "EMMST", "Date peremption", "DatePeremption_EMMST", 310, "Date", "", "Non", "", "Saisie", "", 0, "", "")
    Call AjouterChamp(lo, "EMMST", "Duree peremption", "DureePeremption_EMMST", 320, "Texte", "", "Non", "", "Saisie", "", 0, "", "")
    Call AjouterChamp(lo, "EMMST", "Reutilisable", "Reutilisable_EMMST", 330, "Booleen", "Booleen", "Non", "", "Saisie", "", 0, "", "")

    ' --- Marquage / Manutention ---
    Call AjouterChamp(lo, "EMMST", "Etiquette emballage", "EtiquetteEmballage_EMMST", 400, "Texte", "", "Non", "", "Saisie", "", 0, "", "")
    Call AjouterChamp(lo, "EMMST", "Marquage equipement", "MarquageEquipement_EMMST", 410, "Texte", "", "Non", "", "Saisie", "", 0, "", "")
    Call AjouterChamp(lo, "EMMST", "Manutention equipement", "ManutentionEquipement_EMMST", 420, "Texte", "", "Non", "", "Saisie", "", 0, "", "")
    Call AjouterChamp(lo, "EMMST", "Outillage specifique manutention", "OutillageSpecifiqueManutention_EMMST", 430, "Texte", "", "Non", "", "Saisie", "", 0, "", "")
    Call AjouterChamp(lo, "EMMST", "Precaution manutention", "PrecautionManutention_EMMST", 440, "Texte", "", "Non", "", "Saisie", "", 0, "", "")
    Call AjouterChamp(lo, "EMMST", "Type outil", "TypeOutil_EMMST", 450, "Texte", "", "Non", "", "Saisie", "", 0, "", "")

    ' --- Stockage (booleens + commentaires) ---
    Call AjouterChamp(lo, "EMMST", "Position stockage", "PositionStockage_EMMST", 500, "Texte", "", "Non", "", "Saisie", "", 0, "", "")
    Call AjouterChamp(lo, "EMMST", "Stockage exterieur", "StockageExterieur_EMMST", 510, "Booleen", "Booleen", "Non", "", "Saisie", "", 0, "", "")
    Call AjouterChamp(lo, "EMMST", "Stockage sous abri", "StockageSousAbri_EMMST", 520, "Booleen", "Booleen", "Non", "", "Saisie", "", 0, "", "")
    Call AjouterChamp(lo, "EMMST", "Stockage air libre", "StockageAirLibre_EMMST", 530, "Booleen", "Booleen", "Non", "", "Saisie", "", 0, "", "")
    Call AjouterChamp(lo, "EMMST", "Stockage hors poussiere", "StockageHorsPoussiere_EMMST", 540, "Booleen", "Booleen", "Non", "", "Saisie", "", 0, "", "")
    Call AjouterChamp(lo, "EMMST", "Stockage hors eau", "StockageHorsEau_EMMST", 550, "Booleen", "Booleen", "Non", "", "Saisie", "", 0, "", "")
    Call AjouterChamp(lo, "EMMST", "Stockage interieur", "StockageInterieur_EMMST", 560, "Booleen", "Booleen", "Non", "", "Saisie", "", 0, "", "")
    Call AjouterChamp(lo, "EMMST", "Stockage chaud", "StockageChaud_EMMST", 570, "Booleen", "Booleen", "Non", "", "Saisie", "", 0, "", "")
    Call AjouterChamp(lo, "EMMST", "Stockage froid", "StockageFroid_EMMST", 580, "Booleen", "Booleen", "Non", "", "Saisie", "", 0, "", "")
    Call AjouterChamp(lo, "EMMST", "Stockage hors lumiere", "StockageHorsLumiere_EMMST", 590, "Booleen", "Booleen", "Non", "", "Saisie", "", 0, "", "")
    Call AjouterChamp(lo, "EMMST", "Gerbable", "Gerbable_EMMST", 600, "Booleen", "Booleen", "Non", "", "Saisie", "", 0, "", "")
    Call AjouterChamp(lo, "EMMST", "Moyen stockage", "MoyenStockage_EMMST", 610, "Texte", "", "Non", "", "Saisie", "", 0, "", "")
    Call AjouterChamp(lo, "EMMST", "Precaution stockage", "PrecautionStockage_EMMST", 620, "Texte", "", "Non", "", "Saisie", "", 0, "", "")
    Call AjouterChamp(lo, "EMMST", "Maintenance stockage", "MaintenanceStockage_EMMST", 630, "Texte", "", "Non", "", "Saisie", "", 0, "", "")

    ' --- Transport (presents dans les SQL existants, absents de Champs_BDD) ---
    Call AjouterChamp(lo, "EMMST", "Moyen transport", "MoyenTransport_EMMST", 660, "Texte", "", "Non", "", "Saisie", "", 0, "", "")
    Call AjouterChamp(lo, "EMMST", "Outillage specifique transport", "OutillageSpecifiqueTransport_EMMST", 670, "Texte", "", "Non", "", "Saisie", "", 0, "", "")

    ' --- Commentaires (par section) ---
    Call AjouterChamp(lo, "EMMST", "Commentaire emballage", "CommentaireEmballage_EMMST", 700, "Texte", "", "Non", "", "Saisie", "", 0, "", "")
    Call AjouterChamp(lo, "EMMST", "Commentaire marquage", "CommentaireMarquage_EMMST", 710, "Texte", "", "Non", "", "Saisie", "", 0, "", "")
    Call AjouterChamp(lo, "EMMST", "Commentaire manutention", "CommentaireManutention_EMMST", 720, "Texte", "", "Non", "", "Saisie", "", 0, "", "")
    Call AjouterChamp(lo, "EMMST", "Commentaire stockage", "CommentaireStockage_EMMST", 730, "Texte", "", "Non", "", "Saisie", "", 0, "", "")
    Call AjouterChamp(lo, "EMMST", "Commentaire transport", "CommentaireTransport_EMMST", 740, "Texte", "", "Non", "", "Saisie", "", 0, "", "")

    ' --- Versioning ---
    Call AjouterChamp(lo, "EMMST", "Version", "Version_EMMST", 900, "Texte", "", "Non", "", "Auto", "", 0, "", "Version de la fiche EMMST (geree en interne).")
End Sub

' =========================================================================
' PEUPLEMENT DU CONFIG POUR LA TABLE PROJET SLI
' =========================================================================
' Brief v0.2 section 6.4. Champs utiles au PROVIDOR.
' Briques Formation / BTE / MCO / OBSO exclues (perimetre PATH).
Private Sub PeuplerProjetSLI(lo As ListObject)
    ' ColonneBdD = vrais noms BdD (cf. Champs_BDD.xlsx, colonne Projet_SLI).
    ' Perimetre arrete avec Thomas (2026-06-09) : identite, configuration,
    ' profil de mission, economique, GDP, requis, PSLI, references et
    ' interlocuteurs. Exclus : dates de jalons, livrables PROVIDOR/ATM/
    ' EMMST/MUL, prix Formation/BTE/MCO/OBSO, LCC par sous-domaine,
    ' colonnes formules Excel (MeQ/Correcteur/Date MEQ).

    ' --- Identite projet (Ordre 10-60) -----------------------
    Call AjouterChamp(lo, "ProjetSLI", "Nom projet", "Nom projet_PROJET SLI", 10, "Texte", "", "Oui", "", "Saisie", "", 3, "", "Identifiant unique projet SLI (cle).")
    Call AjouterChamp(lo, "ProjetSLI", "Projet / Offre", "Projet / Offre_PROJET SLI", 20, "Texte", "", "Non", "", "Saisie", "", 1, "", "")
    Call AjouterChamp(lo, "ProjetSLI", "Region", "Region_PROJET SLI", 30, "Texte", "", "Non", "", "Saisie", "", 1, "", "")
    Call AjouterChamp(lo, "ProjetSLI", "CT Code", "CT Code_PROJET SLI", 40, "Texte", "", "Non", "", "Saisie", "", 1, "", "")
    Call AjouterChamp(lo, "ProjetSLI", "Type de train", "Type de train_PROJET SLI", 50, "Liste", "TypeTrain", "Oui", "", "Saisie", "", 2, "", "")
    Call AjouterChamp(lo, "ProjetSLI", "Client final", "Client final_PROJET SLI", 60, "Texte", "", "Non", "", "Saisie", "", 1, "", "")

    ' --- Configuration projet (Ordre 70-150) -----------------
    Call AjouterChamp(lo, "ProjetSLI", "Duree de circulation", "Duree de circulation_PROJET SLI", 70, "Nombre", "", "Oui", "", "Saisie", "", 2, "", "En annees.")
    Call AjouterChamp(lo, "ProjetSLI", "Nombre de trains", "Nombre de train_PROJET SLI", 80, "Nombre", "", "Oui", "", "Saisie", "", 2, "", "Note : nom BdD 'train' au singulier.")
    Call AjouterChamp(lo, "ProjetSLI", "Duree garantie", "Duree de garantie_PROJET SLI", 90, "Nombre", "", "Oui", "", "Saisie", "", 1, "", "En annees.")
    Call AjouterChamp(lo, "ProjetSLI", "Niveau de maintenance", "Niv de maintenance_PROJET SLI", 100, "Liste", "NiveauMaintenance", "Oui", "", "Saisie", "", 1, "", "")
    Call AjouterChamp(lo, "ProjetSLI", "Nombre DTR EAP/PBS", "Nombre de DTR dans l'EAP ou la PBS_PROJET SLI", 110, "Nombre", "", "Non", "", "Saisie", "", 1, "", "")
    Call AjouterChamp(lo, "ProjetSLI", "Nombre fournisseurs CDRL", "Nombre de fournisseur soumis a CDRL_PROJET SLI", 120, "Nombre", "", "Non", "", "Saisie", "", 1, "", "")
    Call AjouterChamp(lo, "ProjetSLI", "Type de pack", "Type de pack (complexe/standard/mini)_PROJET SLI", 130, "Texte", "", "Non", "", "Saisie", "", 1, "", "complexe / standard / mini.")
    Call AjouterChamp(lo, "ProjetSLI", "UO grand pere du pack", "UO ""grand pere"" du pack_PROJET SLI", 140, "Texte", "", "Non", "", "Saisie", "", 1, "", "")
    Call AjouterChamp(lo, "ProjetSLI", "Statut", "Statut (termine/En cours/A planifier)_PROJET SLI", 150, "Texte", "", "Non", "", "Saisie", "", 1, "", "termine / En cours / A planifier.")

    ' --- Profil de mission (Ordre 160-190) -------------------
    Call AjouterChamp(lo, "ProjetSLI", "Servicing/day", "Servicing day by year_PROJET SLI", 160, "Nombre", "", "Oui", "", "Saisie", "", 1, "", "Profil de mission.")
    Call AjouterChamp(lo, "ProjetSLI", "PH/year", "PH by year_PROJET SLI", 170, "Nombre", "", "Oui", "", "Saisie", "", 1, "", "Power Hours par an.")
    Call AjouterChamp(lo, "ProjetSLI", "RH/year", "RH by year_PROJET SLI", 180, "Nombre", "", "Oui", "", "Saisie", "", 1, "", "Running Hours par an.")
    Call AjouterChamp(lo, "ProjetSLI", "Km/year", "Km by year_PROJET SLI", 190, "Nombre", "", "Oui", "", "Saisie", "", 1, "", "Kilometres par an.")

    ' --- Economique (Ordre 200-230) --------------------------
    Call AjouterChamp(lo, "ProjetSLI", "Devise", "Devise_PROJET SLI", 200, "Liste", "Devise", "Oui", "", "Saisie", "", 1, "", "")
    Call AjouterChamp(lo, "ProjetSLI", "Man hour cost", "MAN hour cost_PROJET SLI", 210, "Nombre", "", "Oui", "", "Saisie", "", 1, "", "Cout horaire main d'oeuvre dans la devise du projet.")
    Call AjouterChamp(lo, "ProjetSLI", "LCC scope total", "LCC valeur scope total_PROJET SLI", 220, "Nombre", "", "Oui", "", "Saisie", "", 1, "", "Perimetre LCC en cout total.")

    ' --- Metrologie GDP (Ordre 240-270) ----------------------
    Call AjouterChamp(lo, "ProjetSLI", "Nbre GDP classique", "nbre GDP classique_PROJET SLI", 240, "Nombre", "", "Non", "", "Saisie", "", 0, "", "")
    Call AjouterChamp(lo, "ProjetSLI", "Nbre GDP fournisseur", "nbre GDP fournisseur_PROJET SLI", 250, "Nombre", "", "Non", "", "Saisie", "", 0, "", "")
    Call AjouterChamp(lo, "ProjetSLI", "% GDP classique consommee", "% de GDP classique consommee_PROJET SLI", 260, "Nombre", "", "Non", "", "Saisie", "", 0, "", "")
    Call AjouterChamp(lo, "ProjetSLI", "% GDP fournisseur consommee", "% de GDP fournisseur consommee_PROJET SLI", 270, "Nombre", "", "Non", "", "Saisie", "", 0, "", "")

    ' --- Requis client / fournisseur + PSLI (Ordre 280-330) ---
    ' Retires 2026-06-10 (decision Thomas) : Requis client Acceptance/
    ' Clausing/MoP, Requis fournisseur Clausing, PSLI preliminaire/final.
    ' Hors scope de SIRUS, geres autrement cote process.

    ' --- References (Ordre 340-360) --------------------------
    Call AjouterChamp(lo, "ProjetSLI", "Reference PBS/EAP", "Ref PBS/EAP_PROJET SLI", 340, "Texte", "", "Oui", "", "Saisie", "", 1, "", "")
    Call AjouterChamp(lo, "ProjetSLI", "Reference RAM report", "Ref RAM report projet_PROJET SLI", 350, "Texte", "", "Oui", "", "Saisie", "", 1, "", "")
    Call AjouterChamp(lo, "ProjetSLI", "Version Providor/BASL en BdD", "Version du Providor/BASL integre en BdD_PROJET SLI", 360, "Texte", "", "Non", "", "Saisie", "", 0, "", "")

    ' --- Interlocuteurs RSLI / RASL / LU (Ordre 400-490) -----
    Call AjouterChamp(lo, "ProjetSLI", "RSLI", "RSLI_PROJET SLI", 400, "Texte", "", "Non", "", "Saisie", "", 1, "", "Responsable SLI projet.")
    Call AjouterChamp(lo, "ProjetSLI", "RASL", "RASL_PROJET SLI", 410, "Texte", "", "Non", "", "Saisie", "", 1, "", "Responsable ASL projet.")
    Call AjouterChamp(lo, "ProjetSLI", "LU", "LU_PROJET SLI", 420, "Texte", "", "Non", "", "Saisie", "", 0, "", "")
    Call AjouterChamp(lo, "ProjetSLI", "Interlocuteur LU", "InterlocuteurLU_PROJET SLI", 430, "Texte", "", "Non", "", "Saisie", "", 0, "", "")
    Call AjouterChamp(lo, "ProjetSLI", "ISR Manager LU", "ISR Manager LU_PROJET SLI", 440, "Texte", "", "Non", "", "Saisie", "", 0, "", "")
    Call AjouterChamp(lo, "ProjetSLI", "LSA Engineer LU", "LSA Engineer LU_PROJET SLI", 450, "Texte", "", "Non", "", "Saisie", "", 0, "", "")
    Call AjouterChamp(lo, "ProjetSLI", "Doc Manager LU", "Doc Manager LU_PROJET SLI", 460, "Texte", "", "Non", "", "Saisie", "", 0, "", "")
    Call AjouterChamp(lo, "ProjetSLI", "System Engineer LU", "System Engineer LU_PROJET SLI", 470, "Texte", "", "Non", "", "Saisie", "", 0, "", "")
    Call AjouterChamp(lo, "ProjetSLI", "TC Engineer LU", "TC Engineer LU_PROJET SLI", 480, "Texte", "", "Non", "", "Saisie", "", 0, "", "")
    Call AjouterChamp(lo, "ProjetSLI", "BID Manager LU", "BID Manager LU_PROJET SLI", 490, "Texte", "", "Non", "", "Saisie", "", 0, "", "")

    ' --- Interlocuteurs VBN / BdD (Ordre 500-560) ------------
    Call AjouterChamp(lo, "ProjetSLI", "PM ou BM VBN", "PM ou BM VBN_PROJET SLI", 500, "Texte", "", "Non", "", "Saisie", "", 0, "", "")
    Call AjouterChamp(lo, "ProjetSLI", "PrEM ou BTM", "PrEM ou BTM_PROJET SLI", 510, "Texte", "", "Non", "", "Saisie", "", 0, "", "")
    Call AjouterChamp(lo, "ProjetSLI", "PrPM", "PrPM_PROJET SLI", 520, "Texte", "", "Non", "", "Saisie", "", 0, "", "")
    Call AjouterChamp(lo, "ProjetSLI", "PMM", "PMM_PROJET SLI", 530, "Texte", "", "Non", "", "Saisie", "", 0, "", "")
    Call AjouterChamp(lo, "ProjetSLI", "Procurement", "Procurement_PROJET SLI", 540, "Texte", "", "Non", "", "Saisie", "", 0, "", "")
    Call AjouterChamp(lo, "ProjetSLI", "RAM", "RAM_PROJET SLI", 550, "Texte", "", "Non", "", "Saisie", "", 0, "", "")

    ' --- Design Leaders par scope (Ordre 600-700) ------------
    Call AjouterChamp(lo, "ProjetSLI", "DL PA_PACIS", "DL PA_PACIS_PROJET SLI", 600, "Texte", "", "Non", "", "Saisie", "", 0, "", "")
    Call AjouterChamp(lo, "ProjetSLI", "DL PIS", "DL PIS_PROJET SLI", 610, "Texte", "", "Non", "", "Saisie", "", 0, "", "")
    Call AjouterChamp(lo, "ProjetSLI", "DL CCTV", "DL CCTV_PROJET SLI", 620, "Texte", "", "Non", "", "Saisie", "", 0, "", "")
    Call AjouterChamp(lo, "ProjetSLI", "DL IOB", "DL IOB_PROJET SLI", 630, "Texte", "", "Non", "", "Saisie", "", 0, "", "")
    Call AjouterChamp(lo, "ProjetSLI", "DL APC", "DL APC_PROJET SLI", 640, "Texte", "", "Non", "", "Saisie", "", 0, "", "")
    Call AjouterChamp(lo, "ProjetSLI", "DL DAS", "DL DAS_PROJET SLI", 650, "Texte", "", "Non", "", "Saisie", "", 0, "", "")
    Call AjouterChamp(lo, "ProjetSLI", "DL EVR", "DL EVR_PROJET SLI", 660, "Texte", "", "Non", "", "Saisie", "", 0, "", "")
    Call AjouterChamp(lo, "ProjetSLI", "DL LECTEUR BADGE", "DL LECTEUR BADGE_PROJET SLI", 670, "Texte", "", "Non", "", "Saisie", "", 0, "", "")
    Call AjouterChamp(lo, "ProjetSLI", "DL GSM REPEATER", "DL GSM REPEATER_PROJET SLI", 680, "Texte", "", "Non", "", "Saisie", "", 0, "", "")
    Call AjouterChamp(lo, "ProjetSLI", "DL SRS", "DL SRS_PROJET SLI", 690, "Texte", "", "Non", "", "Saisie", "", 0, "", "")
    Call AjouterChamp(lo, "ProjetSLI", "Autres", "Autres_PROJET SLI", 700, "Texte", "", "Non", "", "Saisie", "", 0, "", "")
End Sub

' =========================================================================
' PEUPLEMENT DU CONFIG POUR LA TABLE PROJET ASL (mise a jour PBS)
' =========================================================================
' Spec Thomas 2026-06-11 (cf. Docs/SIRUS_REGLES.md).
' Obligatoires : Nom projet, Sous Systeme, DTR, Designation PBS, Qte,
' MTBF fonctionnel, Prix Spare archive, Lead Time Spare archive, Testabilite.
' Conditionnels (si NodeType = LRU ou SRU) : Prix Repair archive, Lead Time
' Repair archive (les autres NodeType n ont pas de cycle reparation).
' Heritage Produit (info pre-remplie, non obligatoire) : Node Type,
' MTBF intrinseque Projet, DVL fonctionnel HARD/SOFT.
' Coherence forte : Nom projet ==> doit correspondre a un Projet SLI existant
' (bloque a l envoi via SIRUS_UI_BtnEnvoyerClick / SIRUS_ProjetsASL_Inconnus).
Private Sub PeuplerProjetASL(lo As ListObject)
    ' ColonneBdD = vrais noms BdD (cf. Champs_BDD.xlsx + diag Thomas 2026-06-09).
    ' Toutes les colonnes utiles incluses. Exclus : ID Projet (ASL) (auto),
    ' ISP TIS HORS FRANCE (archive), Mise en qualite / Correcteur / Date de
    ' MEQ (formules Excel BdD).

    ' Ordre verrouille Thomas 2026-06-11 (cf. Docs/SIRUS_REGLES.md).
    ' Note : les noms BdD (ColonneBdD) restent inchanges (lies au schema
    ' Champs_BDD.xlsx) ; seuls les libelles humains "Champ" sont alignes
    ' sur la nomenclature ASL ("archive", "fonctionnel", ...).

    ' --- Cles / identification (Ordre 10-50) -----------------
    Call AjouterChamp(lo, "ProjetASL", "Nom projet", "PROJET_PROJETASL", 10, "Texte", "", "Oui", "", "Saisie", "", 3, "Doit correspondre a un Projet SLI existant", "Coherence forte : doit exister dans ProjetSLI.[Nom projet_PROJET SLI].")
    Call AjouterChamp(lo, "ProjetASL", "Sous Systeme", "Sous Systeme_PROJETASL", 20, "Texte", "", "Oui", "", "Saisie", "", 2, "", "Sous-systeme du projet (obligatoire).")
    Call AjouterChamp(lo, "ProjetASL", "DTR", "DTR_PROJETASL", 30, "Texte", "", "Oui", "", "Saisie", "", 3, "", "Identifiant unique du DTR de la PBS.")
    Call AjouterChamp(lo, "ProjetASL", "Designation PBS", "Designation_PROJETASL", 40, "Texte", "", "Oui", "", "Saisie", "", 2, "", "Designation issue de la PBS du projet.")
    Call AjouterChamp(lo, "ProjetASL", "Qte", "Q PBS_PROJETASL", 50, "Nombre", "", "Oui", "", "Saisie", "", 1, "", "Quantite par train.")

    ' --- Saisies projet techniques (Ordre 60-100) ------------
    Call AjouterChamp(lo, "ProjetASL", "Node Type", "Node typed PROJET_PROJETASL", 60, "Liste", "NodeType", "Non", "", "Saisie", "Produit.Node type Produit (LRU/SRU...)_PRODUIT", 0, "", "Pre-rempli depuis Produit (lecture conseillee). Saisie possible si specifique projet.")
    Call AjouterChamp(lo, "ProjetASL", "MTBF intrinseque Projet", "MTBF intrinseque Projet (h) issu du RAM report projet_PROJETASL", 70, "Nombre", "", "Non", "", "Saisie", "Produit.MTBF intrinseque Produit (h) issu de la datasheet _PRODUIT", 0, "", "MTBF intrinseque projet (RAM report). Pre-rempli depuis Produit si dispo.")
    Call AjouterChamp(lo, "ProjetASL", "MTBF fonctionnel", "MTBF fonctionnel (h) issu du RAM report projet_PROJETASL", 80, "Nombre", "", "Oui", "", "Saisie", "", 2, "", "MTBF fonctionnel projet (peut differer du MTBF intrinseque Produit).")
    ' Note : suffixe '_PROJET ASL' avec espace, confirme Thomas 2026-06-09
    ' (different des autres ProjetASL en '_PROJETASL' sans espace).
    Call AjouterChamp(lo, "ProjetASL", "DVL fonctionnel HARD", "DUREE DE VIE LIMITEE (RAM report specifique projet (CGR)) HARDWARE_PROJET ASL", 90, "Nombre", "", "Non", "", "Saisie", "", 0, "", "DVL Hardware specifique projet (RAM report). Suffixe BdD avec espace.")
    Call AjouterChamp(lo, "ProjetASL", "DVL fonctionnel SOFT", "DUREE DE VIE LIMITEE (RAM report specifique projet (CGR)) SOFTWARE_PROJET ASL", 100, "Nombre", "", "Non", "", "Saisie", "", 0, "", "DVL Software specifique projet (RAM report). Suffixe BdD avec espace.")

    ' --- Prix + Lead times archive (Ordre 110-140) -----------
    '   Prix Spare archive + Lead Time Spare archive : obligatoires (toujours).
    '   Prix Repair archive + Lead Time Repair archive : conditionnels
    '   (obligatoires uniquement si NodeType = LRU ou SRU). Voir G_SIRUS_Regles.
    Call AjouterChamp(lo, "ProjetASL", "Prix Spare archive", "Prix spare (DEVISE du projet (SLI))_PROJETASL", 110, "Nombre", "", "Oui", "", "Saisie", "", 2, "", "Prix SPARE archive dans la devise du projet SLI.")
    Call AjouterChamp(lo, "ProjetASL", "Lead Time Spare archive", "Lead time SPARE (jours)_PROJETASL", 120, "Nombre", "", "Oui", "", "Saisie", "", 1, "", "Lead time SPARE archive en jours.")
    Call AjouterChamp(lo, "ProjetASL", "Prix Repair archive", "Prix repair (DEVISE du projet (SLI))_PROJETASL", 130, "Nombre", "", "Conditionnel", "Obligatoire si NodeType LRU/SRU", "Saisie", "", 1, "", "Prix REPAIR archive dans la devise du projet SLI.")
    Call AjouterChamp(lo, "ProjetASL", "Lead Time Repair archive", "Lead time REPAIR (jours)_PROJETASL", 140, "Nombre", "", "Conditionnel", "Obligatoire si NodeType LRU/SRU", "Saisie", "", 1, "", "Lead time REPAIR archive en jours.")

    ' --- Testabilite + autres prix (Ordre 150-180) -----------
    Call AjouterChamp(lo, "ProjetASL", "Testabilite", "Testabilite (DTR du BTE relatif a l'equipement sur le projet concerne)_PROJETASL", 150, "Liste", "Booleen", "Non", "", "Saisie", "", 0, "", "Liste Oui/Non : indique si le DTR est testable sur le projet. Non obligatoire (decision 2026-06-15) : la testabilite n est pas un critere bloquant pour la MeQ Generation.")
    Call AjouterChamp(lo, "ProjetASL", "Prix Preventif", "Prix preventif (DEVISE du projet (SLI))_PROJETASL", 160, "Nombre", "", "Non", "", "Saisie", "", 0, "", "Prix preventif dans la devise du projet SLI.")
    Call AjouterChamp(lo, "ProjetASL", "Prix After Market Spare", "Prix after market Spare_PROJETASL", 170, "Nombre", "", "Non", "", "Saisie", "", 0, "", "")
    Call AjouterChamp(lo, "ProjetASL", "Prix After Market Repair", "Prix after market Repair_PROJETASL", 180, "Nombre", "", "Non", "", "Saisie", "", 0, "", "")

    ' --- Divers (Ordre 190-200) ------------------------------
    Call AjouterChamp(lo, "ProjetASL", "STBI/DAT", "STBI/DAT_PROJETASL", 190, "Texte", "", "Non", "", "Saisie", "", 0, "", "")
    Call AjouterChamp(lo, "ProjetASL", "LCS", "LCS", 200, "Texte", "", "Non", "", "Saisie", "", 0, "", "Note : nom BdD sans suffixe table.")
End Sub

' =========================================================================
' AJOUT D'UNE LIGNE AU LISTOBJECT DU CONFIG
' =========================================================================
Private Sub AjouterChamp(lo As ListObject, ByVal sTable As String, ByVal sChamp As String, ByVal sColonneBdD As String, ByVal lOrdre As Long, ByVal sTypeChamp As String, ByVal sListeSource As String, ByVal sObligatoire As String, ByVal sConditionDescription As String, ByVal sSource As String, ByVal sSourceHeritage As String, ByVal dPoidsMeQ As Double, ByVal sRegleAlerte As String, ByVal sCommentaire As String)

    Dim r As ListRow
    Set r = lo.ListRows.Add

    r.Range(1, 1).Value = sTable
    r.Range(1, 2).Value = sChamp
    r.Range(1, 3).Value = sColonneBdD
    r.Range(1, 4).Value = lOrdre
    r.Range(1, 5).Value = sTypeChamp
    r.Range(1, 6).Value = sListeSource
    r.Range(1, 7).Value = sObligatoire
    r.Range(1, 8).Value = sConditionDescription
    r.Range(1, 9).Value = sSource
    r.Range(1, 10).Value = sSourceHeritage
    r.Range(1, 11).Value = dPoidsMeQ
    r.Range(1, 12).Value = sRegleAlerte
    r.Range(1, 13).Value = sCommentaire
End Sub

' =========================================================================
' LECTURE DU CONFIG POUR UNE TABLE DONNEE
' =========================================================================
' Renvoie une Collection de Champ_Config pour la table demandee, triee par
' Ordre croissant. Si la feuille n'existe pas ou si la table est inconnue,
' renvoie une Collection vide et trace un message dans la fenetre Immediate.
' Aucune erreur bloquante pour l'utilisateur (lecture diagnostique).
Public Function SIRUS_LireConfig(ByVal sTable As String) As Collection

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim r As ListRow
    Dim col As Collection
    Dim def As SIRUS_Champ_Config
    Dim deja As Boolean
    Dim i As Long
    Dim inserted As Boolean

    Set col = New Collection

    On Error GoTo Sortie

    Set wb = ThisWorkbook

    ' -- Verifier presence de la feuille --------------------
    deja = False
    For Each ws In wb.Worksheets
        If ws.Name = SIRUS_CONFIG_SHEET Then
            deja = True
            Exit For
        End If
    Next ws

    If Not deja Then
        Debug.Print "SIRUS_LireConfig : feuille " & SIRUS_CONFIG_SHEET & " introuvable. Lancer SIRUS_InitConfigSheet d'abord."
        Set SIRUS_LireConfig = col
        Exit Function
    End If

    Set ws = wb.Worksheets(SIRUS_CONFIG_SHEET)
    Set lo = ws.ListObjects(SIRUS_CONFIG_LISTOBJ)

    If lo.ListRows.Count = 0 Then
        Set SIRUS_LireConfig = col
        Exit Function
    End If

    ' -- Parcours des lignes et filtrage par table ----------
    ' Insertion triee par Ordre via Collection.Add ... Before
    ' Important : une NOUVELLE instance par ligne (sinon toutes les
    ' entrees de la Collection pointent vers le meme objet).
    For Each r In lo.ListRows
        If CStr(r.Range(1, 1).Value) = sTable Then
            Set def = New SIRUS_Champ_Config
            def.Table = CStr(r.Range(1, 1).Value)
            def.Champ = CStr(r.Range(1, 2).Value)
            def.ColonneBdD = CStr(r.Range(1, 3).Value)
            def.Ordre = NzLng(r.Range(1, 4).Value)
            def.TypeChamp = CStr(r.Range(1, 5).Value)
            def.ListeSource = CStr(r.Range(1, 6).Value)
            def.Obligatoire = CStr(r.Range(1, 7).Value)
            def.ConditionDescription = CStr(r.Range(1, 8).Value)
            def.Source = CStr(r.Range(1, 9).Value)
            def.SourceHeritage = CStr(r.Range(1, 10).Value)
            def.PoidsMeQ = NzDbl(r.Range(1, 11).Value)
            def.RegleAlerte = CStr(r.Range(1, 12).Value)
            def.Commentaire = CStr(r.Range(1, 13).Value)

            inserted = False
            For i = 1 To col.Count
                If col(i).Ordre > def.Ordre Then
                    col.Add def, , i
                    inserted = True
                    Exit For
                End If
            Next i
            If Not inserted Then col.Add def
        End If
    Next r

    Set SIRUS_LireConfig = col
    Exit Function

Sortie:
    Debug.Print "SIRUS_LireConfig : erreur " & Err.Description
    Set SIRUS_LireConfig = col
End Function

' =========================================================================
' HELPERS PRIVES
' =========================================================================
Private Function NzLng(v As Variant) As Long
    If IsNumeric(v) Then
        NzLng = CLng(v)
    Else
        NzLng = 0
    End If
End Function

Private Function NzDbl(v As Variant) As Double
    If IsNumeric(v) Then
        NzDbl = CDbl(v)
    Else
        NzDbl = 0
    End If
End Function

' =========================================================================
' PROCEDURES DE TEST (executables depuis VBE)
' =========================================================================

' Test_InitConfigSheet : (re)cree la feuille SIRUS_Config et la peuple Produit.
' Lancer depuis VBE apres import du module.
Public Sub Test_InitConfigSheet()
    Call SIRUS_InitConfigSheet(bForcer:=True)
End Sub

' Test_LireConfig : lit le config pour la table Produit et trace les
' definitions dans la fenetre Immediate.
Public Sub Test_LireConfig()
    Dim col As Collection
    Dim def As SIRUS_Champ_Config
    Dim s As String

    Set col = SIRUS_LireConfig(SIRUS_TBL_PRODUIT)

    If col.Count = 0 Then
        Debug.Print "Test_LireConfig : aucun champ pour la table Produit."
        Exit Sub
    End If

    Debug.Print "Test_LireConfig : " & col.Count & " champs lus pour Produit."
    For Each def In col
        s = "  Ordre=" & def.Ordre & " | " & def.Champ & " (BdD=" & def.ColonneBdD & ", Type=" & def.TypeChamp & ", Oblig=" & def.Obligatoire & ", Source=" & def.Source & ", MeQ=" & def.PoidsMeQ & ")"
        If LenB(def.ConditionDescription) > 0 Then
            s = s & " [cond: " & def.ConditionDescription & "]"
        End If
        Debug.Print s
    Next def
End Sub

' Test_LireConfig_Toutes : parcourt les 6 tables connues et trace pour
' chacune le nombre de champs lus + la repartition par categorie
' (Obligatoires Oui / Conditionnels / Non) et par source (Saisie / Auto / Herite).
Public Sub Test_LireConfig_Toutes()

    Dim tables As Variant
    Dim t As Variant
    Dim col As Collection
    Dim def As SIRUS_Champ_Config
    Dim nbOui As Long, nbCond As Long, nbNon As Long
    Dim nbSaisie As Long, nbAuto As Long, nbHerite As Long
    Dim total As Long

    tables = Array(SIRUS_TBL_PRODUIT, SIRUS_TBL_INFOATM, SIRUS_TBL_TACHESATM, SIRUS_TBL_EMMST, SIRUS_TBL_PROJETSLI, SIRUS_TBL_PROJETASL)

    Debug.Print "=== Test_LireConfig_Toutes ==="

    For Each t In tables
        Set col = SIRUS_LireConfig(CStr(t))
        nbOui = 0: nbCond = 0: nbNon = 0
        nbSaisie = 0: nbAuto = 0: nbHerite = 0
        For Each def In col
            Select Case def.Obligatoire
                Case "Oui":          nbOui = nbOui + 1
                Case "Conditionnel": nbCond = nbCond + 1
                Case Else:           nbNon = nbNon + 1
            End Select
            Select Case def.Source
                Case "Saisie": nbSaisie = nbSaisie + 1
                Case "Auto":   nbAuto = nbAuto + 1
                Case "Herite": nbHerite = nbHerite + 1
            End Select
        Next def
        total = total + col.Count
        Debug.Print "  " & t & " : " & col.Count & " champs [Oblig: Oui=" & nbOui & " Cond=" & nbCond & " Non=" & nbNon & "] [Source: Saisie=" & nbSaisie & " Auto=" & nbAuto & " Herite=" & nbHerite & "]"
    Next t

    Debug.Print "TOTAL : " & total & " champs sur " & (UBound(tables) - LBound(tables) + 1) & " tables."
End Sub
