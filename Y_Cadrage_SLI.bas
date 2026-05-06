Attribute VB_Name = "Y_Cadrage_SLI"
Public Sub Cadrage_Infos_ATM(Infos_tab As Variant)
' Macro de cadrage de la feuille Info ATM de la BDD
' Il est possible de changer la disposition des colonnes mais pas la nomination des colonnes

'On récupère les numéros de colonne
Info_type.dtr = Arraypos(Infos_tab, "DTR_INFOATM")
Info_type.DescriptifFR = Arraypos(Infos_tab, "DescriptifATMFR_INFOATM")
Info_type.LocalisationFR = Arraypos(Infos_tab, "LocalisationATM_INFOATM")
Info_type.Plan = Arraypos(Infos_tab, "PLANATM_INFOATM")
Info_type.CommentaireFR = Arraypos(Infos_tab, "CommentaireATM_INFOATM")
Info_type.version = Arraypos(Infos_tab, "Version_INFOATM")
Info_type.DescriptifEN = Arraypos(Infos_tab, "DescriptifATMEN_INFOATM")
Info_type.LocalisationEN = Arraypos(Infos_tab, "LocalisationATMEN_INFOATM")
Info_type.CommentaireEN = Arraypos(Infos_tab, "CommentairesATMEN_INFOATM")
Info_type.Repa = Arraypos(Infos_tab, "Reparabilite_INFOATM")
Info_type.HARD_SOFT = Arraypos(Infos_tab, "HARD/SOFT")
Info_type.Poids = Arraypos(Infos_tab, "Poids_INFOATM")
Info_type.Dureevie = Arraypos(Infos_tab, "Durée de vie limitée générique_INFOATM")
Info_type.Designation = Arraypos(Infos_tab, "DesignationATM_INFOATM")

End Sub
Public Sub Cadrage_Taches_ATM(Taches_tab As Variant)
' Macro de cadrage de la feuille Info ATM de la BDD
' Il est possible de changer la disposition mais pas la nomination des colonnes

'On récupère les numéros de colonne
Taches_type.dtr = Arraypos(Taches_tab, "DTR_TACHEATM")
Taches_type.langue = Arraypos(Taches_tab, "Langue_TACHEATM")
Taches_type.rang = Arraypos(Taches_tab, "Rang_TACHEATM")
Taches_type.HARD_SOFT = Arraypos(Taches_tab, "HARD/SOFT")
Taches_type.categorie = Arraypos(Taches_tab, "Categorie_TACHEATM")
Taches_type.Niveau_Qualif = Arraypos(Taches_tab, "Niveau qualif_TACHEATM")
Taches_type.description = Arraypos(Taches_tab, "Description_TACHEATM")
Taches_type.Critere_Dim = Arraypos(Taches_tab, "CritèresDimensionnels_TACHEATM")
Taches_type.OutillageSpe = Arraypos(Taches_tab, "OutillageSpécifique_TACHEATM")
Taches_type.OutillageStd = Arraypos(Taches_tab, "OutillageStandard_TACHEATM")
Taches_type.PieceRechangeCons = Arraypos(Taches_tab, "PiècedeRechangeConsommable_TACHEATM")
Taches_type.NbOperateur = Arraypos(Taches_tab, "NbOpérateurs_TACHEATM")
Taches_type.TpsEstime = Arraypos(Taches_tab, "TempsEstiméHeures_TACHEATM")

End Sub

Public Sub Cadrage_EMMST(EMMST_tab As Variant)
'Macro de cadrage de la feuille EMMST de la BDD
'Il est possible de changer la disposition mais pas la nomination des colonnes

'On récupère les numéros de colonne
EMMST_type.dtr = Arraypos(EMMST_tab, "DTR_EMMST")
EMMST_type.DimNuLongueur = Arraypos(EMMST_tab, "DimensionNuLongueur_EMMST")
EMMST_type.DimNuLargeur = Arraypos(EMMST_tab, "DimensionNuLargeur_EMMST")
EMMST_type.DimNuHauteur = Arraypos(EMMST_tab, "DimensionNuHauteur_EMMST")
EMMST_type.DimNuPoids = Arraypos(EMMST_tab, "DimensionNuPoids")
EMMST_type.DimCondLongueur = Arraypos(EMMST_tab, "DimensionConditionnéLongueur_EMMST")
EMMST_type.DimCondLargeur = Arraypos(EMMST_tab, "DimensionConditionnéLargeur_EMMST")
EMMST_type.DimCondHauteur = Arraypos(EMMST_tab, "DimensionConditionnéHauteur_EMMST")
EMMST_type.DimCondPoids = Arraypos(EMMST_tab, "DimensionConditionnéPoids_EMMST")
EMMST_type.DatePeremption = Arraypos(EMMST_tab, "DatePeremption_EMMST")
EMMST_type.UniteCond = Arraypos(EMMST_tab, "UnitéConditionnement_EMMST")
EMMST_type.CommEmballage = Arraypos(EMMST_tab, "CommentaireEmballage_EMMST")
EMMST_type.CommMarquage = Arraypos(EMMST_tab, "CommentaireMarquage_EMMST")
EMMST_type.CommManutention = Arraypos(EMMST_tab, "CommentaireManutention_EMMST")
EMMST_type.CommStockage = Arraypos(EMMST_tab, "CommentaireStockage_EMMST")
EMMST_type.CommTransport = Arraypos(EMMST_tab, "CommentaireTransport_EMMST")
EMMST_type.TypeOutil = Arraypos(EMMST_tab, "TypeOutil_EMMST")
EMMST_type.DureePeremption = Arraypos(EMMST_tab, "DuréePéremption_EMMST")
EMMST_type.Reutilisable = Arraypos(EMMST_tab, "Réutilisable_EMMST")
EMMST_type.EtiquetteEmballage = Arraypos(EMMST_tab, "EtiquetteEmballage_EMMST")
EMMST_type.MarquageEquipement = Arraypos(EMMST_tab, "MarquageEquipement_EMMST")
EMMST_type.ManutentionEquipement = Arraypos(EMMST_tab, "ManutentionEquipement_EMMST")
EMMST_type.OutSpeManutention = Arraypos(EMMST_tab, "OutillageSpécifiqueManutention_EMMST")
EMMST_type.PrecautionManutention = Arraypos(EMMST_tab, "PrécautionManutention_EMMST")
EMMST_type.PositionStockage = Arraypos(EMMST_tab, "PositionStockage_EMMST")
EMMST_type.StockageExterieur = Arraypos(EMMST_tab, "StockageExtérieur_EMMST")
EMMST_type.StockageSousAbri = Arraypos(EMMST_tab, "StockageSousAbri_EMMST")
EMMST_type.StockageAirLibre = Arraypos(EMMST_tab, "StockageAirLibre_EMMST")
EMMST_type.StockageHorsPoussiere = Arraypos(EMMST_tab, "StockageHorsPoussière_EMMST")
EMMST_type.StockageHorsEau = Arraypos(EMMST_tab, "StockageHorsEau_EMMST")
EMMST_type.StockageInterieur = Arraypos(EMMST_tab, "StockageIntérieur_EMMST")
EMMST_type.StockageChaud = Arraypos(EMMST_tab, "StockageChaud_EMMST")
EMMST_type.StockageFroid = Arraypos(EMMST_tab, "StockageFroid_EMMST")
EMMST_type.StockageHorsLumiere = Arraypos(EMMST_tab, "StockageHorsLumière_EMMST")
EMMST_type.Gerbable = Arraypos(EMMST_tab, "Gerbable_EMMST")
EMMST_type.MoyenStockage = Arraypos(EMMST_tab, "MoyenStockage_EMMST")
EMMST_type.PrecautionStockage = Arraypos(EMMST_tab, "PrécautionStockage_EMMST")
EMMST_type.MaintenanceStockage = Arraypos(EMMST_tab, "MaintenanceStockage_EMMST")
EMMST_type.MoyenTransport = Arraypos(EMMST_tab, "MoyenTransport_EMMST")
EMMST_type.OutillageSpeTransport = Arraypos(EMMST_tab, "OutillageSpecifiqueTransport_EMMST")
EMMST_type.PrescriptionTransport = Arraypos(EMMST_tab, "PrescriptionTransport_EMMST")
EMMST_type.PrecautionTransport = Arraypos(EMMST_tab, "PrécautionTransport_EMMST")
EMMST_type.VersionEMMST = Arraypos(EMMST_tab, "Version_EMMST")
EMMST_type.DatePeremption_EN = Arraypos(EMMST_tab, "DatePeremptionEN_EMMST")
EMMST_type.UniteCond_EN = Arraypos(EMMST_tab, "UnitéConditionnéEN_EMMST")
EMMST_type.CommEmballage_EN = Arraypos(EMMST_tab, "CommentaireEmballageEN_EMMST")
EMMST_type.CommMarquage_EN = Arraypos(EMMST_tab, "CommentaireMarquageEN_EMMST")
EMMST_type.CommManutention_EN = Arraypos(EMMST_tab, "CommentaireManutentionEN_EMMST")
EMMST_type.CommStockage_EN = Arraypos(EMMST_tab, "CommentaireStockageEN_EMMST")
EMMST_type.CommTransport_EN = Arraypos(EMMST_tab, "CommentaireTransportEN_EMMST")


End Sub

Public Sub Cadrage_ProjetSLI()
' Macro de cadrage de la feuille Info ATM de la BDD
' Il est possible de changer la disposition mais pas la nomination des colonnes

'Déclaration des variables
Dim BDD As Workbook
Dim Projets_SLI As Worksheet
Dim Projets_tab As Variant
 
'Stockage du classeur macro dans une variable
Set BDD = Workbooks("BDD excel original V03.xlsm")
'On stocke La feuille contenant les données dans une variable
Set Taches_ATM = BDD.Worksheets("Projets_SLI")

Projets_tab = Taches_ATM.Range("projets_SLI[#All]")

'On récupère les numéros de colonne


End Sub



