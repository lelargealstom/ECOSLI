'--------------------------------------------------------------------------------------------------------
'
'                                       Generateur Livrables SLI
'                                       Module Z_Variables
'                                       Developpement : LE LARGE Thomas (LGM/Alstom)
'                                       Module regroupant les déclarations de variables publiques
'
'
'-----------------------------------------------------------------------------------------------------------
Option Explicit
Option Compare Text
'variables VBA

Public Espacetravail As DAO.Workspace

'Déclarations des classeurs utilisés
Public Generateur As Workbook
Public Template As Workbook
Public wb As Workbook
Public wb1 As Workbook

'Déclarations des feuilles utilisées
Public DS1 As Worksheet
Public ws As Worksheet
Public index As Worksheet
Public CoverPage As Worksheet
Public dEst As Worksheet
Public Template_EMMST As Worksheet
Public Template_HARD As Worksheet
Public Template_SOFT As Worksheet
'Déclarations des feuilles Providor
Public GP As Worksheet
Public BS As Worksheet
Public ML As Worksheet
Public CM As Worksheet
Public PM As Worksheet
Public CS As Worksheet
Public TL As Worksheet
Public SL As Worksheet
Public LCC As Worksheet

'Déclarations des bibliothèques de mapping BDD
Public MapProjetsSLI As Variant
Public Champs_ProjetsSLI As Object
Public Champs_Dates As Object
Public Champs_ProjetsASL As Object
Public Champs_InfoATM As Object
Public Champs_TachesATM As Object
Public Champs_EMMST As Object
Public Champs_Produit As Object
Public ChampControl, ctrlName As String
Public val As String
'Constantes d'index de sections
Global Const IDX_GENERAL_START = 1
Global Const IDX_GENERAL_END = 23
Global Const IDX_INTERLOCUTEURS_START = 24
Global Const IDX_INTERLOCUTEURS_END = 48
Global Const IDX_DATES_START = 49
Global Const IDX_DATES_END = 79

'variables barre de progression
Public bar_tot As Long
Public bar_cpt As Long
Public bar_prg As Variant
Public bar_temp As Long
'Variables de tableaux des données PBS
Public tDTR As Variant
Public tDES_PBS As Variant
Public tDES_BDD As Variant
Public tDES_MEF As Variant
Public tSS As Variant
Public tMTBF_gen As Variant
Public tMTBF_fct As Variant
Public tMTBF_Intr As Variant
Public tDureevie_Gen_HARD As Variant
Public tDureevie_Gen_SOFT As Variant
Public tDureevie_Ana_HARD As Variant
Public tDureevie_Ana_SOFT As Variant
Public tTps_CorrPrev_Gen_HARD As Variant
Public tTps_CorrPrev_Ana_HARD As Variant
Public tTps_Prev_Gen_SOFT As Variant
Public tTps_Prev_Ana_SOFT As Variant
Public tNTI As Variant
Public tHARD_SOFT As Variant
Public tQTE As Variant
Public tNodeType As Variant
Public tPERIODICITE As Variant
Public tPrix_Cat_Spare As Variant
Public tPrix_Cat_Repair As Variant
Public tPrix_Manuel_Spare As Variant
Public tPrix_Manuel_Repair As Variant
Public tPrix_SpareProjet_Archive As Variant
Public tPrix_RepairProjet_Archive As Variant
Public tPrix_FRS_Spare As Variant
Public tPrix_FRS_Repair As Variant
Public Infos_tab, Taches_tab, EMMST_tab As Variant

' ID_PBS - tableau d'identifiants incrementaux 1..N pour chaque ligne du
' tableau Providor de DATA_ASL. Alimente dans RecupererDonneesPBS_Providor.
' Resout le cas des DTR dupliques dans la PBS (ex. "TBD" avec valeurs
' differentes, ou meme DTR dans 2 sous-systemes) : chaque ligne PBS a un
' ID unique utilise pour le remplissage des livrables, en remplacement
' de la cle DTR. Depuis 2026-11-09.
Public tID_PBS As Variant

' Mappings volatiles Ligne feuille livrable -> ID_PBS, peuples lors de la
' construction BS/ML/CM. Clef = numero de ligne Excel (Long), Valeur = ID_PBS
' (Long). Utilises par RemplirMTBF_Corr, RemplirTempsReparation_Corr,
' RemplirPrixRepair, RemplirPrixSpare. Reinitialises a chaque Main_Providor.
Public dictBSRowToIDPBS As Object
Public dictCMRowToIDPBS As Object

' Index DTR -> ligne pour acces O(1) dans RechercherPlagesDTR_V1/V2.
' Construits une seule fois dans RecupererDonneesBDD_ATM (B_ATM_Commun).
' Si l'indexation echoue (dico Nothing ou cle absente), RechercherPlages
' retombe automatiquement sur une boucle lineaire (safe fallback).
Public dictTachesPMB As Object   ' cle = DTR + "|" + langue -> 1ere ligne dans Taches_tab
Public dictInfosPMJ As Object    ' cle = DTR -> ligne dans Infos_tab

'Profil User de l'environnement Excel
Public UserProfile As Integer
Public versionDisponible As String
'Variables specifiques SLI
Public Code_Projet As String
Public Projet_Abrev As String
Public gate As String
Public langue As String
Public DureeCirculation As Variant
Public Revision As String
Public Version_ATM As Integer
Public Auteur As String
Public Nv_Maint_Projet As String

'Variables spécifiques ATM / EMMST
' =========================================================================
' CONSTANTES DE CONFIGURATION
' =========================================================================
Public Const BIBLIO_PATH As String = "\Alstom\DIS Service VBN Pole ISR - Operations\003-03 Produits\Bibliothèques ATM2.0 illustrées"
Public Const TEMPLATE_V2_FR_PATH As String = "\Alstom\DIS Service VBN Pole ISR - Qualité\005-03 Templates\02_Phase Projet\02. ASL\ATM\ATM 2.0 VPF (exemple Baden)\ATMC2.0_FR-Template.xlsx"
Public Const TEMPLATE_V2_EN_PATH As String = "\Alstom\DIS Service VBN Pole ISR - Qualité\005-03 Templates\02_Phase Projet\02. ASL\ATM\ATM 2.0 VPF (exemple Baden)\ATMC2.0_EN-Template.xlsx"
Public Const TEMPLATE_V1_FR_PATH As String = "\Alstom\DIS Service VBN Pole ISR - Qualité\005-03 Templates\02_Phase Projet\02. ASL\ATM\ATM v1 VBN (MEQ)\ATM_FR-Template VBN v2.xlsx"
Public Const TEMPLATE_V1_EN_PATH As String = "\Alstom\DIS Service VBN Pole ISR - Qualité\005-03 Templates\02_Phase Projet\02. ASL\ATM\ATM v1 VBN (MEQ)\MTA_EN-Template VBN v2.xlsx"
Public Const PROVIDOR_TEMPLATE_PATH As String = "\Alstom\DIS Service VBN Pole ISR - Qualité\005-03 Templates\02_Phase Projet\02. ASL\Providor - LCC\Providor v5.021.xlsb"
'chaine de connexion BDD Excel
Global Const CONNEXION_EXCEL_BDD_CONSULTATION = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=UserProfile\Alstom\DIS Service VBN Pole ISR - Outils Métier ISR\01. BDD SLI\01_BDD_Consultation_Alstom\BDD excel original V4.1.xlsb;Extended Properties='Excel 12.0 Macro;HDR=Yes;IMEX=1';"
Global Const CONNEXION_EXCEL_BDD_TRAVAIL = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=UserProfile\Alstom\DIS Service VBN Pole ISR - Outils Métier ISR\01. BDD SLI\02_BDD_Prod_LGM\BDD excel original V4.1.xlsb;Extended Properties='Excel 12.0 Macro;HDR=Yes;IMEX=1';"
Global Const CONNEXION_SHAREPOINT_LIST = "Provider=Microsoft.ACE.OLEDB.12.0;WSS;IMEX=0;RetrieveIds=Yes;DATABASE=https://alstomgroup.sharepoint.com/sites/FR_VBN_obso_services;LIST={ListeSharePoint};"
Global Const vbOffline As Long = &H1000

' Lignes d'insertion par rang (ATM V1)
Public Const LIGNE_RANG_1 As Integer = 14
Public Const LIGNE_RANG_2 As Integer = 17
Public Const LIGNE_RANG_3 As Integer = 20
Public Const LIGNE_RANG_4 As Integer = 23
Public Const LIGNE_RANG_5 As Integer = 26
' Mapping Rang -> Chapitre (ATM V2)
Public Const RANG_1_CHAPITRE As String = "A"
Public Const RANG_2_CHAPITRE As String = "B"
Public Const RANG_3_4_CHAPITRE As String = "C"
Public Const RANG_5_CHAPITRE As String = "D"

' =========================================================================
' VARIABLES SPÉCIFIQUES ATM
' =========================================================================
' STRUCTURE POUR GÉRER UN CHAPITRE
Public Type ChapitreInfo
    Marqueur As String          ' "A", "B", "C", "D"
    ligneDebut As Integer       ' Ligne du marqueur dans la feuille
    LigneProchainChap As Integer ' Ligne du chapitre suivant
    TotalTemps As Double        ' Somme des temps du chapitre
    NombreInsertions As Integer ' Compteur de lignes insérées
End Type
Public test_taches, test_infos, MTBF_pres, Duree_vie_HARD_pres, Duree_vie_SOFT_pres, RAMReportDATA, trouve, Presence_BDD, test, AvantCGR As Boolean
Public system, system_temp, designationfinal As String
Public estnum, TabIndex As Variant
Public cpt, PMH, PMB, PMJ, nvx, ligvide, ligcat, ligcatsup, DL1, DL2 As Integer
Public ligne, incr, compteur As Integer
' =========================================================================
' VARIABLES SPÉCIFIQUES EMMST
' =========================================================================
Public spare, Tool, designation, dtr, supplier_id, supplier_name, blength, bwidth, bheight, bmass As Range
Public plength, pwidth, pheight, pmass, lapsing_no, lapsing_yes, expiry_date, multiple_packaging As Range
Public non_reusable_box, reusable_box, pack_label_no, pack_label_yes, equip_label_no, equip_label_yes As Range
Public handling_mean_no, handling_mean_yes, spe_tool_no, spe_tool_yes, man_prec_no, man_prec_yes As Range
Public pos_hor, pos_ver, pos_indif, out, out_shelter, out_air, out_dust, out_water, ind, ind_hot, ind_cold, ind_dust, ind_water, ind_light As Range
Public stackable_no, stackable_yes, stor_mean_no, stor_mean_yes, stor_prec_no, stor_prec_yes, stor_maint_no, stor_maint_yes As Range
Public transp_mean_no, transp_mean_yes, transp_tool_no, transp_tool_yes, transp_req_no, transp_req_yes, transp_prec_no, transp_prec_yes As Range
Public pack_com, lab_com, man_com, stor_com, transp_com As Range

' =========================================================================
' VARIABLES SPÉCIFIQUES PROVIDOR
' =========================================================================
Public nextIdx As Long
Public NodeLevel As Long
Public BoolMaintPrev As Boolean
Public addFamily As Boolean
' Dernières lignes feuilles Providor
Public dlBS As Integer
Public dlML As Integer
Public dlCM As Integer
Public dlPM As Integer
Public dlCU As Integer
Public dlGen As Integer
' Coefficients prix fournisseur
Public Coeff_Price_FRS_Spare As Double
Public Coeff_Price_FRS_Repair As Double
' Chemins et fichiers
Public plage As String
Public CheminComplet As String
Public CheminTemplate As String
Public fichierComplet As String
Public nomFichier As String
' Sous-systèmes
Public curSS As String
Public lastSS As String
' Tableaux Material List
Public DTR_tab_ML As Variant
Public Des_tab_ML As Variant
Public LIST_NODE As Variant
'Variables d'objets
Public fso As Object
Public ctrl As MSForms.control
Public Repertoire As FileDialog
Public msg As String
Public key As Variant
Public ai As addin
Public found As Boolean
'Variables de msgbox
Public resp As VbMsgBoxResult
Public respSave As VbMsgBoxResult

'Public Base As ADODB.Connection
Public Base As Object
' Timeout des requêtes ODBC
Global Const STANDARD_TIMEOUT = 120 ' en standard 2 minutes
' lors de fonctions longues commela maj des articles 1 heure
Global Const LONG_TIMEOUT = 3600 ' lors de fonctions longues comme parts synchro

'AR202402: introduction des variables 'color' identique à Tomsys
Global Const COULEUR_NOIR = 0
Global Const COULEUR_BLANC = 16777215
Global Const COULEUR_VERT = 65280 'vert fluo '311586 'vert foncé
Global Const COULEUR_VERT_F = 311586 'vert foncé
Global Const COULEUR_VERT_C = 10092441 'vert clair
Global Const COULEUR_JAUNE = 65535
Global Const COULEUR_JAUNE_C = 6750207 'jaune pâle
Global Const COULEUR_ORANGE = 32255
Global Const COULEUR_ROUGE = 2366701
Global Const COULEUR_BLEU = 15773696 ' bleu ciel '16776960 'bleu clair
Global Const COULEUR_BLEU_F = 10485792 ' utilisé dans les textes
Global Const COULEUR_VIOLET = 10498160
Global Const COULEUR_VIOLET_C = 16751103
Global Const COULEUR_GRIS = 9474192
Global Const COULEUR_GRIS_C = 14277081

'pour les requetes select
Public STRSEL As String
Public STRSEL1 As String
Public STRSEL2 As String
Public STRSEL3 As String
Public STRSEL4 As String
Public REQSEL As Object
Public REQSEL1 As Object
Public REQSEL2 As Object
'pour les requetes insert
Public STRINS As String
'pour les requetes update
Public STRUPD As String
'pour les requetes delete
Public STRDEL As String 'pour les requetes delete


'ID des différentes Listes SharePoint
Global Const INFO_ATM_List = "997ff1da-1ab7-4135-b257-2b6ebdad2549"

'accès aux fichiers
Public cheminbase           As String
Global xlApp                As Excel.Application

'Variables specifiques utilisateur
Public chemin As String
Public userpath As String
Public mode_utilisation  As Variant
Public nomfichierchoisi As String
Public nomdossierchoisi As String

'Variables de boucles
Public h, i, j, k, l As Integer

'variables extract :
Public recherchecasempl
Public typedonnée As Variant
Public ligdeb, ligfin, ligref, ligfinonglet, ligresult
Public nblevel As Variant 'peut être "ALL" ou un nombre
Public colrefinit, colfinPREPA, coldebART, colfinART, coldebSOU, colfinSOU, coldebPDN, colfinPDN, coldebSTK, colfinSTK, coldebCAS, coldebQTY, colfinQTY, colfinCAS, colPRP, coldebPRD, colinterPRD, colfinPRD As Variant
'option REQSELs sur onglet actif
Public resultsuronglet As Boolean
Public coldebonglet, colfinonglet As Variant
'gestion onglets et selection
Public fichier, fichierinit, newfichier, onglet, ongletinit, newonglet, refcherchée As String
Public annulé, couleur, projetactif, projetinactif, présenceartart As Boolean
Public problem As Boolean
'variables colonnes utilisé partout
Public vbsPath As String
Public nbligtot As Long '  nbligtot = Columns(1).Rows.Count 'initialisé dans START
Public nbcoltot As Long '  nbcoltot = Rows(1).Columns.Count 'initialisé dans START
' ------------------------------------------------------------------
' Colonnes des différents Onglets BDD
' Infos ATM
'Constuction d'un type personnalisé pour le stockage des numéros de lignes et de colonnes pour l'onglet Infos_ATM
Public Type Info_Col

    dtr As Integer
    DescriptifFR As Integer
    LocalisationFR As Integer
    Plan As Integer
    CommentaireFR As Integer
    version As Integer
    DescriptifEN As Integer
    CommentaireEN As Integer
    LocalisationEN As Integer
    Repa As Integer
    HARD_SOFT As Integer
    Poids As Integer
    Dureevie As Integer
    designation As Integer

End Type

Public Info_type As Info_Col

' Taches ATM
'Constuction d'un type personnalisé pour le stockage des numéros de lignes et de colonnes pour l'onglet Taches_ATM
Public Type Taches_Col

    dtr As Integer
    langue As Integer
    rang As Integer
    HARD_SOFT As Integer
    categorie As Integer
    Niveau_Qualif As Integer
    description As Integer
    Critere_Dim As Integer
    OutillageSpe As Integer
    OutillageStd As Integer
    PieceRechangeCons As Integer
    NbOperateur As Integer
    TpsEstime As Integer
    TpsCumulee As Integer

End Type

Public Taches_type As Taches_Col

'EMMST
'Construction d'un type personnalisé pour le stockage des numéros de lignes et de colonne pour l'onglet EMMST
Public Type EMMST_Col

    dtr As Integer
    DimNuLongueur As Integer
    DimNuLargeur As Integer
    DimNuHauteur As Integer
    DimNuPoids As Integer
    DimCondLongueur As Integer
    DimCondLargeur As Integer
    DimCondHauteur As Integer
    DimCondPoids As Integer
    DatePeremption As Integer
    UniteCond As Integer
    CommEmballage As Integer
    CommMarquage As Integer
    CommManutention As Integer
    CommStockage As Integer
    CommTransport As Integer
    TypeOutil As Integer
    DureePeremption As Integer
    Reutilisable As Integer
    EtiquetteEmballage As Integer
    MarquageEquipement As Integer
    ManutentionEquipement As Integer
    OutSpeManutention As Integer
    PrecautionManutention As Integer
    PositionStockage As Integer
    StockageExterieur As Integer
    StockageSousAbri As Integer
    StockageAirLibre As Integer
    StockageHorsPoussiere As Integer
    StockageHorsEau As Integer
    StockageInterieur As Integer
    StockageChaud As Integer
    StockageFroid As Integer
    StockageHorsLumiere As Integer
    Gerbable As Integer
    MoyenStockage As Integer
    PrecautionStockage As Integer
    MaintenanceStockage As Integer
    MoyenTransport As Integer
    OutillageSpeTransport As Integer
    PrescriptionTransport As Integer
    PrecautionTransport As Integer
    VersionEMMST As Integer
    UniteCond_EN As Integer
    DatePeremption_EN As Integer
    CommEmballage_EN As Integer
    CommMarquage_EN As Integer
    CommManutention_EN As Integer
    CommStockage_EN As Integer
    CommTransport_EN As Integer
   
End Type

Public EMMST_type As EMMST_Col

'Generateur Livrables
' Taches ATM
'Constuction d'un type personnalisé pour le stockage des numéros de lignes et de colonnes pour l'onglet Taches_ATM
Public Type Projets_SLI_Col

    Nom_Projet As Integer
    Projet_Offre As Integer
    CT_Code_Projet As Integer
    Type_Train As Integer
    Date_debut_production As Integer
    Date_fin_production As Integer
    Date_circulation As Integer
    Nombre_train As Integer
    Servicing_day_byyear As Integer
    PH_byyear As Integer
    RH_byyear As Integer
    Km_byyear As Integer
    Duree_garantie_months As Integer
    devise As Integer
    Man_hour_cost As Integer
    Niveau_maintenance As Integer
    Nb_DTR_PBS As Integer
    Nb_FRS_CDRL As Integer
    Type_Pack As Integer
    UO_grand_pere As Integer
    Statut_Projet As Integer
    Nb_GDP_classique As Integer
    Nb_GDP_FRS As Integer
    Pourc_GDP_classique As Integer
    Pourc_GDP_classique_cons As Integer
    Pourc_GDP_FRS As Integer
    Pourc_GDP_FRS_cons As Integer
    R_client_Acceptance As Integer
    R_client_Clausing As Integer
    R_client_MoP As Integer
    R_FRS_Clausing As Integer
    PSLI_preli As Integer
    PSLI_final As Integer
    Providor_preli As Integer
    Providor_final As Integer
    Providor_update As Integer
    ATM_preli As Integer
    ATM_final As Integer
    EMMST_preli As Integer
    EMMST_final As Integer
    MUL_preli As Integer
    MUL_final As Integer
    MC8 As Integer
    LCC_CCV0 As Integer
    LCC_CCAP As Integer
    Dossier_justi As Integer
    Date_Check_OdyC As Integer
    Date_NTP As Integer
    Date_LGR_VBN As Integer
    Date_LGR_LU As Integer
    Date_SGR_VBN As Integer
    Date_SGR_LU As Integer
    Date_BA1 As Integer
    Date_BA2 As Integer
    Date_BA3 As Integer
    Date_BA4 As Integer
    Date_BA5 As Integer
    Date_IGR_VBN As Integer
    Date_IGR_LU As Integer
    Date_DD1 As Integer
    Date_DD2 As Integer
    Date_DD3 As Integer
    Date_DD4 As Integer
    Date_DD5 As Integer
    Date_PGR_VBN As Integer
    Date_PGR_LU As Integer
    Date_DF1 As Integer
    Date_DF2 As Integer
    Date_DF3 As Integer
    Date_DF4 As Integer
    Date_DF5 As Integer
    Date_CGR_VBN As Integer
    Date_CGR_LU As Integer
    Date_FAI1 As Integer
    Date_FAI2 As Integer
    Date_FAI3 As Integer
    Date_FAI4 As Integer
    Date_FAI5 As Integer
    Date_FEI_VBN As Integer
    Date_GFV_LU As Integer
    Date_IQA_VBN As Integer
    Date_IQA_LU As Integer
    Date_FQA_LU As Integer
    RSLI As Integer
    RASL As Integer
    LU As Integer
    Interloc_LU As Integer
    PM_VBN As Integer
    PreM_VBN As Integer
    PrPM As Integer
    PMM As Integer
    Procurement As Integer
    DL As Integer
    RAM As Integer
    Autres_Interloc As Integer
    Ref_PBS As Integer
    Ref_RAM_Report As Integer
    Version_providor As Integer
    Prix_formation_PACIS As Integer
    Prix_formation_IOB As Integer
    Prix_formation_APC As Integer
    Prix_formation_DAS As Integer
    Prix_formation_SRS As Integer
    Prix_formation_EVR As Integer
    Prix_formation_Lecteur_badge As Integer
    Prix_formation_GSM As Integer
    BTE As Integer
    Ref_BTE As Integer
    Prix_BTE As Integer
    Date_installation_BTE As Integer
    Date_der_visite_BTE As Integer
    Duree_garantie_BTE As Integer
    MCO As Integer
    Duree_MCO As Integer
    Prix_MCO As Integer
    OBSO As Integer
    Duree_couv_OBSO As Integer
    Prix_OBSO As Integer
    LCC_valeur_total As Integer
    LCC_valeur_TCMS As Integer
    LCC_valeur_Traction_BMS As Integer
    LCC_valeur_N_T As Integer
    LCC_valeur_PACIS As Integer
    LCC_valeur_IOB As Integer
    LCC_valeur_APC As Integer
    LCC_valeur_DAS As Integer
    LCC_valeur_SRS As Integer
    LCC_valeur_EVR As Integer
    LCC_valeur_Lecteur_badge As Integer
    LCC_valeur_GSM As Integer
    
End Type

Public Projets_SLI_type As Projets_SLI_Col

' =========================================================================
' TYPE STRUCTURÉ - BREAKDOWN STRUCTURE
' =========================================================================
Public Type bs_col
    NodeLevel As Integer
    NodeIndex As Integer
    NodeName As Integer
    ListNode As Integer
    Quantity As Integer
    Update As Integer
    PartNumber As Integer
    NodeType As Integer
    MaintenanceLevel As Integer
End Type

' =========================================================================
' TYPE STRUCTURÉ - MATERIAL LIST
' =========================================================================
Public Type ml_col
    Update As Integer
    Name As Integer
    PartNumber As Integer
    Cost As Integer
    CostSource As Integer              ' Colonne P : CostValidityDate avec date source + commentaires sources : ex "D&IS Catalogue v25_05"
    CurrencyUnit As Integer
    MaterialType As Integer
End Type

' =========================================================================
' TYPE STRUCTURÉ - CORRECTIVE MAINTENANCE
' =========================================================================
Public Type cm_col
    NodeName As Integer
    NodeIndex As Integer
    SequenceNo As Integer
    FailureName As Integer
    TamponNodeLevel As Integer
    TamponDTR As Integer
    TamponNodeName As Integer
    TamponNodeType As Integer
    FailureRate As Integer
    FailureMultiplier As Integer
    FailureRateUnit As Integer
    CMNo As Integer
    TaskDescription As Integer
    OnOff As Integer
    RepairTime As Integer
    TimeSource As Integer              ' Colonne Procedure No
    MaintenanceLevel As Integer
    SkillName As Integer
    SkillQty As Integer
    MaterialName As Integer
    MaterialQty As Integer
    ToolName As Integer
    ToolQty As Integer
    JustificationTaskFR As Integer
    JustificationTaskEN As Integer
    MTBF As Integer
    MTBFUnit As Integer
    MTBFSource As Integer              ' Commentaires Alstom Taux Défaillance DL
End Type

' =========================================================================
' TYPE STRUCTURÉ - PREVENTIVE MAINTENANCE
' =========================================================================
Public Type pm_col
    NodeIndex As Integer
    SequenceNo As Integer
    TaskDescription As Integer
    PrevMaintType As Integer
    PackageReference As Integer
    OnOff As Integer
    FrequencyValue As Integer
    FrequencyUnit As Integer
    PeriodicCancelledOcc As Integer
    TaskTime As Integer
    TimeSource As Integer
    SkillName As Integer
    SkillQty As Integer
    MaterialName As Integer
    MaterialQty As Integer
    JustificationTaskFR As Integer     ' Colonne EE et EF (FR, EN meme valeur)
    JustificationTaskEN As Integer
End Type

' =========================================================================
' TYPE STRUCTURE - CONTROL SHEET
' =========================================================================
Public Type cs_col
    Content As Integer
    datecontrolsheet As Integer
    Author As Integer
    Issue As Integer
End Type

' =========================================================================
' TYPE STRUCTURÉ - SKILL LIST
' =========================================================================
Public Type sl_col
    Update As Integer
    SkillCategory As Integer
    ProfessionCategory As Integer
    ManHourCost As Integer
    CurrencyUnit As Integer
End Type

' =========================================================================
' TYPE STRUCTURÉ - TOOL LIST
' =========================================================================
Public Type tl_col
    Update As Integer
    Name As Integer
    PartNumber As Integer
    AcquisitionCost As Integer
    Currency As Integer
    ToolType As Integer
End Type

' =========================================================================
' TYPE STRUCTURÉ - GENERAL PARAMETERS
' =========================================================================
Public Type gp_col
    OpUnits_Description As Integer
    OpUnits_Acronym As Integer
    OpUnits_YearlyValue As Integer
    OpUnits_Unit As Integer
    CurrUnit_Currtablename As Integer
    CurrUnit_Currname As Integer
    CurrUnit_Currabbrev As Integer
    CurrUnit_changerate As Integer
    CurrUnit_datereviewtoolrate As Integer
    CurrUnit_pjctchangerate As Integer
    CurrUnit_datereviewpjctrate As Integer
End Type

Public Type gp_lig
    Pjct_number As Integer
    Pjct_name As Integer
    Pjct_acronym As Integer
    Pjct_productname As Integer
    LCC_nbassets As Integer
    LCC_length As Integer
    LCC_warranty As Integer
End Type

' =========================================================================
' INSTANCES PUBLIQUES DES TYPES
' =========================================================================
Public bs_col As bs_col
Public ml_col As ml_col
Public cm_col As cm_col
Public pm_col As pm_col
Public cs_col As cs_col
Public sl_col As sl_col
Public tl_col As tl_col
Public gp_col As gp_col
Public gp_lig As gp_lig
' ---------------------------------------------------------------------------------
