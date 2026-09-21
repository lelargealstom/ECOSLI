Attribute VB_Name = "U_Mapping"
'--------------------------------------------------------------------------------------------------------
'
'                                       Generateur Livrables SLI
'                                       Module H_Mapping
'                                       Developpement : LE LARGE Thomas (LGM/Alstom)
'                                       Module permettant d'effectuer le mapping des données
'
'
'-----------------------------------------------------------------------------------------------------------
Option Explicit
'Table : Projets_SLI
Public Function GetChampsProjetsSLI() As Object

    Set Champs_ProjetsSLI = CreateObject("Scripting.Dictionary")

    ' Page General : 1 à 13
    Champs_ProjetsSLI.Add "Nom projet_PROJET SLI", "Gen_NomProjet"
    Champs_ProjetsSLI.Add "Projet / Offre_PROJET SLI", "Gen_Type"
    Champs_ProjetsSLI.Add "Region_PROJET SLI", "Gen_Region"
    Champs_ProjetsSLI.Add "CT Code_PROJET SLI", "Gen_CTCode"
    Champs_ProjetsSLI.Add "Type de train_PROJET SLI", "Gen_TypeTrain"
    Champs_ProjetsSLI.Add "Client final_PROJET SLI", "Gen_Customer"
    Champs_ProjetsSLI.Add "Date de debut de production_PROJET SLI", "Gen_DateDébut"
    Champs_ProjetsSLI.Add "Date de fin de production_PROJET SLI", "Gen_DateFin"
    Champs_ProjetsSLI.Add "Type de pack (complexe/standard/mini)_PROJET SLI", "Gen_TypePack"
    Champs_ProjetsSLI.Add "UO ""grand pere"" du pack_PROJET SLI", "Gen_UOPack"
    Champs_ProjetsSLI.Add "Statut (termine/En cours/A planifier)_PROJET SLI", "Gen_StatutUO"
    Champs_ProjetsSLI.Add "Nombre de DTR dans l EAP ou la PBS_PROJET SLI", "Gen_NbDTRPBS"
    Champs_ProjetsSLI.Add "Nombre de fournisseur soumis a CDRL_PROJET SLI", "Gen_NbFRS"

    ' Page Profil de mission : 14 à 23
    Champs_ProjetsSLI.Add "Duree de circulation_PROJET SLI", "Gen_DureeCirculation"
    Champs_ProjetsSLI.Add "Nombre de train_PROJET SLI", "Gen_NbTrains"
    Champs_ProjetsSLI.Add "Servicing day by year_PROJET SLI", "Gen_ServicingbyDay"
    Champs_ProjetsSLI.Add "PH by year_PROJET SLI", "Gen_PowerHour"
    Champs_ProjetsSLI.Add "RH by year_PROJET SLI", "Gen_RunningHour"
    Champs_ProjetsSLI.Add "Km by year_PROJET SLI", "Gen_KMbyYear"
    Champs_ProjetsSLI.Add "Duree de garantie_PROJET SLI", "Gen_DureeGarantie"
    Champs_ProjetsSLI.Add "Devise_PROJET SLI", "Gen_Devise"
    Champs_ProjetsSLI.Add "MAN hour cost_PROJET SLI", "Gen_ManHourCost"
    Champs_ProjetsSLI.Add "Niv de maintenance_PROJET SLI", "Gen_NiveauMaintenance"
    
    ' Page Interlocuteurs : 24 à 48
    ' --- Villeurbanne ---
    Champs_ProjetsSLI.Add "RSLI_PROJET SLI", "Gen_RSLI"
    Champs_ProjetsSLI.Add "RASL_PROJET SLI", "Gen_RASL"
    Champs_ProjetsSLI.Add "PM ou BM VBN_PROJET SLI", "Gen_PMBMVBN"
    Champs_ProjetsSLI.Add "PrEM ou BTM_PROJET SLI", "Gen_PrEMBTM"
    Champs_ProjetsSLI.Add "PrPM_PROJET SLI", "Gen_PrPM"
    Champs_ProjetsSLI.Add "PMM_PROJET SLI", "Gen_PMM"
    Champs_ProjetsSLI.Add "Procurement_PROJET SLI", "Gen_Procurement"
    Champs_ProjetsSLI.Add "RAM_PROJET SLI", "Gen_RAM"
    ' --- Leading Unit ---
    Champs_ProjetsSLI.Add "LU_PROJET SLI", "GEN_LU"
    'Champs_ProjetsSLI.Add "Interlocuteur LU_PROJET SLI", "GEN_InterlocuteurLU"
    Champs_ProjetsSLI.Add "ISR Manager LU_PROJET SLI", "GEN_ISRManagerLU"
    Champs_ProjetsSLI.Add "LSA Engineer LU_PROJET SLI", "GEN_LSAEngineerLU"
    Champs_ProjetsSLI.Add "Doc Manager LU_PROJET SLI", "GEN_DocManagerLU"
    Champs_ProjetsSLI.Add "System Engineer LU_PROJET SLI", "GEN_SystemEngineerLU"
    Champs_ProjetsSLI.Add "TC Engineer LU_PROJET SLI", "GEN_TCEngineerLU"
    Champs_ProjetsSLI.Add "BID Manager LU_PROJET SLI", "GEN_BIDManagerLU"
    ' --- Domain Leader ---
    Champs_ProjetsSLI.Add "DL PA_PACIS_PROJET SLI", "GEN_DLPAPACIS"
    'Champs_ProjetsSLI.Add "DL PIS_PROJET SLI", "GEN_DLPIS"
    'Champs_ProjetsSLI.Add "DL CCTV_PROJET SLI", "GEN_DLCCTV"
    Champs_ProjetsSLI.Add "DL IOB_PROJET SLI", "GEN_DLIOB"
    Champs_ProjetsSLI.Add "DL APC_PROJET SLI", "GEN_DLAPC"
    Champs_ProjetsSLI.Add "DL DAS_PROJET SLI", "GEN_DLDAS"
    Champs_ProjetsSLI.Add "DL EVR_PROJET SLI", "GEN_DLEVR"
    Champs_ProjetsSLI.Add "DL LECTEUR BADGE_PROJET SLI", "GEN_DLLECTEURBADGE"
    Champs_ProjetsSLI.Add "DL GSM REPEATER_PROJET SLI", "GEN_DLGSMREPEATER"
    Champs_ProjetsSLI.Add "DL SRS_PROJET SLI", "GEN_DLSRS"

    ' Champs Livrables : 49 à 63
    'Requis client
    Champs_ProjetsSLI.Add "Requis client : Acceptance_PROJET SLI", "Gen_RequisAcceptance"
    Champs_ProjetsSLI.Add "Requis client : Clausing_PROJET SLI", "Gen_RequisClausing"
    Champs_ProjetsSLI.Add "Requis client : MoP_PROJET SLI", "Gen_RequisMoP"
    Champs_ProjetsSLI.Add "Requis fournisseur : Clausing_PROJET SLI", "Gen_RequisFournisseurClausing"
    ' Champs PSLI et Providor
    Champs_ProjetsSLI.Add "PSLI preliminaire_PROJET SLI", "Gen_PSLIPreliminaire"
    Champs_ProjetsSLI.Add "PSLI final_PROJET SLI", "Gen_PSLIFinal"
    Champs_ProjetsSLI.Add "PROVIDOR/BASL preliminaire_PROJET SLI", "Gen_ProvidorPreliminaire"
    Champs_ProjetsSLI.Add "PROVIDOR/BASL final_PROJET SLI", "Gen_ProvidorFinal"
    Champs_ProjetsSLI.Add "PROVIDOR/BASL update_PROJET SLI", "Gen_ProvidorUpdate"
    ' Champs ATM et EMMST
    Champs_ProjetsSLI.Add "ATM preliminaire_PROJET SLI", "Gen_ATMPreliminaire"
    Champs_ProjetsSLI.Add "ATM final_PROJET SLI", "Gen_ATMFinal"
    Champs_ProjetsSLI.Add "EMMST prelimianire_PROJET SLI", "Gen_EMMSTPreliminaire"
    Champs_ProjetsSLI.Add "EMMST final_PROJET SLI", "Gen_EMMSTFinal"
    ' Champs MUL
    Champs_ProjetsSLI.Add "MUL preliminaire_PROJET SLI", "Gen_MULPreliminaire"
    Champs_ProjetsSLI.Add "MUL final_PROJET SLI", "Gen_MULFinal"
    ' Champs LCC et MC8
    'Champs_ProjetsSLI.Add "MC8_PROJET SLI", "Gen_MC8"
    'Champs_ProjetsSLI.Add "LCC CCVO_PROJET SLI", "Gen_LCC_CCVO"
    'Champs_ProjetsSLI.Add "LCC CCAP_PROJET SLI", "Gen_LCC_CCAP"
    'Champs_ProjetsSLI.Add "Dossiers de justification_PROJET SLI", "Gen_DossiersJustification"

    ' Champs Dates : 64 à 146
    ' --- Date DFQ ---
    Champs_ProjetsSLI.Add "Date de check OdyC_PROJET SLI", "Gen_DateCheckOdyC"
    Champs_ProjetsSLI.Add "Date NTP_PROJET SLI", "Gen_DateNTP"
    Champs_ProjetsSLI.Add "Date LGR VBN_PROJET SLI", "Gen_DateLGRVBN"
    Champs_ProjetsSLI.Add "Date SGR VBN_PROJET SLI", "Gen_DateSGRVBN"
    Champs_ProjetsSLI.Add "Date IGR VBN_PROJET SLI", "Gen_DateIGRVBN"
    Champs_ProjetsSLI.Add "Date PGR VBN_PROJET SLI", "Gen_DatePGRVBN"
    Champs_ProjetsSLI.Add "Date CGR (first train) LU_PROJET SLI", "Gen_DateCGRFirstTrain"
    Champs_ProjetsSLI.Add "Date CGR (full train) LU_PROJET SLI", "Gen_DateCGRFullTrain"
    Champs_ProjetsSLI.Add "Date FEI VBN_PROJET SLI", "Gen_DateFEIVBN"
    Champs_ProjetsSLI.Add "Date GFV LU_PROJET SLI", "Gen_DateGFVLU"
    Champs_ProjetsSLI.Add "Date IQA VBN_PROJET SLI", "Gen_DateIQAVBN"
    Champs_ProjetsSLI.Add "Date GSR VBN_PROJET SLI", "Gen_DateGSRVBN"
    Champs_ProjetsSLI.Add "Date FQA VBN_PROJET SLI", "Gen_DateFQAVBN"

    Set GetChampsProjetsSLI = Champs_ProjetsSLI
End Function
Public Function GetChampsDates() As Object

    Set Champs_Dates = CreateObject("Scripting.Dictionary")

    ' --- Date Inter-Gates ---
    ' ---    Date GO RFQ   ---
    Champs_ProjetsSLI.Add "Date GO RFQ PA_PACIS_PROJET SLI", "GEN_DateGORFQPAPACIS"
    Champs_ProjetsSLI.Add "Date GO RFQ PIS_PROJET SLI", "GEN_DateGORFQPIS"
    Champs_ProjetsSLI.Add "Date GO RFQ CCTV_PROJET SLI", "GEN_DateGORFQCCTV"
    Champs_ProjetsSLI.Add "Date GO RFQ IOB_PROJET SLI", "GEN_DateGORFQIOB"
    Champs_ProjetsSLI.Add "Date GO RFQ APC_PROJET SLI", "GEN_DateGORFQAPC"
    Champs_ProjetsSLI.Add "Date GO RFQ DAS_PROJET SLI", "GEN_DateGORFQDAS"
    Champs_ProjetsSLI.Add "Date GO RFQ EVR_PROJET SLI", "GEN_DateGORFQEVR"
    Champs_ProjetsSLI.Add "Date GO RFQ LECTEUR BADGE_PROJET SLI", "GEN_DateGORFQLECTEURBADGE"
    Champs_ProjetsSLI.Add "Date GO RFQ GSM REPEATER_PROJET SLI", "GEN_DateGORFQGSMREPEATER"
    Champs_ProjetsSLI.Add "Date GO RFQ SRS_PROJET SLI", "GEN_DateGORFQSRS"
    ' --- Date BA ---
    Champs_ProjetsSLI.Add "Date BA PA_PACIS_PROJET SLI", "GEN_DateBAPAPACIS"
    Champs_ProjetsSLI.Add "Date BA PIS_PROJET SLI", "GEN_DateBAPIS"
    Champs_ProjetsSLI.Add "Date BA CCTV_PROJET SLI", "GEN_DateBACCTV"
    Champs_ProjetsSLI.Add "Date BA IOB_PROJET SLI", "GEN_DateBAIOB"
    Champs_ProjetsSLI.Add "Date BA APC_PROJET SLI", "GEN_DateBAAPC"
    Champs_ProjetsSLI.Add "Date BA DAS_PROJET SLI", "GEN_DateBADAS"
    Champs_ProjetsSLI.Add "Date BA EVR_PROJET SLI", "GEN_DateBAEVR"
    Champs_ProjetsSLI.Add "Date BA LECTEUR BADGE_PROJET SLI", "GEN_DateBALECTEURBADGE"
    Champs_ProjetsSLI.Add "Date BA GSM REPEATER_PROJET SLI2", "GEN_DateBAGSMREPEATER"
    Champs_ProjetsSLI.Add "Date BA SRS_PROJET SLI", "GEN_DateBASRS" '96
    ' --- Date GO ORDER ---
    Champs_ProjetsSLI.Add "Date GO ORDER PA_PACIS_PROJET SLI", "GEN_DateGOORDERPAPACIS"
    Champs_ProjetsSLI.Add "Date GO ORDER PIS_PROJET SLI", "GEN_DateGOORDERPIS"
    Champs_ProjetsSLI.Add "Date GO ORDER CCTV_PROJET SLI", "GEN_DateGOORDERCCTV"
    Champs_ProjetsSLI.Add "Date GO ORDER IOB_PROJET SLI", "GEN_DateGOORDERIOB"
    Champs_ProjetsSLI.Add "Date GO ORDER APC_PROJET SLI", "GEN_DateGOORDERAPC"
    Champs_ProjetsSLI.Add "Date GO ORDER DAS_PROJET SLI", "GEN_DateGOORDERDAS"
    Champs_ProjetsSLI.Add "Date GO ORDER EVR_PROJET SLI", "GEN_DateGOORDERVR"
    Champs_ProjetsSLI.Add "Date GO ORDER LECTEUR BADGE_PROJET SLI4", "GEN_DateGOORDERLECTEURBADGE"
    Champs_ProjetsSLI.Add "Date GO ORDER GSM REPEATER_PROJET SLI", "GEN_DateGOORDERGSMREPEATER"
    Champs_ProjetsSLI.Add "Date GO ORDER SRS_PROJET SLI", "GEN_DateGOORDERSRS"
    ' --- Date DETAIL DESIGN ---
    Champs_ProjetsSLI.Add "Date DETAIL DESIGN PA_PACIS_PROJET SLI", "GEN_DateDETAILDESIGNPAPACIS"
    Champs_ProjetsSLI.Add "Date DETAIL DESIGN PIS_PROJET SLI", "GEN_DateDETAILDESIGNPIS"
    Champs_ProjetsSLI.Add "Date DETAIL DESIGN CCTV_PROJET SLI", "GEN_DateDETAILDESIGNCCTV"
    Champs_ProjetsSLI.Add "Date DETAIL DESIGN IOB_PROJET SLI", "GEN_DateDETAILDESIGNIOB"
    Champs_ProjetsSLI.Add "Date DETAIL DESIGN APC_PROJET SLI", "GEN_DateDETAILDESIGNAPC"
    Champs_ProjetsSLI.Add "Date DETAIL DESIGN DAS_PROJET SLI", "GEN_DateDETAILDESIGNDAS"
    Champs_ProjetsSLI.Add "Date DETAIL DESIGN EVR_PROJET SLI", "GEN_DateDETAILDESIGNEVR"
    Champs_ProjetsSLI.Add "Date DETAIL DESIGN LECTEUR BADGE_PROJET SLI", "GEN_DateDETAILDESIGNLECTEURBADGE"
    Champs_ProjetsSLI.Add "Date DETAIL DESIGN GSM REPEATER_PROJET SLI", "GEN_DateDETAILDESIGNGSMREPEATER"
    Champs_ProjetsSLI.Add "Date DETAIL DESIGN SRS_PROJET SLI", "GEN_DateDETAILDESIGNSRS"
    ' --- Date DESIGN FREEZE ---
    Champs_ProjetsSLI.Add "Date DESIGN FREEZE PA_PACIS_PROJET SLI", "GEN_DateDESIGNFREEZEPAPACIS"
    Champs_ProjetsSLI.Add "Date DESIGN FREEZE PIS_PROJET SLI", "GEN_DateDESIGNFREEZEPIS"
    Champs_ProjetsSLI.Add "Date DESIGN FREEZE CCTV_PROJET SLI", "GEN_DateDESIGNFREEZECCTV"
    Champs_ProjetsSLI.Add "Date DESIGN FREEZE IOB_PROJET SLI", "GEN_DateDESIGNFREEZEIOB"
    Champs_ProjetsSLI.Add "Date DESIGN FREEZE APC_PROJET SLI", "GEN_DateDESIGNFREEZEAPC"
    Champs_ProjetsSLI.Add "Date DESIGN FREEZE DAS_PROJET SLI", "GEN_DateDESIGNFREEZEDAS"
    Champs_ProjetsSLI.Add "Date DESIGN FREEZE EVR_PROJET SLI", "GEN_DateDESIGNFREEZEEVR"
    Champs_ProjetsSLI.Add "Date DESIGN FREEZE LECTEUR BADGE_PROJET SLI", "GEN_DateDESIGNFREEZELECTEURBADGE"
    Champs_ProjetsSLI.Add "Date DESIGN FREEZE GSM REPEATER_PROJET SLI", "GEN_DateDESIGNFREEZEGSMREPEATER"
    Champs_ProjetsSLI.Add "Date DESIGN FREEZE SRS_PROJET SLI", "GEN_DateDESIGNFREEZESRS"
    ' --- Date FAI ---
    Champs_ProjetsSLI.Add "Date FAI PA_PACIS_PROJET SLI", "GEN_DateFAIPAPACIS"
    Champs_ProjetsSLI.Add "Date FAI PIS_PROJET SLI", "GEN_DateFAIPIS"
    Champs_ProjetsSLI.Add "Date FAI CCTV_PROJET SLI", "GEN_DateFAICCTV"
    Champs_ProjetsSLI.Add "Date FAI IOB_PROJET SLI", "GEN_DateFAIIOB"
    Champs_ProjetsSLI.Add "Date FAI APC_PROJET SLI", "GEN_DateFAIAPC"
    Champs_ProjetsSLI.Add "Date FAI DAS_PROJET SLI", "GEN_DateFAIDAS"
    Champs_ProjetsSLI.Add "Date FAI EVR_PROJET SLI", "GEN_DateFAIEVR"
    Champs_ProjetsSLI.Add "Date FAI LECTEUR BADGE_PROJET SLI", "GEN_DateFAILECTEURBADGE"
    Champs_ProjetsSLI.Add "Date FAI GSM REPEATER_PROJET SLI", "GEN_DateFAIGSMREPEATER"
    Champs_ProjetsSLI.Add "Date FAI SRS_PROJET SLI", "GEN_DateFAISRS"
    ' --- Date IQA ---
    Champs_ProjetsSLI.Add "Date IQA PA_PACIS_PROJET SLI", "GEN_DateIQAPAPACIS"
    Champs_ProjetsSLI.Add "Date IQA PIS_PROJET SLI", "GEN_DateIQAPIS"
    Champs_ProjetsSLI.Add "Date IQA CCTV_PROJET SLI", "GEN_DateIQACCTV"
    Champs_ProjetsSLI.Add "Date IQA IOB_PROJET SLI", "GEN_DateIQAIOB"
    Champs_ProjetsSLI.Add "Date IQA APC_PROJET SLI", "GEN_DateIQAAPC"
    Champs_ProjetsSLI.Add "Date IQA DAS_PROJET SLI", "GEN_DateIQADAS"
    Champs_ProjetsSLI.Add "Date IQA EVR_PROJET SLI", "GEN_DateIQAEVR"
    Champs_ProjetsSLI.Add "Date IQA LECTEUR BADGE_PROJET SLI", "GEN_DateIQALECTEURBADGE"
    Champs_ProjetsSLI.Add "Date IQA GSM REPEATER_PROJET SLI", "GEN_DateIQAGSMREPEATER"
    Champs_ProjetsSLI.Add "Date IQA SRS_PROJET SLI", "GEN_DateIQASRS"
    
    Set GetChampsDates = Champs_Dates

End Function
Public Function MappingProjetsSLI() As Variant
    
    ' Page General : 1 à 13
    ' Page Profil de mission : 14 à 23
    ' Page Interlocuteurs : 24 à 48
    ' --- Villeurbanne ---
    ' --- Leading Unit ---
    ' --- Domain Leader ---
    Dim arr As Variant
    
    ' Initialisation du mapping si non fait
    If Champs_ProjetsSLI Is Nothing Then Call GetChampsProjetsSLI
    arr = DictToArray2D(Champs_ProjetsSLI)
    If IsEmpty(arr) Then Exit Function
    
    
    
End Function


