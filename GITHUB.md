Test_SIRUS_TenterADODBPurPourDTR("DTR0000120004")
======================================================================
Test ADODB pur pour DTR = 'DTR0000120004'
======================================================================

-- Variante 1 : Baseline (config actuelle CONNEXION_EXCEL_BDD_TRAVAIL) --
  MATCH 1 : [Repair] Date=22/06/2026 | Prix= | RefCat='D&IS Catalogue FY26_27_v03'
  MATCH 2 : [Repair] Date=06/03/2026 | Prix= | RefCat='D&IS Catalogue FY26_27_v01'
  MATCH 3 : [Repair] Date=06/03/2026 | Prix= | RefCat='D&IS Catalogue FY26_27_v02'
  MATCH 4 : [Spare] Date=26/10/2025 | Prix=30,52 | RefCat='D&IS Catalogue FY25_26_v05'
  MATCH 5 : [Spare] Date=01/09/2024 | Prix=34,1 | RefCat='FCS 2024_08'
  Total lignes lues : 29788 | DTR alphanum : 29748 | Matches DTR cible : 5
  Exemples DTR alphanum : 'DTR0000007454' ; 'DTR0000007454' ; 'DTR0000007454' ; 'DTR0000007454' ; 'DTR0000007454' ; 

-- Variante 2 : IMEX=1;TypeGuessRows=0 --
  MATCH 1 : [Repair] Date=22/06/2026 | Prix= | RefCat='D&IS Catalogue FY26_27_v03'
  MATCH 2 : [Repair] Date=06/03/2026 | Prix= | RefCat='D&IS Catalogue FY26_27_v01'
  MATCH 3 : [Repair] Date=06/03/2026 | Prix= | RefCat='D&IS Catalogue FY26_27_v02'
  MATCH 4 : [Spare] Date=26/10/2025 | Prix=30,52 | RefCat='D&IS Catalogue FY25_26_v05'
  MATCH 5 : [Spare] Date=01/09/2024 | Prix=34,1 | RefCat='FCS 2024_08'
  Total lignes lues : 29788 | DTR alphanum : 29748 | Matches DTR cible : 5
  Exemples DTR alphanum : 'DTR0000007454' ; 'DTR0000007454' ; 'DTR0000007454' ; 'DTR0000007454' ; 'DTR0000007454' ; 

-- Variante 3 : IMEX=1;MaxScanRows=0 --
  MATCH 1 : [Repair] Date=22/06/2026 | Prix= | RefCat='D&IS Catalogue FY26_27_v03'
  MATCH 2 : [Repair] Date=06/03/2026 | Prix= | RefCat='D&IS Catalogue FY26_27_v01'
  MATCH 3 : [Repair] Date=06/03/2026 | Prix= | RefCat='D&IS Catalogue FY26_27_v02'
  MATCH 4 : [Spare] Date=26/10/2025 | Prix=30,52 | RefCat='D&IS Catalogue FY25_26_v05'
  MATCH 5 : [Spare] Date=01/09/2024 | Prix=34,1 | RefCat='FCS 2024_08'
  Total lignes lues : 29788 | DTR alphanum : 29748 | Matches DTR cible : 5
  Exemples DTR alphanum : 'DTR0000007454' ; 'DTR0000007454' ; 'DTR0000007454' ; 'DTR0000007454' ; 'DTR0000007454' ; 

-- Variante 4 : HDR=No;IMEX=1;TypeGuessRows=0 --
  ERREUR ouverture ou lecture : Aucune valeur donnée pour un ou plusieurs des paramètres requis.

-- Variante 5 : Excel 8.0 (Jet-style typing) --
  MATCH 1 : [Repair] Date=22/06/2026 | Prix= | RefCat='D&IS Catalogue FY26_27_v03'
  MATCH 2 : [Repair] Date=06/03/2026 | Prix= | RefCat='D&IS Catalogue FY26_27_v01'
  MATCH 3 : [Repair] Date=06/03/2026 | Prix= | RefCat='D&IS Catalogue FY26_27_v02'
  MATCH 4 : [Spare] Date=26/10/2025 | Prix=30,52 | RefCat='D&IS Catalogue FY25_26_v05'
  MATCH 5 : [Spare] Date=01/09/2024 | Prix=34,1 | RefCat='FCS 2024_08'
  Total lignes lues : 29788 | DTR alphanum : 29748 | Matches DTR cible : 5
  Exemples DTR alphanum : 'DTR0000007454' ; 'DTR0000007454' ; 'DTR0000007454' ; 'DTR0000007454' ; 'DTR0000007454' ; 

-- Variante 6 : Excel 12.0 Xml (xlsb natif) --
  MATCH 1 : [Repair] Date=22/06/2026 | Prix= | RefCat='D&IS Catalogue FY26_27_v03'
  MATCH 2 : [Repair] Date=06/03/2026 | Prix= | RefCat='D&IS Catalogue FY26_27_v01'
  MATCH 3 : [Repair] Date=06/03/2026 | Prix= | RefCat='D&IS Catalogue FY26_27_v02'
  MATCH 4 : [Spare] Date=26/10/2025 | Prix=30,52 | RefCat='D&IS Catalogue FY25_26_v05'
  MATCH 5 : [Spare] Date=01/09/2024 | Prix=34,1 | RefCat='FCS 2024_08'
  Total lignes lues : 29788 | DTR alphanum : 29748 | Matches DTR cible : 5
  Exemples DTR alphanum : 'DTR0000007454' ; 'DTR0000007454' ; 'DTR0000007454' ; 'DTR0000007454' ; 'DTR0000007454' ; 

Interpretation :
  Variante viable = nText > 0 (au moins un DTR alphanumerique lu)
  Variante ideale = nMatch >= 1 pour le DTR cible
