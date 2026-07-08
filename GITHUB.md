Test_SIRUS_TenterADODBPurPourDTR("DTR0000120004")
======================================================================
Test ADODB pur pour DTR = 'DTR0000120004'
======================================================================

-- Variante 1 : Baseline (config actuelle CONNEXION_EXCEL_BDD_TRAVAIL) --
  Total lignes lues : 29789 | DTR alphanum : 0 | Matches DTR cible : 0

-- Variante 2 : IMEX=1;TypeGuessRows=0 --
  Total lignes lues : 29789 | DTR alphanum : 0 | Matches DTR cible : 0

-- Variante 3 : IMEX=1;MaxScanRows=0 --
  Total lignes lues : 29789 | DTR alphanum : 0 | Matches DTR cible : 0

-- Variante 4 : HDR=No;IMEX=1;TypeGuessRows=0 --
  ERREUR ouverture ou lecture : Aucune valeur donnée pour un ou plusieurs des paramètres requis.

-- Variante 5 : Excel 8.0 (Jet-style typing) --
  Total lignes lues : 29789 | DTR alphanum : 0 | Matches DTR cible : 0

-- Variante 6 : Excel 12.0 Xml (xlsb natif) --
  Total lignes lues : 29789 | DTR alphanum : 0 | Matches DTR cible : 0

Interpretation :
  Variante viable = nText > 0 (au moins un DTR alphanumerique lu)
  Variante ideale = nMatch >= 1 pour le DTR cible

