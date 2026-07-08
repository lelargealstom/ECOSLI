Test_SIRUS_DiagPreFillProduit(DTR0000120004)
======================================================================
Diagnostic pre-fill SnapShotPrix pour DTR = ''
======================================================================

-- 1. Etat cache m_dictSnapShotPrixCache --
  Cache : VIDE (Nothing). Sera charge au 1er appel.

-- 2. Config Produit : champs Source='SnapShotPrix' --
  Nb champs total dans config Produit : 25
  [1] Champ='Prix Spare a date' | ColonneBdD='Prix SPARE catalogue a date_PRODUIT' | Obligatoire=Conditionnel
  [2] Champ='Dates validite Spare' | ColonneBdD='Dates de validite des prix SPARE connus sur Projet (ASL)_PRODUIT' | Obligatoire=Conditionnel
  [3] Champ='MoQ Spare' | ColonneBdD='MoQ spare catalogue_PRODUIT' | Obligatoire=Conditionnel
  [4] Champ='Lead Time Spare' | ColonneBdD='Lead time SPARE catalogue (jours) a date_PRODUIT' | Obligatoire=Conditionnel
  [5] Champ='Prix Repair a date' | ColonneBdD='Prix REPAIR catalogue a date_PRODUIT' | Obligatoire=Conditionnel
  [6] Champ='Dates validite Repair' | ColonneBdD='Dates de validite des prix REPAIR connus sur Projet (ASL)_PRODUIT' | Obligatoire=Conditionnel
  [7] Champ='Lead Time Repair' | ColonneBdD='Lead time REPAIR (jours) manuel_PRODUIT' | Obligatoire=Conditionnel
  -> 7 champs Source='SnapShotPrix' trouves (attendu : 7).

-- 3. Lookup dictSnapShotPrix pour DTR='' --
  Nb DTR total dans le dict : 32
  DTR '' ABSENT du dict SnapShotPrix (pas de snapshot retenu).
  Verifier : le DTR a-t-il des lignes Prix non nul ou OnRequestPrix dans Produit_SnapShotPrix ?
  10 premiers DTR presents en dict :
    '8600361'
    '8600385'
    '8600562'
    '8600602'
    '8600611'
    '8600635'
    '8600711'
    '8600715'
    '8600716'
    '8600723'
