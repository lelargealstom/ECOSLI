Test_SIRUS_DiagPreFillProduit("DTR0000120004")
======================================================================
Diagnostic pre-fill SnapShotPrix pour DTR = 'DTR0000120004'
======================================================================

-- 1. Etat cache m_dictSnapShotPrixCache --
  Cache : POPULATED. 3062 DTR distincts en cache.

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

-- 3. Lookup dictSnapShotPrix pour DTR='DTR0000120004' --
  Nb DTR total dans le dict : 3062
  DTR 'DTR0000120004' PRESENT. Valeurs qui seraient ecrites en cellules :
    Dates de validite des prix REPAIR connus sur Projet (ASL)_PRODUIT = 'D&IS Catalogue FY26_27_v03'
    Lead time REPAIR (jours) manuel_PRODUIT = '120'
    Prix SPARE catalogue a date_PRODUIT = '30,52'
    Dates de validite des prix SPARE connus sur Projet (ASL)_PRODUIT = 'D&IS Catalogue FY25_26_v05'
    MoQ spare catalogue_PRODUIT = ''
    Lead time SPARE catalogue (jours) a date_PRODUIT = ''

-- 4. Match config <-> dict pour ce DTR --
  [OK] 'Prix SPARE catalogue a date_PRODUIT' -> '30,52'
  [OK] 'Dates de validite des prix SPARE connus sur Projet (ASL)_PRODUIT' -> 'D&IS Catalogue FY25_26_v05'
  [OK] 'MoQ spare catalogue_PRODUIT' -> ''
  [OK] 'Lead time SPARE catalogue (jours) a date_PRODUIT' -> ''
  [--] 'Prix REPAIR catalogue a date_PRODUIT' -> ABSENT du dict (cellule restera vide).
  [OK] 'Dates de validite des prix REPAIR connus sur Projet (ASL)_PRODUIT' -> 'D&IS Catalogue FY26_27_v03'
  [OK] 'Lead time REPAIR (jours) manuel_PRODUIT' -> '120'
