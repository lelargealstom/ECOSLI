Public Sub Cadrage_ECOSLI(Infos_tab As Variant)
' Macro de cadrage de la feuille Info ATM de la BDD
' Il est possible de changer la disposition des colonnes mais pas la nomination des colonnes

'On récupère les numéros de colonne
ECOSLI_Type.dtr = 6
ECOSLI_Type.designation = 7
ECOSLI_Type.Designation_BdD = 8
ECOSLI_Type.Designation_Providor = 9
ECOSLI_Type.Sous_syst = 10
ECOSLI_Type.Qte = 12
ECOSLI_Type.NodeType = 13
ECOSLI_Type.HARD_SOFT = 14
ECOSLI_Type.NTI = 15
ECOSLI_Type.RefPlan = 16
ECOSLI_Type.MTBF_Gen = 17
ECOSLI_Type.TxDefaillanceGen = 18
ECOSLI_Type.DureeTacheGenHARD = 19
ECOSLI_Type.MTBF_Intr = 20
ECOSLI_Type.MTBF_Fonct = 21
ECOSLI_Type.TxDefaillanceAna = 22
ECOSLI_Type.DureeTacheAnaHARD = 23
ECOSLI_Type.UniteCorrective = 24
ECOSLI_Type.DureeVieGenHARD = 25
ECOSLI_Type.DureeVieProjetHARD = 26
ECOSLI_Type.DureeVieGenSOFT = 27
ECOSLI_Type.DureeTacheGenSOFT = 28
ECOSLI_Type.DureeVieProjetSOFT = 29
ECOSLI_Type.DureeTacheAnaSOFT = 30
ECOSLI_Type.UnitePreventive = 31
ECOSLI_Type.CoutTotal = 32
ECOSLI_Type.PrixSpare_Catalogue = 33
ECOSLI_Type.LeadTimeSpare_Catalogue = 37
ECOSLI_Type.PrixRepa_Catalogue = 38
ECOSLI_Type.LeadTimeRepa_Catalogue = 39
ECOSLI_Type.PrixSpare_FRS = 40
ECOSLI_Type.LeadTimeSpare_FRS = 41
ECOSLI_Type.PrixRepa_FRS = 42
ECOSLI_Type.LeadTimeRepa_FRS = 43
ECOSLI_Type.PrixSpare_Manuel = 47
ECOSLI_Type.PrixRepa_Manuel = 48
ECOSLI_Type.PrixProjet_Archive = 49

End Sub
