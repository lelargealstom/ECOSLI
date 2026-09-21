# Patch v2 sur ECOSLI_DATA_ASL.xlsx :
# - Supprime le bouton Shape Btn_Rafraichir_DATA_ASL (decision : bouton
#   deplace dans le UserForm ECOSLI a cote de Project_List)
# - Ajoute les formules locales en col I (NTI) et W (Duree tache SOFT
#   generique) qui devaient deja y etre selon la spec mais ont ete
#   oubliees dans le 1er script.

$path = "C:\Users\tomle\OneDrive\Bureau\ECOSLI\Templates\ECOSLI_DATA_ASL.xlsx"
$pwd = "LGM69000"

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.ScreenUpdating = $false

$wb = $xl.Workbooks.Open($path, $false, $false)
$ws = $wb.Worksheets.Item(1)
$wb.Unprotect($pwd)
$ws.Unprotect($pwd)

Write-Host "=== A. Supprimer bouton Btn_Rafraichir_DATA_ASL ==="
try {
    $btn = $ws.Shapes.Item("Btn_Rafraichir_DATA_ASL")
    $btn.Delete()
    Write-Host "  Bouton supprime"
} catch {
    Write-Host "  Bouton non trouve (deja supprime ?) : $($_.Exception.Message)"
}

Write-Host "=== B. Ajouter formules locales col I (NTI) et W (Duree SOFT gen) ==="
$lo = $ws.ListObjects.Item("Providor")

# Col I : NTI = calcul sur Node typed
# Note : utiliser ChrW(8230) = '…' (caractere identique a celui du header)
$nodeColName = "Node typed (LRU/SRU" + [char]8230 + ")"
$formulaNTI = "=IF([@[$nodeColName]]=""LRU"",""3"",IF([@[$nodeColName]]=""LRUNR"",""3"",IF([@[$nodeColName]]=""SRU"",""3+"",IF([@[$nodeColName]]=""SRUNR"",""3+"",""""))))"
$colI = $lo.ListColumns.Item(9).DataBodyRange
$colI.Formula = $formulaNTI
Write-Host "  Col I (NTI) - formule appliquee sur $($colI.Rows.Count) lignes"

# Col W : Duree tache preventive SOFT generique = 0.5 si HARD/SOFT = BOTH
$formulaSOFTgen = "=IF([@[HARD/SOFT]]=""BOTH"",0.5,"""")"
$colW = $lo.ListColumns.Item(23).DataBodyRange
$colW.Formula = $formulaSOFTgen
Write-Host "  Col W (Duree SOFT gen) - formule appliquee sur $($colW.Rows.Count) lignes"

Write-Host "=== C. Pas de re-protection (decision Thomas 2026-06-29) ==="

Write-Host "=== Sauvegarde ==="
$xl.ScreenUpdating = $true
$wb.Save()
$wb.Close($false)
$xl.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null
[GC]::Collect()
Write-Host "Termine"
