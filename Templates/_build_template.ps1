# Script de generation du template DATA_ASL fige (v2)
# Lit ECOSLI_newtemplate.xlsx, ecrit ECOSLI_DATA_ASL.xlsx
# Note : conserve le ListObject "Providor" existant (utilise par
# D_Providor_Main / B_ATM_Main / C_EMMST_Main downstream).

$src = "C:\Users\tomle\OneDrive\Bureau\ECOSLI\Templates\ECOSLI_newtemplate.xlsx"
$dst = "C:\Users\tomle\OneDrive\Bureau\ECOSLI\Templates\ECOSLI_DATA_ASL.xlsx"
$logo = "C:\Users\tomle\OneDrive\Bureau\ECOSLI\Templates\logo_alstom.png"
$pwd = "LGM69000"

Copy-Item -Path $src -Destination $dst -Force

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.ScreenUpdating = $false

$wb = $xl.Workbooks.Open($dst, $false, $false)
$xl.Calculation = -4135   # xlCalculationManual (apres open OK)
$ws = $wb.Worksheets.Item(1)

# Couleurs (Interior.Color en BGR)
$COLOR_VERT       = 5296274    # saisie ASL
$COLOR_ROUGE      = 255        # Refresh BdD
$COLOR_BLEU       = 15773696   # calcul local
$COLOR_GRIS       = 12632256   # constante (PH/Y)
$COLOR_TITRE_BG   = 12419407   # bleu clair pour titre

Write-Host "=== Etape 1 : Fix typos colonnes AA et AE ==="
$ws.Cells.Item(4, 27).Value2 = "Prix Spare Catalogue"
$ws.Cells.Item(4, 31).Value2 = "Prix Repair Catalogue"
Write-Host "  OK"

Write-Host "=== Etape 2 : Fix couleur AL et AM (rouge -> bleu) ==="
$ws.Cells.Item(4, 38).Interior.Color = $COLOR_BLEU
$ws.Cells.Item(4, 39).Interior.Color = $COLOR_BLEU
Write-Host "  OK"

Write-Host "=== Etape 3 : Supprimer colonnes AP a AW (vides) ==="
for ($c = 49; $c -ge 42; $c--) { $ws.Columns.Item($c).Delete() | Out-Null }
Write-Host "  OK (8 colonnes supprimees)"

Write-Host "=== Etape 4 : Labels coefficients ligne 1 ==="
$ws.Cells.Item(1, 28).Value2 = "Coeff Prix calcule"
$ws.Cells.Item(1, 37).Value2 = "Coeff conv"
$ws.Cells.Item(1, 38).Value2 = "Coeff Alstom Spare"
$ws.Cells.Item(1, 39).Value2 = "Coeff Alstom Repair"
foreach ($c in @(28, 37, 38, 39)) {
    $cell = $ws.Cells.Item(1, $c)
    $cell.Font.Bold = $true
    $cell.Font.Size = 8
    $cell.Font.Italic = $true
    $cell.HorizontalAlignment = -4108
}
Write-Host "  OK"

Write-Host "=== Etape 5 : Logo + titre ==="
$ws.Cells.Item(1, 1).Value2 = ""
$ws.Rows.Item(1).RowHeight = 22
$ws.Rows.Item(2).RowHeight = 22
$ws.Columns.Item(1).ColumnWidth = 12
$ws.Columns.Item(2).ColumnWidth = 12

$logoRange = $ws.Range("A1:B2")
$pic = $ws.Shapes.AddPicture($logo, $false, $true, $logoRange.Left, $logoRange.Top, -1, -1)
$pic.LockAspectRatio = -1
if ($pic.Height -gt $logoRange.Height * 0.9) { $pic.Height = $logoRange.Height * 0.9 }
if ($pic.Width -gt $logoRange.Width * 0.9) { $pic.Width = $logoRange.Width * 0.9 }
$pic.Top = $logoRange.Top + ($logoRange.Height - $pic.Height) / 2
$pic.Left = $logoRange.Left + ($logoRange.Width - $pic.Width) / 2
Write-Host "  Logo insere"

$titleRange = $ws.Range("C1:G2")
$titleRange.Merge()
$titleRange.Value2 = "ECOSLI - DATA_ASL PBS Projet"
$titleRange.Font.Size = 16
$titleRange.Font.Bold = $true
$titleRange.Font.Color = 9764864
$titleRange.HorizontalAlignment = -4108
$titleRange.VerticalAlignment = -4108
$titleRange.Interior.Color = $COLOR_TITRE_BG
Write-Host "  Titre OK"

Write-Host "=== Etape 6 : Cellule Projet + Date maj ==="
$ws.Cells.Item(1, 8).Value2 = "Projet :"
$ws.Cells.Item(1, 8).Font.Bold = $true
$ws.Cells.Item(1, 8).HorizontalAlignment = -4152
$ws.Range("I1:J1").Merge()
$ws.Cells.Item(1, 9).Value2 = "(rempli au Rafraichir)"
$ws.Cells.Item(1, 9).Font.Italic = $true
$ws.Cells.Item(1, 9).Font.Color = 8421504

$ws.Cells.Item(2, 8).Value2 = "Derniere maj :"
$ws.Cells.Item(2, 8).Font.Bold = $true
$ws.Cells.Item(2, 8).HorizontalAlignment = -4152
$ws.Range("I2:J2").Merge()
$ws.Cells.Item(2, 9).Value2 = "(jamais)"
$ws.Cells.Item(2, 9).Font.Italic = $true
$ws.Cells.Item(2, 9).Font.Color = 8421504

# Cellules nommees pour ecriture VBA facile
$shName = $ws.Name
$wb.Names.Add("DATA_ASL_PROJET", "='" + $shName + "'!`$I`$1") | Out-Null
$wb.Names.Add("DATA_ASL_DERNIERE_MAJ", "='" + $shName + "'!`$I`$2") | Out-Null
Write-Host "  Cellules nommees creees"

Write-Host "=== Etape 7 : Legende code couleur K1:N2 ==="
$legende = @(
    @{Col=11; Bg=$COLOR_VERT;  Label="Saisie ASL"},
    @{Col=12; Bg=$COLOR_ROUGE; Label="Refresh BdD"},
    @{Col=13; Bg=$COLOR_BLEU;  Label="Calcul local"},
    @{Col=14; Bg=$COLOR_GRIS;  Label="Constante"}
)
foreach ($l in $legende) {
    $cell1 = $ws.Cells.Item(1, $l.Col)
    $cell2 = $ws.Cells.Item(2, $l.Col)
    $cell1.Interior.Color = $l.Bg
    $cell2.Value2 = $l.Label
    $cell2.Font.Size = 8
    $cell2.HorizontalAlignment = -4108
    $borders = $cell1.Borders
    for ($i = 7; $i -le 10; $i++) { $borders.Item($i).LineStyle = 1; $borders.Item($i).Weight = 2 }
    $ws.Columns.Item($l.Col).ColumnWidth = 12
}
Write-Host "  Legende OK"

Write-Host "=== Etape 8 : Bouton Rafraichir Shape ==="
$btnRange = $ws.Range("O1:T2")
$btn = $ws.Shapes.AddShape(5, $btnRange.Left, $btnRange.Top, $btnRange.Width, $btnRange.Height)
$btn.Name = "Btn_Rafraichir_DATA_ASL"
$btn.Fill.ForeColor.RGB = 9764864
$btn.Line.ForeColor.RGB = 9764864
$btn.TextFrame2.TextRange.Text = "Rafraichir DATA_ASL"
$btn.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 16777215
$btn.TextFrame2.TextRange.Font.Bold = $true
$btn.TextFrame2.TextRange.Font.Size = 11
$btn.TextFrame2.HorizontalAnchor = 2
$btn.TextFrame2.VerticalAnchor = 3
$btn.OnAction = "'ECOSLI.xlam'!SIRUS_RafraichirDATA_ASL"
Write-Host "  Bouton cree, macro = 'ECOSLI.xlam'!SIRUS_RafraichirDATA_ASL"

Write-Host "=== Etape 9 : Reutiliser ListObject Providor existant (downstream compat) ==="
$lo = $ws.ListObjects.Item("Providor")
Write-Host "  ListObject existant : $($lo.Name) - $($lo.Range.Address())"
# Resize si besoin pour couvrir A4:AO + derniere ligne actuelle
$lastRow = [Math]::Max(52, $lo.Range.Row + $lo.Range.Rows.Count - 1)
$newRange = $ws.Range("A4:AO" + $lastRow)
$lo.Resize($newRange)
Write-Host "  Resize a $($lo.Range.Address())"
Write-Host "  Note : on conserve le nom 'Providor' (utilise par D_Providor_Main + B_ATM_Main + C_EMMST_Main)"

Write-Host "=== Etape 10 : Freeze panes (ligne 5 + col E) ==="
$ws.Activate()
$ws.Cells.Item(5, 5).Select() | Out-Null
$xl.ActiveWindow.FreezePanes = $true
Write-Host "  OK"

Write-Host "=== Etape 11 : Format cellules par type ==="
function Set-ColFormat($colIdx, $format) {
    $col = $lo.ListColumns.Item($colIdx).DataBodyRange
    if ($col -ne $null) { $col.NumberFormat = $format }
}
Set-ColFormat 1 "@"
Set-ColFormat 2 "@"
Set-ColFormat 3 "@"
Set-ColFormat 4 "0"
Set-ColFormat 5 "@"
Set-ColFormat 6 "@"
Set-ColFormat 7 "@"
Set-ColFormat 8 "@"
Set-ColFormat 9 "@"
Set-ColFormat 10 "@"
Set-ColFormat 11 "0"
Set-ColFormat 12 "0"
Set-ColFormat 13 "0"
Set-ColFormat 14 "0.00E+00"
Set-ColFormat 15 "0.00E+00"
Set-ColFormat 16 "0.000"
Set-ColFormat 17 "0.000"
Set-ColFormat 18 "@"
Set-ColFormat 19 "0"
Set-ColFormat 20 "0"
Set-ColFormat 21 "0"
Set-ColFormat 22 "0"
Set-ColFormat 23 "0.000"
Set-ColFormat 24 "0.000"
Set-ColFormat 25 "@"
Set-ColFormat 26 "#,##0.00"
Set-ColFormat 27 "#,##0.00"
Set-ColFormat 28 "#,##0.00"
Set-ColFormat 29 "@"
Set-ColFormat 30 "0"
Set-ColFormat 31 "#,##0.00"
Set-ColFormat 32 "@"
Set-ColFormat 33 "0"
Set-ColFormat 34 "#,##0.00"
Set-ColFormat 35 "0"
Set-ColFormat 36 "#,##0.00"
Set-ColFormat 37 "0"
Set-ColFormat 38 "#,##0.00"
Set-ColFormat 39 "#,##0.00"
Set-ColFormat 40 "#,##0.00"
Set-ColFormat 41 "#,##0.00"
Write-Host "  41 formats appliques"

Write-Host "=== Etape 12 : Commentaires sur 15 entetes cles ==="
function Add-Comment($colIdx, $text) {
    $cell = $lo.HeaderRowRange.Item(1, $colIdx)
    if ($cell.Comment -ne $null) { $cell.Comment.Delete() }
    $cell.AddComment($text) | Out-Null
    $cell.Comment.Visible = $false
}
Add-Comment 1  "Identifiant unique du DTR - cle de matching BdD. Ne pas modifier apres saisie."
Add-Comment 5  "Auto-rempli depuis Produit.[Designation_PRODUIT] via Refresh."
Add-Comment 7  "Auto-rempli depuis Produit.[Node type Produit] via Refresh."
Add-Comment 8  "Auto-calcule depuis Taches_ATM : BOTH si HARD+SOFT, HARD si HARD seul, SOFT si SOFT seul, vide sinon."
Add-Comment 9  "Calcul local : 3 si LRU/LRUNR, 3+ si SRU/SRUNR, vide sinon."
Add-Comment 11 "Saisie ASL - issu du RAM Report specifique projet (CGR)."
Add-Comment 12 "Saisie ASL - issu du RAM Report specifique projet (CGR)."
Add-Comment 17 "Auto-calcule : SUM des Temps analytiques des taches HARD pour ce DTR."
Add-Comment 24 "Auto-calcule : SUM des Temps analytiques des taches SOFT pour ce DTR."
Add-Comment 27 "Prix Spare au D&IS Catalogue (date la plus recente)."
Add-Comment 29 "Reference du catalogue ayant fourni le Prix Spare retenu."
Add-Comment 31 "Prix Repair au D&IS Catalogue (date la plus recente)."
Add-Comment 32 "Reference du catalogue ayant fourni le Prix Repair retenu."
Add-Comment 40 "Moyenne du prix Spare sur les projets archives ou ce DTR existe (hors projet courant)."
Add-Comment 41 "Moyenne du prix Repair sur les projets archives ou ce DTR existe (hors projet courant)."
Write-Host "  15 commentaires ajoutes"

Write-Host "=== Etape 13 : Mise en forme conditionnelle (cellules Refresh vides) ==="
$refreshCols = @(5, 7, 8, 13, 16, 17, 19, 21, 24, 26, 27, 29, 30, 31, 32, 33, 40, 41)
foreach ($colIdx in $refreshCols) {
    $col = $lo.ListColumns.Item($colIdx).DataBodyRange
    if ($col -ne $null) {
        $fc = $col.FormatConditions.Add(1, 3, "=`"`"")
        $fc.Interior.Color = 11857151   # orange clair
    }
}
Write-Host "  MFC OK sur $($refreshCols.Count) colonnes Refresh"

Write-Host "=== Etape 14 : Deverrouiller cellules saisie ASL (vertes) ==="
$saisieCols = @(1, 2, 3, 4, 11, 12, 20, 22, 34, 35, 36, 37)
$ws.Cells.Locked = $true
foreach ($colIdx in $saisieCols) {
    $col = $lo.ListColumns.Item($colIdx).DataBodyRange
    if ($col -ne $null) { $col.Locked = $false }
}
Write-Host "  $($saisieCols.Count) colonnes vertes deverrouillees"

Write-Host "=== Etape 15 : Nettoyer cellules nommees cassees (#REF!) ==="
$nbDeleted = 0
$namesToDelete = @()
foreach ($n in $wb.Names) {
    if ($n.RefersTo -like "*#REF!*") {
        $namesToDelete += $n.Name
    }
}
foreach ($nm in $namesToDelete) {
    $wb.Names.Item($nm).Delete()
    $nbDeleted++
}
Write-Host "  $nbDeleted nom(s) supprime(s) : $($namesToDelete -join ', ')"

Write-Host "=== Etapes 16-17 : Protection desactivee (decision Thomas 2026-06-29) ==="
Write-Host "  La feuille DATA_ASL et le workbook restent non proteges."

Write-Host "=== Sauvegarde ==="
$xl.Calculation = -4105   # xlCalculationAutomatic
$xl.ScreenUpdating = $true
$wb.Save()
$wb.Close($false)
$xl.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null
[GC]::Collect()
[GC]::WaitForPendingFinalizers()

Write-Host ""
Write-Host "=== FINI ==="
Write-Host "Template genere : $dst"
