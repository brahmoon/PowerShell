Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# --- ヘルパー関数 ---
function SafeStr($v) { if ($null -eq $v) { return "" } try { return [string]$v } catch { return "" } }

function Convert-ToColumnName([int]$number) {
    $columnName = ""
    while ($number -gt 0) {
        $mod = ($number - 1) % 26
        $columnName = [char](65 + $mod) + $columnName
        $number = [math]::Floor(($number - $mod) / 26)
    }
    return $columnName
}

function Get-ExcelApp {
    $app = New-Object -ComObject Excel.Application
    $app.Visible = $false
    $app.DisplayAlerts = $false
    $app.ScreenUpdating = $false
    return $app
}

function Select-ExcelFile($title) {
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Title = $title
    $ofd.Filter = "Excel Files|*.xlsx;*.xlsm;*.xlsb;*.xls"
    if ($ofd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $ofd.FileName
    }
    return $null
}

function Show-SheetSelector($sheetNames) {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "比較シート設定"
    $form.Size = New-Object System.Drawing.Size(350,420)
    $form.StartPosition = "CenterScreen"

    $checkedList = New-Object System.Windows.Forms.CheckedListBox
    $checkedList.Dock = "Top"
    $checkedList.Height = 320
    $sheetNames | ForEach-Object { [void]$checkedList.Items.Add($_, $true) }

    $btnAll = New-Object System.Windows.Forms.Button
    $btnAll.Text = "すべて選択"
    $btnAll.Width = 140
    $btnAll.Location = New-Object System.Drawing.Point(20,330)
    $btnAll.Add_Click({ for ($i=0; $i -lt $checkedList.Items.Count; $i++) { $checkedList.SetItemChecked($i,$true) } })

    $btnNone = New-Object System.Windows.Forms.Button
    $btnNone.Text = "すべて解除"
    $btnNone.Width = 140
    $btnNone.Location = New-Object System.Drawing.Point(180,330)
    $btnNone.Add_Click({ for ($i=0; $i -lt $checkedList.Items.Count; $i++) { $checkedList.SetItemChecked($i,$false) } })

    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Text = "OK"
    $btnOK.Dock = "Bottom"
    $btnOK.Add_Click({ $form.Tag = $checkedList.CheckedItems; $form.Close() })

    $form.Controls.AddRange(@($checkedList,$btnAll,$btnNone,$btnOK))
    [void]$form.ShowDialog()
    return @($form.Tag)
}

function Compare-Sheets($s1, $s2, $sheetName) {
    $r1 = $s1.UsedRange
    $r2 = $s2.UsedRange
    $maxRow = [Math]::Max([int]$r1.Rows.Count, [int]$r2.Rows.Count)
    $maxCol = [Math]::Max([int]$r1.Columns.Count, [int]$r2.Columns.Count)
    if ($maxRow -lt 1) { $maxRow = 1 }
    if ($maxCol -lt 1) { $maxCol = 1 }

    $v1 = $s1.Range("A1",$s1.Cells.Item($maxRow,$maxCol)).Value2
    $v2 = $s2.Range("A1",$s2.Cells.Item($maxRow,$maxCol)).Value2

    $diffs = @()
    for ($r=1; $r -le $maxRow; $r++) {
        for ($c=1; $c -le $maxCol; $c++) {
            $val1 = if ($v1 -is [System.Array]) { $v1[$r,$c] } else { $v1 }
            $val2 = if ($v2 -is [System.Array]) { $v2[$r,$c] } else { $v2 }
            if ($val1 -ne $val2) {
                $diffs += [PSCustomObject]@{
                    Sheet  = $sheetName
                    Row    = $r
                    Column = $c
                    Text1  = $val1
                    Text2  = $val2
                }
            }
        }
    }
    return $diffs
}

function Show-Progress($title, $max) {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = $title
    $form.Size = New-Object System.Drawing.Size(400,120)
    $form.StartPosition = "CenterScreen"
    $form.ControlBox = $false
    $form.TopMost = $true

    $label = New-Object System.Windows.Forms.Label
    $label.Text = "比較を実行しています..."
    $label.Dock = "Top"
    $label.TextAlign = "MiddleCenter"
    $label.Height = 30

    $bar = New-Object System.Windows.Forms.ProgressBar
    $bar.Dock = "Bottom"
    $bar.Minimum = 0
    $bar.Maximum = $max
    $bar.Value = 0

    $form.Controls.AddRange(@($label, $bar))
    $form.Show()
    [System.Windows.Forms.Application]::DoEvents()
    return @{ Form = $form; Bar = $bar; Label = $label }
}

# === GUI構築 ===
$form = New-Object System.Windows.Forms.Form
$form.Text = "Excel差分比較ツール"
$form.Size = New-Object System.Drawing.Size(900,620)
$form.StartPosition = "CenterScreen"
$form.AllowDrop = $true

# --- Book1 / Book2 ラベル縮小 ---
$lbl1 = New-Object System.Windows.Forms.Label
$lbl1.Text = "B1:"
$lbl1.Width = 22
$lbl1.Location = New-Object System.Drawing.Point(20,25)

$txtFile1 = New-Object System.Windows.Forms.TextBox
$txtFile1.Location = New-Object System.Drawing.Point(50,20)
$txtFile1.Width = 670

$btnFile1 = New-Object System.Windows.Forms.Button
$btnFile1.Text = "選択"
$btnFile1.Location = New-Object System.Drawing.Point(740,18)
$btnFile1.Add_Click({ $path = Select-ExcelFile "Book1を選択"; if ($path) { $txtFile1.Text = $path } })

$lbl2 = New-Object System.Windows.Forms.Label
$lbl2.Text = "B2:"
$lbl2.Width = 22
$lbl2.Location = New-Object System.Drawing.Point(20,65)

$txtFile2 = New-Object System.Windows.Forms.TextBox
$txtFile2.Location = New-Object System.Drawing.Point(50,60)
$txtFile2.Width = 670

$btnFile2 = New-Object System.Windows.Forms.Button
$btnFile2.Text = "選択"
$btnFile2.Location = New-Object System.Drawing.Point(740,58)
$btnFile2.Add_Click({ $path = Select-ExcelFile "Book2を選択"; if ($path) { $txtFile2.Text = $path } })

# --- ドラッグ＆ドロップ対応 ---
$form.Add_DragEnter({
    if ($_.Data.GetDataPresent([Windows.Forms.DataFormats]::FileDrop)) { $_.Effect = 'Copy' }
})
$form.Add_DragDrop({
    $files = $_.Data.GetData([Windows.Forms.DataFormats]::FileDrop)
    if ($files.Count -gt 0) {
        if (-not $txtFile1.Text) { $txtFile1.Text = $files[0] }
        elseif (-not $txtFile2.Text) { $txtFile2.Text = $files[0] }
        else { $txtFile1.Text = $files[0] }
    }
})

$btnSheet = New-Object System.Windows.Forms.Button
$btnSheet.Text = "比較シート設定"
$btnSheet.Location = New-Object System.Drawing.Point(20,100)
$btnSheet.Width = 240

$btnCompare = New-Object System.Windows.Forms.Button
$btnCompare.Text = "比較開始"
$btnCompare.Location = New-Object System.Drawing.Point(280,100)

$listView = New-Object System.Windows.Forms.ListView
$listView.View = 'Details'
$listView.FullRowSelect = $true
$listView.Location = New-Object System.Drawing.Point(20,150)
$listView.Size = New-Object System.Drawing.Size(840,400)
[void]$listView.Columns.Add("Sheet",100)
[void]$listView.Columns.Add("Row",60)
[void]$listView.Columns.Add("Column",80)
[void]$listView.Columns.Add("Text1",300)
[void]$listView.Columns.Add("Text2",300)

$form.Controls.AddRange(@($lbl1,$txtFile1,$btnFile1,$lbl2,$txtFile2,$btnFile2,$btnSheet,$btnCompare,$listView))

# --- シート選択 ---
$btnSheet.Add_Click({
    if (-not (Test-Path $txtFile1.Text) -or -not (Test-Path $txtFile2.Text)) {
        [System.Windows.Forms.MessageBox]::Show("両方のファイルを指定してください。")
        return
    }
    $app = Get-ExcelApp
    $wb1 = $app.Workbooks.Open($txtFile1.Text)
    $wb2 = $app.Workbooks.Open($txtFile2.Text)
    $names1 = @($wb1.Sheets | ForEach-Object { $_.Name })
    $names2 = @($wb2.Sheets | ForEach-Object { $_.Name })
    $global:selectedSheets = Show-SheetSelector ($names1 | Where-Object { $names2 -contains $_ })
    $wb1.Close($false)
    $wb2.Close($false)
    $app.Quit()
})

# --- 比較開始 ---
$btnCompare.Add_Click({
    if (-not (Test-Path $txtFile1.Text) -or -not (Test-Path $txtFile2.Text)) {
        [System.Windows.Forms.MessageBox]::Show("ファイルパスが正しくありません。")
        return
    }

    $app = Get-ExcelApp
    $global:wb1 = $app.Workbooks.Open($txtFile1.Text)
    $global:wb2 = $app.Workbooks.Open($txtFile2.Text)

    $names1 = @($wb1.Sheets | ForEach-Object { $_.Name })
    $names2 = @($wb2.Sheets | ForEach-Object { $_.Name })
    $common = $names1 | Where-Object { $names2 -contains $_ }
    $targetSheets = if ($global:selectedSheets) { $global:selectedSheets } else { $common }

    $progress = Show-Progress "比較中..." $targetSheets.Count
    $listView.Items.Clear()

    $i = 0
    foreach ($name in $targetSheets) {
        $i++
        $progress.Label.Text = "比較中: $name ($i / $($targetSheets.Count))"
        $diffs = Compare-Sheets $wb1.Sheets.Item($name) $wb2.Sheets.Item($name) $name
        foreach ($d in $diffs) {
            $item = New-Object System.Windows.Forms.ListViewItem((SafeStr $d.Sheet))
            [void]$item.SubItems.Add((SafeStr $d.Row))
            [void]$item.SubItems.Add((Convert-ToColumnName $d.Column))
            [void]$item.SubItems.Add((SafeStr $d.Text1))
            [void]$item.SubItems.Add((SafeStr $d.Text2))
            [void]$listView.Items.Add($item)
        }
        $progress.Bar.Value = $i
        [System.Windows.Forms.Application]::DoEvents()
    }

    $progress.Form.Close()
    $app.Visible = $false
    [System.Windows.Forms.MessageBox]::Show("比較完了: $($listView.Items.Count) 件の差分が見つかりました。")
})

# --- 結果クリックでセル表示 ---
$listView.Add_Click({
    if ($listView.SelectedItems.Count -eq 0) { return }
    $sel = $listView.SelectedItems[0]
    $sheet = $sel.SubItems[0].Text
    $row = [int]$sel.SubItems[1].Text
    $colName = $sel.SubItems[2].Text

    # 列名→列番号
    $col = 0
    foreach ($ch in $colName.ToCharArray()) { $col = $col * 26 + ([byte][char]$ch - 64) }

    try {
        if ($global:wb1 -and $global:wb2) {

            # --- ① Workbook1 を前面化してセルを選択 ---
            $ws1 = $global:wb1.Sheets.Item($sheet)
            $ws1.Activate()
            $ws1.Cells.Item($row,$col).Select()
            $global:wb1.Activate()
            $global:wb1.Application.Visible = $false
            $global:wb1.Windows.Item(1).Visible = $false
            $global:wb1.Windows.Item(1).WindowState = -4137  # xlNormal

            # --- ② Workbook2 も同様に ---
            $ws2 = $global:wb2.Sheets.Item($sheet)
            $ws2.Activate()
            $ws2.Cells.Item($row,$col).Select()
            $global:wb2.Activate()
            $global:wb2.Application.Visible = $false
            $global:wb2.Windows.Item(1).Visible = $false
            $global:wb2.Windows.Item(1).WindowState = -4137  # xlNormal

            # --- ③ 最後に両方を最大化（灰画面防止のため最後に実行） ---
            $global:wb1.Application.WindowState = 2  # xlMaximized
            $global:wb2.Application.WindowState = 2

            # --- ④ 明示的に最前面へ ---
            $global:wb1.Windows.Item(1).Activate()
            Start-Sleep -Milliseconds 200
            $global:wb2.Windows.Item(1).Activate()
        }
        else {
            [System.Windows.Forms.MessageBox]::Show("ブックが開かれていません。再比較を実行してください。")
        }
    } catch {
        [System.Windows.Forms.MessageBox]::Show("セル表示に失敗しました: $sheet $colName$row`n$($_.Exception.Message)")
    }
})

# --- フォーム閉鎖時にExcelを終了 ---
$form.Add_FormClosed({
    try {
        if ($global:wb1) { $global:wb1.Close($false) }
        if ($global:wb2) { $global:wb2.Close($false) }
        if ($global:wb1.Application) { $global:wb1.Application.Quit() }
        if ($global:wb2.Application) { $global:wb2.Application.Quit() }
    } catch {}
})

[void]$form.ShowDialog()
