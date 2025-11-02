Add-Type -AssemblyName System.Web
Add-Type -AssemblyName System.Drawing

$prefix = "http://127.0.0.1:8080/"
$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add($prefix)
$listener.Start()
Write-Host "✅ PowerShell Excel server running at $prefix"

# ===== 共通ヘルパ =====

function Get-ExcelApp {
    try {
        $app = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
        Write-Host "Excel instance detected: $($app.Name)"
        return $app
    } catch {
        Write-Host "⚠️ Excel not found. Starting new instance..."
        $app = New-Object -ComObject Excel.Application
        $app.Visible = $true
        return $app
    }
}

function Send-Json($context, $obj) {
    $res = $context.Response
    # ---- CORS 対策ヘッダ ----
    $res.Headers.Add("Access-Control-Allow-Origin", "*")
    $res.Headers.Add("Access-Control-Allow-Methods", "POST, GET, OPTIONS")
    $res.Headers.Add("Access-Control-Allow-Headers", "Content-Type")

    # プリフライトリクエスト (OPTIONS)
    if ($context.Request.HttpMethod -eq "OPTIONS") {
        $res.StatusCode = 204
        $res.Close()
        return
    }

    # JSON応答
    $json = ($obj | ConvertTo-Json -Compress -Depth 6)
    $buffer = [System.Text.Encoding]::UTF8.GetBytes($json)
    $res.ContentType = "application/json"
    $res.ContentEncoding = [System.Text.Encoding]::UTF8
    $res.ContentLength64 = $buffer.Length
    $res.OutputStream.Write($buffer, 0, $buffer.Length)
    $res.OutputStream.Flush()
    $res.OutputStream.Close()
}

function Convert-HexToOleColor {
    param(
        [Parameter(Mandatory = $true)][string]$Hex
    )

    if (-not $Hex -or $Hex.Length -ne 7 -or $Hex[0] -ne '#') {
        throw "Invalid hex color: $Hex"
    }

    $r = [Convert]::ToInt32($Hex.Substring(1, 2), 16)
    $g = [Convert]::ToInt32($Hex.Substring(3, 2), 16)
    $b = [Convert]::ToInt32($Hex.Substring(5, 2), 16)
    return ($b -bor ($g -shl 8) -bor ($r -shl 16))
}

function Convert-OleToHex {
    param(
        [Parameter()][object]$Ole
    )

    if ($null -eq $Ole) { return $null }

    try {
        $color = [System.Drawing.ColorTranslator]::FromOle([int]$Ole)
        return ('#{0:X2}{1:X2}{2:X2}' -f $color.R, $color.G, $color.B)
    } catch {
        return $null
    }
}

# ===== 処理メイン =====

$global:ActiveWorkbook = $null

function Handle-Request($context) {
    $req = $context.Request
    $reader = New-Object IO.StreamReader($req.InputStream, [Text.Encoding]::UTF8)
    $body = $reader.ReadToEnd()
    $reader.Close()
    $params = @{}
    if ($body) { $params = ConvertFrom-Json $body }

    $excel = Get-ExcelApp
    $result = @{}

    switch -Regex ($req.Url.AbsolutePath) {

        # --- 開いているブック一覧 ---
        "^/listWorkbooks$" {
            $books = @()
            foreach ($wb in $excel.Workbooks) { $books += $wb.Name }
            $result = @{ ok = $true; workbooks = $books }
        }

        # --- ブック選択 ---
        "^/setWorkbook$" {
            $name = $params.name
            $found = $null
            foreach ($wb in $excel.Workbooks) {
                if ($wb.Name -eq $name) { $found = $wb; break }
            }
            if ($found) {
                $global:ActiveWorkbook = $found
                $result = @{ ok = $true; active = $found.Name }
            } else {
                $result = @{ ok = $false; error = "Workbook not found: $name" }
            }
        }

        # --- 選択セル情報取得 ---
        "^/getSelection$" {
            $wb = $global:ActiveWorkbook
            if (-not $wb) { $wb = $excel.ActiveWorkbook }
            if (-not $wb) {
                $result = @{ ok = $false; error = "No workbook selected" }
            } else {
                try {
                    $sel = $wb.Application.Selection
                    $result = @{
                        ok = $true
                        workbook = $wb.Name
                        sheet = $sel.Worksheet.Name
                        address = $sel.Address()
                        value = $sel.Text
                        font_color = @{
                            ole = $sel.Font.Color
                            hex = Convert-OleToHex $sel.Font.Color
                        }
                        interior_color = @{
                            ole = $sel.Interior.Color
                            hex = Convert-OleToHex $sel.Interior.Color
                        }
                    }
                } catch {
                    $result = @{ ok = $false; error = $_.Exception.Message }
                }
            }
        }

        # --- 選択範囲を塗りつぶし (後方互換含む) ---
        "^/(fillSelection|colorSelection)$" {
            $colorHex = $params.color
            if (-not $colorHex) {
                $result = @{ ok = $false; error = "Color parameter is required" }
                break
            }

            $wb = $global:ActiveWorkbook
            if (-not $wb) { $wb = $excel.ActiveWorkbook }

            if (-not $wb) {
                $result = @{ ok = $false; error = "No workbook selected" }
            } else {
                try {
                    $sel = $wb.Application.Selection
                    if ($sel -ne $null) {
                        $oleColor = Convert-HexToOleColor $colorHex
                        $sel.Interior.Color = $oleColor
                        $result = @{
                            ok = $true
                            workbook = $wb.Name
                            sheet = $sel.Worksheet.Name
                            range = $sel.Address()
                            color = $colorHex
                            target = "fill"
                        }
                    } else {
                        $result = @{ ok = $false; error = "No selection" }
                    }
                } catch {
                    $result = @{ ok = $false; error = $_.Exception.Message }
                }
            }
        }

        # --- 選択範囲のフォント色変更 ---
        "^/colorFont$" {
            $colorHex = $params.color
            if (-not $colorHex) {
                $result = @{ ok = $false; error = "Color parameter is required" }
                break
            }

            $wb = $global:ActiveWorkbook
            if (-not $wb) { $wb = $excel.ActiveWorkbook }

            if (-not $wb) {
                $result = @{ ok = $false; error = "No workbook selected" }
            } else {
                try {
                    $sel = $wb.Application.Selection
                    if ($sel -ne $null) {
                        $oleColor = Convert-HexToOleColor $colorHex
                        $sel.Font.Color = $oleColor
                        $result = @{
                            ok = $true
                            workbook = $wb.Name
                            sheet = $sel.Worksheet.Name
                            range = $sel.Address()
                            color = $colorHex
                            target = "font"
                        }
                    } else {
                        $result = @{ ok = $false; error = "No selection" }
                    }
                } catch {
                    $result = @{ ok = $false; error = $_.Exception.Message }
                }
            }
        }

        default {
            $result = @{ ok = $false; error = "Unknown endpoint: $($req.Url.AbsolutePath)" }
        }
    }

    Send-Json $context $result
}

# ===== メインループ =====
while ($true) {
    $context = $listener.GetContext()
    Handle-Request $context
}
