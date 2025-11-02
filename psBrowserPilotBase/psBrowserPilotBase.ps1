Add-Type -AssemblyName System.Web
Add-Type -AssemblyName System.Management.Automation

$prefix = "http://127.0.0.1:8080/"
$listener = [System.Net.HttpListener]::new()
$listener.Prefixes.Add($prefix)
$listener.Start()
Write-Host "✅ PowerShell browser server running at $prefix"

function Send-Json {
    param(
        [Parameter(Mandatory)][System.Net.HttpListenerContext]$Context,
        [Parameter(Mandatory)]$Data,
        [int]$StatusCode = 200
    )

    $response = $Context.Response

    # Allow cross origin requests to simplify browser-side usage
    $response.Headers["Access-Control-Allow-Origin"] = "*"
    $response.Headers["Access-Control-Allow-Methods"] = "POST, GET, OPTIONS"
    $response.Headers["Access-Control-Allow-Headers"] = "Content-Type"

    if ($Context.Request.HttpMethod -eq "OPTIONS") {
        $response.StatusCode = 204
        $response.Close()
        return
    }

    $json = ($Data | ConvertTo-Json -Compress -Depth 6)
    $buffer = [System.Text.Encoding]::UTF8.GetBytes($json)

    $response.StatusCode = $StatusCode
    $response.ContentType = "application/json"
    $response.ContentEncoding = [System.Text.Encoding]::UTF8
    $response.ContentLength64 = $buffer.Length

    $response.OutputStream.Write($buffer, 0, $buffer.Length)
    $response.OutputStream.Flush()
    $response.OutputStream.Close()
}

function Get-JsonBody {
    param(
        [Parameter(Mandatory)][System.Net.HttpListenerContext]$Context
    )

    $request = $Context.Request
    if (-not $request.HasEntityBody) {
        return $null
    }

    $encoding = $request.ContentEncoding
    $contentType = $request.ContentType

    $charsetSpecified = $false
    if ($contentType) {
        try {
            $parsedContentType = [System.Net.Mime.ContentType]::new($contentType)
            if ($parsedContentType.CharSet) {
                $charsetSpecified = $true
                if (-not $encoding) {
                    try {
                        $encoding = [System.Text.Encoding]::GetEncoding($parsedContentType.CharSet)
                    } catch {
                        $encoding = $null
                    }
                }
            }
        } catch {
            # If the header cannot be parsed, fall back to inspecting the raw string.
            if ($contentType -match 'charset=') {
                $charsetSpecified = $true
                if (-not $encoding) {
                    try {
                        $rawCharset = ($contentType -split 'charset=')[1].Split(';')[0].Trim()
                        if ($rawCharset) {
                            $encoding = [System.Text.Encoding]::GetEncoding($rawCharset)
                        }
                    } catch {
                        $encoding = $null
                    }
                }
            }
        }
    }

    if (-not $charsetSpecified) {
        # Treat UTF-8 as the default when the client does not specify a charset.
        $encoding = [System.Text.Encoding]::UTF8
    } elseif (-not $encoding) {
        $encoding = [System.Text.Encoding]::UTF8
    }

    $reader = [System.IO.StreamReader]::new($request.InputStream, $encoding, $true)
    try {
        $raw = $reader.ReadToEnd()
    } finally {
        $reader.Close()
    }

    if (-not $raw) { return $null }

    try {
        return $raw | ConvertFrom-Json
    } catch {
        throw "Invalid JSON payload"
    }
}

function Invoke-PowerShellScript {
    param(
        [Parameter(Mandatory)][string]$Script
    )

    $ps = [PowerShell]::Create()
    try {
        $null = $ps.AddScript($Script)
        $null = $ps.AddCommand('Out-String')

        $output = $ps.Invoke()
        $errors = @()
        foreach ($err in $ps.Streams.Error) {
            $errors += $err.ToString().Trim()
        }

        $text = ($output | Out-String).TrimEnd()
        $hasErrors = $ps.HadErrors -or ($errors.Count -gt 0)

        return [pscustomobject]@{
            ok = -not $hasErrors
            output = $text
            errors = $errors
        }
    } catch {
        return [pscustomobject]@{
            ok = $false
            output = ""
            errors = @($_.Exception.Message)
        }
    } finally {
        $ps.Dispose()
    }
}

function Handle-Request {
    param(
        [Parameter(Mandatory)][System.Net.HttpListenerContext]$Context
    )

    $request = $Context.Request
    $path = $request.Url.AbsolutePath.ToLowerInvariant()

    switch ($path) {
        '/' {
            $data = @{
                ok = $true
                message = 'psBrowserPilotBase server online'
            }
            Send-Json -Context $Context -Data $data
        }
        '/runscript' {
            if ($request.HttpMethod -ne 'POST') {
                $data = @{ ok = $false; error = 'Use POST with JSON body {"script": "..."}' }
                Send-Json -Context $Context -Data $data -StatusCode 405
                break
            }

            try {
                $payload = Get-JsonBody -Context $Context
            } catch {
                $data = @{ ok = $false; error = $_ }
                Send-Json -Context $Context -Data $data -StatusCode 400
                break
            }

            if (-not $payload -or -not $payload.PSObject.Properties['script'] -or -not $payload.script) {
                $data = @{ ok = $false; error = 'Missing "script" in request body.' }
                Send-Json -Context $Context -Data $data -StatusCode 400
                break
            }

            $result = Invoke-PowerShellScript -Script ([string]$payload.script)
            Send-Json -Context $Context -Data $result
        }
        default {
            $data = @{ ok = $false; error = "Not found: $path" }
            Send-Json -Context $Context -Data $data -StatusCode 404
        }
    }
}

try {
    while ($listener.IsListening) {
        $context = $listener.GetContext()
        Handle-Request -Context $context
    }
} finally {
    $listener.Stop()
    $listener.Close()
}
