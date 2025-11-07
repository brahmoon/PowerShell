Add-Type -AssemblyName System.Web
Add-Type -AssemblyName System.Management.Automation

$script:ContentRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$prefix = "http://127.0.0.1:8787/"
$listener = [System.Net.HttpListener]::new()
$listener.Prefixes.Add($prefix)
$listener.Start()
Write-Host "✅ psBrowserPilot server running at $prefix"
Write-Host "   Serving static files from $script:ContentRoot"

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

function Get-ContentType {
    param(
        [Parameter(Mandatory)][string]$Extension
    )

    switch ($Extension.ToLowerInvariant()) {
        '.html' { 'text/html; charset=utf-8' }
        '.htm' { 'text/html; charset=utf-8' }
        '.js' { 'application/javascript; charset=utf-8' }
        '.mjs' { 'application/javascript; charset=utf-8' }
        '.css' { 'text/css; charset=utf-8' }
        '.json' { 'application/json; charset=utf-8' }
        '.svg' { 'image/svg+xml' }
        '.png' { 'image/png' }
        '.jpg' { 'image/jpeg' }
        '.jpeg' { 'image/jpeg' }
        '.gif' { 'image/gif' }
        '.ico' { 'image/x-icon' }
        '.woff' { 'font/woff' }
        '.woff2' { 'font/woff2' }
        '.ttf' { 'font/ttf' }
        '.map' { 'application/json; charset=utf-8' }
        default { 'application/octet-stream' }
    }
}

function Resolve-StaticPath {
    param(
        [Parameter()][string]$RequestPath
    )

    $decoded = [System.Web.HttpUtility]::UrlDecode($RequestPath)
    $relative = $decoded.TrimStart('/').Replace('/', [System.IO.Path]::DirectorySeparatorChar)

    if ([string]::IsNullOrWhiteSpace($relative)) {
        $relative = 'index.html'
    }

    $fullPath = [System.IO.Path]::GetFullPath((Join-Path $script:ContentRoot $relative))

    if (-not $fullPath.StartsWith($script:ContentRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
        return $null
    }

    if (Test-Path $fullPath -PathType Container) {
        $fullPath = Join-Path $fullPath 'index.html'
    }

    if (Test-Path $fullPath -PathType Leaf) {
        return $fullPath
    }

    return $null
}

function Send-StaticFile {
    param(
        [Parameter(Mandatory)][System.Net.HttpListenerContext]$Context,
        [Parameter(Mandatory)][string]$FilePath,
        [switch]$IsHead
    )

    $response = $Context.Response

    try {
        $buffer = [System.IO.File]::ReadAllBytes($FilePath)
    } catch {
        $response.StatusCode = 500
        $response.StatusDescription = 'Failed to read file'
        $response.Close()
        return
    }

    $extension = [System.IO.Path]::GetExtension($FilePath)
    $response.ContentType = Get-ContentType -Extension $extension
    $response.ContentLength64 = $buffer.Length

    if (-not $IsHead) {
        $response.OutputStream.Write($buffer, 0, $buffer.Length)
    }

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

function Get-NodesDirectory {
    if (-not $script:NodesDirectory) {
        $script:NodesDirectory = Join-Path $script:ContentRoot 'nodes'
    }
    return $script:NodesDirectory
}

function Ensure-NodesDirectory {
    $dir = Get-NodesDirectory
    if (-not (Test-Path $dir)) {
        $null = New-Item -ItemType Directory -Path $dir -Force
    }
    return $dir
}

function Sanitize-NodeId {
    param(
        [Parameter(Mandatory)][string]$Id
    )
    $trimmed = $Id.Trim()
    if (-not $trimmed) { return '' }
    $normalized = $trimmed -replace '\s+', '_' -replace '[^A-Za-z0-9_]', '_'
    return $normalized
}

function Read-CustomNodeSpecs {
    $dir = Ensure-NodesDirectory
    $files = @(Get-ChildItem -Path $dir -Filter '*.json' -ErrorAction SilentlyContinue)
    $nodes = @()
    foreach ($file in $files) {
        try {
            $raw = Get-Content -Path $file.FullName -Raw -Encoding UTF8
            if ([string]::IsNullOrWhiteSpace($raw)) { continue }
            $parsed = $raw | ConvertFrom-Json -ErrorAction Stop
            if ($parsed) {
                $nodes += $parsed
            }
        } catch {
            Write-Warning "Failed to read custom node from $($file.FullName): $($_.Exception.Message)"
        }
    }
    return $nodes
}

function Write-CustomNodeSpec {
    param(
        [Parameter(Mandatory)][psobject]$Spec
    )

    $rawId = [string]$Spec.id
    $sanitizedId = Sanitize-NodeId -Id $rawId
    if (-not $sanitizedId) {
        throw "Spec id is required."
    }
    $Spec.id = $sanitizedId
    $dir = Ensure-NodesDirectory
    $path = Join-Path $dir "$sanitizedId.json"
    try {
        $json = $Spec | ConvertTo-Json -Depth 8
        [System.IO.File]::WriteAllText($path, $json, [System.Text.Encoding]::UTF8)
    } catch {
        throw "Failed to write custom node: $($_.Exception.Message)"
    }
    return $Spec
}

function Remove-CustomNodeSpec {
    param(
        [Parameter(Mandatory)][string]$Id
    )
    $sanitizedId = Sanitize-NodeId -Id $Id
    if (-not $sanitizedId) { return }
    $dir = Ensure-NodesDirectory
    $path = Join-Path $dir "$sanitizedId.json"
    if (Test-Path $path) {
        Remove-Item -Path $path -Force -ErrorAction SilentlyContinue
    }
}

function Invoke-PowerShellScript {
    param(
        [Parameter(Mandatory)][string]$Script
    )

    $runspace = $null
    $ps = $null
    try {
        $runspace = [runspacefactory]::CreateRunspace()
        $runspace.ApartmentState = [Threading.ApartmentState]::STA
        $runspace.ThreadOptions = [System.Management.Automation.Runspaces.PSThreadOptions]::ReuseThread
        $runspace.Open()

        $ps = [PowerShell]::Create()
        $ps.Runspace = $runspace

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
        if ($ps -ne $null) {
            $ps.Dispose()
        }
        if ($runspace -ne $null) {
            $runspace.Close()
            $runspace.Dispose()
        }
    }
}

function Handle-Request {
    param(
        [Parameter(Mandatory)][System.Net.HttpListenerContext]$Context
    )

    $request = $Context.Request
    $path = $request.Url.AbsolutePath

    if ($request.HttpMethod -in 'GET', 'HEAD') {
        $staticFile = Resolve-StaticPath -RequestPath $path
        if ($staticFile) {
            Send-StaticFile -Context $Context -FilePath $staticFile -IsHead:($request.HttpMethod -eq 'HEAD')
            return
        }
    }

    $normalizedPath = $path.ToLowerInvariant()

    switch ($normalizedPath) {
        '/' {
            $data = @{
                ok = $true
                message = 'psBrowserPilot server online'
            }
            Send-Json -Context $Context -Data $data
        }
        '/nodes/list' {
            try {
                $nodes = Read-CustomNodeSpecs
                Send-Json -Context $Context -Data @{ ok = $true; nodes = $nodes }
            } catch {
                Send-Json -Context $Context -Data @{ ok = $false; error = $_.Exception.Message } -StatusCode 500
            }
        }
        '/nodes/save' {
            if ($request.HttpMethod -ne 'POST') {
                Send-Json -Context $Context -Data @{ ok = $false; error = 'Use POST with JSON body {"spec": {...}}' } -StatusCode 405
                break
            }
            try {
                $payload = Get-JsonBody -Context $Context
            } catch {
                Send-Json -Context $Context -Data @{ ok = $false; error = $_ } -StatusCode 400
                break
            }
            if (-not $payload -or -not $payload.PSObject.Properties['spec']) {
                Send-Json -Context $Context -Data @{ ok = $false; error = 'Missing "spec" in request body.' } -StatusCode 400
                break
            }
            try {
                $saved = Write-CustomNodeSpec($payload.spec)
                $nodes = Read-CustomNodeSpecs
                Send-Json -Context $Context -Data @{ ok = $true; spec = $saved; nodes = $nodes }
            } catch {
                Send-Json -Context $Context -Data @{ ok = $false; error = $_.Exception.Message } -StatusCode 500
            }
        }
        '/nodes/delete' {
            if ($request.HttpMethod -ne 'POST') {
                Send-Json -Context $Context -Data @{ ok = $false; error = 'Use POST with JSON body {"id": "..."}' } -StatusCode 405
                break
            }
            try {
                $payload = Get-JsonBody -Context $Context
            } catch {
                Send-Json -Context $Context -Data @{ ok = $false; error = $_ } -StatusCode 400
                break
            }
            $rawId = [string]$payload.id
            if ([string]::IsNullOrWhiteSpace($rawId)) {
                Send-Json -Context $Context -Data @{ ok = $false; error = 'Missing "id" in request body.' } -StatusCode 400
                break
            }
            try {
                Remove-CustomNodeSpec -Id $rawId
                $nodes = Read-CustomNodeSpecs
                Send-Json -Context $Context -Data @{ ok = $true; nodes = $nodes }
            } catch {
                Send-Json -Context $Context -Data @{ ok = $false; error = $_.Exception.Message } -StatusCode 500
            }
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
            if ($request.HttpMethod -eq 'OPTIONS') {
                $Context.Response.Headers["Access-Control-Allow-Origin"] = "*"
                $Context.Response.Headers["Access-Control-Allow-Methods"] = "POST, GET, OPTIONS"
                $Context.Response.Headers["Access-Control-Allow-Headers"] = "Content-Type"
                $Context.Response.StatusCode = 204
                $Context.Response.Close()
            } else {
                $data = @{ ok = $false; error = "Not found: $normalizedPath" }
                Send-Json -Context $Context -Data $data -StatusCode 404
            }
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
pause