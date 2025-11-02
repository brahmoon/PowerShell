#requires -Version 5.1
<#+
.SYNOPSIS
    psBrowserPilot HTTP PowerShell bridge.
.DESCRIPTION
    Exposes a lightweight HTTP API that lets a browser-based NodeFlow client
    create isolated PowerShell sessions (runspaces) and execute scripts.
    The service provides:
      * /health                - Service health information
      * /sessions (GET, POST)  - List or create sessions
      * /sessions/{id} (GET, DELETE) - Inspect or dispose a session
      * /sessions/{id}/history - Retrieve execution history for the session
      * /commands (POST)       - Execute a script (optionally bound to a session)
    Responses are JSON encoded and include stdout/stderr style diagnostics.
#>

Add-Type -AssemblyName System.Web

$ErrorActionPreference = 'Stop'

# =============================
# === Configuration & state ===
# =============================

$listenerPrefix = 'http://127.0.0.1:8080/'
$listener = [System.Net.HttpListener]::new()
$listener.Prefixes.Clear()
$listener.Prefixes.Add($listenerPrefix)
$listener.Start()

Write-Host "✅ psBrowserPilot listening on $listenerPrefix"

$runspacePool = [runspacefactory]::CreateRunspacePool(1, 4)
$runspacePool.Open()

$sessionStore = [System.Collections.Concurrent.ConcurrentDictionary[string, hashtable]]::new()

# =============================
# === Helper functions      ===
# =============================

function Write-JsonResponse {
    param(
        [Parameter(Mandatory)][System.Net.HttpListenerContext]$Context,
        [Parameter(Mandatory)][object]$Data,
        [int]$StatusCode = 200
    )

    $response = $Context.Response
    $response.Headers['Access-Control-Allow-Origin'] = '*'
    $response.Headers['Access-Control-Allow-Methods'] = 'GET, POST, DELETE, OPTIONS'
    $response.Headers['Access-Control-Allow-Headers'] = 'Content-Type'
    $response.StatusCode = $StatusCode

    if ($Context.Request.HttpMethod -eq 'OPTIONS') {
        $response.StatusCode = 204
        $response.Close()
        return
    }

    $json = $Data | ConvertTo-Json -Depth 6
    $buffer = [System.Text.Encoding]::UTF8.GetBytes($json)
    $response.ContentType = 'application/json'
    $response.ContentEncoding = [System.Text.Encoding]::UTF8
    $response.ContentLength64 = $buffer.Length
    $response.OutputStream.Write($buffer, 0, $buffer.Length)
    $response.OutputStream.Flush()
    $response.OutputStream.Close()
}

function Read-JsonBody {
    param(
        [Parameter(Mandatory)][System.Net.HttpListenerRequest]$Request
    )

    if (-not $Request.HasEntityBody) {
        return $null
    }

    $stream = $Request.InputStream
    $encoding = $Request.ContentEncoding
    $reader = New-Object System.IO.StreamReader($stream, $encoding)
    try {
        $body = $reader.ReadToEnd()
        if (-not [string]::IsNullOrWhiteSpace($body)) {
            return $body | ConvertFrom-Json -ErrorAction Stop
        }
        return $null
    } finally {
        $reader.Dispose()
    }
}

function Close-Session {
    param(
        [Parameter(Mandatory)][string]$SessionId,
        [switch]$Silent
    )

    if (-not $sessionStore.ContainsKey($SessionId)) {
        return $false
    }

    $session = $sessionStore[$SessionId]
    if ($null -ne $session.Runspace) {
        try { $session.Runspace.Close() } catch {}
        try { $session.Runspace.Dispose() } catch {}
    }

    $removed = $null
    [void]$sessionStore.TryRemove($SessionId, [ref]$removed)
    if (-not $Silent) {
        Write-Host "🧹 Closed session $SessionId"
    }
    return $true
}

function Convert-StreamRecord {
    param(
        [Parameter(Mandatory)]$Record
    )

    return [pscustomobject]@{
        message   = $Record.ToString()
        category  = $Record.CategoryInfo.Category
        exception = $Record.Exception.Message
        fullyQualifiedErrorId = $Record.FullyQualifiedErrorId
    }
}

function Invoke-BrowserScript {
    param(
        [Parameter(Mandatory)][string]$Script,
        [hashtable]$Variables,
        [object[]]$InputObjects,
        [string]$SessionId
    )

    $session = $null
    $ps = [PowerShell]::Create()
    $scriptStopwatch = [System.Diagnostics.Stopwatch]::StartNew()

    try {
        if ($SessionId) {
            if (-not $sessionStore.TryGetValue($SessionId, [ref]$session)) {
                return [pscustomobject]@{
                    ok       = $false
                    exitCode = 1
                    error    = "Session not found: $SessionId"
                }
            }
            $ps.Runspace = $session.Runspace
            if ($Variables) {
                foreach ($pair in $Variables.GetEnumerator()) {
                    $session.Runspace.SessionStateProxy.SetVariable($pair.Key, $pair.Value)
                }
            }
        } else {
            $ps.RunspacePool = $runspacePool
            if ($Variables) {
                $ps.AddScript(@'
param($__bindings)
foreach ($entry in $__bindings.GetEnumerator()) {
    Set-Variable -Name $entry.Key -Value $entry.Value -Scope Script
}
'@).AddArgument($Variables)
            }
        }

        $ps.AddScript($Script)
        $output = if ($InputObjects) { $ps.Invoke($InputObjects) } else { $ps.Invoke() }

        $stdout = ''
        if ($output) {
            $stdout = ($output | ForEach-Object { $_ | Out-String }) -join ''
        }

        $errors = @()
        if ($ps.Streams.Error.Count -gt 0) {
            $errors = $ps.Streams.Error | ForEach-Object { Convert-StreamRecord $_ }
        }

        $warnings = $ps.Streams.Warning | ForEach-Object { $_.Message }
        $verbose  = $ps.Streams.Verbose | ForEach-Object { $_.Message }
        $information = $ps.Streams.Information | ForEach-Object {
            [pscustomobject]@{
                message = $_.MessageData
                tags    = $_.Tags
            }
        }

        $duration = [math]::Round($scriptStopwatch.Elapsed.TotalMilliseconds, 2)
        $result = [pscustomobject]@{
            ok          = ($ps.HadErrors -eq $false)
            exitCode    = if ($ps.HadErrors) { 1 } else { 0 }
            stdout      = $stdout.TrimEnd()
            objects     = $output | ForEach-Object {
                try {
                    $_ | ConvertTo-Json -Depth 6 -Compress
                } catch {
                    $_.ToString()
                }
            }
            errors      = $errors
            warnings    = $warnings
            verbose     = $verbose
            information = $information
            durationMs  = $duration
        }

        if ($SessionId -and $session) {
            $historyEntry = [pscustomobject]@{
                timestamp = (Get-Date).ToString('o')
                script    = $Script
                ok        = $result.ok
                exitCode  = $result.exitCode
                durationMs = $result.durationMs
            }
            $session.History.Add($historyEntry) | Out-Null
        }

        return $result
    } catch {
        return [pscustomobject]@{
            ok       = $false
            exitCode = 1
            error    = $_.Exception.Message
            detail   = $_.Exception.ToString()
        }
    } finally {
        $scriptStopwatch.Stop()
        $ps.Dispose()
    }
}

function Create-Session {
    param(
        [string]$Name
    )

    $runspace = [runspacefactory]::CreateRunspace()
    $runspace.ApartmentState = 'MTA'
    $runspace.ThreadOptions = 'ReuseThread'
    $runspace.Open()

    $sessionId = [guid]::NewGuid().ToString()
    $sessionData = [hashtable]@{
        Id       = $sessionId
        Name     = $Name
        Created  = Get-Date
        Runspace = $runspace
        History  = New-Object System.Collections.Generic.List[object]
    }
    if (-not $sessionStore.TryAdd($sessionId, $sessionData)) {
        $runspace.Dispose()
        throw "Failed to register session"
    }

    Write-Host "✨ Created session $sessionId ($Name)"

    return $sessionData
}

function Get-SessionSummary {
    param(
        [Parameter(Mandatory)][hashtable]$Session
    )

    return [pscustomobject]@{
        id         = $Session.Id
        name       = if ($Session.Name) { $Session.Name } else { '' }
        created    = ($Session.Created).ToString('o')
        historyLen = $Session.History.Count
    }
}

function Handle-Request {
    param(
        [Parameter(Mandatory)][System.Net.HttpListenerContext]$Context
    )

    $request = $Context.Request
    $path = $request.Url.AbsolutePath.TrimEnd('/')
    if ($path -eq '') { $path = '/' }

    if ($request.HttpMethod -eq 'OPTIONS') {
        Write-JsonResponse -Context $Context -Data @{ ok = $true }
        return
    }

    try {
        switch -Regex ($path) {
            '^/health$' {
                if ($request.HttpMethod -ne 'GET') { break }
                $payload = [pscustomobject]@{
                    ok = $true
                    server = 'psBrowserPilot'
                    version = '1.0'
                    activeSessions = $sessionStore.Count
                    runspacePool = [pscustomobject]@{
                        minThreads = $runspacePool.GetMinRunspaces()
                        maxThreads = $runspacePool.GetMaxRunspaces()
                        available  = $runspacePool.GetAvailableRunspaces()
                    }
                    timestamp = (Get-Date).ToString('o')
                }
                Write-JsonResponse -Context $Context -Data $payload
                return
            }

            '^/sessions$' {
                if ($request.HttpMethod -eq 'GET') {
                    $sessions = $sessionStore.GetEnumerator() | ForEach-Object { Get-SessionSummary $_.Value }
                    Write-JsonResponse -Context $Context -Data @{ ok = $true; sessions = $sessions }
                    return
                }
                if ($request.HttpMethod -eq 'POST') {
                    $body = Read-JsonBody -Request $request
                    $name = $body?.name
                    $session = Create-Session -Name $name
                    Write-JsonResponse -Context $Context -StatusCode 201 -Data @{
                        ok = $true
                        session = Get-SessionSummary $session
                    }
                    return
                }
                break
            }

            '^/sessions/(?<id>[0-9a-fA-F-]+)$' {
                $sessionId = $Matches['id']
                if (-not $sessionStore.ContainsKey($sessionId)) {
                    Write-JsonResponse -Context $Context -StatusCode 404 -Data @{ ok = $false; error = "Session not found" }
                    return
                }
                $session = $sessionStore[$sessionId]
                switch ($request.HttpMethod) {
                    'GET' {
                        $details = Get-SessionSummary $session
                        $details | Add-Member -MemberType NoteProperty -Name 'history' -Value $session.History.ToArray()
                        Write-JsonResponse -Context $Context -Data @{ ok = $true; session = $details }
                        return
                    }
                    'DELETE' {
                        Close-Session -SessionId $sessionId | Out-Null
                        Write-JsonResponse -Context $Context -Data @{ ok = $true; closed = $sessionId }
                        return
                    }
                }
                break
            }

            '^/sessions/(?<id>[0-9a-fA-F-]+)/history$' {
                $sessionId = $Matches['id']
                if (-not $sessionStore.ContainsKey($sessionId)) {
                    Write-JsonResponse -Context $Context -StatusCode 404 -Data @{ ok = $false; error = 'Session not found' }
                    return
                }
                $session = $sessionStore[$sessionId]
                Write-JsonResponse -Context $Context -Data @{
                    ok = $true
                    session = $sessionId
                    history = $session.History.ToArray()
                }
                return
            }

            '^/commands$' {
                if ($request.HttpMethod -ne 'POST') { break }
                $body = Read-JsonBody -Request $request
                if (-not $body) {
                    Write-JsonResponse -Context $Context -StatusCode 400 -Data @{ ok = $false; error = 'Request body is required' }
                    return
                }
                $script = $body.script
                if (-not $script) {
                    Write-JsonResponse -Context $Context -StatusCode 400 -Data @{ ok = $false; error = 'Script is required' }
                    return
                }
                $variables = $null
                if ($body.variables) {
                    if ($body.variables -is [hashtable]) {
                        $variables = $body.variables
                    } else {
                        $variables = @{}
                        foreach ($prop in $body.variables.PSObject.Properties) {
                            $variables[$prop.Name] = $prop.Value
                        }
                    }
                }
                $inputObjects = if ($body.input) {
                    if ($body.input -is [System.Array]) { [object[]]$body.input } else { @($body.input) }
                }

                $result = Invoke-BrowserScript -Script $script -Variables $variables -InputObjects $inputObjects -SessionId $body.sessionId
                if ($result.ok) {
                    Write-JsonResponse -Context $Context -Data $result
                } else {
                    Write-JsonResponse -Context $Context -StatusCode 500 -Data $result
                }
                return
            }
        }

        Write-JsonResponse -Context $Context -StatusCode 404 -Data @{ ok = $false; error = "Endpoint not found: $path" }
    } catch {
        Write-JsonResponse -Context $Context -StatusCode 500 -Data @{
            ok = $false
            error = $_.Exception.Message
            detail = $_.Exception.ToString()
        }
    }
}

# =============================
# === Server loop           ===
# =============================

try {
    while ($listener.IsListening) {
        $context = $listener.GetContext()
        Handle-Request -Context $context
    }
} finally {
    Write-Host '⏹️  Stopping psBrowserPilot...'
    $listener.Stop()
    $listener.Close()
    foreach ($sessionId in @($sessionStore.Keys)) {
        Close-Session -SessionId $sessionId -Silent
    }
    if ($runspacePool) {
        $runspacePool.Close()
        $runspacePool.Dispose()
    }
}