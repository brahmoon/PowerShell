param(
    [int]$Port = 8085,
    [string]$Prefix,
    [switch]$Quiet
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if (-not $Prefix) {
    $Prefix = "http://+:$Port/"
}

if (-not [System.Net.HttpListener]::IsSupported) {
    throw 'HttpListener is not supported on this platform.'
}

$listener = [System.Net.HttpListener]::new()
$null = $listener.Prefixes.Add($Prefix)

$sessionStore = New-Object System.Collections.Concurrent.ConcurrentDictionary[string, object]

function Get-ErrorMessage {
    param($ErrorRecord)

    if ($ErrorRecord -is [System.Management.Automation.ErrorRecord]) {
        return $ErrorRecord.Exception.Message
    }

    if ($ErrorRecord -is [System.Exception]) {
        return $ErrorRecord.Message
    }

    return [string]$ErrorRecord
}

function New-RunspaceSession {
    param(
        [string[]]$Modules,
        [string]$InitialScript
    )

    $runspace = [runspacefactory]::CreateRunspace()
    $runspace.Open()

    if ($Modules) {
        $powershell = [powershell]::Create()
        try {
            $powershell.Runspace = $runspace
            foreach ($moduleName in $Modules) {
                if ([string]::IsNullOrWhiteSpace($moduleName)) {
                    continue
                }
                $powershell.AddCommand('Import-Module').AddArgument($moduleName).AddStatement() | Out-Null
            }
            if ($InitialScript) {
                $powershell.AddScript($InitialScript) | Out-Null
            }
            $powershell.Invoke() | Out-Null
        } finally {
            $powershell.Dispose()
        }
    } elseif ($InitialScript) {
        $powershell = [powershell]::Create()
        try {
            $powershell.Runspace = $runspace
            $powershell.AddScript($InitialScript) | Out-Null
            $powershell.Invoke() | Out-Null
        } finally {
            $powershell.Dispose()
        }
    }

    $sessionId = [guid]::NewGuid().ToString('n')
    $moduleList = if ($Modules) { [string[]]$Modules } else { @() }

    $session = [pscustomobject]@{
        Id        = $sessionId
        Runspace  = $runspace
        Created   = Get-Date
        LastUsed  = Get-Date
        Modules   = $moduleList
        Lock      = New-Object object
    }

    if (-not $sessionStore.TryAdd($sessionId, $session)) {
        $runspace.Dispose()
        throw "Failed to store session $sessionId"
    }

    return $session
}

function Close-RunspaceSession {
    param([pscustomobject]$Session)

    if (-not $Session) {
        return
    }

    try {
        [System.Threading.Monitor]::Enter($Session.Lock)
        if ($Session.Runspace) {
            $Session.Runspace.Close()
            $Session.Runspace.Dispose()
        }
    } finally {
        if ($Session.Lock) {
            [System.Threading.Monitor]::Exit($Session.Lock)
        }
    }
}

function Read-RequestBody {
    param([System.Net.HttpListenerRequest]$Request)

    if (-not $Request.HasEntityBody) {
        return $null
    }

    $encoding = if ($Request.ContentEncoding) { $Request.ContentEncoding } else { [System.Text.Encoding]::UTF8 }
    $reader = New-Object System.IO.StreamReader($Request.InputStream, $encoding)
    try {
        return $reader.ReadToEnd()
    } finally {
        $reader.Dispose()
    }
}

function Write-JsonResponse {
    param(
        [System.Net.HttpListenerContext]$Context,
        [int]$StatusCode,
        $Body
    )

    $response = $Context.Response
    $response.StatusCode = $StatusCode
    $response.ContentType = 'application/json; charset=utf-8'
    $response.Headers['Access-Control-Allow-Origin'] = '*'
    $response.Headers['Access-Control-Allow-Methods'] = 'GET, POST, DELETE, OPTIONS'
    $response.Headers['Access-Control-Allow-Headers'] = 'Content-Type, X-Requested-With'

    $json = if ($Body -is [string]) { $Body } else { ($Body | ConvertTo-Json -Depth 8) }
    $buffer = [System.Text.Encoding]::UTF8.GetBytes($json)
    $response.ContentLength64 = $buffer.Length
    $response.OutputStream.Write($buffer, 0, $buffer.Length)
    $response.OutputStream.Close()
}

function Write-PlainResponse {
    param(
        [System.Net.HttpListenerContext]$Context,
        [int]$StatusCode,
        [string]$Body
    )

    $response = $Context.Response
    $response.StatusCode = $StatusCode
    $response.ContentType = 'text/plain; charset=utf-8'
    $response.Headers['Access-Control-Allow-Origin'] = '*'
    $response.Headers['Access-Control-Allow-Methods'] = 'GET, POST, DELETE, OPTIONS'
    $response.Headers['Access-Control-Allow-Headers'] = 'Content-Type, X-Requested-With'
    $buffer = [System.Text.Encoding]::UTF8.GetBytes($Body)
    $response.ContentLength64 = $buffer.Length
    $response.OutputStream.Write($buffer, 0, $buffer.Length)
    $response.OutputStream.Close()
}

function Parse-JsonBody {
    param([string]$Raw)

    if (-not $Raw) {
        return $null
    }

    try {
        return $Raw | ConvertFrom-Json -ErrorAction Stop
    } catch {
        throw "Invalid JSON payload: $($_.Exception.Message)"
    }
}

function Get-SessionById {
    param([string]$SessionId)

    if ([string]::IsNullOrWhiteSpace($SessionId)) {
        return $null
    }

    $null = $sessionStore.TryGetValue($SessionId, [ref]$session)
    return $session
}

function Execute-SessionCommand {
    param(
        [pscustomobject]$Session,
        [string]$Script,
        [hashtable]$Variables,
        [int]$TimeoutSeconds = 120
    )

    if (-not $Script) {
        throw 'Script payload is empty.'
    }

    $powershell = [powershell]::Create()
    $powershell.Runspace = $Session.Runspace
    $started = Get-Date
    $result = $null
    $hadTimeout = $false

    [System.Threading.Monitor]::Enter($Session.Lock)
    try {
        $Session.LastUsed = Get-Date
        $powershell.AddScript($Script, $true) | Out-Null

        if ($Variables) {
            foreach ($entry in $Variables.GetEnumerator()) {
                $powershell.AddParameter($entry.Key, $entry.Value) | Out-Null
            }
        }

        $asyncResult = $powershell.BeginInvoke($null, $null)
        if (-not $asyncResult.AsyncWaitHandle.WaitOne([TimeSpan]::FromSeconds($TimeoutSeconds))) {
            $hadTimeout = $true
            $powershell.Stop()
        }

        if ($hadTimeout) {
            throw "Script execution timed out after $TimeoutSeconds seconds."
        }

        $result = $powershell.EndInvoke($asyncResult)
    } finally {
        [System.Threading.Monitor]::Exit($Session.Lock)
    }

    $finished = Get-Date

    $stdout = ''
    if ($result) {
        try {
            $stdout = ($result | Out-String).TrimEnd()
        } catch {
            $stdout = ($result | ForEach-Object { $_.ToString() }) -join [Environment]::NewLine
        }
    }

    $structured = @()
    if ($result) {
        foreach ($item in $result) {
            try {
                $structured += [pscustomobject]@{
                    type  = $item.GetType().FullName
                    value = ($item | ConvertTo-Json -Depth 8)
                }
            } catch {
                $structured += [pscustomobject]@{
                    type  = $item.GetType().FullName
                    value = $item.ToString()
                }
            }
        }
    }

    $errors = $powershell.Streams.Error | ForEach-Object { $_.ToString() }
    $verbose = $powershell.Streams.Verbose | ForEach-Object { $_.Message }
    $warning = $powershell.Streams.Warning | ForEach-Object { $_.Message }
    $debug = $powershell.Streams.Debug | ForEach-Object { $_.Message }
    $information = $powershell.Streams.Information | ForEach-Object { $_.ToString() }

    $hadErrors = $powershell.HadErrors -or ($errors -and $errors.Count -gt 0)

    $powershell.Dispose()

    return [pscustomobject]@{
        sessionId        = $Session.Id
        success          = -not $hadErrors
        exitCode         = if ($hadErrors) { 1 } else { 0 }
        startedAt        = $started.ToUniversalTime().ToString('o')
        finishedAt       = $finished.ToUniversalTime().ToString('o')
        durationMs       = [math]::Round(($finished - $started).TotalMilliseconds, 2)
        output           = $stdout
        results          = $structured
        errors           = $errors
        verbose          = $verbose
        warnings         = $warning
        debug            = $debug
        information      = $information
        modules          = $Session.Modules
    }
}

try {
    $listener.Start()
    if (-not $Quiet) {
        Write-Host "psBrowserPilot listening on $($listener.Prefixes -join ', ')" -ForegroundColor Cyan
    }

    while ($listener.IsListening) {
        try {
            $context = $listener.GetContext()
        } catch [System.Net.HttpListenerException] {
            break
        } catch {
            if (-not $Quiet) {
                Write-Warning $_.Exception.Message
            }
            continue
        }

        [System.Threading.ThreadPool]::QueueUserWorkItem({
            param($ctx)
            try {
                $request = $ctx.Request
                $response = $ctx.Response
                $path = $request.Url.AbsolutePath.TrimEnd('/')
                if (-not $path) { $path = '/' }

                if ($request.HttpMethod -eq 'OPTIONS') {
                    $response.StatusCode = 200
                    $response.Headers['Access-Control-Allow-Origin'] = '*'
                    $response.Headers['Access-Control-Allow-Methods'] = 'GET, POST, DELETE, OPTIONS'
                    $response.Headers['Access-Control-Allow-Headers'] = 'Content-Type, X-Requested-With'
                    $response.OutputStream.Close()
                    return
                }

                switch ($request.HttpMethod) {
                    'GET' {
                        if ($path -eq '/' -or $path -eq '/healthz') {
                            Write-PlainResponse -Context $ctx -StatusCode 200 -Body 'psBrowserPilot is running.'
                            return
                        }

                        if ($path -eq '/sessions') {
                            $sessions = $sessionStore.Values | ForEach-Object {
                                [pscustomobject]@{
                                    id       = $_.Id
                                    created  = $_.Created.ToUniversalTime().ToString('o')
                                    lastUsed = $_.LastUsed.ToUniversalTime().ToString('o')
                                    modules  = $_.Modules
                                }
                            }
                            Write-JsonResponse -Context $ctx -StatusCode 200 -Body $sessions
                            return
                        }

                        if ($path -like '/sessions/*') {
                            $sessionId = $path.Split('/')[-1]
                            $session = Get-SessionById -SessionId $sessionId
                            if (-not $session) {
                                Write-JsonResponse -Context $ctx -StatusCode 404 -Body @{ error = 'Session not found.' }
                                return
                            }
                            $body = [pscustomobject]@{
                                id       = $session.Id
                                created  = $session.Created.ToUniversalTime().ToString('o')
                                lastUsed = $session.LastUsed.ToUniversalTime().ToString('o')
                                modules  = $session.Modules
                            }
                            Write-JsonResponse -Context $ctx -StatusCode 200 -Body $body
                            return
                        }

                        Write-JsonResponse -Context $ctx -StatusCode 404 -Body @{ error = 'Not found.' }
                        return
                    }
                    'POST' {
                        if ($path -eq '/sessions') {
                            $raw = Read-RequestBody -Request $request
                            $payload = $null
                            if ($raw) {
                                try {
                                    $payload = Parse-JsonBody -Raw $raw
                                } catch {
                                    Write-JsonResponse -Context $ctx -StatusCode 400 -Body @{ error = Get-ErrorMessage $_ }
                                    return
                                }
                            }
                            $modules = @()
                            $initialScript = $null
                            if ($payload) {
                                if ($payload.modules -and $payload.modules -is [System.Collections.IEnumerable]) {
                                    foreach ($item in $payload.modules) {
                                        if ([string]::IsNullOrWhiteSpace($item)) { continue }
                                        $modules += [string]$item
                                    }
                                }
                                if ($payload.initialScript) {
                                    $initialScript = [string]$payload.initialScript
                                }
                            }
                            try {
                                $session = New-RunspaceSession -Modules $modules -InitialScript $initialScript
                            } catch {
                                Write-JsonResponse -Context $ctx -StatusCode 500 -Body @{ error = Get-ErrorMessage $_ }
                                return
                            }

                            $body = [pscustomobject]@{
                                sessionId = $session.Id
                                createdAt = $session.Created.ToUniversalTime().ToString('o')
                                modules   = $session.Modules
                            }
                            Write-JsonResponse -Context $ctx -StatusCode 201 -Body $body
                            return
                        }

                        if ($path -eq '/commands') {
                            $raw = Read-RequestBody -Request $request
                            if (-not $raw) {
                                Write-JsonResponse -Context $ctx -StatusCode 400 -Body @{ error = 'Missing request body.' }
                                return
                            }
                            try {
                                $payload = Parse-JsonBody -Raw $raw
                            } catch {
                                Write-JsonResponse -Context $ctx -StatusCode 400 -Body @{ error = Get-ErrorMessage $_ }
                                return
                            }

                            $sessionId = [string]$payload.sessionId
                            $script = [string]$payload.script
                            $timeoutSeconds = if ($payload.timeoutSeconds) { [int]$payload.timeoutSeconds } else { 120 }
                            $variables = $null
                            if ($payload.variables) {
                                $variables = @{}
                                foreach ($entry in $payload.variables.psobject.Properties) {
                                    $variables[$entry.Name] = $entry.Value
                                }
                            }

                            $session = Get-SessionById -SessionId $sessionId
                            if (-not $session) {
                                Write-JsonResponse -Context $ctx -StatusCode 404 -Body @{ error = 'Session not found.' }
                                return
                            }

                            try {
                                $result = Execute-SessionCommand -Session $session `
                                    -Script $script `
                                    -Variables $variables `
                                    -TimeoutSeconds $timeoutSeconds
                                Write-JsonResponse -Context $ctx -StatusCode 200 -Body $result
                            } catch {
                                Write-JsonResponse -Context $ctx -StatusCode 500 -Body @{ error = Get-ErrorMessage $_; sessionId = $session.Id }
                            }
                            return
                        }

                        Write-JsonResponse -Context $ctx -StatusCode 404 -Body @{ error = 'Not found.' }
                        return
                    }
                    'DELETE' {
                        if ($path -like '/sessions/*') {
                            $sessionId = $path.Split('/')[-1]
                            $session = Get-SessionById -SessionId $sessionId
                            if (-not $session) {
                                Write-JsonResponse -Context $ctx -StatusCode 404 -Body @{ error = 'Session not found.' }
                                return
                            }

                            $removed = $null
                            $null = $sessionStore.TryRemove($sessionId, [ref]$removed)
                            Close-RunspaceSession -Session $session
                            Write-JsonResponse -Context $ctx -StatusCode 200 -Body @{ sessionId = $sessionId; removed = [bool]$removed }
                            return
                        }

                        Write-JsonResponse -Context $ctx -StatusCode 404 -Body @{ error = 'Not found.' }
                        return
                    }
                    default {
                        Write-JsonResponse -Context $ctx -StatusCode 405 -Body @{ error = 'Method not allowed.' }
                        return
                    }
                }
            } catch {
                try {
                        Write-JsonResponse -Context $ctx -StatusCode 500 -Body @{ error = Get-ErrorMessage $_ }
                    } catch {
                    }
                }
            }, $context) | Out-Null
    }
} finally {
    foreach ($item in $sessionStore.Values) {
        Close-RunspaceSession -Session $item
    }
    $listener.Stop()
    $listener.Close()
    if (-not $Quiet) {
        Write-Host 'psBrowserPilot listener stopped.'
    }
}
