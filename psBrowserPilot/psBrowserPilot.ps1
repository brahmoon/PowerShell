Add-Type -AssemblyName System.Web

$prefix = "http://127.0.0.1:8080/"
$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add($prefix)
$listener.Start()
Write-Host "✅ PowerShell browser server running at $prefix"

$global:Sessions = [System.Collections.Concurrent.ConcurrentDictionary[string, object]]::new()

function Add-CorsHeaders {
    param([System.Net.HttpListenerResponse]$response)
    $response.Headers['Access-Control-Allow-Origin'] = '*'
    $response.Headers['Access-Control-Allow-Methods'] = 'GET, POST, DELETE, OPTIONS'
    $response.Headers['Access-Control-Allow-Headers'] = 'Content-Type'
}

function Send-Json {
    param(
        [System.Net.HttpListenerContext]$context,
        [int]$statusCode = 200,
        $payload = $null
    )

    $response = $context.Response
    Add-CorsHeaders -response $response
    $response.StatusCode = $statusCode

    if ($statusCode -eq 204 -or $null -eq $payload) {
        try { $response.OutputStream.Close() } catch { }
        return
    }

    try {
        $json = $payload | ConvertTo-Json -Depth 12
    } catch {
        $json = '{"ok":false,"error":"Failed to serialize response"}'
    }

    $buffer = [System.Text.Encoding]::UTF8.GetBytes($json)
    $response.ContentType = 'application/json; charset=utf-8'
    $response.ContentEncoding = [System.Text.Encoding]::UTF8
    $response.ContentLength64 = $buffer.Length
    $response.OutputStream.Write($buffer, 0, $buffer.Length)
    $response.OutputStream.Close()
}

function Read-JsonBody {
    param([System.Net.HttpListenerRequest]$request)

    if (-not $request.HasEntityBody) {
        return @{}
    }

    $encoding = $request.ContentEncoding
    if (-not $encoding) {
        $encoding = [System.Text.Encoding]::UTF8
    }

    $reader = New-Object System.IO.StreamReader($request.InputStream, $encoding)
    $body = $reader.ReadToEnd()
    $reader.Close()

    if ([string]::IsNullOrWhiteSpace($body)) {
        return @{}
    }

    try {
        $parsed = ConvertFrom-Json -InputObject $body -Depth 12
    } catch {
        throw "Invalid JSON payload: $($_.Exception.Message)"
    }

    return Convert-ToHashtable $parsed
}

function Convert-ToHashtable {
    param($obj)

    if ($null -eq $obj) {
        return @{}
    }
    if ($obj -is [hashtable]) {
        return $obj
    }

    $hash = @{}
    foreach ($prop in $obj.PSObject.Properties) {
        $hash[$prop.Name] = $prop.Value
    }
    return $hash
}

function Convert-ToArray {
    param($value)

    if ($null -eq $value) { return @() }
    if ($value -is [System.Collections.IEnumerable] -and -not ($value -is [string])) {
        return @($value)
    }
    return @($value)
}

function Convert-ToPlainText {
    param($value)

    if ($null -eq $value) { return '' }
    if ($value -is [string]) { return $value }

    try {
        return ($value | Out-String).TrimEnd("`r", "`n")
    } catch {
        return $value.ToString()
    }
}

function Format-ErrorRecord {
    param([System.Management.Automation.ErrorRecord]$record)

    if ($null -eq $record) { return $null }

    return [ordered]@{
        message = $record.Exception.Message
        category = $record.CategoryInfo.Category.ToString()
        fullyQualifiedErrorId = $record.FullyQualifiedErrorId
        scriptStackTrace = $record.ScriptStackTrace
        target = Convert-ToPlainText $record.TargetObject
    }
}

function Format-ProgressRecord {
    param([System.Management.Automation.ProgressRecord]$record)

    if ($null -eq $record) { return $null }

    return [ordered]@{
        activity = $record.Activity
        status = $record.StatusDescription
        currentOperation = $record.CurrentOperation
        percent = $record.PercentComplete
        secondsRemaining = $record.SecondsRemaining
    }
}

function Close-Runspace {
    param($session)

    if ($null -eq $session) { return }

    $runspace = $session.Runspace
    if ($runspace) {
        try {
            if ($runspace.RunspaceStateInfo.State -ne [System.Management.Automation.Runspaces.RunspaceState]::Closed) {
                $runspace.Close()
            }
        } catch {
        }
        try { $runspace.Dispose() } catch { }
    }
}

function Build-SessionResponse {
    param($session)

    return [ordered]@{
        id = $session.Id
        name = $session.Name
        created = $session.Created.ToString('o')
        metadata = $session.Metadata
        commandCount = $session.History.Count
    }
}

function Handle-SessionCreate {
    param($context)

    $payload = Read-JsonBody -request $context.Request
    $name = if ($payload.ContainsKey('name') -and $payload.name) { [string]$payload.name } else { "browser-session" }
    $metadata = Convert-ToHashtable $payload.metadata

    $iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault2()

    foreach ($module in Convert-ToArray $payload.modules) {
        $moduleName = [string]$module
        if ([string]::IsNullOrWhiteSpace($moduleName)) { continue }
        try {
            [void]$iss.ImportPSModule($moduleName)
        } catch {
            return Send-Json $context 400 ([ordered]@{
                ok = $false
                error = "Failed to import module '$moduleName'"
                detail = $_.Exception.Message
            })
        }
    }

    $runspace = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace($iss)
    $runspace.Open()

    $variables = Convert-ToHashtable $payload.variables
    foreach ($key in $variables.Keys) {
        $runspace.SessionStateProxy.SetVariable($key, $variables[$key])
    }

    $bootstrapResult = $null
    if ($payload.ContainsKey('bootstrapScript') -and $payload.bootstrapScript) {
        $bootstrapPS = [System.Management.Automation.PowerShell]::Create()
        $bootstrapPS.Runspace = $runspace
        $bootstrapPS.AddScript([string]$payload.bootstrapScript) | Out-Null
        $bootstrapOutput = $bootstrapPS.Invoke()
        $bootstrapErrors = $bootstrapPS.Streams.Error
        if ($bootstrapPS.HadErrors) {
            Close-Runspace ([pscustomobject]@{ Runspace = $runspace })
            $messages = ($bootstrapErrors | ForEach-Object { $_.Exception.Message }) -join '; '
            return Send-Json $context 400 ([ordered]@{
                ok = $false
                error = 'Bootstrap script failed'
                detail = $messages
            })
        }
        $bootstrapResult = @{
            output = $bootstrapOutput | ForEach-Object { Convert-ToPlainText $_ }
        }
        $bootstrapPS.Dispose()
    }

    $sessionId = [Guid]::NewGuid().ToString('N')
    $session = [pscustomobject]@{
        Id = $sessionId
        Name = $name
        Metadata = $metadata
        Created = Get-Date
        Runspace = $runspace
        History = New-Object System.Collections.Generic.List[object]
    }

    [void]$global:Sessions.TryAdd($sessionId, $session)

    $response = Build-SessionResponse $session
    if ($bootstrapResult) {
        $response.bootstrap = $bootstrapResult
    }

    Send-Json $context 201 ([ordered]@{
        ok = $true
        session = $response
    })
}

function Handle-SessionList {
    param($context)

    $items = @()
    foreach ($entry in $global:Sessions.GetEnumerator()) {
        $items += Build-SessionResponse $entry.Value
    }

    Send-Json $context 200 ([ordered]@{
        ok = $true
        sessions = $items
    })
}

function Handle-SessionGet {
    param($context, [string]$sessionId)

    $session = $null
    if (-not $global:Sessions.TryGetValue($sessionId, [ref]$session)) {
        return Send-Json $context 404 ([ordered]@{
            ok = $false
            error = "Session not found"
            sessionId = $sessionId
        })
    }

    Send-Json $context 200 ([ordered]@{
        ok = $true
        session = Build-SessionResponse $session
    })
}

function Handle-SessionDelete {
    param($context, [string]$sessionId)

    $removed = $null
    if (-not $global:Sessions.TryRemove($sessionId, [ref]$removed)) {
        return Send-Json $context 404 ([ordered]@{
            ok = $false
            error = 'Session not found'
            sessionId = $sessionId
        })
    }

    Close-Runspace $removed

    Send-Json $context 200 ([ordered]@{
        ok = $true
        sessionId = $sessionId
    })
}

function Handle-SessionHistory {
    param($context, [string]$sessionId)

    $session = $null
    if (-not $global:Sessions.TryGetValue($sessionId, [ref]$session)) {
        return Send-Json $context 404 ([ordered]@{
            ok = $false
            error = 'Session not found'
            sessionId = $sessionId
        })
    }

    $limitParam = $context.Request.QueryString['limit']
    $limit = 50
    if ($limitParam) {
        $tmpLimit = 0
        if ([int]::TryParse($limitParam, [ref]$tmpLimit)) {
            if ($tmpLimit -gt 0) { $limit = $tmpLimit }
        }
    }

    $history = $session.History.ToArray()
    if ($history.Length -gt $limit) {
        $start = $history.Length - $limit
        $history = $history[$start..($history.Length - 1)]
    }

    Send-Json $context 200 ([ordered]@{
        ok = $true
        sessionId = $sessionId
        commandCount = $session.History.Count
        history = $history
    })
}

function Handle-CommandInvoke {
    param($context, [string]$sessionId)

    $session = $null
    if (-not $global:Sessions.TryGetValue($sessionId, [ref]$session)) {
        return Send-Json $context 404 ([ordered]@{
            ok = $false
            error = 'Session not found'
            sessionId = $sessionId
        })
    }

    $payload = Read-JsonBody -request $context.Request
    $script = [string]$payload.script
    if ([string]::IsNullOrWhiteSpace($script)) {
        return Send-Json $context 400 ([ordered]@{
            ok = $false
            error = 'Script is required'
        })
    }

    $metadata = Convert-ToHashtable $payload.metadata
    $parameters = Convert-ToHashtable $payload.parameters
    $inputData = $payload.input
    $commandId = [Guid]::NewGuid().ToString('N')

    $ps = [System.Management.Automation.PowerShell]::Create()
    $ps.Runspace = $session.Runspace
    $ps.AddScript($script, $true) | Out-Null
    foreach ($key in $parameters.Keys) {
        $ps.AddParameter($key, $parameters[$key]) | Out-Null
    }

    $started = Get-Date
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    $output = @()
    $exceptionMessage = $null

    try {
        if ($null -ne $inputData) {
            $invokeInput = Convert-ToArray $inputData
            $output = $ps.Invoke($invokeInput)
        } else {
            $output = $ps.Invoke()
        }
    } catch {
        $exceptionMessage = $_.Exception.Message
    }

    $stopwatch.Stop()
    $ended = Get-Date

    $textOutput = $output | ForEach-Object { Convert-ToPlainText $_ }
    $errors = @()
    foreach ($err in $ps.Streams.Error) {
        $errors += Format-ErrorRecord $err
    }

    if ($exceptionMessage) {
        $errors += [ordered]@{
            message = $exceptionMessage
            category = 'UnhandledException'
            fullyQualifiedErrorId = 'UnhandledException'
            scriptStackTrace = $null
            target = $null
        }
    }

    $warnings = @($ps.Streams.Warning | ForEach-Object { @{ message = $_.Message } })
    $verbose = @($ps.Streams.Verbose | ForEach-Object { @{ message = $_.Message } })
    $debug = @($ps.Streams.Debug | ForEach-Object { @{ message = $_.Message } })
    $information = @($ps.Streams.Information | ForEach-Object {
        @{ message = Convert-ToPlainText $_.MessageData; source = $_.Source }
    })
    $progress = @($ps.Streams.Progress | ForEach-Object { Format-ProgressRecord $_ })

    $hadErrors = $ps.HadErrors -or (-not [string]::IsNullOrEmpty($exceptionMessage))
    $success = -not $hadErrors

    $entry = [ordered]@{
        id = $commandId
        sessionId = $session.Id
        script = $script
        started = $started.ToString('o')
        ended = $ended.ToString('o')
        durationMs = [int][Math]::Round($stopwatch.Elapsed.TotalMilliseconds)
        ok = $success
        hadErrors = $hadErrors
        metadata = $metadata
        result = [ordered]@{
            output = $textOutput
            outputText = [string]::Join("`n", $textOutput)
            streams = [ordered]@{
                error = $errors
                warning = $warnings
                verbose = $verbose
                information = $information
                debug = $debug
                progress = $progress
            }
        }
    }

    if ($information.Count -gt 0) {
        $entry.result.streams.informationText = ($information | ForEach-Object { $_.message }) -join "`n"
    }

    $ps.Dispose()

    [void]$session.History.Add($entry)

    Send-Json $context 200 ([ordered]@{
        ok = $success
        sessionId = $session.Id
        commandCount = $session.History.Count
        command = $entry
    })
}

function Handle-Root {
    param($context)

    Send-Json $context 200 ([ordered]@{
        ok = $true
        service = 'psBrowserPilot'
        endpoints = @(
            '/health',
            '/sessions',
            '/sessions/{id}',
            '/sessions/{id}/history',
            '/sessions/{id}/commands'
        )
    })
}

function Handle-Request {
    param($context)

    try {
        $request = $context.Request
        if ($request.HttpMethod -eq 'OPTIONS') {
            return Send-Json $context 204 $null
        }

        $path = $request.Url.AbsolutePath
        if ($path.Length -gt 1 -and $path.EndsWith('/')) {
            $path = $path.TrimEnd('/')
        }
        $segments = if ($path -eq '/') { @() } else { $path.Trim('/') -split '/' }

        switch -Regex ($request.HttpMethod) {
            'GET' {
                if ($segments.Count -eq 0) { return Handle-Root $context }
                if ($segments.Count -eq 1 -and $segments[0] -eq 'health') {
                    return Send-Json $context 200 (@{ ok = $true; status = 'healthy' })
                }
                if ($segments.Count -eq 1 -and $segments[0] -eq 'sessions') {
                    return Handle-SessionList $context
                }
                if ($segments.Count -eq 2 -and $segments[0] -eq 'sessions') {
                    return Handle-SessionGet $context $segments[1]
                }
                if ($segments.Count -eq 3 -and $segments[0] -eq 'sessions' -and $segments[2] -eq 'history') {
                    return Handle-SessionHistory $context $segments[1]
                }
                break
            }
            'POST' {
                if ($segments.Count -eq 1 -and $segments[0] -eq 'sessions') {
                    return Handle-SessionCreate $context
                }
                if ($segments.Count -eq 3 -and $segments[0] -eq 'sessions' -and $segments[2] -eq 'commands') {
                    return Handle-CommandInvoke $context $segments[1]
                }
                break
            }
            'DELETE' {
                if ($segments.Count -eq 2 -and $segments[0] -eq 'sessions') {
                    return Handle-SessionDelete $context $segments[1]
                }
                break
            }
        }

        Send-Json $context 404 ([ordered]@{
            ok = $false
            error = 'Endpoint not found'
            path = $path
        })
    } catch {
        Write-Warning "Request handling failed: $($_.Exception.Message)"
        try {
            Send-Json $context 500 ([ordered]@{
                ok = $false
                error = $_.Exception.Message
                detail = $_.ScriptStackTrace
            })
        } catch {
        }
    }
}

try {
    while ($listener.IsListening) {
        try {
            $context = $listener.GetContext()
        } catch [System.Net.HttpListenerException] {
            break
        } catch [System.ObjectDisposedException] {
            break
        }

        [System.Threading.ThreadPool]::QueueUserWorkItem({ param($ctx) Handle-Request $ctx }, $context) | Out-Null
    }
} finally {
    Write-Host 'Shutting down listener...'
    $listener.Stop()
    foreach ($entry in $global:Sessions.GetEnumerator()) {
        Close-Runspace $entry.Value
    }
    $listener.Close()
}
