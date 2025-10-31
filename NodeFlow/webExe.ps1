Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

function Get-WebView2AssemblyInfo {
    param(
        [string[]]$HintDirectories = @()
    )

    $candidates = @()

    foreach ($hint in $HintDirectories) {
        if ($hint -and (Test-Path $hint)) {
            $candidates += (Get-Item $hint).FullName
        }
    }

    $programFilesRoots = @($env:ProgramFiles, ${env:ProgramFiles(x86)}) | Where-Object { $_ -and (Test-Path $_) }
    foreach ($root in $programFilesRoots) {
        foreach ($suffix in @('Microsoft\EdgeWebView\Application', 'Microsoft\Edge\Application')) {
            $path = Join-Path $root $suffix
            if (Test-Path $path) {
                $candidates += $path
            }
        }
    }

    $candidates = $candidates | Select-Object -Unique
    foreach ($candidate in $candidates) {
        $versionDirs = Get-ChildItem -Path $candidate -Directory -ErrorAction SilentlyContinue | Sort-Object Name -Descending
        foreach ($version in $versionDirs) {
            $winForms = Join-Path $version.FullName 'Microsoft.Web.WebView2.WinForms.dll'
            $core = Join-Path $version.FullName 'Microsoft.Web.WebView2.Core.dll'
            $loader = Join-Path $version.FullName 'WebView2Loader.dll'

            if ((Test-Path $winForms) -and (Test-Path $core)) {
                return [PSCustomObject]@{
                    WinForms       = (Get-Item $winForms).FullName
                    Core           = (Get-Item $core).FullName
                    Loader         = if (Test-Path $loader) { (Get-Item $loader).FullName } else { $null }
                    LoaderRootPath = $version.FullName
                }
            }
        }
    }

    return $null
}

$localWebViewFolder = Join-Path $PSScriptRoot 'WebView2'
$assemblyInfo = Get-WebView2AssemblyInfo -HintDirectories @($localWebViewFolder)

if (-not $assemblyInfo) {
    [System.Windows.Forms.MessageBox]::Show(
        'WebView2 ランタイムまたはライブラリが見つかりませんでした。Microsoft Edge WebView2 ランタイムをインストールしてください。',
        'NodeFlow',
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null
    return
}

if ($assemblyInfo.LoaderRootPath) {
    $loaderDirectory = $assemblyInfo.LoaderRootPath
    if ($loaderDirectory -and (Test-Path $loaderDirectory)) {
        $currentPath = [System.Environment]::GetEnvironmentVariable('PATH')
        if ($currentPath -notlike "*$loaderDirectory*") {
            [System.Environment]::SetEnvironmentVariable('PATH', "$loaderDirectory;$currentPath")
        }
    }
}

Add-Type -Path $assemblyInfo.WinForms
Add-Type -Path $assemblyInfo.Core

$csCode = @"
using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;

[ComVisible(true)]
[ClassInterface(ClassInterfaceType.AutoDual)]
public class NodeBridge
{
    private static string Escape(string value)
    {
        if (string.IsNullOrEmpty(value))
        {
            return string.Empty;
        }

        return value
            .Replace("\\", "\\\\")
            .Replace("\"", "\\\"")
            .Replace("\r", "\\r")
            .Replace("\n", "\\n");
    }

    public string RunPowerShell(string script)
    {
        if (string.IsNullOrWhiteSpace(script))
        {
            return "{\"exitCode\":-1,\"stdout\":\"\",\"stderr\":\"Script is empty.\"}";
        }

        string tempFile = Path.Combine(Path.GetTempPath(), "nodeflow_" + Guid.NewGuid().ToString("N") + ".ps1");

        try
        {
            File.WriteAllText(tempFile, script ?? string.Empty);

            var psi = new ProcessStartInfo
            {
                FileName = "powershell.exe",
                Arguments = "-NoProfile -ExecutionPolicy Bypass -File \"" + tempFile + "\"",
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };

            using (var process = Process.Start(psi))
            {
                if (process == null)
                {
                    return "{\"exitCode\":-1,\"stdout\":\"\",\"stderr\":\"Failed to start PowerShell process.\"}";
                }

                string stdout = process.StandardOutput.ReadToEnd();
                string stderr = process.StandardError.ReadToEnd();
                process.WaitForExit();

                return "{\"exitCode\":" + process.ExitCode + ",\"stdout\":\"" + Escape(stdout) + "\",\"stderr\":\"" + Escape(stderr) + "\"}";
            }
        }
        catch (Exception ex)
        {
            return "{\"exitCode\":-1,\"stdout\":\"\",\"stderr\":\"" + Escape(ex.Message) + "\"}";
        }
        finally
        {
            try
            {
                if (File.Exists(tempFile))
                {
                    File.Delete(tempFile);
                }
            }
            catch
            {
                // ignore cleanup errors
            }
        }
    }
}
"@

Add-Type -TypeDefinition $csCode -ReferencedAssemblies System.Windows.Forms, System.Drawing, System.Core

[System.Windows.Forms.Application]::EnableVisualStyles()
[System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false)

$indexPath = Join-Path $PSScriptRoot 'index.html'
if (-not (Test-Path $indexPath)) {
    [System.Windows.Forms.MessageBox]::Show(
        "index.html が見つかりません: $indexPath",
        'NodeFlow',
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null
    return
}

$resolvedIndexPath = (Resolve-Path $indexPath).ProviderPath

$form = New-Object System.Windows.Forms.Form
$form.Text = 'NodeFlow'
$form.Size = New-Object System.Drawing.Size(1200, 800)
$form.StartPosition = 'CenterScreen'

$webView = New-Object Microsoft.Web.WebView2.WinForms.WebView2
$webView.Dock = 'Fill'
$webView.DefaultBackgroundColor = [System.Drawing.Color]::FromArgb(255, 24, 24, 24)
$form.Controls.Add($webView)

$bridge = New-Object NodeBridge

try {
    $userDataFolder = Join-Path ([System.IO.Path]::GetTempPath()) 'NodeFlowWebView2'
    if (-not (Test-Path $userDataFolder)) {
        [System.IO.Directory]::CreateDirectory($userDataFolder) | Out-Null
    }

    $environment = [Microsoft.Web.WebView2.Core.CoreWebView2Environment]::CreateAsync($null, $userDataFolder).GetAwaiter().GetResult()
    $null = $webView.EnsureCoreWebView2Async($environment).GetAwaiter().GetResult()

    $webView.CoreWebView2.AddHostObjectToScript('nodeBridge', $bridge)
    $webView.CoreWebView2.Settings.AreDefaultScriptDialogsEnabled = $true
    $webView.CoreWebView2.Settings.AreDevToolsEnabled = $false

    $webView.Source = New-Object System.Uri($resolvedIndexPath)
}
catch {
    [System.Windows.Forms.MessageBox]::Show(
        "WebView2 の初期化に失敗しました: $($_.Exception.Message)",
        'NodeFlow',
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null
    $webView.Dispose()
    $form.Dispose()
    return
}

$form.Add_Shown({ param($sender, $args) $sender.Activate() })
$form.Add_FormClosed({
    param($sender, $args)
    try {
        if ($webView.CoreWebView2) {
            $webView.CoreWebView2.RemoveHostObjectFromScript('nodeBridge')
        }
    } catch {
        # ignore teardown failures
    }
    $webView.Dispose()
})

$form.ShowDialog() | Out-Null
$form.Dispose()
