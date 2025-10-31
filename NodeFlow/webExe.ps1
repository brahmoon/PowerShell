<#
 NodeFlow WebView2 Host - Async Stable Edition
#>

# --- STA強制 ---
if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    $psi = New-Object System.Diagnostics.ProcessStartInfo "powershell.exe"
    $psi.Arguments = "-STA -NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""
    [System.Diagnostics.Process]::Start($psi) | Out-Null
    exit
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()
[System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false)

$LogFile = Join-Path $PSScriptRoot "NodeFlow.log"
"[$(Get-Date)] NodeFlow starting..." | Out-File $LogFile -Encoding utf8

# --- DLLロード ---
$winFormsDll = Join-Path $PSScriptRoot 'Microsoft.Web.WebView2.WinForms.dll'
$coreDll     = Join-Path $PSScriptRoot 'Microsoft.Web.WebView2.Core.dll'
$loaderDll   = Join-Path $PSScriptRoot 'WebView2Loader.dll'
$env:PATH = "$($PSScriptRoot);$env:PATH"

Add-Type -Path $winFormsDll
Add-Type -Path $coreDll

# --- Bridgeクラス ---
$csCode = @"
using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;

[ComVisible(true)]
[ClassInterface(ClassInterfaceType.AutoDual)]
public class NodeBridge {
    public string RunPowerShell(string script) {
        try {
            var psi = new ProcessStartInfo();
            psi.FileName = "powershell.exe";
            psi.Arguments = "-NoProfile -Command " + script;
            psi.RedirectStandardOutput = true;
            psi.UseShellExecute = false;
            psi.CreateNoWindow = true;
            using (var proc = Process.Start(psi)) {
                string output = proc.StandardOutput.ReadToEnd();
                proc.WaitForExit();
                return output;
            }
        } catch (Exception ex) {
            return "Error: " + ex.Message;
        }
    }
}
"@
Add-Type -TypeDefinition $csCode -ReferencedAssemblies System.Windows.Forms, System.Drawing, System.Core

# --- index.html ---
$indexPath = Join-Path $PSScriptRoot 'index.html'
if (-not (Test-Path $indexPath)) {
    [System.Windows.Forms.MessageBox]::Show("index.html not found.", "NodeFlow") | Out-Null
    exit
}
$resolvedIndexPath = (Resolve-Path $indexPath).ProviderPath
$resolvedUri = "file:///" + $resolvedIndexPath.Replace("\","/")

# --- フォーム構築 ---
$form = New-Object System.Windows.Forms.Form
$form.Text = "NodeFlow"
$form.Size = New-Object System.Drawing.Size(1200,800)
$form.StartPosition = "CenterScreen"

$webView = New-Object Microsoft.Web.WebView2.WinForms.WebView2
$webView.Dock = "Fill"
$form.Controls.Add($webView)

$bridge = New-Object NodeBridge

# --- 初期化ロジック ---
$form.Add_Shown({
    try {
        "[$(Get-Date)] Initializing WebView2..." | Out-File $LogFile -Append

        # 環境オブジェクト作成
        $userDataFolder = Join-Path ([System.IO.Path]::GetTempPath()) 'NodeFlowWebView2'
        if (-not (Test-Path $userDataFolder)) { [System.IO.Directory]::CreateDirectory($userDataFolder) | Out-Null }

        $envTask = [Microsoft.Web.WebView2.Core.CoreWebView2Environment]::CreateAsync($null, $userDataFolder)
        $envTask.ContinueWith({
            param($t)
            $form.Invoke({
                try {
                    $env = $t.Result
                    $initTask = $webView.EnsureCoreWebView2Async($env)
                    $initTask.ContinueWith({
                        param($x)
                        $form.Invoke({
                            try {
                                $webView.CoreWebView2.AddHostObjectToScript('nodeBridge', $bridge)
                                $webView.CoreWebView2.Settings.AreDevToolsEnabled = $true
                                $webView.Source = [System.Uri]$resolvedUri
                                "[$(Get-Date)] WebView2 initialized OK -> $resolvedUri" | Out-File $LogFile -Append
                            } catch {
                                "[$(Get-Date)] Error in nested init: $($_.Exception.Message)" | Out-File $LogFile -Append
                            }
                        })
                    })
                } catch {
                    "[$(Get-Date)] Env init error: $($_.Exception.Message)" | Out-File $LogFile -Append
                }
            })
        })
    } catch {
        "[$(Get-Date)] Unexpected init exception: $($_.Exception.Message)" | Out-File $LogFile -Append
    }
})

$form.Add_FormClosed({ "[$(Get-Date)] NodeFlow closed." | Out-File $LogFile -Append })
[System.Windows.Forms.Application]::Run($form)
