Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$csCode = @"
using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;

[ComVisible(true)]
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

$form = New-Object System.Windows.Forms.Form
$form.Text = "NodeFlow"
$form.Size = New-Object System.Drawing.Size(1200, 800)
$form.StartPosition = "CenterScreen"

$browser = New-Object System.Windows.Forms.WebBrowser
$browser.Dock = "Fill"
$browser.ScriptErrorsSuppressed = $true
$browser.AllowWebBrowserDrop = $false
$browser.IsWebBrowserContextMenuEnabled = $false
$browser.WebBrowserShortcutsEnabled = $true
$browser.ObjectForScripting = (New-Object NodeBridge)

$indexPath = Join-Path $PSScriptRoot 'index.html'
if (-not (Test-Path $indexPath)) {
    [System.Windows.Forms.MessageBox]::Show("index.html が見つかりません: $indexPath", "NodeFlow", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    return
}

$uri = New-Object System.Uri((Resolve-Path $indexPath).ProviderPath)
$browser.Url = $uri

$form.Controls.Add($browser)
$form.Add_Shown({ $form.Activate() })

$form.ShowDialog() | Out-Null
$form.Dispose()
