Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$csCode = @"
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;

[ComVisible(true)]
public class NodeBridge
{
    public void RunScript(string script)
    {
        if (string.IsNullOrWhiteSpace(script))
        {
            MessageBox.Show("生成されたスクリプトが空です。", "NodeFlow", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        try
        {
            var psi = new ProcessStartInfo
            {
                FileName = "powershell.exe",
                Arguments = "-NoLogo -NoProfile -ExecutionPolicy Bypass -Command -",
                RedirectStandardInput = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };

            using (var process = Process.Start(psi))
            {
                if (process == null)
                {
                    MessageBox.Show("PowerShell プロセスを起動できませんでした。", "NodeFlow", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                process.StandardInput.WriteLine(script);
                process.StandardInput.WriteLine("exit $LASTEXITCODE");
                process.StandardInput.Flush();
                process.StandardInput.Close();

                string output = process.StandardOutput.ReadToEnd();
                string error = process.StandardError.ReadToEnd();
                process.WaitForExit();

                if (!string.IsNullOrWhiteSpace(error))
                {
                    MessageBox.Show(error.Trim(), "PowerShell 実行エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    string message = string.IsNullOrWhiteSpace(output)
                        ? "PowerShell スクリプトを実行しました。"
                        : output.Trim();
                    MessageBox.Show(message, "PowerShell 実行結果", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                "PowerShell スクリプトの実行に失敗しました: " + ex.Message,
                "NodeFlow",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }
    }
}
"@;

Add-Type -TypeDefinition $csCode -ReferencedAssemblies System.Windows.Forms, System.Drawing, System.Core

[System.Windows.Forms.Application]::EnableVisualStyles()
[System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false)

$form = New-Object System.Windows.Forms.Form
$form.Text = "NodeFlow Editor"
$form.Size = New-Object System.Drawing.Size(1200, 800)
$form.StartPosition = "CenterScreen"

$browser = New-Object System.Windows.Forms.WebBrowser
$browser.Dock = "Fill"
$browser.ScriptErrorsSuppressed = $true
$browser.ObjectForScripting = (New-Object NodeBridge)

$indexPath = Join-Path $PSScriptRoot 'index.html'
if (-not (Test-Path $indexPath)) {
    [System.Windows.Forms.MessageBox]::Show(
        "index.html が見つかりません: $indexPath",
        "NodeFlow",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error)
    return
}

$browser.Url = New-Object System.Uri($indexPath)
$form.Controls.Add($browser)

$form.ShowDialog() | Out-Null
