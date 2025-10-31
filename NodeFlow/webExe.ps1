Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$csCode = @"
using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

[ComVisible(true)]
public class NodeBridge
{
    private readonly string workingDirectory;

    public NodeBridge(string workingDirectory)
    {
        this.workingDirectory = workingDirectory;
    }

    public void RunScript(string script)
    {
        if (string.IsNullOrWhiteSpace(script))
        {
            MessageBox.Show("生成されたスクリプトが空です。", "NodeFlow Executor", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        try
        {
            var psi = new ProcessStartInfo
            {
                FileName = "powershell.exe",
                Arguments = "-NoProfile -ExecutionPolicy Bypass -Command -",
                RedirectStandardInput = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true,
                WorkingDirectory = workingDirectory
            };

            using (var process = Process.Start(psi))
            {
                if (process == null)
                {
                    MessageBox.Show("PowerShell プロセスを開始できませんでした。", "NodeFlow Executor", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                process.StandardInput.WriteLine(script);
                process.StandardInput.WriteLine("exit");
                process.StandardInput.Flush();
                process.StandardInput.Close();

                string output = process.StandardOutput.ReadToEnd();
                string error = process.StandardError.ReadToEnd();
                process.WaitForExit();

                var builder = new StringBuilder();
                if (!string.IsNullOrWhiteSpace(output))
                {
                    builder.AppendLine("[Output]");
                    builder.AppendLine(output.Trim());
                }

                if (!string.IsNullOrWhiteSpace(error))
                {
                    if (builder.Length > 0)
                    {
                        builder.AppendLine();
                    }
                    builder.AppendLine("[Error]");
                    builder.AppendLine(error.Trim());
                }

                var message = builder.Length > 0 ? builder.ToString() : "PowerShell スクリプトが完了しました。";
                MessageBox.Show(message, "NodeFlow Executor", MessageBoxButtons.OK, string.IsNullOrWhiteSpace(error) ? MessageBoxIcon.Information : MessageBoxIcon.Warning);
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show("スクリプトの実行に失敗しました: " + ex.Message, "NodeFlow Executor", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}
"@;

Add-Type -TypeDefinition $csCode -ReferencedAssemblies System.Windows.Forms, System.Drawing

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$indexPath = Join-Path $scriptDir 'index.html'

$form = New-Object System.Windows.Forms.Form
$form.Text = 'NodeFlow'
$form.Size = New-Object System.Drawing.Size(1200, 800)
$form.StartPosition = 'CenterScreen'

$browser = New-Object System.Windows.Forms.WebBrowser
$browser.Dock = 'Fill'
$browser.ScriptErrorsSuppressed = $false
$browser.ObjectForScripting = (New-Object NodeBridge ($scriptDir))
$form.Controls.Add($browser)

$uri = (New-Object System.Uri($indexPath)).AbsoluteUri
$browser.Navigate($uri)

$form.Add_Shown({ $form.Activate() })
[void]$form.ShowDialog()
