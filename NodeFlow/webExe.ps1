Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# --- 子プロセス呼び出し用クラス（Bridge） ---
$csCode = @"
using System;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;

[ComVisible(true)]
public class NodeBridge {
    public void Execute(string json) {
        try {
            // JSONを一時ファイルに保存
            string temp = Path.Combine(Path.GetTempPath(), "node_" + Guid.NewGuid().ToString("N") + ".json");
            File.WriteAllText(temp, json);

            // 子プロセスPowerShellにスクリプト文字列を直接渡す
            string script = @"
                try {
                    \$json = Get-Content '$temp' -Raw | ConvertFrom-Json
                    switch (\$json.op) {
                        'add' { \$r = \$json.a + \$json.b }
                        'sub' { \$r = \$json.a - \$json.b }
                        'mul' { \$r = \$json.a * \$json.b }
                        'div' { 
                            if (\$json.b -eq 0) { throw 'Division by zero' }
                            \$r = \$json.a / \$json.b 
                        }
                        default { throw ('Unknown operation: ' + \$json.op) }
                    }
                    [pscustomobject]@{status='ok';op=\$json.op;result=\$r} | ConvertTo-Json -Compress
                } catch {
                    [pscustomobject]@{status='error';message=\$_.Exception.Message} | ConvertTo-Json -Compress
                }
            ";

            var psi = new ProcessStartInfo();
            psi.FileName = "powershell.exe";
            psi.Arguments = "-NoProfile -ExecutionPolicy Bypass -Command \"" + script.Replace("\"","\\\"") + "\"";
            psi.RedirectStandardOutput = true;
            psi.RedirectStandardError = true;
            psi.UseShellExecute = false;
            psi.CreateNoWindow = true;

            using (var proc = Process.Start(psi)) {
                string output = proc.StandardOutput.ReadToEnd();
                string error = proc.StandardError.ReadToEnd();
                proc.WaitForExit();

                try { File.Delete(temp); } catch {}

                if (!string.IsNullOrWhiteSpace(error)) {
                    MessageBox.Show("Error: " + error, "Executor Error");
                } else {
                    MessageBox.Show("Result: " + output, "Executor Result");
                }
            }
        } catch (Exception ex) {
            MessageBox.Show("Execution failed: " + ex.Message, "Bridge Error");
        }
    }
}
"@

# コンパイル
Add-Type -TypeDefinition $csCode -ReferencedAssemblies System.Windows.Forms, System.Drawing, System.ComponentModel, System.Core, System.Drawing.Design

# --- フォーム構築 ---
$form = New-Object System.Windows.Forms.Form
$form.Text = "Node Editor Host"
$form.Size = New-Object System.Drawing.Size(900,600)
$form.StartPosition = "CenterScreen"

$wb = New-Object System.Windows.Forms.WebBrowser
$wb.Dock = "Fill"
$wb.ScriptErrorsSuppressed = $true
$wb.ObjectForScripting = (New-Object NodeBridge)

# --- HTML ---
$html = @"
<!DOCTYPE html>
<html>
<head>
<meta charset='utf-8'>
<title>Node Editor Demo</title>
<style>
body { font-family: sans-serif; margin: 20px; }
textarea { width: 100%; height: 100px; }
button { padding: 6px 12px; margin-top: 8px; }
</style>
</head>
<body>
<h2>Node Execution (Direct Child Process)</h2>
<textarea id='nodejson'>{ "op":"mul", "a":3, "b":4 }</textarea><br>
<button onclick='runNode()'>Execute Node</button>
<div id='log' style='margin-top:12px;border:1px solid #ccc;padding:8px;'></div>
<script>
function log(msg){
  document.getElementById('log').innerHTML += msg + '<br>';
}
function runNode(){
  var node = document.getElementById('nodejson').value;
  log('Executing: ' + node);
  if(window.external && window.external.Execute){
    window.external.Execute(node);
  } else {
    log('Bridge not available.');
  }
}
</script>
</body>
</html>
"@

$wb.DocumentText = $html
$form.Controls.Add($wb)
$form.ShowDialog()
