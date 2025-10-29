param(
    [string]$ModulesPath = (Join-Path $PSScriptRoot 'Modules')
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$corePath = Join-Path $PSScriptRoot 'Core'
$coreFiles = Get-ChildItem -Path $corePath -Filter '*.cs' | Sort-Object Name

if (-not $coreFiles)
{
    throw "Core definitions are missing in '$corePath'."
}

$references = @(
    'System.dll',
    'System.Core.dll',
    'System.Drawing.dll',
    'System.Windows.Forms.dll',
    'System.Web.Extensions.dll',
    'Microsoft.CSharp.dll'
)

Add-Type -Path $coreFiles.FullName -ReferencedAssemblies $references -Language CSharp

if (-not (Test-Path -Path $ModulesPath))
{
    New-Item -ItemType Directory -Path $ModulesPath | Out-Null
}

[System.Windows.Forms.Application]::EnableVisualStyles()
[System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false)

$form = [DockableModularForm.MainForm]::new($ModulesPath)
[System.Windows.Forms.Application]::Run($form)
