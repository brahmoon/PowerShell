# Legacy entry point retained for backward compatibility.
# This script now delegates to the modular dockable form launcher.

$scriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path
$modularLauncher = Join-Path $scriptDirectory 'ModulerForm/DockableModuleApp.ps1'

if (-not (Test-Path -Path $modularLauncher))
{
    throw "Modular launcher script not found at '$modularLauncher'."
}

& $modularLauncher @PSBoundParameters
