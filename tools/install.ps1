# Installs the plugin into ONLYOFFICE, and gives it the secret it needs to be trusted as local.
#
#   pwsh tools/install.ps1                 # install for the current user
#   pwsh tools/install.ps1 -Uninstall
#
# Two things happen here, and the second is the one that is easy to forget:
#
#   1. The plugin's files are copied into ONLYOFFICE's user plugin folder, named by the guid in
#      config.json. The USER folder, not the one under Program Files: that one needs elevation,
#      and the user folder has its own `v1`, so the `../v1/plugins.js` the page loads still
#      resolves.
#   2. `secret.js` is written from the per-install secret the Nexus PLM Addin Service keeps on
#      this machine. The plugin's page is loaded from disk, so it has an opaque origin, which the
#      service will not trust on its own - any website can obtain one through a sandboxed iframe.
#      The secret is the proof that this caller is really on this machine. A page loaded from disk
#      can read a file beside itself; a remote site cannot read the user's disk.
#
# Without step 2 the plugin installs and draws, and every call answers 403.

param(
    [switch]$Uninstall,
    [string]$PluginSource = (Join-Path $PSScriptRoot ".." "plugin"),
    [string]$SecretFile = (Join-Path $env:LOCALAPPDATA "NexusPLM\browser-addin-secret.txt")
)

$ErrorActionPreference = 'Stop'

$pluginsRoot = Join-Path $env:LOCALAPPDATA "ONLYOFFICE\DesktopEditors\data\sdkjs-plugins"
if (-not (Test-Path $pluginsRoot)) {
    throw "ONLYOFFICE's plugin folder was not found at $pluginsRoot. Is Desktop Editors installed and has it been run once?"
}

# The folder is named by the guid, without the "asc." prefix config.json carries.
$config = Get-Content (Join-Path $PluginSource "config.json") -Raw | ConvertFrom-Json
$folderName = $config.guid -replace '^asc\.', ''
$dest = Join-Path $pluginsRoot $folderName

if ($Uninstall) {
    if (Test-Path $dest) { Remove-Item $dest -Recurse -Force; Write-Output "removed $dest" }
    else { Write-Output "nothing installed at $dest" }
    return
}

if (Test-Path $dest) { Remove-Item $dest -Recurse -Force }
New-Item -ItemType Directory -Path $dest -Force | Out-Null
Copy-Item (Join-Path $PluginSource "*") $dest -Recurse -Force
Write-Output "installed $($config.name) $($config.version) to $dest"

if (Test-Path $SecretFile) {
    $secret = (Get-Content $SecretFile -Raw).Trim()
    # Written as a script, not as data: a browser refuses `fetch` from a file:// page to a file://
    # URL, but will load a sibling <script> quite happily.
    $js = @"
/* Written by tools/install.ps1 from the Addin Service's per-install secret. Not source, not
   secret in any cryptographic sense - it never leaves this machine, and it exists so the service
   can tell a local plugin from a website that forged an opaque origin. */
window.NexusPlmSecret = "$secret";
"@
    Set-Content -Path (Join-Path $dest "secret.js") -Value $js -Encoding UTF8
    Write-Output "wrote secret.js from $SecretFile"
} else {
    Write-Warning @"
No secret at $SecretFile, so secret.js was NOT written and every call will answer 403.
The Nexus PLM Addins tray application writes it at startup - start it once, then run this again.
"@
}

Write-Output ""
Write-Output "Restart ONLYOFFICE, then: Plugins -> Nexus PLM"
