param([string]$PfxPassword = $env:ORION_PFX_PASSWORD)
# Packages Dist\Orion-v3.1.exe as a signed MSIX. Requires certs\OrionSigning.pfx (kept out of git).
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$tools = Get-ChildItem "$PSScriptRoot\tools\bt\bin" -Recurse -Filter makeappx.exe | ? FullName -match 'x64' | Select -First 1
$bin = $tools.DirectoryName
$stage = "$PSScriptRoot\msix\stage"
Remove-Item $stage -Recurse -Force -ErrorAction SilentlyContinue
New-Item -ItemType Directory $stage | Out-Null
Copy-Item "$PSScriptRoot\msix\AppxManifest.xml" $stage
Copy-Item "$PSScriptRoot\msix\Assets" $stage -Recurse
Copy-Item "$root\Dist\Orion-v3.1.exe" $stage
$out = "$root\Dist\Orion-v3.1.msix"
Remove-Item $out -ErrorAction SilentlyContinue
& "$bin\makeappx.exe" pack /d $stage /p $out /o
& "$bin\signtool.exe" sign /fd SHA256 /f "$root\certs\OrionSigning.pfx" /p $PfxPassword $out
Copy-Item "$root\certs\Orion-Certificate.cer" "$root\Dist\Orion-Certificate.cer" -Force
