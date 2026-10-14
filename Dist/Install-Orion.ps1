# Orion one-click installer: trusts the Orion certificate (Trusted People) and installs the app.
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
  Start-Process powershell -Verb RunAs -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$($MyInvocation.MyCommand.Path)`""
  exit
}
try {
  $cer = Join-Path $here 'Orion-Certificate.cer'
  $msix = Join-Path $here 'Orion-v2.1.msix'
  if (-not (Test-Path $cer) -or -not (Test-Path $msix)) { throw "Orion-Certificate.cer and Orion-v2.1.msix must be in the same folder as this installer." }
  Write-Host "Installing Orion certificate (Trusted People)..."
  Import-Certificate -FilePath $cer -CertStoreLocation Cert:\LocalMachine\TrustedPeople | Out-Null
  Write-Host "Installing Orion..."
  Add-AppxPackage -Path $msix
  Write-Host "`nDone. Open 'Orion Core Structural Calculator' from the Start menu." -ForegroundColor Green
} catch {
  Write-Host "`nInstallation failed: $($_.Exception.Message)" -ForegroundColor Red
}
Read-Host "`nPress Enter to close"
