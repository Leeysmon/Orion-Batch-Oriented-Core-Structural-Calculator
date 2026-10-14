$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
Set-Location $root
$builds = @(
  @{ name = 'Orion-v3.0'; cfg = "window.ORION_CONFIG = { version: '3.0', year: 2026, expires: '2026-12-31' };" },
  @{ name = 'Orion-v3.0_admin'; cfg = "window.ORION_CONFIG = { version: '3.0', year: 2026, expires: null };" }
)
foreach ($b in $builds) {
  $stage = Join-Path $root "build_exe\stage_$($b.name)"
  Remove-Item $stage -Recurse -Force -ErrorAction SilentlyContinue
  $app = Join-Path $stage 'app'
  New-Item -ItemType Directory -Force $app | Out-Null
  foreach ($f in 'index.html','style.css','app.js','calc.js','lock.js','manifest.webmanifest') { Copy-Item $f $app }
  Copy-Item vendor, icons -Destination $app -Recurse
  Set-Content (Join-Path $app 'config.js') $b.cfg
  $env:ORION_NAME = $b.name
  python -m PyInstaller --noconfirm --onefile --noconsole --name $b.name --icon "$root\icons\orion.ico" `
    --add-data "$app;app" --distpath "$root\Dist" --workpath "$root\build_exe\work_$($b.name)" --specpath "$stage" "$root\build_exe\launcher.py"
}
