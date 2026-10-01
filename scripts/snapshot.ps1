# 사용: .\snapshot.ps1 -Week 04 -State "줌 배율 미작동 · 카운터보어 축간거리 0"
param([Parameter(Mandatory)][string]$Week, [Parameter(Mandatory)][string]$State)
$root = Join-Path $PSScriptRoot ".."
$name = "W$Week`_$State.zip"
$out  = Join-Path $root "snapshots\$name"
$tmp  = Join-Path $env:TEMP "hc_snap"
Remove-Item $tmp -Recurse -Force -ErrorAction SilentlyContinue
New-Item -ItemType Directory -Force -Path $tmp | Out-Null
Copy-Item (Join-Path $root "src")  $tmp -Recurse -Exclude bin,obj,.vs
Copy-Item (Join-Path $root "docs") $tmp -Recurse
Compress-Archive -Path "$tmp\*" -DestinationPath $out -Force
Write-Host "스냅샷 → $out"
