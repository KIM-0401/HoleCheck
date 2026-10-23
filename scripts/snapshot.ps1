# 사용: .\snapshot.ps1 -Week 04 -State "줌 배율 미작동 · 카운터보어 축간거리 0"
param([Parameter(Mandatory)][string]$Week, [Parameter(Mandatory)][string]$State)
$root = Join-Path $PSScriptRoot ".."
$name = "W$Week`_$State.zip"
$out  = Join-Path $root "snapshots\$name"
$tmp  = Join-Path $env:TEMP "hc_snap"
Remove-Item $tmp -Recurse -Force -ErrorAction SilentlyContinue
New-Item -ItemType Directory -Force -Path $tmp | Out-Null
# Copy-Item -Exclude 는 src 바로 아래 항목에만 적용되어 하위의 bin · obj · .vs 가 들어갔다 (W05 issues)
# 경로 어디에든 bin · obj · .vs 폴더가 있는 파일은 복사하지 않는다 (VS 가 열어 둔 .vs 잠금 파일도 건드리지 않음)
$srcRoot = (Resolve-Path (Join-Path $root "src")).Path
Get-ChildItem $srcRoot -Recurse -File -Force |
    Where-Object { $_.FullName.Substring($srcRoot.Length) -notmatch '(^|\\)(bin|obj|\.vs)\\' } |
    ForEach-Object {
        $dest = Join-Path (Join-Path $tmp "src") $_.FullName.Substring($srcRoot.Length + 1)
        New-Item -ItemType Directory -Force -Path (Split-Path $dest) | Out-Null
        Copy-Item -LiteralPath $_.FullName -Destination $dest -Force
    }
Copy-Item (Join-Path $root "docs") $tmp -Recurse
Compress-Archive -Path "$tmp\*" -DestinationPath $out -Force
Write-Host "스냅샷 → $out"
