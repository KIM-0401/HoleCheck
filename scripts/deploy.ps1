# 빌드 산출물을 Inventor Addins 폴더로 복사 · Inventor 종료 후 실행
$src = Join-Path $PSScriptRoot "..\src\HoleCheckAddIn\bin\Debug"
$dst = Join-Path $env:APPDATA "Autodesk\Inventor 2027\Addins\HoleCheckAddIn"
if (Get-Process Inventor -ErrorAction SilentlyContinue) { Write-Host "Inventor 실행 중. 종료 후 다시 실행." ; exit 1 }
New-Item -ItemType Directory -Force -Path $dst | Out-Null
Copy-Item "$src\*.dll" $dst -Force
Copy-Item "$src\*.addin" $dst -Force
Write-Host "배포 완료 → $dst"
Get-ChildItem $dst | Select-Object Name, LastWriteTime
