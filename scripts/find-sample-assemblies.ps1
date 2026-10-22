# 샘플 폴더에서 "너무 복잡하지 않으면서 구멍이 많은 판재·블록형 조립품" 후보를 찾는다 · 읽기 전용
# 사용: .\scripts\find-sample-assemblies.ps1 -Root "C:\Users\Public\Documents\Autodesk\Inventor 2022\Samples"
#
# - 샘플 파일을 열거나(Inventor) 수정·이동하지 않는다. .iam 바이트를 읽기만 한다
# - 참조 파일명 수는 추정치다. .iam 바이트에서 UTF-16LE·ASCII 문자열을 뽑아 ".ipt"/".iam"으로 끝나는
#   고유 파일명을 센 값이라, 실제 참조 부품 수와 다를 수 있다 (CSV 열 "참조파일명수_추정")
# - 구멍 수는 PowerShell로 알 수 없으므로 계산하지 않는다
param(
    [Parameter(Mandatory)][string]$Root,
    [int]$MinParts = 6,
    [int]$MaxParts = 60,
    [string]$OutCsv = (Join-Path ([System.Environment]::GetFolderPath("DesktopDirectory")) ("HoleCheck_SampleScan_" + (Get-Date -Format "yyyyMMdd_HHmmss") + ".csv")),
    [int]$MaxMB = 200
)

$colKeywords = @("plate", "bracket", "flange", "frame", "panel", "block", "base", "pattern", "perforat", "grid", "hole", "bolt")

# 파일명 한 조각 · 경로 구분자는 포함하지 않으므로 전체 경로에서도 마지막 이름만 잡힌다
$strFileNamePattern = '[A-Za-z0-9_\-\.\(\)\[\] ]{1,150}?\.(?:ipt|iam)\b'

# 같은 폴더의 .ipt/.iam 개수는 폴더당 한 번만 센다
$colFolderCounts = @{}

function Get-FolderCounts([string]$strFolder) {
    if (-not $colFolderCounts.ContainsKey($strFolder)) {
        $intIptCount = @(Get-ChildItem -LiteralPath $strFolder -Filter "*.ipt" -File -ErrorAction SilentlyContinue).Count
        $intIamCount = @(Get-ChildItem -LiteralPath $strFolder -Filter "*.iam" -File -ErrorAction SilentlyContinue).Count
        $colFolderCounts[$strFolder] = @($intIptCount, $intIamCount)
    }
    return $colFolderCounts[$strFolder]
}

function Get-ReferencedNameCount([byte[]]$colBytes, [string]$strSelfName) {
    $colNames = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    $objRegex = New-Object System.Text.RegularExpressions.Regex($strFileNamePattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)

    # UTF-16LE는 짝수·홀수 바이트 시작 두 번 읽는다 (문자열이 홀수 오프셋에 있을 수 있음)
    $colTexts = @(
        [System.Text.Encoding]::Unicode.GetString($colBytes),
        [System.Text.Encoding]::ASCII.GetString($colBytes)
    )
    if ($colBytes.Length -gt 1) {
        $colTexts += [System.Text.Encoding]::Unicode.GetString($colBytes, 1, $colBytes.Length - 1)
    }

    foreach ($strText in $colTexts) {
        foreach ($objMatch in $objRegex.Matches($strText)) {
            $strName = $objMatch.Value.Trim()
            if ($strName.Length -gt 4 -and $strName -ne $strSelfName) {
                [void]$colNames.Add($strName)
            }
        }
    }
    return $colNames.Count
}

function Get-KeywordMatches([string]$strFileName, [string]$strFolderName) {
    $strTarget = ($strFileName + " " + $strFolderName).ToLowerInvariant()
    return @($colKeywords | Where-Object { $strTarget.Contains($_) })
}

if (-not (Test-Path -LiteralPath $Root -PathType Container)) {
    Write-Host "폴더가 없습니다: $Root"
    exit 1
}

$colAssemblies = @(Get-ChildItem -LiteralPath $Root -Filter "*.iam" -File -Recurse -ErrorAction SilentlyContinue)
$colRows = New-Object System.Collections.Generic.List[object]
$intSkippedCount = 0

foreach ($objFile in $colAssemblies) {
    $dblSizeMB = [System.Math]::Round($objFile.Length / 1MB, 2)
    if ($dblSizeMB -gt $MaxMB) {
        Write-Warning "건너뜀 (${dblSizeMB}MB > ${MaxMB}MB): $($objFile.FullName)"
        $intSkippedCount++
        continue
    }

    $strFolder = $objFile.DirectoryName
    $strFolderName = Split-Path -Leaf $strFolder
    $colCounts = Get-FolderCounts $strFolder
    $colMatched = Get-KeywordMatches $objFile.Name $strFolderName

    $objRow = [ordered]@{
        "전체경로"           = $objFile.FullName
        "상위폴더"           = $strFolderName
        "파일명"             = $objFile.Name
        "크기MB"             = $dblSizeMB
        "같은폴더_ipt수"     = $colCounts[0]
        "같은폴더_iam수"     = $colCounts[1]
        "참조파일명수_추정"  = ""
        "키워드일치수"       = $colMatched.Count
        "키워드"             = ($colMatched -join " ")
        "판정"               = ""
        "오류"               = ""
    }

    try {
        $colBytes = [System.IO.File]::ReadAllBytes($objFile.FullName)
        $intReferencedCount = Get-ReferencedNameCount $colBytes $objFile.Name
        $objRow["참조파일명수_추정"] = $intReferencedCount
        if ($intReferencedCount -ge $MinParts -and $intReferencedCount -le $MaxParts) {
            $objRow["판정"] = "후보"
        } else {
            $objRow["판정"] = "범위밖"
        }
    } catch {
        $objRow["판정"] = "읽기실패"
        $objRow["오류"] = $_.Exception.Message
    }

    $colRows.Add([pscustomobject]$objRow)
}

# 후보를 위로 · 그 안에서 키워드 일치 수 많은 순
$colSorted = @($colRows | Sort-Object -Property `
    @{ Expression = { if ($_."판정" -eq "후보") { 0 } else { 1 } }; Ascending = $true },
    @{ Expression = { $_."키워드일치수" }; Descending = $true })

# Windows PowerShell 5.1의 -Encoding UTF8은 BOM을 붙인다
$colSorted | Export-Csv -LiteralPath $OutCsv -NoTypeInformation -Encoding UTF8

$colCandidates = @($colSorted | Where-Object { $_."판정" -eq "후보" })

Write-Host ""
Write-Host "전체 .iam: $($colAssemblies.Count) · 후보: $($colCandidates.Count) · 건너뜀(크기): $intSkippedCount"
Write-Host "상위 10개 (경로 · 참조파일명수_추정 · 키워드)"
foreach ($objRow in ($colCandidates | Select-Object -First 10)) {
    Write-Host ("  {0} · {1} · {2}" -f $objRow."전체경로", $objRow."참조파일명수_추정", $objRow."키워드")
}
Write-Host "CSV → $OutCsv"
