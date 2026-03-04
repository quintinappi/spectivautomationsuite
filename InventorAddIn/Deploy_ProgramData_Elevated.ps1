$ErrorActionPreference = 'Stop'

$srcDll = 'C:\Users\Quintin\Documents\Spectiv\3. Working\INVENTOR_AUTOMATION_SUITE_2026\InventorAddIn\AssemblyClonerAddIn\bin\x64\Release\AssemblyClonerAddIn.dll'
$srcAddin = 'C:\Users\Quintin\Documents\Spectiv\3. Working\INVENTOR_AUTOMATION_SUITE_2026\InventorAddIn\AssemblyClonerAddIn\AssemblyClonerAddIn.addin'
$dstDir = 'C:\ProgramData\Autodesk\Inventor 2026\Addins'
$roamDir2026 = Join-Path $env:APPDATA 'Autodesk\Inventor 2026\Addins'
$roamDirLegacy = Join-Path $env:APPDATA 'Autodesk\Inventor Addins'
$report = 'C:\Users\Quintin\Documents\Spectiv\3. Working\INVENTOR_AUTOMATION_SUITE_2026\InventorAddIn\ProgramData_Deploy_Report.txt'

try {
    New-Item -ItemType Directory -Path $dstDir -Force | Out-Null

    $dstDll = Join-Path $dstDir 'AssemblyClonerAddIn.dll'
    $dstAddin = Join-Path $dstDir 'AssemblyClonerAddIn.addin'

    if (Test-Path $dstDll) { attrib -R $dstDll }
    if (Test-Path $dstAddin) { attrib -R $dstAddin }

    Copy-Item -Path $srcDll -Destination $dstDll -Force
    Copy-Item -Path $srcAddin -Destination $dstAddin -Force

    New-Item -ItemType Directory -Path $roamDir2026 -Force | Out-Null
    New-Item -ItemType Directory -Path $roamDirLegacy -Force | Out-Null

    $roamDll2026 = Join-Path $roamDir2026 'AssemblyClonerAddIn.dll'
    $roamAddin2026 = Join-Path $roamDir2026 'AssemblyClonerAddIn.addin'
    $roamDllLegacy = Join-Path $roamDirLegacy 'AssemblyClonerAddIn.dll'
    $roamAddinLegacy = Join-Path $roamDirLegacy 'AssemblyClonerAddIn.addin'

    Copy-Item -Path $srcDll -Destination $roamDll2026 -Force
    Copy-Item -Path $srcAddin -Destination $roamAddin2026 -Force
    Copy-Item -Path $srcDll -Destination $roamDllLegacy -Force
    Copy-Item -Path $srcAddin -Destination $roamAddinLegacy -Force

    $srcHash = (Get-FileHash $srcDll -Algorithm SHA256).Hash
    $dstHash = (Get-FileHash $dstDll -Algorithm SHA256).Hash
    $roam2026Hash = (Get-FileHash $roamDll2026 -Algorithm SHA256).Hash
    $roamLegacyHash = (Get-FileHash $roamDllLegacy -Algorithm SHA256).Hash
    $dstInfo = Get-Item $dstDll

    @(
        'STATUS=OK'
        "SRC_HASH=$srcHash"
        "PROGRAMDATA_HASH=$dstHash"
        "ROAM_2026_HASH=$roam2026Hash"
        "ROAM_LEGACY_HASH=$roamLegacyHash"
        "PROGRAMDATA_MATCH=$([string]::Equals($srcHash,$dstHash,[System.StringComparison]::OrdinalIgnoreCase))"
        "ROAM_2026_MATCH=$([string]::Equals($srcHash,$roam2026Hash,[System.StringComparison]::OrdinalIgnoreCase))"
        "ROAM_LEGACY_MATCH=$([string]::Equals($srcHash,$roamLegacyHash,[System.StringComparison]::OrdinalIgnoreCase))"
        "DST_LASTWRITE=$($dstInfo.LastWriteTime.ToString('o'))"
        "DST_LENGTH=$($dstInfo.Length)"
    ) | Set-Content -Path $report -Encoding UTF8
}
catch {
    @(
        'STATUS=ERROR'
        "ERROR=$($_.Exception.Message)"
        "DETAIL=$($_ | Out-String)"
    ) | Set-Content -Path $report -Encoding UTF8
    throw
}
