param(
    [string]$DwgInput = "",
    [string]$DxfOutput = "",
    [string]$ProjectRoot = "",

    [string]$OdaExe = "",
    [string]$OutputVersion = "",
    [switch]$Recursive,
    [switch]$Audit,
    [int]$WaitTimeoutSeconds = 0
)

$ErrorActionPreference = "Stop"
$swTotal = [System.Diagnostics.Stopwatch]::StartNew()
. (Join-Path $PSScriptRoot "pipeline_defaults.ps1")
$cfg = Get-NbdPipelineDefaults -ProjectRoot $ProjectRoot

if ([string]::IsNullOrWhiteSpace($DwgInput)) { $DwgInput = $cfg.DwgInput }
if ([string]::IsNullOrWhiteSpace($DxfOutput)) { $DxfOutput = $cfg.DxfOutput }
if ([string]::IsNullOrWhiteSpace($OdaExe)) { $OdaExe = $cfg.OdaExe }
if ([string]::IsNullOrWhiteSpace($OutputVersion)) { $OutputVersion = $cfg.OdaOutputVersion }
if ($WaitTimeoutSeconds -le 0) { $WaitTimeoutSeconds = [int]$cfg.OdaWaitTimeoutSeconds }
if (-not $PSBoundParameters.ContainsKey("Recursive")) { $Recursive = [bool]$cfg.OdaRecursive }
if (-not $PSBoundParameters.ContainsKey("Audit")) { $Audit = [bool]$cfg.OdaAudit }

function Resolve-InputContext {
    param([string]$PathValue)

    $resolved = (Resolve-Path $PathValue).Path
    $item = Get-Item $resolved

    if ($item.PSIsContainer) {
        return @{
            SourceDir = $resolved
            Filter = "*.dwg"
        }
    }

    return @{
        SourceDir = $item.DirectoryName
        Filter = $item.Name
    }
}

function Get-DxfOutputSnapshot {
    param([string]$PathValue)

    $files = Get-ChildItem -Path $PathValue -Recurse -Filter *.dxf -File -ErrorAction SilentlyContinue
    $count = @($files).Count
    $sumLength = 0L
    $latestTicks = 0L

    foreach ($f in $files) {
        $sumLength += [int64]$f.Length
        $ticks = $f.LastWriteTimeUtc.Ticks
        if ($ticks -gt $latestTicks) {
            $latestTicks = $ticks
        }
    }

    return [pscustomobject]@{
        Count = $count
        SumLength = $sumLength
        LatestTicks = $latestTicks
    }
}

if (!(Test-Path $DwgInput)) {
    throw "Input path not found: $DwgInput"
}

if (!(Test-Path $DxfOutput)) {
    New-Item -ItemType Directory -Path $DxfOutput | Out-Null
}

$ctx = Resolve-InputContext -PathValue $DwgInput
$recurseFlag = if ($Recursive) { "1" } else { "0" }
$auditFlag = if ($Audit) { "1" } else { "0" }

Write-Host "ODA converter :" $OdaExe
Write-Host "Source dir    :" $ctx.SourceDir
Write-Host "Target dir    :" (Resolve-Path $DxfOutput).Path
Write-Host "Filter        :" $ctx.Filter
Write-Host "Version       :" $OutputVersion
Write-Host "Recursive     :" $recurseFlag
Write-Host "Audit         :" $auditFlag

# ODA File Converter CLI-syntax:
#   ODAFileConverter <sourceDir> <targetDir> <version> <type> <recursive> <audit> <filter>
$argsList = @(
    $ctx.SourceDir,
    (Resolve-Path $DxfOutput).Path,
    $OutputVersion,
    "DXF",
    $recurseFlag,
    $auditFlag,
    $ctx.Filter
)

# Gebruik PowerShell's call-operator zodat argumenten met spaties correct als losse args worden doorgegeven.
# ODA File Converter kan terugkeren voordat alle bestanden zijn weggeschreven, daarom pollen we hieronder de outputmap.
$expectedDwgCount = (Get-ChildItem -Path $ctx.SourceDir -Recurse:$Recursive -Filter $ctx.Filter -File -ErrorAction SilentlyContinue | Measure-Object).Count
Write-Host "Expected DWGs :" $expectedDwgCount

& $OdaExe @argsList
$exit = $LASTEXITCODE

$dxfCount = 0
$stablePolls = 0
$requiredStablePolls = 4   # 4 * 500ms = ~2s without changes
$prevSnapshot = $null
$deadline = (Get-Date).AddSeconds([Math]::Max(0, $WaitTimeoutSeconds))
do {
    $snapshot = Get-DxfOutputSnapshot -PathValue $DxfOutput
    $dxfCount = [int]$snapshot.Count

    if ($dxfCount -gt 0) {
        if ($null -ne $prevSnapshot -and
            $snapshot.Count -eq $prevSnapshot.Count -and
            $snapshot.SumLength -eq $prevSnapshot.SumLength -and
            $snapshot.LatestTicks -eq $prevSnapshot.LatestTicks) {
            $stablePolls++
        }
        else {
            $stablePolls = 0
        }

        $enoughFiles = ($expectedDwgCount -le 0) -or ($dxfCount -ge $expectedDwgCount)
        if ($enoughFiles -and $stablePolls -ge $requiredStablePolls) {
            break
        }
    }

    $prevSnapshot = $snapshot
    Start-Sleep -Milliseconds 500
} while ((Get-Date) -lt $deadline)

if ($null -ne $exit -and $exit -ne 0 -and $dxfCount -eq 0) {
    throw "ODAFileConverter failed with exit code $exit"
}

if ($dxfCount -eq 0) {
    throw "ODAFileConverter finished but no DXF files were created within $WaitTimeoutSeconds seconds. Check source files and ODA settings/version arguments."
}

if ($expectedDwgCount -gt 0 -and $dxfCount -lt $expectedDwgCount) {
    Write-Warning "Fewer DXF files found than expected DWGs ($dxfCount/$expectedDwgCount) after waiting $WaitTimeoutSeconds seconds."
}

Write-Host "Done. ODA exit code:" $exit
Write-Host "DXF files in output:" $dxfCount
$swTotal.Stop()
Write-Host "Duur (DWG->DXF):" ("{0:00}:{1:00}:{2:00}.{3:000}" -f $swTotal.Elapsed.Hours, $swTotal.Elapsed.Minutes, $swTotal.Elapsed.Seconds, $swTotal.Elapsed.Milliseconds)
