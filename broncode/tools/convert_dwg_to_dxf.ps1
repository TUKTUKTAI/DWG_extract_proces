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
$odaBatchErrorLog = Join-Path $cfg.DoelOutput "oda_batch_errors.csv"

if ([string]::IsNullOrWhiteSpace($DwgInput)) { $DwgInput = $cfg.DwgInput }
if ([string]::IsNullOrWhiteSpace($DxfOutput)) { $DxfOutput = $cfg.DxfOutput }
if ([string]::IsNullOrWhiteSpace($OdaExe)) { $OdaExe = $cfg.OdaExe }
if ([string]::IsNullOrWhiteSpace($OutputVersion)) { $OutputVersion = $cfg.OdaOutputVersion }
$waitTimeoutExplicit = $PSBoundParameters.ContainsKey("WaitTimeoutSeconds")
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
    param(
        [string]$PathValue,
        [datetime]$ModifiedSinceUtc = [datetime]::MinValue
    )

    $files = Get-ChildItem -Path $PathValue -Recurse -Filter *.dxf -File -ErrorAction SilentlyContinue
    if ($ModifiedSinceUtc -ne [datetime]::MinValue) {
        $files = @($files | Where-Object { $_.LastWriteTimeUtc -ge $ModifiedSinceUtc })
    }
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

function Get-AutoWaitTimeoutSeconds {
    param(
        [int]$DefaultSeconds,
        [int]$ExpectedDwgCount
    )

    if ($ExpectedDwgCount -le 0) {
        return [Math]::Max(120, $DefaultSeconds)
    }

    # Heuristiek voor grote batches:
    # basis 120s + ~1s per 20 DWG, gemaximeerd op 30 min.
    $estimated = 120 + [Math]::Ceiling($ExpectedDwgCount / 20.0)
    return [Math]::Min(1800, [Math]::Max($DefaultSeconds, [int]$estimated))
}

function Get-OdaErrFiles {
    param(
        [string]$PathValue,
        [datetime]$ModifiedSinceUtc = [datetime]::MinValue
    )

    $files = Get-ChildItem -Path $PathValue -Recurse -Filter *.err -File -ErrorAction SilentlyContinue
    if ($ModifiedSinceUtc -ne [datetime]::MinValue) {
        $files = @($files | Where-Object { $_.LastWriteTimeUtc -ge $ModifiedSinceUtc })
    }
    return @($files)
}

function Write-OdaBatchErrorLog {
    param(
        [string]$CsvPath,
        [array]$ErrFiles
    )

    if ($null -eq $ErrFiles -or $ErrFiles.Count -eq 0) {
        return
    }

    $dir = Split-Path -Parent $CsvPath
    if ($dir -and !(Test-Path $dir)) {
        New-Item -ItemType Directory -Path $dir | Out-Null
    }

    if (!(Test-Path $CsvPath)) {
        "timestamp;err_file;related_dxf" | Out-File -FilePath $CsvPath -Encoding utf8 -Append
    }

    foreach ($f in $ErrFiles) {
        $relatedDxf = ""
        if ($f.Name.ToLower().EndsWith(".dxf.err")) {
            $relatedDxf = Join-Path $f.DirectoryName ($f.Name.Substring(0, $f.Name.Length - 4))
        }
        $line = "{0};{1};{2}" -f `
            (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), `
            ($f.FullName -replace ';', ','), `
            ($relatedDxf -replace ';', ',')
        Add-Content -Path $CsvPath -Value $line -Encoding utf8
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
if (-not $waitTimeoutExplicit) {
    $WaitTimeoutSeconds = Get-AutoWaitTimeoutSeconds -DefaultSeconds $WaitTimeoutSeconds -ExpectedDwgCount $expectedDwgCount
}
Write-Host "Expected DWGs :" $expectedDwgCount
Write-Host "Wait timeout  :" ($WaitTimeoutSeconds.ToString() + "s")

$runStartUtc = [DateTime]::UtcNow.AddSeconds(-2)
& $OdaExe @argsList
$exit = $LASTEXITCODE

$dxfCount = 0
$totalDxfCount = 0
$stablePolls = 0
$requiredStablePolls = 4   # 4 * 500ms = ~2s without changes
$prevSnapshot = $null
$deadline = (Get-Date).AddSeconds([Math]::Max(0, $WaitTimeoutSeconds))
$nextProgressLog = (Get-Date).AddSeconds(10)
$extensionsUsed = 0
$maxExtensions = 5
do {
    $snapshot = Get-DxfOutputSnapshot -PathValue $DxfOutput -ModifiedSinceUtc $runStartUtc
    $totalSnapshot = Get-DxfOutputSnapshot -PathValue $DxfOutput
    $dxfCount = [int]$snapshot.Count
    $totalDxfCount = [int]$totalSnapshot.Count

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

    if ((Get-Date) -ge $nextProgressLog) {
        if ($expectedDwgCount -gt 0) {
            Write-Host ("Wachten op ODA... DXF's deze run: {0}/{1} (totaal in map: {2})" -f $dxfCount, $expectedDwgCount, $totalDxfCount)
        } else {
            Write-Host ("Wachten op ODA... DXF's deze run: {0} (totaal in map: {1})" -f $dxfCount, $totalDxfCount)
        }
        $nextProgressLog = (Get-Date).AddSeconds(10)
    }

    $prevSnapshot = $snapshot
    Start-Sleep -Milliseconds 500

    if ((Get-Date) -ge $deadline -and $stablePolls -eq 0 -and $dxfCount -gt 0 -and $extensionsUsed -lt $maxExtensions) {
        $deadline = (Get-Date).AddSeconds(60)
        $extensionsUsed++
        Write-Host "ODA nog actief; wachttijd verlengd met 60s (extensie $extensionsUsed/$maxExtensions)."
    }
} while ((Get-Date) -lt $deadline)

if ($null -ne $exit -and $exit -ne 0 -and $dxfCount -eq 0) {
    throw "ODAFileConverter failed with exit code $exit"
}

if ($dxfCount -eq 0) {
    throw "ODAFileConverter finished but no DXF files were created within $WaitTimeoutSeconds seconds. Check source files and ODA settings/version arguments."
}

if ($expectedDwgCount -gt 0 -and $dxfCount -lt $expectedDwgCount) {
    Write-Warning "Fewer DXF files updated/created in this run than expected DWGs ($dxfCount/$expectedDwgCount). Total DXF files currently in map: $totalDxfCount."
}

$errFilesThisRun = Get-OdaErrFiles -PathValue $DxfOutput -ModifiedSinceUtc $runStartUtc
$totalErrFiles = (Get-OdaErrFiles -PathValue $DxfOutput).Count
if ($errFilesThisRun.Count -gt 0) {
    Write-Warning "ODA produced $($errFilesThisRun.Count) .err file(s) in this run. See $odaBatchErrorLog"
    Write-OdaBatchErrorLog -CsvPath $odaBatchErrorLog -ErrFiles $errFilesThisRun
}

Write-Host "Done. ODA exit code:" $exit
Write-Host "DXF files this run :" $dxfCount
Write-Host "DXF files in output:" $totalDxfCount
Write-Host ".err files this run:" $errFilesThisRun.Count
Write-Host ".err files in output:" $totalErrFiles
$swTotal.Stop()
Write-Host "Duur (DWG->DXF):" ("{0:00}:{1:00}:{2:00}.{3:000}" -f $swTotal.Elapsed.Hours, $swTotal.Elapsed.Minutes, $swTotal.Elapsed.Seconds, $swTotal.Elapsed.Milliseconds)
