param(
    [string]$DwgInput = "",
    [string]$DxfOutput = "",
    [string]$ProjectRoot = "",
    [string]$OdaExe = "",
    [string]$OutputVersion = "",
    [switch]$Recursive,
    [switch]$Audit,
    [int]$WaitTimeoutSeconds = 0,
    [string]$ErrorLog = ""
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
if ([string]::IsNullOrWhiteSpace($ErrorLog)) { $ErrorLog = (Join-Path $cfg.DoelOutput "oda_per_file_errors.csv") }
if (-not $PSBoundParameters.ContainsKey("Recursive")) { $Recursive = [bool]$cfg.OdaRecursive }
if (-not $PSBoundParameters.ContainsKey("Audit")) { $Audit = [bool]$cfg.OdaAudit }

function Get-DwgFiles {
    param([string]$InputPath, [switch]$RecursiveFlag)
    $resolved = (Resolve-Path $InputPath).Path
    $item = Get-Item $resolved
    if ($item.PSIsContainer) {
        if ($RecursiveFlag) { return Get-ChildItem -Path $resolved -Recurse -File -Filter *.dwg }
        return Get-ChildItem -Path $resolved -File -Filter *.dwg
    }
    return @($item)
}

function Ensure-Dir {
    param([string]$PathValue)
    if (!(Test-Path $PathValue)) { New-Item -ItemType Directory -Path $PathValue | Out-Null }
}

function Append-ErrorLog {
    param([string]$CsvPath, [string]$DwgFile, [string]$Message)
    if ([string]::IsNullOrWhiteSpace($CsvPath)) { return }
    $dir = Split-Path -Parent $CsvPath
    if ($dir) { Ensure-Dir $dir }
    if (!(Test-Path $CsvPath)) {
        "timestamp;dwg;error" | Out-File -FilePath $CsvPath -Encoding utf8
    }
    $line = "{0};{1};{2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), ($DwgFile -replace ";", ","), ($Message -replace "[\r\n;]", " ")
    $line | Out-File -FilePath $CsvPath -Append -Encoding utf8
}

function Wait-ForFile {
    param([string]$PathValue, [int]$TimeoutSeconds)
    $deadline = (Get-Date).AddSeconds([Math]::Max(0, $TimeoutSeconds))
    do {
        if (Test-Path $PathValue) { return $true }
        Start-Sleep -Milliseconds 500
    } while ((Get-Date) -lt $deadline)
    return (Test-Path $PathValue)
}

if (!(Test-Path $DwgInput)) { throw "Input path not found: $DwgInput" }
Ensure-Dir $DxfOutput

$dxfRoot = (Resolve-Path $DxfOutput).Path
$dwgFiles = @(Get-DwgFiles -InputPath $DwgInput -RecursiveFlag:$Recursive)
if ($dwgFiles.Count -eq 0) { throw "No DWG files found in $DwgInput" }

$inputResolved = (Resolve-Path $DwgInput).Path
$inputIsFolder = (Get-Item $inputResolved).PSIsContainer
$auditFlag = if ($Audit) { "1" } else { "0" }
$ok = 0
$failed = 0

Write-Host "ODA converter :" $OdaExe
Write-Host "DWG count     :" $dwgFiles.Count
Write-Host "Target dir    :" $dxfRoot
Write-Host "Version       :" $OutputVersion
Write-Host "Audit         :" $auditFlag
Write-Host "Mode          : per-file isolation"

for ($i = 0; $i -lt $dwgFiles.Count; $i++) {
    $dwg = $dwgFiles[$i]
    try {
        $sourceDir = $dwg.DirectoryName
        $relativeSubdir = ""
        if ($inputIsFolder) {
            $parent = $dwg.DirectoryName
            if ($parent.StartsWith($inputResolved, [System.StringComparison]::OrdinalIgnoreCase)) {
                $relativeSubdir = $parent.Substring($inputResolved.Length).TrimStart('\')
            }
        }

        $targetDir = if ([string]::IsNullOrWhiteSpace($relativeSubdir)) { $dxfRoot } else { Join-Path $dxfRoot $relativeSubdir }
        Ensure-Dir $targetDir
        $expectedDxf = Join-Path $targetDir ([System.IO.Path]::GetFileNameWithoutExtension($dwg.Name) + ".dxf")

        Write-Host ("[{0}/{1}] {2}" -f ($i + 1), $dwgFiles.Count, $dwg.Name)

        $argsList = @(
            $sourceDir,
            $targetDir,
            $OutputVersion,
            "DXF",
            "0",
            $auditFlag,
            $dwg.Name
        )

        & $OdaExe @argsList
        $exit = $LASTEXITCODE
        $exists = Wait-ForFile -PathValue $expectedDxf -TimeoutSeconds $WaitTimeoutSeconds
        if (-not $exists) {
            throw "No DXF created within $WaitTimeoutSeconds s (exit=$exit)"
        }

        $ok++
    }
    catch {
        $failed++
        $msg = $_.Exception.Message
        Write-Host "  FAIL: $msg" -ForegroundColor Red
        Append-ErrorLog -CsvPath $ErrorLog -DwgFile $dwg.FullName -Message $msg
        continue
    }
}

$swTotal.Stop()
Write-Host ""
Write-Host "Done. Success: $ok / $($dwgFiles.Count); Failed: $failed"
Write-Host "Duur (DWG->DXF per-file):" (Format-NbdDuration $swTotal.Elapsed)
if ($ErrorLog -and (Test-Path $ErrorLog)) {
    Write-Host "Error log: $ErrorLog"
}
