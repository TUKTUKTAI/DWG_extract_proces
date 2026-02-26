param(
    [string]$ProjectRoot = "",
    [switch]$UsePerFileOda,
    [switch]$SkipJavaPostProcessing,
    [switch]$CleanupDxfAfter,
    [switch]$MergeAfter
)

$ErrorActionPreference = "Stop"
$swTotal = [System.Diagnostics.Stopwatch]::StartNew()

. (Join-Path $PSScriptRoot "pipeline_defaults.ps1")

function Resolve-NbdProjectRoot {
    param([string]$Value)
    if ($Value -and $Value.Trim() -ne "") {
        return (Resolve-Path $Value).Path
    }
    return (Resolve-Path (Join-Path $PSScriptRoot "..\..")).Path
}

$root = Resolve-NbdProjectRoot -Value $ProjectRoot
$cfg = Get-NbdPipelineDefaults -ProjectRoot $root
$dxfDir = $cfg.DxfOutput
$doelDir = $cfg.DoelOutput
$eindDir = $cfg.EindOutput

if (!(Test-Path $doelDir)) { New-Item -ItemType Directory -Path $doelDir | Out-Null }
if (!(Test-Path $eindDir)) { New-Item -ItemType Directory -Path $eindDir | Out-Null }

$odaScript = if ($UsePerFileOda) {
    Join-Path $PSScriptRoot "convert_dwg_to_dxf_per_file.ps1"
} else {
    Join-Path $PSScriptRoot "convert_dwg_to_dxf.ps1"
}
$pipelineScript = Join-Path $PSScriptRoot "run_dxf_pipeline.ps1"
$mergeScript = Join-Path $PSScriptRoot "merge_naverwerking_results.py"

Write-Host "Project root:" $root
Write-Host "Mode       :" ($(if ($UsePerFileOda) { "ODA per file" } else { "ODA batch" }))
Write-Host "DXF map     :" $dxfDir
Write-Host "Run Java    :" ($(if ($SkipJavaPostProcessing) { "No" } else { "Yes" }))
Write-Host "Merge after :" ($(if ($MergeAfter) { "Yes" } else { "No" }))
Write-Host "Cleanup DXF :" ($(if ($CleanupDxfAfter) { "Yes" } else { "No" }))

Push-Location $root
try {
    $swOda = [System.Diagnostics.Stopwatch]::StartNew()
    & powershell -ExecutionPolicy Bypass -File $odaScript
    if ($LASTEXITCODE -ne 0) {
        throw "DWG -> DXF step failed with exit code $LASTEXITCODE"
    }
    $swOda.Stop()
    Write-Host "Duur stap 1 (DWG -> DXF):" (Format-NbdDuration $swOda.Elapsed)

    $swPipe = [System.Diagnostics.Stopwatch]::StartNew()
    $pipelineArgs = @(
        "-ExecutionPolicy", "Bypass",
        "-File", $pipelineScript
    )
    if ($SkipJavaPostProcessing) {
        $pipelineArgs += "-RunJavaAfter:$false"
    }
    & powershell @pipelineArgs
    if ($LASTEXITCODE -ne 0) {
        throw "DXF pipeline step failed with exit code $LASTEXITCODE"
    }
    $swPipe.Stop()
    Write-Host "Duur stap 2 (DXF -> extract -> Java):" (Format-NbdDuration $swPipe.Elapsed)

    if ($MergeAfter) {
        $swMerge = [System.Diagnostics.Stopwatch]::StartNew()
        & $cfg.PythonExe $mergeScript
        if ($LASTEXITCODE -ne 0) {
            throw "Merge step failed with exit code $LASTEXITCODE"
        }
        $swMerge.Stop()
        Write-Host "Duur stap 3 (merge):" (Format-NbdDuration $swMerge.Elapsed)
    }

    if ($CleanupDxfAfter -and (Test-Path $dxfDir)) {
        # Keep .gitkeep if present; remove generated DXFs and subfolders.
        Get-ChildItem -Path $dxfDir -Force -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -ne ".gitkeep" } |
            Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
        Write-Host "DXF tijdelijke bestanden opgeschoond (behalve .gitkeep)."
    }
}
finally {
    Pop-Location
    $swTotal.Stop()
    Write-Host "Totale duur (DWG -> Excel wrapper):" (Format-NbdDuration $swTotal.Elapsed)
}

