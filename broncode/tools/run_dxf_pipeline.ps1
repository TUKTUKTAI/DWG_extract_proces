param(
    [string]$DxfInput = "",

    [string]$ProjectRoot = "",
    [string]$PythonExe = "python",
    [switch]$RunJavaAfter,
    [switch]$CompileJavaFirst,
    [switch]$StrictNbdValidation,
    [switch]$MergeAfter
)

$ErrorActionPreference = "Stop"
$swTotal = [System.Diagnostics.Stopwatch]::StartNew()
. (Join-Path $PSScriptRoot "pipeline_defaults.ps1")

function Resolve-ProjectRoot {
    param([string]$Value)
    if ($Value -and $Value.Trim() -ne "") {
        return (Resolve-Path $Value).Path
    }
    # Script is expected in <project>\broncode\tools
    return (Resolve-Path (Join-Path $PSScriptRoot "..\..")).Path
}

$root = Resolve-ProjectRoot -Value $ProjectRoot
$cfg = Get-NbdPipelineDefaults -ProjectRoot $root
$broncode = Join-Path $root "broncode"
$bron = Join-Path $root "Bron"
$doel = Join-Path $root "Doel"
$javaFile = Join-Path $broncode "P22_0002_Main.java"
$javaCp = ".\broncode;.\broncode\lib\*"
$errorLog = Join-Path $doel "dxf_extract_errors.csv"
$mergeScript = Join-Path $broncode "tools\merge_naverwerking_results.py"

if (!(Test-Path $bron)) { New-Item -ItemType Directory -Path $bron | Out-Null }
if (!(Test-Path $doel)) { New-Item -ItemType Directory -Path $doel | Out-Null }

if ([string]::IsNullOrWhiteSpace($DxfInput)) { $DxfInput = $cfg.DxfOutput }
if (-not $PSBoundParameters.ContainsKey("RunJavaAfter")) { $RunJavaAfter = [bool]$cfg.JavaRunAfter }
if (-not $PSBoundParameters.ContainsKey("CompileJavaFirst")) { $CompileJavaFirst = [bool]$cfg.JavaCompileFirst }
if (-not $PSBoundParameters.ContainsKey("StrictNbdValidation")) { $StrictNbdValidation = [bool]$cfg.StrictNbdValidation }

Write-Host "Project root: $root"
Write-Host "DXF input: $DxfInput"
Write-Host "Bron output (extracts): $bron"

Push-Location $root
try {
    $swExtract = [System.Diagnostics.Stopwatch]::StartNew()
    $pyArgs = @(
        ".\broncode\tools\dxf_to_extract_xlsx.py",
        "--input", $DxfInput,
        "--output-dir", $bron,
        "--recursive",
        "--error-log", $errorLog
    )
    if ($StrictNbdValidation) {
        $pyArgs += "--strict-nbd"
    }

    & $PythonExe @pyArgs

    if ($LASTEXITCODE -ne 0) {
        throw "DXF extract step failed with exit code $LASTEXITCODE"
    }
    $swExtract.Stop()
    Write-Host ("Duur DXF -> extract: {0:00}:{1:00}:{2:00}.{3:000}" -f $swExtract.Elapsed.Hours, $swExtract.Elapsed.Minutes, $swExtract.Elapsed.Seconds, $swExtract.Elapsed.Milliseconds)

    if ($RunJavaAfter) {
        $javac = Join-Path $root "jdk-13.0.2\bin\javac.exe"
        $java = Join-Path $root "jdk-13.0.2\bin\java.exe"

        if ($CompileJavaFirst) {
            $swJavac = [System.Diagnostics.Stopwatch]::StartNew()
            & $javac -cp $javaCp $javaFile
            if ($LASTEXITCODE -ne 0) {
                throw "javac failed with exit code $LASTEXITCODE"
            }
            $swJavac.Stop()
            Write-Host ("Duur Java compile: {0:00}:{1:00}:{2:00}.{3:000}" -f $swJavac.Elapsed.Hours, $swJavac.Elapsed.Minutes, $swJavac.Elapsed.Seconds, $swJavac.Elapsed.Milliseconds)
        }

        $swJava = [System.Diagnostics.Stopwatch]::StartNew()
        & $java -cp $javaCp "P22_0002_Main"
        if ($LASTEXITCODE -ne 0) {
            throw "Java post-processing failed with exit code $LASTEXITCODE"
        }
        $swJava.Stop()
        Write-Host ("Duur Java naverwerking: {0:00}:{1:00}:{2:00}.{3:000}" -f $swJava.Elapsed.Hours, $swJava.Elapsed.Minutes, $swJava.Elapsed.Seconds, $swJava.Elapsed.Milliseconds)
    }

    if ($MergeAfter) {
        $swMerge = [System.Diagnostics.Stopwatch]::StartNew()
        & $PythonExe $mergeScript
        if ($LASTEXITCODE -ne 0) {
            throw "Merge step failed with exit code $LASTEXITCODE"
        }
        $swMerge.Stop()
        Write-Host ("Duur merge naverwerking: {0:00}:{1:00}:{2:00}.{3:000}" -f $swMerge.Elapsed.Hours, $swMerge.Elapsed.Minutes, $swMerge.Elapsed.Seconds, $swMerge.Elapsed.Milliseconds)
    }
}
finally {
    Pop-Location
}

$swTotal.Stop()
Write-Host "Pipeline completed."
Write-Host ("Totale duur DXF-pipeline: {0:00}:{1:00}:{2:00}.{3:000}" -f $swTotal.Elapsed.Hours, $swTotal.Elapsed.Minutes, $swTotal.Elapsed.Seconds, $swTotal.Elapsed.Milliseconds)
