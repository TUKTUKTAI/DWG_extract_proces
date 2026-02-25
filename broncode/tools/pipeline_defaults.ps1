function Get-NbdPipelineDefaults {
    param(
        [string]$ProjectRoot = ""
    )

    if ($ProjectRoot -and $ProjectRoot.Trim() -ne "") {
        $root = (Resolve-Path $ProjectRoot).Path
    }
    else {
        # Script expected in <project>\broncode\tools
        $root = (Resolve-Path (Join-Path $PSScriptRoot "..\..")).Path
    }

    $odaCandidates = @(
        "C:\Program Files\ODA\ODAFileConverter 26.12.0\ODAFileConverter.exe",
        "C:\Program Files\ODA\ODAFileConverter\ODAFileConverter.exe",
        "C:\Program Files (x86)\ODA\ODAFileConverter\ODAFileConverter.exe"
        "C:\Program Files\ODA\ODAFileConverter 25.12.0\ODAFileConverter.exe"
    )
    $odaExe = $null
    foreach ($candidate in $odaCandidates) {
        if (Test-Path $candidate) {
            $odaExe = $candidate
            break
        }
    }
    if (-not $odaExe) {
        $odaExe = "ODAFileConverter.exe"
    }

    return @{
        ProjectRoot = $root
        DwgInput = (Join-Path $root "dwgs")
        DxfOutput = (Join-Path $root "DXF")
        BronOutput = (Join-Path $root "Bron")
        DoelOutput = (Join-Path $root "Doel")
        EindOutput = (Join-Path $root "Eindresultaat")
        OdaExe = $odaExe
        OdaOutputVersion = "ACAD2013"
        OdaRecursive = $true
        OdaAudit = $false
        OdaWaitTimeoutSeconds = 120
        StrictNbdValidation = $true
        PythonExe = "python"
        JavaCompileFirst = $true
        JavaRunAfter = $true
    }
}

function Format-NbdDuration {
    param([TimeSpan]$Elapsed)
    "{0:00}:{1:00}:{2:00}.{3:000}" -f $Elapsed.Hours, $Elapsed.Minutes, $Elapsed.Seconds, $Elapsed.Milliseconds
}
