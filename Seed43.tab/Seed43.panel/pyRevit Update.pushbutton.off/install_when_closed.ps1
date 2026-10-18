# install_when_closed.ps1
# Started hidden by pyRevit Update's script.py. Downloads the pyRevit
# installer, checks it, waits until no Revit is running, then opens it.
# Any failure shows a message box; every step is logged next to the download.

param(
    [Parameter(Mandatory = $true)][string]$Url,
    [Parameter(Mandatory = $true)][long]$Size,
    [Parameter(Mandatory = $true)][string]$OutFile,
    [string]$Version = ""
)

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'   # the progress bar makes Invoke-WebRequest many times slower
$workDir = Split-Path -Parent $OutFile
$log = Join-Path $workDir 'update.log'
$pending = Join-Path $workDir 'pending.json'

function Write-Log([string]$text) {
    Add-Content -Path $log -Value ("{0:yyyy-MM-dd HH:mm:ss}  {1}" -f (Get-Date), $text)
}

function Show-Message([string]$text, [string]$icon) {
    Add-Type -AssemblyName System.Windows.Forms
    [void][System.Windows.Forms.MessageBox]::Show($text, 'pyRevit Update', 'OK', $icon)
}

try {
    Write-Log "pyRevit $Version : downloading $Url"
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

    $existing = Get-Item $OutFile -ErrorAction SilentlyContinue
    if (-not $existing -or $existing.Length -ne $Size) {
        Invoke-WebRequest -Uri $Url -OutFile $OutFile -UseBasicParsing
    }

    $actual = (Get-Item $OutFile).Length
    if ($actual -ne $Size) {
        throw "Download is $actual bytes, expected $Size. Try again."
    }

    $sig = Get-AuthenticodeSignature -FilePath $OutFile
    if ($sig.Status -ne 'Valid') {
        Remove-Item $OutFile -Force
        throw "The installer's digital signature is not valid ($($sig.Status)), so it was deleted and not run."
    }
    Write-Log "downloaded and signature valid: $($sig.SignerCertificate.Subject)"

    Write-Log 'waiting for Revit to close'
    while (Get-Process -Name 'Revit' -ErrorAction SilentlyContinue) {
        Start-Sleep -Seconds 5
    }

    Write-Log 'Revit closed, starting installer'
    Start-Process -FilePath $OutFile
}
catch {
    Write-Log "FAILED: $($_.Exception.Message)"
    Show-Message ("pyRevit $Version was not installed.`n`n" + $_.Exception.Message + "`n`nLog: $log") 'Error'
}
finally {
    Remove-Item $pending -Force -ErrorAction SilentlyContinue
}
