param(
    [switch]$SkipPdfInfoInstall
)

$ErrorActionPreference = "Stop"

function Add-ToUserPathIfMissing {
    param(
        [Parameter(Mandatory = $true)]
        [string]$PathEntry
    )

    $currentUserPath = [Environment]::GetEnvironmentVariable("Path", "User")

    if ([string]::IsNullOrWhiteSpace($currentUserPath)) {
        [Environment]::SetEnvironmentVariable("Path", $PathEntry, "User")
        Write-Host "Added to user PATH: $PathEntry"
        return
    }

    if ($currentUserPath.Split(";") -contains $PathEntry) {
        Write-Host "Path already contains: $PathEntry"
        return
    }

    [Environment]::SetEnvironmentVariable("Path", "$currentUserPath;$PathEntry", "User")
    Write-Host "Added to user PATH: $PathEntry"
}

function Find-PdfInfoExecutable {
    $pdfInfoCommand = Get-Command -Name pdfinfo -ErrorAction SilentlyContinue

    if ($null -ne $pdfInfoCommand) {
        return $pdfInfoCommand.Source
    }

    $candidateRoots = @(
        (Join-Path -Path $env:LOCALAPPDATA -ChildPath "Microsoft\WinGet\Packages"),
        (Join-Path -Path $env:ProgramFiles -ChildPath "poppler"),
        (Join-Path -Path $env:ProgramFiles -ChildPath "Poppler")
    )

    foreach ($root in $candidateRoots) {
        if (-not (Test-Path -Path $root -PathType Container)) {
            continue
        }

        $pdfInfoExe = Get-ChildItem -Path $root -Filter "pdfinfo.exe" -Recurse -ErrorAction SilentlyContinue |
            Select-Object -First 1

        if ($null -ne $pdfInfoExe) {
            return $pdfInfoExe.FullName
        }
    }

    return $null
}

if ($PSScriptRoot) {
    $ScriptFolder = $PSScriptRoot
}
elseif ($MyInvocation.MyCommand.Path) {
    $ScriptFolder = Split-Path -Parent $MyInvocation.MyCommand.Path
}
else {
    $ScriptFolder = (Get-Location).Path
}

if (-not (Test-Path -Path $ScriptFolder -PathType Container)) {
    throw "Script folder does not exist: $ScriptFolder"
}

$CurrentUserPath = [Environment]::GetEnvironmentVariable("Path", "User")

Add-ToUserPathIfMissing -PathEntry $ScriptFolder

$PdfInfoPath = Find-PdfInfoExecutable

if (-not [string]::IsNullOrWhiteSpace($PdfInfoPath)) {
    Write-Host "pdfinfo is available at: $PdfInfoPath"

    $pdfInfoDirectory = Split-Path -Parent $PdfInfoPath
    Add-ToUserPathIfMissing -PathEntry $pdfInfoDirectory
}
elseif ($SkipPdfInfoInstall) {
    Write-Warning "pdfinfo is missing and auto-install was skipped. Page counts may be less reliable."
}
else {
    $WingetCommand = Get-Command -Name winget -ErrorAction SilentlyContinue

    if ($null -eq $WingetCommand) {
        Write-Warning "winget is not available, so pdfinfo could not be installed automatically."
    }
    else {
        Write-Host "Attempting to install Poppler (provides pdfinfo) with winget..."

        try {
            & $WingetCommand.Source install --id oschwartz10612.Poppler --exact --accept-package-agreements --accept-source-agreements --silent | Out-Null
            Write-Host "Poppler install command completed."

            $PdfInfoPath = Find-PdfInfoExecutable

            if (-not [string]::IsNullOrWhiteSpace($PdfInfoPath)) {
                Write-Host "pdfinfo is available at: $PdfInfoPath"

                $pdfInfoDirectory = Split-Path -Parent $PdfInfoPath
                Add-ToUserPathIfMissing -PathEntry $pdfInfoDirectory
            }
            else {
                Write-Warning "Poppler was installed, but pdfinfo.exe was not discovered automatically."
            }
        }
        catch {
            Write-Warning "Failed to install Poppler automatically. Error: $($_.Exception.Message)"
        }
    }
}

Write-Host ""
Write-Host "Close and reopen PowerShell for the change to take effect."