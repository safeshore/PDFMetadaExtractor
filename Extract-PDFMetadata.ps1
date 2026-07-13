param(
    [Parameter(Mandatory = $true)]
    [string]$Folder
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path -Path $Folder -PathType Container)) {
    throw "Folder does not exist or is not a directory: $Folder"
}

$Folder = (Resolve-Path -Path $Folder).Path

if ($PSScriptRoot) {
    $ScriptFolder = $PSScriptRoot
}
elseif ($MyInvocation.MyCommand.Path) {
    $ScriptFolder = Split-Path -Parent $MyInvocation.MyCommand.Path
}
else {
    $ScriptFolder = (Get-Location).Path
}

$RunTimestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$TxtOutput    = Join-Path -Path $Folder -ChildPath ("PDFMetadata_{0}.txt" -f $RunTimestamp)
$CsvOutput    = Join-Path -Path $Folder -ChildPath ("PDFMetadata_{0}.csv" -f $RunTimestamp)

function Get-PdfInfoCommand {
    $command = Get-Command -Name pdfinfo -ErrorAction SilentlyContinue

    if ($null -ne $command) {
        return $command
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
            return [pscustomobject]@{ Source = $pdfInfoExe.FullName }
        }
    }

    return $null
}

$PdfInfoCommand = $null

try {
    $PdfInfoCommand = Get-PdfInfoCommand

    if ($null -eq $PdfInfoCommand) {
        throw "pdfinfo not found"
    }

    Write-Host "Using pdfinfo at: $($PdfInfoCommand.Source)"
}
catch {
    Write-Warning "pdfinfo was not found. The script will try intrinsic parsing and Windows Shell metadata instead."
}

$Shell = $null

try {
    $Shell = New-Object -ComObject Shell.Application
}
catch {
    Write-Warning "Could not create Shell.Application COM object. Windows Shell metadata fallback is unavailable. Error: $($_.Exception.Message)"
}

function Get-IntrinsicPdfPageCount {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo]$PdfFile
    )

    try {
        # Parse raw bytes and count '/Type /Page' objects. This avoids external dependencies.
        $bytes = [System.IO.File]::ReadAllBytes($PdfFile.FullName)
        $ascii = [System.Text.Encoding]::ASCII.GetString($bytes)

        $matches = [regex]::Matches($ascii, '/Type\s*/Page(?!s)')

        if ($matches.Count -gt 0) {
            return $matches.Count
        }
    }
    catch {
        Write-Warning "Intrinsic parser failed for: $($PdfFile.FullName). Error: $($_.Exception.Message)"
    }

    return $null
}

function Get-PdfPageCount {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo]$PdfFile,

        [Parameter(Mandatory = $false)]
        $PdfInfoCommand,

        [Parameter(Mandatory = $false)]
        $Shell
    )

    if ($null -ne $PdfInfoCommand) {
        try {
            $pdfInfoOutput = & $PdfInfoCommand.Source $PdfFile.FullName 2>&1

            if ($LASTEXITCODE -eq 0) {
                foreach ($line in $pdfInfoOutput) {
                    if ($line -match '^Pages:\s+(\d+)') {
                        return $Matches[1]
                    }
                }

                Write-Warning "pdfinfo completed but did not return a page count for: $($PdfFile.FullName)"
            }
            else {
                Write-Warning "pdfinfo returned exit code $LASTEXITCODE for: $($PdfFile.FullName)"
                Write-Warning "pdfinfo output: $pdfInfoOutput"
            }
        }
        catch {
            Write-Warning "pdfinfo failed for: $($PdfFile.FullName). Error: $($_.Exception.Message)"
        }
    }

    $intrinsicCount = Get-IntrinsicPdfPageCount -PdfFile $PdfFile

    if ($null -ne $intrinsicCount -and $intrinsicCount -gt 0) {
        return $intrinsicCount
    }

    if ($null -ne $Shell) {
        try {
            $directory = $Shell.Namespace($PdfFile.Directory.FullName)

            if ($null -eq $directory) {
                Write-Warning "Could not open Shell namespace for directory: $($PdfFile.Directory.FullName)"
                return "Unknown"
            }

            $item = $directory.ParseName($PdfFile.Name)

            if ($null -eq $item) {
                Write-Warning "Could not parse Shell item for file: $($PdfFile.FullName)"
                return "Unknown"
            }

            for ($x = 0; $x -lt 400; $x++) {
                $propertyName = $directory.GetDetailsOf($null, $x)

                if ($propertyName -match 'Pages') {
                    $pageCount = $directory.GetDetailsOf($item, $x)

                    if (-not [string]::IsNullOrWhiteSpace($pageCount)) {
                        return $pageCount
                    }

                    Write-Warning "Shell metadata had a Pages field, but it was blank for: $($PdfFile.FullName)"
                    return "Unknown"
                }
            }

            Write-Warning "Shell metadata did not expose a Pages field for: $($PdfFile.FullName)"
        }
        catch {
            Write-Warning "Shell metadata lookup failed for: $($PdfFile.FullName). Error: $($_.Exception.Message)"
        }
    }

    return "Unknown"
}

try {
    $PdfFiles = Get-ChildItem -Path $Folder -Filter "*.pdf" -Recurse -File

    $Results = foreach ($PdfFile in $PdfFiles) {
        $pageCount = Get-PdfPageCount `
            -PdfFile $PdfFile `
            -PdfInfoCommand $PdfInfoCommand `
            -Shell $Shell

        [pscustomobject]@{
            FileName = $PdfFile.Name
            FullPath = $PdfFile.FullName
            Pages    = $pageCount
        }
    }

    $Results = @($Results | Sort-Object FileName)

    if ($Results.Count -eq 0) {
        Write-Warning "No PDF files were found in: $Folder"

        "FileName,FullPath,Pages" | Set-Content -Path $CsvOutput -Encoding UTF8
        "No PDF files were found in: $Folder" | Set-Content -Path $TxtOutput -Encoding UTF8
    }
    else {
        $Results |
            Export-Csv -Path $CsvOutput -NoTypeInformation -Encoding UTF8

        $Results |
            Format-Table -AutoSize |
            Out-String |
            Set-Content -Path $TxtOutput -Encoding UTF8
    }

    Write-Host "PDF metadata export complete."
    Write-Host "PDF files processed: $($Results.Count)"
    Write-Host "CSV output: $CsvOutput"
    Write-Host "TXT output: $TxtOutput"
}
catch {
    Write-Error "Failed to complete PDF metadata export. Error: $($_.Exception.Message)"
    exit 1
}