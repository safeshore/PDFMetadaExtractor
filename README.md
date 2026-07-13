# PDFMetadaExtractor

PowerShell utility for scanning a folder of PDF files and exporting PDF page-count metadata.

This repository provides:
- `Extract-PDFMetadata.ps1`: the main extraction script.
- `Setup.ps1`: optional setup helper to add this repo and `pdfinfo` to your user `PATH`, and optionally install Poppler (`pdfinfo`) via `winget`.
- `Instructions/`: static instruction assets.

## What It Does

`Extract-PDFMetadata.ps1` recursively scans a target folder for `*.pdf` files and generates timestamped output files in that same folder:
- `PDFMetadata_yyyyMMdd_HHmmss.csv`
- `PDFMetadata_yyyyMMdd_HHmmss.txt`

Each output row includes:
- `FileName`
- `FullPath`
- `Pages`

Page count detection is attempted in this order:
1. `pdfinfo` (preferred, most reliable)
2. Intrinsic parser that counts `/Type /Page` objects in the PDF bytes
3. Windows Shell metadata fallback (COM `Shell.Application`)

If no PDFs are found, the script still writes output files with a helpful message.

## Requirements

- Windows PowerShell or PowerShell 7+
- Windows OS (Shell metadata fallback uses COM)
- Optional but recommended: `pdfinfo` from Poppler for best page-count accuracy

## Quick Start

1. (Optional) Run setup to configure `PATH` and install/discover `pdfinfo`:

```powershell
.\Setup.ps1
```

2. Run the extractor against a folder containing PDFs:

```powershell
.\Extract-PDFMetadata.ps1 -Folder "C:\Docs\PDFs"
```

3. Review generated files in the target folder:
- `PDFMetadata_*.csv`
- `PDFMetadata_*.txt`

## Usage Examples

Run against a local docs folder:

```powershell
.\Extract-PDFMetadata.ps1 -Folder "C:\Users\you\Documents\Contracts"
```

Run against the current directory:

```powershell
.\Extract-PDFMetadata.ps1 -Folder "."
```

Run setup but skip automatic Poppler install:

```powershell
.\Setup.ps1 -SkipPdfInfoInstall
```

## Example CSV Output

```csv
"FileName","FullPath","Pages"
"A-1001.pdf","C:\Docs\PDFs\A-1001.pdf","12"
"Invoice-2026-07.pdf","C:\Docs\PDFs\Invoices\Invoice-2026-07.pdf","2"
"ScannedPacket.pdf","C:\Docs\PDFs\ScannedPacket.pdf","Unknown"
```

## Notes

- Output files are written to the folder you pass to `-Folder`.
- Filenames include a timestamp to avoid overwriting previous runs.
- If `pdfinfo` is unavailable, extraction still runs with fallback strategies.
- After running `Setup.ps1`, restart your PowerShell session so `PATH` updates are recognized.

## Troubleshooting

`pdfinfo` not found:
- Run `Setup.ps1`.
- If needed, manually install Poppler and ensure the folder containing `pdfinfo.exe` is in your user `PATH`.

No PDFs discovered:
- Verify the folder path passed to `-Folder` exists and contains `.pdf` files.
- The script searches recursively.
