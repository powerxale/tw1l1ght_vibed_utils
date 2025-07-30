#Requires -Version 5.1
<#
.SYNOPSIS
    Converts all PowerPoint (.pptx, .ppt) files in a specified folder to PDF.

.DESCRIPTION
    This script automates the process of converting PowerPoint presentations to PDF format.
    It requires Microsoft PowerPoint to be installed on the system. The script will
    create a 'PDF' subdirectory within the source folder to store the converted files.

.PARAMETER SourceFolder
    The full path to the folder containing the .pptx and .ppt files to be converted.

.EXAMPLE
    .\Convert-PPTXtoPDF.ps1 -SourceFolder "C:\Users\YourUser\Desktop\Presentations"
#>
param (
    [Parameter(Mandatory=$true, HelpMessage="Enter the path to the folder with PowerPoint files.")]
    [string]$SourceFolder
)

# --- Main Script ---

# Check if the source folder exists
if (-not (Test-Path -Path $SourceFolder -PathType Container)) {
    Write-Error "Error: The source folder was not found: $SourceFolder"
    exit
}

# Define the output directory for PDFs
$outputFolder = Join-Path -Path $SourceFolder -ChildPath "PDF"

# Create the output directory if it doesn't exist
if (-not (Test-Path -Path $outputFolder -PathType Container)) {
    try {
        New-Item -Path $outputFolder -ItemType Directory -ErrorAction Stop | Out-Null
        Write-Host "Successfully created output directory: $outputFolder" -ForegroundColor Green
    } catch {
        Write-Error "Error: Could not create output directory: $outputFolder. Please check permissions."
        exit
    }
}

# Check if PowerPoint is installed by trying to create the COM object
try {
    $powerPoint = New-Object -ComObject PowerPoint.Application -ErrorAction Stop
} catch {
    Write-Error "Error: Microsoft PowerPoint is not installed or cannot be started. Please install PowerPoint to use this script."
    exit
}

# Get all PowerPoint files in the source folder (non-recursive)
$powerPointFiles = Get-ChildItem -Path $SourceFolder -Filter "*.ppt*" | Where-Object { $_.Extension -in ".ppt", ".pptx" }

if ($powerPointFiles.Count -eq 0) {
    Write-Warning "No PowerPoint files (.ppt, .pptx) were found in the specified folder."
    $powerPoint.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($powerPoint) | Out-Null
    exit
}

Write-Host "Found $($powerPointFiles.Count) PowerPoint file(s) to convert."

# Define the PDF format type from the PowerPoint object model
$pdfFormat = 32 # Corresponds to ppSaveAsPDF

foreach ($file in $powerPointFiles) {
    $outputFileName = Join-Path -Path $outputFolder -ChildPath "$($file.BaseName).pdf"
    Write-Host "Converting '$($file.Name)'..." -ForegroundColor Yellow

    # Check if a PDF with the same name already exists
    if (Test-Path -Path $outputFileName) {
        Write-Host "  -> PDF file already exists. Skipping." -ForegroundColor Cyan
        continue
    }

    $presentation = $null
    try {
        # Open the presentation
        $presentation = $powerPoint.Presentations.Open($file.FullName, $true, $false, $false) # (FileName, ReadOnly, Untitled, WithWindow)

        # Save the presentation as PDF
        $presentation.SaveAs($outputFileName, $pdfFormat)

        Write-Host "  -> Successfully converted to '$($outputFileName)'" -ForegroundColor Green
    } catch {
        Write-Error "  -> Failed to convert '$($file.Name)'. It might be open, password-protected, or corrupted."
        Write-Error $_.Exception.Message
    } finally {
        # Close the presentation
        if ($presentation -ne $null) {
            $presentation.Close()
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($presentation) | Out-Null
        }
    }
}

# Quit PowerPoint and release the COM object
$powerPoint.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($powerPoint) | Out-Null

# Force garbage collection to ensure PowerPoint process terminates
[gc]::Collect()
[gc]::WaitForPendingFinalizers()

Write-Host "All conversions complete." -ForegroundColor Green
