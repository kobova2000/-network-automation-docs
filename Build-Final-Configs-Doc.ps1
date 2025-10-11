<#
===============================================================================
Script Name: Build-Final-Configs-Doc.ps1
Author: Frank Abraham
Version: 1.2
License: MIT License

Description:
    Compiles multiple Cisco switch configuration text files into a single,
    professionally formatted Microsoft Word document.

    This script represents Phase 2 of the automation workflow:
        1. get-switch-configs.py  →  collects configs from devices
        2. Build-Final-Configs-Doc.ps1  →  formats and compiles configs for review

Key Features:
    • 30 spaces between hostname and IP
    • Removes all lines containing "!"
    • Single-line spacing (no extra before/after)
    • Runs Word invisibly (no GUI)
    • Console-only status output

Usage:
    1. Place all configuration .txt files in the "final-configs" folder.
    2. Run the script:
        .\Build-Final-Configs-Doc.ps1 -OutputDoc "C:\Path\To\Your\Scripts\final-configs.docx"

Dependencies:
    • Microsoft Word (for COM automation)
    • Windows PowerShell 5.1 or later
===============================================================================
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $false)]
    [string]$OutputDoc = "C:\Path\To\Your\Scripts\final-configs.docx"
)

Write-Host "Starting Build-Final-Configs-Doc..." -ForegroundColor Cyan

# ---------------------------------------------------------------------------
# Folder setup
# ---------------------------------------------------------------------------
$configPath = "C:\Path\To\Your\Scripts\final-configs"

if (-not (Test-Path $configPath)) {
    Write-Host "Error: Config path not found: $configPath" -ForegroundColor Red
    exit 1
}

$configFiles = Get-ChildItem -Path $configPath -Filter *.txt
if (-not $configFiles) {
    Write-Host "No configuration files found in $configPath" -ForegroundColor Red
    exit 1
}

Write-Host "Found $($configFiles.Count) configuration files." -ForegroundColor Yellow

# ---------------------------------------------------------------------------
# Create Word COM object
# ---------------------------------------------------------------------------
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $doc = $word.Documents.Add()
}
catch {
    Write-Host "Error: Could not create Microsoft Word COM object. Is Word installed?" -ForegroundColor Red
    exit 1
}

# ---------------------------------------------------------------------------
# Process each configuration file
# ---------------------------------------------------------------------------
foreach ($file in $configFiles) {
    Write-Host "Processing $($file.Name)..." -ForegroundColor Cyan

    try {
        # Read and clean configuration
        $content = Get-Content -Path $file.FullName -ErrorAction Stop |
                   Where-Object { $_ -notmatch '^!' }   # Remove lines containing "!"

        # Extract hostname
        $hostname = ($content | Select-String -Pattern '^hostname\s+(.+)$' |
                     ForEach-Object { $_.Matches[0].Groups[1].Value })

        if (-not $hostname) { $hostname = "UNKNOWN-HOSTNAME" }

        # Prepare header text with 30 spaces between hostname and IP
        $ip = $file.BaseName
        $headerText = ("{0,-1}" -f $hostname) + (" " * 30) + $ip

        # Insert page break except for first file
        if ($file -ne $configFiles[0]) {
            $doc.Content.InsertBreak([Microsoft.Office.Interop.Word.WdBreakType]::wdPageBreak)
        }

        # Insert header
        $header = $doc.Content.Paragraphs.Add()
        $header.Range.Text = $headerText
        $header.Range.Font.Size = 12
        $header.Range.Font.Bold = $true
        $header.Range.SpaceBefore = 0
        $header.Range.SpaceAfter = 0
        $header.Range.LineSpacingRule = 0
        $header.Range.InsertParagraphAfter()

        # Insert configuration text
        $para = $doc.Content.Paragraphs.Add()
        $para.Range.Text = ($content -join "`r`n")
        $para.Range.Font.Size = 10
        $para.Range.Font.Bold = $false
        $para.Range.SpaceBefore = 0
        $para.Range.SpaceAfter = 0
        $para.Range.LineSpacingRule = 0
        $para.Range.InsertParagraphAfter()
    }
    catch {
        Write-Host "Error processing file: $($file.Name). $_" -ForegroundColor Red
    }
}

# ---------------------------------------------------------------------------
# Save and close Word document
# ---------------------------------------------------------------------------
try {
    Write-Host "Saving document to $OutputDoc ..." -ForegroundColor Yellow
    $doc.SaveAs([ref]$OutputDoc, [ref]16)  # 16 = wdFormatDocumentDefault (.docx)
    $doc.Close()
    $word.Quit()
    Write-Host "Document successfully saved: $OutputDoc" -ForegroundColor Green
}
catch {
    Write-Host "Error while saving document: $_" -ForegroundColor Red
    if ($word) { $word.Quit() }
    exit 1
}

Write-Host "Build-Final-Configs-Doc completed." -ForegroundColor Green
