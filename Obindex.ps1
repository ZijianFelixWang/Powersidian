<#
.SYNOPSIS
Generates Obsidian vault index with automatic linking and organization

.DESCRIPTION
This script automates the creation of index files in an Obsidian vault. Key features include:
- Automatic vault backup with GUID-named directories
- Multi-threaded file copying for fast backups
- Recursive directory processing with assets folder exclusion
- Dual operation modes (standard/book style)
- Progress tracking with nested progress bars
- Obsidian-compatible link generation
- Automatic timestamp callouts
- Create or update portal file with homepage links
- Calculate statistics of definitions/theorems/etc. and write them to a .md file
- Generate Powersidian configuration for Knowledge, Lecture Notes, and Appendices
- Add metadata info block to notes
- Reopen Obsidian after processing (optional)

.PARAMETER VaultPath
Specifies the path to the Obsidian vault directory. Must be a valid directory path.

.PARAMETER Book
When specified, enables book-style formatting with:
- Chapter numbers in file headers
- Hierarchical section numbering (e.g. §1.2.3)
- Structured heading levels

.PARAMETER ObsidianShortcut
When specified, opens Obsidian before processing and waits for the user to close it before continuing.

.PARAMETER InvokePowersidian
When specified, opens Powersidian after processing to export vault into PDF.

.PARAMETER SkipPowerCheck
When specified, Obindex will be forced to run on battery. Not recommended.

.EXAMPLE
PS> .\Obindex.ps1 -VaultPath "C:\MyVault"

Standard mode: Creates basic index files with direct links

.EXAMPLE
PS> .\Obindex.ps1 -VaultPath "C:\MyVault" -Book

Book mode: Generates numbered chapters and sections with § notation

.EXAMPLE
PS> .\Obindex.ps1 -VaultPath "C:\MyVault" -ObsidianShortcut

This will open Obsidian and wait for the user to close it before proceeding.

.EXAMPLE
PS> .\Obindex.ps1 -VaultPath "C:\MyVault" -InvokePowersidian

This will open Powersidian after processing to export the vault into PDF.

.EXAMPLE
PS> .\Obindex.ps1 -VaultPath "C:\MyVault" -SkipPowerCheck

This will force Obindex to run on battery. Not recommended.

.INPUTS
None. You cannot pipe input to this script.

.OUTPUTS
Creates/updates markdown files in the vault directory structure.

.NOTES
Copyright (C) 2025 Nuaptan. All rights reserved.
This program is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

This script requires Robocopy and PowerShell 7.5 or later.

#>

param(
    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path $_ -PathType Container })]
    [string]$VaultPath,

    [switch]$Book,

    [switch]$ObsidianShortcut,

    [switch]$InvokePowersidian,

    [switch]$SkipPowerCheck
)

#region Helper Function 1
function Get-SectionNumber {
    param(
        [int]$CurrentLevel,
        [ref]$Counters
    )
    
    # Ensure array is large enough
    while ($Counters.Value.Count -le $CurrentLevel) {
        $Counters.Value += 0
    }
    
    # Increment current level counter
    $Counters.Value[$CurrentLevel]++
    
    # Reset deeper levels
    for ($i = $CurrentLevel + 1; $i -lt $Counters.Value.Count; $i++) {
        $Counters.Value[$i] = 0
    }
    
    # Build section string
    $section = @()
    for ($i = 0; $i -le $CurrentLevel; $i++) {
        if ($Counters.Value[$i] -gt 0) {
            $section += $Counters.Value[$i]
        }
    }
    
    return $section -join '.'
}
#endregion

#region Helper Function 2
# Function to get the size of a folder in MB
function Get-FolderSizeMB {
    param (
        [string]$FolderPath
    )
    # Get the folder size in bytes
    $FolderSizeBytes = (Get-ChildItem -Path $FolderPath -Recurse -File | Measure-Object -Property Length -Sum).Sum
    # Convert bytes to MB
    $FolderSizeMB = [Math]::Round($FolderSizeBytes / 1MB, 2)
    return $FolderSizeMB
}

# Function to delete the oldest subfolder
function Remove-OldestSubfolder {
    param (
        [string]$FolderPath
    )
    # Get all subfolders sorted by creation time (oldest first)
    $Subfolders = Get-ChildItem -Path $FolderPath -Directory | Sort-Object CreationTime

    # Check if there are any subfolders
    if ($Subfolders.Count -gt 0) {
        # Get the oldest subfolder
        $OldestSubfolder = $Subfolders[0]

        # Remove the oldest subfolder
        try {
            Remove-Item -Path $OldestSubfolder.FullName -Recurse -Force
            Write-Host "Removed oldest subfolder: $($OldestSubfolder.Name)"
        }
        catch {
            Write-Host "Error removing subfolder $($OldestSubfolder.Name): $($_.Exception.Message)"
        }
    }
    else {
        Write-Host "No subfolders found in $($FolderPath)."
    }
}

function Remove-OldBackups {
    param (
        [string]$FolderPath,
        [int]$SizeThresholdGB,
        [int]$TargetSizeMB,
        [int]$MinSubfolders
    )
    Write-Host "Starting folder monitoring for $($FolderPath)..."

    # Get the folder size in MB
    $FolderSizeMB = Get-FolderSizeMB -FolderPath $FolderPath

    # Convert the size threshold from GB to MB
    $SizeThresholdMB = $SizeThresholdGB * 1024

    # Get the number of subfolders
    $SubfolderCount = (Get-ChildItem -Path $FolderPath -Directory).Count

    # Check if the folder size exceeds the threshold
    if ($FolderSizeMB -gt $SizeThresholdMB) {
        Write-Host "Folder size $($FolderSizeMB) MB exceeds the threshold of $($SizeThresholdMB) MB."
        Write-Host "Starting cleanup process..."

        # Loop until the folder size is below the target or the minimum number of subfolders is reached
        while (($FolderSizeMB -gt $TargetSizeMB) -and ($SubfolderCount -gt $MinSubfolders)) {
            # Remove the oldest subfolder
            Remove-OldestSubfolder -FolderPath $FolderPath

            # Update the folder size
            $FolderSizeMB = Get-FolderSizeMB -FolderPath $FolderPath

            # Update the subfolder count
            $SubfolderCount = (Get-ChildItem -Path $FolderPath -Directory).Count

            Write-Host "Current folder size: $($FolderSizeMB) MB. Subfolder count: $($SubfolderCount)"
        }

        Write-Host "Cleanup process completed."
        Write-Host "Final folder size: $($FolderSizeMB) MB. Final subfolder count: $($SubfolderCount)"
    }
    else {
        Write-Host "Folder size $($FolderSizeMB) MB is within the threshold of $($SizeThresholdMB) MB."
    }
}
#endregion

#region Helper Function 3
function ConvertTo-Roman {
    param(
        [Parameter(Mandatory)]
        [int]$Number
    )
    
    if ($Number -lt 1 -or $Number -gt 3999) {
        throw "Number must be between 1 and 3999."
    }
    
    $roman = ""
    
    $map = @(
        @{Value = 1000; Symbol = "M" },
        @{Value = 900; Symbol = "CM" },
        @{Value = 500; Symbol = "D" },
        @{Value = 400; Symbol = "CD" },
        @{Value = 100; Symbol = "C" },
        @{Value = 90; Symbol = "XC" },
        @{Value = 50; Symbol = "L" },
        @{Value = 40; Symbol = "XL" },
        @{Value = 10; Symbol = "X" },
        @{Value = 9; Symbol = "IX" },
        @{Value = 5; Symbol = "V" },
        @{Value = 4; Symbol = "IV" },
        @{Value = 1; Symbol = "I" }
    )
    
    foreach ($entry in $map) {
        while ($Number -ge $entry.Value) {
            $roman += $entry.Symbol
            $Number -= $entry.Value
        }
    }
    
    return $roman
}

function Remove-SpaceBeforeRoman {
    param(
        [Parameter(Mandatory)]
        [string]$InputString
    )

    if ($InputString -match "^(.*)\s([IVXLCDM]+)$") {
        return "$($matches[1])$($matches[2])"
    }
    else {
        return $InputString
    }
}
#endregion

#region Initial Checks and Setup
Write-Host "Obindex. Do not close this window."
Write-Host ""

if (-not (Join-Path -Path $VaultPath -ChildPath "Notes Root\Knowledge" | Test-Path -PathType Container)) {
    Write-Error "Invalid vault path. Ensure the directory structure is correct."
    if ($ObsidianShortcut) {
        Pause
    }
    exit 1
}
if (-not (Join-Path -Path $VaultPath -ChildPath "Notes Root\Portals" | Test-Path -PathType Container)) {
    Write-Error "Invalid vault path. Ensure the directory structure is correct."
    if ($ObsidianShortcut) {
        Pause
    }
    exit 1
}

# Check for Obsidian process
if (-not $ObsidianShortcut) {
    $obsidianProcess = Get-Process obsidian -ErrorAction SilentlyContinue
    $reopenObsidian = $null -ne $obsidianProcess
    if ($obsidianProcess) {
        Write-Host "Stopping Obsidian process..."
        $obsidianProcess | Stop-Process -Force
    }
}
else {
    if (-not (Join-Path -Path $env:LOCALAPPDATA -ChildPath "Programs\obsidian\Obsidian.exe" | Test-Path)) {
        Write-Error "Obsidian not found. Ensure the application is installed."
        Pause
        exit 1
    }
    Start-Process -FilePath (Join-Path -Path $env:LOCALAPPDATA -ChildPath "Programs\obsidian\Obsidian.exe") -RedirectStandardOutput "Temp:\$([Guid]::NewGuid().ToString())" -Wait
}

# AC power check
if (-not $SkipPowerCheck) {
    if ((Get-WmiObject -Class Win32_Battery).BatteryStatus -eq 1) {
        Write-Host "Obindex will not run because AC cable is not connected."
        if ($ObsidianShortcut) {
            Start-Sleep -Seconds 1
        }
        Exit
    }
}

# Create backup
$backupParent = Join-Path ([Environment]::GetFolderPath('MyDocuments')) "ObsidianBackups"
$tempBackup = Join-Path $backupParent ([Guid]::NewGuid().ToString())

Remove-OldBackups -FolderPath $backupParent -SizeThresholdGB 2 -TargetSizeMB 512 -MinSubfolders 4

try {
    New-Item -Path $tempBackup -ItemType Directory -Force | Out-Null
    Write-Host "Backing up vault to: $tempBackup"
    robocopy $VaultPath $tempBackup /MIR /MT:16 /R:3 /W:10 /NP /NFL /NDL /NJH /NJS
    Write-Host "Backup completed successfully. Location: $tempBackup"
}
catch {
    Write-Error "Backup failed: $_"
    exit 1
}

Write-Host "`nPlease do not open Obsidian or modify the vault during processing.`n"
#endregion

#region Directory Processing
$indexedFiles = @()
$homepageNames = @()
$allDirs = Get-ChildItem (Join-Path -Path $VaultPath -ChildPath "Notes Root\Knowledge") -Directory -Recurse | 
Where-Object {
    $_.FullName -notmatch '\\assets\\?' -and 
        (Get-ChildItem $_.FullName -File -Filter *.md)
}

$dirCount = $allDirs.Count
$dirIndex = 0

foreach ($dir in $allDirs) {
    $dirIndex++
    $dirProgress = @{
        Id              = 0
        Activity        = "Processing Directories"
        Status          = "Processing $($dir.Name) ($dirIndex/$dirCount)"
        PercentComplete = ($dirIndex / $dirCount) * 100
    }
    Write-Progress @dirProgress
    
    # Create homepage file
    $homepage = Join-Path $dir.FullName "$($dir.Name) Homepage.md"
    if (Test-Path $homepage) {
        Remove-Item $homepage -Force 
    }
    New-Item -Path $homepage -ItemType File -Force | Out-Null
    
    $homepageNames += "$($dir.Name) Homepage"

    $header = @"
<!-- automatically generated -->
# $($dir.Name)

"@
    $header | Out-File $homepage -Encoding utf8
    $indexedFiles += $homepage
    
    # Process MD files
    $files = Get-ChildItem $dir.FullName -Filter *.md |
    Where-Object { 
        $_.Name -ne "$($dir.Name) Homepage.md" -and
        -not (Select-String -Path $_.FullName -Pattern '<!-- automatically generated -->')
    } |
    Sort-Object CreationTime
    
    $fileCount = $files.Count
    $fileIndex = 0
    
    foreach ($file in $files) {
        $fileIndex++
        $fileProgress = @{
            Id              = 1
            Activity        = "Processing Files"
            Status          = "Processing $($file.Name) ($fileIndex/$fileCount)"
            PercentComplete = ($fileIndex / $fileCount) * 100
        }
        Write-Progress @fileProgress
        
        # Add file section
        $chapterPrefix = if ($Book) { "Chapter $fileIndex " } else { "" }
        Add-Content $homepage "## ${chapterPrefix}$($file.BaseName)"
        
        # Process headings
        $content = Get-Content $file.FullName
        $counters = [ref]@()
        $currentLevel = 0
        
        foreach ($line in $content) {
            if ($line -match '^(?<level>#+)\s+(?<title>.+)$') {
                $headingLevel = $matches.level.Length - 1  # Original heading level (0-based)
                
                # Calculate section number
                $sectionNumber = if ($Book) {
                    $sn = Get-SectionNumber -CurrentLevel $headingLevel -Counters $counters
                    "§$fileIndex.$sn "
                }
                else { "" }
                
                # Create link
                $newHeadingLevel = $matches.level.Length + 2
                $headingText = $matches.title.Trim()
                $link = "$('#' * $newHeadingLevel) ${sectionNumber}$headingText [[$($file.BaseName)#$headingText|→]]"
                
                Add-Content $homepage "$link"
            }
        }
        Add-Content $homepage ""
    }
    
    # Add timestamp callout
    $timestamp = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
    Add-Content $homepage @"
> [!info]
> Index generated at $timestamp.
"@
}

Write-Progress -Activity "Processing Files" -Completed
Write-Progress -Activity "Processing Directories" -Completed
#endregion

#region Generate Portal
Write-Host "Indexing complete. Created/updated files:"
$indexedFiles | ForEach-Object { Write-Host "- $_" }
Write-Host ""

$portal = Join-Path -Path $VaultPath -ChildPath "Notes Root\Portals\Portal → Knowledge Base.md"
if (Test-Path $portal) {
    Remove-Item $portal -Force 
}
New-Item -Path $portal -ItemType File -Force | Out-Null

Add-Content $portal "<!-- automatically generated -->"
foreach ($page in $homepageNames) {
    Add-Content $portal "![[$page]]"
}
If (Join-Path -Path $VaultPath -ChildPath "Notes Root\Portals\Portal → Knowledge Base Special.md" | Test-Path) {
    Add-Content $portal "![[Portal → Knowledge Base Special]]"
}
Add-Content $portal ""
$timestamp = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
Add-Content $portal @"
> [!info]
> Index generated at $timestamp.
"@

Write-Host "Portal updated: $portal"
Write-Host ""
#endregion

#region Generate Statistics
# Initialize counters
$counts = [ordered]@{
    "Definitions"  = 0
    "Theorems"     = 0
    "Lemmas"       = 0
    "Propositions" = 0
    "Corollaries"  = 0
    "Examples"     = 0
    "Cautions"     = 0
    "Questions"    = 0
    "Axioms"       = 0
}

# Create regex patterns for each type
$patterns = @{
    "Definitions"  = '^>\s*\[\!abstract\]\s+Definition\b'
    "Theorems"     = '^>\s*\[\!note\]\s+Theorem\b'
    "Lemmas"       = '^>\s*\[\!note\]\s+Lemma\b'
    "Propositions" = '^>\s*\[\!note\]\s+Proposition\b'
    "Corollaries"  = '^>\s*\[\!note\]\s+Corollary\b'
    "Examples"     = '^>\s*\[\!example\]\s+Example\b'
    "Cautions"     = '^>\s*\[\!caution\]\s+Caution\b'
    "Questions"    = '^>\s*\[\!question\]\s+Question\b'
    "Axioms"       = '^>\s*\[\!abstract\]\s+Axiom\b'
}

# Compile regex objects
$regexes = @{}
foreach ($key in $patterns.Keys) {
    $regexes[$key] = [regex]$patterns[$key]
}

# Process all markdown files

$notes = Get-ChildItem -Path $vaultPath -Recurse -Filter *.md
$noteCount = $notes.Count
$noteIndex = 0

$notes | ForEach-Object {
    $noteIndex++
    $noteProgress = @{
        Id              = 1
        Activity        = "Calculating Statistics for Notes"
        Status          = "Processing $($_.Name) ($noteIndex/$noteCount)"
        PercentComplete = ($noteIndex / $noteCount) * 100
    }
    Write-Progress @noteProgress

    $content = Get-Content $_.FullName -Raw
    foreach ($line in $content -split "`r?`n") {
        foreach ($key in $counts.Keys) {
            if ($regexes[$key].IsMatch($line)) {
                $counts[$key]++
                break
            }
        }
    }
}
Write-Progress -Activity "Calculating Statistics for Notes" -Completed
Write-Progress -Id 1 -Completed

# Create output path
$outputDir = Join-Path $vaultPath "Notes Root\Automatic Files"
$outputPath = Join-Path $outputDir "Statistics.md"

# Delete existing file if exists
if (Test-Path $outputPath) { Remove-Item $outputPath -Force }

# Create directory if needed
if (-not (Test-Path $outputDir)) { New-Item -Path $outputDir -ItemType Directory | Out-Null }

# Create markdown table
$table = @"
# Vault Statistics
| Type          | Count |
|---------------|-------|
| Definitions   | $($counts.Definitions) |
| Theorems      | $($counts.Theorems) |
| Lemmas        | $($counts.Lemmas) |
| Propositions  | $($counts.Propositions) |
| Corollaries   | $($counts.Corollaries) |
| Examples      | $($counts.Examples) |
| Cautions      | $($counts.Cautions) |
| Questions     | $($counts.Questions) |
| Axioms        | $($counts.Axioms) |
| **Total**         | **$(($counts.Values | Measure-Object -Sum).Sum)** |

> [!info]
> Statistics generated at $(Get-Date -Format "MM/dd/yyyy HH:mm:ss")

"@

# Write to file
$table | Out-File -FilePath $outputPath -Encoding utf8

Write-Host "Statistics generated at: $outputPath"
Write-Host ""
#endregion

#region Generate Metaddata for Notes
$kfiles = Get-ChildItem (Join-Path -Path $VaultPath -ChildPath "Notes Root\Knowledge") -Recurse -Filter *.md -Exclude *Homepage*
$afiles = Get-ChildItem (Join-Path -Path $VaultPath -ChildPath "Notes Root\Lecture Notes") -Recurse -Filter *.md -Exclude *Homepage*, "*Assignments Pad*" -Include *Assignments*
$lfiles = Get-ChildItem (Join-Path -Path $VaultPath -ChildPath "Notes Root\Lecture Notes") -Recurse -Filter *.md -Exclude *Homepage* -Include "*Lecture Notes*"

$totalFiles = $kfiles.Length + $afiles.Length + $lfiles.Length
$currentFile = 0

$kfiles | ForEach-Object {
    $currentFile++
    $metadataProgress = @{
        Id              = 1
        Activity        = "Adding Metadata to Notes (1/3)"
        Status          = "Processing $($_.Name) ($currentFile/$totalFiles)"
        PercentComplete = ($currentFile / $totalFiles) * 100
    }
    Write-Progress @metadataProgress

    if ((Get-Content $_).Contains("<!-- End of Obindex Metadata -->")) {
        $endIndex = ((Get-Content $_) | Select-String -Pattern "<!-- End of Obindex Metadata -->").LineNumber
        $x = Get-Content $_ 
        $y = $x | Select-Object -Index ($endIndex..$x.Length) -ErrorAction SilentlyContinue
        $y | Set-Content $_
    }

    $cannName = "K_$($_.BaseName.Replace(' ', '_'))"
    $banner = @"
> [!info] Metadata
> Canonical Name: ``$cannName``
> Created Time:   $($_.CreationTime.ToString())
> Modified Time:  $($_.LastWriteTime.ToString())
> Indexed Time:   $((Get-Date).ToString())
> This metadata is generated by script Obindex.ps1 locally.

<!-- End of Obindex Metadata -->
"@
    $banner, (Get-Content $_) | Set-Content $_
    Start-Sleep -Milliseconds 50
}

$lfiles | ForEach-Object {
    $currentFile++
    $metadataProgress = @{
        Id              = 1
        Activity        = "Adding Metadata to Notes (2/3)"
        Status          = "Processing $($_.Name) ($currentFile/$totalFiles)"
        PercentComplete = ($currentFile / $totalFiles) * 100
    }
    Write-Progress @metadataProgress

    if ((Get-Content $_).Contains("<!-- End of Obindex Metadata -->")) {
        $endIndex = ((Get-Content $_) | Select-String -Pattern "<!-- End of Obindex Metadata -->").LineNumber
        # Write-Debug "x1/y1 processing: $($_.Name)"
        $x1 = Get-Content $_ 
        $y1 = $x1 | Select-Object -Index ($endIndex..$x1.Length) -ErrorAction SilentlyContinue
        $y1 | Set-Content $_
    }

    $cannName = "L_$($_.BaseName.Replace(' ', '_'))"
    $banner = @"
> [!info] Metadata
> Canonical Name: ``$cannName``
> Created Time:   $($_.CreationTime.ToString())
> Modified Time:  $($_.LastWriteTime.ToString())
> Indexed Time:   $((Get-Date).ToString())
> This metadata is generated by script Obindex.ps1 locally.

<!-- End of Obindex Metadata -->
"@
    $banner, (Get-Content $_) | Set-Content $_
    Start-Sleep -Milliseconds 50
}

$afiles | ForEach-Object {
    $currentFile++
    $metadataProgress = @{
        Id              = 1
        Activity        = "Adding Metadata to Notes (3/3)"
        Status          = "Processing $($_.Name) ($currentFile/$totalFiles)"
        PercentComplete = ($currentFile / $totalFiles) * 100
    }
    Write-Progress @metadataProgress

    if ((Get-Content $_).Contains("<!-- End of Obindex Metadata -->")) {
        $endIndex = ((Get-Content $_) | Select-String -Pattern "<!-- End of Obindex Metadata -->").LineNumber
        $x2 = Get-Content $_ 
        $y2 = $x2 | Select-Object -Index ($endIndex..$x2.Length) -ErrorAction SilentlyContinue
        $y2 | Set-Content $_
    }

    $cannName = "A_$($_.BaseName.Replace(' ', '_'))"
    $banner = @"
> [!info] Metadata
> Canonical Name: ``$cannName``
> Created Time:   $($_.CreationTime.ToString())
> Modified Time:  $($_.LastWriteTime.ToString())
> Indexed Time:   $((Get-Date).ToString())
> This metadata is generated by script Obindex.ps1 locally.

<!-- End of Obindex Metadata -->
"@
    $banner, (Get-Content $_) | Set-Content $_
    Start-Sleep -Milliseconds 50
}

Write-Progress -Id 1 -Completed
Write-Host "Done adding metadata to notes."
#endregion

#region Write Powersidian Configuration
$configPath = Join-Path -Path $VaultPath -ChildPath "assets\powersidian.txt"
if (Test-Path $configPath) {
    Remove-Item $configPath -Force
}
New-Item -Path $configPath -ItemType File -Force | Out-Null
Add-Content $configPath "Notes Root/Support Files/Powersidian/Noton soez Nuaptan"

# Generate configuration for Knowledge subfolder
$knowledgeParts = (Get-ChildItem (Join-Path -Path $VaultPath -ChildPath "Notes Root\Knowledge") -Directory | Sort-Object CreationTime).Name
for ($i = 0; $i -lt $knowledgeParts.Count; $i++) {
    $knowledgeProgress = @{
        Id              = 1
        Activity        = "Generating Powersidian Configuration for Knowledge Base"
        Status          = "Processing $($knowledgeParts[$i]) ($($i + 1)/$($knowledgeParts.Count))"
        PercentComplete = (($i + 1) / $knowledgeParts.Count) * 100
    }
    Write-Progress @knowledgeProgress

    $partFileName = Join-Path -Path $VaultPath -ChildPath "Notes Root\Support Files\Powersidian\Part $(ConvertTo-Roman -Number ($i + 1)).md"
    if (Test-Path $partFileName) {
        Remove-Item $partFileName -Force
    }
    New-Item -Path $partFileName -ItemType File -Force | Out-Null

    if (Join-Path -Path $VaultPath -ChildPath "Notes Root\Knowledge\$($knowledgeParts[$i])\$($knowledgeParts[$i]) Homepage.md" | Test-Path) {
        Add-Content $partFileName "![[$($knowledgeParts[$i]) Homepage#$($knowledgeParts[$i])]]"
    }
    else {
        Add-Content $partFileName "# $($knowledgeParts[$i])"
    }
    Add-Content $partFileName ""

    Add-Content $configPath "Notes Root/Support Files/Powersidian/Part$(ConvertTo-Roman -Number ($i + 1))"

    $partFiles = Get-ChildItem (Join-Path -Path $VaultPath -ChildPath "Notes Root\Knowledge\$($knowledgeParts[$i])") -File -Filter *.md | Sort-Object CreationTime
    $partFiles | ForEach-Object -Process {
        if ($_.BaseName.Contains("Homepage")) {
            continue
        }
        Add-Content $configPath "Notes Root/Knowledge/$($knowledgeParts[$i])/$(Remove-SpaceBeforeRoman -InputString $_.BaseName)"   
    }
}

Write-Progress -Activity "Generating Powersidian Configuration for Knowledge Base" -Completed
Write-Progress -Id 1 -Completed

# Generate configuration for Lecture Notes subfolder
$lectureParts = (Get-ChildItem (Join-Path -Path $VaultPath -ChildPath "Notes Root\Lecture Notes") -Directory | Sort-Object CreationTime).Name
for ($i = 0; $i -lt $lectureParts.Count; $i++) {
    $lectureProgress = @{
        Id              = 1
        Activity        = "Generating Powersidian Configuration for Lecture Notes"
        Status          = "Processing $($lectureParts[$i]) ($($i + 1)/$($lectureParts.Count))"
        PercentComplete = (($i + 1) / $lectureParts.Count) * 100
    }
    Write-Progress @lectureProgress

    $partFileName = Join-Path -Path $VaultPath -ChildPath "Notes Root\Support Files\Powersidian\Part $(ConvertTo-Roman -Number ($i + $knowledgeParts.Count + 1)).md"
    if (Test-Path $partFileName) {
        Remove-Item $partFileName -Force
    }
    New-Item -Path $partFileName -ItemType File -Force | Out-Null

    $lecHomepage = (Get-ChildItem (Join-Path -Path $VaultPath -ChildPath "Notes Root\Lecture Notes\$($lectureParts[$i])") -Filter "*Homepage.md")
    if ($lecHomepage) {
        Add-Content $partFileName "# Lecture Notes - $($lectureParts[$i])"
        Add-Content $partFileName "![[$($lecHomepage[0].BaseName)]]"
    }
    else {
        Add-Content $partFileName "# Lecture Notes - $($lectureParts[$i])"
    }

    Add-Content $partFileName ""

    Add-Content $configPath "Notes Root/Support Files/Powersidian/Part$(ConvertTo-Roman -Number ($i + $knowledgeParts.Count + 1))"

    if ($lectureParts[$i].Contains("Science Lectures")) {
        continue
    }

    $partFiles = Get-ChildItem (Join-Path -Path $VaultPath -ChildPath "Notes Root\Lecture Notes\$($lectureParts[$i])") -File -Filter *.md | Sort-Object CreationTime
    $partFiles | ForEach-Object -Process {
        if (-not($_.BaseName.Contains("Homepage") -or $_.BaseName.Contains("REV"))) {
            Add-Content $configPath "Notes Root/Lecture Notes/$($lectureParts[$i])/$(Remove-SpaceBeforeRoman -InputString $_.BaseName)"   
        }
    }
}

Write-Progress -Activity "Generating Powersidian Configuration for Lecture Notes" -Completed
Write-Progress -Id 1 -Completed

# Generate configuration for Appendices
$i = $knowledgeParts.Count + $lectureParts.Count + 1
$partFileName = Join-Path -Path $VaultPath -ChildPath "Notes Root\Support Files\Powersidian\Part $(ConvertTo-Roman -Number $i).md"
if (Test-Path $partFileName) {
    Remove-Item $partFileName -Force
}
New-Item -Path $partFileName -ItemType File -Force | Out-Null
Add-Content $partFileName "# Appendices"
Add-Content $partFileName ""
Add-Content $configPath "Notes Root/Support Files/Powersidian/Part$(ConvertTo-Roman -Number $i)"
Add-Content $configPath "Notes Root/Portals/Portal → Portals"
Add-Content $configPath "Notes Root/Portals/Portal → Special Pages"
Add-Content $configPath "Notes Root/Portals/Portal → Curricula"
Add-Content $configPath "Notes Root/Support Files/Book Titles"

Write-Host ""
Write-Host "Powersidian configuration generated at: $configPath"
Write-Host "Powersidian configuration file size: $((Get-Item $configPath).Length) bytes."

if ($InvokePowersidian) {
    Write-Host "Initializing Obsidian."
    Invoke-Command -ScriptBlock {
        Write-Host ""
        .\Powersidian.ps1 -NotesFilePath $configPath -ObindexMode
    }
}
#endregion

#region Reopen Obsidian
if (-not $ObsidianShortcut) {
    if ($reopenObsidian -and (Join-Path -Path $env:LOCALAPPDATA -ChildPath "Programs\obsidian\Obsidian.exe" | Test-Path)) {
        Start-Process -FilePath (Join-Path -Path $env:LOCALAPPDATA -ChildPath "Programs\obsidian\Obsidian.exe") -RedirectStandardOutput "Temp:\$([Guid]::NewGuid().ToString())"
    }
}
#endregion