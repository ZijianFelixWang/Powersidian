# PowersidianVM = Power Obsidian VM Host Script
# Copyright 2024 Nuaptan.
# 
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
# 
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
# 
# You should have received a copy of the GNU General Public License
# along with this program.  If not, see <https://www.gnu.org/licenses/>.

[CmdletBinding()]
param (
    [string]$ExportFolderPath = "E:\Powersidian-Temp", # Change this to your desired export folder path
    [string]$Source = "E:\Notes", # Change this to your notes folder path
    [string]$Destination = $(Join-Path -Path $ExportFolderPath -ChildPath "Notes-Backup"), # Change this to your desired backup folder path
    [string]$OutputFilePath = $("$ExportFolderPath\Obsidian2024.pdf"), # Change this to your desired output file path
    [string]$NotesFilePath = $("$ExportFolderPath\Catalog.txt"), # Change this to your notes catalog path.
    [string]$VMName = "PowersidianVM", # Change this to your VM name
    [string]$VMSnapshotName = "Ready", # Change this to your VM snapshot name
    [string]$VMCLIXML = [string]::Empty, # Change this to your VM credential XML file path.
    [string]$PowersidianPath = $(Join-Path -Path (Get-Location) -ChildPath "Powersidian.ps1"), # Change this to your Powersidian script path.
    [string]$TempZipName = "workspace.zip") # Change this to your temporary zip file name.

$debug = $false

Write-Output "PowersidianVM  Copyright 2024 Nuaptan."
Write-Output "This program comes with ABSOLUTELY NO WARRANTY."
Write-Output "This is free software, and you are welcome to redistribute it"
Write-Output "under certain conditions. For details view the source code."
Write-Output ""

Start-Sleep -Milliseconds 1

if ($PSVersionTable.PSVersion -lt [Version]7.4.6) {
    Write-Error "PowersidianVM requires PowerShell 7.4.6 or later. Please update your PowerShell."
    Exit
}

function Test-Administrator {
    $currentUser = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    return $currentUser.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}
if (-not (Test-Administrator)) {
    Write-Output "Trying to invoke UAC."
    try {
        Start-Process pwsh.exe -Verb RunAs -ArgumentList ("-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" " + ($args -join ' '))
    }
    catch {
        Write-Error "Failed to invoke UAC. Please run this script as an administrator."
        Exit
    }
    Exit
}

if (-not (Test-Path -Path $Source -PathType Container)) {
    Write-Error "Source folder $Source not found."
    Exit
}
if (-not (Test-Path -Path $NotesFilePath -PathType Leaf)) {
    Write-Error "Notes catalog file $NotesFilePath not found."
    Exit
}
if (-not (Test-Path -Path $PowersidianPath -PathType Leaf)) {
    Write-Error "Powersidian script $PowersidianPath not found."
    Exit
}

[pscredential]$VMCred = $null
if (($VMCLIXML -eq [string]::Empty) -or (-not (Test-Path -Path $VMCLIXML))) {
    Write-Output "You will need a VM credential for this."
    $VMCred = Get-Credential -Message "Enter your VM credential. Note: for username, use the format 'domain\username' or 'username@domain'."
}
else {
    $VMCred = Import-Clixml -Path $VMCLIXML
}

if (-not (Get-VM $VMName)) {
    Write-Error "VM $VMName not found."
    Exit
}
Write-Output "Detected VM $VMName."

if (-not $debug) {
    if ((Get-VM $VMName).State -ne "Off") {
        Write-Output "Shutting down VM $VMName."
        try {
            Stop-VM -Name $VMName -Force -Confirm:$false
        }
        catch {
            Write-Error "Failed to shut down VM $VMName."
            Exit
        }
    }

    try {
        Write-Output "Trying to restore snapshot $VMSnapshotName."
        Restore-VMSnapshot -Name $VMSnapshotName -VMName $VMName -Confirm:$false
    }
    catch {
        Write-Error "Failed to restore snapshot $VMSnapshotName."
        Exit
    }

    Write-Output "Starting VM $VMName."
    try {
        Start-VM -Name $VMName
    }
    catch {
        Write-Error "Failed to start VM $VMName."
        Exit
    }
    do {
        Start-Sleep -Seconds 1
        Write-Output "Waiting for VM $VMName to start."
    } while ((Get-VM $VMName).State -ne "Running")
}

if (-not (Test-Path -Path $ExportFolderPath)) {
    New-Item -Path $ExportFolderPath -ItemType Directory | Out-Null
    Write-Output "Created export folder at $ExportFolderPath."
}

if (Test-Path -Path $TempZipName) {
    Remove-Item -Path $TempZipName -Force
    Write-Output "Removed temporary zip file $TempZipName."
}
if (-not (Test-Path -Path (Join-Path -Path (Get-Location) -ChildPath "Packages") -PathType Container)) {
    Write-Error "Packages folder not found."
    Exit
}
Write-Output "Preparing relevant files."
Compress-Archive -Path $Source, $NotesFilePath, $PowersidianPath, (Join-Path -Path (Get-Location) -ChildPath "Packages"), (Join-Path -Path (Get-Location) -ChildPath "obsidian.json")  -DestinationPath (Join-Path -Path $ExportFolderPath -ChildPath $TempZipName) -Force
Write-Output "Compressed relevant files to $TempZipName."
try {
    Copy-VMFile -Name $VMName -SourcePath (Join-Path -Path $ExportFolderPath -ChildPath $TempZipName) -DestinationPath "C:\Users\Public\Documents" -FileSource Host -Force
    Write-Output "Copied $TempZipName to VM $VMName."
}
catch {
    Write-Error "Failed to copy $TempZipName to VM $VMName."
    Exit
}
try {
    Invoke-Command -VMName $VMName -ScriptBlock { Expand-Archive -Path "C:\Users\Public\Documents\*.zip" -Destination "C:\Users\Public\Documents" -Force } -Credential $VMCred
    Write-Output "Extracted $TempZipName on VM $VMName."
}
catch {
    Write-Error "Failed to extract $TempZipName on VM $VMName."
    Exit
}

try {
    Write-Output "Overwriting obsidian.json on VM $VMName."
    Invoke-Command -VMName $VMName -ScriptBlock { Copy-Item -Path "C:\Users\Public\Documents\obsidian.json" -Destination "C:\Users\felix\AppData\Roaming\obsidian\obsidian.json" -Force } -Credential $VMCred
}
catch {
    Write-Error "Failed to overwrite obsidian.json on VM $VMName."
    Exit
}
Write-Output "Creating temporary folder on VM $VMName."
Invoke-Command -VMName $VMName -ScriptBlock {
    mkdir "C:\Users\Public\Documents\Backup" | Out-Null
    mkdir "C:\Users\Public\Documents\Temp" | Out-Null
    } -Credential $VMCred
Write-Output "Launching Powersidian on VM $VMName."
Invoke-Command -VMName $VMName -ScriptBlock {
    Set-Location "C:\Users\Public\Documents"
    & "C:\Users\Public\Documents\Powersidian.ps1" -VM -Source "C:\Users\Public\Documents\Notes" -NotesFilePath "C:\Users\Public\Documents\Catalog.txt" -Destination "C:\Users\Public\Documents\Backup" -ExportFolderPath "C:\Users\Public\Documents\Temp"
    } -Credential $VMCred
Pause

# .\Powersidian.ps1 -VM -Source .\Notes\ -NotesFilePath .\Catalog.txt -Destination .\Backup\ -ExportFolderPath "C:\Users\Public\Documents\Temp\"