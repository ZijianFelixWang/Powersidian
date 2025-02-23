# Powersidian = Power Obsidian
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
    [string]$OutputFilePath = $("$ExportFolderPath\TmpExport2025.pdf"), # Change this to your desired output file path
    [string]$NotesFilePath = $("$ExportFolderPath\Catalog.txt"), # Change this to your notes catalog path.
    [string]$ObsidianPath = $(Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) -ChildPath "Programs\obsidian\Obsidian.exe"),
    [switch]$ObindexMode = $false,
    [switch]$VM) # This is a virtual machine switch. If you are running this script on a virtual machine, add -VM to the command line.

$debug = $false
$pdfDebug = $false

Write-Output "Powersidian  Copyright 2024 Nuaptan."
Write-Output "This program comes with ABSOLUTELY NO WARRANTY."
Write-Output "This is free software, and you are welcome to redistribute it"
Write-Output "under certain conditions. For details view the source code."
Write-Output ""

Start-Sleep -Milliseconds 1
Write-Progress -Completed

if ($VM) {
    Write-Output "Running on a virtual machine. Write-Progress and Write-Error are hooked."
    function Write-Progress {
        param (
            [string]$Activity,
            [string]$Status,
            [int]$PercentComplete,
            [string]$CurrentOperation = [string]::Empty
        )
        if ($CurrentOperation -eq [string]::Empty) {
            Write-Output "$($Activity): $Status ($PercentComplete%)"
        }
        else {
            Write-Output "$($Activity): $Status ($PercentComplete%) - $CurrentOperation"
        }
    }

    function Write-Error {
        param (
            [string]$Message
        )
        Write-Output "Error: $Message"
    }
}

Write-Output "Checking system requirements."

if (-not (Test-Path -Path $ObsidianPath)) {
    Write-Error "Obsidian is not installed. Please install Obsidian."
    Exit
}

$pdftk = Get-Command "pdftk" -ErrorAction SilentlyContinue
if ($null -ne $pdftk) {
    Write-Output "pdftk is installed."
}
else {
    Write-Error "pdftk is not installed. Please install pdftk."
    Exit
}

if ($PSVersionTable.PSVersion -lt [Version]7.4.6) {
    Write-Error "Powersidian requires PowerShell 7.4.6 or later. Please update your PowerShell."
    Exit
}

if ((-not $VM) -and (-not $ObindexMode)) {
    Add-Type -AssemblyName PresentationFramework
    $response = [System.Windows.MessageBox]::Show(
        "Close irrelevant programs and do NOT use your computer while Powersidian is running. Click Yes to proceed and click No to exit.",
        "Powersidian",
        [System.Windows.MessageBoxButton]::YesNo,
        [System.Windows.MessageBoxImage]::Warning
    )
    if ($response -eq [System.Windows.MessageBoxResult]::No) {
        Write-Output "Program exits as requested."
        Exit
    }
}
else {
    Write-Output "Skipping user interaction."
}

if ((-not $VM) -and (-not $ObindexMode)) {
    if (Test-Path $Destination) {
        Write-Output "Removing existing backup folder."
        Remove-Item -Path $Destination -Recurse -Force
        Write-Progress -Completed
    }
    $items = Get-ChildItem -Path $Source -Recurse
    $totalItems = $items.Count
    $currentItem = 0
    Write-Output "Backing up $Source to $Destination."

    foreach ($item in $items) {
        $currentItem++
        $percentComplete = [math]::Round(($currentItem / $totalItems) * 100, 2)
        Write-Progress -Activity "Backing Up Files" `
            -Status "Processing: $($item.FullName)" `
            -PercentComplete $percentComplete
        $target = $item.FullName -replace [regex]::Escape($Source), $Destination
        if ($item.PSIsContainer) {
            if (-not (Test-Path -Path $target)) {
                New-Item -ItemType Directory -Path $target | Out-Null
            }
        }
        else {
            Copy-Item -Path $item.FullName -Destination $target -Force
        }
    }
    Write-Output "Backup completed."
    Write-Progress -Completed
}
else {
    Write-Output "Skipping backup."
}

if (-not $debug) {
    $obsidianProcess = Get-Process -Name "Obsidian" -ErrorAction SilentlyContinue
    if ($obsidianProcess) {
        Write-Output "Obsidian is running. Terminating Obsidian."
        Stop-Process -Id $obsidianProcess.Id -Force
    }
    while ($true) {
        $obStill = Get-Process -Name "Obsidian" -ErrorAction SilentlyContinue
        if (-not ($obStill)) {
            Write-Output "Obsidian terminated."
            break
        }
        else {
            Write-Output "Obsidian is still running."
            Start-Sleep -Seconds 1
        }
    }

    Start-Process -FilePath $ObsidianPath -ArgumentList "--remote-debugging-port=9222" -RedirectStandardOutput "Temp:\$([Guid]::NewGuid().ToString())"
    Write-Output "Started Obsidian with developer access on port 9222."

    for ($cd = 6; $cd -gt 0; $cd--) {
        Write-Progress -Activity "Launching Obsidian" -Status "Count down: $cd" -PercentComplete (((6 - $cd) / 6) * 100)
        Start-Sleep -Seconds 1
    }
    Write-Progress -Completed

    while ($true) {
        $tcp = (Test-NetConnection 127.0.0.1 -Port 9222 -ErrorAction SilentlyContinue).TcpTestSucceeded
        if ($tcp) {
            Write-Output "Obsidian initialized successfully."
            break
        }
        else {
            Write-Output "Waiting for Obsidian to initialize."
            Start-Sleep -Seconds 1
        }
    }
}

$notes = Get-Content -Path $NotesFilePath

if (-not $pdfDebug) {
    if (-not (Test-Path -Path $ExportFolderPath)) {
        New-Item -Path $ExportFolderPath -ItemType Directory | Out-Null
        Write-Output "Created export folder at $ExportFolderPath."
    }
    else {
        $pdfFiles = Get-ChildItem -Path $ExportFolderPath -Filter "*.pdf" -File
        if ($pdfFiles) {
            Write-Output "Temp folder is not empty. Deleting unnecessary files in temp folder."
            foreach ($file in $pdfFiles) {
                Remove-Item -Path $file.FullName -Force
            }
        }
    }
}

if ($VM) {
    Set-Location "C:\Users\Public\Documents"
}

[Reflection.Assembly]::LoadFrom((Join-Path (Get-Location) -ChildPath "Packages\PuppeteerSharp.20.0.5\lib\net8.0\PuppeteerSharp.dll")) | Out-Null
[Reflection.Assembly]::LoadFrom((Join-Path (Get-Location) -ChildPath "Packages\Microsoft.Extensions.DependencyInjection.8.0.0\lib\net8.0\Microsoft.Extensions.DependencyInjection.dll")) | Out-Null
[Reflection.Assembly]::LoadFrom((Join-Path (Get-Location) -ChildPath "Packages\Microsoft.Extensions.DependencyInjection.Abstractions.8.0.0\lib\net8.0\Microsoft.Extensions.DependencyInjection.Abstractions.dll")) | Out-Null
[Reflection.Assembly]::LoadFrom((Join-Path (Get-Location) -ChildPath "Packages\Microsoft.Extensions.Logging.8.0.0\lib\net8.0\Microsoft.Extensions.Logging.dll")) | Out-Null
[Reflection.Assembly]::LoadFrom((Join-Path (Get-Location) -ChildPath "Packages\Microsoft.Extensions.Logging.Abstractions.8.0.0\lib\net8.0\Microsoft.Extensions.Logging.Abstractions.dll")) | Out-Null
[Reflection.Assembly]::LoadFrom((Join-Path (Get-Location) -ChildPath "Packages\Microsoft.Extensions.Options.8.0.0\lib\net8.0\Microsoft.Extensions.Options.dll")) | Out-Null
[Reflection.Assembly]::LoadFrom((Join-Path (Get-Location) -ChildPath "Packages\Microsoft.Extensions.Primitives.8.0.0\lib\net8.0\Microsoft.Extensions.Primitives.dll")) | Out-Null
Write-Output "PuppeteerSharp loaded."

Add-Type -AssemblyName Microsoft.VisualBasic
Add-Type -AssemblyName System.Windows.Forms

if ($VM) {
    [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")    # to avoid "Trust this vault?" dialog
}

$totalSteps = $notes.Count
$currentStep = 0

if (-not $pdfDebug) {
    $browser = [PuppeteerSharp.Puppeteer]::ConnectAsync([PuppeteerSharp.ConnectOptions]@{ BrowserURL = "http://localhost:9222" }).GetAwaiter().GetResult()
    if ($browser.IsConnected) {
        Write-Output "Connection established successfully."
    }
    else {
        Write-Error "Connection establishment failure."
        Exit
    }

    Write-Output ""

    $page = ($browser.PagesAsync().GetAwaiter().GetResult())[0]
    foreach ($note in $notes) {
        $currentStep++

        $sProcess = Get-Process -Name "SumatraPDF" -ErrorAction SilentlyContinue
        if ($sProcess) {
            Stop-Process -Id $sProcess.Id -Force
        }

        Write-Progress -Activity "Exporting notes" -Status "Opening note: $note ($currentStep/$totalSteps)" -PercentComplete (($currentStep / $totalSteps) * 100) 
        $page.BringToFrontAsync().GetAwaiter().GetResult() | Out-Null
        $page.Keyboard.DownAsync("Control").GetAwaiter().GetResult() | Out-Null
        $page.Keyboard.PressAsync("O").GetAwaiter().GetResult() | Out-Null
        $page.Keyboard.UpAsync("Control").GetAwaiter().GetResult() | Out-Null
        $page.Keyboard.TypeAsync($note).GetAwaiter().GetResult() | Out-Null
        $page.Keyboard.PressAsync("Enter").GetAwaiter().GetResult() | Out-Null

        Write-Progress -Activity "Exporting notes" -Status "Invoking PDF exporting dialog: $note ($currentStep/$totalSteps)" -PercentComplete (($currentStep / $totalSteps) * 100)
        $page.Keyboard.DownAsync("Control").GetAwaiter().GetResult() | Out-Null
        $page.Keyboard.PressAsync("E").GetAwaiter().GetResult() | Out-Null
        $page.Keyboard.UpAsync("Control").GetAwaiter().GetResult() | Out-Null

        for ($i = 0; $i -lt 5; $i++) {
            $page.Keyboard.PressAsync("Tab").GetAwaiter().GetResult() | Out-Null
        }
        $page.Keyboard.PressAsync("Enter").GetAwaiter().GetResult() | Out-Null

        Write-Progress -Activity "Exporting notes" -Status "Setting export location: $note ($currentStep/$totalSteps)" -PercentComplete (($currentStep / $totalSteps) * 100)
        Start-Sleep -Milliseconds 700
        while ($true) {
            try {
                [Microsoft.VisualBasic.Interaction]::AppActivate("Save As")
                break
            }
            catch {
                Write-Output "$(Get-Date -Format "HH:mm:ss") Save As dialog not found. Retrying."

                Start-Sleep -Milliseconds 500
                $page.Keyboard.PressAsync("Escape").GetAwaiter().GetResult() | Out-Null
                Start-Sleep -Milliseconds 100
                $page.Keyboard.PressAsync("Escape").GetAwaiter().GetResult() | Out-Null
                Start-Sleep -Milliseconds 100
                $page.Keyboard.PressAsync("Escape").GetAwaiter().GetResult() | Out-Null
                Start-Sleep -Milliseconds 100

                for ($j = 0; $j -lt 7; $j++) {
                    $page.Keyboard.DownAsync("Control").GetAwaiter().GetResult() | Out-Null
                    $page.Keyboard.PressAsync("Z").GetAwaiter().GetResult() | Out-Null
                    $page.Keyboard.UpAsync("Control").GetAwaiter().GetResult() | Out-Null
                    Start-Sleep -Milliseconds 300
                }

                $page.Keyboard.DownAsync("Control").GetAwaiter().GetResult() | Out-Null
                $page.Keyboard.PressAsync("E").GetAwaiter().GetResult() | Out-Null
                $page.Keyboard.UpAsync("Control").GetAwaiter().GetResult() | Out-Null

                for ($i = 0; $i -lt 5; $i++) {
                    $page.Keyboard.PressAsync("Tab").GetAwaiter().GetResult() | Out-Null
                }
                $page.Keyboard.PressAsync("Enter").GetAwaiter().GetResult() | Out-Null
                Start-Sleep -Milliseconds 500
            }
        }
        [System.Windows.Forms.SendKeys]::SendWait("^a")
        Start-Sleep -Milliseconds 100
        [System.Windows.Forms.SendKeys]::SendWait("{BACKSPACE}")
        $fn = "{0:D4}.pdf" -f $currentStep
        [System.Windows.Forms.SendKeys]::SendWait($fn)
        Start-Sleep -Milliseconds 100
        [System.Windows.Forms.SendKeys]::SendWait("{F4}")
        Start-Sleep -Milliseconds 500
        [System.Windows.Forms.SendKeys]::SendWait("^a")
        Start-Sleep -Milliseconds 100
        [System.Windows.Forms.SendKeys]::SendWait("{BACKSPACE}")
        [System.Windows.Forms.SendKeys]::SendWait($ExportFolderPath)
        Start-Sleep -Milliseconds 100
        [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
        Start-Sleep -Milliseconds 100
        [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
        Start-Sleep -Milliseconds 100
        [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
        Start-Sleep -Milliseconds 100
        [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")

        Write-Progress -Activity "Exporting notes" -Status "Exporting note: $note ($currentStep/$totalSteps)" -PercentComplete (($currentStep / $totalSteps) * 100)
        while (-not (Test-Path -Path "$ExportFolderPath\$fn")) {
            Start-Sleep -Milliseconds 100
        }

        # Write-Output "$currentStep/$totalSteps Exported note $note successfully."
        Start-Sleep -Seconds 1
    }
}

$sProcess = Get-Process -Name "SumatraPDF" -ErrorAction SilentlyContinue
if ($sProcess) {
    Stop-Process -Id $sProcess.Id -Force
}

$exportedFiles = Get-ChildItem -Path $ExportFolderPath -Filter "*.pdf" -File | Sort-Object Name
$eLength = $exportedFiles.Length
Write-Output "$eLength files exported."
$inputFiles = $exportedFiles -join " "

Write-Output "Combining PDF files into $OutputFilePath"
Invoke-Expression "pdftk $inputFiles cat output $OutputFilePath"
Write-Output "PDF files combined successfully."

Write-Output "Closing Obsidian."
if (-not $debug) {
    $browser.CloseAsync() | Out-Null
}
