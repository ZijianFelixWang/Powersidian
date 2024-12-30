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

$debug = $false
$pdfDebug = $false

$exportFolderPath = "E:\Powersidian-Temp"   # Change this to your desired export folder path
$source = "E:\Notes"    # Change this to your notes folder path
$destination = Join-Path -Path $exportFolderPath -ChildPath "Notes-Backup" # Change this to your desired backup folder path
$outputFilePath = "$exportFolderPath\CombinedOutput.pdf" # Change this to your desired output file path
$notesFilePath = "$exportFolderPath\Catalog.txt" # Change this to your notes catalog path.

$obsidianPath = Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) -ChildPath "Programs\obsidian\Obsidian.exe"

Write-Host "Powersidian  Copyright 2024 Nuaptan."
Write-Host "This program comes with ABSOLUTELY NO WARRANTY."
Write-Host "This is free software, and you are welcome to redistribute it"
Write-Host "under certain conditions. For details view the source code."
Write-Host ""

Start-Sleep -Milliseconds 1
Write-Host "Checking system requirements."

if (-not (Test-Path -Path $obsidianPath)) {
    Write-Error "Obsidian is not installed. Please install Obsidian."
    Exit-PSHostProcess
}

$pdftk = Get-Command "pdftk" -ErrorAction SilentlyContinue
if ($null -ne $pdftk) {
    Write-Host "pdftk is installed."
}
else {
    Write-Error "pdftk is not installed. Please install pdftk."
    Exit-PSHostProcess
}

if ($PSVersionTable.PSVersion -lt [Version]7.4.6) {
    Write-Error "Powersidian requires PowerShell 7.4.6 or later. Please update your PowerShell."
    Exit-PSHostProcess
}

Add-Type -AssemblyName PresentationFramework
$response = [System.Windows.MessageBox]::Show(
    "Close irrelevant programs and do NOT use your computer while Powersidian is running. Click Yes to proceed and click No to exit.",
    "Powersidian",
    [System.Windows.MessageBoxButton]::YesNo,
    [System.Windows.MessageBoxImage]::Warning
)
if ($response -eq [System.Windows.MessageBoxResult]::No) {
    Write-Host "Program exits as requested."
    Exit-PSHostProcess
}

if (Test-Path $destination) {
    Write-Host "Removing existing backup folder."
    Remove-Item -Path $destination -Recurse -Force
}
$items = Get-ChildItem -Path $source -Recurse
$totalItems = $items.Count
$currentItem = 0
Write-Host "Backing up $source to $destination."

foreach ($item in $items) {
    $currentItem++
    $percentComplete = [math]::Round(($currentItem / $totalItems) * 100, 2)
    Write-Progress -Activity "Backing Up Files" `
        -Status "Processing: $($item.FullName)" `
        -PercentComplete $percentComplete
    $target = $item.FullName -replace [regex]::Escape($source), $destination
    if ($item.PSIsContainer) {
        if (-not (Test-Path -Path $target)) {
            New-Item -ItemType Directory -Path $target | Out-Null
        }
    }
    else {
        Copy-Item -Path $item.FullName -Destination $target -Force
    }
}
Write-Host "Backup completed."

if (-not $debug) {
    $obsidianProcess = Get-Process -Name "Obsidian" -ErrorAction SilentlyContinue
    if ($obsidianProcess) {
        Write-Host "Obsidian is running. Terminating Obsidian."
        Stop-Process -Id $obsidianProcess.Id -Force
    }
    while ($true) {
        $obStill = Get-Process -Name "Obsidian" -ErrorAction SilentlyContinue
        if (-not ($obStill)) {
            Write-Host "Obsidian terminated."
            break
        }
        else {
            Write-Host "Obsidian is still running."
            Start-Sleep -Seconds 1
        }
    }

    Start-Process -FilePath $obsidianPath -ArgumentList "--remote-debugging-port=9222" | Out-Null
    Write-Host "Started Obsidian with developer access on port 9222."

    for ($cd = 6; $cd -gt 0; $cd--) {
        Write-Progress -Activity "Launching Obsidian" -Status "Count down: $cd" -PercentComplete (((6 - $cd) / 6) * 100)
        Start-Sleep -Seconds 1
    }

    while ($true) {
        $tcp = (Test-NetConnection 127.0.0.1 -Port 9222 -ErrorAction SilentlyContinue).TcpTestSucceeded
        if ($tcp) {
            Write-Host "Obsidian initialized successfully."
            break
        }
        else {
            Write-Host "Waiting for Obsidian to initialize."
            Start-Sleep -Seconds 1
        }
    }
}

$notes = Get-Content -Path $notesFilePath

if (-not $pdfDebug) {
    if (-not (Test-Path -Path $exportFolderPath)) {
        New-Item -Path $exportFolderPath -ItemType Directory | Out-Null
        Write-Host "Created export folder at $exportFolderPath."
    }
    else {
        $pdfFiles = Get-ChildItem -Path $exportFolderPath -Filter "*.pdf" -File
        if ($pdfFiles) {
            Write-Host "Temp folder is not empty. Deleting unnecessary files in temp folder."
            foreach ($file in $pdfFiles) {
                Remove-Item -Path $file.FullName -Force
            }
        }
    }
}

[Reflection.Assembly]::LoadFrom((Join-Path (Get-Location) -ChildPath "Packages\PuppeteerSharp.20.0.5\lib\net8.0\PuppeteerSharp.dll")) | Out-Null
[Reflection.Assembly]::LoadFrom((Join-Path (Get-Location) -ChildPath "Packages\Microsoft.Extensions.DependencyInjection.8.0.0\lib\net8.0\Microsoft.Extensions.DependencyInjection.dll")) | Out-Null
[Reflection.Assembly]::LoadFrom((Join-Path (Get-Location) -ChildPath "Packages\Microsoft.Extensions.DependencyInjection.Abstractions.8.0.0\lib\net8.0\Microsoft.Extensions.DependencyInjection.Abstractions.dll")) | Out-Null
[Reflection.Assembly]::LoadFrom((Join-Path (Get-Location) -ChildPath "Packages\Microsoft.Extensions.Logging.8.0.0\lib\net8.0\Microsoft.Extensions.Logging.dll")) | Out-Null
[Reflection.Assembly]::LoadFrom((Join-Path (Get-Location) -ChildPath "Packages\Microsoft.Extensions.Logging.Abstractions.8.0.0\lib\net8.0\Microsoft.Extensions.Logging.Abstractions.dll")) | Out-Null
[Reflection.Assembly]::LoadFrom((Join-Path (Get-Location) -ChildPath "Packages\Microsoft.Extensions.Options.8.0.0\lib\net8.0\Microsoft.Extensions.Options.dll")) | Out-Null
[Reflection.Assembly]::LoadFrom((Join-Path (Get-Location) -ChildPath "Packages\Microsoft.Extensions.Primitives.8.0.0\lib\net8.0\Microsoft.Extensions.Primitives.dll")) | Out-Null
Write-Host "PuppeteerSharp loaded."

Add-Type -AssemblyName Microsoft.VisualBasic
Add-Type -AssemblyName System.Windows.Forms

$totalSteps = $notes.Count
$currentStep = 0

if (-not $pdfDebug) {
    $browser = [PuppeteerSharp.Puppeteer]::ConnectAsync([PuppeteerSharp.ConnectOptions]@{ BrowserURL = "http://localhost:9222" }).GetAwaiter().GetResult()
    if ($browser.IsConnected) {
        Write-Host "Connection established successfully."
    }
    else {
        Write-Error "Connection establishment failure."
        Exit-PSHostProcess
    }
    $page = ($browser.PagesAsync().GetAwaiter().GetResult())[0]
    foreach ($note in $notes) {
        $currentStep++

        $sProcess = Get-Process -Name "SumatraPDF" -ErrorAction SilentlyContinue
        if ($sProcess) {
            Stop-Process -Id $sProcess.Id -Force
        }

        Write-Progress -Activity "Exporting notes" -Status "Processing note: $note" -PercentComplete (($currentStep / $totalSteps) * 100) -CurrentOperation "Opening note $note"
        $page.BringToFrontAsync().GetAwaiter().GetResult() | Out-Null
        $page.Keyboard.DownAsync("Control").GetAwaiter().GetResult() | Out-Null
        $page.Keyboard.PressAsync("O").GetAwaiter().GetResult() | Out-Null
        $page.Keyboard.UpAsync("Control").GetAwaiter().GetResult() | Out-Null
        $page.Keyboard.TypeAsync($note).GetAwaiter().GetResult() | Out-Null
        $page.Keyboard.PressAsync("Enter").GetAwaiter().GetResult() | Out-Null

        Write-Progress -Activity "Exporting notes" -Status "Processing note: $note" -PercentComplete (($currentStep / $totalSteps) * 100) -CurrentOperation "Invoking PDF exporting dialog"
        $page.Keyboard.DownAsync("Control").GetAwaiter().GetResult() | Out-Null
        $page.Keyboard.PressAsync("E").GetAwaiter().GetResult() | Out-Null
        $page.Keyboard.UpAsync("Control").GetAwaiter().GetResult() | Out-Null

        for ($i = 0; $i -lt 5; $i++) {
            $page.Keyboard.PressAsync("Tab").GetAwaiter().GetResult() | Out-Null
        }
        $page.Keyboard.PressAsync("Enter").GetAwaiter().GetResult() | Out-Null

        Write-Progress -Activity "Exporting notes" -Status "Processing note: $note" -PercentComplete (($currentStep / $totalSteps) * 100) -CurrentOperation "Setting export location"
        Start-Sleep -Milliseconds 500
        while ($true) {
            try {
                [Microsoft.VisualBasic.Interaction]::AppActivate("Save As")
                break
            }
            catch {
                Write-Host "Save As dialog not found. Retrying."
                Start-Sleep -Milliseconds 500
                $page.Keyboard.PressAsync("Escape").GetAwaiter().GetResult() | Out-Null
                Start-Sleep -Milliseconds 100
                $page.Keyboard.PressAsync("Escape").GetAwaiter().GetResult() | Out-Null
                Start-Sleep -Milliseconds 100
                $page.Keyboard.PressAsync("Escape").GetAwaiter().GetResult() | Out-Null
                Start-Sleep -Milliseconds 100
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
        [System.Windows.Forms.SendKeys]::SendWait($exportFolderPath)
        Start-Sleep -Milliseconds 100
        [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
        Start-Sleep -Milliseconds 100
        [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
        Start-Sleep -Milliseconds 100
        [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
        Start-Sleep -Milliseconds 100
        [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")

        Write-Progress -Activity "Exporting notes" -Status "Processing note: $note" -PercentComplete (($currentStep / $totalSteps) * 100) -CurrentOperation "Exporting note $note"
        while (-not (Test-Path -Path "$exportFolderPath\$fn")) {
            Start-Sleep -Milliseconds 100
        }

        Write-Host "Exported note $note successfully."
        Start-Sleep -Seconds 1
    }
}

$sProcess = Get-Process -Name "SumatraPDF" -ErrorAction SilentlyContinue
if ($sProcess) {
    Stop-Process -Id $sProcess.Id -Force
}

$exportedFiles = Get-ChildItem -Path $exportFolderPath -Filter "*.pdf" -File | Sort-Object Name
$eLength = $exportedFiles.Length
Write-Host "$eLength files exported."
$inputFiles = $exportedFiles -join " "

Write-Host "Combining PDF files into $outputFilePath"
Invoke-Expression "pdftk $inputFiles cat output $outputFilePath"
Write-Host "PDF files combined successfully."

Write-Host "Closing Obsidian."
if (-not $debug) {
    $browser.CloseAsync() | Out-Null
}
