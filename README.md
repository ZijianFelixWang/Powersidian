# Powersidian & Obindex
## Introduction
**Powersidian** is a PowerShell script that helps you export a whole Obsidian vault of notes easily and combine the exported PDFs into a single book.

**Obindex** is a PowerShell script that does the following.
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
- Reopen Obsidian after processing

These two tools should be useful managing huge Obsidian notes vaults. 

## Requirements
- Obsidian.
- PowerShell. (version >= 7.5.2)
- pdftk Server
- This script only works on Windows, as it uses things like `System.Windows.Forms.SendKeys` for automation about file saving dialogs.

## How it Works & How to Use
The source code should be clear enough, and you can always use `Get-Help` cmdlet to view the help info.

**Also:** I wrote this for myself, so the code only works for my vault structure currently. e.g. `[vault root]\Notes Root\Knowledge`, `[vault root]\Notes Root\Lecture Notes`, etc. You may (very possibly) need to modify the code to make it useful to you.

**Notes on `PowersidianVM.ps1`:** I originally wanted to create a Hyper-V Windows VM so that `Powersidian.ps1` works there. In this way I don't have to wait for over 40 minutes waiting for the script. (Since it uses `sendkeys` I cannot use my computer while it runs.) This feature is only *partially* finished, and is *not* usable now. I don't have time to work on this part now, perhaps I will implement the remaining logic some time later...
