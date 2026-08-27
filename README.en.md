# Windows Scripts

[Deutsch](README.md) | **English**

This repository contains small PowerShell tools for storage analysis and literal text searches in files. Every tool is available in German and English. The graphical tools are designed for Windows and do not require additional PowerShell modules or a separate WinDirStat installation.

## Quick start

The easiest entry point is the bilingual launcher:

1. Download or clone the complete repository.
2. Double-click [`Start-WindowsScripts.cmd`](Start-WindowsScripts.cmd).
3. Select **English** or **Deutsch**, choose a tool, and click **Launch**.

The batch file removes possible Internet zone markers from itself and from [`Start-WindowsScripts.ps1`](Start-WindowsScripts.ps1), then starts the graphical launcher. Before launching a tool, the launcher also calls `Unblock-File` for the selected script. This permanently removes that file's Mark of the Web or `Zone.Identifier`. The launcher and selected tool use `ExecutionPolicy Bypass` only in their respective new processes; no machine-wide execution policy is changed and no administrator rights are granted. Only run downloaded scripts after reviewing or otherwise trusting their source.

## Files and languages

| Tool | German | English | Purpose |
| --- | --- | --- | --- |
| StorageTree | [`StorageTree.ps1`](StorageTree.ps1) | [`StorageTree.en.ps1`](StorageTree.en.ps1) | Compact graphical storage analyzer |
| WinDirStat | [`WinDirStat.ps1`](WinDirStat.ps1) | [`WinDirStat.en.ps1`](WinDirStat.en.ps1) | Folder tree, file view, and top 200 files |
| WinDirStat 2 | [`WinDirStat2.ps1`](WinDirStat2.ps1) | [`WinDirStat2.en.ps1`](WinDirStat2.en.ps1) | Adds context menus and file actions |
| WinDirStat 3 | [`WinDirStat3.ps1`](WinDirStat3.ps1) | [`WinDirStat3.en.ps1`](WinDirStat3.en.ps1) | Adds an interactive SequoiaView-style treemap |
| File content search | [`suche.ps1`](suche.ps1) | [`suche.en.ps1`](suche.en.ps1) | Recursive literal text search with a GUI |
| Archive.org CDX search | [`CdxSearchGui.ps1`](CdxSearchGui.ps1) | [`CdxSearchGui.en.ps1`](CdxSearchGui.en.ps1) | Finds archived files for a domain |
| Encoding Doctor | [`Encoding-Doctor.de.ps1`](Encoding-Doctor.de.ps1) | [`Encoding-Doctor.ps1`](Encoding-Doctor.ps1) | Checks and repairs text encodings with backups |
| Profile folder sizes | [`ProfileSize.ps1`](ProfileSize.ps1) | [`ProfileSize.en.ps1`](ProfileSize.en.ps1) | Console summary of profile subfolder sizes |

Historic filenames without a language suffix are normally the German versions, while English variants use `.en.ps1`. The later addition `Encoding-Doctor.ps1` was already written in English, so its German variant exceptionally uses `.de.ps1`. The launcher itself is bilingual.

## Requirements

The graphical scripts require:

- Windows with an interactive desktop session
- Windows PowerShell 5.1 or a newer PowerShell version on Windows
- the WPF, Windows Forms, and .NET desktop components provided with Windows

`ProfileSize.ps1` and `ProfileSize.en.ps1` are console-only and do not require desktop components. All analyses require read access to the selected folders and files. Administrator rights are not mandatory, but protected, locked, or unavailable entries may be skipped without them.

The Archive.org CDX search additionally needs Internet access to `web.archive.org`. Encoding Doctor operates only on local files.

Despite their names, the WinDirStat scripts are self-contained PowerShell applications. They do not launch or depend on the separate WinDirStat program.

## Manual launch

Change to the repository folder first:

```powershell
Set-Location "C:\path\to\Windows-Scripts"
```

For example, launch the English treemap version with:

```powershell
powershell.exe -NoProfile -STA -File .\WinDirStat3.en.ps1
```

If the file is trusted and only its Internet zone marker prevents execution, remove that marker once:

```powershell
Unblock-File .\WinDirStat3.en.ps1
```

Alternatively, `-ExecutionPolicy Bypass` can be used for one new PowerShell process only:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File .\WinDirStat3.en.ps1
```

This does not change a machine-wide policy. Group Policy, AppLocker, application control, access permissions, and elevation requirements still apply.

None of the analysis scripts has public command-line parameters. Scan paths are selected in the graphical storage analyzers. The search and profile-size scripts use fixed starting locations described below.

## Storage analyzers

### StorageTree

The compact base version displays the size of a selected folder and its subfolders as a tree.

- The current user profile is preselected.
- Enter a folder or choose one with **Browse**.
- **Scan** or Enter starts the analysis.
- A background runspace keeps the interface responsive.
- Subfolders are sorted by size and show their percentage of the parent folder.
- Double-clicking a folder opens it in File Explorer.
- **Cancel** requests a cooperative stop and displays a partial result when possible.

### WinDirStat

This version adds a file pane to StorageTree:

- It displays up to the 200 largest files found in the complete scan.
- Selecting a folder shows the files located directly in that folder, sorted by size.
- **Top 200 in scan** returns to the global list.
- Double-clicking a file opens it with its registered Windows application.

### WinDirStat 2

Version 2 adds Explorer-like interactions. A file context menu provides:

- open the file
- reveal the file in File Explorer
- copy its complete path
- move it to the Recycle Bin
- invoke additional shell actions registered on the local Windows system

Folders provide **Open in Explorer** and **Copy path**. The Delete key also sends the selected file to the Recycle Bin. Whether Windows asks for confirmation and whether a Recycle Bin is available depends on system and drive settings.

### WinDirStat 3: treemap

Version 3 keeps the tree, file list, and file actions from version 2 and adds an interactive treemap inspired by SequoiaView:

- rectangle area represents logical file size
- a squarified layout favors rectangles that are easier to compare and select
- colors group common file extensions and file types
- tooltips show the complete path and formatted size
- clicking a rectangle synchronizes the selected file where possible
- double-clicking a rectangle opens the file
- in global mode, the view visualizes up to the 200 largest files in the scan; for a selected folder, it shows files located directly in that folder
- selectable display limits are 100, 250, 500, 750, or 1,000 areas; the default is 500
- above the limit, the aggregate area consumes one slot itself: at the default limit, the 499 largest files are shown individually and all remaining files share one proportional, non-clickable area
- zero-byte files may appear in the list but have no visible treemap area

The implementation is a lightweight WPF interpretation rather than a pixel-identical reimplementation of SequoiaView. It uses gradients and borders for a cushion-like appearance without external rendering libraries.

The visualization is based on the principles described in the Eindhoven University of Technology publications [Cushion treemaps](https://research.tue.nl/en/publications/cushion-treemaps-visualization-of-hierarchical-information/) and [Squarified treemaps](https://research.tue.nl/en/publications/squarified-treemaps/).

### Shared behavior and limitations

All storage analyzers sum logical file lengths. During a scan they report scanned files and folders, bytes read, caught access/read errors, and the current path. The progress indicator is intentionally indeterminate because the tools do not perform a complete counting pass first. The global top-200 list is part of that scan snapshot; the direct file list for a selected folder is read again when the folder is selected.

Keep these limitations in mind:

- Subdirectories with the Reparse Point attribute, such as junctions and directory symbolic links, are omitted to avoid loops and duplicate counting. An explicitly selected root directory is exempt from that check.
- Locked, unreadable, offline, or concurrently removed entries can be skipped.
- The UI value labeled **skipped** counts caught errors; it is not an exact count of every omitted file and folder.
- Logical file lengths are not necessarily equal to allocated disk space. Compression, sparse or cloud files, hard links, metadata, and allocation-unit rounding can produce different physical usage.
- Concurrent file-system changes can make a scan an inconsistent snapshot.
- Very large or slow directory trees require time and memory. Cancellation is cooperative and can be delayed by a blocking file-system operation.
- There is no CSV or report export.

## File content search

`suche.ps1` and `suche.en.ps1` open a Windows Forms interface and recursively search matching files in a background runspace.

The interface accepts:

- **File filter:** a case-insensitive wildcard such as `*.txt`, `*.ps1`, or `*config*`; an empty filter becomes `*`.
- **Search term:** literal, case-insensitive text that is checked line by line; it is not a regular expression.

Results contain the relative file path, line number, and trimmed line text:

```text
File: Subfolder\Example.ps1 | Line 42: Get-ChildItem -Recurse
```

When run normally as a script file, the search root is the folder containing the script, including all subfolders. If PowerShell cannot determine the script folder, the current working directory is used. The GUI cannot select another root folder.

Unreadable directories and files are mostly skipped silently. A broad filter such as `*` may also attempt to read binary files as text. Files without a byte order mark are interpreted as UTF-8; legacy encodings such as Windows-1252 may be displayed incorrectly. The file filter and search term are trimmed, so leading or trailing whitespace cannot be part of the query.

## Profile folder sizes

`ProfileSize.ps1` and `ProfileSize.en.ps1` sum the files below every direct subfolder of `%USERPROFILE%` and print formatted gigabyte values to the console:

```text
    1.25 GB  C:\Users\Name\Documents
```

This is intended as a quick overview. Without `-Force`, hidden folders and files are normally omitted, which can exclude `AppData`. Files directly in the user-profile root are not counted. Errors during recursive file enumeration within a profile subfolder are suppressed. The output has no total or sorting and consists of formatted strings rather than reusable PowerShell objects. Decimal and thousands separators follow the system's regional settings. It also reports logical file lengths rather than physical disk usage.

## Archive.org CDX search

`CdxSearchGui.ps1` and `CdxSearchGui.en.ps1` query the Internet Archive CDX service and produce clickable Wayback download links for archived files from a domain.

The interface provides:

- a domain or URL input
- the CDX match types `domain`, `prefix`, `host`, and `exact`
- file extensions such as `exe`, `*.exe,*.zip`, or a detected regular expression
- optional from/to dates
- results with archive timestamp, reported size, HTTP status, and a direct `id_` link

The query restricts results to HTTP status 200 and collapses identical URL keys, so it displays at most one result per URL rather than a complete list of every archived capture. It runs synchronously on the UI thread, so the window can become unresponsive during a slow request until the 120-second timeout. Search values are transmitted to `web.archive.org`. Generated `id_` links target archived raw files, and clicking one opens it through the Windows default application. In particular, do not open or execute archived programs or compressed files without checking them. API availability, rate limits, and very large result sets are outside the script's control.

## Encoding Doctor

`Encoding-Doctor.ps1` is the English variant and `Encoding-Doctor.de.ps1` is the German variant. Both files intentionally use ASCII-only source text so that they remain readable even under incorrect ANSI/UTF-8 detection.

Recommended first invocation:

```powershell
.\Encoding-Doctor.ps1 -DryRun
```

Without `-DryRun`, the tool works recursively from its own script folder and may modify files:

- valid UTF-8 PowerShell files without a BOM receive a UTF-8 BOM for Windows PowerShell 5.1; existing bytes remain unchanged when no text damage is found
- common UTF-8/Windows-1252 mojibake patterns are detected heuristically and repaired when possible
- PowerShell files detected as legacy encoding are converted to UTF-8 with BOM under the assumption that they use Windows-1252
- every changed file is backed up under `_Encoding_Backup_<timestamp>` before modification
- binary files, unsupported extensions, files larger than 25 MB, the running Doctor script, and previous backup folders are skipped

In the launcher, **Yes** selects repair mode and **No** performs a dry run; **No** is the safe default, and **Cancel** does not launch the tool. Mojibake repair is heuristic and can misinterpret legitimate special characters. Before repair mode, also create a Git commit or external backup. Review the resulting diff and retain the generated backup folder until verification is complete.

The current version also adds a BOM to ASCII-only PowerShell source even though ASCII is already unambiguous without one. This is harmless but can create additional file changes. Local backup folders are excluded from Git uploads through `.gitignore`.

## Data safety

- StorageTree, WinDirStat, the search tools, and the profile-size tools do not modify files during normal use.
- WinDirStat 2 and 3 can move the selected file to the Recycle Bin through their context menu or the Delete key.
- Additional Windows shell actions are loaded from the local shell without a security allowlist. Depending on the registered action, they may modify files or external state.
- After a file is removed, the file list and treemap are refreshed. Folder sizes still represent the previous scan until a new scan is performed, and the global top-200 list may contain fewer than 200 entries in the meantime.
- The Archive.org CDX search transmits the entered domain, filters, and date boundaries to `web.archive.org` but does not modify local files itself.
- Encoding Doctor can recode or repair many local text files; run it with `-DryRun` first and rely on its backups when reviewing changes.
