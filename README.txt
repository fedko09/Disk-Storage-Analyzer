Directory Explorer – PowerShell
===============================

Description
-----------
Directory Explorer is a lightweight PowerShell-based GUI tool for Windows 10/11
that lets power users, sysadmins, and regular users quickly inspect the contents
of any directory.

It provides a sortable grid view of files and folders with detailed metadata and
several convenience functions for cleanup and analysis.

Features
--------
- Browse any directory path or use quick-jump locations (Desktop, Documents,
  Downloads, System drive).
- Single-level or recursive scans (Recurse subfolders).
- Hide hidden/system items by default, with an option to show them.
- Filter by:
  - Text (name, extension, or full path)
  - Minimum file size (MB) for “what’s eating this drive” analysis.
- Preset scan profiles:
  - User profile – big files
  - Downloads – big files
  - System drive – very large files
- Detailed columns:
  - Name
  - Extension
  - Type (File / Directory / Other)
  - Size (KB)
  - Created / Modified timestamps
  - Attributes
  - Full path
- Context menu on items:
  - Open (file or folder)
  - Open containing folder in Explorer and select item
  - Copy path(s) to clipboard
  - Delete selected items (with confirmation)
- Toolbar actions:
  - Delete Selected…
  - Export CSV… (exports the currently displayed grid)
- “Busy” overlay with indeterminate progress bar during scans and delete
  operations, so the UI clearly shows that work is in progress.

Requirements
------------
- Windows 10 or Windows 11
- PowerShell 5.x or later
- Ability to run local scripts:
  - You may need to set the execution policy, for example:
    Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned

Usage
-----
1. Save the script as: DirectoryExplorer.ps1
2. Right-click the file and choose:
   "Run with PowerShell"
   or run from a PowerShell console:
   .\DirectoryExplorer.ps1
3. Use the Path box, Browse button, Quick list, or Profile list to choose a
   target directory, then click Load.
4. Use filters, size limits, and context-menu actions as needed.

Notes
-----
- Recursive scans and large folders can take time; the tool will show a
  “Scanning…” overlay while it works.
- Deletions are permanent. Make sure you understand what you’re removing before
  confirming the Delete Selected action.

------------------------------
Created by: Bogdan Fedko
