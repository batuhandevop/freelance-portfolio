# Batch File Processor - VBA Macro

## What It Does

This VBA macro scans a folder for `.xlsx` and `.csv` files, opens each one, and automatically builds a **master summary workbook** containing:

- **Table of Contents** -- a clickable index of every processed file with row counts, column counts, file sizes, and status indicators.
- **Per-file detail sheets** -- each file gets its own sheet showing metadata, column headers, summary statistics (min / max / average for numeric columns), and a data preview of the first five rows.
- **Hyperlinked navigation** -- click any file name in the TOC to jump to its detail sheet; each detail sheet has a "Back to Table of Contents" link.

## Key Features

| Feature | Details |
|---|---|
| Folder picker | User selects the target folder via a native dialog |
| Format support | `.xlsx` and `.csv` files |
| Error handling | Locked, corrupt, or inaccessible files are logged with an error status instead of crashing the macro |
| Progress indicator | Status bar shows current file number, percentage, and file name |
| Performance | Screen updating, events, and auto-calculation are disabled during processing for speed |
| Professional formatting | Styled headers, alternate row shading, auto-fit columns, freeze panes, borders |

## How to Use

1. Open Excel and press **Alt + F11** to open the VBA editor.
2. Go to **File > Import File** and select `batch_file_processor.bas`.
3. (Optional) In **Tools > References**, enable *Microsoft Scripting Runtime* -- the macro uses late binding so this is not strictly required, but early binding gives better IntelliSense during development.
4. Close the VBA editor and press **Alt + F8**.
5. Select **ProcessFolder** and click **Run**.
6. Choose the folder containing your data files and confirm.
7. The macro creates a new workbook with the consolidated results.

## Requirements

- Microsoft Excel 2016, 2019, or 365 (Windows)
- Macros must be enabled (Trust Center settings)
