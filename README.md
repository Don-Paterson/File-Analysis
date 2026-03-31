# File-Analysis

File and folder analysis tool for Windows. Designed to give a full picture of
how files are organised (or not) across a user's profile — useful for filing
consultations, PC clean-ups, and storage audits.

Generates a self-contained HTML report that opens in any browser.

## Quick Start

Run this one-liner in a PowerShell session (no Admin required):

```powershell
irm https://raw.githubusercontent.com/Don-Paterson/File-Analysis/main/run-file-analysis.ps1 | iex
```

This will:
1. Download `Analyse-Files.ps1` to your Desktop
2. Run it immediately
3. Save the HTML report to the same folder

## What It Analyses

| Section | Details |
|---|---|
| Summary | Total files, total size, Desktop item count (colour coded), duplicate count, reclaimable space |
| Desktop Inventory | Every file and folder on the Desktop listed with type, size, and date. Alert shown if 20+ or 50+ items present |
| Profile Disk Map | Proportional bar chart of every top-level folder in the user profile (WinDirStat-style) |
| File Type Breakdown | Per-location summary, breakdown by category (Documents, Images, Video etc), top 30 extensions |
| Duplicate Files | Files with matching name and size across all scanned locations, with reclaimable space calculation |
| File Timeline | Activity by year, 20 oldest files, 20 most recently modified |
| Large Files | Top 40 largest files across all scanned locations |
| Filing Advice | Suggested folder structure based on what was found, plus plain-English filing principles |

## Locations Scanned

- Desktop
- Documents (including subfolder tree)
- Downloads
- Pictures
- Videos
- Music
- OneDrive (auto-detected if present)
- Full user profile (top-level disk map overview)

## Output

A single self-contained `.html` file saved to the current directory:

```
File-Analysis-<hostname>-<yyyyMMdd-HHmm>.html
```

Open in any browser. Features dark theme, collapsible sections, fixed sidebar
navigation, and colour-coded alerts.

## Requirements

- Windows 10 or Windows 11
- PowerShell 5.1+ or pwsh 7.x
- No Administrator rights required

## Files

| File | Purpose |
|---|---|
| `run-file-analysis.ps1` | Launcher — downloads script to Desktop and runs it |
| `Analyse-Files.ps1` | Main analysis script |
