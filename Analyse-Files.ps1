#Requires -Version 5.1
<#
.SYNOPSIS
    Analyse-Files.ps1 - File & Folder Analysis for Filing Consultation
.DESCRIPTION
    Scans Desktop, Documents, Downloads, Pictures, Videos, Music, OneDrive,
    and the full user profile. Produces a self-contained HTML report covering:
    - Desktop inventory (every file and folder listed)
    - File type breakdown per location
    - Duplicate detection (same name + same size)
    - Timeline of activity (oldest/newest)
    - Large files
    - WinDirStat-style disk usage overview
    - Suggested folder structure based on what was found
.NOTES
    Compatible : Windows 10 / Windows 11
    PowerShell : 5.1+ or pwsh 7.x
    Output     : .\File-Analysis-<hostname>-<date>.html
    Run As     : Standard user is fine; Admin not required
#>

Set-StrictMode -Off
$ErrorActionPreference = 'SilentlyContinue'
$ProgressPreference    = 'SilentlyContinue'

Write-Host ""
Write-Host "  File & Folder Analysis" -ForegroundColor Cyan
Write-Host "  This may take a few minutes on large profiles..." -ForegroundColor DarkCyan
Write-Host ""

$ReportDate = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
$ReportFile = ".\File-Analysis-$($env:COMPUTERNAME)-$(Get-Date -Format 'yyyyMMdd-HHmm').html"

# -------------------------------------------------------------------------------
# LOCATIONS TO SCAN
# -------------------------------------------------------------------------------
$userName  = $env:USERNAME
$userHome  = $env:USERPROFILE

$locations = [ordered]@{
    Desktop   = [Environment]::GetFolderPath('Desktop')
    Documents = [Environment]::GetFolderPath('MyDocuments')
    Downloads = Join-Path $userHome 'Downloads'
    Pictures  = [Environment]::GetFolderPath('MyPictures')
    Videos    = [Environment]::GetFolderPath('MyVideos')
    Music     = [Environment]::GetFolderPath('MyMusic')
}

# OneDrive - check multiple possible paths
$oneDrivePath = $null
foreach ($candidate in @(
    $env:OneDrive,
    $env:OneDriveConsumer,
    $env:OneDriveCommercial,
    (Join-Path $userHome 'OneDrive'),
    (Join-Path $userHome 'OneDrive - Personal'))) {
    if ($candidate -and (Test-Path $candidate)) {
        $oneDrivePath = $candidate
        break
    }
}
if ($oneDrivePath) { $locations['OneDrive'] = $oneDrivePath }

# -------------------------------------------------------------------------------
# HELPER FUNCTIONS
# -------------------------------------------------------------------------------
function HE($s) {
    if ($null -eq $s) { return '' }
    [string]$s -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;' -replace '"','&quot;'
}

function FmtSize($bytes) {
    if ($null -eq $bytes -or $bytes -eq 0) { return '0 B' }
    $b = [long]$bytes
    if ($b -gt 1GB) { return "$([math]::Round($b/1GB,2)) GB" }
    if ($b -gt 1MB) { return "$([math]::Round($b/1MB,1)) MB" }
    if ($b -gt 1KB) { return "$([math]::Round($b/1KB,0)) KB" }
    return "$b B"
}

function FmtDate($d) {
    if ($null -eq $d) { return '' }
    try { return ($d | Get-Date -Format 'yyyy-MM-dd') } catch { return '' }
}

function GetExt($file) {
    $ext = $file.Extension.ToLower()
    if ($ext -eq '') { return '(no extension)' }
    return $ext
}

function CatFromExt($ext) {
    switch -Regex ($ext) {
        '^(\.docx?|\.odt|\.rtf|\.txt|\.pdf|\.md)$'           { return 'Documents' }
        '^(\.xlsx?|\.ods|\.csv)$'                              { return 'Spreadsheets' }
        '^(\.pptx?|\.odp|\.key)$'                             { return 'Presentations' }
        '^(\.jpg|\.jpeg|\.png|\.gif|\.bmp|\.tiff?|\.webp|\.heic|\.raw|\.cr2|\.nef)$' { return 'Images' }
        '^(\.mp4|\.mov|\.avi|\.mkv|\.wmv|\.flv|\.m4v|\.webm)$' { return 'Videos' }
        '^(\.mp3|\.wav|\.flac|\.aac|\.m4a|\.wma|\.ogg)$'     { return 'Audio' }
        '^(\.zip|\.rar|\.7z|\.tar|\.gz|\.bz2|\.iso)$'        { return 'Archives' }
        '^(\.exe|\.msi|\.bat|\.cmd|\.ps1|\.vbs|\.reg|\.dll)$' { return 'Executables/Scripts' }
        '^(\.html?|\.css|\.js|\.php|\.py|\.sql|\.json|\.xml|\.yml|\.yaml)$' { return 'Code/Web' }
        '^(\.lnk|\.url)$'                                     { return 'Shortcuts' }
        '^(\.pst|\.ost|\.eml|\.msg)$'                         { return 'Email' }
        default                                                { return 'Other' }
    }
}

# -------------------------------------------------------------------------------
# SCAN EACH LOCATION
# -------------------------------------------------------------------------------
$allFiles   = [System.Collections.Generic.List[PSCustomObject]]::new()
$scanResults= [ordered]@{}

foreach ($loc in $locations.GetEnumerator()) {
    $name = $loc.Key
    $path = $loc.Value
    Write-Host "  Scanning $name..." -ForegroundColor DarkCyan

    if (-not (Test-Path $path)) {
        $scanResults[$name] = [PSCustomObject]@{
            Path=($path); Exists=$false; Files=@(); TotalBytes=0
        }
        continue
    }

    $files = Get-ChildItem $path -Recurse -File -ErrorAction SilentlyContinue |
        ForEach-Object {
            $rel = $_.FullName.Substring($path.Length).TrimStart('\','/')
            $obj = [PSCustomObject]@{
                Location    = $name
                Name        = $_.Name
                RelPath     = $rel
                FullPath    = $_.FullName
                SizeBytes   = $_.Length
                Extension   = GetExt $_
                Category    = CatFromExt (GetExt $_)
                Created     = $_.CreationTime
                Modified    = $_.LastWriteTime
                Accessed    = $_.LastAccessTime
            }
            $allFiles.Add($obj)
            $obj
        }

    $totalBytes = ($files | Measure-Object SizeBytes -Sum).Sum
    $scanResults[$name] = [PSCustomObject]@{
        Path       = $path
        Exists     = $true
        Files      = $files
        TotalBytes = [long]($totalBytes)
    }
}

# -------------------------------------------------------------------------------
# FULL PROFILE SCAN (top-level folders with sizes - WinDirStat overview)
# -------------------------------------------------------------------------------
Write-Host "  Scanning full profile for disk map..." -ForegroundColor DarkCyan
$profileTopFolders = Get-ChildItem $userHome -Directory -ErrorAction SilentlyContinue |
    ForEach-Object {
        $sz = (Get-ChildItem $_.FullName -Recurse -File -ErrorAction SilentlyContinue |
               Measure-Object Length -Sum).Sum
        [PSCustomObject]@{ Name=$_.Name; Path=$_.FullName; Bytes=[long]($sz) }
    } | Sort-Object Bytes -Descending

$profileFilesAtRoot = Get-ChildItem $userHome -File -ErrorAction SilentlyContinue
$profileTotalBytes  = ($profileTopFolders | Measure-Object Bytes -Sum).Sum

# -------------------------------------------------------------------------------
# DESKTOP DEEP DIVE (special case - list everything)
# -------------------------------------------------------------------------------
Write-Host "  Deep scan of Desktop..." -ForegroundColor DarkCyan
$desktopPath   = $locations['Desktop']
$desktopTop    = Get-ChildItem $desktopPath -ErrorAction SilentlyContinue | Sort-Object PSIsContainer,Name -Descending
$desktopFolders= $desktopTop | Where-Object { $_.PSIsContainer }
$desktopFiles  = $desktopTop | Where-Object { -not $_.PSIsContainer }

$desktopFolderDetails = $desktopFolders | ForEach-Object {
    $items = Get-ChildItem $_.FullName -Recurse -ErrorAction SilentlyContinue
    $sz    = ($items | Where-Object {-not $_.PSIsContainer} | Measure-Object Length -Sum).Sum
    [PSCustomObject]@{
        Name       = $_.Name
        FileCount  = ($items | Where-Object {-not $_.PSIsContainer}).Count
        FolderCount= ($items | Where-Object { $_.PSIsContainer }).Count
        Bytes      = [long]($sz)
        Modified   = $_.LastWriteTime
    }
}

# -------------------------------------------------------------------------------
# DUPLICATE DETECTION (same name + same size across all scanned files)
# -------------------------------------------------------------------------------
Write-Host "  Detecting duplicates..." -ForegroundColor DarkCyan
$dupeGroups = $allFiles |
    Where-Object { $_.SizeBytes -gt 0 } |
    Group-Object { "$($_.Name)|$($_.SizeBytes)" } |
    Where-Object { $_.Count -gt 1 } |
    Sort-Object { ($_.Group | Measure-Object SizeBytes -Sum).Sum } -Descending |
    Select-Object -First 50

# -------------------------------------------------------------------------------
# LARGE FILES (top 40 across all scanned locations)
# -------------------------------------------------------------------------------
$largeFiles = $allFiles | Sort-Object SizeBytes -Descending | Select-Object -First 40

# -------------------------------------------------------------------------------
# TIMELINE
# -------------------------------------------------------------------------------
$timeline = $allFiles | Where-Object { $_.Modified -ne $null } | Sort-Object Modified

$oldest = $timeline | Select-Object -First 20
$newest = $timeline | Sort-Object Modified -Descending | Select-Object -First 20

# Activity by year
$byYear = $allFiles | Where-Object { $_.Modified } |
    Group-Object { $_.Modified.Year } |
    Sort-Object Name |
    ForEach-Object {
        $bytes = ($_.Group | Measure-Object SizeBytes -Sum).Sum
        [PSCustomObject]@{ Year=$_.Name; Count=$_.Count; Bytes=[long]$bytes }
    }

# -------------------------------------------------------------------------------
# FILE TYPE BREAKDOWN (across all locations combined)
# -------------------------------------------------------------------------------
$byCategory = $allFiles |
    Group-Object Category |
    Sort-Object Count -Descending |
    ForEach-Object {
        $bytes = ($_.Group | Measure-Object SizeBytes -Sum).Sum
        [PSCustomObject]@{ Category=$_.Name; Count=$_.Count; Bytes=[long]$bytes }
    }

$byExtension = $allFiles |
    Group-Object Extension |
    Sort-Object Count -Descending |
    Select-Object -First 30 |
    ForEach-Object {
        $bytes = ($_.Group | Measure-Object SizeBytes -Sum).Sum
        [PSCustomObject]@{ Extension=$_.Name; Count=$_.Count; Bytes=[long]$bytes }
    }

# -------------------------------------------------------------------------------
# SUGGESTED FOLDER STRUCTURE
# -------------------------------------------------------------------------------
# Build suggestions based on what categories were actually found
$foundCats = $byCategory | ForEach-Object { $_.Category }
$suggestions = [System.Collections.Generic.List[PSCustomObject]]::new()

$suggestions.Add([PSCustomObject]@{
    Folder = 'Work'
    SubFolders = @('Clients', 'Projects', 'Admin', 'Finance', 'Contracts', 'HR')
    Reason = 'Central home for all work-related documents'
})
$suggestions.Add([PSCustomObject]@{
    Folder = 'Personal'
    SubFolders = @('Finance', 'Health', 'Legal', 'Home', 'Travel', 'Family')
    Reason = 'Personal paperwork and records'
})
if ('Images' -in $foundCats -or 'Videos' -in $foundCats) {
    $suggestions.Add([PSCustomObject]@{
        Folder = 'Photos & Videos'
        SubFolders = @('YYYY-MM Events', 'Family', 'Holidays', 'Scanned Documents')
        Reason = "$(($allFiles | Where-Object {$_.Category -in 'Images','Videos'}).Count) images/videos found across scanned folders"
    })
}
if ('Audio' -in $foundCats) {
    $suggestions.Add([PSCustomObject]@{
        Folder = 'Music'
        SubFolders = @('Albums', 'Playlists', 'Podcasts')
        Reason = "$(($allFiles | Where-Object {$_.Category -eq 'Audio'}).Count) audio files found"
    })
}
if ('Archives' -in $foundCats) {
    $suggestions.Add([PSCustomObject]@{
        Folder = 'Archives & Installers'
        SubFolders = @('Software', 'Backups', 'Zipped Projects')
        Reason = "$(($allFiles | Where-Object {$_.Category -eq 'Archives'}).Count) archive files found - review before deleting"
    })
}
if ('Email' -in $foundCats) {
    $suggestions.Add([PSCustomObject]@{
        Folder = 'Email Archives'
        SubFolders = @('Inbox Archives', 'Sent Archives')
        Reason = "$(($allFiles | Where-Object {$_.Category -eq 'Email'}).Count) .pst/.eml files found - these can be large"
    })
}
$suggestions.Add([PSCustomObject]@{
    Folder = '_Inbox (Staging Area)'
    SubFolders = @()
    Reason = 'Temporary drop zone - anything unprocessed goes here, Desktop should stay empty'
})
$suggestions.Add([PSCustomObject]@{
    Folder = '_Archive'
    SubFolders = @('2020', '2021', '2022', '2023', '2024', '2025')
    Reason = 'Completed or old work filed by year - out of the way but findable'
})

# -------------------------------------------------------------------------------
# BUILD HTML
# -------------------------------------------------------------------------------
function SectionHtml($id, $title, $icon, $content) {
    @"
<section id="$id">
  <div class="sec-head" onclick="toggle('$id')">
    <span class="sec-icon">$icon</span><h2>$title</h2>
    <span class="arrow" id="arr-$id">&#9660;</span>
  </div>
  <div class="sec-body" id="bdy-$id">$content</div>
</section>
"@
}

function SubHtml($title, $content) {
    "<div class='sub'><h3>$title</h3>$content</div>"
}

function TRow([string[]]$cells, [string]$cls='') {
    $tds = ($cells | ForEach-Object { "<td>$(HE $_)</td>" }) -join ''
    "<tr$(if($cls){" class='$cls'"})>$tds</tr>"
}

# -- SUMMARY CARDS -----------------------------------------------------------
$totalFileCount = $allFiles.Count
$totalBytes2    = ($allFiles | Measure-Object SizeBytes -Sum).Sum
$dupeCount      = ($dupeGroups | ForEach-Object { $_.Count - 1 } | Measure-Object -Sum).Sum
$dupeBytes      = ($dupeGroups | ForEach-Object {
    $sz = $_.Group[0].SizeBytes
    $sz * ($_.Count - 1)
} | Measure-Object -Sum).Sum
$desktopCount   = $desktopFiles.Count + $desktopFolders.Count

$cards = @"
<div class="cards">
  <div class="card"><div class="card-val">$totalFileCount</div><div class="card-lbl">Total Files Scanned</div></div>
  <div class="card"><div class="card-val">$(FmtSize $totalBytes2)</div><div class="card-lbl">Total Size</div></div>
  <div class="card $(if($desktopCount -gt 50){'card-warn'}elseif($desktopCount -gt 20){'card-notice'}else{'card-ok'})">
    <div class="card-val">$desktopCount</div><div class="card-lbl">Items on Desktop</div></div>
  <div class="card $(if($dupeBytes -gt 500MB){'card-warn'}else{'card-notice'})">
    <div class="card-val">$dupeCount</div><div class="card-lbl">Likely Duplicates</div></div>
  <div class="card card-notice"><div class="card-val">$(FmtSize $dupeBytes)</div><div class="card-lbl">Reclaimable (Dupes)</div></div>
  <div class="card"><div class="card-val">$($locations.Count)</div><div class="card-lbl">Locations Scanned</div></div>
</div>
<p class='meta'>User: <strong>$userName</strong> &nbsp;|&nbsp; Profile: $userHome &nbsp;|&nbsp; Report: $ReportDate</p>
"@

# -- DESKTOP -----------------------------------------------------------------
$deskFilesHtml = ''
if ($desktopFiles) {
    $rows = ($desktopFiles | Sort-Object Name | ForEach-Object {
        $cat = CatFromExt (GetExt $_)
        $cls = if ($_.Extension -eq '.lnk') { 'muted-row' } else { '' }
        TRow @($_.Name, $_.Extension.ToLower(), $cat, (FmtSize $_.Length), (FmtDate $_.LastWriteTime)) $cls
    }) -join ''
    $deskFilesHtml = "<table><thead><tr><th>Filename</th><th>Type</th><th>Category</th><th>Size</th><th>Modified</th></tr></thead><tbody>$rows</tbody></table>"
} else {
    $deskFilesHtml = "<p class='ok-text'>No loose files on Desktop.</p>"
}

$deskFolderHtml = ''
if ($desktopFolderDetails) {
    $rows = ($desktopFolderDetails | Sort-Object Bytes -Descending | ForEach-Object {
        TRow @($_.Name, $_.FileCount, $_.FolderCount, (FmtSize $_.Bytes), (FmtDate $_.Modified))
    }) -join ''
    $deskFolderHtml = "<table><thead><tr><th>Folder Name</th><th>Files</th><th>Subfolders</th><th>Size</th><th>Modified</th></tr></thead><tbody>$rows</tbody></table>"
} else {
    $deskFolderHtml = "<p class='ok-text'>No folders on Desktop.</p>"
}

$deskNote = if ($desktopCount -gt 50) {
    "<div class='alert critical'><strong>ACTION NEEDED:</strong> $desktopCount items on Desktop. This is significantly impacting PC performance and makes files very hard to find. Everything here should be moved into Documents.</div>"
} elseif ($desktopCount -gt 20) {
    "<div class='alert warning'><strong>TIDY UP:</strong> $desktopCount items on Desktop. Recommend moving files to Documents and keeping Desktop for shortcuts only.</div>"
} elseif ($desktopCount -gt 0) {
    "<div class='alert notice'>$desktopCount items on Desktop. Relatively manageable but worth reviewing what can be moved to Documents.</div>"
} else {
    "<div class='alert ok'>Desktop is clean.</div>"
}

$desktopHtml = $deskNote +
    (SubHtml "Files on Desktop ($($desktopFiles.Count))" $deskFilesHtml) +
    (SubHtml "Folders on Desktop ($($desktopFolderDetails.Count))" $deskFolderHtml)

# -- DISK MAP ----------------------------------------------------------------
$diskMapHtml = "<p class='muted'>Size of each top-level folder in your profile ($userHome)</p>"

# Build proportional bars
$maxBytes = ($profileTopFolders | Select-Object -First 1).Bytes
if ($maxBytes -gt 0) {
    $diskMapHtml += "<div class='diskmap'>"
    foreach ($f in $profileTopFolders) {
        $pct    = [math]::Max(1, [math]::Round($f.Bytes / $maxBytes * 100, 0))
        $gb     = FmtSize $f.Bytes
        $isKnown= $f.Name -in $locations.Keys
        $colour = if ($isKnown) { '#4f8ef7' } else { '#6c757d' }
        $diskMapHtml += @"
<div class='dm-row'>
  <div class='dm-label'>$(HE $f.Name)</div>
  <div class='dm-bar-wrap'>
    <div class='dm-bar' style='width:${pct}%;background:$colour'></div>
  </div>
  <div class='dm-size'>$gb</div>
</div>
"@
    }
    if ($profileFilesAtRoot) {
        $rootSz = ($profileFilesAtRoot | Measure-Object Length -Sum).Sum
        $diskMapHtml += "<div class='dm-row'><div class='dm-label muted'>(root files)</div><div class='dm-bar-wrap'><div class='dm-bar' style='width:1%;background:#555'></div></div><div class='dm-size muted'>$(FmtSize $rootSz)</div></div>"
    }
    $diskMapHtml += "</div>"
    $diskMapHtml += "<p class='muted' style='margin-top:.6em'>Blue = scanned in detail &nbsp;|&nbsp; Grey = present but not deep-scanned</p>"
} else {
    $diskMapHtml += "<p class='muted'>Could not enumerate profile folders.</p>"
}

# -- FILE TYPE BREAKDOWN -----------------------------------------------------
$catColours = @{
    'Documents'='#4f8ef7'; 'Spreadsheets'='#27ae60'; 'Presentations'='#f39c12';
    'Images'='#e74c3c'; 'Videos'='#9b59b6'; 'Audio'='#1abc9c';
    'Archives'='#e67e22'; 'Executables/Scripts'='#c0392b'; 'Code/Web'='#2ecc71';
    'Shortcuts'='#7f8c8d'; 'Email'='#3498db'; 'Other'='#95a5a6'
}

$catRows = ($byCategory | ForEach-Object {
    $col = if ($catColours[$_.Category]) { $catColours[$_.Category] } else { '#95a5a6' }
    $pct = if ($totalFileCount -gt 0) { [math]::Round($_.Count/$totalFileCount*100,1) } else { 0 }
    "<tr><td><span class='dot' style='background:$col'></span>$(HE $_.Category)</td>" +
    "<td>$($_.Count)</td><td>$pct%</td><td>$(FmtSize $_.Bytes)</td></tr>"
}) -join ''
$catTable = "<table><thead><tr><th>Category</th><th>Files</th><th>% of total</th><th>Size</th></tr></thead><tbody>$catRows</tbody></table>"

$extRows = ($byExtension | ForEach-Object {
    TRow @($_.Extension, $_.Count, (FmtSize $_.Bytes))
}) -join ''
$extTable = "<table><thead><tr><th>Extension</th><th>Count</th><th>Total Size</th></tr></thead><tbody>$extRows</tbody></table>"

# Per-location summary
$locRows = ($scanResults.GetEnumerator() | ForEach-Object {
    $r = $_.Value
    if (-not $r.Exists) {
        "<tr class='muted-row'><td>$(HE $_.Key)</td><td class='muted' colspan='4'>Not found / not applicable</td></tr>"
    } else {
        $fc = $r.Files.Count
        $cats = ($r.Files | Group-Object Category | Sort-Object Count -Descending |
                 Select-Object -First 3 | ForEach-Object { $_.Name }) -join ', '
        "<tr><td>$(HE $_.Key)</td><td>$(HE $r.Path)</td><td>$fc</td><td>$(FmtSize $r.TotalBytes)</td><td>$cats</td></tr>"
    }
}) -join ''
$locTable = "<table><thead><tr><th>Location</th><th>Path</th><th>Files</th><th>Size</th><th>Top Categories</th></tr></thead><tbody>$locRows</tbody></table>"

$typeHtml = (SubHtml 'Summary by Location'   $locTable) +
            (SubHtml 'Summary by Category'   $catTable) +
            (SubHtml 'Top 30 File Extensions' $extTable)

# -- DUPLICATES --------------------------------------------------------------
$dupeHtml = ''
if ($dupeGroups) {
    $dupeHtml += "<div class='alert warning'>Found <strong>$($dupeGroups.Count) duplicate groups</strong> " +
                 "($dupeCount extra copies, $(FmtSize $dupeBytes) reclaimable).</div>"
    $rows = ($dupeGroups | ForEach-Object {
        $g    = $_.Group
        $sz   = FmtSize $g[0].SizeBytes
        $name = $g[0].Name
        $locs = ($g | ForEach-Object { "$($_.Location): $($_.RelPath)" }) -join '<br>'
        "<tr><td>$(HE $name)</td><td>$($g.Count)</td><td>$sz</td><td>$locs</td></tr>"
    }) -join ''
    $dupeHtml += "<table><thead><tr><th>Filename</th><th>Copies</th><th>Size Each</th><th>Locations</th></tr></thead><tbody>$rows</tbody></table>"
} else {
    $dupeHtml = "<p class='ok-text'>No duplicates detected across scanned locations.</p>"
}

# -- TIMELINE ----------------------------------------------------------------
$yearRows = ($byYear | ForEach-Object {
    $maxC = ($byYear | Measure-Object Count -Maximum).Maximum
    $pct  = if ($maxC -gt 0) { [math]::Round($_.Count/$maxC*100,0) } else { 0 }
    $age  = (Get-Date).Year - [int]$_.Year
    $cls  = if ($age -gt 5) { 'muted-row' } else { '' }
    "<tr class='$cls'><td>$($_.Year)</td><td>$($_.Count)</td><td>$(FmtSize $_.Bytes)</td>" +
    "<td><div class='bar-wrap'><div class='bar-fill' style='width:${pct}%;background:#4f8ef7'></div></div></td></tr>"
}) -join ''
$yearTable = "<table><thead><tr><th>Year</th><th>Files Modified</th><th>Total Size</th><th>Volume</th></tr></thead><tbody>$yearRows</tbody></table>"

$oldRows = ($oldest | ForEach-Object {
    TRow @($_.Name, $_.Location, (FmtDate $_.Created), (FmtDate $_.Modified), $_.Category)
}) -join ''
$oldTable = "<table><thead><tr><th>Filename</th><th>Location</th><th>Created</th><th>Modified</th><th>Category</th></tr></thead><tbody>$oldRows</tbody></table>"

$newRows = ($newest | ForEach-Object {
    TRow @($_.Name, $_.Location, (FmtDate $_.Modified), $_.Category, (FmtSize $_.SizeBytes))
}) -join ''
$newTable = "<table><thead><tr><th>Filename</th><th>Location</th><th>Modified</th><th>Category</th><th>Size</th></tr></thead><tbody>$newRows</tbody></table>"

$timeHtml = (SubHtml 'Activity by Year (files last modified)' $yearTable) +
            (SubHtml '20 Oldest Files'                        $oldTable) +
            (SubHtml '20 Most Recently Modified Files'        $newTable)

# -- LARGE FILES -------------------------------------------------------------
$lgRows = ($largeFiles | ForEach-Object {
    TRow @($_.Name, $_.Location, $_.RelPath, $_.Category, (FmtSize $_.SizeBytes), (FmtDate $_.Modified))
}) -join ''
$lgHtml = "<table><thead><tr><th>Filename</th><th>Location</th><th>Path</th><th>Category</th><th>Size</th><th>Modified</th></tr></thead><tbody>$lgRows</tbody></table>"

# -- SUGGESTIONS -------------------------------------------------------------
$sugHtml = @"
<div class='alert notice'>
  <strong>Suggested approach:</strong> Start with the Desktop. Move everything into a single
  staging folder called <strong>_Inbox</strong> in Documents, then sort from there.
  The Desktop should only ever hold shortcuts, not files.
</div>
<div class='suggest-grid'>
"@

foreach ($sug in $suggestions) {
    $subList = if ($sug.SubFolders.Count -gt 0) {
        "<ul>$(($sug.SubFolders | ForEach-Object { "<li>$_</li>" }) -join '')</ul>"
    } else { '' }
    $sugHtml += @"
<div class='sug-card'>
  <div class='sug-title'>&#128193; $(HE $sug.Folder)</div>
  $subList
  <div class='sug-reason'>$(HE $sug.Reason)</div>
</div>
"@
}
$sugHtml += "</div>"

$sugHtml += @"
<div class='sub' style='margin-top:1.5em'>
  <h3>Key Principles to Share</h3>
  <ul class='principles'>
    <li><strong>Desktop = shortcuts only.</strong> No files, no folders. Treat it like a physical desk - clear it daily.</li>
    <li><strong>One place for everything.</strong> Documents is the home base. Everything branches from there.</li>
    <li><strong>Name files to be findable.</strong> Use dates: <code>2025-03-Invoice-BT.pdf</code> not <code>Invoice(2)FINAL.pdf</code></li>
    <li><strong>Year-based archiving.</strong> At the end of each year, move completed work into <code>_Archive\YYYY</code></li>
    <li><strong>Downloads is a staging area.</strong> Nothing lives there permanently - it gets sorted weekly.</li>
    <li><strong>Duplicates waste space and cause confusion.</strong> Keep one copy, in one place, with a clear name.</li>
  </ul>
</div>
"@

# -- ASSEMBLE FULL HTML ------------------------------------------------------
$navItems = @(
    @('summary',   'Summary'),
    @('desktop',   'Desktop'),
    @('diskmap',   'Disk Map'),
    @('filetypes', 'File Types'),
    @('dupes',     'Duplicates'),
    @('timeline',  'Timeline'),
    @('large',     'Large Files'),
    @('suggest',   'Suggestions')
)
$navHtml = ($navItems | ForEach-Object {
    "<a href='#$($_[0])' onclick='scrollTo(""$($_[0])"");return false;'>$($_[1])</a>"
}) -join ''

$allSections =
    (SectionHtml 'summary'   'Summary'                          '#'  $cards) +
    (SectionHtml 'desktop'   'Desktop Inventory'                'D'  $desktopHtml) +
    (SectionHtml 'diskmap'   'Profile Disk Map'                 'M'  $diskMapHtml) +
    (SectionHtml 'filetypes' 'File Type Breakdown'              'T'  $typeHtml) +
    (SectionHtml 'dupes'     'Duplicate Files'                  'X'  $dupeHtml) +
    (SectionHtml 'timeline'  'File Timeline'                    'Y'  $timeHtml) +
    (SectionHtml 'large'     'Large Files (top 40)'             'L'  $lgHtml) +
    (SectionHtml 'suggest'   'Filing Advice & Suggested Layout' 'S'  $sugHtml)

$html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>File Analysis - $userName - $ReportDate</title>
<style>
:root{--bg:#0f1117;--surface:#1a1d27;--surface2:#22263a;--border:#2e3350;
  --text:#e2e8f0;--muted:#8892a4;--accent:#4f8ef7;--ok:#27ae60;--warn:#f39c12;--crit:#e74c3c;--hover:#2a3050;}
*{box-sizing:border-box;margin:0;padding:0;}
body{font-family:'Segoe UI',system-ui,sans-serif;background:var(--bg);color:var(--text);font-size:14px;line-height:1.5;display:flex;min-height:100vh;}
nav{width:180px;min-width:180px;background:var(--surface);border-right:1px solid var(--border);position:fixed;top:0;left:0;height:100vh;overflow-y:auto;z-index:100;}
.nav-head{padding:1em;border-bottom:1px solid var(--border);}
.nav-head h1{font-size:13px;color:var(--accent);font-weight:700;text-transform:uppercase;letter-spacing:.04em;}
.nav-head p{font-size:11px;color:var(--muted);margin-top:.2em;}
nav a{display:block;padding:.4em 1.1em;color:var(--muted);text-decoration:none;font-size:12.5px;border-left:3px solid transparent;transition:all .15s;}
nav a:hover,nav a.active{background:var(--hover);color:var(--text);border-left-color:var(--accent);}
main{margin-left:180px;padding:1.8em 2.2em;width:100%;max-width:1350px;}
.report-header{background:var(--surface);border:1px solid var(--border);border-radius:10px;padding:1.2em 1.8em;margin-bottom:1.2em;}
.report-header h1{font-size:19px;font-weight:700;}
.report-header p{color:var(--muted);font-size:12px;margin-top:.3em;}
section{background:var(--surface);border:1px solid var(--border);border-radius:9px;margin-bottom:1.1em;overflow:hidden;}
.sec-head{display:flex;align-items:center;gap:.65em;padding:.85em 1.2em;cursor:pointer;user-select:none;background:var(--surface2);border-bottom:1px solid var(--border);transition:background .15s;}
.sec-head:hover{background:var(--hover);}
.sec-head h2{font-size:13.5px;font-weight:600;flex:1;}
.sec-icon{font-size:11px;background:var(--border);color:var(--muted);padding:.15em .45em;border-radius:3px;font-weight:700;}
.arrow{color:var(--muted);font-size:11px;transition:transform .2s;}
.arrow.closed{transform:rotate(-90deg);}
.sec-body{padding:1.1em 1.3em;}
.sec-body.hidden{display:none;}
.sub{margin-bottom:1.4em;}
.sub h3{font-size:11px;text-transform:uppercase;letter-spacing:.07em;color:var(--accent);margin-bottom:.55em;padding-bottom:.3em;border-bottom:1px solid var(--border);}
table{width:100%;border-collapse:collapse;font-size:13px;margin-top:.3em;}
th{background:var(--surface2);color:var(--muted);font-weight:600;font-size:11px;text-transform:uppercase;letter-spacing:.05em;padding:.45em .75em;text-align:left;border-bottom:1px solid var(--border);}
td{padding:.4em .75em;border-bottom:1px solid var(--border);vertical-align:top;word-break:break-word;}
tr:last-child td{border-bottom:none;}
tr:hover td{background:var(--hover);}
.muted-row td{opacity:.45;}
.alert{padding:.65em .95em;border-radius:6px;margin-bottom:.7em;font-size:13px;border-left:4px solid;}
.alert.critical{background:rgba(231,76,60,.12);border-color:var(--crit);}
.alert.warning{background:rgba(243,156,18,.12);border-color:var(--warn);}
.alert.notice{background:rgba(79,142,247,.12);border-color:var(--accent);}
.alert.ok{background:rgba(39,174,96,.12);border-color:var(--ok);}
.ok-text{color:var(--ok);font-weight:600;}
.warn-text{color:var(--warn);font-weight:600;}
.meta{color:var(--muted);font-size:11px;margin-top:.8em;}
.muted{color:var(--muted);}
ul{padding-left:1.3em;}li{padding:.12em 0;font-size:13px;}
p{font-size:13px;margin:.25em 0;}
code{background:var(--surface2);padding:.1em .35em;border-radius:3px;font-size:12px;font-family:'Consolas','Courier New',monospace;}
.cards{display:flex;gap:.8em;flex-wrap:wrap;margin-bottom:.8em;}
.card{background:var(--surface2);border:1px solid var(--border);border-radius:8px;padding:.9em 1.2em;min-width:130px;flex:1;}
.card-val{font-size:22px;font-weight:700;color:var(--text);}
.card-lbl{font-size:11px;color:var(--muted);margin-top:.2em;}
.card-warn{border-color:var(--crit)!important;}.card-warn .card-val{color:var(--crit);}
.card-notice{border-color:var(--warn)!important;}.card-notice .card-val{color:var(--warn);}
.card-ok{border-color:var(--ok)!important;}.card-ok .card-val{color:var(--ok);}
.dot{display:inline-block;width:10px;height:10px;border-radius:50%;margin-right:6px;vertical-align:middle;}
.bar-wrap{background:var(--surface2);border-radius:3px;height:12px;position:relative;min-width:80px;overflow:hidden;}
.bar-fill{height:100%;border-radius:3px;background:var(--accent);}
.diskmap{margin-top:.5em;}
.dm-row{display:flex;align-items:center;gap:.7em;margin-bottom:.35em;}
.dm-label{width:160px;min-width:160px;font-size:12.5px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
.dm-bar-wrap{flex:1;background:var(--surface2);border-radius:3px;height:18px;overflow:hidden;}
.dm-bar{height:100%;border-radius:3px;min-width:2px;transition:width .3s;}
.dm-size{width:80px;text-align:right;font-size:12px;color:var(--muted);white-space:nowrap;}
.suggest-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(220px,1fr));gap:.8em;margin:.8em 0;}
.sug-card{background:var(--surface2);border:1px solid var(--border);border-radius:8px;padding:1em;}
.sug-title{font-weight:700;font-size:13.5px;margin-bottom:.5em;color:var(--text);}
.sug-reason{font-size:11px;color:var(--muted);margin-top:.6em;font-style:italic;}
.sug-card ul{padding-left:1.1em;}
.sug-card li{font-size:12px;padding:.1em 0;}
.principles li{margin-bottom:.4em;padding:.2em 0;}
::-webkit-scrollbar{width:5px;height:5px;}
::-webkit-scrollbar-thumb{background:var(--border);border-radius:3px;}
@media(max-width:800px){nav{display:none;}main{margin-left:0;padding:1em;}}
</style>
</head>
<body>
<nav>
  <div class="nav-head">
    <h1>File Analysis</h1>
    <p>$userName</p>
  </div>
  $navHtml
  <div style="padding:1em;font-size:10.5px;color:var(--muted);border-top:1px solid var(--border);margin-top:.4em">$ReportDate</div>
</nav>
<main>
  <div class="report-header">
    <h1>&#128193; File &amp; Folder Analysis</h1>
    <p>$userName &nbsp;&bull;&nbsp; $userHome &nbsp;&bull;&nbsp; $ReportDate</p>
    <p style="margin-top:.4em">Scanned: $(($locations.GetEnumerator() | ForEach-Object { $_.Key }) -join ', ')</p>
  </div>
  $allSections
</main>
<script>
function toggle(id){
  var b=document.getElementById('bdy-'+id);
  var a=document.getElementById('arr-'+id);
  a.classList.toggle('closed',b.classList.toggle('hidden'));
}
function scrollTo(id){
  var el=document.getElementById(id);
  if(el)el.scrollIntoView({behavior:'smooth',block:'start'});
  document.querySelectorAll('nav a').forEach(function(a){a.classList.remove('active');});
  var lnk=document.querySelector('nav a[href="#'+id+'"]');
  if(lnk)lnk.classList.add('active');
}
window.addEventListener('scroll',function(){
  var cur='';
  document.querySelectorAll('section[id]').forEach(function(s){
    if(window.scrollY>=s.offsetTop-120)cur=s.id;
  });
  document.querySelectorAll('nav a').forEach(function(a){
    a.classList.toggle('active',a.getAttribute('href')==='#'+cur);
  });
});
</script>
</body>
</html>
"@

$html | Out-File -FilePath $ReportFile -Encoding UTF8 -NoNewline

Write-Host ""
Write-Host "  [OK] Analysis complete." -ForegroundColor Green
Write-Host "  HTML Report : $ReportFile" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Open with   : Invoke-Item '$ReportFile'" -ForegroundColor DarkCyan
Write-Host ""
