<#
.SYNOPSIS
    Migrates Windows user profile data, application settings, browser bookmarks, and Outlook PST files.

.DESCRIPTION
    Two-phase migration script:
    1. Backup Mode: Gathers user shell folders (detecting OneDrive redirections), AppData\Roaming,
       Outlook PST archives, and browser bookmarks (Chrome, Edge, Brave, Firefox). Caches and lock
       files are excluded. A portable .7z archive is created with relative folder mapping.
    2. Restore Mode: Merges the archive into the signed-in profile on the new computer.
       Existing destination files are skipped (-aos and Robocopy /XC /XN /XO) so new files are
       never overwritten. OneDrive-redirected folders are automatically detected and populated.
       Chromium bookmarks are merged into a dedicated folder without overwriting existing bookmarks.

.PARAMETER Mode
    Operation mode: 'Backup' to create the archive on the old PC, or 'Restore' to merge on the new laptop.

.PARAMETER ArchivePath
    Path to the .7z archive file.
    In Backup mode: Destination path (e.g., 'D:\Migration\ColleagueProfile.7z'). Defaults to current directory.
    In Restore mode: Source path to the existing archive to extract.

.PARAMETER SevenZipPath
    Optional manual path to 7z.exe if not in standard locations or system PATH.

.PARAMETER ExcludeDownloads
    Switch to skip the Downloads folder if it contains large, unneeded installers.

.EXAMPLE
    # On Old Computer (run while colleague is signed in):
    .\Migrate-UserProfile.ps1 -Mode Backup -ArchivePath "D:\ColleagueBackup.7z"

.EXAMPLE
    # On New Laptop (run while colleague is signed in):
    .\Migrate-UserProfile.ps1 -Mode Restore -ArchivePath "D:\ColleagueBackup.7z"
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [ValidateSet('Backup', 'Restore')]
    [string]$Mode,

    [Parameter(Position = 1)]
    [string]$ArchivePath,

    [string]$SourceProfile = $env:USERPROFILE,

    [string]$TargetProfile = $env:USERPROFILE,

    [string]$SevenZipPath,

    [switch]$ExcludeDownloads
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Write-Log {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        [ValidateSet('INFO', 'WARN', 'ERROR', 'SUCCESS')]
        [string]$Level = 'INFO'
    )
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $color = switch ($Level) {
        'INFO'    { 'Cyan' }
        'WARN'    { 'Yellow' }
        'ERROR'   { 'Red' }
        'SUCCESS' { 'Green' }
    }
    Write-Host "[$timestamp] [$Level] $Message" -ForegroundColor $color
}

function Find-7Zip {
    param([string]$CustomPath)

    if ($CustomPath -and (Test-Path -LiteralPath $CustomPath)) {
        return (Resolve-Path -LiteralPath $CustomPath).Path
    }

    $paths = @(
        (Get-Command 7z.exe, 7za.exe -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Source -First 1),
        "$env:ProgramFiles\7-Zip\7z.exe",
        "${env:ProgramFiles(x86)}\7-Zip\7z.exe",
        "$env:ProgramFiles\NVIDIA Corporation\NVIDIA App\7z.exe",
        "$env:ProgramData\chocolatey\bin\7z.exe",
        "$env:LOCALAPPDATA\Programs\7-Zip\7z.exe",
        "$env:USERPROFILE\scoop\shims\7z.exe"
    ) | Where-Object { $_ -and (Test-Path -LiteralPath $_) }

    $candidates = @($paths)
    if ($candidates.Count -gt 0) {
        return (Resolve-Path -LiteralPath $candidates[0]).Path
    }

    return $null
}

function Get-KnownFolderRegistryMap {
    param([string]$ProfileRoot = $env:USERPROFILE)

    $standardNames = @('Desktop', 'Documents', 'Downloads', 'Pictures', 'Music', 'Videos', 'Favorites')
    $resolved = [ordered]@{}

    $isCurrentUser = $false
    if ($ProfileRoot -and (Test-Path -LiteralPath $ProfileRoot) -and (Test-Path -LiteralPath $env:USERPROFILE)) {
        $isCurrentUser = ((Resolve-Path -LiteralPath $ProfileRoot).Path -ieq (Resolve-Path -LiteralPath $env:USERPROFILE).Path)
    }

    if ($isCurrentUser) {
        $regKey = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders'
        $props = Get-ItemProperty -Path $regKey -ErrorAction SilentlyContinue

        $map = [ordered]@{
            'Desktop'   = $props.Desktop
            'Documents' = $props.Personal
            'Downloads' = $props.'{374DE290-123F-4565-9164-39C4925E467B}'
            'Pictures'  = $props.'My Pictures'
            'Music'     = $props.'My Music'
            'Videos'    = $props.'My Video'
            'Favorites' = $props.Favorites
        }

        foreach ($name in $map.Keys) {
            $raw = $map[$name]
            if ($raw) {
                $expanded = [System.Environment]::ExpandEnvironmentVariables($raw)
                if (Test-Path -LiteralPath $expanded) {
                    $resolved[$name] = (Resolve-Path -LiteralPath $expanded).Path
                    continue
                }
            }
            $defaultFallback = Join-Path $ProfileRoot $name
            if (Test-Path -LiteralPath $defaultFallback) {
                $resolved[$name] = (Resolve-Path -LiteralPath $defaultFallback).Path
            }
        }
    } else {
        foreach ($name in $standardNames) {
            $oneDrivePath = Join-Path $ProfileRoot "OneDrive\$name"
            if (Test-Path -LiteralPath $oneDrivePath) {
                $resolved[$name] = (Resolve-Path -LiteralPath $oneDrivePath).Path
                continue
            }
            $directPath = Join-Path $ProfileRoot $name
            if (Test-Path -LiteralPath $directPath) {
                $resolved[$name] = (Resolve-Path -LiteralPath $directPath).Path
            }
        }
    }
    return $resolved
}

function Test-OpenApplications {
    $apps = @('chrome', 'msedge', 'firefox', 'brave', 'outlook', 'teams')
    $running = Get-Process -Name $apps -ErrorAction SilentlyContinue
    if ($running) {
        $names = ($running | Select-Object -ExpandProperty ProcessName -Unique) -join ', '
        Write-Log "Open applications detected: $names" -Level WARN
        Write-Log "Close these applications to prevent file-lock warnings during migration." -Level WARN
    }
}

function Merge-ChromiumBookmarks {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourceBookmarksFile,
        [Parameter(Mandatory = $true)]
        [string]$TargetBookmarksFile,
        [Parameter(Mandatory = $true)]
        [string]$BrowserName
    )

    if (-not (Test-Path -LiteralPath $SourceBookmarksFile)) {
        return
    }

    $targetDir = Split-Path -Parent $TargetBookmarksFile
    if (-not (Test-Path -LiteralPath $targetDir)) {
        New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
    }

    if (-not (Test-Path -LiteralPath $TargetBookmarksFile)) {
        Copy-Item -LiteralPath $SourceBookmarksFile -Destination $TargetBookmarksFile -Force
        Write-Log "Copied $BrowserName bookmarks to new profile." -Level SUCCESS
        return
    }

    try {
        $srcJson = Get-Content -LiteralPath $SourceBookmarksFile -Raw -Encoding UTF8 | ConvertFrom-Json
        $tgtJson = Get-Content -LiteralPath $TargetBookmarksFile -Raw -Encoding UTF8 | ConvertFrom-Json

        $maxId = [ref]0
        $getMaxId = {
            param($node, [ref]$refMax)
            if ($node.PSObject.Properties['id'] -and $node.id) {
                $parsed = 0
                if ([int]::TryParse($node.id.ToString(), [ref]$parsed) -and $parsed -gt $refMax.Value) {
                    $refMax.Value = $parsed
                }
            }
            if ($node.PSObject.Properties['children'] -and $node.children) {
                foreach ($child in $node.children) {
                    & $getMaxId $child $refMax
                }
            }
        }

        if ($tgtJson.roots.PSObject.Properties['bookmark_bar'] -and $tgtJson.roots.bookmark_bar) { & $getMaxId $tgtJson.roots.bookmark_bar $maxId }
        if ($tgtJson.roots.PSObject.Properties['other'] -and $tgtJson.roots.other) { & $getMaxId $tgtJson.roots.other $maxId }
        if ($tgtJson.roots.PSObject.Properties['synced'] -and $tgtJson.roots.synced) { & $getMaxId $tgtJson.roots.synced $maxId }

        $reassignIds = {
            param($node, [ref]$counter)
            $counter.Value++
            $node.id = $counter.Value.ToString()
            if ($node.PSObject.Properties['children'] -and $node.children) {
                foreach ($child in $node.children) {
                    & $reassignIds $child $counter
                }
            }
        }

        $sourceChildren = [System.Collections.Generic.List[object]]::new()
        if ($srcJson.roots.PSObject.Properties['bookmark_bar'] -and $srcJson.roots.bookmark_bar -and $srcJson.roots.bookmark_bar.PSObject.Properties['children'] -and $srcJson.roots.bookmark_bar.children) {
            foreach ($c in $srcJson.roots.bookmark_bar.children) { $sourceChildren.Add($c) }
        }
        if ($srcJson.roots.PSObject.Properties['other'] -and $srcJson.roots.other -and $srcJson.roots.other.PSObject.Properties['children'] -and $srcJson.roots.other.children) {
            foreach ($c in $srcJson.roots.other.children) { $sourceChildren.Add($c) }
        }

        if ($sourceChildren.Count -eq 0) {
            return
        }

        $mergedFolder = [PSCustomObject]@{
            children      = $sourceChildren.ToArray()
            date_added    = [string][long]((Get-Date).ToFileTimeUtc() * 10)
            date_modified = "0"
            id            = ""
            name          = "Migrated Bookmarks (Old PC)"
            type          = "folder"
        }

        & $reassignIds $mergedFolder $maxId

        if (-not $tgtJson.roots.PSObject.Properties['other']) {
            $tgtJson.roots | Add-Member -MemberType NoteProperty -Name 'other' -Value ([PSCustomObject]@{
                children = @()
                name     = "Other bookmarks"
                type     = "folder"
            })
        }

        $existingOther = [System.Collections.Generic.List[object]]::new()
        if ($tgtJson.roots.other.PSObject.Properties['children'] -and $tgtJson.roots.other.children) {
            foreach ($item in $tgtJson.roots.other.children) {
                $existingOther.Add($item)
            }
        }
        $existingOther.Add($mergedFolder)
        $tgtJson.roots.other.children = $existingOther.ToArray()

        $tempOut = "$TargetBookmarksFile.migrated.tmp"
        $tgtJson | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $tempOut -Encoding UTF8
        Move-Item -LiteralPath $tempOut -Destination $TargetBookmarksFile -Force
        Write-Log "Merged $BrowserName bookmarks into 'Migrated Bookmarks (Old PC)' under Other Bookmarks." -Level SUCCESS
    } catch {
        Write-Log "Automated merge for $BrowserName bookmarks encountered an error: $_" -Level WARN
        Write-Log "The original bookmarks remain safely accessible in Documents\Migrated Browser Bookmarks." -Level INFO
    }
}

# --- Resolve 7-Zip Executable ---
$exe7z = Find-7Zip -CustomPath $SevenZipPath
if (-not $exe7z) {
    Write-Log "7-Zip executable (7z.exe) could not be found." -Level ERROR
    Write-Log "Please install 7-Zip (e.g. run 'winget install 7zip.7zip') or specify -SevenZipPath 'C:\Path\To\7z.exe'" -Level ERROR
    throw "7-Zip not found."
}
Write-Log "Using 7-Zip: $exe7z" -Level INFO

# ==============================================================================
# BACKUP MODE (Old Computer)
# ==============================================================================
if ($Mode -eq 'Backup') {
    $sourceProfile = if ($SourceProfile) { (Resolve-Path -LiteralPath $SourceProfile).Path } else { (Resolve-Path -LiteralPath $env:USERPROFILE).Path }
    $sourceUser = [System.IO.Path]::GetFileName($sourceProfile)
    $sourceComputer = $env:COMPUTERNAME

    Write-Log "Initiating backup for user: $sourceUser on $sourceComputer" -Level INFO
    Test-OpenApplications

    if (-not $ArchivePath) {
        $dateStr = Get-Date -Format 'yyyyMMdd_HHmmss'
        $ArchivePath = Join-Path (Get-Location).Path "UserProfile_${sourceUser}_${dateStr}.7z"
    }

    $archiveDir = Split-Path -Parent $ArchivePath
    if ($archiveDir -and -not (Test-Path -LiteralPath $archiveDir)) {
        New-Item -ItemType Directory -Path $archiveDir -Force | Out-Null
    }

    Write-Log "Target archive: $ArchivePath" -Level INFO

    $stagingGuid = [System.Guid]::NewGuid().ToString('N')
    $stagingDir = Join-Path $env:TEMP "ProfileMigration_Staging_$stagingGuid"
    $excListFile = Join-Path $env:TEMP "7z_exclude_$stagingGuid.txt"
    $createdJunctions = [System.Collections.Generic.List[string]]::new()

    try {
        New-Item -ItemType Directory -Path $stagingDir -Force | Out-Null

        # 1. Map Known Folders (Resolving OneDrive redirection)
        $knownMap = Get-KnownFolderRegistryMap -ProfileRoot $sourceProfile
        Write-Log "Detected Windows known folders (including OneDrive redirection):" -Level INFO
        foreach ($k in $knownMap.Keys) {
            if ($ExcludeDownloads -and $k -eq 'Downloads') {
                Write-Log "  - ${k}: SKIPPED (Excluded via -ExcludeDownloads)" -Level INFO
                continue
            }
            $targetPath = $knownMap[$k]
            Write-Log "  + $k -> $targetPath" -Level INFO
            $linkPath = Join-Path $stagingDir $k
            cmd.exe /c mklink /J "$linkPath" "$targetPath" > $null 2>&1
            if (Test-Path -LiteralPath $linkPath) {
                $createdJunctions.Add($linkPath)
            } else {
                # Fallback to copy if junction is not supported
                Copy-Item -LiteralPath $targetPath -Destination $linkPath -Recurse -Force
            }
        }

        # 1b. Map User-Created Profile Folders (e.g. Projects, Code, Workspace, Scripts, .ssh)
        $knownFolderNames = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($name in $knownMap.Keys) { [void]$knownFolderNames.Add($name) }

        $systemReservedFolders = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        $reservedList = @(
            'AppData', 'Application Data', 'Local Settings', 'NetHood', 'PrintHood',
            'Recent', 'SendTo', 'Start Menu', 'Templates', 'My Documents', 'Cookies'
        )
        foreach ($r in $reservedList) { [void]$systemReservedFolders.Add($r) }

        $customProfileDirs = Get-ChildItem -LiteralPath $sourceProfile -Directory -Force -ErrorAction SilentlyContinue
        if ($customProfileDirs) {
            Write-Log "Scanning for user-created profile folders..." -Level INFO
            foreach ($dir in $customProfileDirs) {
                if ($dir.Attributes -band [System.IO.FileAttributes]::ReparsePoint) { continue }
                if ($systemReservedFolders.Contains($dir.Name) -or $knownFolderNames.Contains($dir.Name)) { continue }

                Write-Log "  + User-created folder: $($dir.Name) -> $($dir.FullName)" -Level INFO
                $linkPath = Join-Path $stagingDir $dir.Name
                cmd.exe /c mklink /J "$linkPath" "$($dir.FullName)" > $null 2>&1
                if (Test-Path -LiteralPath $linkPath) {
                    $createdJunctions.Add($linkPath)
                } else {
                    Copy-Item -LiteralPath $dir.FullName -Destination $linkPath -Recurse -Force
                }
            }
        }

        # 2. AppData\Roaming
        $srcRoaming = Join-Path $sourceProfile 'AppData\Roaming'
        if (Test-Path -LiteralPath $srcRoaming) {
            $appDataStaging = Join-Path $stagingDir 'AppData'
            New-Item -ItemType Directory -Path $appDataStaging -Force | Out-Null
            $roamingLink = Join-Path $appDataStaging 'Roaming'
            cmd.exe /c mklink /J "$roamingLink" "$srcRoaming" > $null 2>&1
            if (Test-Path -LiteralPath $roamingLink) {
                $createdJunctions.Add($roamingLink)
            } else {
                Copy-Item -LiteralPath $srcRoaming -Destination $roamingLink -Recurse -Force
            }
        }

        # 3. Outlook PST Files
        $pstStaging = Join-Path $stagingDir 'OutlookFiles'
        $pstFiles = @()
        $srcLocal = Join-Path $sourceProfile 'AppData\Local'
        $outlookLocal = Join-Path $srcLocal 'Microsoft\Outlook'
        if (Test-Path -LiteralPath $outlookLocal) {
            $pstFiles += Get-ChildItem -LiteralPath $outlookLocal -Filter '*.pst' -File -ErrorAction SilentlyContinue
        }
        if ($knownMap['Documents']) {
            $docsOutlook = Join-Path $knownMap['Documents'] 'Outlook Files'
            if (Test-Path -LiteralPath $docsOutlook) {
                # These are already inside Documents junction, but tracked in manifest
                $pstFiles += Get-ChildItem -LiteralPath $docsOutlook -Filter '*.pst' -File -ErrorAction SilentlyContinue
            }
        }

        if ($pstFiles.Count -gt 0) {
            New-Item -ItemType Directory -Path $pstStaging -Force | Out-Null
            Write-Log "Found Outlook PST files to archive:" -Level INFO
            foreach ($pst in $pstFiles) {
                Write-Log "  + PST: $($pst.FullName) ($([math]::Round($pst.Length / 1MB, 1)) MB)" -Level INFO
                # Only copy if outside Documents to avoid duplication
                if (-not ($pst.FullName.StartsWith($knownMap['Documents'], [System.StringComparison]::OrdinalIgnoreCase))) {
                    Copy-Item -LiteralPath $pst.FullName -Destination $pstStaging -Force
                }
            }
        }

        # 4. Browser Bookmarks
        $bookmarkStaging = Join-Path $stagingDir 'BrowserBookmarks'
        New-Item -ItemType Directory -Path $bookmarkStaging -Force | Out-Null

        $browserSpecs = @(
            @{ Name = 'Chrome'; Path = (Join-Path $srcLocal 'Google\Chrome\User Data\Default\Bookmarks') },
            @{ Name = 'Edge';   Path = (Join-Path $srcLocal 'Microsoft\Edge\User Data\Default\Bookmarks') },
            @{ Name = 'Brave';  Path = (Join-Path $srcLocal 'BraveSoftware\Brave-Browser\User Data\Default\Bookmarks') }
        )

        foreach ($spec in $browserSpecs) {
            if (Test-Path -LiteralPath $spec.Path) {
                $destDir = Join-Path $bookmarkStaging $spec.Name
                New-Item -ItemType Directory -Path $destDir -Force | Out-Null
                Copy-Item -LiteralPath $spec.Path -Destination (Join-Path $destDir 'Bookmarks') -Force
                Write-Log "Captured $($spec.Name) bookmarks." -Level INFO
            }
        }

        # Firefox bookmarks
        $firefoxProfiles = Join-Path $srcRoaming 'Mozilla\Firefox\Profiles'
        if (Test-Path -LiteralPath $firefoxProfiles) {
            $fxProfiles = Get-ChildItem -LiteralPath $firefoxProfiles -Directory -ErrorAction SilentlyContinue
            if ($fxProfiles) {
                $fxDest = Join-Path $bookmarkStaging 'Firefox'
                New-Item -ItemType Directory -Path $fxDest -Force | Out-Null
                foreach ($fxp in $fxProfiles) {
                    $places = Join-Path $fxp.FullName 'places.sqlite'
                    if (Test-Path -LiteralPath $places) {
                        Copy-Item -LiteralPath $places -Destination (Join-Path $fxDest "places_$($fxp.Name).sqlite") -Force
                    }
                    $backups = Join-Path $fxp.FullName 'bookmarkbackups'
                    if (Test-Path -LiteralPath $backups) {
                        Copy-Item -LiteralPath $backups -Destination (Join-Path $fxDest "backups_$($fxp.Name)") -Recurse -Force
                    }
                }
                Write-Log "Captured Firefox bookmark databases and backups." -Level INFO
            }
        }

        # 5. Metadata Manifest
        $manifest = [PSCustomObject]@{
            SourceComputer       = $sourceComputer
            SourceUser           = $sourceUser
            SourceProfile        = $sourceProfile
            BackupDateUtc        = (Get-Date).ToUniversalTime().ToString('u')
            WindowsVersion       = [System.Environment]::OSVersion.VersionString
            SourceKnownFolders   = $knownMap
            HasOutlookFiles      = (Test-Path -LiteralPath $pstStaging)
        }
        $manifestFile = Join-Path $stagingDir 'MigrationManifest.json'
        $manifest | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $manifestFile -Encoding UTF8

        # 6. Exclusions List (Temporary files, caches, locks, registry hives)
        $exclusions = @(
            'NTUSER.DAT*',
            'ntuser.dat*',
            'UsrClass.dat*',
            'usrclass.dat*',
            '*.lock',
            '*.tmp',
            'desktop.ini',
            'thumbs.db',
            'Thumbs.db',
            'IconCache.db',
            'Cache',
            'Cache\*',
            'Code Cache',
            'Code Cache\*',
            'GPUCache',
            'GPUCache\*',
            'ShaderCache',
            'ShaderCache\*',
            'Crashpad',
            'Crashpad\*',
            'CacheStorage',
            'CacheStorage\*',
            'ScriptCache',
            'ScriptCache\*',
            'WebCache*',
            'Temp',
            'Temp\*'
        )
        [System.IO.File]::WriteAllLines($excListFile, $exclusions, [System.Text.UTF8Encoding]::new($false))

        # 7. Compress with 7-Zip
        Write-Log "Compressing profile archive to $ArchivePath..." -Level INFO

        $7zArgs = @(
            'a',
            '-t7z',
            '-mx=5',
            '-ssw',
            "`"$ArchivePath`"",
            "`"$stagingDir\*`"",
            "`"-xr@$excListFile`""
        )

        if ($PSCmdlet.ShouldProcess($ArchivePath, "Create profile archive")) {
            $proc = Start-Process -FilePath $exe7z -ArgumentList ($7zArgs -join ' ') -NoNewWindow -PassThru -Wait
            $exitCode = $proc.ExitCode

            if ($exitCode -eq 0) {
                Write-Log "Archive created successfully!" -Level SUCCESS
            } elseif ($exitCode -eq 1) {
                Write-Log "Archive completed with warnings (exit code 1). Locked files were skipped." -Level WARN
            } else {
                throw "7-Zip failed with exit code $exitCode."
            }

            $archInfo = Get-Item -LiteralPath $ArchivePath
            $sizeMb = [math]::Round($archInfo.Length / 1MB, 2)
            $sizeGb = [math]::Round($archInfo.Length / 1GB, 2)
            if ($sizeGb -ge 1) {
                Write-Log "Final Archive Size: $sizeGb GB" -Level INFO
            } else {
                Write-Log "Final Archive Size: $sizeMb MB" -Level INFO
            }
        }
    } finally {
        # Safely remove junctions before deleting staging directory
        foreach ($junc in $createdJunctions) {
            if (Test-Path -LiteralPath $junc) {
                try { [System.IO.Directory]::Delete($junc) } catch {}
            }
        }
        if (Test-Path -LiteralPath $stagingDir) {
            Remove-Item -LiteralPath $stagingDir -Recurse -Force -ErrorAction SilentlyContinue
        }
        if (Test-Path -LiteralPath $excListFile) {
            Remove-Item -LiteralPath $excListFile -Force -ErrorAction SilentlyContinue
        }
    }
}

# ==============================================================================
# RESTORE / MERGE MODE (New Laptop)
# ==============================================================================
elseif ($Mode -eq 'Restore') {
    if (-not $ArchivePath -or -not (Test-Path -LiteralPath $ArchivePath)) {
        throw "Archive file not found: $ArchivePath"
    }

    $targetProfile = if ($TargetProfile) { (Resolve-Path -LiteralPath $TargetProfile).Path } else { (Resolve-Path -LiteralPath $env:USERPROFILE).Path }
    $targetUser = [System.IO.Path]::GetFileName($targetProfile)
    $targetComputer = $env:COMPUTERNAME

    Write-Log "Initiating restore for user: $targetUser on $targetComputer" -Level INFO
    Write-Log "Target profile: $targetProfile" -Level INFO
    Write-Log "Conflict rule: Skip existing files (-aos / Robocopy /XC /XN /XO). Existing files are preserved." -Level INFO

    Test-OpenApplications

    # Verify archive integrity
    Write-Log "Verifying archive integrity..." -Level INFO
    $testProc = Start-Process -FilePath $exe7z -ArgumentList "t `"$ArchivePath`" -r MigrationManifest.json" -NoNewWindow -PassThru -Wait
    if ($testProc.ExitCode -gt 1) {
        throw "Archive integrity test failed (code $($testProc.ExitCode))."
    }

    $targetKnownMap = Get-KnownFolderRegistryMap -ProfileRoot $targetProfile
    Write-Log "Detected target Windows known folders (including OneDrive redirection):" -Level INFO
    foreach ($k in $targetKnownMap.Keys) {
        Write-Log "  + $k -> $($targetKnownMap[$k])" -Level INFO
    }

    if ($PSCmdlet.ShouldProcess($targetProfile, "Extract and merge archive $ArchivePath")) {
        # 1. Extract archive directly to target profile with -aos (skip existing)
        Write-Log "Extracting archive to profile root (skipping existing files)..." -Level INFO
        $extractArgs = @(
            'x',
            "`"$ArchivePath`"",
            "`"-o$targetProfile`"",
            '-aos',
            '-y'
        )
        $proc = Start-Process -FilePath $exe7z -ArgumentList ($extractArgs -join ' ') -NoNewWindow -PassThru -Wait
        if ($proc.ExitCode -gt 1) {
            throw "7-Zip extraction failed with code $($proc.ExitCode)."
        }

        # 2. Redirect Known Folders to active locations if redirected (e.g. OneDrive)
        foreach ($folderName in $targetKnownMap.Keys) {
            $resolvedTarget = $targetKnownMap[$folderName]
            $standardPath = Join-Path $targetProfile $folderName

            if ($resolvedTarget -and ($resolvedTarget -ne $standardPath) -and (Test-Path -LiteralPath $standardPath)) {
                Write-Log "Relocating extracted $folderName into active redirected path: $resolvedTarget" -Level INFO
                if (-not (Test-Path -LiteralPath $resolvedTarget)) {
                    New-Item -ItemType Directory -Path $resolvedTarget -Force | Out-Null
                }
                # Robocopy: /E (recursive), /XC /XN /XO (skip existing files), /MOVE (relocate and clean up source)
                robocopy "$standardPath" "$resolvedTarget" /E /XC /XN /XO /MOVE /R:1 /W:1 /NFL /NDL /NP | Out-Null
                if (Test-Path -LiteralPath $standardPath) {
                    Remove-Item -LiteralPath $standardPath -Recurse -Force -ErrorAction SilentlyContinue
                }
            }
        }

        # 3. Handle Outlook PST Files
        $extractedPstDir = Join-Path $targetProfile 'OutlookFiles'
        if (Test-Path -LiteralPath $extractedPstDir) {
            $docsRoot = if ($targetKnownMap['Documents']) { $targetKnownMap['Documents'] } else { Join-Path $targetProfile 'Documents' }
            $outlookTargetDir = Join-Path $docsRoot 'Outlook Files'
            if (-not (Test-Path -LiteralPath $outlookTargetDir)) {
                New-Item -ItemType Directory -Path $outlookTargetDir -Force | Out-Null
            }
            Write-Log "Merging Outlook PST files into $outlookTargetDir (skipping existing)..." -Level INFO
            robocopy "$extractedPstDir" "$outlookTargetDir" /E /XC /XN /XO /MOVE /R:1 /W:1 /NFL /NDL /NP | Out-Null
            Remove-Item -LiteralPath $extractedPstDir -Recurse -Force -ErrorAction SilentlyContinue
        }

        # 4. Handle Browser Bookmarks
        $extractedBookmarks = Join-Path $targetProfile 'BrowserBookmarks'
        if (Test-Path -LiteralPath $extractedBookmarks) {
            $docsRoot = if ($targetKnownMap['Documents']) { $targetKnownMap['Documents'] } else { Join-Path $targetProfile 'Documents' }
            $backupExportDir = Join-Path $docsRoot 'Migrated Browser Bookmarks'
            New-Item -ItemType Directory -Path $backupExportDir -Force | Out-Null
            $tgtLocal = Join-Path $targetProfile 'AppData\Local'

            # Chrome
            $chromeSrc = Join-Path $extractedBookmarks 'Chrome\Bookmarks'
            if (Test-Path -LiteralPath $chromeSrc) {
                $chromeTgt = Join-Path $tgtLocal 'Google\Chrome\User Data\Default\Bookmarks'
                Copy-Item -LiteralPath $chromeSrc -Destination (Join-Path $backupExportDir 'Chrome_Bookmarks_OldPC.json') -Force
                Merge-ChromiumBookmarks -SourceBookmarksFile $chromeSrc -TargetBookmarksFile $chromeTgt -BrowserName 'Google Chrome'
            }

            # Edge
            $edgeSrc = Join-Path $extractedBookmarks 'Edge\Bookmarks'
            if (Test-Path -LiteralPath $edgeSrc) {
                $edgeTgt = Join-Path $tgtLocal 'Microsoft\Edge\User Data\Default\Bookmarks'
                Copy-Item -LiteralPath $edgeSrc -Destination (Join-Path $backupExportDir 'Edge_Bookmarks_OldPC.json') -Force
                Merge-ChromiumBookmarks -SourceBookmarksFile $edgeSrc -TargetBookmarksFile $edgeTgt -BrowserName 'Microsoft Edge'
            }

            # Brave
            $braveSrc = Join-Path $extractedBookmarks 'Brave\Bookmarks'
            if (Test-Path -LiteralPath $braveSrc) {
                $braveTgt = Join-Path $tgtLocal 'BraveSoftware\Brave-Browser\User Data\Default\Bookmarks'
                Copy-Item -LiteralPath $braveSrc -Destination (Join-Path $backupExportDir 'Brave_Bookmarks_OldPC.json') -Force
                Merge-ChromiumBookmarks -SourceBookmarksFile $braveSrc -TargetBookmarksFile $braveTgt -BrowserName 'Brave'
            }

            # Firefox
            $fxSrc = Join-Path $extractedBookmarks 'Firefox'
            if (Test-Path -LiteralPath $fxSrc) {
                $fxExport = Join-Path $backupExportDir 'Firefox'
                New-Item -ItemType Directory -Path $fxExport -Force | Out-Null
                Copy-Item -LiteralPath "$fxSrc\*" -Destination $fxExport -Recurse -Force
                Write-Log "Firefox bookmark databases and backup files saved to: $fxExport" -Level INFO
                Write-Log "To restore in Firefox: Open Firefox -> Bookmarks -> Manage Bookmarks -> Import & Backup -> Restore -> Choose File..." -Level INFO
            }

            Remove-Item -LiteralPath $extractedBookmarks -Recurse -Force -ErrorAction SilentlyContinue
        }

        # 5. Manifest Summary & Cleanup
        $manifestPath = Join-Path $targetProfile 'MigrationManifest.json'
        if (Test-Path -LiteralPath $manifestPath) {
            try {
                $manifestInfo = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json
                Write-Log "Restored data from: $($manifestInfo.SourceUser)@$($manifestInfo.SourceComputer) (Backup Date: $($manifestInfo.BackupDateUtc))" -Level INFO
            } catch {}
            Remove-Item -LiteralPath $manifestPath -Force -ErrorAction SilentlyContinue
        }

        Write-Log "Profile migration and merge completed successfully!" -Level SUCCESS
        Write-Log "Have your colleague sign out and sign back in to refresh icons and app states." -Level SUCCESS
    }
}
