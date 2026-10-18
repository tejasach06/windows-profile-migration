# Windows profile migration

PowerShell script for backing up a Windows user profile to a 7-Zip archive and merging its files into another computer's profile.

## Run a migration

Install 7-Zip on both computers. Sign in as the user whose profile you want to migrate. Close browsers, Outlook, and other applications before each operation.

On the old computer:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\Migrate-UserProfile.ps1" -Mode Backup -ArchivePath "E:\ProfileBackup.7z"
```

Copy the script and archive to the new computer. Sign in as the destination user, then run:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\Migrate-UserProfile.ps1" -Mode Restore -ArchivePath "E:\ProfileBackup.7z"
```

Replace `E:` with the migration drive letter. The execution-policy setting applies to this PowerShell process. Organization policy can still block execution.

## Included data

- Desktop, Documents, Downloads, Pictures, Music, Videos, and Favorites, with registry-based folder resolution for the signed-in user.
- Additional top-level profile directories, including hidden directories, except reserved folders and top-level reparse points.
- AppData\\Roaming, subject to cache and temporary-file exclusions.
- Outlook PST files in the standard local Outlook directory and files already included through Documents.
- Default-profile Chrome, Edge, and Brave bookmarks, plus Firefox bookmark databases and backups.

Existing destination files are skipped during extraction. Chromium bookmarks are an exception: the script attempts to add old bookmarks to a folder inside the existing bookmark file. Firefox recovery files are saved under `Documents\Migrated Browser Bookmarks`.

## Limits and precautions

- This is a file migration, not a Windows account clone. It does not install applications or migrate registry settings, account identities, or guaranteed working passwords and sessions.
- Archives are not encrypted. Hidden profile directories can contain SSH keys and other sensitive files. Protect the archive and migration drive.
- Keep the old computer and a separate backup until you verify the restored files and applications.
- OneDrive cloud-only files must be downloaded before backup. Redirection detection does not prove that every cloud file is available locally.
- The script excludes names such as `Cache`, `Temp`, and `*.tmp` throughout the archive, including user-created directories.
- Review the script before use. Redirected-folder restore currently moves and removes the conventional profile folder, and cleanup does not check Robocopy exit codes. Do not use that path on a profile containing important files without a separate backup.
- `-WhatIf` is not a reliable end-to-end preview because native commands do not consistently honor it.

## Options

Use `-SevenZipPath "C:\Program Files\7-Zip\7z.exe"` to specify the archiver. Use `-ExcludeDownloads` to omit Downloads. `-SourceProfile` and `-TargetProfile` override the profile roots, but only the signed-in profile receives registry-based known-folder resolution.

Listed in the [Helper Scripts catalog](https://github.com/tejasach06/helper-scripts).
