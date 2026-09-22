[CmdletBinding()]
param(
    [switch]$IUnderstandThisDeletesLocalSyncedFolder
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$TargetFolderName = 'Kursusafdelingen - Dokumenter'
$TargetFolderPatterns = @(
    'Kursusafdelingen - Dokumenter',
    'Kursusafdelingen*Dokumenter',
    '*Kursusafdelingen*'
)
$ScriptVersion = '2026-09-18-05'
$RestartOneDrive = $true
$LogRoot = Join-Path -Path $env:LOCALAPPDATA -ChildPath 'OneDriveSharePointCleanup'
$LogPath = Join-Path -Path $LogRoot -ChildPath 'Kursusafdelingen-RealTest.log'

function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO', 'WARN', 'ERROR')][string]$Level = 'INFO'
    )

    if (-not (Test-Path -LiteralPath $LogRoot)) {
        New-Item -Path $LogRoot -ItemType Directory -Force | Out-Null
    }

    $line = '{0:s} [{1}] {2}' -f (Get-Date), $Level, $Message
    Add-Content -LiteralPath $LogPath -Value $line -Encoding UTF8
    Write-Host $line
}

function Get-RegistryValueProperties {
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        return @()
    }

    $item = Get-ItemProperty -LiteralPath $Path
    return $item.PSObject.Properties | Where-Object {
        $_.Name -notin @('PSPath', 'PSParentPath', 'PSChildName', 'PSDrive', 'PSProvider') -and
        $null -ne $_.Value -and
        [string]$_.Value -ne ''
    }
}

function ConvertTo-NativeRegistryPath {
    param([Parameter(Mandatory)][string]$Path)

    return $Path -replace '^Microsoft\.PowerShell\.Core\\Registry::HKEY_CURRENT_USER', 'HKCU' `
                 -replace '^HKCU:', 'HKCU' `
                 -replace '/', '\'
}

function Export-RegistryBackup {
    $timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    foreach ($registryPath in @('HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\SyncRootManager', 'HKCU:\Software\Microsoft\OneDrive\Accounts')) {
        if (-not (Test-Path -LiteralPath $registryPath)) {
            continue
        }

        $nativePath = ConvertTo-NativeRegistryPath -Path $registryPath
        $safeName = ($nativePath -replace '[\\/:*?"<>| ]+', '_').Trim('_')
        $target = Join-Path -Path $LogRoot -ChildPath ("{0}-{1}.reg" -f $safeName, $timestamp)
        $output = & reg.exe export $nativePath $target /y 2>&1
        if ($LASTEXITCODE -eq 0) {
            Write-Log "Registry-backup gemt: $target"
        } else {
            Write-Log "Kunne ikke eksportere $nativePath. reg.exe sagde: $output" 'WARN'
        }
    }
}

function Get-OneDriveExecutablePath {
    $process = Get-Process -Name OneDrive -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($process -and $process.Path) {
        return $process.Path
    }

    $candidates = @(
        (Join-Path $env:LOCALAPPDATA 'Microsoft\OneDrive\OneDrive.exe'),
        (Join-Path ${env:ProgramFiles} 'Microsoft OneDrive\OneDrive.exe'),
        (Join-Path ${env:ProgramFiles(x86)} 'Microsoft OneDrive\OneDrive.exe')
    )

    return $candidates | Where-Object { $_ -and (Test-Path -LiteralPath $_) } | Select-Object -First 1
}

function Test-TargetLocalPath {
    param([Parameter(Mandatory)][string]$LocalPath)

    $folderName = Split-Path -Leaf $LocalPath
    foreach ($pattern in $TargetFolderPatterns) {
        if ($folderName -like $pattern -or $LocalPath -like "*$pattern*") {
            return $true
        }
    }

    return $false
}

function Get-TargetSyncEntries {
    $entries = New-Object System.Collections.Generic.List[object]
    $allBusinessPaths = New-Object System.Collections.Generic.List[string]
    $syncRootManagerPath = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\SyncRootManager'
    if (Test-Path -LiteralPath $syncRootManagerPath) {
        foreach ($syncRootKey in Get-ChildItem -LiteralPath $syncRootManagerPath) {
            $syncRootId = $syncRootKey.PSChildName
            if ($syncRootId -notmatch 'Business|SPO|SharePoint|ODSP') {
                continue
            }

            $userSyncRootsPath = Join-Path -Path $syncRootKey.PSPath -ChildPath 'UserSyncRoots'
            foreach ($property in Get-RegistryValueProperties -Path $userSyncRootsPath) {
                $localPath = [Environment]::ExpandEnvironmentVariables([string]$property.Value)
                $allBusinessPaths.Add("SyncRootManager: $localPath")
                if (Test-TargetLocalPath -LocalPath $localPath) {
                    $entries.Add([pscustomobject]@{
                        Source = 'SyncRootManager'
                        LocalPath = $localPath
                        RegistryKeyPath = $syncRootKey.PSPath
                        RegistryValueKeyPath = $userSyncRootsPath
                        RegistryValueName = $property.Name
                    })
                }
            }
        }
    }

    $accountsPath = 'HKCU:\Software\Microsoft\OneDrive\Accounts'
    if (Test-Path -LiteralPath $accountsPath) {
        foreach ($accountKey in Get-ChildItem -LiteralPath $accountsPath | Where-Object { $_.PSChildName -like 'Business*' }) {
            $cachePath = Join-Path -Path $accountKey.PSPath -ChildPath 'ScopeIdToMountPointPathCache'
            foreach ($property in Get-RegistryValueProperties -Path $cachePath) {
                $localPath = [Environment]::ExpandEnvironmentVariables([string]$property.Value)
                $allBusinessPaths.Add("OneDriveAccountCache: $localPath")
                if (Test-TargetLocalPath -LocalPath $localPath) {
                    $entries.Add([pscustomobject]@{
                        Source = 'OneDriveAccountCache'
                        LocalPath = $localPath
                        RegistryKeyPath = $cachePath
                        RegistryValueKeyPath = $cachePath
                        RegistryValueName = $property.Name
                    })
                }
            }
        }
    }

    foreach ($businessPath in ($allBusinessPaths | Sort-Object -Unique)) {
        Write-Log "Set sync-sti: $businessPath"
    }

    $knownTenantRoot = Join-Path -Path $env:USERPROFILE -ChildPath 'Dansk Blindesamfund'
    if (Test-Path -LiteralPath $knownTenantRoot) {
        foreach ($folder in Get-ChildItem -LiteralPath $knownTenantRoot -Directory -ErrorAction SilentlyContinue) {
            Write-Log "Set lokal tenant-mappe: $($folder.FullName)"
            if (Test-TargetLocalPath -LocalPath $folder.FullName) {
                $alreadyMatched = @($entries | Where-Object { $_.LocalPath -eq $folder.FullName }).Count -gt 0
                if (-not $alreadyMatched) {
                    $entries.Add([pscustomobject]@{
                        Source = 'FileSystemOnly'
                        LocalPath = $folder.FullName
                        RegistryKeyPath = $null
                        RegistryValueKeyPath = $null
                        RegistryValueName = $null
                    })
                }
            }
        }
    } else {
        Write-Log "Lokal tenant-rodmappe findes ikke: $knownTenantRoot" 'WARN'
    }

    $entries | Sort-Object Source, LocalPath, RegistryValueName -Unique
}

function Remove-SyncRegistryEntry {
    param([Parameter(Mandatory)]$Entry)

    if ($Entry.Source -eq 'FileSystemOnly') {
        Write-Log "Ingen matchende registry-post fundet for filesystem-match: $($Entry.LocalPath)" 'WARN'
        return
    }

    if ($Entry.Source -eq 'SyncRootManager') {
        Remove-Item -LiteralPath $Entry.RegistryKeyPath -Recurse -Force -ErrorAction SilentlyContinue
        Write-Log "Fjernede SyncRootManager-nøgle: $($Entry.RegistryKeyPath)"
        return
    }

    Remove-ItemProperty -LiteralPath $Entry.RegistryValueKeyPath -Name $Entry.RegistryValueName -Force -ErrorAction SilentlyContinue
    Write-Log "Fjernede OneDrive cache-værdi: $($Entry.RegistryValueName)"
}

function Remove-LocalFolderRobust {
    param([Parameter(Mandatory)][string]$LocalPath)

    if (-not (Test-Path -LiteralPath $LocalPath)) {
        Write-Log "Lokal mappe fandtes ikke: $LocalPath"
        return
    }

    try {
        Get-ChildItem -LiteralPath $LocalPath -Force -Recurse -ErrorAction SilentlyContinue |
            ForEach-Object { $_.Attributes = 'Normal' }
        (Get-Item -LiteralPath $LocalPath -Force).Attributes = 'Normal'
    } catch {
        Write-Log "Kunne ikke nulstille filattributter for ${LocalPath}: $($_.Exception.Message)" 'WARN'
    }

    try {
        Remove-Item -LiteralPath $LocalPath -Recurse -Force -ErrorAction Stop
    } catch {
        Write-Log "Remove-Item fejlede for ${LocalPath}: $($_.Exception.Message). Prøver cmd.exe rmdir fallback." 'WARN'
        $cmdOutput = & cmd.exe /c rd /s /q "\\?\$LocalPath" 2>&1
        if ($LASTEXITCODE -ne 0) {
            Write-Log "cmd.exe rmdir fallback returnerede exitkode $LASTEXITCODE for $LocalPath. Output: $cmdOutput" 'WARN'
        }
    }

    if (Test-Path -LiteralPath $LocalPath) {
        $quarantinePath = Join-Path -Path (Split-Path -Parent $LocalPath) -ChildPath ("_OneDriveCleanup_DeleteMe_{0:yyyyMMddHHmmss}" -f (Get-Date))
        try {
            Rename-Item -LiteralPath $LocalPath -NewName (Split-Path -Leaf $quarantinePath) -Force -ErrorAction Stop
            Write-Log "Mappen kunne ikke slettes direkte og blev omdøbt til: $quarantinePath" 'WARN'

            $cmdOutput = & cmd.exe /c rd /s /q "\\?\$quarantinePath" 2>&1
            if ($LASTEXITCODE -ne 0 -and (Test-Path -LiteralPath $quarantinePath)) {
                throw "cmd.exe rmdir kunne heller ikke slette karantænemappen. Output: $cmdOutput"
            }
        } catch {
            throw "Kunne ikke slette eller omdøbe ${LocalPath}. $($_.Exception.Message)"
        }
    }

    if (Test-Path -LiteralPath $LocalPath) {
        throw "Mappen findes stadig efter sletningsforsøg: $LocalPath"
    }

    Write-Log "Slettede lokal mappe: $LocalPath"
}

try {
    if (-not $IUnderstandThisDeletesLocalSyncedFolder) {
        Write-Log 'Dette kontrollerede REAL TEST-script er deaktiveret som standard. Det sletter ikke noget uden -IUnderstandThisDeletesLocalSyncedFolder.' 'WARN'
        exit 0
    }

    Write-Log "Starter kontrolleret REAL TEST version $ScriptVersion for $TargetFolderName."
    Write-Log "Kører som bruger: $([Security.Principal.WindowsIdentity]::GetCurrent().Name) / USERPROFILE=$env:USERPROFILE / LOCALAPPDATA=$env:LOCALAPPDATA"
    $entries = @(Get-TargetSyncEntries)
    if ($entries.Count -eq 0) {
        Write-Log "Ingen sync-poster fundet for $TargetFolderName."
        exit 0
    }

    foreach ($entry in $entries) {
        Write-Log "Matcher sync-post: $($entry.Source) / $($entry.LocalPath)"
    }

    Export-RegistryBackup
    Get-Process -Name OneDrive -ErrorAction SilentlyContinue | Stop-Process -Force

    foreach ($entry in $entries) {
        Remove-SyncRegistryEntry -Entry $entry
    }

    foreach ($localPath in @($entries.LocalPath | Sort-Object -Unique)) {
        Remove-LocalFolderRobust -LocalPath $localPath
    }

    if ($RestartOneDrive) {
        $oneDrivePath = Get-OneDriveExecutablePath
        if ($oneDrivePath) {
            Start-Process -FilePath $oneDrivePath
            Write-Log 'OneDrive blev startet igen.'
        } else {
            Write-Log 'OneDrive.exe blev ikke fundet. Start OneDrive manuelt.' 'WARN'
        }
    }

    Write-Log "Kontrolleret REAL TEST afsluttet for $TargetFolderName."
    exit 0
} catch {
    Write-Log $_.Exception.Message 'ERROR'
    exit 1
}
