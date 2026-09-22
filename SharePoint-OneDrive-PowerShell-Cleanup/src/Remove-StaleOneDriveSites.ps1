[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [string[]]$NameOrPathPattern,
    [switch]$OnlyUnavailableLocalPath,
    [switch]$IncludePersonal,
    [switch]$RemoveRegistryEntries,
    [switch]$RemoveLocalFolders,
    [switch]$RestartOneDrive,
    [string]$BackupPath = (Join-Path -Path (Split-Path -Parent $PSScriptRoot) -ChildPath ("backup\onedrive-syncroot-backup-{0:yyyyMMdd-HHmmss}.reg" -f (Get-Date)))
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

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
    param(
        [Parameter(Mandatory)][string[]]$RegistryPaths,
        [Parameter(Mandatory)][string]$Destination
    )

    $backupFolder = Split-Path -Parent $Destination
    New-Item -Path $backupFolder -ItemType Directory -Force | Out-Null

    foreach ($registryPath in $RegistryPaths) {
        $nativePath = ConvertTo-NativeRegistryPath -Path $registryPath
        $safeName = ($nativePath -replace '[\\/:*?"<>| ]+', '_').Trim('_')
        $target = Join-Path -Path $backupFolder -ChildPath ("{0}-{1}" -f $safeName, (Split-Path -Leaf $Destination))

        if ($PSCmdlet.ShouldProcess($nativePath, "Export registry backup to $target")) {
            $output = & reg.exe export $nativePath $target /y 2>&1
            if ($LASTEXITCODE -ne 0) {
                Write-Warning "Kunne ikke eksportere $nativePath. reg.exe sagde: $output"
            }
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

function Get-SyncRootManagerEntries {
    $basePath = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\SyncRootManager'
    if (-not (Test-Path -LiteralPath $basePath)) {
        return @()
    }

    Get-ChildItem -LiteralPath $basePath | ForEach-Object {
        $syncRootKey = $_
        $syncRootId = $syncRootKey.PSChildName
        $isBusiness = $syncRootId -match 'Business|SPO|SharePoint|ODSP'

        if (-not $IncludePersonal -and -not $isBusiness) {
            return
        }

        $userSyncRootsPath = Join-Path -Path $syncRootKey.PSPath -ChildPath 'UserSyncRoots'
        foreach ($property in Get-RegistryValueProperties -Path $userSyncRootsPath) {
            $localPath = [Environment]::ExpandEnvironmentVariables([string]$property.Value)
            [pscustomobject]@{
                Source = 'SyncRootManager'
                AccountType = if ($isBusiness) { 'Business' } else { 'Personal' }
                DisplayName = $syncRootId
                SyncRootId = $syncRootId
                LocalPath = $localPath
                LocalPathExists = Test-Path -LiteralPath $localPath
                RegistryKeyPath = $syncRootKey.PSPath
                RegistryValueKeyPath = $userSyncRootsPath
                RegistryValueName = $property.Name
            }
        }
    }
}

function Get-OneDriveAccountCacheEntries {
    $basePath = 'HKCU:\Software\Microsoft\OneDrive\Accounts'
    if (-not (Test-Path -LiteralPath $basePath)) {
        return @()
    }

    Get-ChildItem -LiteralPath $basePath | ForEach-Object {
        $accountKey = $_
        $isBusiness = $accountKey.PSChildName -match '^Business'

        if (-not $IncludePersonal -and -not $isBusiness) {
            return
        }

        $cachePath = Join-Path -Path $accountKey.PSPath -ChildPath 'ScopeIdToMountPointPathCache'
        foreach ($property in Get-RegistryValueProperties -Path $cachePath) {
            $localPath = [Environment]::ExpandEnvironmentVariables([string]$property.Value)
            [pscustomobject]@{
                Source = 'OneDriveAccountCache'
                AccountType = if ($isBusiness) { 'Business' } else { 'Personal' }
                DisplayName = Split-Path -Leaf $localPath
                SyncRootId = $property.Name
                LocalPath = $localPath
                LocalPathExists = Test-Path -LiteralPath $localPath
                RegistryKeyPath = $cachePath
                RegistryValueKeyPath = $cachePath
                RegistryValueName = $property.Name
            }
        }
    }
}

function Select-StaleOneDriveEntry {
    param([Parameter(Mandatory, ValueFromPipeline)]$Entry)

    process {
        if ($OnlyUnavailableLocalPath -and $Entry.LocalPathExists) {
            return
        }

        if ($NameOrPathPattern -and $NameOrPathPattern.Count -gt 0) {
            $matched = $false
            foreach ($pattern in $NameOrPathPattern) {
                if ($Entry.DisplayName -like $pattern -or $Entry.LocalPath -like $pattern -or $Entry.SyncRootId -like $pattern) {
                    $matched = $true
                    break
                }
            }

            if (-not $matched) {
                return
            }
        }

        $Entry
    }
}

function Remove-OneDriveRegistryEntry {
    param([Parameter(Mandatory)]$Entry)

    if ($Entry.Source -eq 'SyncRootManager') {
        if ($PSCmdlet.ShouldProcess($Entry.RegistryKeyPath, 'Remove SyncRootManager key')) {
            Remove-Item -LiteralPath $Entry.RegistryKeyPath -Recurse -Force -ErrorAction SilentlyContinue
        }
        return
    }

    if ($PSCmdlet.ShouldProcess($Entry.RegistryValueKeyPath, "Remove registry value $($Entry.RegistryValueName)")) {
        Remove-ItemProperty -LiteralPath $Entry.RegistryValueKeyPath -Name $Entry.RegistryValueName -Force -ErrorAction SilentlyContinue
    }
}

function Remove-OneDriveLocalFolder {
    param([Parameter(Mandatory)]$Entry)

    if (-not $Entry.LocalPathExists) {
        return
    }

    if ($PSCmdlet.ShouldProcess($Entry.LocalPath, 'Remove local synced folder')) {
        Remove-Item -LiteralPath $Entry.LocalPath -Recurse -Force
    }
}

$entries = @(
    Get-SyncRootManagerEntries
    Get-OneDriveAccountCacheEntries
) | Where-Object { $_ } | Sort-Object Source, LocalPath, SyncRootId -Unique

$selectedEntries = @($entries | Select-StaleOneDriveEntry)

if (-not $RemoveRegistryEntries -and -not $RemoveLocalFolders -and -not $RestartOneDrive) {
    $selectedEntries | Sort-Object AccountType, Source, DisplayName | Format-Table Source, AccountType, DisplayName, LocalPathExists, LocalPath -AutoSize
    return
}

if ($selectedEntries.Count -eq 0) {
    Write-Warning 'Ingen OneDrive/SharePoint sync-poster matchede filteret.'
    return
}

if ($RemoveRegistryEntries) {
    Export-RegistryBackup -RegistryPaths @(
        'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\SyncRootManager',
        'HKCU:\Software\Microsoft\OneDrive\Accounts'
    ) -Destination $BackupPath

    foreach ($entry in $selectedEntries) {
        Remove-OneDriveRegistryEntry -Entry $entry
    }
}

if ($RemoveLocalFolders) {
    foreach ($entry in ($selectedEntries | Sort-Object LocalPath -Unique)) {
        Remove-OneDriveLocalFolder -Entry $entry
    }
}

if ($RestartOneDrive) {
    $oneDrivePath = Get-OneDriveExecutablePath
    if (-not $oneDrivePath) {
        Write-Warning 'Kunne ikke finde OneDrive.exe. Genstart OneDrive manuelt.'
        return
    }

    if ($PSCmdlet.ShouldProcess('OneDrive', 'Restart OneDrive client')) {
        Get-Process -Name OneDrive -ErrorAction SilentlyContinue | Stop-Process -Force
        Start-Process -FilePath $oneDrivePath
    }
}

$selectedEntries | Sort-Object AccountType, Source, DisplayName | Format-Table Source, AccountType, DisplayName, LocalPathExists, LocalPath -AutoSize
