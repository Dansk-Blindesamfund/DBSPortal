[CmdletBinding()]
param(
    [string]$UserPrincipalName,
    [string]$SharePointHostName,
    [string[]]$NameOrPathPattern,
    [switch]$OnlyUnavailableLocalPath,
    [switch]$IncludePersonal,
    [string]$OutputPath
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

function Select-AccessCheckEntry {
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

        if ($Entry.AccountType -ne 'Business') {
            return
        }

        $Entry
    }
}

function Get-CurrentUserPrincipalName {
    $upn = $null
    $whoami = & whoami /upn 2>$null
    if ($LASTEXITCODE -eq 0 -and $whoami -and $whoami.Count -gt 0) {
        $upn = ([string]$whoami[0]).Trim()
    }

    if ([string]::IsNullOrWhiteSpace($upn)) {
        $upn = $env:USERNAME
    }

    return $upn
}

if ([string]::IsNullOrWhiteSpace($UserPrincipalName)) {
    $UserPrincipalName = Get-CurrentUserPrincipalName
}

if ([string]::IsNullOrWhiteSpace($UserPrincipalName)) {
    throw 'UserPrincipalName er ikke angivet og kunne ikke bestemmes automatisk.'
}

if ([string]::IsNullOrWhiteSpace($SharePointHostName)) {
    throw 'SharePointHostName er påkrævet. Angiv fx "blindesamfund.sharepoint.com".'
}

$entries = @(
    Get-SyncRootManagerEntries
    Get-OneDriveAccountCacheEntries
) | Where-Object { $_ } | Sort-Object Source, LocalPath, SyncRootId -Unique

$selectedEntries = @($entries | Select-AccessCheckEntry)

if ($selectedEntries.Count -eq 0) {
    Write-Warning 'Ingen relevante OneDrive/SharePoint sync-poster blev fundet.'
    return
}

$payload = [pscustomobject]@{
    sharePointHostName = $SharePointHostName
    userPrincipalName = $UserPrincipalName
    entries = @(
        $selectedEntries | ForEach-Object {
            [pscustomobject]@{
                source = $_.Source
                displayName = $_.DisplayName
                localPath = $_.LocalPath
                siteUrl = ''
            }
        }
    )
}

$json = $payload | ConvertTo-Json -Depth 8

if (-not [string]::IsNullOrWhiteSpace($OutputPath)) {
    $parentFolder = Split-Path -Parent $OutputPath
    if (-not [string]::IsNullOrWhiteSpace($parentFolder)) {
        New-Item -Path $parentFolder -ItemType Directory -Force | Out-Null
    }

    $json | Set-Content -Path $OutputPath -Encoding UTF8
    Write-Host "Payload gemt til $OutputPath"
}

$json
