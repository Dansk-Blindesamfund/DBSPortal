[CmdletBinding()]
param(
    [switch]$ScheduledRun
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Intune configuration. Upload this file as a platform script and edit these values first.
$TenantId = '7269da2b-d73e-4791-ba91-3675fa4b83f0'
$ClientId = 'b90b44fd-4ef4-4f70-82ab-e548aa190281'
$CertificateThumbprint = ''
$CertificateStorePath = 'Cert:\LocalMachine\My'
$SharePointHostName = 'blindesamfund.sharepoint.com'
$DryRun = $false
$RemoveLocalFolders = $true
$RestartOneDrive = $false
$ScriptVersion = '2026-09-21-v18-live-test'
$TenantFolderName = 'Dansk Blindesamfund'
$AllowFileSystemOnlyDeletion = $false
$EnablePeriodicCheck = $false
$PeriodicCheckMinutes = 30
$AccessCheckFunctionUrl = 'https://func-dbs-onedrive-cleanup-test.azurewebsites.net/api/sharepoint/access-check'
$AccessCheckFunctionKey = $env:DBS_ACCESS_CHECK_FUNCTION_KEY

# Optional explicit mappings. Recommended for the first production rollout.
# Example:
# $SiteMappings = @(
#     @{ LocalPathPattern = '*Fælles - Dokumenter*'; SiteUrl = 'https://contoso.sharepoint.com/sites/Faelles' }
# )
$SiteMappings = @(
    @{ LocalPathPattern = '*IT - Dokumenter*'; SiteUrl = 'https://blindesamfund.sharepoint.com/sites/IT' },
    @{ LocalPathPattern = '*Kursusafdelingen*'; SiteUrl = 'https://blindesamfund.sharepoint.com/sites/Kursus' }
)

$logBasePath = $env:LOCALAPPDATA
if ([string]::IsNullOrWhiteSpace($logBasePath)) {
    $logBasePath = Join-Path -Path $env:ProgramData -ChildPath 'OneDriveSharePointCleanup'
}

$LogRoot = Join-Path -Path $logBasePath -ChildPath 'OneDriveSharePointCleanup'
$LogPath = Join-Path -Path $LogRoot -ChildPath 'AccessCheckCleanup.log'
$BackupPath = Join-Path -Path $LogRoot -ChildPath ("onedrive-syncroot-backup-{0:yyyyMMdd-HHmmss}.reg" -f (Get-Date))

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

function ConvertTo-Base64Url {
    param([Parameter(Mandatory)][byte[]]$Bytes)

    return [Convert]::ToBase64String($Bytes).TrimEnd('=').Replace('+', '-').Replace('/', '_')
}

function New-ClientAssertion {
    param(
        [Parameter(Mandatory)][string]$Authority,
        [Parameter(Mandatory)][string]$ApplicationId,
        [Parameter(Mandatory)][System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificate
    )

    $now = [DateTimeOffset]::UtcNow
    $header = @{
        alg = 'RS256'
        typ = 'JWT'
        x5t = ConvertTo-Base64Url -Bytes $Certificate.GetCertHash()
    } | ConvertTo-Json -Compress

    $payload = @{
        aud = $Authority
        exp = $now.AddMinutes(10).ToUnixTimeSeconds()
        iss = $ApplicationId
        jti = [guid]::NewGuid().ToString()
        nbf = $now.AddMinutes(-1).ToUnixTimeSeconds()
        sub = $ApplicationId
    } | ConvertTo-Json -Compress

    $unsignedToken = '{0}.{1}' -f `
        (ConvertTo-Base64Url -Bytes ([Text.Encoding]::UTF8.GetBytes($header))), `
        (ConvertTo-Base64Url -Bytes ([Text.Encoding]::UTF8.GetBytes($payload)))

    $rsa = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($Certificate)
    if (-not $rsa) {
        throw 'Certifikatet har ikke en tilgængelig privat nøgle.'
    }

    $signature = $rsa.SignData(
        [Text.Encoding]::UTF8.GetBytes($unsignedToken),
        [Security.Cryptography.HashAlgorithmName]::SHA256,
        [Security.Cryptography.RSASignaturePadding]::Pkcs1
    )

    return '{0}.{1}' -f $unsignedToken, (ConvertTo-Base64Url -Bytes $signature)
}

function Get-AppOnlyAccessToken {
    param(
        [Parameter(Mandatory)][string]$ResourceScope,
        [Parameter(Mandatory)][System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificate
    )

    $authority = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
    $assertion = New-ClientAssertion -Authority $authority -ApplicationId $ClientId -Certificate $Certificate
    $body = @{
        client_id = $ClientId
        client_assertion = $assertion
        client_assertion_type = 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer'
        grant_type = 'client_credentials'
        scope = $ResourceScope
    }

    return (Invoke-RestMethod -Method Post -Uri $authority -Body $body -ContentType 'application/x-www-form-urlencoded').access_token
}

function Get-CleanupCertificate {
    if ([string]::IsNullOrWhiteSpace($CertificateThumbprint)) {
        throw 'CertificateThumbprint mangler i script-konfigurationen.'
    }

    $path = Join-Path -Path $CertificateStorePath -ChildPath $CertificateThumbprint
    $certificate = Get-Item -LiteralPath $path -ErrorAction Stop
    if (-not $certificate.HasPrivateKey) {
        throw "Certifikatet $CertificateThumbprint har ingen privat nøgle på klienten."
    }

    return $certificate
}

function Get-CurrentUserPrincipalName {
    $accountPath = 'HKCU:\Software\Microsoft\OneDrive\Accounts'
    if (Test-Path -LiteralPath $accountPath) {
        foreach ($account in Get-ChildItem -LiteralPath $accountPath | Where-Object { $_.PSChildName -like 'Business*' }) {
            $properties = Get-ItemProperty -LiteralPath $account.PSPath
            foreach ($name in @('UserEmail', 'UserName')) {
                if ($properties.PSObject.Properties.Name -contains $name -and $properties.$name -match '@') {
                    return [string]$properties.$name
                }
            }
        }
    }

    $whoAmI = (& whoami.exe /upn 2>$null)
    if ($LASTEXITCODE -eq 0 -and $whoAmI -match '@') {
        return [string]$whoAmI.Trim()
    }

    throw 'Kunne ikke finde brugerens UPN. Kør scriptet i bruger-kontekst via Intune.'
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
    param([Parameter(Mandatory)][string[]]$RegistryPaths)

    foreach ($registryPath in $RegistryPaths) {
        if (-not (Test-Path -LiteralPath $registryPath)) {
            continue
        }

        $nativePath = ConvertTo-NativeRegistryPath -Path $registryPath
        $safeName = ($nativePath -replace '[\\/:*?"<>| ]+', '_').Trim('_')
        $target = Join-Path -Path $LogRoot -ChildPath ("{0}-{1}" -f $safeName, (Split-Path -Leaf $BackupPath))
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

function Restart-OneDriveAsStandardUser {
    param([Parameter(Mandatory)][string]$OneDrivePath)

    $taskName = "OneDriveCleanupRestart-$([guid]::NewGuid().ToString('N'))"
    $userId = [Security.Principal.WindowsIdentity]::GetCurrent().Name
    $action = New-ScheduledTaskAction -Execute $OneDrivePath
    $principal = New-ScheduledTaskPrincipal -UserId $userId -LogonType Interactive -RunLevel Limited
    $task = New-ScheduledTask -Action $action -Principal $principal -Description 'Start OneDrive without elevation after SharePoint cleanup.'

    try {
        Register-ScheduledTask -TaskName $taskName -InputObject $task -Force | Out-Null
        Start-ScheduledTask -TaskName $taskName
        Write-Log "OneDrive restart planlagt uden administratorrettigheder for $userId."
    } finally {
        Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue
    }
}

function Install-PeriodicCheckTask {
    if ($ScheduledRun -or -not $EnablePeriodicCheck) {
        return
    }

    if ([string]::IsNullOrWhiteSpace($PSCommandPath) -or -not (Test-Path -LiteralPath $PSCommandPath)) {
        Write-Log 'Periodic check kunne ikke installeres, fordi scriptets sti ikke er tilgængelig.' 'WARN'
        return
    }

    $workerRoot = Join-Path -Path $env:LOCALAPPDATA -ChildPath 'OneDriveSharePointCleanup'
    $workerPath = Join-Path -Path $workerRoot -ChildPath 'AccessCheckWorker.ps1'
    $taskName = 'OneDrive SharePoint Access Check'
    $userId = [Security.Principal.WindowsIdentity]::GetCurrent().Name

    New-Item -Path $workerRoot -ItemType Directory -Force | Out-Null
    Copy-Item -LiteralPath $PSCommandPath -Destination $workerPath -Force

    $action = New-ScheduledTaskAction -Execute (Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe') -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$workerPath`" -ScheduledRun"
    $principal = New-ScheduledTaskPrincipal -UserId $userId -LogonType InteractiveToken -RunLevel Limited
    $logonTrigger = New-ScheduledTaskTrigger -AtLogOn -User $userId
    $repeatTrigger = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(1) -RepetitionInterval ([TimeSpan]::FromMinutes($PeriodicCheckMinutes)) -RepetitionDuration ([TimeSpan]::FromDays(3650))
    $task = New-ScheduledTask -Action $action -Trigger @($logonTrigger, $repeatTrigger) -Principal $principal -Description 'Checks SharePoint access and removes inaccessible registry-backed OneDrive sync entries.'

    Register-ScheduledTask -TaskName $taskName -InputObject $task -Force | Out-Null
    Write-Log "Periodic adgangstjek installeret for $userId. Interval: $PeriodicCheckMinutes minutter."
}

function Get-SiteSearchQueries {
    param([Parameter(Mandatory)][string]$DisplayName)

    $queries = New-Object System.Collections.Generic.List[string]
    $cleanName = $DisplayName.Trim()
    if (-not [string]::IsNullOrWhiteSpace($cleanName)) {
        $queries.Add($cleanName)
    }

    $withoutLibrarySuffix = ($cleanName -replace '\s+-\s+(Dokumenter|Documents|General|Delte dokumenter|Shared Documents)$', '').Trim()
    if (-not [string]::IsNullOrWhiteSpace($withoutLibrarySuffix) -and $withoutLibrarySuffix -ne $cleanName) {
        $queries.Add($withoutLibrarySuffix)
    }

    $beforeDash = ($cleanName -replace '\s+-\s+.*$', '').Trim()
    if (-not [string]::IsNullOrWhiteSpace($beforeDash) -and $beforeDash -ne $cleanName -and $beforeDash -ne $withoutLibrarySuffix) {
        $queries.Add($beforeDash)
    }

    return $queries | Select-Object -Unique
}

function Add-SiteMapping {
    param([Parameter(Mandatory)]$Entry)

    foreach ($mapping in $SiteMappings) {
        if ($Entry.LocalPath -like $mapping.LocalPathPattern -or $Entry.DisplayName -like $mapping.LocalPathPattern) {
            $Entry.SiteUrl = $mapping.SiteUrl.TrimEnd('/')
            return
        }
    }
}

function Get-SyncedSharePointEntry {
    $entries = New-Object System.Collections.Generic.List[object]
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
                $entries.Add([pscustomobject]@{
                    Source = 'SyncRootManager'
                    DisplayName = Split-Path -Leaf $localPath
                    SyncRootId = $syncRootId
                    LocalPath = $localPath
                    LocalPathExists = Test-Path -LiteralPath $localPath
                    RegistryKeyPath = $syncRootKey.PSPath
                    RegistryValueKeyPath = $userSyncRootsPath
                    RegistryValueName = $property.Name
                    SiteUrl = $null
                })
            }
        }
    }

    $accountsPath = 'HKCU:\Software\Microsoft\OneDrive\Accounts'
    if (Test-Path -LiteralPath $accountsPath) {
        foreach ($accountKey in Get-ChildItem -LiteralPath $accountsPath | Where-Object { $_.PSChildName -like 'Business*' }) {
            $cachePath = Join-Path -Path $accountKey.PSPath -ChildPath 'ScopeIdToMountPointPathCache'
            foreach ($property in Get-RegistryValueProperties -Path $cachePath) {
                $localPath = [Environment]::ExpandEnvironmentVariables([string]$property.Value)
                $entries.Add([pscustomobject]@{
                    Source = 'OneDriveAccountCache'
                    DisplayName = Split-Path -Leaf $localPath
                    SyncRootId = $property.Name
                    LocalPath = $localPath
                    LocalPathExists = Test-Path -LiteralPath $localPath
                    RegistryKeyPath = $cachePath
                    RegistryValueKeyPath = $cachePath
                    RegistryValueName = $property.Name
                    SiteUrl = $null
                })
            }
        }
    }

    $tenantRoot = Join-Path -Path $env:USERPROFILE -ChildPath $TenantFolderName
    if (Test-Path -LiteralPath $tenantRoot) {
        foreach ($folder in Get-ChildItem -LiteralPath $tenantRoot -Directory -ErrorAction SilentlyContinue) {
            $alreadyKnown = @($entries | Where-Object { $_.LocalPath -eq $folder.FullName }).Count -gt 0
            if ($alreadyKnown) {
                continue
            }

            $entries.Add([pscustomobject]@{
                Source = 'FileSystemOnly'
                DisplayName = $folder.Name
                SyncRootId = $null
                LocalPath = $folder.FullName
                LocalPathExists = $true
                RegistryKeyPath = $null
                RegistryValueKeyPath = $null
                RegistryValueName = $null
                SiteUrl = $null
            })
        }
    } else {
        Write-Log "Lokal tenant-rodmappe findes ikke: $tenantRoot" 'WARN'
    }

    foreach ($entry in ($entries | Sort-Object LocalPath, Source, SyncRootId -Unique)) {
        Add-SiteMapping -Entry $entry
        $entry
    }
}

function Search-GraphSiteUrl {
    param(
        [Parameter(Mandatory)][string]$DisplayName,
        [Parameter(Mandatory)][string]$GraphToken
    )

    if ([string]::IsNullOrWhiteSpace($SharePointHostName)) {
        return $null
    }

    foreach ($query in Get-SiteSearchQueries -DisplayName $DisplayName) {
        $uri = 'https://graph.microsoft.com/v1.0/sites?search={0}' -f [uri]::EscapeDataString($query)
        $response = Invoke-RestMethod -Method Get -Uri $uri -Headers @{ Authorization = "Bearer $GraphToken" }
        $matches = @($response.value | Where-Object { $_.webUrl -like "https://$SharePointHostName/*" })

        $siteName = ($DisplayName -replace '\s+-\s+.*$', '').Trim()
        $exactMatches = @($matches | Where-Object { $_.displayName -eq $siteName })
        if ($exactMatches.Count -eq 1) {
            Write-Log "Site-URL fundet via eksakt Graph-match '$siteName': $($exactMatches[0].webUrl)"
            return [string]$exactMatches[0].webUrl.TrimEnd('/')
        }

        if ($matches.Count -eq 1) {
            Write-Log "Site-URL fundet via Graph search '$query': $($matches[0].webUrl)"
            return [string]$matches[0].webUrl.TrimEnd('/')
        }

        if ($matches.Count -gt 1) {
            Write-Log "Graph search '$query' gav flere site-match for $DisplayName. Springer over for at undgå forkert sletning." 'WARN'
        }
    }

    return $null
}

function Test-SharePointUserAccess {
    param(
        [Parameter(Mandatory)][string]$SiteUrl,
        [Parameter(Mandatory)][string]$UserPrincipalName,
        [Parameter(Mandatory)][string]$DisplayName,
        [Parameter(Mandatory)][string]$SharePointToken
    )

    $claim = "i:0#.f|membership|$UserPrincipalName"
    $encodedClaim = [uri]::EscapeDataString("'$claim'")
    $libraryMatch = [regex]::Match($DisplayName.Trim(), '^.+?\s+-\s+(.+)$')
    if ($libraryMatch.Success) {
        $libraryName = $libraryMatch.Groups[1].Value.Trim()
        $encodedLibraryName = [uri]::EscapeDataString($libraryName.Replace("'", "''"))
        $uri = "{0}/_api/web/lists/getbytitle('{1}')/GetUserEffectivePermissions(@u)?@u={2}" -f $SiteUrl.TrimEnd('/'), $encodedLibraryName, $encodedClaim
        Write-Log "Tester brugeradgang til dokumentbiblioteket '$libraryName' på $SiteUrl."
    } else {
        $uri = "{0}/_api/web/GetUserEffectivePermissions(@u)?@u={1}" -f $SiteUrl.TrimEnd('/'), $encodedClaim
        Write-Log "Tester brugeradgang til SharePoint-site $SiteUrl."
    }

    try {
        $response = Invoke-RestMethod -Method Get -Uri $uri -Headers @{
            Authorization = "Bearer $SharePointToken"
            Accept = 'application/json;odata=nometadata'
        }
    } catch {
        Write-Log "Adgangstjek fejlede for ${SiteUrl}: $($_.Exception.Message)" 'WARN'
        return $true
    }

    $permissions = if ($response.GetUserEffectivePermissions) { $response.GetUserEffectivePermissions } else { $response }
    $low = [uint64]$permissions.Low
    $hasViewListItems = ($low -band ([uint64]1 -shl 0)) -ne 0
    $hasOpen = ($low -band ([uint64]1 -shl 17)) -ne 0
    $hasViewPages = ($low -band ([uint64]1 -shl 18)) -ne 0

    return ($hasViewListItems -or $hasOpen -or $hasViewPages)
}

function Invoke-ServerSideAccessCheck {
    param(
        [Parameter(Mandatory)][string]$UserPrincipalName,
        [Parameter(Mandatory)][object[]]$Entries
    )

    $payloadEntries = @($Entries | ForEach-Object {
        [pscustomobject]@{
            source = $_.Source
            displayName = $_.DisplayName
            localPath = $_.LocalPath
            siteUrl = $_.SiteUrl
        }
    })

    $payload = @{
        sharePointHostName = $SharePointHostName
        userPrincipalName = $UserPrincipalName
        entries = $payloadEntries
    } | ConvertTo-Json -Depth 8

    $headers = @{}
    if (-not [string]::IsNullOrWhiteSpace($AccessCheckFunctionKey)) {
        $headers['x-functions-key'] = $AccessCheckFunctionKey
    }

    return Invoke-RestMethod -Method Post -Uri $AccessCheckFunctionUrl -Headers $headers -Body $payload -ContentType 'application/json'
}

function Remove-SyncRegistryEntry {
    param([Parameter(Mandatory)]$Entry)

    if ($DryRun) {
        Write-Log "DRY RUN: Ville fjerne registry for $($Entry.LocalPath)"
        return
    }

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

function Remove-LocalSyncFolder {
    param([Parameter(Mandatory)]$Entry)

    if ($Entry.Source -eq 'FileSystemOnly' -and -not $AllowFileSystemOnlyDeletion) {
        Write-Log "Springer filesystem-only mappe over, fordi der ikke er en matchende OneDrive registry-post: $($Entry.LocalPath)" 'WARN'
        return
    }

    if (-not $Entry.LocalPathExists) {
        Write-Log "Lokal mappe findes ikke længere: $($Entry.LocalPath)"
        return
    }

    if ($DryRun) {
        Write-Log "DRY RUN: Ville slette lokal mappe $($Entry.LocalPath)"
        return
    }

    try {
        Get-ChildItem -LiteralPath $Entry.LocalPath -Force -Recurse -ErrorAction SilentlyContinue |
            ForEach-Object { $_.Attributes = 'Normal' }
        (Get-Item -LiteralPath $Entry.LocalPath -Force).Attributes = 'Normal'
    } catch {
        Write-Log "Kunne ikke nulstille filattributter for $($Entry.LocalPath): $($_.Exception.Message)" 'WARN'
    }

    try {
        Remove-Item -LiteralPath $Entry.LocalPath -Recurse -Force -ErrorAction Stop
    } catch {
        Write-Log "Remove-Item fejlede for $($Entry.LocalPath): $($_.Exception.Message). Prøver cmd.exe rmdir fallback." 'WARN'
        $cmdOutput = & cmd.exe /c rd /s /q "\\?\$($Entry.LocalPath)" 2>&1
        if ($LASTEXITCODE -ne 0) {
            Write-Log "cmd.exe rmdir fallback returnerede exitkode $LASTEXITCODE for $($Entry.LocalPath). Output: $cmdOutput" 'WARN'
        }
    }

    if (Test-Path -LiteralPath $Entry.LocalPath) {
        $quarantinePath = Join-Path -Path (Split-Path -Parent $Entry.LocalPath) -ChildPath ("_OneDriveCleanup_DeleteMe_{0:yyyyMMddHHmmss}" -f (Get-Date))
        try {
            Rename-Item -LiteralPath $Entry.LocalPath -NewName (Split-Path -Leaf $quarantinePath) -Force -ErrorAction Stop
            Write-Log "Mappen kunne ikke slettes direkte og blev omdøbt til: $quarantinePath" 'WARN'

            $cmdOutput = & cmd.exe /c rd /s /q "\\?\$quarantinePath" 2>&1
            if ($LASTEXITCODE -ne 0 -and (Test-Path -LiteralPath $quarantinePath)) {
                throw "cmd.exe rmdir kunne heller ikke slette karantænemappen. Output: $cmdOutput"
            }
        } catch {
            throw "Kunne ikke slette eller omdøbe $($Entry.LocalPath). $($_.Exception.Message)"
        }
    }

    if (Test-Path -LiteralPath $Entry.LocalPath) {
        throw "Mappen findes stadig efter sletningsforsøg: $($Entry.LocalPath)"
    }

    Write-Log "Slettede lokal synkroniseret mappe: $($Entry.LocalPath)"
}

try {
    Install-PeriodicCheckTask
    Write-Log "Starter SharePoint/OneDrive sync-oprydning version $ScriptVersion."
    Write-Log "Kører som bruger: $([Security.Principal.WindowsIdentity]::GetCurrent().Name) / USERPROFILE=$env:USERPROFILE / LOCALAPPDATA=$env:LOCALAPPDATA"
    if ($DryRun) {
        Write-Log 'DryRun er slået til. Ingen registry-poster eller mapper bliver slettet.'
    }

    $useServerSideAccessCheck = -not [string]::IsNullOrWhiteSpace($AccessCheckFunctionUrl)
    $missingAppOnlyConfig = [string]::IsNullOrWhiteSpace($TenantId) -or
                            [string]::IsNullOrWhiteSpace($ClientId) -or
                            [string]::IsNullOrWhiteSpace($CertificateThumbprint) -or
                            [string]::IsNullOrWhiteSpace($SharePointHostName)

    if (-not $useServerSideAccessCheck -and $missingAppOnlyConfig) {
        Write-Log 'App/cert-konfiguration mangler. Kører kun Intune inventory uden SharePoint-adgangstjek for at undgå hard-fail i live-test.' 'WARN'
        $inventoryEntries = @(Get-SyncedSharePointEntry)
        foreach ($entry in $inventoryEntries) {
            Write-Log "Fundet sync-post: $($entry.Source) / $($entry.DisplayName) / LocalPathExists=$($entry.LocalPathExists) / SiteUrl=$($entry.SiteUrl) / $($entry.LocalPath)"
        }
        Write-Log 'Inventory afsluttet uden sletning, fordi der ikke er konfiguration til adgangstjek.'
        exit 0
    }

    $userPrincipalName = Get-CurrentUserPrincipalName
    $entries = @(Get-SyncedSharePointEntry)

    if ($entries.Count -eq 0) {
        Write-Log 'Ingen SharePoint/OneDrive for Business sync-poster fundet.'
        exit 0
    }

    $entriesToRemove = New-Object System.Collections.Generic.List[object]
    if ($useServerSideAccessCheck) {
        Write-Log "Bruger server-side adgangstjek via Azure Function: $AccessCheckFunctionUrl"
        $accessCheckResponse = Invoke-ServerSideAccessCheck -UserPrincipalName $userPrincipalName -Entries $entries
        foreach ($result in @($accessCheckResponse.results)) {
            Write-Log "Adgangstjek-resultat: $($result.accessState) / $($result.displayName) / SiteUrl=$($result.siteUrl) / $($result.reason)"
            if ($result.accessState -ne 'MissingAccess') {
                continue
            }

            $matchingEntry = $entries | Where-Object { $_.LocalPath -eq $result.localPath } | Select-Object -First 1
            if ($matchingEntry) {
                if ($matchingEntry.Source -eq 'FileSystemOnly') {
                    Write-Log "Brugeren mangler adgang, men posten er filesystem-only og bliver ikke fjernet: $($matchingEntry.LocalPath)" 'WARN'
                    continue
                }

                $matchingEntry.SiteUrl = $result.siteUrl
                Write-Log "Brugeren mangler adgang. Markerer til fjernelse: $($matchingEntry.SiteUrl) / $($matchingEntry.LocalPath)" 'WARN'
                $entriesToRemove.Add($matchingEntry)
            }
        }
    } else {
        $certificate = Get-CleanupCertificate
        $graphToken = Get-AppOnlyAccessToken -ResourceScope 'https://graph.microsoft.com/.default' -Certificate $certificate
        $sharePointToken = Get-AppOnlyAccessToken -ResourceScope "https://$SharePointHostName/.default" -Certificate $certificate
        foreach ($entry in $entries) {
            Write-Log "Vurderer sync-post: $($entry.Source) / $($entry.DisplayName) / $($entry.LocalPath)"
            if (-not $entry.SiteUrl) {
                $entry.SiteUrl = Search-GraphSiteUrl -DisplayName $entry.DisplayName -GraphToken $graphToken
            }

            if (-not $entry.SiteUrl) {
                Write-Log "Springer over, fordi site-URL ikke kunne bestemmes sikkert: $($entry.DisplayName) / $($entry.LocalPath)" 'WARN'
                continue
            }

            $hasAccess = Test-SharePointUserAccess -SiteUrl $entry.SiteUrl -UserPrincipalName $userPrincipalName -DisplayName $entry.DisplayName -SharePointToken $sharePointToken
            if ($hasAccess) {
                Write-Log "Brugeren har stadig adgang: $($entry.SiteUrl)"
                continue
            }

            if ($entry.Source -eq 'FileSystemOnly') {
                Write-Log "Brugeren mangler adgang, men posten er filesystem-only og bliver ikke fjernet: $($entry.LocalPath)" 'WARN'
                continue
            }

            Write-Log "Brugeren mangler adgang. Markerer til fjernelse: $($entry.SiteUrl) / $($entry.LocalPath)" 'WARN'
            $entriesToRemove.Add($entry)
        }
    }

    if ($entriesToRemove.Count -eq 0) {
        Write-Log 'Ingen utilgængelige sync-poster skulle fjernes.'
        exit 0
    }

    if (-not $DryRun) {
        Export-RegistryBackup -RegistryPaths @(
            'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\SyncRootManager',
            'HKCU:\Software\Microsoft\OneDrive\Accounts'
        )
        Get-Process -Name OneDrive -ErrorAction SilentlyContinue | Stop-Process -Force
    }

    foreach ($entry in $entriesToRemove) {
        Remove-SyncRegistryEntry -Entry $entry
    }

    if ($RemoveLocalFolders) {
        foreach ($entry in ($entriesToRemove | Sort-Object LocalPath -Unique)) {
            Remove-LocalSyncFolder -Entry $entry
        }
    }

    if ($RestartOneDrive -and -not $DryRun) {
        $oneDrivePath = Get-OneDriveExecutablePath
        if ($oneDrivePath) {
            Restart-OneDriveAsStandardUser -OneDrivePath $oneDrivePath
        } else {
            Write-Log 'OneDrive.exe blev ikke fundet. Brugeren skal starte OneDrive manuelt.' 'WARN'
        }
    }

    Write-Log 'SharePoint/OneDrive sync-oprydning afsluttet.'
    exit 0
} catch {
    Write-Log $_.Exception.Message 'ERROR'
    Write-Log (($_ | Out-String).Trim()) 'ERROR'
    exit 1
}
