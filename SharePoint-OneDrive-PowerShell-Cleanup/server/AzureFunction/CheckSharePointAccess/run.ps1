using namespace System.Net

param($Request, $TriggerMetadata)

function Get-AccessToken {
    param([Parameter(Mandatory)][string]$ResourceUrl)

    if ([string]::IsNullOrWhiteSpace($env:IDENTITY_ENDPOINT) -or [string]::IsNullOrWhiteSpace($env:IDENTITY_HEADER)) {
        throw 'Managed identity endpoint is not available. Enable system-assigned managed identity on the Function App.'
    }

    $tokenUri = '{0}?api-version=2019-08-01&resource={1}' -f $env:IDENTITY_ENDPOINT, [uri]::EscapeDataString($ResourceUrl)
    $tokenResponse = Invoke-RestMethod -Method Get -Uri $tokenUri -Headers @{ 'X-IDENTITY-HEADER' = $env:IDENTITY_HEADER }
    return $tokenResponse.access_token
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

function Search-GraphSiteUrl {
    param(
        [Parameter(Mandatory)][string]$DisplayName,
        [Parameter(Mandatory)][string]$SharePointHostName,
        [Parameter(Mandatory)][string]$GraphToken
    )

    foreach ($query in Get-SiteSearchQueries -DisplayName $DisplayName) {
        $uri = 'https://graph.microsoft.com/v1.0/sites?search={0}' -f [uri]::EscapeDataString($query)
        $response = Invoke-RestMethod -Method Get -Uri $uri -Headers @{ Authorization = "Bearer $GraphToken" }
        $matches = @($response.value | Where-Object { $_.webUrl -like "https://$SharePointHostName/*" })

        $siteName = ($DisplayName -replace '\s+-\s+.*$', '').Trim()
        $exactMatches = @($matches | Where-Object { $_.displayName -eq $siteName })
        if ($exactMatches.Count -eq 1) {
            return [string]$exactMatches[0].webUrl.TrimEnd('/')
        }

        if ($matches.Count -eq 1) {
            return [string]$matches[0].webUrl.TrimEnd('/')
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
    } else {
        $uri = "{0}/_api/web/GetUserEffectivePermissions(@u)?@u={1}" -f $SiteUrl.TrimEnd('/'), $encodedClaim
    }

    $response = Invoke-RestMethod -Method Get -Uri $uri -Headers @{
        Authorization = "Bearer $SharePointToken"
        Accept = 'application/json;odata=nometadata'
    }

    $permissions = if ($response.GetUserEffectivePermissions) { $response.GetUserEffectivePermissions } else { $response }
    $low = [uint64]$permissions.Low
    $hasViewListItems = ($low -band ([uint64]1 -shl 0)) -ne 0
    $hasOpen = ($low -band ([uint64]1 -shl 17)) -ne 0
    $hasViewPages = ($low -band ([uint64]1 -shl 18)) -ne 0

    return ($hasViewListItems -or $hasOpen -or $hasViewPages)
}

try {
    $body = $Request.Body
    if ($body -is [string]) {
        $body = $body | ConvertFrom-Json
    }

    if ([string]::IsNullOrWhiteSpace($body.userPrincipalName)) {
        throw 'userPrincipalName is required.'
    }

    if ([string]::IsNullOrWhiteSpace($body.sharePointHostName)) {
        throw 'sharePointHostName is required.'
    }

    if (-not $body.entries) {
        throw 'entries is required.'
    }

    $sharePointHostName = [string]$body.sharePointHostName
    $userPrincipalName = [string]$body.userPrincipalName
    $graphToken = Get-AccessToken -ResourceUrl 'https://graph.microsoft.com'
    $sharePointToken = Get-AccessToken -ResourceUrl "https://$sharePointHostName"
    $results = New-Object System.Collections.Generic.List[object]

    foreach ($entry in @($body.entries)) {
        $siteUrl = if ($entry.siteUrl) { [string]$entry.siteUrl } else { $null }
        if ([string]::IsNullOrWhiteSpace($siteUrl)) {
            $siteUrl = Search-GraphSiteUrl -DisplayName ([string]$entry.displayName) -SharePointHostName $sharePointHostName -GraphToken $graphToken
        }

        if ([string]::IsNullOrWhiteSpace($siteUrl)) {
            $results.Add([pscustomobject]@{
                source = $entry.source
                displayName = $entry.displayName
                localPath = $entry.localPath
                siteUrl = $null
                accessState = 'Unknown'
                reason = 'Site URL could not be resolved safely'
            })
            continue
        }

        try {
            $hasAccess = Test-SharePointUserAccess -SiteUrl $siteUrl -UserPrincipalName $userPrincipalName -DisplayName ([string]$entry.displayName) -SharePointToken $sharePointToken
            $results.Add([pscustomobject]@{
                source = $entry.source
                displayName = $entry.displayName
                localPath = $entry.localPath
                siteUrl = $siteUrl
                accessState = if ($hasAccess) { 'HasAccess' } else { 'MissingAccess' }
                reason = if ($hasAccess) { 'User has view/open permissions' } else { 'User has no view/open permissions' }
            })
        } catch {
            $results.Add([pscustomobject]@{
                source = $entry.source
                displayName = $entry.displayName
                localPath = $entry.localPath
                siteUrl = $siteUrl
                accessState = 'Unknown'
                reason = $_.Exception.Message
            })
        }
    }

    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = [HttpStatusCode]::OK
        Headers = @{ 'Content-Type' = 'application/json' }
        Body = (@{
            userPrincipalName = $userPrincipalName
            results = $results
        } | ConvertTo-Json -Depth 8)
    })
} catch {
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = [HttpStatusCode]::BadRequest
        Headers = @{ 'Content-Type' = 'application/json' }
        Body = (@{ error = $_.Exception.Message } | ConvertTo-Json)
    })
}
