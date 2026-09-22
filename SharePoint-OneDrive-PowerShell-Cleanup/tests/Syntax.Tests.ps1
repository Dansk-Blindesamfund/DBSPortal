$projectRoot = Split-Path -Parent $PSScriptRoot
$scriptPaths = Get-ChildItem -Path $projectRoot -Filter '*.ps1' -Recurse | Where-Object { $_.FullName -ne $PSCommandPath }

foreach ($scriptPath in $scriptPaths) {
    $tokens = $null
    $parseErrors = $null
    [System.Management.Automation.Language.Parser]::ParseFile($scriptPath.FullName, [ref]$tokens, [ref]$parseErrors) | Out-Null

    if ($parseErrors.Count -gt 0) {
        $parseErrors | Format-List | Out-String | Write-Error
        exit 1
    }

    Write-Host "Syntax OK: $($scriptPath.FullName)"
}
