[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [switch]$RestoreLatestRegistryBackup,
    [switch]$ResetOneDrive,
    [string]$CleanupLogRoot = (Join-Path -Path $env:LOCALAPPDATA -ChildPath 'OneDriveSharePointCleanup')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Write-Step {
    param([Parameter(Mandatory)][string]$Message)
    Write-Host ('[{0:T}] {1}' -f (Get-Date), $Message)
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

function Import-LatestCleanupRegistryBackup {
    if (-not (Test-Path -LiteralPath $CleanupLogRoot)) {
        Write-Warning "Backup-mappen findes ikke: $CleanupLogRoot"
        return
    }

    $backupFiles = Get-ChildItem -LiteralPath $CleanupLogRoot -Filter '*.reg' -File -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending

    if (-not $backupFiles) {
        Write-Warning "Der blev ikke fundet .reg-backups i $CleanupLogRoot"
        return
    }

    foreach ($backupFile in $backupFiles) {
        if ($PSCmdlet.ShouldProcess($backupFile.FullName, 'Import registry backup')) {
            Write-Step "Importerer registry-backup: $($backupFile.FullName)"
            $output = & reg.exe import $backupFile.FullName 2>&1
            if ($LASTEXITCODE -ne 0) {
                Write-Warning "reg.exe import fejlede for $($backupFile.FullName). Output: $output"
            }
        }
    }
}

$oneDrivePath = Get-OneDriveExecutablePath
if (-not $oneDrivePath) {
    throw 'OneDrive.exe blev ikke fundet på maskinen.'
}

Write-Step 'Stopper OneDrive.'
Get-Process -Name OneDrive -ErrorAction SilentlyContinue | Stop-Process -Force

if ($RestoreLatestRegistryBackup) {
    Import-LatestCleanupRegistryBackup
}

if ($ResetOneDrive) {
    Write-Step 'Kører OneDrive /reset.'
    Start-Process -FilePath $oneDrivePath -ArgumentList '/reset'
    Write-Step 'Vent 1-2 minutter. Start derefter OneDrive igen, hvis den ikke åbner automatisk.'
} else {
    Write-Step 'Starter OneDrive.'
    Start-Process -FilePath $oneDrivePath
}

Write-Step 'Færdig. Hvis OneDrive beder om login, skal brugeren tilslutte sin arbejdskonto igen.'
