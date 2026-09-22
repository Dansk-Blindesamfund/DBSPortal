$scriptPath = Join-Path -Path $PSScriptRoot -ChildPath '..\src\Remove-StaleOneDriveSites.ps1'

& $scriptPath -NameOrPathPattern '*Eksempel Site*' -RemoveRegistryEntries -WhatIf
