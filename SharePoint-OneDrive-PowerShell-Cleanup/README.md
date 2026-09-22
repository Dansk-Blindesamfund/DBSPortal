# SharePoint/OneDrive PowerShell Cleanup

PowerShell-projekt til at finde og fjerne lokale OneDrive/SharePoint sync-registreringer for sites, som brugeren ikke længere har adgang til.

Scriptet er lavet til Windows-klienter med OneDrive for Business. Det sletter ikke lokale filer som standard og kører bedst først som preview.

Projektet indeholder også et Intune-klart platform script: `intune/Remove-InaccessibleSharePointSync.Intune.ps1`.

## Hvad scriptet gør

- Finder OneDrive/SharePoint sync-roots i brugerens registry.
- Finder OneDrive for Business mount point cache-poster.
- Kan filtrere på site-/mappenavn med wildcard-mønstre.
- Kan kun vise lokale sync-roots hvor mount point-stien ikke længere findes.
- Kan eksportere registry-backup før ændringer.
- Kan fjerne valgte registry-poster og eventuelt lokale mapper.
- Kan genstarte OneDrive efter oprydning.

## Begrænsning

Microsoft har ikke en stabil, offentlig PowerShell-cmdlet til at stoppe synkronisering af enkelte SharePoint-biblioteker i OneDrive-klienten. Derfor rydder scriptet lokale Explorer/OneDrive sync-registreringer. Brug det især når OneDrive allerede fejler på et site, adgangen er fjernet, eller en gammel sync-root hænger tilbage.

## Brug

Åbn PowerShell i projektmappen:

```powershell
Set-Location .\SharePoint-OneDrive-PowerShell-Cleanup
```

### 1) Generér payload til Azure Function access-check

Dette giver en JSON-payload med de lokale sync-poster, som klienten vil sende til server-side adgangstjek:

```powershell
.\tools\Export-OneDriveAccessCheckPayload.ps1 `
  -UserPrincipalName "teb@blind.dk" `
  -SharePointHostName "blindesamfund.sharepoint.com" `
  -OutputPath ".\out\onedrive-access-check-payload.json"
```

Outputtet kan sendes til Azure Function endpointet i `server/AzureFunction/CheckSharePointAccess`.

### 2) Vis alle fundne OneDrive for Business sync-poster

```powershell
.\src\Remove-StaleOneDriveSites.ps1
```

Vis kun poster hvor den lokale mappe mangler:

```powershell
.\src\Remove-StaleOneDriveSites.ps1 -OnlyUnavailableLocalPath
```

Preview fjernelse af et bestemt site eller en bestemt mappe:

```powershell
.\src\Remove-StaleOneDriveSites.ps1 -NameOrPathPattern "*Eksempel Site*" -RemoveRegistryEntries -WhatIf
```

Udfør registry-oprydning med backup:

```powershell
.\src\Remove-StaleOneDriveSites.ps1 -NameOrPathPattern "*Eksempel Site*" -RemoveRegistryEntries -RestartOneDrive
```

Fjern også den lokale mappe. Brug kun dette, når du er sikker på at lokale filer ikke skal bevares:

```powershell
.\src\Remove-StaleOneDriveSites.ps1 -NameOrPathPattern "*Eksempel Site*" -RemoveRegistryEntries -RemoveLocalFolders -RestartOneDrive
```

## Anbefalet arbejdsgang

1. Kør scriptet uden ændrings-switches og find de relevante poster.
2. Kør igen med `-NameOrPathPattern` og `-WhatIf`.
3. Kør uden `-WhatIf`, når outputtet kun matcher de sites, der skal fjernes.
4. Genstart OneDrive, hvis du ikke bruger `-RestartOneDrive`.

Backup-filer gemmes i `backup/` som `.reg`-filer.

## Intune platform script

Status: Test-scripts er ryddet op i Intune efter den kontrollerede slet-test. Upload ikke de gamle REAL TEST-scripts igen uden at ændre dem til DryRun først.

Brug `intune/Remove-InaccessibleSharePointSync.Intune.ps1`, når scriptet skal udrulles via Intune til en sikkerhedsgruppe.

Vigtige Intune-indstillinger:

- Run this script using the logged on credentials: `Yes`
- Enforce script signature check: efter jeres signing-politik
- Run script in 64-bit PowerShell Host: `Yes`
- Assignments: vælg den sikkerhedsgruppe, der skal have oprydningen

Scriptet skal køre i bruger-kontekst, fordi OneDrive sync-roots ligger under `HKCU` og brugerens lokale profil.

### Krævet app-registrering

For at kunne tjekke om brugeren stadig har adgang til et SharePoint-site uden interaktiv login skal Intune-scriptet bruge en Entra ID app registration med certifikatbaseret app-only auth.

Vigtigt: udrul ikke app-certifikatets private key bredt til klienterne. Med de nuværende app permissions kan den private key bruges til at hente app-only tokens mod SharePoint. En lokal administrator eller kompromitteret klient vil derfor kunne misbruge certifikatet. Intune kan teknisk set udrulle PFX-certifikater til en gruppe, men det bør ikke bruges til dette app-only certifikat i produktion.

Udfyld øverst i scriptet:

```powershell
$TenantId = '7269da2b-d73e-4791-ba91-3675fa4b83f0'
$ClientId = 'b90b44fd-4ef4-4f70-82ab-e548aa190281'
$CertificateThumbprint = 'CERT_THUMBPRINT_HER'
$SharePointHostName = 'blindesamfund.sharepoint.com'
$DryRun = $true
```

Hvis adgangstjekket køres via Azure Function, skal klient-scriptet i stedet bruge function endpointet:

```powershell
$AccessCheckFunctionUrl = 'https://<function-app>.azurewebsites.net/api/sharepoint/access-check'
$AccessCheckFunctionKey = $env:DBS_ACCESS_CHECK_FUNCTION_KEY
```

Intune-scriptet kan hentes fra repositoryets `main`-branch, så det ikke er bundet til en lokal OneDrive-mappe. Function key skal fortsat leveres separat som miljøvariablen `DBS_ACCESS_CHECK_FUNCTION_KEY` og må ikke lægges i GitHub.

Med Azure Function-modellen skal klienterne ikke have app-certifikatets private key. Function key er stadig en hemmelighed og bør kun give adgang til dette afgrænsede adgangstjek, ikke direkte SharePoint/Graph-tokens.

Appen skal have admin-consent til relevante Graph/SharePoint application permissions. Brug mindst Graph site search og SharePoint site permission check. I praksis vil mange tenants bruge `Sites.Read.All` til Graph og SharePoint app-only adgang til at læse effektive permissions.

Certifikatet med privat nøgle skal findes på klienten i `Cert:\LocalMachine\My`, eller også skal `$CertificateStorePath` ændres.

Aktuel testopsætning i Intune:

- Testgruppe: `OneDrive sletning`
- All-sites DryRun script: `OneDrive SharePoint cleanup - All sites - DryRun`
- All-sites DryRun script-id: `9be794ca-c265-4cdb-bfb5-27b970c426e3`
- App registration: `OneDrive SharePoint Cleanup Access Check`
- Client ID: `b90b44fd-4ef4-4f70-82ab-e548aa190281`

Før `$DryRun` sættes til `$false`, skal certifikatet med privat nøgle være udrullet til klienterne, og `$CertificateThumbprint` skal udfyldes med klientcertifikatets thumbprint.

Produktionsscriptet sletter som standard ikke `FileSystemOnly`-mapper, altså mapper der kun findes i filsystemet, men ikke som OneDrive registry-post. Det er bevidst forsigtigt, fordi sådan en mappe kan være en aktiv OneDrive-resttilstand. Sæt kun `$AllowFileSystemOnlyDeletion = $true`, hvis den konkrete sync-type er testet.

OneDrive genstartes ikke automatisk efter oprydning (`$RestartOneDrive = $false`). Det forhindrer klienten i at genoprette en sync, hvis OneDrive stadig har en aktiv sync-registrering. Genstart skal først aktiveres igen, når sync-registreringen er valideret fjernet.

### Anbefalet sikker produktionsmodel

Den sikre produktionsmodel er at holde app-certifikatets private key væk fra klienterne:

1. Klient-scriptet kører i Intune og laver inventory over synkede SharePoint-mapper.
2. En server-side komponent, fx Azure Automation Runbook eller Azure Function med managed identity/certifikat, laver SharePoint-adgangstjekket.
3. Klient-scriptet får kun en liste over de konkrete lokale sync-poster, der må ryddes.
4. Klienten fjerner kun registry-baserede OneDrive sync-poster og sletter ikke `FileSystemOnly`-mapper som standard.

Det betyder, at en kompromitteret klient ikke får et genbrugeligt SharePoint app-only certifikat.

Der er tilføjet et PowerShell Azure Function-projekt i `server/AzureFunction`, som implementerer server-side adgangstjekket med managed identity.

## Recovery efter test

Hvis OneDrive bliver ustabil efter en slet-test, kan recovery-scriptet bruges på klienten:

```powershell
.\tools\Repair-OneDriveAfterCleanup.ps1 -RestoreLatestRegistryBackup -ResetOneDrive
```

Hvis OneDrive bare skal genstartes uden registry-restore:

```powershell
.\tools\Repair-OneDriveAfterCleanup.ps1
```

### Anbefalet første udrulning

Start med `$DryRun = $true`. Scriptet logger til:

```text
%LOCALAPPDATA%\OneDriveSharePointCleanup\AccessCheckCleanup.log
```

Når loggen viser de rigtige sites og mapper, ændres `$DryRun = $false`.

Hvis registry på klienterne ikke indeholder nok information til sikkert at finde site-URL, springer scriptet posten over. Til første produktion anbefales eksplicit mapping:

```powershell
$SiteMappings = @(
	@{ LocalPathPattern = '*Fælles - Dokumenter*'; SiteUrl = 'https://tenant.sharepoint.com/sites/Faelles' }
)
```

Det er bevidst konservativt: scriptet skal ikke gætte sig til et site og slette en lokal mappe på et forkert match.
