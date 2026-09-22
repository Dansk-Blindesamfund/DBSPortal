# Azure Function: SharePoint Access Check

PowerShell Azure Function til server-side adgangstjek for OneDrive/SharePoint cleanup-scriptet.

Formålet er at holde SharePoint app-permissions væk fra klienterne. Klienten sender kun inventory over lokale sync-poster; Azure Function bruger managed identity til at tjekke, om brugeren stadig har adgang til de tilhørende SharePoint-sites.

## Hvorfor Azure Function

Azure Function passer bedst her, fordi Intune-scriptet har lokal viden om maskinens sync-mapper, mens adgangstjekket bør ske server-side. Alternativet Azure Automation er godt til planlagte jobs, men mindre oplagt når klienten skal spørge on-demand for lige præcis de mapper, der findes lokalt.

## Deployment-overblik

1. Opret en Function App med PowerShell runtime.
2. Slå system-assigned managed identity til.
3. Giv managed identity application permissions til SharePoint/Graph.
4. Deploy indholdet af denne mappe til Function App'en.
5. Sæt Intune-scriptets `$AccessCheckFunctionUrl` til function endpointet.
6. Gem function key sikkert og indsæt den i `$AccessCheckFunctionKey`, eller brug en bedre auth-front som Easy Auth/APIM i produktion.

## Nødvendige permissions

Managed identity skal kunne:

- Søge/læse sites via Microsoft Graph, fx `Sites.Read.All`.
- Kalde SharePoint REST `GetUserEffectivePermissions`, fx SharePoint `Sites.FullControl.All` eller en smallere model hvis tenant'en understøtter det i jeres setup.

## Request

En typisk request kan bygges med helper-scriptet i `tools/Export-OneDriveAccessCheckPayload.ps1`:

```powershell
.\tools\Export-OneDriveAccessCheckPayload.ps1 `
  -UserPrincipalName "teb@blind.dk" `
  -SharePointHostName "blindesamfund.sharepoint.com" `
  -OutputPath ".\out\onedrive-access-check-payload.json"
```

Eksempel på JSON payload:

```json
{
  "sharePointHostName": "blindesamfund.sharepoint.com",
  "userPrincipalName": "teb@blind.dk",
  "entries": [
    {
      "source": "OneDriveAccountCache",
      "displayName": "Fælles - Dokumenter",
      "localPath": "C:\\Users\\teb_dbs\\Dansk Blindesamfund\\Fælles - Dokumenter",
      "siteUrl": ""
    }
  ]
}
```

## Response

```json
{
  "userPrincipalName": "teb@blind.dk",
  "results": [
    {
      "source": "OneDriveAccountCache",
      "displayName": "Fælles - Dokumenter",
      "localPath": "C:\\Users\\teb_dbs\\Dansk Blindesamfund\\Fælles - Dokumenter",
      "siteUrl": "https://blindesamfund.sharepoint.com/sites/Faelles",
      "accessState": "HasAccess",
      "reason": "User has view/open permissions"
    }
  ]
}
```

`accessState` kan være:

- `HasAccess`
- `MissingAccess`
- `Unknown`

Klienten må kun fjerne sync-poster med `MissingAccess`.
