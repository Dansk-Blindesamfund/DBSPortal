# DBS Portal Drift

## Adresser

- Offentlig portal: https://dbs-itsupport-portal-4377.grayforest-78bf8bbb.westeurope.azurecontainerapps.io/
- Lokal portal: http://localhost:8080/
- Lokal testredirect til SPA er ogsaa registreret: http://localhost:5500/index.html

## Offentlig auth

Den offentlige portal bruger Azure Container App Authentication foran selve appen.

- Provider: Microsoft / Entra ID
- App registration: DBS Portal
- Client ID: 2ad3b090-e894-456c-8831-5761163ae173
- Container App: dbs-itsupport-portal-4377
- Resource group: DBS

Vigtige redirect-URL'er i Entra:

- https://dbs-itsupport-portal-4377.grayforest-78bf8bbb.westeurope.azurecontainerapps.io/
- https://dbs-itsupport-portal-4377.grayforest-78bf8bbb.westeurope.azurecontainerapps.io/.auth/login/aad/callback
- http://localhost:8080/
- http://localhost:5500/index.html

## Secrets og konfiguration

Trackede filer indeholder ikke laengere de foelsomme vaerdier.

Lokal maskine:

- Brug ITSupportSystem/appsettings.local.json til lokale secrets.
- Filen er ignoreret i Git.

Azure Container App:

- Microsoft auth secret: microsoft-provider-authentication-secret
- Backend Graph secret boer ligge som Container App secret og bindes via env var.
- Backend Blob connection string boer ligge som Container App secret og bindes via env var.

Anbefalede env vars til containeren:

- GRAPH_TENANT_ID
- GRAPH_CLIENT_ID
- GRAPH_CLIENT_SECRET
- AZURE_STORAGE_CONNECTION_STRING
- TICKET_BLOB_ENABLED=true
- TICKET_BLOB_ACCOUNT_URL
- TICKET_BLOB_CONTAINER=itsupport
- TICKET_BLOB_NAME=tickets.json
- SUPPORT_MAILBOX_ADDRESS
- SUPPORT_MAILBOX_DISPLAY_NAME
- FEATURE_USE_TEST_MAILBOX
- FEATURE_TEST_MAILBOX_ADDRESS

Autorisation (roller og adgang):

- AUTHZ_DEV_BYPASS=false i public/prod (true kan bruges lokalt).
- Standardadgang er "kun egne sager" for alle brugere.
- Sikkerhedsgruppen `sg_itsupportsystem` er master og faar adgang til alle sager.
- AUTHZ_MASTER_GROUP_NAMES=sg_itsupportsystem
- AUTHZ_MASTER_GROUP_IDS=<entra-group-object-id> (anbefalet i produktion)
- AUTHZ_ADMIN_EMAILS=mail1@blind.dk,mail2@blind.dk
- AUTHZ_AGENT_EMAILS=agent1@blind.dk,agent2@blind.dk
- AUTHZ_VIEWER_EMAILS=viewer1@blind.dk
- AUTHZ_REQUESTER_EMAILS=req1@blind.dk

Valgfrit kan hele rollemodellen styres via JSON i en enkelt env var:

- AUTHZ_RULES_JSON

Eksempel:

```json
{
	"defaultRole": "requester",
	"masterGroupNames": ["sg_itsupportsystem"],
	"masterGroupIds": ["<entra-group-object-id>"],
	"rolePermissions": {
		"admin": ["*"],
		"agent": ["portal.mail", "portal.word", "portal.excel", "portal.teams", "portal.sharepoint", "portal.onedrive", "portal.forms", "portal.cases.view", "portal.cases.reply", "portal.cases.close", "portal.cases.read.all"],
		"viewer": ["portal.forms", "portal.cases.view", "portal.cases.read.all"],
		"requester": ["portal.forms", "portal.cases.view", "portal.cases.read.own"]
	},
	"roleMappings": {
		"admin": ["admin@blind.dk"],
		"agent": ["agent@blind.dk"],
		"viewer": ["viewer@blind.dk"],
		"requester": ["requester@blind.dk"]
	}
}
```

## Deploy

Byg og deploy fra projektroden:

```powershell
.\scripts\deploy-azure-portal.ps1 -ResourceGroup DBS -Location westeurope -AcrName <acr-navn> -ContainerAppName <app-navn> -ContainerEnvName <env-navn> -ImageTag v1
```

Hvis du kun vil opdatere den eksisterende offentlige app efter kodeaendringer:

```powershell
az acr build --registry dbsitsupacr4377 --image itsupport-portal:v1 --file deploy/portal.Dockerfile .
az containerapp update -g DBS -n dbs-itsupport-portal-4377 --image dbsitsupacr4377.azurecr.io/itsupport-portal:v1

Aktuel offentlig release i denne session er bygget/deployet som v2.
```

## Restart

Genstart offentlig container app:

```powershell
az containerapp revision restart -g DBS -n dbs-itsupport-portal-4377 --revision $(az containerapp revision list -g DBS -n dbs-itsupport-portal-4377 --query "[0].name" -o tsv)
```

Genstart lokal monitor:

```powershell
Set-Location .\ITSupportSystem
dotnet .\bin\Debug\net8.0\ITSupportSystem.dll run
```

## Health og verifikation

Tjek offentlig health endpoint:

```powershell
Invoke-WebRequest -Uri "https://dbs-itsupport-portal-4377.grayforest-78bf8bbb.westeurope.azurecontainerapps.io/healthz"
```

Tjek auth-konfiguration:

```powershell
az containerapp auth show -g DBS -n dbs-itsupport-portal-4377
```

## Token protection

Tenantens tokenbeskyttelse kan blokere browserbaserede Graph-kald fra SPA-funktioner.

Derfor er den offentlige portal nu lavet saadan:

- Portal-login virker via server-side auth i Container App.
- Rene browser-Graph-funktioner kan deaktiveres eller vise forklarende fejl.
- Sager, formularer og lokale editorer kan stadig bruges uden de browser-Graph-kald.

Hvis Graph-funktioner senere skal virke offentligt trods token protection, boer de flyttes til server-side API-kald.
