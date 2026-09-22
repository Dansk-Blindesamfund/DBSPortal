# IT Support System

Dette program overvager en support-postkasse og opretter sager automatisk.

## Funktioner

- Lytter efter nye mails i postkassen.
- Opretter sag og sender kvittering ved modtagelse.
- Genkender eksisterende sager via sagsnummer i emnefeltet.
- Sender mails til brugeren ved svar eller lukning af sag.
- Understotter Microsoft Graph som primaer metode.
- Har feature toggles til test (aktiver/deaktiver funktioner).
- Kan nemt skifte til testpostkasse, fx `teb@blind.dk`.
- Udsendte mails indeholder knapperne `Svar pa sag` og `Luk sagen`.
- Naar en sag lukkes, flyttes relaterede mails til papirkurven i den aktive supportmailbox.
- Supportmailboxen faar en samlet `[SUPPORT]`-samtale pr. sag med brugeropdateringer og supportens svar.
- Sager kan gemmes centralt i Azure Blob Storage med lokal `tickets.json` som cache/fallback.

## Opsaetning

1. Kopier `appsettings.example.json` til `appsettings.json`.
2. Laeg hemmelige vaerdier i `appsettings.local.json` eller miljoevariabler.
3. Udfyld Graph loginoplysninger (`TenantId`, `ClientId`, `ClientSecret`).
4. Hvis sager skal gemmes i Azure, udfyld `BlobStorage` og saet `BlobStorage.Enabled` til `true`.
Du kan enten bruge `ConnectionString`, eller `AccountUrl` sammen med Azure AD / Managed Identity / `az login`.
5. Koer programmet:

```bash
dotnet run --project ITSupportSystem -- run
```

## Graph rettigheder

App-registreringen i Azure skal have Application permissions:

- `Mail.ReadWrite`
- `Mail.Send`

Og der skal gives admin consent.

## Feature toggles

I `appsettings.json` kan du styre test/scenarier:

- `Features.EnableTicketPolling`: slaa polling til/fra.
- `Features.EnableAutoAcknowledgements`: slaa kvitteringsmails til/fra.
- `Features.EnableUpdateAcknowledgements`: slaa kvittering ved opdateringer til/fra (anbefalet `false`).
- `Features.EnableOutgoingNotifications`: slaa svar/luk-mails til/fra.
- `Features.MoveClosedTicketToTrash`: flyt mails for lukkede sager til papirkurven.
- `Features.UseTestMailbox`: brug testpostkasse i stedet for `MailboxAddress`.
- `Features.TestMailboxAddress`: fx `teb@blind.dk`.
- `Features.SendSupportCardOnUpdates`: send support-kort ved hver opdatering (anbefalet `false` for mindre mailstoej).
- `Features.MoveProcessedIncomingToTrash`: flyt allerede behandlede indgaaende mails til papirkurven, saa inbox ikke dubleres.

## Azure Blob Storage

Hvis `BlobStorage.Enabled` er `true`, laeser systemet sager fra Azure Blob Storage ved opstart og gemmer derefter tilbage til samme blob ved hver opdatering.

- `BlobStorage.AccountUrl`: fx `https://minstorage.blob.core.windows.net`.
- `BlobStorage.ConnectionString`: connection string til storage account.
- `BlobStorage.ContainerName`: container til sagsdata, fx `itsupport`.
- `BlobStorage.BlobName`: blob-filnavn, fx `tickets.json`.

Den lokale `DataFilePath` bruges stadig som cache og fallback, hvis blob ikke kan laeses.

## Lokal secret-fil og miljoevariabler

`appsettings.local.json` bliver laest oven paa `appsettings.json` og er beregnet til lokale hemmeligheder, som ikke skal i Git.

Eksempel:

```json
{
	"Graph": {
		"ClientSecret": "lokal-hemmelighed"
	},
	"BlobStorage": {
		"ConnectionString": "lokal-connection-string"
	}
}
```

Følgende miljoevariabler kan bruges i stedet for filer:

- `GRAPH_TENANT_ID`
- `GRAPH_CLIENT_ID`
- `GRAPH_CLIENT_SECRET`
- `AZURE_STORAGE_CONNECTION_STRING`
- `TICKET_BLOB_ENABLED`
- `TICKET_BLOB_ACCOUNT_URL`
- `TICKET_BLOB_CONTAINER`
- `TICKET_BLOB_NAME`
- `FEATURE_USE_TEST_MAILBOX`
- `FEATURE_TEST_MAILBOX_ADDRESS`

## Kommandoer

- `run`: Kontinuerlig overvaagning af mailbox.
- `poll-once`: Hent nye mails en gang.
- `list`: Vis alle sager.
- `reply <TicketId> <Besked>`: Send svarmail i en sag.
- `close <TicketId> <Besked>`: Luk sagen og send afslutningsmail.

## Eksempler

```bash
dotnet run --project ITSupportSystem -- list
dotnet run --project ITSupportSystem -- reply SAG-20260723-0001 "Vi har nulstillet din adgangskode."
dotnet run --project ITSupportSystem -- close SAG-20260723-0001 "Problemet er loest."
```

## Knapper i mail

- `Svar pa sag` opretter en kommandomail til supportpostkassen i samme `[SUPPORT]`-samtaletrad for sagen.
- Systemet bruger teksten efter `#SVARSAG` som svar til brugeren og logger svaret pa sagen.
- `Luk sagen` opretter en kommandomail i samme `[SUPPORT]`-samtaletrad for sagen.
- Systemet lukker automatisk en aaben sag, naar der kommer en mail med `LUK [SAG-xxxx]` i emnet eller `#LUKSAG` i mailteksten.
