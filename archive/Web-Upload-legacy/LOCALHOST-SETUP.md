# Lokal test af DBS Portal (Microsoft login)

## 1) Start lokal webserver
Kør fra projektroden:

```powershell
npm --prefix Website run serve
```

Portalen kører herefter på:
- http://localhost:5500/index.html
- http://127.0.0.1:5500/index.html (kan bruges i browseren, men redirect URI i Entra skal stadig være localhost)

## 2) Konfigurer Entra App Registration
I Entra Admin Center for jeres app registration, tilføj redirect URI:
- http://localhost:5500/index.html

Platformtype bør være:
- Single-page application (SPA)

Tilføj API permissions (Microsoft Graph, Delegated):
- User.Read
- Sites.Read.All
- Files.Read.All

Giv admin consent hvis jeres politik kræver det.

## 3) Sæt Client ID i portalen
I filen script.js, sæt:
- clientId = "<JERES_APP_CLIENT_ID>"

Tenant er allerede sat i koden til DBS tenant-ID.

Fil:
- Website/script.js

## 4) Test flow
1. Åbn http://localhost:5500/index.html
2. Klik "Log ind med Microsoft"
3. Log ind med konto fra DBS-tenant
4. Åbn menuen SharePoint og klik "Vis Fælles i portal"
5. Bekræft at mapper/filer vises i sektionen "SharePoint i portal"

## Fejlfinding
- Hvis login ikke virker, tjek at redirect URI matcher 100% (inkl. path).
- Entra tillader typisk ikke HTTP-IP som redirect URI for SPA. Brug kun:
	- http://localhost:5500/index.html
- Hvis SharePoint-listen ikke kan hentes, mangler der typisk consent til Sites.Read.All eller Files.Read.All.
