# Ekonomi

Veckovis koll på sparande och avkastning (Avanza, SAVR, bank).
Ersätter Joel_Ekonomi.xlsm. Datan sparas i din egen OneDrive (Appar › Ekonomi), inte här.

**Lägg aldrig upp `ekonomi-startdata.json` eller andra filer med dina siffror i det här repot – det är publikt.**

## Installation

1. **GitHub Pages**: Settings → Pages → Source: *Deploy from a branch* → `main` / `(root)` → Save.
   Adressen blir `https://ANVÄNDARNAMN.github.io/ekonomi/` (utan klient-ID öppnas appen i demoläge).
2. **Appregistrering hos Microsoft** (portal.azure.com → *App registrations* → *New registration*):
   - Name: `Ekonomi`
   - Supported account types: *Personal Microsoft accounts only*
   - Redirect URI: plattform *Single-page application (SPA)*, `https://ANVÄNDARNAMN.github.io/ekonomi/`
   - Kopiera *Application (client) ID*.
3. Redigera `config.js` här på GitHub och klistra in ID:t i `clientId: ''`.
4. Öppna adressen på datorn → Logga in → Importera `ekonomi-startdata.json`.
5. Android: öppna adressen i Chrome → ⋮ → *Lägg till på startskärmen* / *Installera app*.

## Uppdatera appen
Ladda upp de nya filerna över de gamla, men **skriv inte över `config.js`** (den innehåller ditt klient-ID).

## Filer
- `index.html`, `app.js` – gränssnittet
- `calc.js` – alla beräkningar (samma logik som den rättade Excel-filen)
- `store.js` – inloggning (OAuth2 PKCE) och lagring i OneDrive-appmappen
- `xlsx.js` – skapar Excel-backupen
- `sw.js`, `manifest.webmanifest`, `icons/` – installation och offlinestöd
- `config.js` – klient-ID

Testläge: lägg till `?demo` i adressen för att prova utan inloggning (data sparas bara i webbläsaren).
