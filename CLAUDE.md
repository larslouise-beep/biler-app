# Frimann Biler – Bil- og nøgleoversigt

Intern webapp til at holde styr på biler og bilnøgler på flere lokationer. Al tekst i appen er på dansk.
Design: Toyota brand visual identity (Signature Grey #282830, Signature Red #FF0022, TME Grey 1–6, skarpe hjørner, lette store overskrifter).

## Arkitektur
- `docs/index.html` – hele appen (HTML/CSS/JS i én fil). Udgives med GitHub Pages fra mappen `/docs` på branch `main`.
  - Taler med serveren via `fetch(API_URL, POST, Content-Type text/plain)` med `{action, token, payload}`.
  - `API_URL` øverst i scriptet = Apps Script-webappens `/exec`-URL.
  - Bruger telefonens GPS (navigator.geolocation) og kamera (getUserMedia + BarcodeDetector/ZXing) – kræver https.
  - QR-etiketter (75 × 25 mm, printes med window.print og @page) indeholder links: `<app-url>?b=<bil-id>` for bil, `&n=<nøgle-nr>` for nøgler.
  - Login: navn + PIN. Token gemmes i localStorage og udløber efter 10 min. inaktivitet (AUTO_LOGUD_MIN).
- `apps-script/Code.gs` + `appsscript.json` – serveren (Google Apps Script bundet til et Google-regneark).
  - `doPost` → `api(action, token, payload)`. Faner: Medarbejdere, Lokationer, Biler, Historik.
  - PIN gemmes som SHA-256 med salt. Sessioner i CacheService (max 6 t). Skrivninger beskyttes med LockService.
  - `opsaetning()` opretter/udvider fanerne og første admin (FOERSTE_ADMIN). Skal køres én gang manuelt i editoren (kræver godkendelse af tilladelser).

## Udgivelse
- Server: `clasp push` fra `apps-script/`. Opdatér den EKSISTERENDE implementering, så URL'en ikke ændres:
  `clasp deploy -i <DEPLOYMENT_ID>` (clasp 3: `clasp update-deployment <DEPLOYMENT_ID>`). Lav aldrig en ny implementering uden at opdatere API_URL.
- App: commit + push af `docs/index.html` til `main` → GitHub Pages opdaterer på 1–2 min.
- Deployment-id og URL'er noteres herunder, når de er oprettet.

## Noter fra opsætning
- Regneark-URL: https://docs.google.com/spreadsheets/d/1UrxFzWHeYyuXETgCfJteIR3d02yuSbl9ycsii5wHacM/edit
- Apps Script deployment-id: AKfycbwN2dV9b_Os1h5gzwxBpuffXHuAHctgAK30E7TaNP9qznT1j9ra-OxJ0tpspZKEr_ItDA (script-id: 1A0hOVX08LDpV3GeSZLJ_1LnNJYqCaUtgke5EcuI0Ofv8z2ohMOcXOfQS)
- API_URL (/exec): https://script.google.com/macros/s/AKfycbwN2dV9b_Os1h5gzwxBpuffXHuAHctgAK30E7TaNP9qznT1j9ra-OxJ0tpspZKEr_ItDA/exec
- GitHub Pages-URL:

## Regler
- Bevar dataformatet i arket (kolonnenavne i række 1). Nye kolonner tilføjes bagerst og i `ARK` i Code.gs.
- Læg aldrig PIN-koder, tokens eller .clasprc.json i git.
- Test ændringer i index.html i en browser, før de pushes.
- Webappens adgang styres af `webapp.access` i `appsscript.json` (skal være `ANYONE_ANONYMOUS`). Gem ikke implementeringer fra en gammel editor-fane – den kan nulstille manifestet til `MYSELF`, så API'et svarer med Google-login. Tjek med `curl -sL <API_URL>` → `{"ok":true,...}`.
