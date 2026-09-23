# IZY Print Dashboard

A simple web dashboard that reads live data from Google Sheets.

## How to run

1. Make sure your Google Sheet is shared as "Anyone with the link can view"
2. Open `index.html` in a browser — **but use a local server**, not by double-clicking the file (due to browser security restrictions on file:// requests)

### Easiest way to run locally

If you have Node.js installed:
```
npx serve .
```
Then open http://localhost:3000

Or with Python:
```
python3 -m http.server 8080
```
Then open http://localhost:8080

## Sheet ID
`1vIERVGUheXWkMS155VWfBEuCrUV4qXGYSUM9mIdppfc`

## Files
- `index.html` — dashboard layout
- `style.css` — styling
- `app.js` — data fetching, filtering, rendering

## Reliable Add Job — frontend v119

The existing Apps Script deployment now runs v153. The Add Job flow sends a stable UUID, checks saved receipts, and resumes incomplete files/sleeves/mockups without creating a second job. Partial failures keep the form intact. If a previous uncertain attempt has different input, the app asks before treating it as a new job. A browser reload preserves request references, not form contents or attachments.

Backend source and tests for the reliability layer are maintained with the integrated IZY workspace (`integrations/print/ReliableJobs.gs`). Apps Script still requires its own deployment; GitHub Pages only publishes this frontend. Do not redeploy the historical Code.gs in this repository alone: the live backend also requires the ReliableJobs module and wrapper patch.

Normal refresh loads app.js?v=119. Existing Google login and staff permissions remain unchanged.

## v120 — Eerste en laatste fles

De actieve printwachtrij opent de fotocontrole van de bestaande IZY-suite. Jim koppelt Ivans telefoon eenmalig via **Telefoon & toegang** op `/print/controle`. Foto’s, beoordelingen en tijdregistratie worden daar bewaard. De bestaande Workfile en PrintLog blijven ongewijzigd. De Google Apps Script-backend hoeft voor deze wijziging niet opnieuw gedeployed te worden.
