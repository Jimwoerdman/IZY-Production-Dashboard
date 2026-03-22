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
