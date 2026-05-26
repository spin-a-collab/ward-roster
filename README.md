# WardRoster v3

Clinical ward rostering system for nursing staff.

## Project Structure

```
ward-roster/
├── index.html          — App entry point + PWA config
├── netlify.toml        — Netlify deployment config
├── package.json        — Dependencies
├── vite.config.js      — Build config
├── public/
│   └── manifest.json   — PWA manifest (home screen icon)
└── src/
    ├── main.jsx        — React entry point
    └── App.jsx         — Main application (all components)
```

## To update the app

1. Go to your GitHub repository
2. Click `src/App.jsx`
3. Click the pencil ✏️ icon (Edit)
4. Select all, paste new code
5. Click "Commit changes"
6. Netlify redeploys automatically in ~60 seconds

## Tech stack

- React 18
- Vite 5
- SheetJS (xlsx) for Excel import/export
- localStorage for data persistence (Supabase migration coming)

## iPad installation

1. Open your Netlify URL in Safari on iPad
2. Tap the Share button
3. Tap "Add to Home Screen"
4. Tap "Add"
5. Open from home screen — runs full screen
