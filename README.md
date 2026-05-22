# Li-ion Weld Pattern Reference Atlas v3

Netlify-ready Vite/React app for lithium-ion battery weld-pattern provenance.

## New in v3

- Add Entry tab for creating atlas entries in the website
- Entries persist in browser local storage
- Delete recently added entries
- Export current atlas to CSV
- Import atlas data from .xlsx, .xls, or .csv
- Side-by-side questioned evidence vs known exemplar image review
- Tailwind pinned to v3.4.17 to avoid Tailwind v4 PostCSS errors on Netlify

## Netlify build settings

Build command: npm run build

Publish directory: dist

Base directory: blank if these files are in the repository root.

## Local setup

npm install
npm run dev

## Important

Uploaded images and new entries are local to the browser. Export the atlas CSV to back up or share new entries. For production use, add authentication, server storage, evidence audit logs, and controlled known-exemplar media storage.
