# Li-ion Weld Pattern Reference Atlas v5

Netlify-ready Vite/React app for lithium-ion battery weld-pattern provenance.

## New in v5

- Add Entry form now includes photo upload for new atlas entries
- Multiple photos can be attached to one atlas entry
- Photo previews appear before saving
- Saved atlas cards display attached photos
- Custom entry list shows photo count and thumbnails
- Atlas search includes photo filenames
- CSV export includes photo count and photo names
- Photos are stored in browser local storage as data URLs
- Existing dropdown-assisted entry fields remain available
- Tailwind pinned to v3.4.17 to avoid the Tailwind v4 PostCSS error

## Local setup

```bash
npm install
npm run dev
```

## Netlify build settings

Build command:

```bash
npm run build
```

Publish directory:

```text
dist
```

Base directory should be blank if these files are in the repository root.

## Deploy by GitHub drag-and-drop

Upload these items to the root of the GitHub repo:

```text
src/
package.json
postcss.config.js
tailwind.config.js
netlify.toml
index.html
README.md
```

Do not upload `node_modules`, `dist`, or `package-lock.json`.

## Important

Photos are saved locally in the browser. This static prototype does not upload photos or case data to a server. Large images can fill browser storage. For production use, add authentication, server storage, evidence audit logs, and controlled known-exemplar media storage.
