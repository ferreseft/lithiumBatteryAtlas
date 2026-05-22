# Li-ion Weld Pattern Reference Atlas

Netlify-ready Vite/React app for lithium-ion battery weld-pattern provenance, evidence intake, comparison scoring, SOP workflow, and glossary reference.

## Local setup

```bash
npm install
npm run dev
```

## Netlify deployment

### Option A: Drag-and-drop deploy

1. Run:
   ```bash
   npm install
   npm run build
   ```
2. Go to Netlify > Sites > Add new site > Deploy manually.
3. Drag the generated `dist` folder into Netlify.

### Option B: GitHub deploy

1. Create a GitHub repository and push this project.
2. In Netlify, choose Add new site > Import an existing project.
3. Select the repository.
4. Build command: `npm run build`
5. Publish directory: `dist`

The included `netlify.toml` already sets those values.

## Production notes

- Uploaded evidence images are previewed locally in the browser only.
- This is a static frontend prototype. It does not store cases or upload files to a server.
- For production use, add authentication, persistent case storage, audit logs, and controlled exemplar image storage.
