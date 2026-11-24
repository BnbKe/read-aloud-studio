# Read Aloud Studio

Listen to PDFs, slides, spreadsheets and Word docs; ask questions with cited answers; get writing feedback; analyze data; export to Word/PDF/PowerPoint/Excel.

## Structure
- `public/` – the website (served as static files)
  - `index.html` – reader, Q&A, voices
  - `studio-pro.js` – Review, Edit, Convert & export
  - `studio-data.js` – Analyze data (stats, charts, ask-the-data, reports)
  - `privacy.html`, `_headers` (security headers)
- `src/worker.js` – one Cloudflare Worker: serves `public/` and the AI helper at `/api/ask`
- `wrangler.jsonc` – Worker config: Workers AI binding, per-visitor rate limit (20/min), model
- `tests/` – `npm test` runs the helper's security and behaviour tests

## Security
- `/api/ask` only answers pages from its own address (plus `EXTRA_ORIGINS`)
- Rate limit per visitor IP: 20 requests per minute
- Request size and conversation length caps; answer length caps
- Document text is treated as study material, never as instructions

## Develop and deploy
- `npm install`, then `npm test`
- `npx wrangler dev` runs locally (Workers AI calls go to Cloudflare)
- Deploys automatically: the Worker is connected to this GitHub repo with Workers Builds, so every push to `main` publishes.
