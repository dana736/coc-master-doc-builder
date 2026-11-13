# Cash Out Challenge | Company Master (CIM) Builder

A standalone Netlify project derived from the working Cash Out KPI Builder architecture. Produces an editable `.docx` company master directory with section navigation, linked document records, model-specific operations, a missing-files tracker, and an action plan. This is **not** a full broker-style Confidential Information Memorandum; it follows the supplied FollowUp.com **Master** reference.

## Deploy
1. Create a new private GitHub repo, or upload the entire contents of this ZIP into the root of a new repository. Do **not** overwrite your separate KPI Builder.
2. Connect that repository to a new Netlify site. `netlify.toml` handles `public/`, serverless functions and `/api/*` routes. Netlify should install Node 20+ dependencies automatically.
3. In Netlify **Environment variables**, set `TURNSTILE_SITE_KEY` and `TURNSTILE_SECRET_KEY` from Cloudflare Turnstile (configure your new Netlify hostname and any custom domain). Redeploy. Open the Netlify URL, complete the challenge, and test a DOCX download. Embed that URL in the Kajabi lesson using `kajabi-embed.html`.

Optional: `DAILY_PER_IP_LIMIT` (default 3), `DAILY_GLOBAL_LIMIT` (default 200). The DOCX tool has **no OpenAI dependency** and does not need `OPENAI_API_KEY`.

## Run locally
`npm install && npm test`; use `npm run dev` for Netlify Functions locally. Local download still requires valid Turnstile variables and an appropriate testing configuration.

## Member link format
Paste one record per line, `Exact document name | https://... | Department`. The last segment is optional. A recognized department gets the link in that section; matching an existing catalog title updates it; other records appear in `Other Existing Documents`. This does not upload, access, or validate linked private files. Existing links are hyperlinks only and access remains controlled by the source storage.

## Security and scope
- Turnstile is verified **server-side**. Per-IP and global daily Netlify Blobs counters are best-effort, not atomic.
- Generated documents are ephemeral responses, not stored on the app. Avoid pasting confidential document contents; paste links only.
- Kajabi's iframe does **not** establish membership authorization on the independently accessible Netlify URL.
- Do not share a buyer-facing copy with sensitive HR or financial information without reviewing permissions and redacting data.
- Suggested records are recommendations, **not assertions** that the records exist or are complete.
- The same Kajabi iframe remains valid only when reusing the exact Netlify URL. A separate deployment needs the new URL.
