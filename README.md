# OneAI — Gemini-integrated full-stack prototype

This package turns the OneAI frontend into a backend-connected application with a universal Gemini command box, specialist vertical instructions, persistent wallet accounting, generated artifacts, optional native provider APIs, and billing-email support.

## Automatic Gemini resilience

The backend now treats `408`, `429`, `500`, `502`, `503`, `504`, and `529` as transient provider failures. For normal Gemini `generateContent` calls it:

1. retries the current model with bounded exponential backoff and `Retry-After` support,
2. falls back to compatible Gemini models supplied in `GEMINI_FALLBACK_MODELS`,
3. can auto-discover compatible `generateContent` models from the Gemini Models API so fallback names do not have to be hard-coded, and
4. only releases the wallet hold after all configured retry/fallback attempts are exhausted.

Raw 429/503 provider payloads are not shown to customers. If all attempts fail, the run receives a clean temporary-unavailability message and the full hold is released with zero charge. Meeting transcription uses the same model retry/fallback path. Gemini TTS retries its configured TTS model and can use optional `GEMINI_TTS_FALLBACK_MODELS`.

Useful tuning variables are `GEMINI_RETRY_ATTEMPTS_PER_MODEL`, `GEMINI_RETRY_BASE_SECONDS`, `GEMINI_RETRY_MAX_SECONDS`, `GEMINI_MAX_FALLBACK_MODELS`, and `GEMINI_AUTO_DISCOVER_FALLBACKS`.

## The home search box now does two jobs

The hero input is no longer only a keyword filter. `POST /api/command` sends the prompt to a server-side Gemini router:

- A question such as **“What changed in India’s EV policy this week?”** routes to **OneAI Search**, which uses Gemini with Google Search grounding and returns an answer plus source links.
- **“Make a 10-slide investor deck on…”** routes directly to Gamma/Presentations, reserves credits, generates the PPTX, settles the final charge, releases the remainder, and records/sends the bill.
- The same direct execution works for PRDs, web apps, videos, voiceovers, ad creatives, and social drafts.
- Tasks that genuinely need extra input do not fake completion. Meeting transcription routes to the Fireflies screen so the user can upload the recording. Browse extraction requires a public URL.

The router and every vertical have separate server-side system instructions. See `VERTICAL_INSTRUCTIONS.md`.

## One Gemini key can power all nine verticals

| Vertical | With only `GEMINI_API_KEY` | Optional native route |
|---|---|---|
| Lovable / Build | Gemini generates working standalone HTML | none assumed |
| ChatPRD / Product | Gemini generates an implementation-ready PRD | plan-dependent ChatPRD can be added later |
| Gamma / Presentations | Gemini plans the deck + OneAI creates a real `.pptx` | `GAMMA_API_KEY` |
| HeyGen / Video | Gemini creates scenes + OneAI renders a real motion-card `.mp4` | `HEYGEN_API_KEY`; optional Gemini Veo |
| ElevenLabs / Voice | Gemini TTS creates `.wav` | `ELEVENLABS_API_KEY` |
| Browse AI / Research & Data | safe public-page fetch + Gemini structured extraction | `BROWSEAI_API_KEY` + robot ID |
| Fireflies / Meetings | Gemini transcribes uploaded audio/video and produces actions | native webhook mode can be added later |
| AdCreative.ai / Marketing | Gemini writes variants + OneAI renders 3 creative sizes | partner API can be added later |
| Typefully / Social | Gemini creates the platform-ready draft | `TYPEFULLY_API_KEY` |

The app always labels the actual route used. A Gemini/local fallback is not presented as a native third-party result.

## Wallet flow

For every generated task, including OneAI Search:

1. Backend estimates the OneAI charge.
2. `estimate × 1.25` is reserved as a hold.
3. The selected vertical runs server-side.
4. On success, the final OneAI cost is captured and unused reserved credits are released.
5. On failure, usage capture is `0` and the full hold is released.
6. The run, ledger, output links and bill are persisted in SQLite.
7. If Resend is configured, the bill is actually emailed. Otherwise it is kept in the Bills screen as `not_configured`.

## API-key setup

Do **not** put a Gemini key inside `public/index.html`, JavaScript, GitHub, or any file served to the browser. This build reads it only from the server environment.

```bash
cd oneai-gemini-integrated
python -m venv .venv
source .venv/bin/activate          # Windows: .venv\\Scripts\\activate
pip install -r requirements.txt
cp .env.example .env
```

Then put your key in the local `.env`:

```text
GEMINI_API_KEY=your_key_here
```

`python-dotenv` is included, so the backend loads `.env` automatically. `.env` is ignored by git.

Run:

```bash
uvicorn main:app --reload --host 0.0.0.0 --port 8000
```

Open `http://localhost:8000`.

## Railway

1. Push this folder to GitHub.
2. Create a Railway service from the repository.
3. In **Railway → Variables**, create `GEMINI_API_KEY` and paste the key there. Do not upload `.env`.
4. Add a persistent volume mounted at `/data` and set `DATA_DIR=/data`.
5. Set `PUBLIC_BASE_URL=https://<your-domain>` so emailed artifact links are absolute.
6. Add `RESEND_API_KEY` + `RESEND_FROM` if bills should be delivered by email.
7. Add native provider keys only when you want the native branded route.

The included `railway.toml` starts `uvicorn main:app --host 0.0.0.0 --port $PORT` and uses `/api/health` as the health check.

## Important environment variables

```text
GEMINI_API_KEY=                 # required for live Gemini routes
GEMINI_MODEL=gemini-3.8-flash
GEMINI_TTS_MODEL=gemini-3.8-flash-lite-tts
GEMINI_TTS_VOICE=Kore
DATA_DIR=./data
PUBLIC_BASE_URL=http://localhost:8000
INITIAL_CREDITS=500
HOLD_BUFFER=0.25
MOCK_MODE=false
ALLOW_TEST_TOPUPS=true

# Optional native routes
GAMMA_API_KEY=
HEYGEN_API_KEY=
ELEVENLABS_API_KEY=
BROWSEAI_API_KEY=
BROWSEAI_ROBOT_ID=
TYPEFULLY_API_KEY=
TYPEFULLY_SOCIAL_SET_ID=
ENABLE_GEMINI_VEO=false

# Real billing email
RESEND_API_KEY=
RESEND_FROM=OneAI <billing@your-verified-domain.com>
```

## What remains prototype-grade

SQLite, anonymous browser IDs, test top-ups, in-process FastAPI background jobs, local artifact storage, and permissive CORS are fine for a demo but not for real money. Before launch, move wallet/run state to PostgreSQL, jobs to a durable queue, files to S3/R2, add authentication and rate limits, disable test top-ups, validate provider webhooks, and restrict CORS.

## Security

Provider keys never need to enter the browser. `.env` is already ignored. The Browse fallback blocks private/loopback/link-local/reserved addresses before fetching a URL to reduce SSRF risk. For production, also enforce outbound network controls and scan uploads.
