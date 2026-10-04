# OneAI Gemini Integration — QA report

## Static validation
- `main.py` compiles successfully with `python -m py_compile`.
- Browser JavaScript extracted from `public/index.html` passes `node --check`.
- The Gemini key is **not** present in any public/client file or committed config. The backend reads `GEMINI_API_KEY` from the server environment / `.env`.

## End-to-end backend orchestration (mock provider mode)
All routes were executed through the same `/api/runs` wallet flow. Each completed, captured the configured final OneAI charge, released the 25% unused hold, produced an artifact, and created a bill/outbox record:
- Lovable → working `.html`
- ChatPRD → `.md` PRD
- Gamma → valid `.pptx`
- HeyGen → valid `.mp4`
- ElevenLabs → valid `.wav`
- Browse AI → `.json`
- Fireflies → `.md` transcript/summary
- AdCreative.ai → copy `.md` + square/landscape/portrait `.png`
- Typefully → `.md` social draft
- OneAI Search → grounded-answer-shaped result + `.md`

## Universal command box
`POST /api/command` was tested for:
- research/search request → OneAI Search run
- presentation request → Gamma run
- app build request → Lovable run
- meeting transcription request → `needs_input`, correctly requiring an upload instead of faking execution

## Gemini-only fallback path tests
External network access is not available in this build environment, so the supplied live Gemini credential was not transmitted or validated. The live Gemini-dependent code paths were tested with mocked Gemini/API responses:
- Gamma → Gemini deck JSON → local PPTX
- Lovable → Gemini HTML → standalone app
- ChatPRD → Gemini PRD
- HeyGen → Gemini scene plan → OneAI motion-card MP4 fallback
- ElevenLabs → Gemini TTS response → WAV
- Browse AI → page text → Gemini structured extraction
- Fireflies → Gemini multimodal transcript response
- AdCreative.ai → Gemini copy → rendered image set
- Typefully → Gemini social draft
- OneAI Search → Gemini search answer + source metadata

## Video fallback
The local motion-card video renderer was executed with FFmpeg. The produced QA MP4 was valid and 9 seconds long.

## Production caveats
SQLite, anonymous browser IDs, test credit top-ups, in-process background jobs, local artifact storage, and permissive CORS are still prototype choices. See `README.md` before using real money or scaling beyond a single instance.

## Retry/fallback regression checks

- Simulated primary-model `503` twice, then verified automatic fallback model success without surfacing the 503 to the caller.
- Simulated repeated `429` responses across primary + fallback models and verified the final error is sanitized (no raw provider JSON/status payload) and is safe for the UI.
- Re-ran Python compilation and frontend JavaScript syntax checks after the resilience changes.
- Re-ran a mock end-to-end PRD job through the HTTP API: 18.75 credits held, 15 charged, 3.75 released, artifact produced.
- Popular/All tool cards no longer display the `Gemini route ready` badge. Tool detail/Integrations can still show backend status.
- Cropped transparent whitespace from the OneAI logo asset and re-centered the hero orbit core so the visible logo is actually centered in the white core.
