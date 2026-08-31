# Agency Deck Translator — Backend

FastAPI service wrapping the existing `ai_utils.py` / `pptx_utils.py` / `docx_utils.py` translation logic, for the Next.js frontend in `../frontend`.

## Development

```bash
pip install -r requirements.txt
export ANTHROPIC_API_KEY=...
uvicorn backend.main:app --reload --port 8000   # run from the repo root
```

## Deployment (Render)

Deployed via Docker using `./Dockerfile` (build context = repo root, since it needs the shared root-level Python files). See `../render.yaml`.

Required environment variables:
- `ANTHROPIC_API_KEY`
- `FRONTEND_ORIGIN` — the deployed frontend's origin, for CORS (e.g. `https://your-app.vercel.app`)
