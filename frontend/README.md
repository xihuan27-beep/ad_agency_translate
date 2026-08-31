# Agency Deck Translator — Frontend

Next.js (App Router) frontend for the Agency Deck Translator. Talks to the FastAPI backend in `../backend`.

## Development

```bash
npm install
cp .env.example .env.local   # set NEXT_PUBLIC_API_URL to your local backend
npm run dev
```

The backend must be running separately (see `../backend`, defaults to `http://localhost:8000`).

## Deployment

Deploy this directory to Vercel as the project root, with `NEXT_PUBLIC_API_URL` set to the deployed backend's URL (e.g. the Render service URL). The backend is deployed separately — see `../render.yaml`.
