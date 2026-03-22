# POD Generator

Generate Proof of Delivery (POD) PDFs from an Excel file. Auto-splits into ZIPs of 3000 PDFs each. Deploy on **Railway** for 1000+ orders.

## Deploy on Railway

Railway supports long-running requests and enough memory for Puppeteer, so you can process **1000–26k+ orders** per upload.

### Deploy with Docker (recommended)

1. Push the repo to GitHub.
2. Go to [Railway](https://railway.app) → **New Project** → **Deploy from GitHub repo**.
3. Select this repo. Railway will detect the **Dockerfile** and deploy.
4. Add a **public domain** in the service → **Settings** → **Networking** (e.g. `your-app.up.railway.app`).
5. Deploy. No extra env vars needed; the Dockerfile sets `PUPPETEER_EXECUTABLE_PATH` and installs Chromium.

### Env vars (optional on Railway)

| Variable | Default | Use |
|----------|---------|-----|
| `CONCURRENT_PDFS` | 1 | Parallel PDFs (try 2–3 if you have more RAM). |
| `BROWSER_RESTART_EVERY` | 600 | Restart browser every N PDFs to avoid memory issues. |
| `PDF_TIMEOUT_MS` | 45000 | Timeout per PDF (ms). |

---

## Run locally

```bash
npm install
npm start
```

Uses full Puppeteer and defaults to concurrency 5. Open `http://localhost:3000`, upload an Excel file, get a ZIP of POD PDFs.
