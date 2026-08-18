# GraphGen Pro

Upload a survey spreadsheet (.xlsx or .csv) and GraphGen Pro will:

1. Auto-detect **independent variables** (demographics: Age, Gender, Education,
   Occupation, etc.) and **dependent variables** (attitude/Likert questions).
2. Generate an **SPSS-style grouped bar chart** for every independent × dependent
   variable pair (percent-within-group, data labels, legend).
3. Run a **one-way ANOVA** comparing an overall composite attitude score across
   the primary independent variable's groups (F, p, SPSS-style ANOVA table).
4. Let you **download a Word (.docx) report** containing every chart with an
   AI-written legend (via Groq's free API — see below), the ANOVA table, an
   interpretation, and a Results/Discussion section.

It's a single small Flask app + one HTML page. No database. Groq's free tier
needs only a free signup to get an API key (no credit card, no paid plan). If
you skip that setup entirely, the app still works using a built-in offline
text writer as a fallback (see below).

## Project structure

```
graphgen_pro/
├── app.py                 # Flask backend: analysis, charts, ANOVA, docx export
├── templates/
│   └── index.html         # The entire frontend (one page, plain HTML/CSS/JS)
├── requirements.txt
├── Procfile                # gunicorn start command (used by Render/Heroku)
├── render.yaml              # optional one-click Render blueprint
└── README.md
```

## Run it locally

```bash
cd graphgen_pro
python3 -m venv venv && source venv/bin/activate      # optional
pip install -r requirements.txt
python app.py
```

Open http://localhost:5000, upload a spreadsheet, click **Analyze Data**, then
**Download Full Word Report**.

## How the AI works (Groq, free tier)

Text (every chart's legend + the Results and Discussion paragraphs) is written
by [Groq](https://console.groq.com) — an LLM API with a **free tier**, no
credit card required.

1. Go to https://console.groq.com/keys and create a free API key.
2. Set it as an environment variable named `GROQ_API_KEY`:
   ```bash
   export GROQ_API_KEY=gsk_xxxxxxxxxxxxxxxxxxxx   # locally
   ```
   On Render: **Dashboard → your service → Environment → Add Environment
   Variable** → key `GROQ_API_KEY`, value = your key. Redeploy.
3. That's it — the app automatically detects the key and uses AI-written text.

**What actually happens under the hood:** all the numbers (which group had the
highest %, the ANOVA F/p, etc.) are computed in Python with pandas/scipy —
the AI is never asked to do math, only to turn already-computed facts into
natural sentences. All charts' legends plus the Results/Discussion section are
requested in a **single batched API call** (not one call per chart), so even
a report with 50-60 charts stays comfortably inside Groq's free rate limits.

**No key set, or the Groq call fails/times out?** The app automatically falls
back to a built-in, offline, rule-based sentence generator (`auto_legend_text`
in `app.py`) — the app never breaks and never requires payment. The web UI
shows a small badge ("✨ Written by Groq AI" vs. "Built-in writer") so you can
always see which one produced a given report.

You can change the model via the `GROQ_MODEL` env var (default:
`llama-3.3-70b-versatile`). See https://console.groq.com/docs/models for the
current list of free models.

## How variables are detected

- Any column whose name contains a keyword like `age`, `gender`, `education`,
  `occupation`, `income`, `platform`, `frequency of`, etc. → **independent
  variable**.
- Everything else (typically Likert-scale attitude questions) → **dependent
  variable**.
- Likert-style text answers (Strongly Disagree → Strongly Agree, Very Low →
  Very High, etc.) are automatically converted to a 1–5 numeric scale for the
  ANOVA; if the app can't recognize the wording pattern, it falls back to an
  order-of-appearance numbering (flagged internally as "not a verified
  scale" — check results if your columns use unusual custom wording).

If your spreadsheet uses very different column names, you can rename the
`IV_KEYWORDS` list near the top of `app.py` to match your survey's wording.

## Deploying to Render (free tier)

1. **Push this folder to a GitHub repo** (Render deploys from GitHub/GitLab).
   ```bash
   git init
   git add .
   git commit -m "GraphGen Pro"
   git branch -M main
   git remote add origin https://github.com/<you>/graphgen-pro.git
   git push -u origin main
   ```

2. **Create the Render service:**
   - Go to https://dashboard.render.com → **New** → **Web Service**.
   - Connect your GitHub repo.
   - Render will detect `render.yaml` automatically and pre-fill the settings.
     If it doesn't, set these manually:
     - **Environment**: `Python 3`
     - **Build Command**: `pip install -r requirements.txt`
     - **Start Command**: `gunicorn app:app --workers 2 --timeout 120`
     - **Instance Type**: `Free`

3. Click **Create Web Service**. Render will build and deploy; your app will
   be live at `https://graphgen-pro.onrender.com` (or similar) in a couple of
   minutes.

4. **Notes for the free tier:**
   - Free instances sleep after inactivity — the first request after idling
     can take ~30–60 seconds to wake up.
   - Free instances have limited RAM (512 MB). If your spreadsheet is very
     large or has many columns, you may want to lower `MAX_CHARTS` in `app.py`
     (default 60) to keep memory/response time low.
   - Uploaded files and generated reports are kept **in memory only** (not
     written to disk, not stored in a database), and are cleared when the
     service restarts or after the in-memory cache fills up (`MAX_CACHE = 8`
     most-recent reports). This is fine for a personal/demo tool; add a
     database or object storage if you need reports to persist long-term or
     across multiple server instances.

That's it — no API keys, no paid services, nothing else to configure.
