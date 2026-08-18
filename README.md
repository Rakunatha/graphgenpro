# GraphGen Pro

Upload a survey spreadsheet (.xlsx or .csv) and GraphGen Pro will:

1. Auto-detect **independent variables** (demographics: Age, Gender, Education,
   Occupation, etc.) and **dependent variables** (attitude/Likert questions).
2. Generate an **SPSS-style grouped bar chart** for every independent × dependent
   variable pair (percent-within-group, data labels, legend).
3. Run a **one-way ANOVA** comparing an overall composite attitude score across
   the primary independent variable's groups (F, p, SPSS-style ANOVA table).
4. Let you **download a Word (.docx) report** containing every chart with an
   auto-written legend, the ANOVA table, an interpretation, and a
   Results/Discussion section — written by a free, built-in rule-based text
   generator (no paid AI API key needed).

It's a single small Flask app + one HTML page. No database, no signup, no
API keys required to run it.

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

## How the "AI" works (and why it's free)

No external AI API call is required. `auto_legend_text()` and the
Results/Discussion section in `build_docx()` in `app.py` generate natural-language
sentences from the actual computed percentages and ANOVA statistics (a
template/rule-based generator). This means:

- Zero API cost, zero API keys, works fully offline.
- If you *do* want nicer, more varied prose later, you can optionally wire in
  a free-tier LLM call (e.g. set an `ANTHROPIC_API_KEY` env var and post the
  same figures/stats to `/v1/messages`, then use that text instead of the
  template sentence) — the code is structured so that's a drop-in swap inside
  `auto_legend_text()` and `build_docx()`. This is optional and the app works
  fully without it.

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
