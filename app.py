"""
GraphGen Pro
============
Upload a survey spreadsheet -> auto-detect independent vs dependent variables ->
generate SPSS-style grouped bar charts for every IV x DV pair -> run a one-way
ANOVA -> export a Word document with figures, legends, an ANOVA table and an
auto-written Results & Discussion section.

100% free to run. Text is written by Groq (https://console.groq.com), which has
a free API tier — set a GROQ_API_KEY environment variable to turn it on. If no
key is set, or the Groq call fails/times out for any reason, the app silently
falls back to a built-in rule-based sentence generator so it always keeps
working with zero configuration and zero cost.
"""

import io
import os
import gc
import json
import base64
import uuid
import traceback

import pandas as pd
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

from flask import Flask, request, render_template, jsonify, send_file
from scipy import stats

from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

# --- Groq (free-tier AI) -----------------------------------------------------
# pip install groq. Get a free key at https://console.groq.com/keys
# Set it as an environment variable: GROQ_API_KEY=gsk_...
GROQ_MODEL = os.environ.get("GROQ_MODEL", "llama-3.3-70b-versatile")
try:
    from groq import Groq
    _groq_client = Groq(api_key=os.environ["GROQ_API_KEY"]) if os.environ.get("GROQ_API_KEY") else None
except Exception:
    _groq_client = None

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 8 * 1024 * 1024  # keep Render free-tier requests bounded

# In-memory store for the last few analyses so /download can rebuild the docx.
# Fine for a single small Render instance; swap for redis/db if you need more.
REPORT_CACHE = {}
MAX_CACHE = 2

# ---------------------------------------------------------------------------
# Column classification
# ---------------------------------------------------------------------------

IV_KEYWORDS = [
    "age", "gender", "sex", "education", "qualification", "occupation",
    "job", "income", "marital", "experience", "platform", "frequency of",
    "region", "location", "designation", "year of study", "grade",
]

LIKERT_SCALES = [
    (["strongly disagree", "disagree", "neutral", "agree", "strongly agree"], 1),
    (["strongly disagree", "disagree", "neutral", "agree", "strongly agree", "strongly"], 1),
    (["very low", "low", "moderate", "high", "very high"], 1),
    (["do not trust at all", "trust a little", "neutral", "trust"], 1),
    (["never", "occasionally", "monthly", "weekly", "daily"], 1),
    (["no", "maybe", "yes"], 1),
]


def _norm(v):
    return str(v).strip().lower()


def classify_columns(df):
    """Return (iv_cols, dv_cols) using simple, transparent heuristics."""
    iv_cols, dv_cols = [], []
    for col in df.columns:
        lc = col.lower().strip()
        if lc == "timestamp" or lc.startswith("unnamed"):
            continue
        if any(k in lc for k in IV_KEYWORDS):
            iv_cols.append(col)
        else:
            dv_cols.append(col)
    return iv_cols, dv_cols


def build_ordinal_map(series):
    """Try to map a categorical Likert-like column to an ordered 1..N scale.
    Falls back to order-of-first-appearance if no known scale matches."""
    values = [v for v in series.dropna().unique().tolist()]
    norm_values = {_norm(v) for v in values}

    for scale, _ in LIKERT_SCALES:
        scale_norm = [s for s in scale]
        if norm_values.issubset(set(scale_norm)) and len(norm_values) >= 2:
            ordered = [s for s in scale_norm if s in norm_values]
            mapping = {}
            for v in values:
                nv = _norm(v)
                if nv in ordered:
                    mapping[v] = ordered.index(nv) + 1
            if len(mapping) == len(values):
                return mapping, True

    # fallback: alphabetical-ish, first-seen order (not a real scale)
    mapping = {v: i + 1 for i, v in enumerate(values)}
    return mapping, False


def is_numeric_series(series):
    return pd.api.types.is_numeric_dtype(series)


def to_numeric_dv(df, col):
    """Return a numeric Series for a DV column, and whether it is a real scale."""
    s = df[col]
    if is_numeric_series(s):
        return s.astype(float), True
    mapping, is_real_scale = build_ordinal_map(s.dropna())
    return s.map(mapping).astype(float), is_real_scale


# ---------------------------------------------------------------------------
# Chart generation (SPSS-style grouped bar chart, % within each IV group)
# ---------------------------------------------------------------------------

SPSS_PALETTE = ["#4472C4", "#ED7D31", "#A5A5A5", "#FFC000", "#5B9BD5", "#70AD47"]


def make_bar_chart(df, iv, dv):
    """Grouped bar chart: % of each DV category, split by IV group."""
    sub = df[[iv, dv]].dropna()
    if sub.empty or sub[iv].nunique() < 2 or sub[dv].nunique() < 2:
        return None, None

    ct = pd.crosstab(sub[iv], sub[dv], normalize="index") * 100
    ct = ct.round(1)

    # Smaller figsize/DPI keeps per-chart memory low -- important on Render's
    # free 512MB instances where a report with 20+ charts can otherwise push
    # the worker over its memory limit and get killed by the platform (which
    # shows up to the browser as a bare, non-JSON "Internal Server Error"
    # page rather than the app's own JSON error response).
    fig, ax = plt.subplots(figsize=(5.4, 3.0), dpi=80)
    try:
        ct.plot(kind="bar", ax=ax, color=SPSS_PALETTE[: len(ct.columns)], edgecolor="black", linewidth=0.5)

        for container in ax.containers:
            ax.bar_label(container, fmt="%.0f%%", fontsize=7, padding=1)

        ax.set_ylabel("Percent within group (%)", fontsize=9)
        ax.set_xlabel(iv, fontsize=9)
        ax.set_title(f"{dv}\nby {iv}", fontsize=10, wrap=True)
        ax.legend(title=dv, fontsize=7, title_fontsize=7, bbox_to_anchor=(1.02, 1), loc="upper left")
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)
        plt.xticks(rotation=20, ha="right", fontsize=8)
        plt.tight_layout()

        buf = io.BytesIO()
        fig.savefig(buf, format="png", bbox_inches="tight")
        buf.seek(0)
        return buf.getvalue(), ct
    finally:
        # Always release the figure, even if plotting raised partway through --
        # otherwise a single bad column can leak a full Figure per request.
        plt.close(fig)


def auto_legend_text(iv, dv, ct):
    """Rule-based legend/result sentence — the offline fallback, always available."""
    try:
        top_group = ct.max(axis=1).idxmax()
        top_cat = ct.loc[top_group].idxmax()
        top_pct = ct.loc[top_group, top_cat]
        return (
            f"The chart shows responses to \u201c{dv}\u201d broken down by {iv}. "
            f"The {top_group} group shows the strongest single pattern: {top_pct:.0f}% selected "
            f"\u201c{top_cat}\u201d, higher than any other group/category combination in this comparison."
        )
    except Exception:
        return f"The chart compares {dv} across {iv} groups."


def auto_result_sentence(iv, dv, ct):
    """Rule-based, offline 'Results'-style sentence for one chart (objective, factual)."""
    try:
        top_group = ct.max(axis=1).idxmax()
        top_cat = ct.loc[top_group].idxmax()
        top_pct = ct.loc[top_group, top_cat]
        return (
            f"According to the chart, {top_pct:.0f}% of {top_group} respondents selected "
            f"\u201c{top_cat}\u201d for \u201c{dv}\u201d, the strongest single pattern across {iv} groups."
        )
    except Exception:
        return f"The chart reports responses to \u201c{dv}\u201d across {iv} groups."


def auto_discussion_sentence(iv, dv, ct):
    """Rule-based, offline 'Discussion'-style sentence for one chart (interpretive)."""
    try:
        top_group = ct.max(axis=1).idxmax()
        top_cat = ct.loc[top_group].idxmax()
        top_pct = ct.loc[top_group, top_cat]
        return (
            f"This suggests {top_group} respondents feel most strongly about \u201c{dv}\u201d, "
            f"with {top_pct:.0f}% converging on \u201c{top_cat}\u201d -- a pattern worth weighing "
            f"against the other {iv} groups shown."
        )
    except Exception:
        return f"This points to a notable pattern in how {iv} groups responded to \u201c{dv}\u201d."


def chart_fact(iv, dv, ct):
    """Compact numeric summary of one chart, used as input to the AI writer."""
    try:
        top_group = ct.max(axis=1).idxmax()
        top_cat = ct.loc[top_group].idxmax()
        top_pct = float(ct.loc[top_group, top_cat])
        return {"iv": iv, "dv": dv, "top_group": str(top_group), "top_category": str(top_cat), "top_pct": round(top_pct, 1)}
    except Exception:
        return {"iv": iv, "dv": dv, "top_group": None, "top_category": None, "top_pct": None}


def ai_generate_narrative(facts, anova, composite_label):
    """
    One batched call to Groq's free-tier API that writes every chart legend
    plus one Results-style and one Discussion-style sentence per chart (in
    the same "ADDITIONAL ANALYSIS" report style used at this org: a running
    narrative that cites each chart as it goes, e.g. "...(fig: 3)"). Keeps
    this well within free rate limits even for 50+ charts since it's one
    batched request. Returns None on any failure so the caller falls back
    to the rule-based text -- the (fig: N) citations themselves are always
    appended in code afterward, not trusted to the model, so numbering
    never drifts even if the model's prose does.
    """
    if _groq_client is None:
        return None

    anova_summary = None
    if anova is not None:
        anova_summary = {
            "grouping_variable": anova["iv"],
            "outcome": composite_label,
            "F": round(anova["F"], 3),
            "p": round(anova["p"], 4),
            "significant": anova["significant"],
            "df_between": anova["df_between"],
            "df_within": anova["df_within"],
        }

    prompt = (
        "You are writing an academic-style SPSS survey report in the following house style: "
        "a numbered FIGURE/LEGEND block per chart, followed by a RESULT section and a DISCUSSION "
        "section that are each ONE continuous narrative walking through every chart in order "
        "(each chart gets 1-3 sentences citing concrete percentages/groups, then the narrative "
        "moves to the next chart) -- NOT a short high-level summary paragraph. "
        "Given the JSON facts below, respond with ONLY valid JSON (no markdown, no code fences) "
        "matching this exact shape:\n"
        '{"legends": ["...", ...], "results_sentences": ["...", ...], "discussion_sentences": ["...", ...]}\n\n'
        "Rules:\n"
        "- \"legends\" must have exactly one string per item in chart_facts, in the same order, "
        "each 1-2 sentences describing that specific chart's pattern (mention the top group/category/percent).\n"
        "- \"results_sentences\" must ALSO have exactly one string per item in chart_facts, in the same order: "
        "an objective 1-3 sentence description of that chart's data (concrete percentages/groups). "
        "Do not add a figure citation yourself -- it is appended automatically afterward.\n"
        "- \"discussion_sentences\" must ALSO have exactly one string per item in chart_facts, in the same order: "
        "a more interpretive 1-3 sentence take on what that chart's pattern suggests. Do not add a citation.\n"
        "- If anova_summary is present, make the very last discussion_sentences item also mention the ANOVA "
        "F, df, and p value in APA style and what it means for significance.\n"
        "- Do not invent numbers not present in the facts.\n\n"
        f"chart_facts = {json.dumps(facts)}\n"
        f"anova_summary = {json.dumps(anova_summary)}\n"
    )

    try:
        resp = _groq_client.chat.completions.create(
            model=GROQ_MODEL,
            messages=[
                {"role": "system", "content": "You write precise, concise, factual survey-analysis reports. Output strict JSON only."},
                {"role": "user", "content": prompt},
            ],
            temperature=0.4,
            max_tokens=4000,
            response_format={"type": "json_object"},
            timeout=20,
        )
        text = resp.choices[0].message.content
        data = json.loads(text)
        if (
            isinstance(data.get("legends"), list) and len(data["legends"]) == len(facts)
            and isinstance(data.get("results_sentences"), list) and len(data["results_sentences"]) == len(facts)
            and isinstance(data.get("discussion_sentences"), list) and len(data["discussion_sentences"]) == len(facts)
        ):
            return data
        return None
    except Exception:
        traceback.print_exc()
        return None


# ---------------------------------------------------------------------------
# ANOVA (one-way, SPSS-style table)
# ---------------------------------------------------------------------------

def run_anova(df, iv, dv_numeric_col_name, dv_values):
    work = pd.DataFrame({iv: df[iv], "dv": dv_values}).dropna()
    groups = [g["dv"].values for _, g in work.groupby(iv) if len(g) > 1]
    labels = [name for name, g in work.groupby(iv) if len(g) > 1]
    if len(groups) < 2:
        return None

    grand_mean = work["dv"].mean()
    n_total = len(work)
    k = len(groups)

    ss_between = sum(len(g) * (g.mean() - grand_mean) ** 2 for g in groups)
    ss_within = sum(((g - g.mean()) ** 2).sum() for g in groups)
    ss_total = ss_between + ss_within

    df_between = k - 1
    df_within = n_total - k

    ms_between = ss_between / df_between if df_between else float("nan")
    ms_within = ss_within / df_within if df_within else float("nan")
    F = ms_between / ms_within if ms_within else float("nan")
    p = stats.f.sf(F, df_between, df_within) if df_within > 0 else float("nan")

    group_stats = work.groupby(iv)["dv"].agg(["count", "mean", "std"]).round(2)

    return {
        "iv": iv,
        "dv_label": dv_numeric_col_name,
        "ss_between": ss_between, "ss_within": ss_within, "ss_total": ss_total,
        "df_between": df_between, "df_within": df_within,
        "ms_between": ms_between, "ms_within": ms_within,
        "F": F, "p": p,
        "group_stats": group_stats,
        "significant": bool(p < 0.05) if p == p else False,
    }


# ---------------------------------------------------------------------------
# Word report builder
# ---------------------------------------------------------------------------

def _set_base_font(doc):
    """Match the reference report's font: Times New Roman 12pt body text."""
    normal = doc.styles["Normal"]
    normal.font.name = "Times New Roman"
    normal.font.size = Pt(12)


def _labeled_paragraph(doc, label, text):
    """A paragraph like '**LEGEND:** some text', matching the reference style."""
    p = doc.add_paragraph()
    p.add_run(f"{label}: ").bold = True
    p.add_run(text)
    return p


def build_docx(title, iv_cols, dv_cols, chart_records, anova, composite_label,
                ai_results_sentences=None, ai_discussion_sentences=None, ai_used=False):
    doc = Document()
    _set_base_font(doc)

    # Title -- bold plain paragraph, not a big Word Heading style.
    title_p = doc.add_paragraph()
    title_p.add_run("ADDITIONAL ANALYSIS").bold = True
    doc.add_paragraph()

    fig_num = 0

    # One FIGURE / image / LEGEND block per chart, in order.
    for rec in chart_records:
        fig_num += 1
        fig_p = doc.add_paragraph()
        fig_p.add_run(f"FIGURE {fig_num}:").bold = True

        img_stream = io.BytesIO(rec["png"])
        doc.add_picture(img_stream, width=Inches(3.15))

        _labeled_paragraph(doc, "LEGEND", rec["legend"])
        doc.add_paragraph()

    # ANOVA presented as its own numbered figure, SPSS "ANOVA" table style.
    if anova is not None:
        fig_num += 1
        fig_p = doc.add_paragraph()
        fig_p.add_run(f"FIGURE {fig_num}:").bold = True

        cap = doc.add_paragraph()
        cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cap_run = cap.add_run("ANOVA")
        cap_run.bold = True

        table = doc.add_table(rows=4, cols=6)
        table.style = "Table Grid"
        hdr = table.rows[0].cells
        for i, txt in enumerate(["Source", "Sum of Squares", "df", "Mean Square", "F", "Sig."]):
            hdr[i].text = txt
            hdr[i].paragraphs[0].runs[0].bold = True
        rows_data = [
            ("Between Groups", anova["ss_between"], anova["df_between"], anova["ms_between"], f"{anova['F']:.3f}", f"{anova['p']:.3f}"),
            ("Within Groups", anova["ss_within"], anova["df_within"], anova["ms_within"], "", ""),
            ("Total", anova["ss_total"], anova["df_between"] + anova["df_within"], "", "", ""),
        ]
        for r, row in enumerate(rows_data, start=1):
            cells = table.rows[r].cells
            for c, val in enumerate(row):
                cells[c].text = f"{val:.3f}" if isinstance(val, float) else str(val)
        doc.add_paragraph()

        sig_txt = (
            f"The model shows F({anova['df_between']}, {anova['df_within']}) = {anova['F']:.3f}, "
            f"p = {anova['p']:.3f}, which is {'below' if anova['significant'] else 'above'} the "
            f"conventional 0.05 threshold. This indicates that {anova['iv']} "
            f"{'does' if anova['significant'] else 'does not'} have a statistically significant "
            f"effect on {composite_label}."
        )
        _labeled_paragraph(
            doc, "LEGEND",
            f"This ANOVA table evaluates whether {anova['iv']} significantly predicts {composite_label}."
        )
        _labeled_paragraph(doc, "INTERPRETATION", sig_txt)
        doc.add_paragraph()

    # RESULT: one continuous narrative, one entry per chart, each citing (fig: N).
    result_label_p = doc.add_paragraph()
    result_label_p.add_run("RESULT:").bold = True
    result_p = doc.add_paragraph()
    for idx, rec in enumerate(chart_records, start=1):
        sentence = (
            ai_results_sentences[idx - 1] if ai_results_sentences and idx - 1 < len(ai_results_sentences)
            else auto_result_sentence(rec["iv"], rec["dv"], rec["ct"])
        )
        result_p.add_run(sentence.strip() + " ")
        result_p.add_run(f"(fig: {idx}) ").bold = True
    doc.add_paragraph()

    # DISCUSSION: same shape, more interpretive tone, plus an ANOVA callout at the end.
    disc_label_p = doc.add_paragraph()
    disc_label_p.add_run("DISCUSSION").bold = True
    disc_p = doc.add_paragraph()
    for idx, rec in enumerate(chart_records, start=1):
        sentence = (
            ai_discussion_sentences[idx - 1] if ai_discussion_sentences and idx - 1 < len(ai_discussion_sentences)
            else auto_discussion_sentence(rec["iv"], rec["dv"], rec["ct"])
        )
        disc_p.add_run(sentence.strip() + " ")
        disc_p.add_run(f"(fig: {idx}) ").bold = True
    if anova is not None:
        closing = (
            f"Taken together with the ANOVA above (F = {anova['F']:.3f}, p = {anova['p']:.3f}), "
            f"{anova['iv']} {'does' if anova['significant'] else 'does not'} appear to have a "
            f"statistically significant relationship with {composite_label}, so any differences "
            f"seen across the charts should be weighed against that result and the usual caveats "
            f"of survey data (sample size, response bias, and measurement limitations)."
        )
        disc_p.add_run(closing)

    out = io.BytesIO()
    doc.save(out)
    out.seek(0)
    return out


# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/analyze", methods=["POST"])
def analyze():
    try:
        f = request.files.get("file")
        if not f:
            return jsonify({"error": "No file uploaded."}), 400

        # Cap extremely large spreadsheets *before* fully loading them --
        # capping rows only after pd.read_csv/read_excel has already parsed
        # the entire file into memory doesn't help; a large/wide file can
        # OOM the 512MB Render free instance during the read itself,
        # especially .xlsx (openpyxl commonly uses 10-20x the file's raw
        # size in RAM). Reading with nrows keeps peak memory bounded by the
        # cap, not by the uploaded file's size.
        MAX_ROWS = int(os.environ.get("MAX_ROWS", 1500))

        filename = f.filename.lower()
        if filename.endswith(".csv"):
            df = pd.read_csv(f, nrows=MAX_ROWS)
        else:
            df = pd.read_excel(f, nrows=MAX_ROWS)

        df = df.dropna(axis=1, how="all")
        if df.empty:
            return jsonify({"error": "The file has no usable data."}), 400

        # Clean up messy header text (stray newlines/whitespace from form exports)
        df.columns = [" ".join(str(c).split()) for c in df.columns]

        iv_cols, dv_cols = classify_columns(df)
        if not iv_cols or not dv_cols:
            return jsonify({"error": "Could not detect independent/dependent variables automatically."}), 400

        # cap total charts for a snappy, low-memory free-tier response
        # (override with the MAX_CHARTS env var on Render if you need more/fewer)
        MAX_CHARTS = int(os.environ.get("MAX_CHARTS", 6))
        chart_records = []
        preview_records = []
        facts = []
        count = 0
        for dv in dv_cols:
            for iv in iv_cols:
                if count >= MAX_CHARTS:
                    break
                png, ct = make_bar_chart(df, iv, dv)
                if png is None:
                    continue
                fallback_legend = auto_legend_text(iv, dv, ct)
                chart_records.append({"iv": iv, "dv": dv, "png": png, "legend": fallback_legend, "ct": ct})
                facts.append(chart_fact(iv, dv, ct))
                count += 1
            if count >= MAX_CHARTS:
                break

        # Build a composite DV score (mean of all Likert-encoded DVs) for the ANOVA
        numeric_cols = []
        for dv in dv_cols:
            numeric_series, is_scale = to_numeric_dv(df, dv)
            if is_scale:
                numeric_cols.append(numeric_series.rename(dv))
        anova = None
        composite_label = "overall attitude score"
        if numeric_cols:
            composite = pd.concat(numeric_cols, axis=1).mean(axis=1)
            primary_iv = iv_cols[0]
            anova = run_anova(df, primary_iv, composite_label, composite)

        # Try one batched Groq call to write every legend + per-chart Results/Discussion
        # sentences. Falls back silently (ai_used=False) if no key is set or the call fails.
        ai_used = False
        ai_results_sentences = None
        ai_discussion_sentences = None
        ai_data = ai_generate_narrative(facts, anova, composite_label)
        if ai_data:
            ai_used = True
            for rec, legend in zip(chart_records, ai_data["legends"]):
                rec["legend"] = legend
            ai_results_sentences = ai_data["results_sentences"]
            ai_discussion_sentences = ai_data["discussion_sentences"]

        # Build the DOCX once during analysis. Rebuilding all charts inside
        # /download could exceed Render's request timeout and return an HTML 500 page.
        title = "GraphGen Pro - Automated Analysis Report"
        report_docx = build_docx(
            title, iv_cols, dv_cols, chart_records, anova, composite_label,
            ai_results_sentences=ai_results_sentences,
            ai_discussion_sentences=ai_discussion_sentences,
            ai_used=ai_used,
        )
        report_bytes = report_docx.getvalue()
        del report_docx  # the in-memory Document + all its embedded images

        report_id = str(uuid.uuid4())
        REPORT_CACHE[report_id] = {
            "docx_bytes": report_bytes,
        }
        if len(REPORT_CACHE) > MAX_CACHE:
            REPORT_CACHE.pop(next(iter(REPORT_CACHE)))

        # Drop the (potentially large) source dataframe now that everything
        # needed from it has been extracted, and force a GC pass so this
        # worker's memory footprint doesn't keep climbing across requests.
        n_rows = len(df)
        del df
        gc.collect()

        # Only send a small preview to the browser; full images stay server-side.
        for rec in chart_records[:6]:
            preview_records.append({
                "iv": rec["iv"], "dv": rec["dv"], "legend": rec["legend"],
                "img": "data:image/png;base64," + base64.b64encode(rec["png"]).decode(),
            })

        anova_summary = None
        if anova is not None:
            anova_summary = {
                "iv": anova["iv"], "F": round(anova["F"], 3), "p": round(anova["p"], 3),
                "significant": anova["significant"],
            }

        return jsonify({
            "report_id": report_id,
            "n_rows": int(n_rows),
            "iv_cols": iv_cols, "dv_cols": dv_cols,
            "total_charts": len(chart_records),
            "preview_charts": preview_records,
            "anova": anova_summary,
            "ai_used": ai_used,
        })

    except MemoryError:
        traceback.print_exc()
        gc.collect()
        return jsonify({
            "error": "The file is too large/complex to process on this server's memory limit. "
                     "Try a smaller file, fewer columns, or lower MAX_ROWS/MAX_CHARTS."
        }), 500
    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": f"Something went wrong: {e}"}), 500


@app.route("/download/<report_id>")
def download(report_id):
    data = REPORT_CACHE.get(report_id)
    if not data:
        return jsonify({
            "error": "Report expired or not found. Please re-analyze your file."
        }), 404

    # Lightweight endpoint: the DOCX was already generated during /analyze.
    docx_bytes = data.get("docx_bytes")
    if not docx_bytes:
        return jsonify({
            "error": "The report file is unavailable. Please re-analyze your file."
        }), 500

    return send_file(
        io.BytesIO(docx_bytes),
        as_attachment=True,
        download_name="GraphGen_Pro_Report.docx",
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        max_age=0,
    )


@app.errorhandler(413)
def request_too_large(error):
    return jsonify({
        "error": "File is too large. Please upload an XLSX/CSV file smaller than 8 MB."
    }), 413


@app.errorhandler(Exception)
def handle_unexpected_error(error):
    # API routes return JSON instead of Flask/Render's HTML error page.
    traceback.print_exc()
    if request.path.startswith("/analyze") or request.path.startswith("/download/"):
        return jsonify({"error": "Internal server error. Check the server logs for details."}), 500
    return error


@app.route("/health")
def health():
    return "ok"


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
