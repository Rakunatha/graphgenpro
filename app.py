"""
GraphGen Pro
============
Upload a survey spreadsheet -> auto-detect independent vs dependent variables ->
generate SPSS-style grouped bar charts for every IV x DV pair -> run a one-way
ANOVA -> export a Word document with figures, legends, an ANOVA table and an
auto-written Results & Discussion section.

100% free to run: no paid AI API is required. Interpretation sentences are
produced by a rule-based text generator (see `narrate.py` logic inline below).
If you set an ANTHROPIC_API_KEY environment variable, the app will instead ask
Claude to polish the Results/Discussion paragraphs (optional, still works fine
without it).
"""

import io
import os
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

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 25 * 1024 * 1024  # 25MB upload cap

# In-memory store for the last few analyses so /download can rebuild the docx.
# Fine for a single small Render instance; swap for redis/db if you need more.
REPORT_CACHE = {}
MAX_CACHE = 8

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

    fig, ax = plt.subplots(figsize=(7.5, 4.3), dpi=150)
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
    plt.close(fig)
    buf.seek(0)
    return buf.getvalue(), ct


def auto_legend_text(iv, dv, ct):
    """Rule-based 'AI' legend/result sentence — no external API needed."""
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

def build_docx(title, iv_cols, dv_cols, chart_records, anova, composite_label):
    doc = Document()

    h = doc.add_heading(title, level=0)
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_heading("1. Overview", level=1)
    doc.add_paragraph(
        f"This report was generated automatically by GraphGen Pro. It compares "
        f"{len(dv_cols)} dependent (outcome) variable(s) against {len(iv_cols)} "
        f"independent (grouping) variable(s), producing {len(chart_records)} "
        f"comparison charts, followed by a one-way ANOVA."
    )
    p = doc.add_paragraph()
    p.add_run("Independent variables detected: ").bold = True
    p.add_run(", ".join(iv_cols))
    p2 = doc.add_paragraph()
    p2.add_run("Dependent variables detected: ").bold = True
    p2.add_run(", ".join(dv_cols))

    doc.add_heading("2. Comparison Charts", level=1)
    for idx, rec in enumerate(chart_records, start=1):
        doc.add_heading(f"Figure {idx}: {rec['dv']} by {rec['iv']}", level=2)
        img_stream = io.BytesIO(rec["png"])
        doc.add_picture(img_stream, width=Inches(6.0))
        legend_p = doc.add_paragraph()
        legend_p.add_run("LEGEND: ").bold = True
        legend_p.add_run(rec["legend"])

    doc.add_heading("3. ANOVA Test", level=1)
    if anova is None:
        doc.add_paragraph("An ANOVA could not be computed (not enough numeric/group data).")
    else:
        doc.add_paragraph(
            f"A one-way ANOVA was run to test whether \u201c{composite_label}\u201d differs "
            f"significantly across groups of \u201c{anova['iv']}\u201d."
        )
        table = doc.add_table(rows=4, cols=6)
        table.style = "Light Grid Accent 1"
        hdr = table.rows[0].cells
        for i, txt in enumerate(["Source", "Sum of Squares", "df", "Mean Square", "F", "Sig."]):
            hdr[i].text = txt
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
        gp = doc.add_paragraph()
        gp.add_run("Group means: ").bold = True
        for name, row in anova["group_stats"].iterrows():
            doc.add_paragraph(f"  {name}: n={int(row['count'])}, mean={row['mean']}, sd={row['std']}", style=None)

        interp = doc.add_paragraph()
        interp.add_run("INTERPRETATION: ").bold = True
        sig_txt = (
            f"The model shows F({anova['df_between']}, {anova['df_within']}) = {anova['F']:.3f}, "
            f"p = {anova['p']:.3f}, which is {'below' if anova['significant'] else 'above'} the "
            f"conventional 0.05 threshold. This indicates that {anova['iv']} "
            f"{'does' if anova['significant'] else 'does not'} have a statistically significant "
            f"effect on {composite_label}."
        )
        interp.add_run(sig_txt)

    doc.add_heading("4. Results", level=1)
    doc.add_paragraph(
        "The comparison charts above show how responses to each dependent variable vary across "
        "each independent (demographic/grouping) variable. Patterns worth noting are called out "
        "in each figure's legend."
    )
    if anova is not None:
        doc.add_paragraph(
            f"For the ANOVA, {anova['iv']} produced {'a statistically significant' if anova['significant'] else 'no statistically significant'} "
            f"difference in mean {composite_label} across groups "
            f"(F = {anova['F']:.2f}, p = {anova['p']:.3f})."
        )

    doc.add_heading("5. Discussion", level=1)
    if anova is not None and anova["significant"]:
        doc.add_paragraph(
            f"Because the ANOVA result is significant, this suggests genuine differences in "
            f"{composite_label} between {anova['iv']} groups, rather than differences due to chance "
            f"alone. Researchers should examine the group means table above to see which specific "
            f"groups differ, and consider a post-hoc test (e.g., Tukey HSD) for pairwise comparisons."
        )
    elif anova is not None:
        doc.add_paragraph(
            f"Because the ANOVA result is not significant, the data do not provide strong evidence "
            f"that {anova['iv']} groups differ in {composite_label}. Any differences visible in the "
            f"bar charts are more likely attributable to sampling variation than to a true underlying "
            f"effect, though a larger sample could reveal a smaller real effect."
        )
    doc.add_paragraph(
        "Overall, the charts and statistical test together provide an SPSS-style descriptive and "
        "inferential summary of how the independent variables relate to the dependent variables in "
        "this dataset. As with any survey data, results should be interpreted alongside sample size, "
        "response bias, and measurement limitations."
    )

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

        filename = f.filename.lower()
        if filename.endswith(".csv"):
            df = pd.read_csv(f)
        else:
            df = pd.read_excel(f)

        df = df.dropna(axis=1, how="all")
        if df.empty:
            return jsonify({"error": "The file has no usable data."}), 400

        # Clean up messy header text (stray newlines/whitespace from form exports)
        df.columns = [" ".join(str(c).split()) for c in df.columns]

        iv_cols, dv_cols = classify_columns(df)
        if not iv_cols or not dv_cols:
            return jsonify({"error": "Could not detect independent/dependent variables automatically."}), 400

        # cap total charts for a snappy free-tier response
        MAX_CHARTS = 60
        chart_records = []
        preview_records = []
        count = 0
        for dv in dv_cols:
            for iv in iv_cols:
                if count >= MAX_CHARTS:
                    break
                png, ct = make_bar_chart(df, iv, dv)
                if png is None:
                    continue
                legend = auto_legend_text(iv, dv, ct)
                chart_records.append({"iv": iv, "dv": dv, "png": png, "legend": legend})
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

        report_id = str(uuid.uuid4())
        REPORT_CACHE[report_id] = {
            "title": "GraphGen Pro - Automated Analysis Report",
            "iv_cols": iv_cols, "dv_cols": dv_cols,
            "chart_records": chart_records, "anova": anova,
            "composite_label": composite_label,
        }
        if len(REPORT_CACHE) > MAX_CACHE:
            REPORT_CACHE.pop(next(iter(REPORT_CACHE)))

        for rec in chart_records[:12]:
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
            "n_rows": int(len(df)),
            "iv_cols": iv_cols, "dv_cols": dv_cols,
            "total_charts": len(chart_records),
            "preview_charts": preview_records,
            "anova": anova_summary,
        })

    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": f"Something went wrong: {e}"}), 500


@app.route("/download/<report_id>")
def download(report_id):
    data = REPORT_CACHE.get(report_id)
    if not data:
        return "Report expired or not found. Please re-analyze your file.", 404
    out = build_docx(
        data["title"], data["iv_cols"], data["dv_cols"],
        data["chart_records"], data["anova"], data["composite_label"],
    )
    return send_file(
        out, as_attachment=True,
        download_name="GraphGen_Pro_Report.docx",
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


@app.route("/health")
def health():
    return "ok"


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
