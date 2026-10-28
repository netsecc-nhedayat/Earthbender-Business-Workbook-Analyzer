"""
Earthbender Business Insights — prototype for Earthbender Fencing / Saliani Group.

Built against the real Earthbender workbook schema:
    Customer, Email, JobID, Address, Suburb, Scope, Type, Material, Sleepers,
    Length in m, Height total in m, Quote $, Deposit, Final Pay, Total Pay,
    Job Start, Job End, Start2End Days, Days on site, Workers?, Hours worked,
    Worker Cost $, Mat Cost $, Total Cost $, Profit $, Profit %, Charged pm,
    Cost pm, Profit pm, Profit ph, Lead Source, Learnings

Run:  streamlit run app.py
"""

import io
import re
from datetime import datetime

import pandas as pd
import plotly.express as px
import streamlit as st

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, Image,
)

# ---------------------------------------------------------------------------
st.set_page_config(page_title="Earthbender Business Insights", layout="wide")
APP_TITLE = "Earthbender Business Insights"
APP_SUB = "Earthbender Fencing / Saliani Group — prototype"

# ---------------------------------------------------------------------------
# Canonical field detection. Patterns matched against lower-cased headers.
# ---------------------------------------------------------------------------
FIELD_LABELS = {
    "job_id":      "Job ID",
    "customer":    "Customer",
    "start_date":  "Job Start",
    "end_date":    "Job End",
    "material":    "Material",
    "type":        "Type",
    "scope":       "Scope",
    "lead_source": "Lead Source",
    "revenue":     "Quote $",
    "total_pay":   "Total Pay",
    "cost":        "Total Cost $",
    "profit":      "Profit $",
    "profit_pct":  "Profit %",
    "hours":       "Hours worked",
    "workers":     "Workers?",
    "suburb":      "Suburb",
    "learnings":   "Learnings",
}

FIELD_PATTERNS = {
    "job_id":      [r"^job\s*id$", r"^job\s*no", r"^id$", r"reference"],
    "customer":    [r"^customer$", r"^client$", r"^name$"],
    "start_date":  [r"job\s*start", r"^start", r"commence", r"date\s*in"],
    "end_date":    [r"job\s*end", r"^end", r"complet", r"finish", r"date\s*out"],
    "material":    [r"^material$", r"product"],
    "type":        [r"^type$"],
    "scope":       [r"^scope$"],
    "lead_source": [r"lead\s*source", r"^lead$", r"^source$", r"channel", r"referr"],
    "revenue":     [r"^quote", r"^revenue$", r"^amount$", r"^price$"],
    "total_pay":   [r"^total\s*pay$", r"^paid$", r"^invoice$"],
    "cost":        [r"total\s*cost", r"^cost$", r"^expense$"],
    "profit":      [r"^profit\s*\$", r"^profit$", r"gross\s*profit"],
    "profit_pct":  [r"profit\s*%", r"margin\s*%"],
    "hours":       [r"hours"],
    "workers":     [r"workers?$", r"^workers"],
    "suburb":      [r"^suburb$", r"^location$", r"^city$"],
    "learnings":   [r"learnings?", r"^notes?$", r"^comments?$"],
}

def humanize(c):
    return re.sub(r"[_\s]+", " ", str(c)).strip().title()

def detect_columns(df):
    found = {}
    for key, pats in FIELD_PATTERNS.items():
        matches = []
        for col in df.columns:
            name = str(col).strip().lower()
            if any(re.search(p, name) for p in pats):
                matches.append(col)
        if matches:
            found[key] = matches
    return found

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
MONEY_RE = re.compile(r"[^\d\.\-]")

def to_money(v):
    """Parse '$2,700' / '1,234.5' / 2700 / '' → float or NaN."""
    if pd.isna(v):
        return float("nan")
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip()
    if not s or s.lower() in {"n/a", "#div/0!", "na", "nan", "-"}:
        return float("nan")
    s = MONEY_RE.sub("", s)
    if s in {"", "-", "."}:
        return float("nan")
    try:
        return float(s)
    except ValueError:
        return float("nan")

def to_pct(v):
    """Parse '30%' / '0.30' / 30 → fraction 0.30 or NaN."""
    if pd.isna(v):
        return float("nan")
    if isinstance(v, (int, float)):
        # already a fraction if < 1.5, else percent
        return float(v) / 100 if abs(v) > 1.5 else float(v)
    s = str(v).strip().replace("%", "")
    try:
        f = float(s)
    except ValueError:
        return float("nan")
    return f / 100 if abs(f) > 1.5 else f

def parse_dates(series):
    return pd.to_datetime(series, errors="coerce", dayfirst=False)

def load_uploaded(file):
    name = file.name.lower()
    try:
        if name.endswith(".csv"):
            raw = file.read()
            for enc in ("utf-8", "utf-8-sig", "latin-1"):
                try:
                    return pd.read_csv(io.BytesIO(raw), encoding=enc), None
                except UnicodeDecodeError:
                    continue
                except pd.errors.EmptyDataError:
                    return None, "The CSV file is empty."
            return None, "Could not decode CSV."
        if name.endswith((".xlsx", ".xlsm", ".xls")):
            return pd.read_excel(file, engine="openpyxl"), None
        return None, f"Unsupported file type: '{file.name}'."
    except Exception as e:
        return None, f"Could not read the file: {e}"

# ---------------------------------------------------------------------------
# Session
# ---------------------------------------------------------------------------
def init_state():
    for k in ("df", "mapping", "source_name", "detected", "analysis", "meta"):
        st.session_state.setdefault(k, None)

# ---------------------------------------------------------------------------
# Section 1 — Import
# ---------------------------------------------------------------------------
def section_import():
    st.header("1. Import workbook")
    st.caption("Upload a CSV or Excel file. The original file is never modified.")
    up = st.file_uploader(
        "Choose a file",
        type=["csv", "xlsx", "xlsm", "xls"],
    )
    if up is None:
        return

    df, err = load_uploaded(up)
    if err:
        st.error(err); return
    if df is None or df.empty:
        st.error("The file loaded but contained no rows."); return

    df.columns = [str(c).strip() for c in df.columns]
    st.session_state["df"] = df
    st.session_state["source_name"] = up.name
    st.session_state["detected"] = detect_columns(df)

    c1, c2, c3 = st.columns(3)
    c1.metric("Rows", f"{len(df):,}")
    c2.metric("Columns", df.shape[1])
    c3.metric("Missing cells", f"{int(df.isna().sum().sum()):,}")

    with st.expander("Available columns", expanded=False):
        st.write([str(c) for c in df.columns])

    with st.expander("Data preview (first 20 rows)", expanded=True):
        st.dataframe(df.head(20), use_container_width=True)

    # Date validity
    date_candidates = [c for c in df.columns if re.search(r"date|day|start|end", str(c).lower())]
    if date_candidates:
        st.subheader("Date column check")
        for c in date_candidates:
            parsed = parse_dates(df[c])
            ok = int(parsed.notna().sum())
            bad = int(len(parsed) - ok)
            st.write(f"- `{c}`: {ok} valid · {bad} invalid")

# ---------------------------------------------------------------------------
# Section 2 — Mapping
# ---------------------------------------------------------------------------
def mapping_ui(df):
    st.header("2. Column mapping")
    st.caption("Columns are auto-detected where possible. Adjust any mapping manually.")
    detected = st.session_state.get("detected") or detect_columns(df)
    options = ["— none —"] + [str(c) for c in df.columns]

    def picker(key, label):
        default_list = detected.get(key, [])
        default = str(default_list[0]) if default_list else "— none —"
        if default not in options:
            default = "— none —"
        return st.selectbox(label, options, index=options.index(default), key=f"map_{key}")

    mapping = {}
    keys = list(FIELD_LABELS.keys())
    cols = st.columns(3)
    for i, k in enumerate(keys):
        with cols[i % 3]:
            v = picker(k, FIELD_LABELS[k])
            mapping[k] = None if v == "— none —" else v

    st.session_state["mapping"] = mapping
    return mapping

# ---------------------------------------------------------------------------
# Analysis build
# ---------------------------------------------------------------------------
def build_analysis(df, m):
    a = df.copy()
    meta = {"usable": {}, "notes": []}

    # Dates
    for k in ("start_date", "end_date"):
        col = m.get(k)
        if col and col in a.columns:
            a[f"_parsed_{k}"] = parse_dates(a[col])
            meta["usable"][k] = a[f"_parsed_{k}"].notna().any()
        else:
            a[f"_parsed_{k}"] = pd.NaT
            meta["usable"][k] = False

    # Numbers
    for k, fn in (("revenue", to_money), ("total_pay", to_money),
                  ("cost", to_money), ("profit", to_money),
                  ("hours", to_money)):
        col = m.get(k)
        if col and col in a.columns:
            a[f"_num_{k}"] = a[col].apply(fn)
            meta["usable"][k] = a[f"_num_{k}"].notna().any()
        else:
            a[f"_num_{k}"] = float("nan")
            meta["usable"][k] = False

    # Percent
    col = m.get("profit_pct")
    if col and col in a.columns:
        a["_num_profit_pct"] = a[col].apply(to_pct)
        meta["usable"]["profit_pct"] = a["_num_profit_pct"].notna().any()
    else:
        a["_num_profit_pct"] = float("nan")
        meta["usable"]["profit_pct"] = False

    # Effective date for month chart: prefer start, else end
    if meta["usable"]["start_date"]:
        a["_eff_date"] = a["_parsed_start_date"]
        a["_eff_src"] = m["start_date"]
    elif meta["usable"]["end_date"]:
        a["_eff_date"] = a["_parsed_end_date"]
        a["_eff_src"] = m["end_date"]
    else:
        a["_eff_date"] = pd.NaT
        a["_eff_src"] = None

    # Derived profit if not present
    if (not meta["usable"].get("profit")) and meta["usable"].get("revenue") and meta["usable"].get("cost"):
        a["_num_profit"] = a["_num_revenue"] - a["_num_cost"]
        meta["usable"]["profit"] = True
        meta["notes"].append("Profit was derived as Quote − Total Cost (no profit column mapped).")

    return a, meta

# ---------------------------------------------------------------------------
# Section 3 — Dashboard
# ---------------------------------------------------------------------------
def section_dashboard(a, meta, m):
    st.header("3. Business dashboard")

    f1, f2, f3 = st.columns(3)
    date_range = None
    if a["_eff_date"].notna().any():
        dmin, dmax = a["_eff_date"].min().date(), a["_eff_date"].max().date()
        date_range = f1.date_input("Date range", value=(dmin, dmax),
                                   min_value=dmin, max_value=dmax)

    mat_sel, lead_sel = [], []
    if m.get("material") in a.columns:
        mat_sel = f2.multiselect("Material",
            sorted(a[m["material"]].dropna().astype(str).unique()), default=[])
    if m.get("lead_source") in a.columns:
        lead_sel = f3.multiselect("Lead source",
            sorted(a[m["lead_source"]].dropna().astype(str).unique()), default=[])

    view = a.copy()
    if date_range and len(date_range) == 2:
        lo = pd.Timestamp(date_range[0])
        hi = pd.Timestamp(date_range[1]) + pd.Timedelta(days=1)
        view = view[(view["_eff_date"] >= lo) & (view["_eff_date"] < hi)]
    if mat_sel:
        view = view[view[m["material"]].astype(str).isin(mat_sel)]
    if lead_sel:
        view = view[view[m["lead_source"]].astype(str).isin(lead_sel)]

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Jobs (filtered)", f"{len(view):,}")
    k2.metric("Jobs (total)", f"{len(a):,}")
    k3.metric("Revenue", f"${view['_num_revenue'].sum():,.0f}" if meta["usable"].get("revenue") else "n/a")
    k4.metric("Profit", f"${view['_num_profit'].sum():,.0f}" if meta["usable"].get("profit") else "n/a")

    # Jobs by month
    st.subheader("Jobs by month")
    if view["_eff_date"].notna().any():
        months = (view.dropna(subset=["_eff_date"])
                  .assign(month=lambda d: d["_eff_date"].dt.to_period("M").dt.to_timestamp())
                  .groupby("month").size().reset_index(name="jobs"))
        st.plotly_chart(px.bar(months, x="month", y="jobs", text="jobs",
                               title="Jobs per month"), use_container_width=True)
        st.caption(f"Date column used: `{a['_eff_src'].iloc[0]}`")
    else:
        st.info("No usable date column — cannot plot jobs by month.")

    # Counts
    st.subheader("Job counts by category")
    cA, cB = st.columns(2)
    for w, key, title in ((cA, "material", "Jobs by material"),
                          (cB, "lead_source", "Jobs by lead source")):
        col = m.get(key)
        if col and col in view.columns:
            counts = (view[col].astype(str).replace("nan", "(missing)")
                      .value_counts().reset_index())
            counts.columns = [FIELD_LABELS[key], "jobs"]
            w.plotly_chart(px.bar(counts, x=FIELD_LABELS[key], y="jobs",
                                  text="jobs", title=title), use_container_width=True)
        else:
            w.info(f"No column mapped for {title.lower()}.")

    # Revenue & profit by material
    if meta["usable"].get("revenue") and m.get("material"):
        st.subheader("Revenue & profit by material")
        grp = view.groupby(view[m["material"]].astype(str)).agg(
            revenue=("_num_revenue", "sum"),
            profit=("_num_profit", "sum") if meta["usable"].get("profit") else ("_num_revenue", "sum"),
            jobs=("_num_revenue", "size"),
        ).reset_index().rename(columns={"_num_revenue": "material"})
        grp = grp[grp["material"] != "nan"]
        st.plotly_chart(px.bar(grp, x="material", y=["revenue", "profit"],
                               barmode="group",
                               title="Revenue vs Profit by material"),
                        use_container_width=True)
        if meta["usable"].get("profit"):
            grp["margin_%"] = grp.apply(
                lambda r: (r["profit"] / r["revenue"] * 100) if r["revenue"] else None,
                axis=1)
            st.dataframe(grp[["material", "jobs", "revenue", "profit", "margin_%"]],
                         use_container_width=True)
            st.caption("Margin % = profit ÷ quote × 100 (only when quote ≠ 0).")
        else:
            st.info("Profit column not mapped — margin % not shown.")
    else:
        st.info("Revenue + material required for this chart.")

    st.session_state["dash_view"] = view

# ---------------------------------------------------------------------------
# Section 4 — Errors
# ---------------------------------------------------------------------------
def section_errors(a, meta, m):
    st.header("4. Potential errors")
    st.caption("Flags are signals for review, not proof of a business mistake.")
    flags = []

    # 1. End before start
    if meta["usable"].get("start_date") and meta["usable"].get("end_date"):
        mask = (a["_parsed_start_date"].notna() & a["_parsed_end_date"].notna()
                & (a["_parsed_end_date"] < a["_parsed_start_date"]))
        for idx in a[mask].index:
            flags.append({"row": idx, "reason": "End date before start date",
                          "suggestion": "Check both dates; likely a typo or swapped year."})

    # 2. Missing job IDs
    if m.get("job_id") and m["job_id"] in a.columns:
        mask = a[m["job_id"]].isna() | (a[m["job_id"]].astype(str).str.strip() == "")
        for idx in a[mask].index:
            flags.append({"row": idx, "reason": "Missing job ID",
                          "suggestion": "Assign a unique JobID for traceability."})

    # 3. Duplicate job IDs
    if m.get("job_id") and m["job_id"] in a.columns:
        dup = a[a[m["job_id"]].notna() &
                a.duplicated(subset=[m["job_id"]], keep=False)]
        for idx in dup.index:
            flags.append({"row": idx, "reason": "Duplicate JobID",
                          "suggestion": "Confirm if legitimate repeat or a re-entry."})

    # 4. Missing financials
    for key, label in (("revenue", "Quote $"), ("cost", "Total Cost $"),
                       ("profit", "Profit $")):
        if meta["usable"].get(key):
            mask = a[f"_num_{key}"].isna()
            for idx in a[mask].index:
                flags.append({"row": idx, "reason": f"Missing {label}",
                              "suggestion": "Fill in or mark as n/a."})

    # 5. Unusual margins (only if both exist)
    if meta["usable"].get("revenue") and meta["usable"].get("profit"):
        rev = a["_num_revenue"]; pr = a["_num_profit"]
        marg = pr / rev.where(rev != 0)
        bad = marg.notna() & ((marg < -0.5) | (marg > 1.5))
        for idx in a[bad].index:
            flags.append({"row": idx,
                          "reason": f"Unusual profit margin ({marg[idx]:.0%})",
                          "suggestion": "Check quote/cost — could be a data entry or discount."})

    if not flags:
        st.success("No potential errors detected with the current mapping."); return

    flag_df = pd.DataFrame(flags)
    st.warning(f"{len(flag_df)} potential issue(s) flagged.")
    st.dataframe(flag_df, use_container_width=True)
    with st.expander("Preview flagged rows"):
        st.dataframe(a.loc[flag_df['row'].unique()], use_container_width=True)

# ---------------------------------------------------------------------------
# Section 5 — Recommendations
# ---------------------------------------------------------------------------
def build_recommendations(a, view, meta, m):
    recs = []
    def add(area, text, evidence, metric, ok=True):
        recs.append(dict(area=area, text=text, evidence=evidence, metric=metric, supported=ok))

    if len(view) == 0:
        add("Data", "Filters returned no jobs.", "0 rows matched.",
            "Apply wider filters."); return recs

    # Marketing
    if m.get("lead_source") and m["lead_source"] in view.columns:
        ls = view[m["lead_source"]].astype(str).replace("nan", "(missing)")
        cnt = ls.value_counts()
        top = cnt.index[0]; share = cnt.iloc[0] / cnt.sum() * 100
        add("Marketing", f"Lead source '{top}' = {share:.0f}% of filtered jobs.",
            f"{cnt.iloc[0]} of {cnt.sum()} jobs.",
            "Share of new jobs per lead source; track cost per lead.")
        if len(cnt) >= 2:
            add("Marketing",
                f"Run a small test on '{cnt.index[1]}' to reduce single-source reliance.",
                f"Runner-up '{cnt.index[1]}' = {cnt.iloc[1]} jobs vs '{top}' {cnt.iloc[0]}.",
                "Jobs per lead source per month.")
    else:
        add("Marketing", "Cannot assess — no Lead Source mapped.",
            "No lead source column.", "n/a", ok=False)

    # Referrals
    if m.get("lead_source") and m["lead_source"] in view.columns:
        ref = view[m["lead_source"]].astype(str).str.lower().str.contains("referr").sum()
        if ref > 0:
            add("Customer referrals",
                f"{ref} referral job(s) in filtered set.",
                f"{ref} of {len(view)} jobs.",
                "Referral jobs per month; referral conversion rate.")
        else:
            add("Customer referrals",
                "No referral-type leads detected — consider a formal referral scheme.",
                "No lead source matches 'referral'.",
                "Referral jobs per month.")
    else:
        add("Customer referrals", "Cannot assess — no Lead Source mapped.",
            "No lead source column.", "n/a", ok=False)

    # Profitability
    if meta["usable"].get("revenue") and m.get("material"):
        rev = view.groupby(view[m["material"]].astype(str))["_num_revenue"].sum()
        rev = rev[rev.index != "nan"].sort_values(ascending=False)
        if not rev.empty:
            top = rev.index[0]; share = rev.iloc[0] / rev.sum() * 100 if rev.sum() else 0
            add("Profitability",
                f"Top revenue material '{top}' = {share:.0f}% of revenue.",
                f"Quote total ${rev.iloc[0]:,.0f}.",
                "Revenue per material per month.")
        if meta["usable"].get("profit"):
            mg = (view.groupby(view[m["material"]].astype(str))
                  .apply(lambda d: d["_num_profit"].sum() / d["_num_revenue"].sum()
                         if d["_num_revenue"].sum() else None).dropna().sort_values())
            if not mg.empty:
                add("Profitability",
                    f"Review pricing on '{mg.index[0]}' (lowest margin material).",
                    f"Margin = {mg.iloc[0]:.0%}.",
                    "Margin % per material, tracked monthly.")
    else:
        add("Profitability", "Cannot assess — map Quote $ and Profit $.",
            "Missing revenue/profit.", "n/a", ok=False)

    # Cost control
    if meta["usable"].get("cost") and meta["usable"].get("revenue"):
        r = view["_num_cost"].sum() / view["_num_revenue"].sum() if view["_num_revenue"].sum() else None
        if r is not None:
            add("Cost control", f"Overall cost ratio = {r:.0%} of quote.",
                "Total Cost ÷ Quote.", "Cost ratio trend by month.")
    else:
        add("Cost control", "Cannot assess — map Total Cost $.",
            "No cost column.", "n/a", ok=False)

    # Data quality
    dq = 0
    if m.get("job_id") and m["job_id"] in a.columns:
        dq += int((a[m["job_id"]].isna() |
                   (a[m["job_id"]].astype(str).str.strip() == "")).sum())
    if meta["usable"].get("start_date") and meta["usable"].get("end_date"):
        dq += int((a["_parsed_start_date"].notna() & a["_parsed_end_date"].notna()
                   & (a["_parsed_end_date"] < a["_parsed_start_date"])).sum())
    if dq:
        add("Data quality", f"{dq} row(s) flagged as data-quality issues.",
            "See Potential errors tab.", "Flag count per import; aim downward.")
    else:
        add("Data quality", "No obvious data-quality issues with current mapping.",
            "0 flags.", "Flag count per import.")

    return recs

def section_recommendations(a, view, meta, m):
    st.header("5. Business recommendations")
    st.caption("Generated with plain Python rules from the calculated results above.")
    recs = build_recommendations(a, view, meta, m)
    st.session_state["recs"] = recs
    for r in recs:
        if r["supported"]:
            st.markdown(f"**{r['area']}** — {r['text']}")
        else:
            st.info(f"{r['area']} — {r['text']}")
        st.caption(f"Evidence: {r['evidence']}  |  Suggested metric: {r['metric']}")
        st.divider()

# ---------------------------------------------------------------------------
# Section 6 — Export
# ---------------------------------------------------------------------------
def make_chart_pngs(view, meta, m):
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    out = []
    if view["_eff_date"].notna().any():
        s = (view.dropna(subset=["_eff_date"])
             .assign(month=lambda d: d["_eff_date"].dt.to_period("M").dt.to_timestamp())
             .groupby("month").size())
        fig, ax = plt.subplots(figsize=(7, 3.5))
        s.plot(kind="bar", ax=ax, color="#203864"); ax.set_title("Jobs per month")
        fig.tight_layout(); b = io.BytesIO(); fig.savefig(b, format="png", dpi=120)
        plt.close(fig); out.append(("Jobs per month", b.getvalue()))
    if meta["usable"].get("revenue") and m.get("material") in view.columns:
        rev = view.groupby(view[m["material"]].astype(str))["_num_revenue"].sum()
        rev = rev[rev.index != "nan"]
        if not rev.empty:
            fig, ax = plt.subplots(figsize=(7, 3.5))
            rev.plot(kind="bar", ax=ax, color="#4C7A34")
            ax.set_title("Revenue by material")
            fig.tight_layout(); b = io.BytesIO(); fig.savefig(b, format="png", dpi=120)
            plt.close(fig); out.append(("Revenue by material", b.getvalue()))
    return out

def build_pdf(summary_lines, kpis, chart_imgs, warnings_text, recommendations):
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4,
                            leftMargin=18*mm, rightMargin=18*mm,
                            topMargin=18*mm, bottomMargin=18*mm)
    styles = getSampleStyleSheet()
    story = [
        Paragraph(f"<b>{APP_TITLE}</b>", styles["Title"]),
        Paragraph(APP_SUB, styles["Normal"]),
        Paragraph(f"Generated {datetime.now():%Y-%m-%d %H:%M}", styles["Italic"]),
        Spacer(1, 8),
        Paragraph("<b>Summary</b>", styles["Heading2"]),
    ]
    for line in summary_lines:
        story.append(Paragraph(line, styles["Normal"]))
    story.append(Spacer(1, 6))

    if kpis:
        story.append(Paragraph("<b>Key metrics</b>", styles["Heading2"]))
        t = Table([["Metric", "Value"]] + kpis, hAlign="LEFT")
        t.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#203864")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
            ("FONTSIZE", (0, 0), (-1, -1), 9),
        ]))
        story.append(t); story.append(Spacer(1, 6))

    if chart_imgs:
        story.append(PageBreak())
        story.append(Paragraph("<b>Charts</b>", styles["Heading2"]))
        for name, img in chart_imgs:
            story.append(Paragraph(name, styles["Heading3"]))
            story.append(Image(io.BytesIO(img), width=160*mm, height=90*mm))
            story.append(Spacer(1, 6))

    story.append(PageBreak())
    story.append(Paragraph("<b>Data-quality warnings</b>", styles["Heading2"]))
    story.append(Paragraph(warnings_text, styles["Normal"]))
    story.append(Spacer(1, 6))
    story.append(Paragraph("<b>Recommendations</b>", styles["Heading2"]))
    for r in recommendations:
        story.append(Paragraph(
            f"<b>{r['area']}</b> — {r['text']}<br/>"
            f"<i>Evidence:</i> {r['evidence']}<br/>"
            f"<i>Test metric:</i> {r['metric']}", styles["Normal"]))
        story.append(Spacer(1, 4))

    doc.build(story)
    return buf.getvalue()

def section_export(a, view, meta, m):
    st.header("6. Export report")
    st.caption("Downloads are built from the analysed data — no fabricated results.")
    recs = st.session_state.get("recs", [])

    summary = [
        f"Source file: {st.session_state.get('source_name', 'unknown')}",
        f"Rows analysed: {len(a)}",
        f"Rows after filters: {len(view)}",
        f"Columns: {a.shape[1]}",
    ]
    kpis = [["Jobs (filtered)", f"{len(view)}"]]
    if meta["usable"].get("revenue"):
        kpis.append(["Revenue", f"${view['_num_revenue'].sum():,.0f}"])
    if meta["usable"].get("profit"):
        kpis.append(["Profit", f"${view['_num_profit'].sum():,.0f}"])

    warn = []
    if meta["usable"].get("start_date") and meta["usable"].get("end_date"):
        n = int((a["_parsed_start_date"].notna() & a["_parsed_end_date"].notna()
                 & (a["_parsed_end_date"] < a["_parsed_start_date"])).sum())
        warn.append(f"End-before-start: {n} row(s).")
    if m.get("job_id") and m["job_id"] in a.columns:
        n = int((a[m["job_id"]].isna() |
                 (a[m["job_id"]].astype(str).str.strip() == "")).sum())
        warn.append(f"Missing JobID: {n} row(s).")
    warnings_text = "<br/>".join(warn) if warn else "No warnings with current mapping."

    if st.button("Generate PDF report"):
        imgs = make_chart_pngs(view, meta, m)
        pdf = build_pdf(summary, kpis, imgs, warnings_text, recs)
        st.download_button("⬇️ Download PDF report", data=pdf,
                           file_name="earthbender_report.pdf", mime="application/pdf")

    st.subheader("Cleaned analysis table")
    csv = a.to_csv(index=False).encode("utf-8")
    st.download_button("⬇️ Download analysis CSV", data=csv,
                       file_name="earthbender_analysis.csv", mime="text/csv")
    st.dataframe(a.head(20), use_container_width=True)

# ---------------------------------------------------------------------------
def main():
    init_state()
    st.title(APP_TITLE); st.caption(APP_SUB)

    section_import()
    df = st.session_state.get("df")
    if df is None:
        st.info("Upload a workbook to begin."); st.stop()

    m = mapping_ui(df)
    a, meta = build_analysis(df, m)
    st.session_state["analysis"] = a
    st.session_state["meta"] = meta

    st.divider(); section_dashboard(a, meta, m)
    st.divider(); section_errors(a, meta, m)
    st.divider()
    view = st.session_state.get("dash_view", a)
    section_recommendations(a, view, meta, m)
    st.divider(); section_export(a, view, meta, m)

if __name__ == "__main__":
    main()
