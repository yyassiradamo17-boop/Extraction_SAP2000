"""
Streamlit app — Element Forces Analyser
Extracts min/max F11 (+ M22) and min/max F22 (+ M11) from an Excel shell-forces report.

Run with:
    streamlit run extract_forces_app.py
"""

import openpyxl
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

# ─────────────────────────────────────────────
# Core extraction logic (unchanged from CLI)
# ─────────────────────────────────────────────


def extract_force_extremes(filepath) -> tuple[dict, pd.DataFrame]:
    wb = openpyxl.load_workbook(filepath)
    ws = wb["Element Forces - Area Shells"]

    # Locate column indices from header row (row 2)
    headers = [cell.value for cell in ws[2]]
    col = {name: headers.index(name) for name in ("F11", "F22", "M11", "M22")}

    results = {
        "F11_max": {"F11": None, "M22": None, "row": None},
        "F11_min": {"F11": None, "M22": None, "row": None},
        "F22_max": {"F22": None, "M11": None, "row": None},
        "F22_min": {"F22": None, "M11": None, "row": None},
    }

    rows = []
    for row_idx, row in enumerate(ws.iter_rows(min_row=4, values_only=True), start=4):
        f11 = row[col["F11"]]
        f22 = row[col["F22"]]
        m11 = row[col["M11"]]
        m22 = row[col["M22"]]

        if not all(isinstance(v, (int, float)) for v in (f11, f22, m11, m22)):
            continue

        rows.append({"row": row_idx, "F11": f11,
                    "F22": f22, "M11": m11, "M22": m22})

        if results["F11_max"]["F11"] is None or f11 > results["F11_max"]["F11"]:
            results["F11_max"] = {"F11": f11, "M22": m22, "row": row_idx}
        if results["F11_min"]["F11"] is None or f11 < results["F11_min"]["F11"]:
            results["F11_min"] = {"F11": f11, "M22": m22, "row": row_idx}
        if results["F22_max"]["F22"] is None or f22 > results["F22_max"]["F22"]:
            results["F22_max"] = {"F22": f22, "M11": m11, "row": row_idx}
        if results["F22_min"]["F22"] is None or f22 < results["F22_min"]["F22"]:
            results["F22_min"] = {"F22": f22, "M11": m11, "row": row_idx}

    df = pd.DataFrame(rows)
    return results, df


# ─────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────

def metric_delta(value: float, ref: float = 0.0) -> str:
    """Return a signed delta string for st.metric."""
    diff = value - ref
    return f"{diff:+.4f}"


def bar_chart(labels, values, colors, title, unit):
    fig = go.Figure(go.Bar(
        x=labels,
        y=values,
        marker_color=colors,
        text=[f"{v:.4f}" for v in values],
        textposition="outside",
    ))
    fig.update_layout(
        title=title,
        yaxis_title=unit,
        plot_bgcolor="rgba(0,0,0,0)",
        paper_bgcolor="rgba(0,0,0,0)",
        font=dict(color="#e0e0e0"),
        title_font=dict(size=16),
        margin=dict(t=50, b=40),
    )
    fig.update_yaxes(gridcolor="rgba(255,255,255,0.1)")
    return fig


def scatter_chart(df, x_col, y_col, x_label, y_label, highlights: list[dict]):
    fig = go.Figure()

    # All data points
    fig.add_trace(go.Scatter(
        x=df[x_col], y=df[y_col],
        mode="markers",
        marker=dict(size=5, color="steelblue", opacity=0.6),
        name="All rows",
    ))

    # Highlight min/max points
    for h in highlights:
        fig.add_trace(go.Scatter(
            x=[h["x"]], y=[h["y"]],
            mode="markers+text",
            marker=dict(size=12, color=h["color"], symbol="star"),
            text=[h["label"]],
            textposition="top center",
            name=h["label"],
        ))

    fig.update_layout(
        xaxis_title=x_label,
        yaxis_title=y_label,
        plot_bgcolor="rgba(0,0,0,0)",
        paper_bgcolor="rgba(0,0,0,0)",
        font=dict(color="#e0e0e0"),
        margin=dict(t=30, b=40),
        legend=dict(bgcolor="rgba(0,0,0,0)"),
    )
    fig.update_xaxes(gridcolor="rgba(255,255,255,0.1)")
    fig.update_yaxes(gridcolor="rgba(255,255,255,0.1)")
    return fig


# ─────────────────────────────────────────────
# Streamlit UI
# ─────────────────────────────────────────────

st.set_page_config(
    page_title="Element Forces Analyser",
    page_icon="🏗️",
    layout="wide",
)

st.title("🏗️ Element Forces Analyser")
st.markdown(
    "Upload your **Element Forces – Area Shells** Excel file to explore extreme force and moment values.")

uploaded = st.file_uploader("Upload Excel file (.xlsx)", type=["xlsx"])

if uploaded is None:
    st.info("👆 Upload an `.xlsx` file to get started.")
    st.stop()

# ── Run extraction ──────────────────────────
with st.spinner("Reading and analysing…"):
    results, df = extract_force_extremes(uploaded)

st.success(f"✅ Analysed **{len(df):,}** valid data rows.")

# ── Summary metrics ─────────────────────────
st.markdown("---")
st.subheader("📊 Summary of Extremes")

col1, col2, col3, col4 = st.columns(4)

with col1:
    st.metric(
        label="Max F11 (KN/m)",
        value=f"{results['F11_max']['F11']:.4f}",
        delta=f"M22 = {results['F11_max']['M22']:.4f} KN-m/m",
    )
with col2:
    st.metric(
        label="Min F11 (KN/m)",
        value=f"{results['F11_min']['F11']:.4f}",
        delta=f"M22 = {results['F11_min']['M22']:.4f} KN-m/m",
        delta_color="inverse",
    )
with col3:
    st.metric(
        label="Max F22 (KN/m)",
        value=f"{results['F22_max']['F22']:.4f}",
        delta=f"M11 = {results['F22_max']['M11']:.4f} KN-m/m",
    )
with col4:
    st.metric(
        label="Min F22 (KN/m)",
        value=f"{results['F22_min']['F22']:.4f}",
        delta=f"M11 = {results['F22_min']['M11']:.4f} KN-m/m",
        delta_color="inverse",
    )

# ── Detail tables ────────────────────────────
st.markdown("---")
col_a, col_b = st.columns(2)

with col_a:
    st.subheader("F11 → M22")
    table_f11 = pd.DataFrame([
        {"Extreme": "Max F11", "F11 (KN/m)": results["F11_max"]["F11"],
         "M22 (KN-m/m)": results["F11_max"]["M22"], "Excel Row": results["F11_max"]["row"]},
        {"Extreme": "Min F11", "F11 (KN/m)": results["F11_min"]["F11"],
         "M22 (KN-m/m)": results["F11_min"]["M22"], "Excel Row": results["F11_min"]["row"]},
    ])
    st.dataframe(table_f11, use_container_width=True, hide_index=True)

with col_b:
    st.subheader("F22 → M11")
    table_f22 = pd.DataFrame([
        {"Extreme": "Max F22", "F22 (KN/m)": results["F22_max"]["F22"],
         "M11 (KN-m/m)": results["F22_max"]["M11"], "Excel Row": results["F22_max"]["row"]},
        {"Extreme": "Min F22", "F22 (KN/m)": results["F22_min"]["F22"],
         "M11 (KN-m/m)": results["F22_min"]["M11"], "Excel Row": results["F22_min"]["row"]},
    ])
    st.dataframe(table_f22, use_container_width=True, hide_index=True)

# ── Bar charts ───────────────────────────────
st.markdown("---")
st.subheader("📈 Bar Charts")

col_c, col_d = st.columns(2)

with col_c:
    fig_f11 = bar_chart(
        labels=["Max F11", "Min F11"],
        values=[results["F11_max"]["F11"], results["F11_min"]["F11"]],
        colors=["#2ecc71", "#e74c3c"],
        title="F11 Extremes",
        unit="KN/m",
    )
    st.plotly_chart(fig_f11, use_container_width=True)

with col_d:
    fig_f22 = bar_chart(
        labels=["Max F22", "Min F22"],
        values=[results["F22_max"]["F22"], results["F22_min"]["F22"]],
        colors=["#3498db", "#e67e22"],
        title="F22 Extremes",
        unit="KN/m",
    )
    st.plotly_chart(fig_f22, use_container_width=True)

# ── Scatter plots ────────────────────────────
st.markdown("---")
st.subheader("🔍 Scatter Plots — All Data with Highlighted Extremes")

col_e, col_f = st.columns(2)

with col_e:
    fig_s1 = scatter_chart(
        df, x_col="F11", y_col="M22",
        x_label="F11 (KN/m)", y_label="M22 (KN-m/m)",
        highlights=[
            {"x": results["F11_max"]["F11"], "y": results["F11_max"]["M22"],
             "label": "Max F11", "color": "#2ecc71"},
            {"x": results["F11_min"]["F11"], "y": results["F11_min"]["M22"],
             "label": "Min F11", "color": "#e74c3c"},
        ],
    )
    st.plotly_chart(fig_s1, use_container_width=True)

with col_f:
    fig_s2 = scatter_chart(
        df, x_col="F22", y_col="M11",
        x_label="F22 (KN/m)", y_label="M11 (KN-m/m)",
        highlights=[
            {"x": results["F22_max"]["F22"], "y": results["F22_max"]["M11"],
             "label": "Max F22", "color": "#3498db"},
            {"x": results["F22_min"]["F22"], "y": results["F22_min"]["M11"],
             "label": "Min F22", "color": "#e67e22"},
        ],
    )
    st.plotly_chart(fig_s2, use_container_width=True)

# ── Raw data table ───────────────────────────
st.markdown("---")
with st.expander("📋 View Raw Data Table"):
    st.dataframe(df, use_container_width=True, hide_index=True)
