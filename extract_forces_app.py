"""
Streamlit app — Element Forces Analyser
Extracts min/max F11 (+ M22) and min/max F22 (+ M11) from an Excel shell-forces report.
Also analyses M11/M22 extremes, identifies their Area Shells, and extracts F22/F11 from those areas.

Run with:
    streamlit run extract_forces_app.py
"""

import openpyxl
import pandas as pd
import plotly.graph_objects as go
import streamlit as st


# ─────────────────────────────────────────────
# Core extraction logic
# ─────────────────────────────────────────────

def extract_all(filepath) -> tuple[dict, dict, pd.DataFrame]:
    wb = openpyxl.load_workbook(filepath)
    ws = wb["Element Forces - Area Shells"]

    # Locate column indices from header row (row 2)
    headers = [cell.value for cell in ws[2]]
    col = {name: headers.index(name) for name in ("Area", "F11", "F22", "M11", "M22")}

    # --- Trackers for F11/F22 extremes (original analysis) ---
    res = {
        "F11_max": {"F11": None, "M22": None, "row": None},
        "F11_min": {"F11": None, "M22": None, "row": None},
        "F22_max": {"F22": None, "M11": None, "row": None},
        "F22_min": {"F22": None, "M11": None, "row": None},
    }

    # --- Trackers for M11/M22 extremes (new analysis) ---
    m_res = {
        "M11_max": {"M11": None, "Area": None, "row": None},
        "M11_min": {"M11": None, "Area": None, "row": None},
        "M22_max": {"M22": None, "Area": None, "row": None},
        "M22_min": {"M22": None, "Area": None, "row": None},
    }

    rows = []
    for row_idx, row in enumerate(ws.iter_rows(min_row=4, values_only=True), start=4):
        area = row[col["Area"]]
        f11  = row[col["F11"]]
        f22  = row[col["F22"]]
        m11  = row[col["M11"]]
        m22  = row[col["M22"]]

        if not all(isinstance(v, (int, float)) for v in (f11, f22, m11, m22)):
            continue

        rows.append({"row": row_idx, "Area": area, "F11": f11, "F22": f22, "M11": m11, "M22": m22})

        # F11 extremes → paired M22
        if res["F11_max"]["F11"] is None or f11 > res["F11_max"]["F11"]:
            res["F11_max"] = {"F11": f11, "M22": m22, "row": row_idx}
        if res["F11_min"]["F11"] is None or f11 < res["F11_min"]["F11"]:
            res["F11_min"] = {"F11": f11, "M22": m22, "row": row_idx}

        # F22 extremes → paired M11
        if res["F22_max"]["F22"] is None or f22 > res["F22_max"]["F22"]:
            res["F22_max"] = {"F22": f22, "M11": m11, "row": row_idx}
        if res["F22_min"]["F22"] is None or f22 < res["F22_min"]["F22"]:
            res["F22_min"] = {"F22": f22, "M11": m11, "row": row_idx}

        # M11 extremes → paired Area
        if m_res["M11_max"]["M11"] is None or m11 > m_res["M11_max"]["M11"]:
            m_res["M11_max"] = {"M11": m11, "Area": area, "row": row_idx}
        if m_res["M11_min"]["M11"] is None or m11 < m_res["M11_min"]["M11"]:
            m_res["M11_min"] = {"M11": m11, "Area": area, "row": row_idx}

        # M22 extremes → paired Area
        if m_res["M22_max"]["M22"] is None or m22 > m_res["M22_max"]["M22"]:
            m_res["M22_max"] = {"M22": m22, "Area": area, "row": row_idx}
        if m_res["M22_min"]["M22"] is None or m22 < m_res["M22_min"]["M22"]:
            m_res["M22_min"] = {"M22": m22, "Area": area, "row": row_idx}

    df = pd.DataFrame(rows)

    # --- Derive F22 from M11 areas, and F11 from M22 areas ---
    def f22_for_area(area):
        sub = df[df["Area"] == area]["F22"]
        return {"max": sub.max(), "min": sub.min(), "area": area}

    def f11_for_area(area):
        sub = df[df["Area"] == area]["F11"]
        return {"max": sub.max(), "min": sub.min(), "area": area}

    area_res = {
        "M11_max_area_F22": f22_for_area(m_res["M11_max"]["Area"]),
        "M11_min_area_F22": f22_for_area(m_res["M11_min"]["Area"]),
        "M22_max_area_F11": f11_for_area(m_res["M22_max"]["Area"]),
        "M22_min_area_F11": f11_for_area(m_res["M22_min"]["Area"]),
    }

    return res, {"m_res": m_res, "area_res": area_res}, df


# ─────────────────────────────────────────────
# Chart helpers
# ─────────────────────────────────────────────

def bar_chart(labels, values, colors, title, unit):
    fig = go.Figure(go.Bar(
        x=labels, y=values,
        marker_color=colors,
        text=[f"{v:.4f}" for v in values],
        textposition="outside",
    ))
    fig.update_layout(
        title=title, yaxis_title=unit,
        plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
        font=dict(color="#e0e0e0"), title_font=dict(size=16),
        margin=dict(t=50, b=40),
    )
    fig.update_yaxes(gridcolor="rgba(255,255,255,0.1)")
    return fig


def scatter_chart(df, x_col, y_col, x_label, y_label, highlights):
    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=df[x_col], y=df[y_col],
        mode="markers",
        marker=dict(size=5, color="steelblue", opacity=0.6),
        name="All rows",
    ))
    for h in highlights:
        fig.add_trace(go.Scatter(
            x=[h["x"]], y=[h["y"]],
            mode="markers+text",
            marker=dict(size=12, color=h["color"], symbol="star"),
            text=[h["label"]], textposition="top center",
            name=h["label"],
        ))
    fig.update_layout(
        xaxis_title=x_label, yaxis_title=y_label,
        plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
        font=dict(color="#e0e0e0"), margin=dict(t=30, b=40),
        legend=dict(bgcolor="rgba(0,0,0,0)"),
    )
    fig.update_xaxes(gridcolor="rgba(255,255,255,0.1)")
    fig.update_yaxes(gridcolor="rgba(255,255,255,0.1)")
    return fig


def grouped_bar(title, groups, unit):
    """groups = [{"label": str, "max": float, "min": float, "color_max": str, "color_min": str}]"""
    fig = go.Figure()
    fig.add_trace(go.Bar(
        name="Max",
        x=[g["label"] for g in groups],
        y=[g["max"] for g in groups],
        marker_color=[g["color_max"] for g in groups],
        text=[f"{g['max']:.4f}" for g in groups],
        textposition="outside",
    ))
    fig.add_trace(go.Bar(
        name="Min",
        x=[g["label"] for g in groups],
        y=[g["min"] for g in groups],
        marker_color=[g["color_min"] for g in groups],
        text=[f"{g['min']:.4f}" for g in groups],
        textposition="outside",
    ))
    fig.update_layout(
        barmode="group", title=title, yaxis_title=unit,
        plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
        font=dict(color="#e0e0e0"), title_font=dict(size=16),
        margin=dict(t=50, b=40),
        legend=dict(bgcolor="rgba(0,0,0,0)"),
    )
    fig.update_yaxes(gridcolor="rgba(255,255,255,0.1)")
    return fig


# ─────────────────────────────────────────────
# Streamlit UI
# ─────────────────────────────────────────────

st.set_page_config(page_title="Element Forces Analyser", page_icon="🏗️", layout="wide")

st.title("🏗️ Element Forces Analyser")
st.markdown("Upload your **Element Forces – Area Shells** Excel file to explore extreme force and moment values.")

uploaded = st.file_uploader("Upload Excel file (.xlsx)", type=["xlsx"])

if uploaded is None:
    st.info("👆 Upload an `.xlsx` file to get started.")
    st.stop()

with st.spinner("Reading and analysing…"):
    res, new_res, df = extract_all(uploaded)

m_res    = new_res["m_res"]
area_res = new_res["area_res"]

st.success(f"✅ Analysed **{len(df):,}** valid data rows.")

# ════════════════════════════════════════════
# SECTION 1 — F11 / F22 Original Analysis
# ════════════════════════════════════════════
st.markdown("---")
st.header("📐 Section 1 — F11 & F22 Extremes")

c1, c2, c3, c4 = st.columns(4)
with c1:
    st.metric("Max F11 (KN/m)", f"{res['F11_max']['F11']:.4f}",
              f"M22 = {res['F11_max']['M22']:.4f} KN-m/m")
with c2:
    st.metric("Min F11 (KN/m)", f"{res['F11_min']['F11']:.4f}",
              f"M22 = {res['F11_min']['M22']:.4f} KN-m/m", delta_color="inverse")
with c3:
    st.metric("Max F22 (KN/m)", f"{res['F22_max']['F22']:.4f}",
              f"M11 = {res['F22_max']['M11']:.4f} KN-m/m")
with c4:
    st.metric("Min F22 (KN/m)", f"{res['F22_min']['F22']:.4f}",
              f"M11 = {res['F22_min']['M11']:.4f} KN-m/m", delta_color="inverse")

ca, cb = st.columns(2)
with ca:
    st.subheader("F11 → M22 Detail")
    st.dataframe(pd.DataFrame([
        {"Extreme": "Max F11", "F11 (KN/m)": res["F11_max"]["F11"],
         "M22 (KN-m/m)": res["F11_max"]["M22"], "Excel Row": res["F11_max"]["row"]},
        {"Extreme": "Min F11", "F11 (KN/m)": res["F11_min"]["F11"],
         "M22 (KN-m/m)": res["F11_min"]["M22"], "Excel Row": res["F11_min"]["row"]},
    ]), use_container_width=True, hide_index=True)

with cb:
    st.subheader("F22 → M11 Detail")
    st.dataframe(pd.DataFrame([
        {"Extreme": "Max F22", "F22 (KN/m)": res["F22_max"]["F22"],
         "M11 (KN-m/m)": res["F22_max"]["M11"], "Excel Row": res["F22_max"]["row"]},
        {"Extreme": "Min F22", "F22 (KN/m)": res["F22_min"]["F22"],
         "M11 (KN-m/m)": res["F22_min"]["M11"], "Excel Row": res["F22_min"]["row"]},
    ]), use_container_width=True, hide_index=True)

cc, cd = st.columns(2)
with cc:
    st.plotly_chart(bar_chart(
        ["Max F11", "Min F11"],
        [res["F11_max"]["F11"], res["F11_min"]["F11"]],
        ["#2ecc71", "#e74c3c"], "F11 Extremes", "KN/m"
    ), use_container_width=True)
with cd:
    st.plotly_chart(bar_chart(
        ["Max F22", "Min F22"],
        [res["F22_max"]["F22"], res["F22_min"]["F22"]],
        ["#3498db", "#e67e22"], "F22 Extremes", "KN/m"
    ), use_container_width=True)

ce, cf = st.columns(2)
with ce:
    st.plotly_chart(scatter_chart(
        df, "F11", "M22", "F11 (KN/m)", "M22 (KN-m/m)",
        highlights=[
            {"x": res["F11_max"]["F11"], "y": res["F11_max"]["M22"], "label": "Max F11", "color": "#2ecc71"},
            {"x": res["F11_min"]["F11"], "y": res["F11_min"]["M22"], "label": "Min F11", "color": "#e74c3c"},
        ]
    ), use_container_width=True)
with cf:
    st.plotly_chart(scatter_chart(
        df, "F22", "M11", "F22 (KN/m)", "M11 (KN-m/m)",
        highlights=[
            {"x": res["F22_max"]["F22"], "y": res["F22_max"]["M11"], "label": "Max F22", "color": "#3498db"},
            {"x": res["F22_min"]["F22"], "y": res["F22_min"]["M11"], "label": "Min F22", "color": "#e67e22"},
        ]
    ), use_container_width=True)


# ════════════════════════════════════════════
# SECTION 2 — M11 Extremes → Area → F22
# ════════════════════════════════════════════
st.markdown("---")
st.header("🔎 Section 2 — M11 Extremes → Identified Area → F22 Range")
st.markdown(
    "The **min and max M11** values are located, their **Area Shell** is identified, "
    "then the **min and max F22** within those areas are extracted."
)

m11_max_area = m_res["M11_max"]["Area"]
m11_min_area = m_res["M11_min"]["Area"]

s2c1, s2c2 = st.columns(2)

with s2c1:
    st.subheader(f"Max M11 — Area **{m11_max_area}**")
    st.metric("Max M11 (KN-m/m)", f"{m_res['M11_max']['M11']:.4f}",
              f"Area Shell: {m11_max_area}")
    ar = area_res["M11_max_area_F22"]
    st.dataframe(pd.DataFrame([
        {"Metric": "Max M11",        "Value (KN-m/m)": m_res["M11_max"]["M11"], "Excel Row": m_res["M11_max"]["row"]},
        {"Metric": "Identified Area","Value": m11_max_area},
        {"Metric": "Max F22 in Area","Value (KN/m)": ar["max"]},
        {"Metric": "Min F22 in Area","Value (KN/m)": ar["min"]},
    ]), use_container_width=True, hide_index=True)
    st.plotly_chart(bar_chart(
        [f"Max F22\n(Area {m11_max_area})", f"Min F22\n(Area {m11_max_area})"],
        [ar["max"], ar["min"]],
        ["#2ecc71", "#e74c3c"],
        f"F22 in Area {m11_max_area} (from Max M11)", "KN/m"
    ), use_container_width=True)

with s2c2:
    st.subheader(f"Min M11 — Area **{m11_min_area}**")
    st.metric("Min M11 (KN-m/m)", f"{m_res['M11_min']['M11']:.4f}",
              f"Area Shell: {m11_min_area}", delta_color="inverse")
    ar = area_res["M11_min_area_F22"]
    st.dataframe(pd.DataFrame([
        {"Metric": "Min M11",        "Value (KN-m/m)": m_res["M11_min"]["M11"], "Excel Row": m_res["M11_min"]["row"]},
        {"Metric": "Identified Area","Value": m11_min_area},
        {"Metric": "Max F22 in Area","Value (KN/m)": ar["max"]},
        {"Metric": "Min F22 in Area","Value (KN/m)": ar["min"]},
    ]), use_container_width=True, hide_index=True)
    st.plotly_chart(bar_chart(
        [f"Max F22\n(Area {m11_min_area})", f"Min F22\n(Area {m11_min_area})"],
        [ar["max"], ar["min"]],
        ["#3498db", "#e67e22"],
        f"F22 in Area {m11_min_area} (from Min M11)", "KN/m"
    ), use_container_width=True)

# Scatter — M11 vs F22, highlight the two areas
st.subheader("M11 vs F22 — Highlighted Areas from Extremes")
fig_m11 = go.Figure()
for area, color, label in [
    (m11_max_area, "#2ecc71", f"Area {m11_max_area} (Max M11)"),
    (m11_min_area, "#e74c3c", f"Area {m11_min_area} (Min M11)"),
]:
    sub = df[df["Area"] == area]
    fig_m11.add_trace(go.Scatter(
        x=sub["F22"], y=sub["M11"], mode="markers",
        marker=dict(size=9, color=color), name=label,
    ))
other = df[~df["Area"].isin([m11_max_area, m11_min_area])]
fig_m11.add_trace(go.Scatter(
    x=other["F22"], y=other["M11"], mode="markers",
    marker=dict(size=4, color="steelblue", opacity=0.3), name="Other areas",
))
fig_m11.update_layout(
    xaxis_title="F22 (KN/m)", yaxis_title="M11 (KN-m/m)",
    plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
    font=dict(color="#e0e0e0"), legend=dict(bgcolor="rgba(0,0,0,0)"),
    margin=dict(t=30, b=40),
)
fig_m11.update_xaxes(gridcolor="rgba(255,255,255,0.1)")
fig_m11.update_yaxes(gridcolor="rgba(255,255,255,0.1)")
st.plotly_chart(fig_m11, use_container_width=True)


# ════════════════════════════════════════════
# SECTION 3 — M22 Extremes → Area → F11
# ════════════════════════════════════════════
st.markdown("---")
st.header("🔎 Section 3 — M22 Extremes → Identified Area → F11 Range")
st.markdown(
    "The **min and max M22** values are located, their **Area Shell** is identified, "
    "then the **min and max F11** within those areas are extracted."
)

m22_max_area = m_res["M22_max"]["Area"]
m22_min_area = m_res["M22_min"]["Area"]

s3c1, s3c2 = st.columns(2)

with s3c1:
    st.subheader(f"Max M22 — Area **{m22_max_area}**")
    st.metric("Max M22 (KN-m/m)", f"{m_res['M22_max']['M22']:.4f}",
              f"Area Shell: {m22_max_area}")
    ar = area_res["M22_max_area_F11"]
    st.dataframe(pd.DataFrame([
        {"Metric": "Max M22",        "Value (KN-m/m)": m_res["M22_max"]["M22"], "Excel Row": m_res["M22_max"]["row"]},
        {"Metric": "Identified Area","Value": m22_max_area},
        {"Metric": "Max F11 in Area","Value (KN/m)": ar["max"]},
        {"Metric": "Min F11 in Area","Value (KN/m)": ar["min"]},
    ]), use_container_width=True, hide_index=True)
    st.plotly_chart(bar_chart(
        [f"Max F11\n(Area {m22_max_area})", f"Min F11\n(Area {m22_max_area})"],
        [ar["max"], ar["min"]],
        ["#9b59b6", "#1abc9c"],
        f"F11 in Area {m22_max_area} (from Max M22)", "KN/m"
    ), use_container_width=True)

with s3c2:
    st.subheader(f"Min M22 — Area **{m22_min_area}**")
    st.metric("Min M22 (KN-m/m)", f"{m_res['M22_min']['M22']:.4f}",
              f"Area Shell: {m22_min_area}", delta_color="inverse")
    ar = area_res["M22_min_area_F11"]
    st.dataframe(pd.DataFrame([
        {"Metric": "Min M22",        "Value (KN-m/m)": m_res["M22_min"]["M22"], "Excel Row": m_res["M22_min"]["row"]},
        {"Metric": "Identified Area","Value": m22_min_area},
        {"Metric": "Max F11 in Area","Value (KN/m)": ar["max"]},
        {"Metric": "Min F11 in Area","Value (KN/m)": ar["min"]},
    ]), use_container_width=True, hide_index=True)
    st.plotly_chart(bar_chart(
        [f"Max F11\n(Area {m22_min_area})", f"Min F11\n(Area {m22_min_area})"],
        [ar["max"], ar["min"]],
        ["#e67e22", "#2980b9"],
        f"F11 in Area {m22_min_area} (from Min M22)", "KN/m"
    ), use_container_width=True)

# Scatter — M22 vs F11, highlight the two areas
st.subheader("M22 vs F11 — Highlighted Areas from Extremes")
fig_m22 = go.Figure()
for area, color, label in [
    (m22_max_area, "#9b59b6", f"Area {m22_max_area} (Max M22)"),
    (m22_min_area, "#1abc9c", f"Area {m22_min_area} (Min M22)"),
]:
    sub = df[df["Area"] == area]
    fig_m22.add_trace(go.Scatter(
        x=sub["F11"], y=sub["M22"], mode="markers",
        marker=dict(size=9, color=color), name=label,
    ))
other = df[~df["Area"].isin([m22_max_area, m22_min_area])]
fig_m22.add_trace(go.Scatter(
    x=other["F11"], y=other["M22"], mode="markers",
    marker=dict(size=4, color="steelblue", opacity=0.3), name="Other areas",
))
fig_m22.update_layout(
    xaxis_title="F11 (KN/m)", yaxis_title="M22 (KN-m/m)",
    plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
    font=dict(color="#e0e0e0"), legend=dict(bgcolor="rgba(0,0,0,0)"),
    margin=dict(t=30, b=40),
)
fig_m22.update_xaxes(gridcolor="rgba(255,255,255,0.1)")
fig_m22.update_yaxes(gridcolor="rgba(255,255,255,0.1)")
st.plotly_chart(fig_m22, use_container_width=True)


# ════════════════════════════════════════════
# Raw Data
# ════════════════════════════════════════════
st.markdown("---")
with st.expander("📋 View Raw Data Table"):
    st.dataframe(df, use_container_width=True, hide_index=True)
