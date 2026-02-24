import pandas as pd
import streamlit as st
from io import BytesIO
import plotly.graph_objects as go
from datetime import datetime

# ================= PAGE CONFIG =================
st.set_page_config(page_title="Lean Line Optimizer", layout="wide")
st.title("🏭 Lean Manufacturing Line Optimizer")
st.caption("Operator Balancing • Kaizen Simulation • Bottleneck Analysis • Order Savings")

# ================= SIDEBAR =================
st.sidebar.header("⚙️ Control Panel")
if st.sidebar.button("🔄 Refresh Data"):
    st.cache_data.clear()
    st.success("Data refreshed!")

# ================= LOAD EXCEL DATA =================
@st.cache_data
def load_data():
    xls = pd.ExcelFile("Route File Lean.xlsx")
    cograde = pd.read_excel(xls, "cograde")
    rodetail = pd.read_excel(xls, "rodetail")
    apnrn = pd.read_excel(xls, "apnrn")
    immaster = pd.read_excel(xls, "immaster")
    for df in [cograde, rodetail, apnrn, immaster]:
        df.columns = df.columns.str.strip().str.lower()
    return cograde, rodetail, apnrn, immaster

cograde, rodetail, apnrn, immaster = load_data()

# ================= SIDEBAR INPUTS =================
item_list = sorted(immaster["item"].dropna().unique())
item_input = st.sidebar.selectbox("📦 Select Item", ["--Select--"] + item_list)
customer_demand = st.sidebar.number_input("📈 Demand per Shift", value=400)
order_qty = st.sidebar.number_input("📦 Order Quantity", min_value=1, value=500)
shift_hours = st.sidebar.number_input(
    "⏱ Shift Duration (hours)", min_value=1.0, value=8.0, step=0.5
)
shift_sec = shift_hours * 3600

# ================= MAIN =================
if item_input != "--Select--":

    # Fetch item route
    route_row = apnrn[apnrn["partno"] == item_input]
    if route_row.empty:
        st.error("No route found for this item.")
        st.stop()

    route_no = route_row.iloc[0]["routeno"]
    route_ops = rodetail[rodetail["routeno"] == route_no].copy()
    route_ops["cycletime"] = pd.to_numeric(route_ops["cycletime"], errors="coerce").fillna(0)
    route_ops = route_ops.sort_values("opno").reset_index(drop=True)
    route_ops["station"] = [(i + 1) * 10 for i in range(len(route_ops))]

    # ===== Merge hour rate per operation =====
    route_ops["laborgrade"] = route_ops["laborgrade"].astype(str).str.strip().str.lower()
    cograde["grade"] = cograde["grade"].astype(str).str.strip().str.lower()
    route_ops = route_ops.merge(
        cograde[["grade", "hourrate"]],
        left_on="laborgrade",
        right_on="grade",
        how="left"
    )
    route_ops["hourrate"] = route_ops["hourrate"].fillna(0)
    route_ops.drop(columns=["grade"], inplace=True)

    # Current metrics
    bottleneck_time = route_ops["cycletime"].max()
    route_ops["bottleneck"] = route_ops["cycletime"] == bottleneck_time
    takt_time = shift_sec / customer_demand if customer_demand else 0
    throughput = 3600 / bottleneck_time if bottleneck_time else 0

    # ================= EXECUTIVE SUMMARY =================
    st.subheader("📊 Executive Summary")
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Current Bottleneck (sec)", f"{bottleneck_time:.1f}")
    k2.metric("Throughput (units/hr)", f"{throughput:.2f}")
    k3.metric("Takt Time (sec)", f"{takt_time:.1f}")
    k4.metric("Order Quantity", f"{order_qty}")
    st.divider()

    # ================= IMPROVEMENT INPUTS =================
    left, right = st.columns([1.2, 1])
    with left:
        st.subheader("🛠 Improvement Inputs")
        input_df = route_ops[["opno", "descrip", "cycletime", "hourrate", "station"]].copy()
        input_df["Extra Operators"] = 0
        input_df["Time Saving (sec)"] = 0.0
        input_df["Improvement %"] = 0.0

        edited_df = st.data_editor(input_df, use_container_width=True, height=400)
        edited_df = edited_df.copy()

    # ================= CALCULATIONS =================
    improved_cycle = []
    reduction = []
    savings_per_unit = []
    savings_order = []

    for i, row in edited_df.iterrows():
        original = row["cycletime"]
        operators = row["Extra Operators"]
        saving_sec = row["Time Saving (sec)"]
        improvement_pct = row["Improvement %"]
        rate = row["hourrate"]

        new_time = original / (operators + 1)
        new_time = max(new_time - saving_sec, 0)
        if improvement_pct > 0:
            new_time = new_time * (1 - improvement_pct / 100)
        new_time = max(new_time, 0)
        improved_cycle.append(new_time)

        red = original - new_time
        reduction.append(red)

        save_unit = (red / 3600) * rate
        savings_per_unit.append(save_unit)
        savings_order.append(save_unit * order_qty)

    improved_cycle = pd.Series(improved_cycle)
    reduction = pd.Series(reduction)
    savings_per_unit = pd.Series(savings_per_unit)
    savings_order = pd.Series(savings_order)

    new_bottleneck_time = improved_cycle.max()
    new_bottleneck_flag = improved_cycle == new_bottleneck_time
    new_throughput = 3600 / new_bottleneck_time if new_bottleneck_time else 0

    # ================= RESULTS =================
    with right:
        st.subheader("📊 Improvement Results")
        r1, r2 = st.columns(2)
        r1.metric("New Bottleneck (sec)", f"{new_bottleneck_time:.1f}")
        r2.metric("New Throughput (units/hr)", f"{new_throughput:.2f}")

        st.markdown("### 🧠 Insights")
        old_bn = route_ops.loc[route_ops["bottleneck"], "opno"].values[0]
        new_bn = route_ops.loc[new_bottleneck_flag, "opno"].values[0]
        if old_bn != new_bn:
            st.success(f"Bottleneck moved: Op {old_bn} → Op {new_bn}")
        else:
            st.warning("Bottleneck unchanged")

        if new_bottleneck_time > takt_time:
            st.error("⚠ Cannot meet customer demand")
        else:
            st.success("✅ Meets customer demand")

        total_time_saved = reduction.sum()
        total_savings_order = sum(savings_order)
        st.info(f"Total Time Saved: {total_time_saved:.1f} sec")
        st.info(f"Estimated Total Savings for Order: ${total_savings_order:,.2f}")

    st.divider()

    # ================= DYNAMIC YAMAZUMI CHART =================
    st.subheader("📊 Yamazumi Chart")
    fig = go.Figure()

    for i, row in edited_df.iterrows():
        work_color = "red" if improved_cycle.iloc[i] > takt_time else ("orange" if new_bottleneck_flag.iloc[i] else "steelblue")
        save_color = "green" if reduction.iloc[i] > 0 else "rgba(0,0,0,0)"

        hover_text = (
            f"Op: {row['opno']}<br>"
            f"Description: {row['descrip']}<br>"
            f"Original Cycle: {row['cycletime']:.1f} sec<br>"
            f"Hour Rate: ${row['hourrate']:.2f}/hr<br>"
            f"Improved Cycle: {improved_cycle.iloc[i]:.1f} sec<br>"
            f"Time Saved: {reduction.iloc[i]:.1f} sec<br>"
            f"Improvement %: {edited_df['Improvement %'].iloc[i]:.1f}%<br>"
            f"Savings $ per Unit: ${savings_per_unit.iloc[i]:.2f}<br>"
            f"Savings $ Order: ${savings_order.iloc[i]:.2f}"
        )

        fig.add_bar(
            x=[row["station"]],
            y=[improved_cycle.iloc[i]],
            name="Work Content",
            marker_color=work_color,
            showlegend=False,
            hovertemplate=hover_text
        )
        fig.add_bar(
            x=[row["station"]],
            y=[reduction.iloc[i]],
            name="Time Saved",
            marker_color=save_color,
            showlegend=False,
            hovertemplate=hover_text
        )

    takt_line_color = "red" if new_bottleneck_time > takt_time else "black"
    fig.add_hline(
        y=takt_time,
        line_dash="dash",
        line_color=takt_line_color,
        annotation_text="Takt",
        annotation_font_color=takt_line_color
    )

    fig.update_layout(
        barmode='stack',
        height=500,
        xaxis_title="Station",
        yaxis_title="Cycle Time (sec)"
    )
    st.plotly_chart(fig, use_container_width=True)

    st.divider()

    # ================= LEAN & SIX SIGMA RECOMMENDATIONS =================
    st.subheader("📌 Lean & Six Sigma Recommendations")
    recommendations = []

    if new_bottleneck_time > takt_time:
        recommendations.append("🔴 Bottleneck exceeds Takt → Line Balancing + Yamazumi.")
    if total_time_saved > 0:
        recommendations.append("📈 Improvement validated → Update Standard Work.")
    if route_ops["cycletime"].std() / route_ops["cycletime"].mean() > 0.2:  # example high variation threshold
        recommendations.append("📊 High variation → Run Six Sigma DMAIC.")
    if route_ops.get("setup_time") is not None and route_ops["setup_time"].max() > 300:  # high setup example
        recommendations.append("⚙️ High setup → Apply SMED.")
    if new_throughput < customer_demand:
        recommendations.append("📉 Low throughput → Kaizen waste elimination.")

    if recommendations:
        for rec in recommendations:
            st.info(rec)
    else:
        st.success("✅ Process is balanced. No immediate Lean/Six Sigma actions required.")

    # ================= DETAILED TABLE =================
    st.subheader("📋 Detailed Data")
    route_ops["improved_cycle"] = improved_cycle
    route_ops["saving_sec"] = reduction
    route_ops["improvement_%"] = edited_df["Improvement %"]
    route_ops["new_bottleneck"] = new_bottleneck_flag
    route_ops["savings_$ per unit"] = savings_per_unit
    route_ops["savings_$ order"] = savings_order

    detailed_cols = [
        "opno", "descrip", "cycletime", "hourrate", "station",
        "improved_cycle", "saving_sec", "improvement_%",
        "new_bottleneck", "savings_$ per unit", "savings_$ order"
    ]
    st.dataframe(route_ops[detailed_cols], use_container_width=True)

    # ================= EXPORT =================
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        route_ops[detailed_cols].to_excel(writer, index=False)
    now = datetime.now().strftime("%Y%m%d_%H%M")
    st.download_button(
        "⬇ Download Report",
        data=output.getvalue(),
        file_name=f"{item_input}_Lean_Report_{now}.xlsx"
    )

else:
    st.info("👈 Select an item from the sidebar to begin")