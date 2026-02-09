import pandas as pd
import streamlit as st
from io import BytesIO
import plotly.graph_objects as go
from datetime import datetime

st.set_page_config(layout="wide", page_title="Lean Manufacturing AI Dashboard")
st.title("🏭 Lean Manufacturing AI Dashboard")

# ------------------------
# Refresh Excel Data
# ------------------------
if st.button("🔄 Refresh Excel Data"):
    st.cache_data.clear()
    st.success("Data refreshed!")

# ------------------------
# Load Data from Excel
# ------------------------
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

# ------------------------
# Item Selection
# ------------------------
item_list = sorted(immaster["item"].dropna().unique())
item_input = st.selectbox("🔎 Select Item", ["--Select--"] + item_list)

if item_input != "--Select--":

    item_info = immaster[immaster["item"] == item_input].iloc[0]
    route_data = apnrn[apnrn["partno"] == item_input]

    if route_data.empty:
        st.warning("No route assigned.")
        st.stop()

    route_no = route_data.iloc[0]["routeno"]

    st.subheader(f"Item: {item_input}")
    st.write(f"**Description:** {item_info.get('descrip','N/A')}")
    st.write(f"**MLO Status (misc05):** {item_info.get('misc05','')}")
    st.write(f"**LAB-OVH Status (misc10):** {item_info.get('misc10','')}")

    # ------------------------
    # Route Operations
    # ------------------------
    route_ops = rodetail[rodetail["routeno"] == route_no].copy()
    route_ops["cycletime"] = pd.to_numeric(route_ops["cycletime"], errors="coerce").fillna(0)
    route_ops = route_ops.sort_values("opno").reset_index(drop=True)
    route_ops["station"] = [(i + 1) * 10 for i in range(len(route_ops))]

    bottleneck_time = route_ops["cycletime"].max()
    route_ops["bottleneck"] = route_ops["cycletime"] == bottleneck_time
    total_cycle_time = route_ops["cycletime"].sum()

    # ------------------------
    # Production Metrics
    # ------------------------
    available_time = 8 * 60 * 60
    customer_demand = st.number_input("Customer Demand per Shift", value=400)
    takt_time = available_time / customer_demand if customer_demand else 0
    throughput = 3600 / bottleneck_time if bottleneck_time else 0

    st.subheader("📊 Production Metrics")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Bottleneck Time", f"{bottleneck_time:.1f} sec")
    c2.metric("Throughput", f"{throughput:.2f} units/hr")
    c3.metric("Takt Time", f"{takt_time:.1f} sec")
    c4.metric("Lead Time", f"{total_cycle_time:.1f} sec")

    # ------------------------
    # Extra Operators
    # ------------------------
    st.subheader("🔧 Adjust Operators at Bottleneck")
    extra_ops = st.number_input("Enter number of extra operators at the bottleneck", min_value=0, value=0)

    # ------------------------
    # Improved Cycle Time
    # ------------------------
    improved_cycle = route_ops["cycletime"].copy()
    if extra_ops > 0:
        improved_cycle[route_ops["bottleneck"]] = bottleneck_time / (extra_ops + 1)

    new_bottleneck_time = improved_cycle.max()
    new_throughput = 3600 / new_bottleneck_time if new_bottleneck_time else 0
    new_lead_time = improved_cycle.sum()

    st.subheader("📊 Updated Metrics with Extra Operators")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("New Bottleneck Time", f"{new_bottleneck_time:.1f} sec")
    c2.metric("New Throughput", f"{new_throughput:.2f} units/hr")
    c3.metric("Takt Time", f"{takt_time:.1f} sec")
    c4.metric("New Lead Time", f"{new_lead_time:.1f} sec")

    # ------------------------
    # Bottleneck Chart
    # ------------------------
    st.subheader("🔴 Bottleneck Improvement Simulation")

    colors = ["red" if x else "steelblue" for x in route_ops["bottleneck"]]

    fig = go.Figure()
    fig.add_bar(x=route_ops["station"], y=route_ops["cycletime"], marker_color=colors, name="Current Cycle Time")
    fig.add_scatter(x=route_ops["station"], y=improved_cycle, mode="lines+markers",
                    name="With Extra Operator", line=dict(color="green", dash="dash"))

    fig.update_layout(xaxis_title="Station", yaxis_title="Cycle Time (sec)", height=450)
    st.plotly_chart(fig, use_container_width=True)

    # ------------------------
    # Bottleneck Deep Dive
    # ------------------------
    st.subheader("🔍 Bottleneck Operation Details")
    bottleneck_row = route_ops.loc[route_ops["cycletime"].idxmax()]
    setup_col = next((col for col in route_ops.columns if "setup" in col), None)
    setup_time_value = bottleneck_row[setup_col] if setup_col else "N/A"

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Operation No", int(bottleneck_row["opno"]))
    col2.metric("Cycle Time (sec)", f"{bottleneck_row['cycletime']:.1f}")
    col3.metric("Setup Time", setup_time_value)
    col4.metric("Labor Grade", bottleneck_row.get("laborgrade", "N/A"))
    st.write(f"**Description:** {bottleneck_row.get('descrip', 'N/A')}")

    # ------------------------
    # Throughput per Operation
    # ------------------------
    route_ops["throughput_per_op"] = route_ops["cycletime"].apply(lambda x: 3600/x if x > 0 else 0)
    improved_df = route_ops.copy()
    improved_df["cycletime"] = improved_cycle
    improved_df["throughput_per_op"] = improved_df["cycletime"].apply(lambda x: 3600/x if x > 0 else 0)

    route_ops["bottleneck_flag"] = route_ops["bottleneck"].apply(lambda x: "🔴" if x else "")
    improved_df["bottleneck_flag"] = improved_df["cycletime"].apply(lambda x: "🔴" if x == new_bottleneck_time else "")

    # ------------------------
    # Tables
    # ------------------------
    st.subheader("📈 Value Stream Map Data")
    st.write("**Current Cycle Time**")
    st.dataframe(route_ops[["station", "opno", "descrip", "cycletime", "laborgrade", "throughput_per_op", "bottleneck_flag"]])

    st.write("**Improved Cycle Time**")
    st.dataframe(improved_df[["station", "opno", "descrip", "cycletime", "laborgrade", "throughput_per_op", "bottleneck_flag"]])

    # ------------------------
    # 💡 Lean & Six Sigma Recommendations
    # ------------------------
    st.subheader("💡 Lean & Six Sigma Recommendations")
    recommendations = []

    if bottleneck_time > takt_time:
        recommendations.append("🔴 Bottleneck exceeds Takt → Line Balancing + Yamazumi.")
    if new_bottleneck_time < bottleneck_time:
        recommendations.append("📈 Improvement validated → Update Standard Work.")
    if route_ops["cycletime"].std() > 0.4 * route_ops["cycletime"].mean():
        recommendations.append("📊 High variation → Run Six Sigma DMAIC.")
    if setup_col and route_ops[setup_col].max() > 0.3 * total_cycle_time:
        recommendations.append("⚙️ High setup → Apply SMED.")
    if len(route_ops) > 8:
        recommendations.append("📋 Many operations → Use 5S + Cellular Layout.")
    if throughput < (customer_demand / 8):
        recommendations.append("📉 Low throughput → Kaizen waste elimination.")
    if not recommendations:
        recommendations.append("✅ Line balanced → Maintain with SPC.")

    for rec in recommendations:
        st.write(rec)

    # ------------------------
    # Download Report
    # ------------------------
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        route_ops.to_excel(writer, sheet_name='CurrentCycle', index=False)
        improved_df.to_excel(writer, sheet_name='ImprovedCycle', index=False)

    now = datetime.now().strftime("%Y%m%d_%H%M")
    st.download_button("⬇ Download Route & Improvement Report",
                       data=output.getvalue(),
                       file_name=f"{item_input}_Route_Report_{now}.xlsx")

else:
    st.info("Please select an item.")
