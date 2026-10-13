"""
Nation Wide Wholesale - Automation Tool
Two tabs: Demand Forecast + PO Generator
"""
import streamlit as st
import tempfile
import os
import sys
import io
import zipfile
import shutil
import traceback
from pathlib import Path
from contextlib import redirect_stdout, redirect_stderr
from datetime import datetime, timezone, timedelta

st.set_page_config(page_title="Nation Wide Automation", page_icon="📦", layout="wide")

# ─────────────────────────── helpers ───────────────────────────

def save_upload(uploaded_file, directory, name=None):
    """Save an uploaded file to directory, return its path (or empty string if None)."""
    if uploaded_file is None:
        return ""
    fname = name or uploaded_file.name
    path = os.path.join(directory, fname)
    with open(path, "wb") as f:
        f.write(uploaded_file.getbuffer())
    return path


def zip_folder(folder_path):
    """Zip an entire folder and return bytes."""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for root, dirs, files in os.walk(folder_path):
            for file in files:
                full = os.path.join(root, file)
                arcname = os.path.relpath(full, folder_path)
                zf.write(full, arcname)
    buf.seek(0)
    return buf.read()


# ─────────────────────── available columns ─────────────────────
ALL_ORDERING_COLUMNS = [
    "Department", "Brand", "box upc", "master upc", "description",
    "PB description", "Pack of", "QTY", "Negative QTY flag",
    "Incoming QTY", "Suggested order QTY", "Adjustment", "Final order QTY",
    "Priority", "Planning notes", "ABC class", "Item class",
    "Auto order?", "Safety factor", "Reference order QTY",
    "Days cover", "Stock at arrival", "Daily demand", "WH trend",
    "Lead days", "Reorder below days", "Target cover days", "Order multiple",
    "Last buying price", "Last vendor", "Lowest buying price", "Lowest price vendor",
    "Status", "Current month avg/day",
]

DEFAULT_ORDERING_COLUMNS = [
    "Department", "Brand", "box upc", "master upc", "description", "PB description",
    "Pack of", "QTY", "Incoming QTY", "Suggested order QTY", "Adjustment",
    "Final order QTY", "Priority",
    "Last buying price", "Last vendor", "Lowest buying price", "Lowest price vendor",
]

ALL_PRIORITIES = [
    "Urgent - order", "Order", "Urgent - expedite incoming",
    "Incoming covers target", "Review - not auto",
    "No demand - review", "Review data", "Hold",
]

DEFAULT_PRIORITIES = ["Urgent - order", "Order"]

# ─────────────────────────── UI ───────────────────────────────

st.title("📦 Nation Wide Wholesale - Automation")
st.caption("Just upload your files,add the setting, press Run  — download the output!")

tab1, tab2 = st.tabs(["📊 Demand Forecast", "📋 PO Generator"])

# ═══════════════════════ TAB 1: DEMAND FORECAST ═══════════════════════
with tab1:
    st.header("Demand Forecast")
    st.markdown("Do Upload Inventory, purchases, sales and price book— script automatically makes the forecast.")

    st.subheader("📁 Upload files")
    col_left, col_right = st.columns(2)
    with col_left:
        st.markdown("**Required Files:**")
        f_inventory = st.file_uploader("Inventory", type=["xlsx"], key="df_inv")
        f_purchase = st.file_uploader("6 Month Purchase", type=["xlsx"], key="df_pur")
        f_sales = st.file_uploader("6 Month Sales", type=["xlsx"], key="df_sal")
        f_pricebook = st.file_uploader("Master Price Book", type=["xlsx"], key="df_pb")
        f_stores = st.file_uploader("6 Months PDI Sales (Stores)", type=["xlsx"], key="df_stores")

    with col_right:
        st.markdown("**Optional Files:**")
        f_pur_append = st.file_uploader("Current Purchase Warehouse (optional)", type=["xlsx"], key="df_pur_app")
        f_sal_append = st.file_uploader("Current Sales Warehouse (optional)", type=["xlsx"], key="df_sal_app")
        f_open_po = st.file_uploader("Open PO's (optional)", type=["xlsx"], key="df_po")
        f_planning = st.file_uploader("Planning Inputs (optional)", type=["xlsx"], key="df_plan")

    # ─────── Settings ───────
    st.subheader("⚙️ Settings")

    with st.expander("Demand Planning Settings", expanded=True):
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            lead_time = st.number_input("Lead Time (days)", value=3, min_value=0, key="df_lead")
        with c2:
            reorder_below = st.number_input("Reorder Below Days", value=10, min_value=1, key="df_reorder")
        with c3:
            target_days = st.number_input("Target Cover Days", value=20, min_value=1, key="df_target")
        with c4:
            window_days = st.number_input("Window Days", value=30, min_value=1, key="df_window")

        st.markdown("**Weights** (3 windows, newest to oldest — total = 1.0)")
        wc1, wc2, wc3 = st.columns(3)
        with wc1:
            w1 = st.number_input("Weight 1 (newest)", value=0.5, min_value=0.0, max_value=1.0, step=0.05, format="%.2f", key="df_w1")
        with wc2:
            w2 = st.number_input("Weight 2 (middle)", value=0.3, min_value=0.0, max_value=1.0, step=0.05, format="%.2f", key="df_w2")
        with wc3:
            w3 = st.number_input("Weight 3 (oldest)", value=0.2, min_value=0.0, max_value=1.0, step=0.05, format="%.2f", key="df_w3")

        weight_sum = round(w1 + w2 + w3, 4)
        if abs(weight_sum - 1.0) > 0.001:
            st.error(f"Weights ka total {weight_sum} hai — exactly 1.0 hona chahiye!")

    with st.expander("Ordering Columns & Priorities"):
        selected_columns = st.multiselect(
            "Ordering Sheet Columns (order matters — pehle select karo jo pehle chahiye)",
            options=ALL_ORDERING_COLUMNS,
            default=DEFAULT_ORDERING_COLUMNS,
            key="df_cols",
        )
        selected_priorities = st.multiselect(
            "Ordering Priorities (ye rows ordering sheet me aayengi)",
            options=ALL_PRIORITIES,
            default=DEFAULT_PRIORITIES,
            key="df_pri",
        )

    # ─────── Run ───────
    required_ok = all([f_inventory, f_purchase, f_sales, f_pricebook, f_stores])

    if not required_ok:
        st.info("Sab required files upload karo (left column ki 5 files) phir Run button active hoga.")

    if st.button("▶️  Run Demand Forecast", type="primary", disabled=not required_ok, key="df_run"):
        if abs(weight_sum - 1.0) > 0.001:
            st.error("Pehle weights fix karo — total 1.0 hona chahiye.")
            st.stop()

        with st.spinner("Processing... ye kuch minutes le sakta hai, wait karo..."):
            tmp = tempfile.mkdtemp(prefix="nw_df_")
            try:
                # Save uploads
                paths = {
                    "inventory": save_upload(f_inventory, tmp, "inventory.xlsx"),
                    "purchase": save_upload(f_purchase, tmp, "purchase.xlsx"),
                    "sales": save_upload(f_sales, tmp, "sales.xlsx"),
                    "pricebook": save_upload(f_pricebook, tmp, "pricebook.xlsx"),
                    "stores_sales": save_upload(f_stores, tmp, "stores_sales.xlsx"),
                    "purchase_append": save_upload(f_pur_append, tmp, "purchase_append.xlsx"),
                    "sales_append": save_upload(f_sal_append, tmp, "sales_append.xlsx"),
                    "open_po": save_upload(f_open_po, tmp, "open_po.xlsx"),
                    "planning_inputs": save_upload(f_planning, tmp, "planning_inputs.xlsx"),
                    "output": os.path.join(tmp, "Warehouse_Monthly_Sales.xlsx"),
                    "output_ordering": os.path.join(tmp, "Warehouse_Ordering_Sheet.xlsx"),
                    "history": os.path.join(tmp, "history.json"),
                }

                # Import and configure
                # Fresh import each run to reset module state
                if "demand_forecast" in sys.modules:
                    del sys.modules["demand_forecast"]
                sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
                import demand_forecast as df

                # Apply user settings
                df.DEMAND["lead_time"] = lead_time
                df.DEMAND["reorder_below_days"] = reorder_below
                df.DEMAND["target_days"] = target_days
                df.DEMAND["window_days"] = window_days
                df.DEMAND["weights"] = (w1, w2, w3)
                df.ORDERING_COLUMNS = list(selected_columns)
                df.ORDERING_PRIORITIES = list(selected_priorities)

                # Capture output
                log_buf = io.StringIO()
                with redirect_stdout(log_buf), redirect_stderr(log_buf):
                    df.run(paths)

                log_text = log_buf.getvalue()
                st.success("✅ Demand Forecast complete!")

                # Show log
                with st.expander("📋 Processing Log", expanded=False):
                    st.code(log_text, language="text")

                # Download buttons
                st.subheader("⬇️ Download Output Files")
                dc1, dc2 = st.columns(2)

                output_path = paths["output"]
                ordering_path = paths["output_ordering"]

                if os.path.isfile(output_path):
                    with open(output_path, "rb") as f:
                        dc1.download_button(
                            "📥 Monthly Sales Workbook",
                            data=f.read(),
                            file_name="Warehouse_Monthly_Sales.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            type="primary",
                        )

                if os.path.isfile(ordering_path):
                    with open(ordering_path, "rb") as f:
                        dc2.download_button(
                            "📥 Ordering Sheet Workbook",
                            data=f.read(),
                            file_name="Warehouse_Ordering_Sheet.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            type="primary",
                        )

            except Exception as e:
                st.error(f"❌ Error: {e}")
                with st.expander("Full Error Details"):
                    st.code(traceback.format_exc(), language="text")
            finally:
                shutil.rmtree(tmp, ignore_errors=True)


# ═══════════════════════ TAB 2: PO GENERATOR ═══════════════════════
with tab2:
    st.header("PO Generator")
    st.markdown("""
    Ordering Sheet workbook upload karo (Demand Forecast ka output, ya Excel me manually mark kiya hua).
    Jis row ko order karna hai uske PO column me **"PO"** likho, phir yahan upload karo.
    """)

    st.subheader("📁 File Upload Karo")
    f_ordering_wb = st.file_uploader("Ordering Sheet Workbook (.xlsx)", type=["xlsx"], key="po_wb")

    st.subheader("⚙️ Settings")
    price_basis = st.radio(
        "Price Basis (vendor ki kaunsi price use ho?)",
        options=["last", "lowest"],
        format_func=lambda x: "Last buying price + Last vendor" if x == "last" else "Lowest buying price + Lowest price vendor",
        horizontal=True,
        key="po_price",
    )

    if not f_ordering_wb:
        st.info("Ordering Sheet workbook upload karo phir Run button active hoga.")

    if st.button("▶️  Generate POs", type="primary", disabled=not f_ordering_wb, key="po_run"):
        with st.spinner("POs generate ho rahe hain..."):
            tmp = tempfile.mkdtemp(prefix="nw_po_")
            try:
                wb_path = save_upload(f_ordering_wb, tmp, "ordering_workbook.xlsx")
                output_folder = os.path.join(tmp, "PO_output")
                os.makedirs(output_folder, exist_ok=True)

                # Import and configure
                if "po_generator" in sys.modules:
                    del sys.modules["po_generator"]
                sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
                import po_generator as pg

                pg.PO_PRICE_BASIS = price_basis

                # Run PO generation (inline from main() logic)
                from datetime import date

                po_date = datetime.now(timezone(timedelta(hours=5))).date()
                log_buf = io.StringIO()

                with redirect_stdout(log_buf), redirect_stderr(log_buf):
                    print(f"Reading ordering workbook (prices: {price_basis.upper()} buying)", flush=True)
                    lines, skipped = pg.read_marked_lines(wb_path)

                    folder = Path(output_folder) / f"PO {po_date.isoformat()}"
                    folder.mkdir(parents=True, exist_ok=True)

                    orders = []
                    for number, (vendor, items) in enumerate(pg.group_by_vendor(lines), 1):
                        po_no = f"{pg.PO_PREFIX}-{po_date:%Y%m%d}-{number:02d}"
                        pdf = folder / pg.pdf_file_name(vendor, po_date)
                        total_qty, total_amount = pg.write_pdf(pdf, po_no, po_date, vendor, items)
                        orders.append((po_no, vendor, items, total_qty, total_amount))
                        print(f"  {po_no}  {vendor}: {len(items)} lines, qty {pg.qty_text(total_qty)}, {pg.money(total_amount)}", flush=True)

                    merged = folder / f"All_POs_{po_date.isoformat()}.xlsx"
                    pg.write_merged_excel(merged, po_date, orders, skipped)

                    print(f"\n{len(orders)} PO(s), {len(lines)} lines, total {pg.money(sum(o[4] for o in orders))}", flush=True)
                    if skipped:
                        print(f"WARNING: {len(skipped)} marked row(s) could not be ordered", flush=True)

                log_text = log_buf.getvalue()

                st.success(f"✅ {len(orders)} PO(s) generated — {len(lines)} total lines!")

                with st.expander("📋 Processing Log", expanded=False):
                    st.code(log_text, language="text")

                # Zip all output and offer download
                st.subheader("⬇️ Download Output")
                zip_data = zip_folder(str(folder))
                st.download_button(
                    f"📥 Download All POs (ZIP) — {len(orders)} vendor(s)",
                    data=zip_data,
                    file_name=f"PO_{po_date.isoformat()}.zip",
                    mime="application/zip",
                    type="primary",
                )

                # Also offer merged Excel separately
                if merged.is_file():
                    with open(merged, "rb") as f:
                        st.download_button(
                            "📥 All POs Merged Excel",
                            data=f.read(),
                            file_name=f"All_POs_{po_date.isoformat()}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        )

                if skipped:
                    with st.expander(f"⚠️ {len(skipped)} rows skipped (reason dekhein)"):
                        for item in skipped:
                            st.write(f"Row {item['row']}: **{item['description']}** → {item['reason']}")

            except Exception as e:
                st.error(f"❌ Error: {e}")
                with st.expander("Full Error Details"):
                    st.code(traceback.format_exc(), language="text")
            finally:
                shutil.rmtree(tmp, ignore_errors=True)

# ─────── Footer ───────
st.divider()
st.caption("Nation Wide Wholesale Automation Tool — Streamlit App")
