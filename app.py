"""
Nation Wide Wholesale - Automation Tool
Three tabs: Demand Forecast + PO Generator + PDI Sales Sorter
"""
import streamlit as st
import tempfile
import os
import sys
import io
import zipfile
import shutil
import traceback
import time
import threading
from pathlib import Path
from contextlib import redirect_stdout, redirect_stderr
from datetime import datetime, timezone, timedelta

print("APP: imports done", flush=True)

st.set_page_config(page_title="Nation Wide Automation", page_icon="📦", layout="wide")

print("APP: page_config done", flush=True)

# ─────────────────────────── custom CSS ───────────────────────────

st.markdown("""
<style>
/* ── Import a clean font ── */
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

/* ── Hide Streamlit/GitHub chrome (Fork, GitHub link, menu, viewer badge) for teammates ── */
[data-testid="stToolbar"], [data-testid="stDecoration"], [data-testid="stStatusWidget"],
#MainMenu, footer, [class*="viewerBadge"], [class*="profileContainer"], [class*="profilePreview"],
[class*="ProfileContainer"], [class*="ProfilePreview"], [data-testid="manage-app-button"], a[href*="share.streamlit.io/user/"],
a[href*="streamlit.io/user/"], iframe[title*="profile" i] {
    display: none !important;
    visibility: hidden !important;
}

/* ── Global tweaks ── */
.stApp {
    font-family: 'Inter', sans-serif;
}

/* ── Hero banner ── */
.hero-banner {
    background: linear-gradient(135deg, #1B3A5C 0%, #2E6B9E 50%, #4A90D9 100%);
    padding: 2rem 2.5rem;
    border-radius: 16px;
    margin-bottom: 1.5rem;
    color: white;
    position: relative;
    overflow: hidden;
}
.hero-banner::before {
    content: '';
    position: absolute;
    top: -50%;
    right: -20%;
    width: 400px;
    height: 400px;
    background: radial-gradient(circle, rgba(255,255,255,0.08) 0%, transparent 70%);
    border-radius: 50%;
}
.hero-banner h1 {
    font-size: 2rem;
    font-weight: 700;
    margin: 0 0 0.3rem 0;
    color: white !important;
    letter-spacing: -0.5px;
}
.hero-banner p {
    font-size: 1rem;
    opacity: 0.85;
    margin: 0;
    font-weight: 400;
}

/* ── Tab styling ── */
.stTabs [data-baseweb="tab-list"] {
    gap: 8px;
    background: #f0f2f6;
    padding: 6px;
    border-radius: 12px;
}
.stTabs [data-baseweb="tab"] {
    border-radius: 8px;
    padding: 10px 24px;
    font-weight: 600;
    font-size: 0.95rem;
}
.stTabs [aria-selected="true"] {
    background: white;
    box-shadow: 0 2px 8px rgba(0,0,0,0.08);
}

/* ── Section cards ── */
.section-card {
    background: #ffffff;
    border: 1px solid #e8ecf1;
    border-radius: 12px;
    padding: 1.5rem;
    margin-bottom: 1rem;
    box-shadow: 0 1px 4px rgba(0,0,0,0.04);
}
.section-card h3 {
    font-size: 1.1rem;
    font-weight: 600;
    color: #1B3A5C;
    margin: 0 0 0.8rem 0;
    display: flex;
    align-items: center;
    gap: 8px;
}

/* ── Metric chips (for file counts, status) ── */
.metric-row {
    display: flex;
    gap: 12px;
    flex-wrap: wrap;
    margin: 0.8rem 0;
}
.metric-chip {
    background: #EEF2FF;
    border: 1px solid #C7D2FE;
    border-radius: 8px;
    padding: 8px 16px;
    font-size: 0.85rem;
    font-weight: 500;
    color: #3730A3;
}
.metric-chip.green {
    background: #ECFDF5;
    border-color: #A7F3D0;
    color: #065F46;
}
.metric-chip.amber {
    background: #FFFBEB;
    border-color: #FDE68A;
    color: #92400E;
}

/* ── Info callout ── */
.info-callout {
    background: linear-gradient(135deg, #EEF2FF 0%, #E0E7FF 100%);
    border-left: 4px solid #6366F1;
    border-radius: 0 8px 8px 0;
    padding: 1rem 1.2rem;
    margin: 0.8rem 0;
    font-size: 0.9rem;
    color: #312E81;
}

/* ── Footer ── */
.custom-footer {
    text-align: center;
    padding: 1.5rem 0 0.5rem 0;
    color: #9CA3AF;
    font-size: 0.8rem;
    border-top: 1px solid #E5E7EB;
    margin-top: 2rem;
}
.custom-footer span {
    font-weight: 600;
    color: #6B7280;
}

/* ── Better file uploader ── */
[data-testid="stFileUploader"] {
    border-radius: 10px;
}
[data-testid="stFileUploader"] section {
    border-radius: 10px;
}

/* ── Button polish ── */
.stButton > button[kind="primary"] {
    border-radius: 10px;
    font-weight: 600;
    padding: 0.6rem 2rem;
    font-size: 1rem;
    letter-spacing: 0.2px;
}

/* ── Download button ── */
.stDownloadButton > button {
    border-radius: 10px;
    font-weight: 600;
}

/* ── Expander ── */
.streamlit-expanderHeader {
    font-weight: 600;
    font-size: 0.95rem;
}

/* ── Success message ── */
.stAlert [data-testid="stMarkdownContainer"] {
    font-weight: 500;
}
</style>
""", unsafe_allow_html=True)

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


class TeeWriter:
    """Write to a buffer, to stderr (Streamlit Cloud server logs) AND, live, to a box on the page
    so every user sees the progress while a run is going (not only the app owner)."""
    def __init__(self, buf, original, placeholder=None, tail_lines=30, interval=0.8):
        self.buf = buf
        self.original = original
        self.placeholder = placeholder
        self.tail_lines = tail_lines
        self.interval = interval
        self._last = 0.0
        self._busy = False

    def _refresh(self, force=False):
        if self.placeholder is None or self._busy:
            return
        now = time.time()
        if not force and now - self._last < self.interval:
            return
        self._busy = True
        try:
            lines = self.buf.getvalue().splitlines()[-self.tail_lines:]
            self.placeholder.code("\n".join(lines) or "Starting...", language="text")
            self._last = now
        except Exception:
            self.placeholder = None
        finally:
            self._busy = False

    def write(self, s):
        self.buf.write(s)
        self.original.write(s)
        self._refresh()
        return len(s)

    def flush(self):
        self.buf.flush()
        self.original.flush()
        self._refresh()

    def finish(self, clear=True):
        """Show the last lines once more; optionally remove the live box (the full log is shown afterwards)."""
        if self.placeholder is None:
            return
        if clear:
            try:
                self.placeholder.empty()
            except Exception:
                pass
        else:
            self._refresh(force=True)


@st.cache_resource
def _activity_store():
    """One store shared by every visitor/session of the running app (lost when the app restarts)."""
    return {"lock": threading.Lock(), "entries": []}


PKT = timezone(timedelta(hours=5))
MAX_ACTIVITY = 60          # keep the latest runs only
MAX_LOG_CHARS = 30000      # per-run log kept in memory (the tail is kept)


def log_activity(tool, user, status, seconds, log_text="", detail=""):
    """Record one run so everybody can see it in the Activity Log tab."""
    try:
        store = _activity_store()
        text = log_text or ""
        if len(text) > MAX_LOG_CHARS:
            text = "... (earlier lines cut) ...\n" + text[-MAX_LOG_CHARS:]
        entry = {
            "time": datetime.now(PKT).strftime("%d %b %Y, %I:%M:%S %p"),
            "tool": tool,
            "user": (user or "").strip() or "Unknown",
            "status": status,
            "seconds": round(seconds, 1),
            "detail": detail,
            "log": text,
        }
        with store["lock"]:
            store["entries"].insert(0, entry)
            del store["entries"][MAX_ACTIVITY:]
    except Exception as _e:
        print(f"APP: could not write activity log: {_e}", flush=True)


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
    "Lowest price date",
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

print("APP: constants done", flush=True)

# ─────────────────────────── UI ───────────────────────────────

try:
    # ── Hero Banner ──
    st.markdown("""
    <div class="hero-banner">
        <h1>📦 Nation Wide Automated Ordering</h1>
        <p>Upload your files, adjust settings, press Run — download the output.</p>
    </div>
    """, unsafe_allow_html=True)
    print("APP: title done", flush=True)

    user_name = st.text_input(
        "Your name (shown in the Activity Log)",
        key="user_name",
        placeholder="e.g. Ali",
    )

    tab1, tab2, tab3, tab4 = st.tabs(["📊 Demand Forecast", "📋 PO Generator", "🏪 PDI Sales Sorter", "🗂️ Activity Log"])
    print("APP: tabs created", flush=True)

    # ═══════════════════════ TAB 1: DEMAND FORECAST ═══════════════════════
    with tab1:
        st.markdown('<div class="section-card"><h3>📊 Demand Forecast</h3>'
                    'Upload Inventory, purchases, sales and price book — '
                    'script automatically makes the forecast.</div>',
                    unsafe_allow_html=True)

        col_left, col_right = st.columns(2)
        with col_left:
            st.markdown("##### Required Files")
            f_inventory = st.file_uploader("Inventory", type=["xlsx"], key="df_inv")
            f_purchase = st.file_uploader("6 Month Purchase", type=["xlsx"], key="df_pur")
            f_sales = st.file_uploader("6 Month Sales", type=["xlsx"], key="df_sal")
            f_pricebook = st.file_uploader("Master Price Book", type=["xlsx"], key="df_pb")
            f_stores = st.file_uploader("6 Months PDI Sales (Stores)", type=["xlsx"], key="df_stores")

        with col_right:
            st.markdown("##### Optional Files")
            f_pur_append = st.file_uploader("Current Purchase Warehouse (optional)", type=["xlsx"], key="df_pur_app")
            f_sal_append = st.file_uploader("Current Sales Warehouse (optional)", type=["xlsx"], key="df_sal_app")
            f_stores_append = st.file_uploader("Current Stores Sales (optional)", type=["xlsx"], key="df_stores_app")
            f_open_po = st.file_uploader("Open PO's (optional)", type=["xlsx"], key="df_po")
            f_planning = st.file_uploader("Planning Inputs (optional)", type=["xlsx"], key="df_plan")
            f_manual_inactive = st.file_uploader("Manual Inactive List (optional)", type=["xlsx"], key="df_manual_inactive",
                                                 help="SKUs listed here are forced to Inactive even if they have purchases in the window.")

        print("APP: tab1 uploaders done", flush=True)

        # ─────── Settings ───────
        st.markdown("---")
        st.markdown("##### ⚙️ Settings")

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

            po_alert_days = st.number_input(
                "Open PO Alert Days",
                value=0,
                min_value=0,
                key="df_po_alert",
                help="A PO older than this many days and still in the Open PO sheet is flagged as an alert. 0 = use Lead Time.",
            )

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
                st.error(f"Weights total is {weight_sum} — must be exactly 1.0!")

        with st.expander("Ordering Columns & Priorities"):
            selected_columns = st.multiselect(
                "Ordering Sheet Columns (order matters)",
                options=ALL_ORDERING_COLUMNS,
                default=DEFAULT_ORDERING_COLUMNS,
                key="df_cols",
            )
            selected_priorities = st.multiselect(
                "Ordering Priorities (these rows will appear in ordering sheet)",
                options=ALL_PRIORITIES,
                default=DEFAULT_PRIORITIES,
                key="df_pri",
            )

        print("APP: tab1 settings done", flush=True)

        # ─────── Run ───────
        required_ok = all([f_inventory, f_purchase, f_sales, f_pricebook, f_stores])

        # Show file upload status
        uploaded_count = sum(1 for f in [f_inventory, f_purchase, f_sales, f_pricebook, f_stores] if f)
        if not required_ok:
            st.markdown(
                f'<div class="info-callout">📁 Upload all 5 required files to activate the Run button. '
                f'<strong>{uploaded_count}/5</strong> uploaded so far.</div>',
                unsafe_allow_html=True,
            )

        st.markdown("")  # spacing
        if st.button("▶️  Run Demand Forecast", type="primary", disabled=not required_ok, key="df_run"):
            if abs(weight_sum - 1.0) > 0.001:
                st.error("Fix weights first — total must be 1.0.")
                st.stop()

            st.session_state.pop("df_result", None)
            with st.spinner("Processing... this may take a few minutes, please wait..."):
                tmp = tempfile.mkdtemp(prefix="nw_df_")
                _t0 = time.time()
                log_buf = None
                try:
                    paths = {
                        "inventory": save_upload(f_inventory, tmp, "inventory.xlsx"),
                        "purchase": save_upload(f_purchase, tmp, "purchase.xlsx"),
                        "sales": save_upload(f_sales, tmp, "sales.xlsx"),
                        "pricebook": save_upload(f_pricebook, tmp, "pricebook.xlsx"),
                        "stores_sales": save_upload(f_stores, tmp, "stores_sales.xlsx"),
                        "stores_sales_append": save_upload(f_stores_append, tmp, "stores_sales_append.xlsx"),
                        "purchase_append": save_upload(f_pur_append, tmp, "purchase_append.xlsx"),
                        "sales_append": save_upload(f_sal_append, tmp, "sales_append.xlsx"),
                        "open_po": save_upload(f_open_po, tmp, "open_po.xlsx"),
                        "planning_inputs": save_upload(f_planning, tmp, "planning_inputs.xlsx"),
                        "manual_inactive": save_upload(f_manual_inactive, tmp, "manual_inactive.xlsx"),
                        "output": os.path.join(tmp, "Warehouse_Monthly_Sales.xlsx"),
                        "output_ordering": os.path.join(tmp, "Warehouse_Ordering_Sheet.xlsx"),
                        "history": os.path.join(tmp, "history.json"),
                    }

                    # Verify saved files before processing
                    for label, fpath in paths.items():
                        if fpath and os.path.isfile(fpath):
                            sz = os.path.getsize(fpath)
                            print(f"APP: saved {label} -> {fpath} ({sz:,} bytes)", flush=True)
                        elif fpath:
                            print(f"APP: WARNING {label} -> {fpath} NOT FOUND", flush=True)

                    # Quick header check on the sales file
                    try:
                        import openpyxl as _oxl
                        _wb = _oxl.load_workbook(paths["sales"], read_only=True, data_only=True)
                        print(f"APP: sales file sheets: {_wb.sheetnames}", flush=True)
                        for _ws in _wb.worksheets:
                            for _rn, _row in enumerate(_ws.iter_rows(min_row=1, max_row=8), 1):
                                _vals = [c.value for c in _row if c.value is not None]
                                if _vals:
                                    print(f"APP: sales sheet '{_ws.title}' row {_rn}: {_vals[:6]}", flush=True)
                        _wb.close()
                    except Exception as _e:
                        print(f"APP: sales file check error: {_e}", flush=True)

                    print("APP: files saved, importing demand_forecast...", flush=True)

                    if "demand_forecast" in sys.modules:
                        del sys.modules["demand_forecast"]
                    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
                    import demand_forecast as df

                    print("APP: demand_forecast imported OK", flush=True)

                    df.DEMAND["lead_time"] = lead_time
                    df.DEMAND["reorder_below_days"] = reorder_below
                    df.DEMAND["target_days"] = target_days
                    df.DEMAND["window_days"] = window_days
                    df.DEMAND["weights"] = (w1, w2, w3)
                    df.OPEN_PO_ALERT_DAYS = int(po_alert_days) if po_alert_days else None
                    # Skip the raw/current copies in the output: much smaller file and far less RAM when saving
                    df.INCLUDE_RAW_TABS = False
                    df.ORDERING_COLUMNS = list(selected_columns)
                    df.ORDERING_PRIORITIES = list(selected_priorities)

                    log_buf = io.StringIO()
                    live_box = st.empty()
                    tee_out = TeeWriter(log_buf, sys.__stderr__, live_box)

                    print("APP: running df.run(paths)...", flush=True)
                    with redirect_stdout(tee_out), redirect_stderr(tee_out):
                        df.run(paths)
                    tee_out.finish(clear=True)

                    log_text = log_buf.getvalue()
                    print("APP: df.run() completed!", flush=True)

                    # Build the ZIP and keep everything in session_state BEFORE any st.* call, so the result
                    # survives even if the browser reconnected and reran the page during the long run.
                    output_path = paths["output"]
                    ordering_path = paths["output_ordering"]
                    detail_path = paths.get("output_detail", "")
                    zip_buf = io.BytesIO()
                    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                        if os.path.isfile(output_path):
                            zf.write(output_path, "Warehouse_Monthly_Sales.xlsx")
                        if os.path.isfile(ordering_path):
                            zf.write(ordering_path, "Warehouse_Ordering_Sheet.xlsx")
                        if detail_path and os.path.isfile(detail_path):
                            zf.write(detail_path, "Warehouse_Monthly_Sales_Detail_Sheets.xlsx")
                    st.session_state["df_result"] = {"zip": zip_buf.getvalue(), "log": log_text}
                    log_activity("Demand Forecast", user_name, "Success", time.time() - _t0, log_text,
                                 f"Inventory: {f_inventory.name}")

                except Exception as e:
                    print(f"APP: ERROR in demand forecast: {e}", flush=True)
                    print(traceback.format_exc(), flush=True)
                    log_activity("Demand Forecast", user_name, "Failed", time.time() - _t0,
                                 (log_buf.getvalue() if log_buf else "") + "\n" + traceback.format_exc(), str(e))
                    st.error(f"❌ Error: {e}")
                    if log_buf is not None:
                        with st.expander("Processing log up to the error", expanded=True):
                            st.code("\n".join(log_buf.getvalue().splitlines()[-60:]), language="text")
                    with st.expander("Full Error Details"):
                        st.code(traceback.format_exc(), language="text")
                finally:
                    shutil.rmtree(tmp, ignore_errors=True)

        # Show the stored result (also after a browser reconnect / page rerun)
        df_result = st.session_state.get("df_result")
        if df_result:
            st.success("✅ Demand Forecast complete!")
            with st.expander("📋 Processing Log", expanded=False):
                st.code(df_result["log"], language="text")
            st.markdown("##### ⬇️ Download Output Files")
            st.download_button(
                "📥 Download All Output Files (ZIP)",
                data=df_result["zip"],
                file_name="Demand_Forecast_Output.zip",
                mime="application/zip",
                type="primary",
                key="df_download",
            )

    print("APP: tab1 complete", flush=True)

    # ═══════════════════════ TAB 2: PO GENERATOR ═══════════════════════
    with tab2:
        st.markdown(
            '<div class="section-card"><h3>📋 PO Generator</h3>'
            'Upload the Ordering Sheet workbook (output from Demand Forecast, or manually marked). '
            'Write <strong>"PO"</strong> in the PO column for each row you want to order, then upload here.</div>',
            unsafe_allow_html=True,
        )

        f_ordering_wb = st.file_uploader("Ordering Sheet Workbook (.xlsx)", type=["xlsx"], key="po_wb")

        st.markdown("---")
        st.markdown("##### ⚙️ Settings")
        price_basis = st.radio(
            "Vendor Basis (which vendor to order from?)",
            options=["last", "lowest"],
            format_func=lambda x: "Last vendor" if x == "last" else "Lowest price vendor",
            horizontal=True,
            key="po_price",
        )

        if not f_ordering_wb:
            st.markdown(
                '<div class="info-callout">📁 Upload the Ordering Sheet workbook to activate the Run button.</div>',
                unsafe_allow_html=True,
            )

        st.markdown("")  # spacing
        if st.button("▶️  Generate POs", type="primary", disabled=not f_ordering_wb, key="po_run"):
            with st.spinner("Generating POs..."):
                tmp = tempfile.mkdtemp(prefix="nw_po_")
                _t0 = time.time()
                log_buf = None
                try:
                    wb_path = save_upload(f_ordering_wb, tmp, "ordering_workbook.xlsx")
                    output_folder = os.path.join(tmp, "PO_output")
                    os.makedirs(output_folder, exist_ok=True)

                    if "po_generator" in sys.modules:
                        del sys.modules["po_generator"]
                    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
                    import po_generator as pg

                    pg.PO_PRICE_BASIS = price_basis

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
                            total_qty = pg.write_pdf(pdf, po_no, po_date, vendor, items)
                            orders.append((po_no, vendor, items, total_qty))
                            print(f"  {po_no}  {vendor}: {len(items)} lines, qty {pg.qty_text(total_qty)}", flush=True)

                        merged = folder / f"All_POs_{po_date.isoformat()}.xlsx"
                        pg.write_merged_excel(merged, po_date, orders, skipped)

                        print(f"\n{len(orders)} PO(s), {len(lines)} lines, total qty {pg.qty_text(sum(o[3] for o in orders))}", flush=True)
                        if skipped:
                            print(f"WARNING: {len(skipped)} marked row(s) could not be ordered", flush=True)

                    log_text = log_buf.getvalue()
                    log_activity("PO Generator", user_name, "Success", time.time() - _t0, log_text,
                                 f"{len(orders)} PO(s), {len(lines)} lines")

                    st.success(f"✅ {len(orders)} PO(s) generated — {len(lines)} total lines!")

                    with st.expander("📋 Processing Log", expanded=False):
                        st.code(log_text, language="text")

                    st.markdown("##### ⬇️ Download Output")
                    zip_data = zip_folder(str(folder))
                    st.download_button(
                        f"📥 Download All POs (ZIP) — {len(orders)} vendor(s)",
                        data=zip_data,
                        file_name=f"PO_{po_date.isoformat()}.zip",
                        mime="application/zip",
                        type="primary",
                    )

                    if merged.is_file():
                        with open(merged, "rb") as f:
                            st.download_button(
                                "📥 All POs Merged Excel",
                                data=f.read(),
                                file_name=f"All_POs_{po_date.isoformat()}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            )

                    if skipped:
                        with st.expander(f"⚠️ {len(skipped)} rows skipped (see reasons)"):
                            for item in skipped:
                                st.write(f"Row {item['row']}: **{item['description']}** → {item['reason']}")

                except Exception as e:
                    print(f"APP: ERROR in PO generator: {e}", flush=True)
                    print(traceback.format_exc(), flush=True)
                    log_activity("PO Generator", user_name, "Failed", time.time() - _t0,
                                 (log_buf.getvalue() if log_buf else "") + "\n" + traceback.format_exc(), str(e))
                    st.error(f"❌ Error: {e}")
                    with st.expander("Full Error Details"):
                        st.code(traceback.format_exc(), language="text")
                finally:
                    shutil.rmtree(tmp, ignore_errors=True)

    print("APP: tab2 complete", flush=True)

    # ═══════════════════════ TAB 3: PDI SALES SORTER ═══════════════════════
    with tab3:
        st.markdown(
            '<div class="section-card"><h3>🏪 PDI Sales Sorter</h3>'
            'Upload one or more PDI store sales files (.xlsx). '
            'Filters rows by department keyword, groups by description, '
            'and produces a summary workbook with one sheet per file.</div>',
            unsafe_allow_html=True,
        )

        pdi_uploads = st.file_uploader(
            "PDI Sales Files (upload any number of .xlsx files)",
            type=["xlsx"],
            accept_multiple_files=True,
            key="pdi_files",
        )

        # Show uploaded file count
        if pdi_uploads:
            st.markdown(
                f'<div class="metric-row">'
                f'<span class="metric-chip green">✓ {len(pdi_uploads)} file(s) uploaded</span>'
                f'</div>',
                unsafe_allow_html=True,
            )

        st.markdown("---")
        st.markdown("##### ⚙️ Settings")
        sc1, sc2 = st.columns([2, 1])
        with sc1:
            dept_keyword = st.text_input(
                "Department Keyword (case-insensitive filter)",
                value="novelty",
                key="pdi_dept",
            )
        with sc2:
            save_detail = st.checkbox(
                "Generate detail workbooks",
                value=False,
                key="pdi_detail",
                help="Creates per-file Merged + Filtered detail sheets alongside the summary.",
            )

        if not pdi_uploads:
            st.markdown(
                '<div class="info-callout">📁 Upload at least one PDI sales file to activate the Run button.</div>',
                unsafe_allow_html=True,
            )

        st.markdown("")  # spacing
        if st.button("▶️  Sort PDI Sales", type="primary", disabled=not pdi_uploads, key="pdi_run"):
            if not dept_keyword.strip():
                st.error("Enter a department keyword to filter by.")
                st.stop()

            with st.spinner("Processing PDI sales files..."):
                tmp = tempfile.mkdtemp(prefix="nw_pdi_")
                _t0 = time.time()
                log_buf = None
                try:
                    # Save uploaded files and build input_files dict (label -> path)
                    input_map = {}
                    for uploaded in pdi_uploads:
                        label = Path(uploaded.name).stem
                        # Handle duplicate filenames
                        base_label = label
                        counter = 2
                        while label in input_map:
                            label = f"{base_label}_{counter}"
                            counter += 1
                        fpath = save_upload(uploaded, tmp, f"{label}.xlsx")
                        input_map[label] = fpath

                    file_names = ", ".join(input_map.keys())
                    st.markdown(
                        f'<div class="metric-row">'
                        f'<span class="metric-chip">🔍 Filter: "{dept_keyword.strip()}"</span>'
                        f'<span class="metric-chip">{len(input_map)} file(s)</span>'
                        f'</div>',
                        unsafe_allow_html=True,
                    )

                    if "pdi_sales_sorter" in sys.modules:
                        del sys.modules["pdi_sales_sorter"]
                    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
                    import pdi_sales_sorter as pdi

                    output_dir = os.path.join(tmp, "pdi_output")
                    log_buf = io.StringIO()
                    live_box = st.empty()
                    tee_out = TeeWriter(log_buf, sys.__stderr__, live_box)

                    print("APP: running pdi_sales_sorter.run()...", flush=True)
                    with redirect_stdout(tee_out), redirect_stderr(tee_out):
                        result = pdi.run(
                            input_files=input_map,
                            output_dir=output_dir,
                            department_contains=dept_keyword.strip(),
                            save_detail=save_detail,
                        )
                    tee_out.finish(clear=True)

                    log_text = log_buf.getvalue()
                    print("APP: pdi_sales_sorter.run() completed!", flush=True)
                    log_activity("PDI Sales Sorter", user_name, "Success", time.time() - _t0, log_text,
                                 f"{len(input_map)} file(s), filter: {dept_keyword.strip()}")
                    st.success(f"✅ PDI Sales Sorter complete — {len(input_map)} file(s) processed!")

                    with st.expander("📋 Processing Log", expanded=False):
                        st.code(log_text, language="text")

                    # Per-file reports
                    if result.get("reports"):
                        with st.expander("📊 Per-File Reports"):
                            for report in result["reports"]:
                                st.text(report)

                    st.markdown("##### ⬇️ Download Output")

                    # Summary workbook always available
                    summary_path = result.get("summary", "")
                    if summary_path and os.path.isfile(summary_path):
                        with open(summary_path, "rb") as f:
                            st.download_button(
                                "📥 Download Summary Workbook",
                                data=f.read(),
                                file_name="Stores_Sales_Summary.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                type="primary",
                            )

                    # If detail files exist, bundle everything into a ZIP
                    detail_files = result.get("details", [])
                    if detail_files:
                        zip_buf = io.BytesIO()
                        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                            if summary_path and os.path.isfile(summary_path):
                                zf.write(summary_path, "Stores_Sales_Summary.xlsx")
                            for dp in detail_files:
                                if os.path.isfile(dp):
                                    zf.write(dp, os.path.basename(dp))
                        zip_buf.seek(0)
                        st.download_button(
                            "📥 Download All Files (ZIP)",
                            data=zip_buf.getvalue(),
                            file_name="PDI_Sales_Sorted.zip",
                            mime="application/zip",
                        )

                except Exception as e:
                    print(f"APP: ERROR in PDI sales sorter: {e}", flush=True)
                    print(traceback.format_exc(), flush=True)
                    log_activity("PDI Sales Sorter", user_name, "Failed", time.time() - _t0,
                                 (log_buf.getvalue() if log_buf else "") + "\n" + traceback.format_exc(), str(e))
                    st.error(f"❌ Error: {e}")
                    with st.expander("Full Error Details"):
                        st.code(traceback.format_exc(), language="text")
                finally:
                    shutil.rmtree(tmp, ignore_errors=True)

    print("APP: tab3 complete", flush=True)

    # ═══════════════════════ TAB 4: ACTIVITY LOG ═══════════════════════
    with tab4:
        st.markdown(
            '<div class="section-card"><h3>🗂️ Activity Log</h3>'
            'Every run by every user (latest first): who ran what, when, whether it worked, and the full processing log. '
            'Open a row to read its log. The list resets when the app restarts.</div>',
            unsafe_allow_html=True,
        )
        st.button("🔄 Refresh", key="act_refresh")
        _store = _activity_store()
        with _store["lock"]:
            _entries = list(_store["entries"])
        if not _entries:
            st.markdown('<div class="info-callout">No runs yet since the app last started.</div>', unsafe_allow_html=True)
        for _i, _e in enumerate(_entries):
            _icon = "✅" if _e["status"] == "Success" else "❌"
            _title = f'{_icon} {_e["time"]} — {_e["tool"]} — {_e["user"]} ({_e["seconds"]}s)'
            with st.expander(_title, expanded=False):
                if _e["detail"]:
                    st.caption(_e["detail"])
                st.code(_e["log"] or "(no log)", language="text")

    print("APP: tab4 complete", flush=True)

    # ─────── Footer ───────
    st.markdown(
        '<div class="custom-footer">'
        '<span>Nation Wide Wholesale</span> — Automation Tool &nbsp;|&nbsp; '
        'Demand Forecast &bull; PO Generator &bull; PDI Sales Sorter'
        '</div>',
        unsafe_allow_html=True,
    )
    print("APP: fully rendered!", flush=True)

except Exception as e:
    print(f"APP: FATAL ERROR: {e}", flush=True)
    print(traceback.format_exc(), flush=True)
    st.error(f"❌ App crashed: {e}")
    st.code(traceback.format_exc())
