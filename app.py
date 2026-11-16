"""
Furniture Asset Tracker  -  Streamlit + Supabase + pre-printed blank QR codes
=============================================================================

SETUP
-----
1) requirements.txt
       streamlit>=1.35
       supabase>=2.4
       streamlit-qrcode-scanner>=0.1.2
       pandas>=2.0
       openpyxl>=3.1
       reportlab>=4.0
       qrcode[pil]>=7.4

2) Supabase: run this SQL once in  Dashboard -> SQL Editor
   ---------------------------------------------------------------------------
   create extension if not exists "pgcrypto";

   create table if not exists public.assets (
     id            uuid primary key default gen_random_uuid(),
     qr_code_id    text not null unique,
     category      text not null,
     location      text not null,
     assigned_user text,
     purchase_date date,
     status        text not null default 'Active'
                   check (status in ('Active','Repair','Damaged','Disposed')),
     created_at    timestamptz not null default now()
   );

   create table if not exists public.inspection_logs (
     log_id           bigint generated always as identity primary key,
     qr_code_id       text not null
                      references public.assets(qr_code_id)
                      on update cascade on delete cascade,
     condition_status text not null
                      check (condition_status in ('Good','Fair','Needs Repair','Broken')),
     quantity         integer not null default 1 check (quantity >= 0),
     remarks          text,
     checked_by       text not null,
     "timestamp"      timestamptz not null default now()
   );

   create index if not exists idx_logs_qr_time
     on public.inspection_logs (qr_code_id, "timestamp" desc);

   -- Lock the tables down. The service_role key used by this server-side app
   -- bypasses RLS, while the public anon key gets no access at all.
   alter table public.assets          enable row level security;
   alter table public.inspection_logs enable row level security;
   ---------------------------------------------------------------------------

3) Secrets  ->  .streamlit/secrets.toml  (local)  or  App settings -> Secrets
   (Streamlit Community Cloud). NEVER commit this file to git.

       [supabase]
       url = "https://YOUR-PROJECT-ID.supabase.co"
       key = "YOUR_SERVICE_ROLE_KEY"     # Project Settings -> API -> service_role

       [admin]
       password = "choose-a-strong-password"   # protects the QR generator page

   The key stays on the Streamlit server, so the service_role key is never
   exposed to browsers. Do not use it in any client-side code.

4) Run:   streamlit run app.py
   NOTE: browsers only allow camera access over HTTPS (or localhost). Streamlit
   Community Cloud is HTTPS by default.
"""

import hmac
import io
import re
from datetime import date, datetime
from xml.sax.saxutils import escape

import pandas as pd
import qrcode
import streamlit as st
from qrcode.constants import ERROR_CORRECT_M
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
from streamlit_qrcode_scanner import qrcode_scanner
from supabase import Client, create_client

# --------------------------------------------------------------------------- #
# Page config (must be the first Streamlit call)
# --------------------------------------------------------------------------- #
st.set_page_config(
    page_title="Furniture Asset Tracker",
    page_icon="🪑",
    layout="wide",
    initial_sidebar_state="auto",
)

# --------------------------------------------------------------------------- #
# Configuration - edit these lists to match your factory / office
# --------------------------------------------------------------------------- #
CATEGORIES = ["Chair", "Table", "Desk", "Locker", "Cabinet", "Rack", "Bench", "Sofa", "Other"]
LOCATIONS = [
    "Admin Floor 1", "Admin Floor 2", "Production Floor 1", "Production Floor 2",
    "Store Room", "Canteen", "Other",
]
CONDITIONS = ["Good", "Fair", "Needs Repair", "Broken"]
STATUSES = ["Active", "Repair", "Damaged", "Disposed"]

# Status assigned automatically when a new asset is registered
CONDITION_TO_STATUS = {"Good": "Active", "Fair": "Active", "Needs Repair": "Repair", "Broken": "Damaged"}

DEFAULT_PREFIX = "RBL-FURN"
QR_ID_PATTERN = re.compile(r"^[A-Z0-9][A-Z0-9_\-]{2,49}$")

# Slightly tighter padding for phones
st.markdown(
    "<style>.block-container{padding-top:1.5rem;padding-bottom:3rem;}</style>",
    unsafe_allow_html=True,
)


# --------------------------------------------------------------------------- #
# Supabase helpers
# --------------------------------------------------------------------------- #
@st.cache_resource(show_spinner=False)
def _create_client(url: str, key: str) -> Client:
    """One shared Supabase client per server process."""
    return create_client(url, key)


def get_supabase() -> Client:
    """Read credentials from st.secrets and return a client (stops app if missing)."""
    try:
        url = st.secrets["supabase"]["url"]
        key = st.secrets["supabase"]["key"]
    except Exception:
        st.error(
            "Supabase credentials are missing. Add a `[supabase]` section with "
            "`url` and `key` to `.streamlit/secrets.toml` (see the setup notes at "
            "the top of app.py)."
        )
        st.stop()
    return _create_client(url, key)


def fetch_all(table: str, order_col: str) -> list[dict]:
    """Fetch every row of a table, paging past Supabase's 1000-row default limit."""
    sb, rows, start, page = get_supabase(), [], 0, 1000
    while True:
        res = (
            sb.table(table).select("*").order(order_col)
            .range(start, start + page - 1).execute()
        )
        batch = res.data or []
        rows.extend(batch)
        if len(batch) < page:
            return rows
        start += page


INVENTORY_COLS = [
    "qr_code_id", "category", "location", "assigned_user", "purchase_date",
    "status", "condition", "quantity", "checked_by", "last_checked",
]
DISPLAY_NAMES = {
    "qr_code_id": "QR Code ID", "category": "Category", "location": "Location",
    "assigned_user": "Assigned User", "purchase_date": "Purchase Date",
    "status": "Status", "condition": "Condition", "quantity": "Qty",
    "checked_by": "Last Checked By", "last_checked": "Last Checked",
}


@st.cache_data(ttl=30, show_spinner="Loading inventory...")
def load_inventory() -> pd.DataFrame:
    """Join every asset with its most recent inspection log."""
    assets = pd.DataFrame(fetch_all("assets", "created_at"))
    if assets.empty:
        return pd.DataFrame(columns=INVENTORY_COLS)

    logs = pd.DataFrame(fetch_all("inspection_logs", "timestamp"))
    if logs.empty:
        assets["condition"] = "Not inspected"
        for col in ("quantity", "checked_by", "last_checked"):
            assets[col] = None
    else:
        logs["timestamp"] = pd.to_datetime(logs["timestamp"], utc=True, errors="coerce")
        latest = (
            logs.sort_values("timestamp")
            .drop_duplicates("qr_code_id", keep="last")
            .rename(columns={"condition_status": "condition", "timestamp": "last_checked"})
            [["qr_code_id", "condition", "quantity", "checked_by", "last_checked"]]
        )
        # Naive timestamps keep Excel / display formatting simple
        latest["last_checked"] = latest["last_checked"].dt.tz_localize(None)
        assets = assets.merge(latest, on="qr_code_id", how="left")
        assets["condition"] = assets["condition"].fillna("Not inspected")

    return assets[INVENTORY_COLS].sort_values("qr_code_id").reset_index(drop=True)


# --------------------------------------------------------------------------- #
# Small UI helpers
# --------------------------------------------------------------------------- #
def with_current(options: list[str], current: str | None) -> list[str]:
    """Make sure a value already stored in the DB is selectable in a dropdown."""
    opts = list(options)
    if current and current not in opts:
        opts.insert(0, current)
    return opts


def resolve_choice(selected: str, custom: str) -> str:
    """If 'Other' is selected, use the custom text box instead."""
    return custom.strip() if selected == "Other" else selected


def normalize_code(raw: str) -> str:
    return (raw or "").strip().upper()


def finish_and_return_to_scanner(message: str) -> None:
    """Show a flash message on the next run and reset the scanner for the next item."""
    st.session_state["flash"] = message
    st.session_state["active_code"] = None
    st.session_state["scan_nonce"] += 1  # new widget key -> fresh scanner
    st.rerun()


def reset_scanner() -> None:
    st.session_state["active_code"] = None
    st.session_state["scan_nonce"] += 1


def set_active_code(raw: str) -> None:
    """Validate a scanned / typed code and make it the active one."""
    code = normalize_code(raw)
    if not QR_ID_PATTERN.match(code):
        st.error(
            f"'{raw}' does not look like one of our asset QR codes "
            "(letters, digits, '-' or '_' only). Please scan again."
        )
        return
    st.session_state["active_code"] = code
    st.rerun()


# --------------------------------------------------------------------------- #
# PAGE 1 - Scan & Register / Update
# --------------------------------------------------------------------------- #
def render_register_form(code: str, inspector: str) -> None:
    sb = get_supabase()
    st.info(f"🆕 **{code}** is not registered yet. Fill in the details to register it.")

    with st.form("register_form"):
        category = st.selectbox("Category", CATEGORIES)
        category_custom = st.text_input("Custom category (only if you chose 'Other')")
        location = st.selectbox("Location", LOCATIONS)
        location_custom = st.text_input("Custom location (only if you chose 'Other')")
        assigned_user = st.text_input("Assigned user (employee name)")
        purchase_date = st.date_input("Purchase date (optional)", value=None, max_value=date.today())
        condition = st.selectbox("Current condition", CONDITIONS)
        quantity = st.number_input("Quantity", min_value=1, value=1, step=1)
        remarks = st.text_area("Remarks (optional)")
        submitted = st.form_submit_button("✅ Register furniture", type="primary")

    if not submitted:
        return

    category_final = resolve_choice(category, category_custom)
    location_final = resolve_choice(location, location_custom)
    if not inspector.strip():
        st.error("Please enter your name in the sidebar ('Inspector name') first.")
        return
    if not category_final or not location_final:
        st.error("Category and location are required.")
        return

    try:
        sb.table("assets").insert({
            "qr_code_id": code,
            "category": category_final,
            "location": location_final,
            "assigned_user": assigned_user.strip() or None,
            "purchase_date": purchase_date.isoformat() if purchase_date else None,
            "status": CONDITION_TO_STATUS[condition],
        }).execute()
    except Exception as exc:  # e.g. unique violation if someone registered it a second ago
        st.error(f"Could not save the asset: {exc}")
        return

    try:
        sb.table("inspection_logs").insert({
            "qr_code_id": code,
            "condition_status": condition,
            "quantity": int(quantity),
            "remarks": remarks.strip() or "Initial registration",
            "checked_by": inspector.strip(),
        }).execute()
    except Exception as exc:
        st.warning(f"Asset saved, but the first inspection log failed: {exc}")

    load_inventory.clear()
    finish_and_return_to_scanner(f"✅ {code} registered as {category_final} in {location_final}.")


def render_update_form(asset: dict, inspector: str) -> None:
    sb = get_supabase()
    code = asset["qr_code_id"]

    # Latest logs (used for details panel, defaults and history)
    logs = (
        sb.table("inspection_logs").select("*").eq("qr_code_id", code)
        .order("timestamp", desc=True).limit(10).execute().data or []
    )
    last = logs[0] if logs else {}

    st.success(f"✔ **{code}** is registered.")
    if asset["status"] == "Disposed":
        st.warning("This asset is marked as **Disposed**. Updating it will still be logged.")

    c1, c2, c3 = st.columns(3)
    c1.metric("Category", asset["category"])
    c2.metric("Location", asset["location"])
    c3.metric("Assigned user", asset.get("assigned_user") or "—")
    c4, c5, c6 = st.columns(3)
    c4.metric("Status", asset["status"])
    c5.metric("Last condition", last.get("condition_status", "—"))
    c6.metric("Purchase date", asset.get("purchase_date") or "—")

    location_opts = with_current(LOCATIONS, asset["location"])
    cond_default = last.get("condition_status", "Good")

    st.subheader("Log inspection / update")
    with st.form("update_form"):
        condition = st.selectbox(
            "Condition", CONDITIONS,
            index=CONDITIONS.index(cond_default) if cond_default in CONDITIONS else 0,
        )
        quantity = st.number_input("Quantity", min_value=0, value=int(last.get("quantity") or 1), step=1)
        location = st.selectbox("Location", location_opts, index=location_opts.index(asset["location"]))
        location_custom = st.text_input("Custom location (only if you chose 'Other')")
        assigned_user = st.text_input("Assigned user", value=asset.get("assigned_user") or "")
        status = st.selectbox("Status", STATUSES, index=STATUSES.index(asset["status"]))
        remarks = st.text_area("Remarks (optional)")
        st.caption(
            "If the asset is still 'Active' and you report 'Needs Repair' or 'Broken', "
            "its status is switched to 'Repair' / 'Damaged' automatically."
        )
        submitted = st.form_submit_button("💾 Save inspection", type="primary")

    if submitted:
        if not inspector.strip():
            st.error("Please enter your name in the sidebar ('Inspector name') first.")
            return
        new_location = resolve_choice(location, location_custom)
        if not new_location:
            st.error("Please type the custom location.")
            return

        # Auto-escalate status for problem conditions
        new_status = status
        if status == "Active":
            new_status = {"Needs Repair": "Repair", "Broken": "Damaged"}.get(condition, status)

        # Work out what actually changed (used for the update and the audit remark)
        updates, notes = {}, []
        new_user = assigned_user.strip() or None
        if new_location != asset["location"]:
            updates["location"] = new_location
            notes.append(f"Location: {asset['location']} → {new_location}")
        if new_user != (asset.get("assigned_user") or None):
            updates["assigned_user"] = new_user
            notes.append(f"User: {asset.get('assigned_user') or '—'} → {new_user or '—'}")
        if new_status != asset["status"]:
            updates["status"] = new_status
            notes.append(f"Status: {asset['status']} → {new_status}")

        full_remarks = " | ".join(filter(None, [remarks.strip(), "; ".join(notes)])) or None

        try:
            if updates:
                sb.table("assets").update(updates).eq("qr_code_id", code).execute()
            sb.table("inspection_logs").insert({
                "qr_code_id": code,
                "condition_status": condition,
                "quantity": int(quantity),
                "remarks": full_remarks,
                "checked_by": inspector.strip(),
            }).execute()
        except Exception as exc:
            st.error(f"Could not save the update: {exc}")
            return

        load_inventory.clear()
        finish_and_return_to_scanner(f"✅ {code} updated - condition logged as '{condition}'.")

    if logs:
        with st.expander("Recent inspection history", expanded=False):
            hist = pd.DataFrame(logs)[
                ["timestamp", "condition_status", "quantity", "remarks", "checked_by"]
            ]
            hist["timestamp"] = (
                pd.to_datetime(hist["timestamp"], utc=True).dt.tz_localize(None).dt.strftime("%Y-%m-%d %H:%M")
            )
            st.dataframe(hist, hide_index=True)


def page_scan(inspector: str) -> None:
    st.header("📷 Scan & Register / Update")
    ss = st.session_state

    if ss.get("flash"):
        st.success(ss.pop("flash"))

    # ---- Step 1: obtain a code ------------------------------------------------
    if ss["active_code"] is None:
        tab_cam, tab_manual = st.tabs(["📷 Camera scan", "⌨️ Type code"])
        with tab_cam:
            st.caption("Allow camera access, then point it at the QR sticker.")
            scanned = qrcode_scanner(key=f"scanner_{ss['scan_nonce']}")
            if scanned:
                set_active_code(scanned)
        with tab_manual:
            with st.form("manual_form"):
                manual = st.text_input("QR code ID", placeholder=f"{DEFAULT_PREFIX}-0001")
                if st.form_submit_button("Look up") and manual:
                    set_active_code(manual)
        return

    # ---- Step 2: register or update ------------------------------------------
    code = ss["active_code"]
    st.button("🔄 Scan another item", on_click=reset_scanner)

    try:
        found = get_supabase().table("assets").select("*").eq("qr_code_id", code).limit(1).execute().data
    except Exception as exc:
        st.error(f"Database error: {exc}")
        return

    if found:
        render_update_form(found[0], inspector)
    else:
        render_register_form(code, inspector)


# --------------------------------------------------------------------------- #
# Export helpers (Excel + PDF)
# --------------------------------------------------------------------------- #
def _fmt(value) -> str:
    """Safe string for any cell (handles NaN / None / timestamps)."""
    if value is None or (not isinstance(value, str) and pd.isna(value)):
        return ""
    if isinstance(value, (pd.Timestamp, datetime)):
        return value.strftime("%Y-%m-%d %H:%M")
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def build_inventory_excel(df: pd.DataFrame) -> bytes:
    out = io.BytesIO()
    export = df.rename(columns=DISPLAY_NAMES)
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        export.to_excel(writer, index=False, sheet_name="Assets")
        ws = writer.sheets["Assets"]
        ws.freeze_panes = "A2"
        for idx, col in enumerate(export.columns, start=1):  # auto column width
            longest = max([len(str(col))] + [len(_fmt(v)) for v in export[col]])
            ws.column_dimensions[ws.cell(row=1, column=idx).column_letter].width = min(longest + 3, 45)
    return out.getvalue()


def build_inventory_pdf(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    page = landscape(A4)
    doc = SimpleDocTemplate(
        buf, pagesize=page, leftMargin=10 * mm, rightMargin=10 * mm,
        topMargin=12 * mm, bottomMargin=12 * mm, title="Furniture Inventory Report",
    )
    styles = getSampleStyleSheet()
    cell = ParagraphStyle("cell", parent=styles["BodyText"], fontSize=7, leading=9)
    head = ParagraphStyle("head", parent=cell, fontName="Helvetica-Bold", textColor=colors.white)

    cols = ["qr_code_id", "category", "location", "assigned_user", "purchase_date",
            "status", "condition", "quantity", "last_checked"]
    widths = [32, 25, 45, 42, 24, 22, 26, 12, 32]  # mm
    data = [[Paragraph(DISPLAY_NAMES[c], head) for c in cols]]
    for _, row in df.iterrows():
        data.append([Paragraph(escape(_fmt(row[c])), cell) for c in cols])

    table = Table(data, colWidths=[w * mm for w in widths], repeatRows=1)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F3A5F")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F2F5F9")]),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.lightgrey),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]))

    def footer(c, d):
        c.setFont("Helvetica", 7)
        c.drawRightString(page[0] - 10 * mm, 6 * mm, f"Page {d.page}")

    story = [
        Paragraph("Furniture Inventory Report", styles["Title"]),
        Paragraph(f"Generated {datetime.now():%Y-%m-%d %H:%M} &nbsp;|&nbsp; {len(df)} assets", styles["Normal"]),
        Spacer(1, 6 * mm),
        table,
    ]
    doc.build(story, onFirstPage=footer, onLaterPages=footer)
    return buf.getvalue()


# --------------------------------------------------------------------------- #
# PAGE 2 - Inventory & Dashboard
# --------------------------------------------------------------------------- #
def page_inventory() -> None:
    st.header("📊 Asset Inventory & Dashboard")
    if st.button("🔄 Refresh data"):
        load_inventory.clear()

    df = load_inventory()
    if df.empty:
        st.info("No assets registered yet. Use the scanner page to register your first item.")
        return

    # ---- Summary metrics -------------------------------------------------------
    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("Total assets", len(df))
    m2.metric("Good", int((df["condition"] == "Good").sum()))
    m3.metric("Fair", int((df["condition"] == "Fair").sum()))
    m4.metric("Needs repair", int((df["condition"] == "Needs Repair").sum()))
    m5.metric("Broken", int((df["condition"] == "Broken").sum()))

    # ---- Filters ----------------------------------------------------------------
    search = st.text_input("🔎 Search (QR ID, user, location, category)")
    f1, f2, f3, f4 = st.columns(4)
    sel_cat = f1.multiselect("Category", sorted(df["category"].dropna().unique()))
    sel_loc = f2.multiselect("Location", sorted(df["location"].dropna().unique()))
    sel_cond = f3.multiselect("Condition", sorted(df["condition"].dropna().unique()))
    sel_status = f4.multiselect("Status", STATUSES)

    mask = pd.Series(True, index=df.index)
    if search.strip():
        needle = search.strip().lower()
        haystack = df[["qr_code_id", "assigned_user", "location", "category"]].fillna("").astype(str)
        mask &= haystack.apply(lambda col: col.str.lower().str.contains(needle, regex=False)).any(axis=1)
    if sel_cat:
        mask &= df["category"].isin(sel_cat)
    if sel_loc:
        mask &= df["location"].isin(sel_loc)
    if sel_cond:
        mask &= df["condition"].isin(sel_cond)
    if sel_status:
        mask &= df["status"].isin(sel_status)
    view = df[mask]

    st.caption(f"Showing {len(view)} of {len(df)} assets")
    st.dataframe(
        view.rename(columns=DISPLAY_NAMES), hide_index=True,
        column_config={"Last Checked": st.column_config.DatetimeColumn(format="YYYY-MM-DD HH:mm")},
    )

    # ---- Exports (respect the active filters) ------------------------------------
    stamp = datetime.now().strftime("%Y%m%d_%H%M")
    d1, d2 = st.columns(2)
    d1.download_button(
        "⬇️ Download Excel", data=build_inventory_excel(view),
        file_name=f"furniture_inventory_{stamp}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    d2.download_button(
        "⬇️ Download PDF", data=build_inventory_pdf(view),
        file_name=f"furniture_inventory_{stamp}.pdf", mime="application/pdf",
    )

    with st.expander("Charts"):
        ca, cb = st.columns(2)
        ca.caption("Assets by condition")
        ca.bar_chart(df["condition"].value_counts())
        cb.caption("Assets by category")
        cb.bar_chart(df["category"].value_counts())


# --------------------------------------------------------------------------- #
# PAGE 3 - QR code generator (admin only)
# --------------------------------------------------------------------------- #
def admin_gate() -> bool:
    """Simple password gate for admin tools. Password lives in st.secrets['admin']['password']."""
    if st.session_state.get("is_admin"):
        return True
    try:
        expected = st.secrets["admin"]["password"]
    except Exception:
        st.error("Admin password not configured. Add an `[admin]` section with `password` to your secrets.")
        return False
    with st.form("admin_login"):
        pw = st.text_input("Admin password", type="password")
        if st.form_submit_button("Unlock"):
            if hmac.compare_digest(pw.encode(), str(expected).encode()):
                st.session_state["is_admin"] = True
                st.rerun()
            else:
                st.error("Wrong password.")
    return False


def qr_png_bytes(text: str) -> bytes:
    qr = qrcode.QRCode(error_correction=ERROR_CORRECT_M, box_size=10, border=1)
    qr.add_data(text)
    qr.make(fit=True)
    buf = io.BytesIO()
    qr.make_image(fill_color="black", back_color="white").save(buf, format="PNG")
    return buf.getvalue()


def build_qr_sheet_pdf(prefix: str, start: int, end: int, cols: int) -> bytes:
    """A4 grid of QR labels with the ID printed underneath each code."""
    page_w, page_h = A4
    margin = 10 * mm
    cell_w = (page_w - 2 * margin) / cols
    qr_size = min(cell_w - 8 * mm, 45 * mm)
    cell_h = qr_size + 13 * mm
    rows = max(1, int((page_h - 2 * margin) // cell_h))
    per_page = rows * cols

    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    c.setTitle(f"QR labels {prefix} {start}-{end}")

    for idx, number in enumerate(range(start, end + 1)):
        pos = idx % per_page
        if idx > 0 and pos == 0:
            c.showPage()
        row, col = divmod(pos, cols)
        x0 = margin + col * cell_w
        y_top = page_h - margin - row * cell_h
        code = f"{prefix}-{number:04d}"

        # Dashed cut line around each label
        c.setStrokeColor(colors.grey)
        c.setLineWidth(0.4)
        c.setDash(1, 2)
        c.rect(x0, y_top - cell_h, cell_w, cell_h)
        c.setDash()

        c.drawImage(
            ImageReader(io.BytesIO(qr_png_bytes(code))),
            x0 + (cell_w - qr_size) / 2, y_top - 3 * mm - qr_size, qr_size, qr_size,
        )
        c.setFillColor(colors.black)
        c.setFont("Helvetica-Bold", 9)
        c.drawCentredString(x0 + cell_w / 2, y_top - cell_h + 3.5 * mm, code)

    c.save()
    return buf.getvalue()


def page_qr_generator() -> None:
    st.header("🏷️ QR Code Generator (Admin)")
    if not admin_gate():
        return

    st.caption("Generates blank QR labels that contain only the ID. Details are added later by scanning.")
    c1, c2, c3, c4 = st.columns(4)
    prefix = c1.text_input("Prefix", value=DEFAULT_PREFIX).strip().upper()
    start = c2.number_input("From number", min_value=1, value=1, step=1)
    end = c3.number_input("To number", min_value=1, value=100, step=1)
    cols = c4.slider("Columns per row", 3, 5, 4)

    if not QR_ID_PATTERN.match(f"{prefix}-{int(start):04d}"):
        st.error("Prefix may only contain letters, digits, '-' and '_'.")
        return
    if end < start:
        st.error("'To' must be greater than or equal to 'From'.")
        return
    count = int(end - start + 1)
    if count > 1000:
        st.error("Please generate at most 1000 labels per PDF.")
        return

    first, last = f"{prefix}-{int(start):04d}", f"{prefix}-{int(end):04d}"
    st.write(f"**{count}** labels: `{first}` → `{last}`")
    st.image(qr_png_bytes(first), caption=f"Preview: {first}", width=150)

    if st.button("Generate PDF", type="primary"):
        with st.spinner("Building PDF..."):
            st.session_state["qr_pdf"] = (
                build_qr_sheet_pdf(prefix, int(start), int(end), cols),
                f"qr_labels_{prefix}_{int(start):04d}-{int(end):04d}.pdf",
            )
    if "qr_pdf" in st.session_state:
        data, name = st.session_state["qr_pdf"]
        st.download_button("⬇️ Download printable PDF", data=data, file_name=name, mime="application/pdf")

    if st.sidebar.button("🔒 Lock admin tools"):
        st.session_state["is_admin"] = False
        st.session_state.pop("qr_pdf", None)
        st.rerun()


# --------------------------------------------------------------------------- #
# Main
# --------------------------------------------------------------------------- #
def main() -> None:
    get_supabase()  # fail early with a friendly message if secrets are missing

    ss = st.session_state
    ss.setdefault("active_code", None)
    ss.setdefault("scan_nonce", 0)

    st.sidebar.title("🪑 Furniture Tracker")
    page = st.sidebar.radio(
        "Navigate",
        ["Scan & Register / Update", "Inventory & Dashboard", "QR Code Generator"],
    )
    inspector = st.sidebar.text_input("Inspector name", key="inspector", placeholder="Your name")

    if page == "Scan & Register / Update":
        page_scan(inspector)
    elif page == "Inventory & Dashboard":
        page_inventory()
    else:
        page_qr_generator()


if __name__ == "__main__":
    main()