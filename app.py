
import streamlit as st
import pandas as pd
import numpy as np
import datetime as dt
from dateutil import parser as dateparser
from openpyxl import load_workbook
from io import BytesIO

st.set_page_config(page_title="Cosmic Generator", layout="wide")

# ---------- Helpers ----------
@st.cache_data(show_spinner=False)
def load_workbook_bytes(b: bytes):
    return load_workbook(filename=BytesIO(b), data_only=True)

@st.cache_data(show_spinner=False)
def load_default_workbook():
    with open("data/cosmic_generator_v25.xlsx", "rb") as f:
        return f.read()

def get_sheet_df(wb, name):
    if name not in wb.sheetnames:
        return pd.DataFrame()
    ws = wb[name]
    data = ws.values
    rows = list(data)
    if not rows:
        return pd.DataFrame()
    header = rows[0]
    df = pd.DataFrame(rows[1:], columns=header)
    df = df.dropna(how="all")
    return df

def universal_day_number(d: dt.date, keep_master=True):
    s = d.strftime("%Y%m%d")
    n = sum(int(ch) for ch in s)
    if keep_master and n in (11, 22, 33):
        return n
    return 9 if n % 9 == 0 else n % 9

def weekday_name(d: dt.date):
    return d.strftime("%A")

def pick_primary(zone_element: str):
    if not isinstance(zone_element, str):
        return ""
    return zone_element.split("+")[0].strip()

def sun_sign_from_date(d: dt.date):
    # Western Tropical Sun sign by date ranges (approx, ignores year-specific edge cases)
    m, day = d.month, d.day
    if   (m==3 and day>=21) or (m==4 and day<=19): return "Aries"
    elif (m==4 and day>=20) or (m==5 and day<=20): return "Taurus"
    elif (m==5 and day>=21) or (m==6 and day<=20): return "Gemini"
    elif (m==6 and day>=21) or (m==7 and day<=22): return "Cancer"
    elif (m==7 and day>=23) or (m==8 and day<=22): return "Leo"
    elif (m==8 and day>=23) or (m==9 and day<=22): return "Virgo"
    elif (m==9 and day>=23) or (m==10 and day<=22): return "Libra"
    elif (m==10 and day>=23) or (m==11 and day<=21): return "Scorpio"
    elif (m==11 and day>=22) or (m==12 and day<=21): return "Sagittarius"
    elif (m==12 and day>=22) or (m==1 and day<=19): return "Capricorn"
    elif (m==1 and day>=20) or (m==2 and day<=18): return "Aquarius"
    elif (m==2 and day>=19) or (m==3 and day<=20): return "Pisces"
    return ""

# ---------- Load data (Sidebar) ----------
st.sidebar.title("Data Source")
st.sidebar.write("You can keep using the bundled workbook or upload a newer version anytime.")
uploaded = st.sidebar.file_uploader("Upload a Cosmic Generator workbook (.xlsx)", type=["xlsx"], help="Upload v26+ here to use it live. No data is stored.")
keep_master = st.sidebar.checkbox("Numerology: Keep master numbers (11/22/33)?", value=True)

if uploaded:
    raw = uploaded.getvalue()
    st.sidebar.success("Using uploaded workbook.")
else:
    raw = load_default_workbook()
    st.sidebar.info("Using bundled cosmic_generator_v25.xlsx")

wb = load_workbook_bytes(raw)

# Pull key sheets
def get_first_sheet_with(tokens):
    for name in wb.sheetnames:
        if all(t.lower() in name.lower() for t in tokens):
            return name
    return None

df_data         = get_sheet_df(wb, "Data")
df_audit        = get_sheet_df(wb, "AuditData")
df_elem_items   = get_sheet_df(wb, "Element_Items")
df_zones        = get_sheet_df(wb, "House_Zones")
df_rel          = get_sheet_df(wb, "Element_Relations")
df_pref         = get_sheet_df(wb, "Element_Preferences")
df_shapes       = get_sheet_df(wb, "Shape_Elements")
df_guide        = get_sheet_df(wb, "Activity_Day_Guide")
master_name     = get_first_sheet_with(["Master", "Correspondence"])
df_master       = get_sheet_df(wb, master_name) if master_name else pd.DataFrame()

# ---------- Global Inputs Tab ----------
tabs = st.tabs(["Inputs", "Life Audit", "Activity Timing", "House Zone Checker", "Items Browser"])

with tabs[0]:
    st.subheader("Inputs")
    colA, colB, colC = st.columns(3)
    with colA:
        birthdate = st.date_input("Birthdate", dt.date(1990,1,1))
    with colB:
        sign_source = st.selectbox("Sign Source", ["Sun (auto from birthdate)", "Moon (manual)", "Manual"])
    with colC:
        manual_sign = ""
        signs_list = sorted(df_data["Astrological Sign"].dropna().unique().tolist()) if "Astrological Sign" in df_data else []
        if sign_source == "Sun (auto from birthdate)":
            st.write("Sun Sign is auto-calculated from date.")
        else:
            manual_sign = st.selectbox("Choose Sign", [""] + signs_list)

    # Determine selected_sign
    if sign_source == "Sun (auto from birthdate)":
        selected_sign = sun_sign_from_date(birthdate)
    else:
        selected_sign = manual_sign

    # Pull Shape & Element for dashboard
    sel_shape = ""
    sel_element = ""
    if selected_sign and not df_master.empty:
        mrow = df_master.loc[df_master["Astrological Sign"]==selected_sign]
        if not mrow.empty:
            sel_shape = str(mrow.iloc[0].get("Shape","") or "")
            sel_element = str(mrow.iloc[0].get("Element","") or "")

    # Store in session state
    st.session_state["selected_sign"] = selected_sign
    st.session_state["selected_shape"] = sel_shape
    st.session_state["selected_element"] = sel_element

    # Show summary
    st.info(f"**Selected Sign:** {selected_sign or '—'}  |  **Element:** {sel_element or '—'}  |  **Shape:** {sel_shape or '—'}")
    if sel_element:
        best_z = df_pref.loc[df_pref["Element"]==sel_element]["Best Zones"]
        best_z_txt = best_z.iloc[0] if not best_z.empty else ""
        st.caption(f"Best Zones for {sel_element}: {best_z_txt}")

# ---------- Life Audit ----------
with tabs[1]:
    st.subheader("Life Audit - Conflicts & Remedies")
    selected_sign = st.session_state.get("selected_sign", "")
    if not selected_sign:
        st.warning("Pick a sign on the Inputs tab to run the audit.")
    # Buckets to check
    cols = st.columns(3)
    cat_items = {}
    categories = [
        ("Colours / Décor", "Avoid Household (strong)", "Avoid Household (mild)", "Colour", "Household Items"),
        ("Foods", "Avoid Foods (strong)", "Avoid Foods (mild)", "Foods", "Foods"),
        ("Crystals & Gemstones", "Avoid Crystals (strong)", "Avoid Crystals (mild)", "Primary Crystals", "Alternative Crystals / Gemstones"),
        ("Activities", "Avoid Activities (strong)", "Avoid Activities (mild)", "Favorable Activities", "Favorable Activities"),
        ("Elements", "Avoid Elements (strong)", "Avoid Elements (mild)", "Element", "Element"),
        ("People (Signs)", "Enemy Signs (strong)", "Enemy Signs (mild)", None, None),
    ]
    st.caption("Enter the item(s) you actually use/have. The audit flags STRONG / MILD conflicts and suggests remedies.")
    for i,(label,strong_col,mild_col,rem1,rem2) in enumerate(categories):
        with cols[i%3]:
            st.write(f"**{label}**")
            val = st.text_input(f"My {label}", key=f"la_{i}")
            cat_items[label] = val

    def audit_lookup(sign, col):
        if sign and col in df_audit.columns:
            row = df_audit.loc[df_audit["Astrological Sign"]==sign]
            if not row.empty:
                return str(row.iloc[0][col] or "")
        return ""

    results = []
    for (label,strong_col,mild_col,rem1,rem2) in categories:
        my = cat_items[label]
        strong_list = audit_lookup(selected_sign, strong_col)
        mild_list = audit_lookup(selected_sign, mild_col)
        conflict = "OK"
        why = "—"
        if my:
            my_low = my.lower()
            if any(tok.strip() and tok.strip().lower() in my_low for tok in str(strong_list).split(",")):
                conflict = "STRONG"
                why = strong_list
            elif any(tok.strip() and tok.strip().lower() in my_low for tok in str(mild_list).split(",")):
                conflict = "MILD"
                why = mild_list
        remedy_txt = ""
        row_d = df_data.loc[df_data["Astrological Sign"]==selected_sign]
        if not row_d.empty:
            if rem1 and rem1 in row_d.columns:
                r1 = str(row_d.iloc[0][rem1] or "")
                remedy_txt += r1
            if rem2 and rem2 in row_d.columns and str(row_d.iloc[0][rem2] or "") not in remedy_txt:
                if remedy_txt: remedy_txt += ", "
                remedy_txt += str(row_d.iloc[0][rem2] or "")
        results.append((label, my, conflict, why, remedy_txt))

    if selected_sign:
        st.write(pd.DataFrame(results, columns=["Category","My Item(s)","Conflict Level","Why flagged","Suggested Remedy"]))

# ---------- Activity Timing ----------
with tabs[2]:
    st.subheader("Activity Timing - Astrology × Numerology")
    activities = sorted(df_guide["Activity"].dropna().unique().tolist()) if "Activity" in df_guide else []
    activity = st.selectbox("Activity", activities)
    date = st.date_input("Planned Date", dt.date.today())
    if activity:
        wd = weekday_name(date)
        ud = universal_day_number(date, keep_master=st.session_state.get("keep_master", True))
        row = df_guide.loc[df_guide["Activity"]==activity].iloc[0]
        good_days = str(row.get("Good Days (Astrology)",""))
        avoid_days = str(row.get("Avoid Days (Astrology)",""))
        good_nums = str(row.get("Good Numbers (Numerology)",""))
        avoid_nums = str(row.get("Avoid Numbers (Numerology)",""))
        notes = str(row.get("Synergy Notes",""))
        afit = "Yes" if wd in good_days else ("No" if wd in avoid_days else "Maybe")
        nfit = "Yes" if str(ud) in good_nums else ("No" if str(ud) in avoid_nums else "Maybe")
        verdict = "Strong Cosmic Timing" if (afit=="Yes" and nfit=="Yes") else ("Weak / Reschedule Suggested" if ("No" in (afit,nfit)) else "Moderate Timing")
        st.write(f"**Weekday:** {wd}   |   **Universal Day:** {ud}")
        st.write(f"**Astrology Fit:** {afit}   |   **Numerology Fit:** {nfit}")
        st.write(f"**Overall Verdict:** {verdict}")
        if notes:
            st.info(notes)

# ---------- House Zone Checker ----------
with tabs[3]:
    st.subheader("House Zone Checker - 5 Elements + Space (with Shapes)")
    items = sorted(df_elem_items["Item Name"].dropna().unique().tolist()) if "Item Name" in df_elem_items else []
    zones = df_zones["Zone"].dropna().tolist() if "Zone" in df_zones else []
    shapes = df_shapes["Shape"].dropna().tolist() if "Shape" in df_shapes else []

    col1, col2, col3 = st.columns([1,1,1])
    with col1:
        item = st.selectbox("Item", items)
    with col2:
        zone = st.selectbox("House Zone", zones)
    with col3:
        sign_shape = st.session_state.get("selected_shape","")
        shape = st.selectbox("Shape (optional)", [""] + shapes, index=([""]+shapes).index(sign_shape) if sign_shape in shapes else 0)

    # Sign dashboard
    selected_sign = st.session_state.get("selected_sign","")
    selected_element = st.session_state.get("selected_element","")
    if selected_sign:
        best_z = df_pref.loc[df_pref["Element"]==selected_element]["Best Zones"]
        best_z_txt = best_z.iloc[0] if not best_z.empty else ""
        st.info(f"Sign: **{selected_sign}** | Element: **{selected_element or '—'}** | Best Zones: **{best_z_txt or '—'}**")

    verdict = ""
    remedy = ""
    if item and zone:
        item_elem = df_elem_items.loc[df_elem_items["Item Name"]==item].iloc[0]["Element"]
        zone_elem = df_zones.loc[df_zones["Zone"]==zone].iloc[0]["Primary Element"]
        zone_primary = pick_primary(zone_elem)
        try:
            rel_row = df_rel.loc[df_rel.iloc[:,0]==item_elem].iloc[0]
            rel = rel_row.get(zone_primary, "Neutral")
        except Exception:
            rel = "Neutral"
        verdict = rel

        best = df_pref.loc[df_pref["Element"]==item_elem]["Best Zones"]
        best_txt = best.iloc[0] if not best.empty else ""
        if zone == "Centre" and item_elem != "Space":
            remedy = "Keep centre open—relocate item to its best zones"
        elif str(rel).startswith("Avoid"):
            remedy = f"Move to: {best_txt}"
        elif rel == "Mild Avoid":
            remedy = "Balance with supportive colours/crystals of the zone element or relocate later"
        else:
            remedy = "OK / Supportive — keep as is"

        shape_elem = ""
        shape_rel = ""
        rec_shapes = ""
        if shape:
            srow = df_shapes.loc[df_shapes["Shape"]==shape]
            if not srow.empty:
                shape_elem = srow.iloc[0]["Element"]
                try:
                    rrow = df_rel.loc[df_rel.iloc[:,0]==shape_elem].iloc[0]
                    shape_rel = rrow.get(zone_primary, "")
                except Exception:
                    shape_rel = ""
        rec_shapes_list = df_shapes.loc[df_shapes["Element"]==zone_primary]["Shape"].dropna().tolist()
        if rec_shapes_list:
            rec_shapes = ", ".join(rec_shapes_list[:10])

        st.write(f"**Item Element:** {item_elem}   |   **Zone Element (Primary):** {zone_primary}")
        st.write(f"**Verdict:** {verdict}")
        st.write(f"**Remedy:** {remedy}")
        if shape:
            st.write(f"**Shape Element:** {shape_elem}   |   **Shape Verdict vs Zone:** {shape_rel}")
        if rec_shapes:
            st.caption(f"Recommended shapes for this zone: {rec_shapes}")

# ---------- Items Browser ----------
with tabs[4]:
    st.subheader("Element Items Browser")
    elems = ["All"] + sorted(df_elem_items["Element"].dropna().unique().tolist())
    cats = ["All"] + sorted(df_elem_items["Category"].dropna().unique().tolist())
    col1, col2 = st.columns(2)
    with col1:
        fe = st.selectbox("Element", elems, index=0)
    with col2:
        fc = st.selectbox("Category", cats, index=0)
    df_view = df_elem_items.copy()
    if fe != "All":
        df_view = df_view[df_view["Element"]==fe]
    if fc != "All":
        df_view = df_view[df_view["Category"]==fc]
    st.dataframe(df_view.reset_index(drop=True), use_container_width=True)
