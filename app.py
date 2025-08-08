
import streamlit as st
import pandas as pd
import numpy as np
import datetime as dt
from openpyxl import load_workbook
from io import BytesIO
import math
import swisseph as swe

st.set_page_config(page_title="Cosmic Generator", layout="wide")

# ---------- Helpers ----------
@st.cache_data(show_spinner=False)
def load_workbook_bytes(b: bytes):
    return load_workbook(filename=BytesIO(b), data_only=True)

@st.cache_data(show_spinner=False)
def load_default_workbook_bytes():
    # Try v21 fallback if bundled
    try:
        with open("data/cosmic_generator_v25.xlsx", "rb") as f:
            return f.read()
    except Exception:
        try:
            with open("data/cosmic_generator_v21.xlsx", "rb") as f:
                return f.read()
        except Exception:
            return b""

def get_sheet_df(wb, name):
    if not wb or name not in wb.sheetnames:
        return pd.DataFrame()
    ws = wb[name]
    rows = list(ws.values)
    if not rows:
        return pd.DataFrame()
    header = rows[0]
    df = pd.DataFrame(rows[1:], columns=header)
    return df.dropna(how="all")

def universal_day_number(d: dt.date, keep_master=True):
    s = d.strftime("%Y%m%d")
    n = sum(int(ch) for ch in s)
    if keep_master and n in (11, 22, 33):
        return n
    return 9 if n % 9 == 0 else n % 9

def weekday_name(d: dt.date):
    return d.strftime("%A")

def sun_sign_from_date(d: dt.date):
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

def sign_from_longitude(lon):
    signs = ["Aries","Taurus","Gemini","Cancer","Leo","Virgo","Libra","Scorpio","Sagittarius","Capricorn","Aquarius","Pisces"]
    idx = int(lon // 30) % 12
    return signs[idx]

def sun_moon_sign_exact(dt_local: dt.datetime, tz_offset_hours: float):
    # Convert local to UTC
    dt_utc = dt_local - dt.timedelta(hours=tz_offset_hours)
    # Julian Day UT
    jd = swe.julday(dt_utc.year, dt_utc.month, dt_utc.day,
                    dt_utc.hour + dt_utc.minute/60 + dt_utc.second/3600.0,
                    swe.GREG_CAL)
    # Swiss Ephemeris calc (tropical, default)
    lon_sun = swe.calc_ut(jd, swe.SUN)[0][0]
    lon_moon = swe.calc_ut(jd, swe.MOON)[0][0]
    return sign_from_longitude(lon_sun), lon_sun, sign_from_longitude(lon_moon), lon_moon

# ---------- Load data (Sidebar) ----------
st.sidebar.title("Data Source")
uploaded = st.sidebar.file_uploader("Upload a Cosmic Generator workbook (.xlsx)", type=["xlsx"], help="Upload your latest vXX to use it live. No data is stored.")
keep_master = st.sidebar.checkbox("Numerology: Keep master numbers (11/22/33)?", value=True)
st.session_state["keep_master"] = keep_master

if uploaded:
    raw = uploaded.getvalue()
    st.sidebar.success("Using uploaded workbook.")
else:
    raw = load_default_workbook_bytes()
    if raw:
        st.sidebar.info("Using bundled fallback workbook. For full features, upload the latest file.")
    else:
        st.sidebar.warning("No default workbook bundled. Please upload your workbook.")

wb = load_workbook_bytes(raw) if raw else None

def get_first_sheet_with(tokens):
    if not wb: return None
    for name in wb.sheetnames:
        if all(t.lower() in name.lower() for t in tokens):
            return name
    return None

# Pull key sheets (if available)
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

# ---------- Tabs ----------
tabs = st.tabs(["Inputs", "Life Audit", "Activity Timing", "House Zone Checker", "Items Browser"])

# ===== Inputs =====
with tabs[0]:
    st.subheader("Inputs (Birthdate, Moon, Sun)")

    colA, colB, colC = st.columns(3)
    with colA:
        birthdate = st.date_input("Birthdate", dt.date(1990,1,1))
        birthtime = st.time_input("Birth time (local)", dt.time(12,0))
        tz_offset = st.number_input("Timezone offset (hours vs UTC)", value=0.0, step=0.5, help="e.g., India = +5.5, London winter=0, summer=+1")
    with colB:
        st.markdown("**Moon sign**")
        moon_mode = st.selectbox("How to set Moon sign?", ["Calculate (exact) from birthdate+time", "Manual"])
        if moon_mode == "Manual":
            signs_list = sorted(df_data.get("Astrological Sign", pd.Series()).dropna().unique().tolist()) if not df_data.empty else ["Aries","Taurus","Gemini","Cancer","Leo","Virgo","Libra","Scorpio","Sagittarius","Capricorn","Aquarius","Pisces"]
            moon_sign = st.selectbox("Choose Moon Sign", signs_list)
        else:
            dt_local = dt.datetime.combine(birthdate, birthtime)
            sun_auto, sun_lon, moon_auto, moon_lon = sun_moon_sign_exact(dt_local, tz_offset)
            moon_sign = moon_auto
            st.caption(f"Computed Moon longitude: {moon_lon:.2f}° → **{moon_sign}**")
    with colC:
        st.markdown("**Sun sign**")
        sun_mode = st.selectbox("How to set Sun sign?", ["Calculate (exact) from birthdate+time", "Manual", "Simple by Birthdate Range"])
        if sun_mode == "Manual":
            signs_list = sorted(df_data.get("Astrological Sign", pd.Series()).dropna().unique().tolist()) if not df_data.empty else ["Aries","Taurus","Gemini","Cancer","Leo","Virgo","Libra","Scorpio","Sagittarius","Capricorn","Aquarius","Pisces"]
            sun_sign = st.selectbox("Choose Sun Sign", signs_list)
            sun_lon = None
        elif sun_mode == "Simple by Birthdate Range":
            sun_sign = sun_sign_from_date(birthdate)
            sun_lon = None
        else:
            dt_local = dt.datetime.combine(birthdate, birthtime)
            sun_auto, sun_lon, moon_auto_x, moon_lon_x = sun_moon_sign_exact(dt_local, tz_offset)
            sun_sign = sun_auto
            st.caption(f"Computed Sun longitude: {sun_lon:.2f}° → **{sun_sign}**")

    # Display both + choose which to use
    st.info(f"**Moon Sign:** {moon_sign}   |   **Sun Sign:** {sun_sign}")
    use_sign = st.radio("Use which sign for the rest of the app?", ["Moon", "Sun"], horizontal=True)

    # Pull details for both
    def sign_details(s):
        elem, shape = "",""
        if s and not df_master.empty and "Astrological Sign" in df_master.columns:
            r = df_master.loc[df_master["Astrological Sign"]==s]
            if not r.empty:
                elem = str(r.iloc[0].get("Element","") or "")
                shape = str(r.iloc[0].get("Shape","") or "")
        return elem, shape

    moon_elem, moon_shape = sign_details(moon_sign)
    sun_elem, sun_shape = sign_details(sun_sign)
    st.write(f"**Moon Element/Shape:** {moon_elem or '—'} / {moon_shape or '—'}")
    st.write(f"**Sun Element/Shape:** {sun_elem or '—'} / {sun_shape or '—'}")

    # Store selection
    selected_sign = moon_sign if use_sign=="Moon" else sun_sign
    selected_element = moon_elem if use_sign=="Moon" else sun_elem
    selected_shape = moon_shape if use_sign=="Moon" else sun_shape
    st.session_state["selected_sign"] = selected_sign
    st.session_state["selected_element"] = selected_element
    st.session_state["selected_shape"] = selected_shape

# ===== Life Audit =====
with tabs[1]:
    st.subheader("Life Audit – Conflicts & Remedies")
    selected_sign = st.session_state.get("selected_sign", "")
    if not selected_sign:
        st.warning("Set Moon/Sun on the Inputs tab first.")
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
        if not df_audit.empty and sign and col in df_audit.columns:
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
        if not df_data.empty:
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

# ===== Activity Timing =====
with tabs[2]:
    st.subheader("Activity Timing – Astrology × Numerology")
    keep_master = st.sidebar.session_state.get("keep_master", True)
    df = df_guide
    if df.empty:
        st.warning("Add Activity_Day_Guide sheet to your workbook for this tool.")
    else:
        activities = sorted(df["Activity"].dropna().unique().tolist())
        activity = st.selectbox("Activity", activities)
        date = st.date_input("Planned Date", dt.date.today())
        if activity:
            wd = date.strftime("%A")
            srow = df.loc[df["Activity"]==activity].iloc[0]
            ud = universal_day_number(date, keep_master=keep_master)
            afit = "Yes" if wd in str(srow.get("Good Days (Astrology)","")) else ("No" if wd in str(srow.get("Avoid Days (Astrology)","")) else "Maybe")
            nfit = "Yes" if str(ud) in str(srow.get("Good Numbers (Numerology)","")) else ("No" if str(ud) in str(srow.get("Avoid Numbers (Numerology)","")) else "Maybe")
            verdict = "Strong Cosmic Timing" if (afit=="Yes" and nfit=="Yes") else ("Weak / Reschedule Suggested" if ("No" in (afit,nfit)) else "Moderate Timing")
            st.write(f"**Weekday:** {wd}   |   **Universal Day:** {ud}")
            st.write(f"**Astrology Fit:** {afit}   |   **Numerology Fit:** {nfit}")
            st.write(f"**Overall Verdict:** {verdict}")
            notes = str(srow.get("Synergy Notes",""))
            if notes: st.info(notes)

# ===== House Zone Checker =====
with tabs[3]:
    st.subheader("House Zone Checker – 5 Elements + Space (with Shapes)")
    if df_elem_items.empty or df_zones.empty or df_rel.empty or df_pref.empty:
        st.warning("Please upload a workbook that includes Element_Items, House_Zones, Element_Relations, Element_Preferences, Shape_Elements.")
    else:
        items = sorted(df_elem_items["Item Name"].dropna().unique().tolist())
        zones = df_zones["Zone"].dropna().tolist()
        shapes = df_shapes["Shape"].dropna().tolist() if not df_shapes.empty else []

        selected_sign = st.session_state.get("selected_sign","")
        selected_element = st.session_state.get("selected_element","")
        selected_shape = st.session_state.get("selected_shape","")

        col1, col2, col3 = st.columns([1,1,1])
        with col1: item = st.selectbox("Item", items)
        with col2: zone = st.selectbox("House Zone", zones)
        with col3:
            choices = [""] + shapes
            idx = choices.index(selected_shape) if selected_shape in choices else 0
            shape = st.selectbox("Shape (optional)", choices, index=idx)

        if selected_sign and selected_element and not df_pref.empty:
            best_z = df_pref.loc[df_pref["Element"]==selected_element]["Best Zones"]
            best_z_txt = best_z.iloc[0] if not best_z.empty else ""
            st.info(f"Using Sign: **{selected_sign}** | Element: **{selected_element}** | Best Zones: **{best_z_txt}**")

        if item and zone:
            item_elem = df_elem_items.loc[df_elem_items["Item Name"]==item].iloc[0]["Element"]
            zone_elem = df_zones.loc[df_zones["Zone"]==zone].iloc[0]["Primary Element"]
            zone_primary = zone_elem.split("+")[0].strip() if isinstance(zone_elem, str) else ""

            rel = "Neutral"
            try:
                r = df_rel[df_rel.iloc[:,0]==item_elem]
                if not r.empty:
                    rel = r.iloc[0].get(zone_primary, "Neutral")
            except Exception: pass

            best = df_pref.loc[df_pref["Element"]==item_elem]["Best Zones"]
            best_txt = best.iloc[0] if not best.empty else ""
            if zone == "Centre" and item_elem != "Space":
                remedy = "Keep centre open—relocate item to its best zones"
            elif isinstance(rel,str) and rel.startswith("Avoid"):
                remedy = f"Move to: {best_txt}"
            elif rel == "Mild Avoid":
                remedy = "Balance with supportive colours/crystals of the zone element or relocate later"
            else:
                remedy = "OK / Supportive — keep as is"

            # Shapes
            shape_elem = ""
            shape_rel = ""
            rec_shapes = ""
            if shape and not df_shapes.empty:
                srow = df_shapes.loc[df_shapes["Shape"]==shape]
                if not srow.empty:
                    shape_elem = srow.iloc[0]["Element"]
                    try:
                        rr = df_rel[df_rel.iloc[:,0]==shape_elem].iloc[0]
                        shape_rel = rr.get(zone_primary, "")
                    except Exception:
                        shape_rel = ""
                rec_list = df_shapes.loc[df_shapes["Element"]==zone_primary]["Shape"].dropna().tolist()
                if rec_list:
                    rec_shapes = ", ".join(rec_list[:10])

            st.write(f"**Item Element:** {item_elem}   |   **Zone Element (Primary):** {zone_primary}")
            st.write(f"**Verdict:** {rel}")
            st.write(f"**Remedy:** {remedy}")
            if shape:
                st.write(f"**Shape Element:** {shape_elem}   |   **Shape Verdict vs Zone:** {shape_rel}")
            if rec_shapes:
                st.caption(f"Recommended shapes for this zone: {rec_shapes}")

# ===== Items Browser =====
with tabs[4]:
    st.subheader("Element Items Browser")
    if df_elem_items.empty:
        st.warning("Upload a workbook with Element_Items sheet to browse items.")
    else:
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
