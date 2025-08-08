
import streamlit as st
import pandas as pd
import numpy as np
import datetime as dt
from io import BytesIO
from openpyxl import load_workbook
import math

st.set_page_config(page_title="Cosmic Generator", layout="wide")

# ---------- Swiss Ephemeris (optional) ----------
HAVE_SWE = False
try:
    import swisseph as swe
    HAVE_SWE = True
except Exception:
    HAVE_SWE = False

# ---------- Helpers ----------
@st.cache_data(show_spinner=False)
def load_workbook_bytes(b: bytes):
    return load_workbook(filename=BytesIO(b), data_only=True)

@st.cache_data(show_spinner=False)
def load_default_workbook():
    try:
        with open("data/cosmic_generator_v25.xlsx", "rb") as f:
            return f.read()
    except Exception:
        return b""

def get_sheet_df(wb, name):
    try:
        if not wb or name not in wb.sheetnames:
            return pd.DataFrame()
        ws = wb[name]
        rows = list(ws.values)
        if not rows:
            return pd.DataFrame()
        header = rows[0]
        df = pd.DataFrame(rows[1:], columns=header)
        return df.dropna(how="all")
    except Exception:
        return pd.DataFrame()

ZODIAC = ["Aries","Taurus","Gemini","Cancer","Leo","Virgo","Libra","Scorpio","Sagittarius","Capricorn","Aquarius","Pisces"]

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

def _deg_to_rad(x): 
    return x * math.pi / 180.0
def _rev(x): 
    return x % 360.0

def moon_longitude_approx_noon_utc(d: dt.date):
    # Approximate Moon longitude at 12:00 UTC for given date (fallback)
    year, month, day = d.year, d.month, d.day
    hour = 12.0
    if month <= 2:
        year -= 1
        month += 12
    A = int(year/100)
    B = 2 - A + int(A/4)
    JD = int(365.25*(year + 4716)) + int(30.6001*(month + 1)) + (day + hour/24.0) + B - 1524.5
    T = (JD - 2451545.0)/36525.0
    Lp = _rev(218.3164477 + 481267.88123421*T - 0.0015786*T*T)
    D  = _rev(297.8501921 + 445267.1114034*T - 0.0018819*T*T)
    M  = _rev(357.5291092 + 35999.0502909*T - 0.0001536*T*T)
    Mp = _rev(134.9633964 + 477198.8675055*T + 0.0087414*T*T)
    F  = _rev(93.2720950 + 483202.0175233*T - 0.0036539*T*T)
    Dr = _deg_to_rad(D); Mr = _deg_to_rad(M); Mpr = _deg_to_rad(Mp); Fr = _deg_to_rad(F)
    lon = (Lp
        + 6.289 * math.sin(Mpr)
        + 1.274 * math.sin(2*Dr - Mpr)
        + 0.658 * math.sin(2*Dr)
        + 0.214 * math.sin(2*Mpr)
        - 0.186 * math.sin(Mr)
        - 0.114 * math.sin(2*Fr)
        + 0.059 * math.sin(2*Dr - 2*Mpr)
        + 0.057 * math.sin(2*Dr - Mr - Mpr)
        + 0.053 * math.sin(2*Dr + Mpr)
        + 0.046 * math.sin(2*Dr - Mr)
        + 0.041 * math.sin(Mr + Mpr))
    return _rev(lon)

def moon_sign_exact(birthdate: dt.date, birthtime: dt.time, tz_offset: float):
    # Return (sign, longitude, used_exact_bool)
    try:
        import swisseph as swe  # local import to avoid failing if package missing
        hour_local = birthtime.hour + birthtime.minute/60 + birthtime.second/3600
        hour_utc = hour_local - float(tz_offset)
        jd = swe.julday(birthdate.year, birthdate.month, birthdate.day, hour_utc)
        pos = swe.calc_ut(jd, swe.MOON)[0]
        lon = float(pos[0])
        idx = int(lon // 30) % 12
        return ZODIAC[idx], lon, True
    except Exception:
        lon = moon_longitude_approx_noon_utc(birthdate)
        idx = int(lon // 30) % 12
        return ZODIAC[idx], lon, False

def universal_day_number(d: dt.date, keep_master=True):
    s = d.strftime("%Y%m%d")
    n = sum(int(ch) for ch in s)
    if keep_master and n in (11, 22, 33):
        return n
    return 9 if n % 9 == 0 else n % 9

def weekday_name(d: dt.date):
    return d.strftime("%A")

def safe_unique_list(series):
    try:
        return sorted([x for x in series.dropna().unique().tolist() if x])
    except Exception:
        return []

# ---------------- Sidebar: Data ----------------
st.sidebar.title("Data Source")
uploaded = st.sidebar.file_uploader(
    "Upload a Cosmic Generator workbook (.xlsx)", 
    type=["xlsx"], 
    help="Upload your latest vXX here to use it live. No data is stored."
)
keep_master = st.sidebar.checkbox("Numerology: Keep master numbers (11/22/33)?", value=True)
st.session_state["keep_master"] = keep_master

raw = uploaded.getvalue() if uploaded else load_default_workbook()
wb = load_workbook_bytes(raw) if raw else None

# Load sheets
df_data         = get_sheet_df(wb, "Data")
df_audit        = get_sheet_df(wb, "AuditData")
df_elem_items   = get_sheet_df(wb, "Element_Items")
df_zones        = get_sheet_df(wb, "House_Zones")
df_rel          = get_sheet_df(wb, "Element_Relations")
df_pref         = get_sheet_df(wb, "Element_Preferences")
df_shapes       = get_sheet_df(wb, "Shape_Elements")
df_guide        = get_sheet_df(wb, "Activity_Day_Guide")

# Try find Master Correspondence
master_name = None
if wb:
    for name in wb.sheetnames:
        if "Master" in name and "Correspondence" in name:
            master_name = name
            break
df_master = get_sheet_df(wb, master_name) if master_name else pd.DataFrame()

signs_list = safe_unique_list(df_data["Astrological Sign"]) if not df_data.empty and "Astrological Sign" in df_data.columns else ZODIAC

# ---------------- Tabs ----------------
tabs = st.tabs(["Inputs", "Life Audit", "Activity Timing", "House Zone Checker", "Items Browser"])

# ===== Inputs =====
with tabs[0]:
    st.subheader("Inputs")
    birthdate = st.date_input("Birthdate", dt.date(1990,1,1))
    birthtime = st.time_input("Birth time (local)", dt.time(12,0))

    # Timezone dropdown
    tz_options = [
        ("UTC (0)", 0.0),
        ("London Winter (0)", 0.0),
        ("London Summer (+1)", 1.0),
        ("Paris / Berlin (+1 winter / +2 summer)", 1.0),
        ("Athens (+2)", 2.0),
        ("Moscow (+3)", 3.0),
        ("Dubai (+4)", 4.0),
        ("India IST (+5.5)", 5.5),
        ("Bangladesh (+6)", 6.0),
        ("Thailand (+7)", 7.0),
        ("Singapore / Hong Kong (+8)", 8.0),
        ("Japan (+9)", 9.0),
        ("Sydney (+10)", 10.0),
        ("Auckland (+12)", 12.0),
        ("New York (-5 winter / -4 summer)", -5.0),
        ("Chicago (-6 winter / -5 summer)", -6.0),
        ("Denver (-7 winter / -6 summer)", -7.0),
        ("Los Angeles (-8 winter / -7 summer)", -8.0),
        ("Mexico City (-6)", -6.0),
        ("Sao Paulo (-3)", -3.0),
        ("Johannesburg (+2)", 2.0),
        ("Cairo (+2)", 2.0),
        ("Tehran (+3.5)", 3.5),
        ("Kathmandu (+5.75)", 5.75),
        ("Yangon (+6.5)", 6.5),
    ]
    tz_labels = [x[0] for x in tz_options]
    tz_map = {k:v for k,v in tz_options}
    tz_choice = st.selectbox("Timezone (pick closest)", tz_labels, index=7)  # default IST
    tz_offset = tz_map.get(tz_choice, 0.0)

    # Compute signs
    sun_sign = sun_sign_from_date(birthdate) or ""
    moon_sign, moon_lon, used_exact = moon_sign_exact(birthdate, birthtime, tz_offset)

    st.info(f"Sun: {sun_sign or '—'}   |   Moon: {moon_sign or '—'} {'(exact)' if used_exact else '(approx)'}   |   TZ: {tz_choice}")

    # Choose which to use globally
    use_choice = st.radio("Use which sign across the app?", ["Sun", "Moon"], horizontal=True)
    selected_sign = sun_sign if use_choice == "Sun" else moon_sign
    st.session_state["selected_sign"] = selected_sign

    # Show ALL info for selected sign from Data tab
    if not df_data.empty and "Astrological Sign" in df_data.columns and selected_sign:
        row = df_data.loc[df_data["Astrological Sign"] == selected_sign]
        if not row.empty:
            info = row.iloc[0].to_dict()
            info_clean = {k: v for k,v in info.items() if pd.notna(v) and str(v) != ""}
            df_show = pd.DataFrame(list(info_clean.items()), columns=["Field","Value"])
            st.write("**Sign Details (from Data sheet):**")
            st.dataframe(df_show, use_container_width=True, hide_index=True)
            # Cache for defaults elsewhere if present
            st.session_state["selected_element"] = str(info_clean.get("Element",""))
            st.session_state["selected_shape"] = str(info_clean.get("Shape",""))
    else:
        st.warning("Data sheet missing or sign not found. Upload a generator with a 'Data' sheet to see full details.")

# ===== Life Audit =====
with tabs[1]:
    st.subheader("Life Audit - Conflicts & Remedies")
    selected_sign = st.session_state.get("selected_sign","")
    if not selected_sign:
        st.warning("Go to Inputs and pick Sun/Moon first.")
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
        if sign and not df_audit.empty and "Astrological Sign" in df_audit.columns and col in df_audit.columns:
            row = df_audit.loc[df_audit["Astrological Sign"]==sign]
            if not row.empty:
                return str(row.iloc[0][col] or "")
        return ""

    results = []
    for (label,strong_col,mild_col,rem1,rem2) in categories:
        my = cat_items[label]
        strong_list = audit_lookup(selected_sign, strong_col)
        mild_list = audit_lookup(selected_sign, mild_col)
        conflict, why = "OK", "—"
        if my:
            my_low = my.lower()
            strong_tokens = [t.strip().lower() for t in str(strong_list).split(",") if t.strip()]
            mild_tokens   = [t.strip().lower() for t in str(mild_list).split(",") if t.strip()]
            if any(tok in my_low for tok in strong_tokens):
                conflict = "STRONG"; why = strong_list
            elif any(tok in my_low for tok in mild_tokens):
                conflict = "MILD"; why = mild_list
        remedy_txt = ""
        if not df_data.empty and "Astrological Sign" in df_data.columns:
            row_d = df_data.loc[df_data["Astrological Sign"]==selected_sign]
            if not row_d.empty:
                if rem1 and rem1 in row_d.columns:
                    r1 = str(row_d.iloc[0][rem1] or "")
                    remedy_txt += r1
                if rem2 and rem2 in row_d.columns:
                    r2 = str(row_d.iloc[0][rem2] or "")
                    if r2 and r2 not in remedy_txt:
                        remedy_txt += (", " if remedy_txt else "") + r2
        results.append((label, my, conflict, why, remedy_txt))

    if selected_sign:
        st.dataframe(pd.DataFrame(results, columns=["Category","My Item(s)","Conflict Level","Why flagged","Suggested Remedy"]), use_container_width=True)

# ===== Activity Timing =====
with tabs[2]:
    st.subheader("Activity Timing - Astrology x Numerology")
    if df_guide.empty:
        st.warning("Activity_Day_Guide sheet not found in workbook.")
    else:
        activities = safe_unique_list(df_guide["Activity"]) if "Activity" in df_guide.columns else []
        activity = st.selectbox("Activity", activities) if activities else ""
        date = st.date_input("Planned Date", dt.date.today())
        if activity:
            row = df_guide.loc[df_guide["Activity"]==activity]
            if row.empty:
                st.warning("Selected activity not found in guide."); 
            else:
                row = row.iloc[0]
                wd = weekday_name(date)
                ud = universal_day_number(date, keep_master=st.session_state.get("keep_master", True))
                good_days = str(row.get("Good Days (Astrology)",""))
                avoid_days = str(row.get("Avoid Days (Astrology)",""))
                good_nums = str(row.get("Good Numbers (Numerology)",""))
                avoid_nums = str(row.get("Avoid Numbers (Numerrology)","")) if "Avoid Numbers (Numerrology)" in row.index else str(row.get("Avoid Numbers (Numerology)",""))
                notes = str(row.get("Synergy Notes",""))
                afit = "Yes" if wd in good_days else ("No" if wd in avoid_days else "Maybe")
                nfit = "Yes" if str(ud) in good_nums else ("No" if str(ud) in avoid_nums else "Maybe")
                verdict = "Strong Cosmic Timing" if (afit=="Yes" and nfit=="Yes") else ("Weak / Reschedule Suggested" if ("No" in (afit,nfit)) else "Moderate Timing")
                st.write(f"Weekday: {wd}   |   Universal Day: {ud}")
                st.write(f"Astrology Fit: {afit}   |   Numerology Fit: {nfit}")
                st.write(f"Overall Verdict: {verdict}")
                if notes:
                    st.info(notes)

# ===== House Zone Checker =====
with tabs[3]:
    st.subheader("House Zone Checker - 5 Elements + Space (with Shapes)")
    if df_elem_items.empty or df_zones.empty or df_rel.empty:
        st.warning("Missing one of: Element_Items, House_Zones, Element_Relations sheets.")
    else:
        items = safe_unique_list(df_elem_items["Item Name"]) if "Item Name" in df_elem_items.columns else []
        zones = safe_unique_list(df_zones["Zone"]) if "Zone" in df_zones.columns else []
        shapes = safe_unique_list(df_shapes["Shape"]) if "Shape" in df_shapes.columns else []
        col1, col2, col3 = st.columns(3)
        with col1:
            item = st.selectbox("Item", items) if items else ""
        with col2:
            zone = st.selectbox("House Zone", zones) if zones else ""
        with col3:
            sign_shape = st.session_state.get("selected_shape","")
            choices = [""] + shapes
            idx = choices.index(sign_shape) if sign_shape in choices else 0
            shape = st.selectbox("Shape (optional)", choices, index=idx)

        if item and zone:
            try:
                item_elem = df_elem_items.loc[df_elem_items["Item Name"]==item].iloc[0]["Element"]
            except Exception:
                item_elem = ""
            try:
                zone_elem = df_zones.loc[df_zones["Zone"]==zone].iloc[0]["Primary Element"]
            except Exception:
                zone_elem = ""
            zone_primary = str(zone_elem).split("+")[0].strip() if zone_elem else ""

            rel = "Neutral"
            try:
                row = df_rel[df_rel.iloc[:,0]==item_elem]
                if not row.empty and zone_primary in row.columns:
                    rel = row.iloc[0][zone_primary]
            except Exception:
                pass

            remedy = "OK / Supportive — keep as is"
            if zone == "Centre" and item_elem != "Space":
                remedy = "Keep centre open — relocate item to its best zones"
            elif isinstance(rel, str) and rel.startswith("Avoid"):
                if not df_pref.empty:
                    rowp = df_pref.loc[df_pref["Element"]==item_elem]
                    best = str(rowp.iloc[0]["Best Zones"]) if not rowp.empty else ""
                    remedy = f"Move to: {best}"
                else:
                    remedy = "Move to a better-suited zone for the item's element"
            elif rel == "Mild Avoid":
                remedy = "Balance with supportive colours/crystals of the zone element or relocate later"

            shape_elem = ""
            shape_rel = ""
            rec_shapes = ""
            if shape and not df_shapes.empty:
                srow = df_shapes.loc[df_shapes["Shape"]==shape]
                if not srow.empty:
                    shape_elem = str(srow.iloc[0]["Element"])
                    try:
                        rrow = df_rel[df_rel.iloc[:,0]==shape_elem]
                        if not rrow.empty and zone_primary in rrow.columns:
                            shape_rel = rrow.iloc[0][zone_primary]
                    except Exception:
                        shape_rel = ""
                rec_list = df_shapes.loc[df_shapes["Element"]==zone_primary]["Shape"].dropna().tolist() if zone_primary else []
                if rec_list:
                    rec_shapes = ", ".join(rec_list[:10])

            st.write(f"Item Element: {item_elem or '—'}   |   Zone Element (Primary): {zone_primary or '—'}")
            st.write(f"Verdict: {rel}")
            st.write(f"Remedy: {remedy}")
            if shape:
                st.write(f"Shape Element: {shape_elem or '—'}   |   Shape Verdict vs Zone: {shape_rel or '—'}")
            if rec_shapes:
                st.caption(f"Recommended shapes for this zone: {rec_shapes}")

# ===== Items Browser =====
with tabs[4]:
    st.subheader("Element Items Browser")
    if df_elem_items.empty:
        st.warning("Element_Items sheet missing.")
    else:
        elems = ["All"] + safe_unique_list(df_elem_items["Element"]) if "Element" in df_elem_items.columns else ["All"]
        cats  = ["All"] + safe_unique_list(df_elem_items["Category"]) if "Category" in df_elem_items.columns else ["All"]
        col1, col2 = st.columns(2)
        with col1:
            fe = st.selectbox("Element", elems, index=0)
        with col2:
            fc = st.selectbox("Category", cats, index=0)
        df_view = df_elem_items.copy()
        if "Element" in df_view.columns and fe != "All":
            df_view = df_view[df_view["Element"]==fe]
        if "Category" in df_view.columns and fc != "All":
            df_view = df_view[df_view["Category"]==fc]
        st.dataframe(df_view.reset_index(drop=True), use_container_width=True)
