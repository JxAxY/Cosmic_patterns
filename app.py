
import streamlit as st
import pandas as pd
import numpy as np
import datetime as dt
from io import BytesIO
from openpyxl import load_workbook
import math
import re

st.set_page_config(page_title="Cosmic Generator", layout="wide")

# --- Optional Swiss Ephemeris ---
try:
    import swisseph as swe
    HAVE_SWE = True
except Exception:
    HAVE_SWE = False

# --- Helpers ---
@st.cache_data(show_spinner=False)
def load_workbook_bytes(b: bytes):
    return load_workbook(filename=BytesIO(b), data_only=True)

@st.cache_data(show_spinner=False)
def load_default_workbook():
    with open("data/cosmic_generator_v25.xlsx", "rb") as f:
        return f.read()

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

def _rev(x): return x % 360.0
def moon_longitude_approx_noon_utc(d: dt.date):
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
    Dr = math.radians(D); Mr = math.radians(M); Mpr = math.radians(Mp); Fr = math.radians(F)
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
    try:
        if HAVE_SWE:
            hour_local = birthtime.hour + birthtime.minute/60 + birthtime.second/3600
            hour_utc = hour_local - float(tz_offset)
            jd = swe.julday(birthdate.year, birthdate.month, birthdate.day, hour_utc)
            pos = swe.calc_ut(jd, swe.MOON)[0]
            lon = float(pos[0])
            idx = int(lon // 30) % 12
            return ZODIAC[idx], lon, True
        lon = moon_longitude_approx_noon_utc(birthdate)
        idx = int(lon // 30) % 12
        return ZODIAC[idx], lon, False
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

# --- Sidebar ---
st.sidebar.title("Data Source")
uploaded = st.sidebar.file_uploader("Upload a Cosmic Generator workbook (.xlsx)", type=["xlsx"])
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

# Tabs
tabs = st.tabs(["Inputs", "Life Audit", "Activity Timing", "House Zone Checker", "Items Browser"])

# ===== Inputs =====
with tabs[0]:
    st.subheader("Inputs")
    birthdate = st.date_input("Birthdate", dt.date(1990,1,1))
    birthtime = st.time_input("Birth time (local)", dt.time(12,0))

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
    tz_choice = st.selectbox("Timezone (pick closest)", tz_labels, index=7)
    tz_offset = tz_map.get(tz_choice, 0.0)

    sun_sign = sun_sign_from_date(birthdate) or ""
    moon_sign, moon_lon, used_exact = moon_sign_exact(birthdate, birthtime, tz_offset)
    st.info(f"Sun: {sun_sign or '—'} | Moon: {moon_sign or '—'} {'(exact)' if used_exact else '(approx)'} | TZ: {tz_choice}")

    use_choice = st.radio("Use which sign across the app?", ["Sun","Moon"], horizontal=True)
    selected_sign = sun_sign if use_choice=="Sun" else moon_sign
    st.session_state["selected_sign"] = selected_sign

    if not df_data.empty and "Astrological Sign" in df_data.columns and selected_sign:
        row = df_data.loc[df_data["Astrological Sign"]==selected_sign]
        if not row.empty:
            info = row.iloc[0].to_dict()
            info_clean = {k: v for k,v in info.items() if pd.notna(v) and str(v) != ""}
            df_show = pd.DataFrame(list(info_clean.items()), columns=["Field","Value"])
            st.write("**Sign Details (from Data sheet):**")
            st.dataframe(df_show, use_container_width=True, hide_index=True)

# ===== Life Audit (v9 matching) =====
with tabs[1]:
    st.subheader("Life Audit – Conflicts & Remedies")
    selected_sign = st.session_state.get("selected_sign","")
    if not selected_sign:
        st.warning("Go to Inputs and choose Sun or Moon first.")
    st.caption("Comma-separated lists. Smarter matching with phrases and word boundaries.")

    config = [
        ("Colours / Décor", "Avoid Household (strong)", "Avoid Household (mild)", "Colour", "Household Items", "general"),
        ("Foods", "Avoid Foods (strong)", "Avoid Foods (mild)", "Foods", "Foods", "general"),
        ("Crystals & Gemstones", "Avoid Crystals (strong)", "Avoid Crystals (mild)", "Primary Crystals", "Alternative Crystals / Gemstones", "general"),
        ("Activities", "Avoid Activities (strong)", "Avoid Activities (mild)", "Favorable Activities", "Favorable Activities", "names_only"),
        ("Elements", "Avoid Elements (strong)", "Avoid Elements (mild)", "Element", "Element", "general"),
        ("People (Signs)", "Enemy Signs (strong)", "Enemy Signs (mild)", None, None, "names_only"),
    ]

    def tokenize(user_text, mode):
        if not user_text: return []
        # split only on commas/newlines; keep phrases
        parts = re.split(r"[,\\n]", user_text)
        tokens = []
        for p in parts:
            t = re.sub(r"\s+", " ", p.strip().lower())
            if t:
                tokens.append(t)
        return tokens

    def normalize_list(csv_str):
        if not csv_str: return []
        terms = [re.sub(r"\s+", " ", x.strip().lower()) for x in str(csv_str).split(",")]
        return [t for t in terms if t]

    def match_token(tok, strong_terms, mild_terms):
        # exact first
        if tok in strong_terms: return "STRONG", tok
        if tok in mild_terms:   return "MILD", tok
        # word-boundary contains: token in term or term in token
        for term in strong_terms:
            if re.search(rf"\\b{re.escape(tok)}\\b", term) or re.search(rf"\\b{re.escape(term)}\\b", tok):
                return "STRONG", term
        for term in mild_terms:
            if re.search(rf"\\b{re.escape(tok)}\\b", term) or re.search(rf"\\b{re.escape(term)}\\b", tok):
                return "MILD", term
        # fallback substring (less strict)
        for term in strong_terms:
            if tok in term or term in tok:
                return "STRONG", term
        for term in mild_terms:
            if tok in term or term in tok:
                return "MILD", term
        return "OK", ""

    results_rows = []
    summary_rows = []

    for (label, col_strong, col_mild, rem1, rem2, mode) in config:
        user_text = st.text_area(f"My {label}", key=f"la9_{label}", height=60, placeholder="e.g., blue, rose gold, marble" if mode!="names_only" else "e.g., Aries, Scorpio")
        tokens = tokenize(user_text, mode)

        if df_audit.empty or "Astrological Sign" not in df_audit.columns:
            st.caption(f"_No rules found for {label} (AuditData missing)._")
            continue
        row = df_audit.loc[df_audit["Astrological Sign"]==selected_sign]
        if row.empty:
            st.caption(f"_No rules for {selected_sign} in {label}._")
            continue

        strong_terms = normalize_list(row.iloc[0].get(col_strong, ""))
        mild_terms   = normalize_list(row.iloc[0].get(col_mild, ""))

        strong_hits = mild_hits = 0
        if tokens:
            for tok in tokens:
                verdict, src = match_token(tok, strong_terms, mild_terms)
                if verdict=="STRONG": strong_hits += 1
                elif verdict=="MILD": mild_hits += 1
                results_rows.append((label, tok, verdict, src))
        else:
            results_rows.append((label, "", "OK", ""))

        fixes = []
        # Pull remedies from Data sheet if present
        if not df_data.empty and "Astrological Sign" in df_data.columns:
            drow = df_data.loc[df_data["Astrological Sign"]==selected_sign]
            if not drow.empty:
                drow = drow.iloc[0]
                if strong_hits>0:
                    if label=="Colours / Décor" and "Colour" in drow:
                        fixes.append(f"Switch to favourable colours: {drow['Colour']}")
                    if label=="Foods" and "Foods" in drow:
                        fixes.append(f"Prioritize: {drow['Foods']}")
                    if label=="Crystals & Gemstones":
                        if "Primary Crystals" in drow: fixes.append(f"Carry/wear: {drow['Primary Crystals']}")
                        if "Alternative Crystals / Gemstones" in drow: fixes.append(f"Alternate set: {drow['Alternative Crystals / Gemstones']}")
                    if label=="Activities" and "Favorable Activities" in drow:
                        fixes.append(f"Swap to: {drow['Favorable Activities']}")
                    if label=="Elements" and "Element" in drow:
                        fixes.append(f"Emphasize {drow['Element']} items")
                    if label=="People (Signs)":
                        fixes.append("Reduce strong enemy-sign dynamics; choose neutral/shared-element activities")
                elif mild_hits>0:
                    if label in ("Colours / Décor","Elements","Crystals & Gemstones"):
                        fixes.append("Balance with supportive colours/crystals of your element")
                    if label=="Foods":
                        fixes.append("Moderate intake; avoid combos with other flagged foods")
                    if label=="Activities":
                        fixes.append("Do on a favourable weekday/number to offset mild conflicts")
                    if label=="People (Signs)":
                        fixes.append("Choose neutral settings; keep interactions short")
        fixes = fixes[:3]
        summary_rows.append((label, strong_hits, mild_hits, ", ".join(fixes)))

    # Render results
    colA, colB = st.columns([2,1])
    with colA:
        if results_rows:
            df_res = pd.DataFrame(results_rows, columns=["Category","Token","Conflict","Why flagged (matched term)"])
            st.write("**Token-by-token results**")
            st.dataframe(df_res, use_container_width=True)
    with colB:
        if summary_rows:
            df_sum = pd.DataFrame(summary_rows, columns=["Category","STRONG","MILD","Top fixes"])
            st.write("**Category summary**")
            st.dataframe(df_sum, use_container_width=True)

# ===== Activity Timing =====
with tabs[2]:
    st.subheader("Activity Timing – Astrology x Numerology")
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
                avoid_nums = str(row.get("Avoid Numbers (Numerology)",""))
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
    st.subheader("House Zone Checker – 5 Elements + Space (with Shapes)")
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
