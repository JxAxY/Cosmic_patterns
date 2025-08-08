# Cosmic Generator (Streamlit) – Exact Moon Sign

This app shows **both Sun and Moon signs** and lets the user choose which one to use.
Moon sign uses **Swiss Ephemeris (pyswisseph)** for **exact** calculation from birthdate+time+timezone.

## Deploy on Streamlit Community Cloud
1. Create a public GitHub repo and upload:
   - `app.py`
   - `requirements.txt`
   - (optional) `data/cosmic_generator_vXX.xlsx` as a default
2. On Streamlit Cloud, set main file path to `app.py` and deploy.

## Using your latest workbook
- The app works best if you **upload your latest `cosmic_generator_vXX.xlsx`** via the sidebar uploader at runtime.
- This repo may include an older fallback in `/data`, but upload the newest for full features (Shape_Elements, etc.).
