# Cosmic Generator (Streamlit)

Public Streamlit app for the Cosmic Generator project. Loads the `cosmic_generator_v25.xlsx` workbook and provides:
- Life Audit (conflicts + remedies)
- Activity Timing Checker (Astrology × Numerology)
- House Zone Checker (5 elements + Space) with Shapes
- Element Items browser

## How to deploy on Streamlit Community Cloud

1. Create a **public GitHub repo** (e.g., `cosmic-generator-app`).
2. Add these files:
   - `app.py`
   - `requirements.txt`
   - `data/cosmic_generator_v25.xlsx`
   - `README.md`
3. Go to [share.streamlit.io](https://share.streamlit.io/) and connect your GitHub.
4. Select your repo, set **Main file path** = `app.py`, and click **Deploy**.

## Updating the data
- Replace `data/cosmic_generator_v25.xlsx` with a newer file (e.g., v26) and redeploy/restart the app.
- Or allow the app's sidebar uploader to use a custom XLSX for your session.

## Privacy
- This sample app does not store user data or uploads on the server.
