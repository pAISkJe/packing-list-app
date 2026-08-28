# Packing List Generator (Streamlit)

A small web app that turns a parcel master workbook (`.xlsx`) into a ready-made
**`<Metal> Packing List`** workbook — computed columns, product photos, and rows
sorted by purity. Built so a team can use it from a browser, with no Python
install.

## What it does

1. Upload the parcel master workbook.
2. Pick the metal sheet (Gold / Silver / Brass …) and a few options.
3. Download a 2-sheet workbook: the **packing list** + a copy of the **source
   sheet** (photos preserved).

Columns are matched by **header text**, so it adapts to each metal sheet's
shifted layout automatically.

## Files

| File | Purpose |
|------|---------|
| `streamlit_app.py` | The web UI (Streamlit entry point) |
| `packing_core.py`  | Reusable generation logic (no UI) |
| `requirements.txt` | Python dependencies |
| `.gitignore`       | Keeps junk / data files out of git |

## Run locally (optional, to test before deploying)

```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```

Then open the local URL it prints (usually http://localhost:8501).

## Deploy on Streamlit Community Cloud

1. **Create a GitHub repo** (private is fine) and add these files
   (`streamlit_app.py`, `packing_core.py`, `requirements.txt`, `.gitignore`,
   `README.md`) to the repo root. Push.

   ```bash
   cd packing-list-app
   git init
   git add .
   git commit -m "Packing list generator"
   git branch -M main
   git remote add origin https://github.com/<you>/<repo>.git
   git push -u origin main
   ```

2. Go to **https://share.streamlit.io** and sign in with GitHub.

3. Click **Create app → Deploy a public app from GitHub** (you'll restrict
   access in the next step), and set:
   - **Repository**: your repo
   - **Branch**: `main`
   - **Main file path**: `streamlit_app.py`

   Click **Deploy**. First build takes a couple of minutes while it installs
   `requirements.txt`.

4. **Make it private to your team.** In the app's **Settings → Sharing**, turn
   off public access and add your team's Google/GitHub emails (India + US) as
   **viewers**. Only those addresses will be able to open the URL.

5. Share the app URL with the team. They upload a parcel file and download the
   packing list — nothing to install.

## Updating the app

Push new commits to the `main` branch; Streamlit Cloud redeploys automatically.

## Notes & limits

- The app reads **cached cell values**. If you just edited formulas in the
  master, open & save it once in Excel first so the values are current.
- The generated USD columns stay blank until you enter the **USD-to-INR rate**
  in cell **E1** of the packing-list sheet (Excel recalculates on open).
- These are business inventory files — keep the app **private** (step 4) and
  confirm uploading to a third-party host is allowed under your data policy.
