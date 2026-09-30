# Tracker — no authentication

A Streamlit personal workspace with **no users, no passwords, and no login**.

## Streamlit Community Cloud

1. Push `app.py` and `requirements.txt` to GitHub.
2. Deploy `app.py` on Streamlit Community Cloud.
3. Open **Manage app → Settings → Secrets**.
4. Paste the contents of `.streamlit/secrets.toml.example` and replace placeholders with your real Google Sheets/service-account values.
5. Save/reboot the app.
6. Share the Google Sheet with the service-account `client_email` as Editor.

The app uses one shared Google Sheet workspace. Worksheets are named `Tasks`, `Workout`, `Ideas`, etc.; there are no per-user sheets.

For local development, if Google Sheets secrets are absent, the app falls back to `data/tracker.json`.
