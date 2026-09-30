# Tracker — two-user Streamlit workspace

A Notion-style personal tracker for exactly two users.

## What changed

- Text columns accept unlimited text (`max_chars=None`).
- `st.data_editor` inputs are rebuilt with stable/reset indexes.
- Every row has a UUID instead of relying on row position.
- No insecure default passwords.
- No custom JavaScript authentication cookies.
- Local JSON storage uses atomic file replacement.
- Cloud persistence uses Google Sheets, not a database.
- Excel remains available as an export/backup format.
- User-provided roadmap HTML is escaped.
- Duplicate widget keys are avoided.
- Task start/end times are validated.
- Job links are validated.
- Habit completion stores the minutes at the time of completion.
- Historical habit minutes therefore do not change when the planned habit duration changes.
- Reminder deletion uses stable IDs.
- Workout target buttons use stable hashed widget keys.
- Data-editor indexes are reset before editing.
- Cloud writes are batched rather than rewriting one Excel workbook on every keystroke.
- The app explicitly requires secrets instead of falling back to public default passwords.

## Local run

```bash
pip install -r requirements.txt
streamlit run app.py
```

Create:

```text
.streamlit/secrets.toml
```

using `secrets.toml.example`.

If Google Sheets credentials are not configured, the app uses `./data/<user>.json`.

## Streamlit Community Cloud

For persistent cloud storage, configure Google Sheets in the app's Secrets settings.

1. Create a Google Cloud service account.
2. Enable Google Sheets API and Google Drive API.
3. Create a Google Sheet.
4. Share that Sheet with the service-account email as Editor.
5. Put the service-account JSON fields and `spreadsheet_id` in Streamlit Secrets.
6. Deploy with `requirements.txt`.

The application automatically creates separate worksheets such as:

```text
you__Tasks
you__Ideas
you__Jobs
partner__Tasks
partner__Ideas
partner__Jobs
```

The Google Sheet is not exposed through the UI; it is the persistence layer.

## Important hosting note

Do not rely on local JSON files for permanent data on Streamlit Community Cloud. Local filesystem persistence is not guaranteed there. Use the Google Sheets backend for the deployed version.
